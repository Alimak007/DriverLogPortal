<#
.SYNOPSIS
    Explains why a request escalated to "Awaiting Coordinator" instead of
    broadcasting.

.DESCRIPTION
    Replays the exact eligibility evaluation the web part runs, against the
    live lists, and prints the verdict for every driver plus the vehicle check.

    "Awaiting Coordinator" means broadcast() ran all the way through without
    throwing - so the column names are right and this is a DATA problem. One of
    two things is true:

        no driver passed all five checks,   or
        no vehicle of the requested type was free.

    This tells you which, and why.

    READ-ONLY. It changes nothing.

.EXAMPLE
    .\5-Diagnose-Eligibility.ps1 -SiteUrl "https://pomae.sharepoint.com/sites/travel-dev/driverlogportal"
    .\5-Diagnose-Eligibility.ps1 -SiteUrl $url -RequestId 7
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SiteUrl,

    # Defaults to the most recent request that is Submitted or Awaiting Coordinator.
    [int] $RequestId
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell -DisableNameChecking
Connect-PnPOnline -Url $SiteUrl -Interactive

function F { param($Item, [string] $Name) $Item.FieldValues[$Name] }
function Head { param($t) Write-Host ""; Write-Host $t -ForegroundColor Cyan }

# ---------------------------------------------------------------- the request
if (-not $RequestId) {
    $candidates = Get-PnPListItem -List 'TransportRequests' -PageSize 500 |
        Where-Object { (F $_ 'ReqStatus') -in 'Submitted', 'Awaiting Coordinator' } |
        Sort-Object Id -Descending
    if (-not $candidates) { throw "No request is Submitted or Awaiting Coordinator. Pass -RequestId explicitly." }
    $RequestId = $candidates[0].Id
    Write-Host "Using the most recent stuck request: item $RequestId" -ForegroundColor Yellow
}

$req = Get-PnPListItem -List 'TransportRequests' -Id $RequestId
$ref        = F $req 'Title'
$status     = F $req 'ReqStatus'
$propertyId = (F $req 'Property').LookupId
$propName   = (F $req 'Property').LookupValue
$typeId     = (F $req 'VehicleTypeRequested').LookupId
$typeName   = (F $req 'VehicleTypeRequested').LookupValue
$windowFrom = [datetime](F $req 'RequiredDateTime')
$ret        = F $req 'ReturnDateTime'
$windowTo   = if ($ret) { [datetime]$ret } else { $windowFrom.AddHours(2) }

Head "REQUEST"
"  $ref  (item $RequestId)"
"  Status        : $status"
"  Property      : $propName (id $propertyId)"
"  Vehicle type  : $typeName (id $typeId)"
"  Trip window   : $($windowFrom.ToString('u')) -> $($windowTo.ToString('u'))"

if (-not $propertyId) {
    Write-Host "`n  STOP: the request has no Property. The booking form sent an empty property." -ForegroundColor Red
    Write-Host "  Set 'Default property ID' in the Book a Ride web part property pane." -ForegroundColor Red
    Disconnect-PnPOnline; return
}
if (-not $typeId) {
    Write-Host "`n  STOP: the request has no Vehicle Type." -ForegroundColor Red
    Disconnect-PnPOnline; return
}

# ------------------------------------------------- check 1: DLClientMaster
Head "CHECK 1 - DLClientMaster (which drivers serve this property)"
$links = Get-PnPListItem -List 'DLClientMaster' -PageSize 2000
$forProperty = $links | Where-Object { (F $_ 'PropertyRef').LookupId -eq $propertyId }
$allocated = @($forProperty | ForEach-Object { (F $_ 'DriverID').LookupId } | Where-Object { $_ })

"  rows in DLClientMaster        : $($links.Count)"
"  rows for property $propertyId : $($forProperty.Count)"
"  distinct driver ids           : $(($allocated | Select-Object -Unique) -join ', ')"

if ($links.Count -gt 0 -and $forProperty.Count -eq 0) {
    Write-Host "  >>> No row points at this property." -ForegroundColor Red
    Write-Host "      Either PropertyRef is blank (backfill it from PropertyAllocation)," -ForegroundColor Red
    Write-Host "      or every row points at a different property." -ForegroundColor Red
    $sample = $links | Select-Object -First 5
    "      First rows' PropertyRef values:"
    foreach ($l in $sample) {
        "        item $($l.Id): PropertyRef=$((F $l 'PropertyRef').LookupValue) [$((F $l 'PropertyRef').LookupId)]  PropertyAllocation=$(F $l 'PropertyAllocation')"
    }
}
if ($allocated.Count -eq 0 -and $forProperty.Count -gt 0) {
    Write-Host "  >>> Rows match the property but DriverID is empty on all of them." -ForegroundColor Red
}

# ----------------------------------------------------------- the other inputs
$blocks = Get-PnPListItem -List 'AvailabilityBlocks' -PageSize 500 | Where-Object {
    $f = F $_ 'FromDateTime'; $t = F $_ 'ToDateTime'
    $f -and $t -and ([datetime]$f -lt $windowTo) -and ([datetime]$t -gt $windowFrom)
}
$openTrips = Get-PnPListItem -List 'TripAssignments' -PageSize 500 |
    Where-Object { (F $_ 'TripStatus') -in 'Assigned', 'En Route', 'In Transit' }

$rule = Get-PnPListItem -List 'AllocationRules' -PageSize 200 | Where-Object {
    (F $_ 'Property').LookupId -eq $propertyId -and (F $_ 'RequestType') -eq (F $req 'RequestType') -and (F $_ 'IsActive')
} | Select-Object -First 1
$ruleCap = if ($rule) { [int](F $rule 'MaxConcurrentTripsPerDriver') } else { 1 }
if ($ruleCap -le 0) { $ruleCap = 1 }

Head "CHECK 2-5 - every driver, every check"
$drivers = Get-PnPListItem -List 'DriversMaster' -PageSize 500
$eligible = 0

foreach ($d in $drivers) {
    $name = F $d 'Title'
    $why  = @()

    if (-not (F $d 'IsActive'))           { $why += 'IsActive = No' }
    if ((F $d 'Status') -eq 'Inactive')   { $why += 'Status = Inactive' }

    if ($d.Id -notin $allocated) { $why += 'not allocated to this property (DLClientMaster)' }

    $quals = @(F $d 'QualifiedVehicleTypes' | ForEach-Object { $_.LookupId })
    if ($quals.Count -eq 0)        { $why += 'QualifiedVehicleTypes is EMPTY' }
    elseif ($typeId -notin $quals) { $why += "not qualified for $typeName (has: $(($quals) -join ','))" }

    $blocked = $blocks | Where-Object {
        (F $_ 'Driver').LookupId -eq $d.Id -and (F $_ 'BlockType') -ne 'Shift'
    }
    if ($blocked) { $why += "blocked in the window ($((F $blocked[0] 'BlockType')))" }

    $exp = F $d 'LicenceExpiry'
    if ($exp -and [datetime]$exp -lt $windowTo) {
        $why += "licence expires $([datetime]$exp | Get-Date -Format 'yyyy-MM-dd'), before the window ends"
    }

    $held = @($openTrips | Where-Object { (F $_ 'Driver').LookupId -eq $d.Id }).Count
    $cap  = [int](F $d 'MaxConcurrentTrips'); if ($cap -le 0) { $cap = $ruleCap }
    if ($held -ge $cap) { $why += "at the cap (holding $held of $cap)" }

    if ($why.Count -eq 0) {
        $eligible++
        Write-Host ("  ELIGIBLE  {0}" -f $name) -ForegroundColor Green
    }
    else {
        Write-Host ("  excluded  {0,-24} {1}" -f $name, ($why -join '; ')) -ForegroundColor DarkGray
    }
}

# -------------------------------------------------------- the vehicle check
Head "VEHICLE CHECK - is any vehicle of this type free?"
$vehicles = Get-PnPListItem -List 'VehicleMaster' -PageSize 500
$candidates = $vehicles | Where-Object {
    (F $_ 'VehicleType').LookupId -eq $typeId -and
    (F $_ 'PropertyRef').LookupId -eq $propertyId -and
    (F $_ 'VehStatus') -notin 'Retired', 'Under Maintenance'
}
"  vehicles in VehicleMaster            : $($vehicles.Count)"
"  of type '$typeName' at this property : $($candidates.Count)"

$busyIds = @($openTrips | ForEach-Object { (F $_ 'Vehicle').LookupId })
$free = $candidates | Where-Object {
    $_.Id -notin $busyIds -and
    -not ($blocks | Where-Object { (F $_ 'Vehicle').LookupId -eq $_.Id -and (F $_ 'BlockType') -ne 'Shift' })
}
"  free right now                       : $(@($free).Count)"

if (@($candidates).Count -eq 0) {
    Write-Host "  >>> No vehicle matched. Check VehicleMaster.PropertyRef is populated" -ForegroundColor Red
    Write-Host "      and VehicleType points at '$typeName'." -ForegroundColor Red
    $byType = $vehicles | Where-Object { (F $_ 'VehicleType').LookupId -eq $typeId }
    "      vehicles of this type at ANY property: $(@($byType).Count)"
    foreach ($v in ($byType | Select-Object -First 5)) {
        "        $(F $v 'Title'): PropertyRef=$((F $v 'PropertyRef').LookupValue) [$((F $v 'PropertyRef').LookupId)]  VehStatus=$(F $v 'VehStatus')"
    }
}

# ------------------------------------------------------------------- verdict
Head "VERDICT"
if ($eligible -eq 0 -and @($free).Count -eq 0) {
    Write-Host "  Both failed: no eligible driver AND no free vehicle." -ForegroundColor Red
}
elseif ($eligible -eq 0) {
    Write-Host "  No eligible driver. The reasons above are the cause." -ForegroundColor Red
}
elseif (@($free).Count -eq 0) {
    Write-Host "  Drivers were eligible ($eligible) but NO VEHICLE was free." -ForegroundColor Red
}
else {
    Write-Host "  $eligible eligible driver(s) and $(@($free).Count) free vehicle(s)." -ForegroundColor Green
    Write-Host "  This request should have broadcast. If it did not, re-book with the" -ForegroundColor Yellow
    Write-Host "  browser console open - broadcast() now logs its reasoning." -ForegroundColor Yellow
}
Write-Host ""

Disconnect-PnPOnline
