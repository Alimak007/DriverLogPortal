<#
.SYNOPSIS
    Compares the real INTERNAL column names on the site against the names the
    web part actually queries, and reports every mismatch.

.DESCRIPTION
    Display names are not internal names. A column created as "Request Status"
    keeps the internal name Request_x0020_Status forever, even after it is
    renamed. The code only ever sees internal names, so a list that looks right
    in the browser can still fail every query.

    This reads the schema and prints, per list:
      OK       - the code's name exists
      MISSING  - the code's name does NOT exist, with near-miss suggestions
      (extra)  - columns on the list the code never touches (informational)

    READ-ONLY. It changes nothing.

.EXAMPLE
    .\4-Check-Column-Names.ps1 -SiteUrl "https://pomae.sharepoint.com/sites/travel-dev/driverlogportal"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SiteUrl,
    [switch] $ShowExtras
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell -DisableNameChecking
Connect-PnPOnline -Url $SiteUrl -Interactive

# What the code queries. Taken from src/common/services/mappers.ts and
# DispatchService.ts - if you change a name there, change it here too.
$EXPECTED = [ordered]@{
    'TransportRequests' = @(
        'Title','RequestType','Requester','Property','PickupLocation','DropOffLocation',
        'RequiredDateTime','ReturnDateTime','NoOfPax','VehicleTypeRequested',
        'OtherVehicleText','ItemDescription','ItemNotes','Priority','ReqStatus',
        'OfferClosesAt','ReleaseAt','AcceptedTrip','Remarks','RejectionReason',
        'CancellationReason'
    )
    'RideOffers' = @(
        'Title','Request','Driver','Vehicle','OfferedOn','ResponseStatus',
        'RejectionReason','RejectionNotes','RespondedOn'
    )
    'TripAssignments' = @(
        'Title','Request','Driver','Vehicle','AcceptedOn','TripStatus','EtaMinutes',
        'EtaUpdatedOn','ActualStartTime','ActualPickupTime','ActualDropTime',
        'OpeningReading','ClosingReading','AllocationMode','AllocatedBy',
        'IsOverride','OverrideReason','CancellationReason'
    )
    'TripStatusHistory' = @(
        'Title','Trip','FromStatus','ToStatus','ChangedBy','ChangedOn','Comment'
    )
    'DriversMaster' = @(
        'Title','UserAccount','DriverID','Phone','LicenceNumber',
        'LicenceExpiry','QualifiedVehicleTypes','HomeProperty','Status',
        'MaxConcurrentTrips','IsActive'
    )
    'VehicleMaster' = @(
        'Title','PlateNo','VehicleType','Make','Model','Colour','Capacity',
        'DefaultDriver','PropertyRef','VehStatus','IsActive'
    )
    'VehicleTypes'       = @('Title','SuitableFor','Capacity','SortOrder','IsActive')
    'Locations'          = @('Title','Property','Address','Latitude','Longitude','Category','IsActive')
    'AllocationRules'    = @(
        'Title','Property','RequestType','Mode','OfferWindowMinutes',
        'MaxConcurrentTripsPerDriver','AdvanceReleaseHours','IsActive',
        'NotifyEnRoute','NotifyInTransit'
    )
    'AvailabilityBlocks' = @('Title','Driver','Vehicle','FromDateTime','ToDateTime','BlockType','Property','Notes')
    'DLClientMaster'     = @('DriverID','PropertyRef','PropertyAllocation')
    'PropertyMaster'     = @('Title')
    'Templates'          = @('Title','Language','Subject','Body','IsActive')
}

# Columns every list has; never report them as extras.
$BUILTIN = @(
    'ContentType','Attachments','Edit','LinkTitleNoMenu','LinkTitle','DocIcon','ItemChildCount',
    'FolderChildCount','_ComplianceFlags','_ComplianceTag','_ComplianceTagWrittenTime',
    '_ComplianceTagUserId','_IsRecord','AppAuthor','AppEditor','Author','Editor','Created',
    'Modified','ID','_UIVersionString','ComplianceAssetId','_CommentCount','_LikeCount',
    '_DisplayName','_ColorHex','_ColorTag','_Flags','TaxCatchAll','TaxCatchAllLabel',
    'MediaServiceImageTags','_ShortcutUrl','_ShortcutSiteId','_ShortcutWebId','_ShortcutUniqueId'
)

function Suggest {
    param([string] $Wanted, [string[]] $Actual)
    # Near misses: same letters ignoring case, spaces and the _x0020_ encoding.
    $norm = { param($x) ($x -replace '_x0020_','' -replace '[^A-Za-z0-9]','').ToLower() }
    $w = & $norm $Wanted
    $hits = $Actual | Where-Object { (& $norm $_) -eq $w -or (& $norm $_) -like "*$w*" -or $w -like "*$(& $norm $_)*" }
    if ($hits) { return ($hits -join ', ') }
    return ''
}

$problems = 0

foreach ($listName in $EXPECTED.Keys) {
    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "`n### $listName" -ForegroundColor Red
        Write-Host "    LIST NOT FOUND on this web." -ForegroundColor Red
        $problems++
        continue
    }

    $fields = Get-PnPField -List $listName |
              Where-Object { -not $_.Hidden } |
              Select-Object -ExpandProperty InternalName

    Write-Host "`n### $listName" -ForegroundColor Cyan

    foreach ($want in $EXPECTED[$listName]) {
        if ($fields -contains $want) {
            Write-Host ("    OK       {0}" -f $want) -ForegroundColor DarkGray
        }
        else {
            $problems++
            $hint = Suggest -Wanted $want -Actual $fields
            if ($hint) {
                Write-Host ("    MISSING  {0}   <-- did you mean: {1}" -f $want, $hint) -ForegroundColor Yellow
            }
            else {
                Write-Host ("    MISSING  {0}" -f $want) -ForegroundColor Red
            }
        }
    }

    if ($ShowExtras) {
        $extras = $fields | Where-Object { $_ -notin $EXPECTED[$listName] -and $_ -notin $BUILTIN }
        foreach ($e in $extras) {
            Write-Host ("    (extra)  {0}" -f $e) -ForegroundColor DarkGray
        }
    }
}

Write-Host ""
if ($problems -eq 0) {
    Write-Host "No mismatches. Every column the code queries exists." -ForegroundColor Green
}
else {
    Write-Host "$problems mismatch(es). Each MISSING line is a query that will throw." -ForegroundColor Yellow
    Write-Host "Send me this output and I will correct the names in the code." -ForegroundColor Yellow
}

Disconnect-PnPOnline
