<#
.SYNOPSIS
    Loads dummy data for testing the Fleet Booking and Dispatch Portal.

.DESCRIPTION
    Every row here exists to exercise a specific screen. The six seeded
    requests cover six different statuses on purpose, so the dashboards have
    something in every state the moment you open them.

    NOT IDEMPOTENT. It checks whether the lists already hold data and stops
    rather than creating duplicates. Use -Force to wipe the transactional
    lists and reseed.

.EXAMPLE
    .\2-Seed-TestData.ps1 -SiteUrl "https://pomae.sharepoint.com/sites/travel-dev" `
                          -TesterUpn "aliakbar.makada@queinfotech.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SiteUrl,

    # The account you will sign in as. Gets its own DriversMaster record, so
    # you can test the Driver Dashboard as yourself.
    [Parameter(Mandatory = $true)][string] $TesterUpn,

    # Extra real accounts to use as drivers 2 and 3. Optional - without them
    # the other drivers get no UserAccount and can be allocated but cannot
    # sign in. That is fine for testing the Coordinator views.
    [string[]] $OtherDriverUpns = @(),

    [switch] $Force
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell -DisableNameChecking

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# Everything is stored UTC. Gulf Standard Time is UTC+4, so a trip the
# interface shows at 09:00 is stored at 05:00Z.
function U { param([datetime] $d) $d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
$now      = (Get-Date).ToUniversalTime()
$todayGst = [datetime]::SpecifyKind($now.Date, 'Utc')

function Seed {
    param([string] $List, [hashtable] $Values)
    Add-PnPListItem -List $List -Values $Values
}

function Get-Id {
    param([string] $List, [string] $Where)
    $item = Get-PnPListItem -List $List -Query "<View><Query><Where>$Where</Where></Query><RowLimit>1</RowLimit></View>"
    if (-not $item) { throw "No item found in $List matching $Where" }
    $item.Id
}

function TitleWhere { param($t) "<Eq><FieldRef Name='Title'/><Value Type='Text'>$t</Value></Eq>" }

# -----------------------------------------------------------------------------
# Guard
# -----------------------------------------------------------------------------

$transactional = 'TripStatusHistory', 'RideOffers', 'TripAssignments', 'TransportRequests', 'NotificationLog'
$existing = Get-PnPListItem -List 'TransportRequests' -PageSize 1
if ($existing -and -not $Force) {
    Write-Warning "TransportRequests already holds data. Re-run with -Force to wipe and reseed, or seed a clean site."
    Disconnect-PnPOnline
    return
}

if ($Force) {
    Write-Host "Clearing transactional lists ..." -ForegroundColor Yellow
    foreach ($l in $transactional) {
        $items = Get-PnPListItem -List $l -PageSize 500
        foreach ($i in $items) { Remove-PnPListItem -List $l -Identity $i.Id -Force | Out-Null }
        Write-Host "  cleared $l ($($items.Count) items)" -ForegroundColor DarkGray
    }
}

# -----------------------------------------------------------------------------
# 1. PropertyMaster
# -----------------------------------------------------------------------------

Write-Host "`n=== PropertyMaster ===" -ForegroundColor Cyan
$properties = 'Dubai Head Office', 'Abu Dhabi Site'
foreach ($p in $properties) {
    if (-not (Get-PnPListItem -List 'PropertyMaster' -Query "<View><Query><Where>$(TitleWhere $p)</Where></Query></View>")) {
        Seed 'PropertyMaster' @{ Title = $p } | Out-Null
        Write-Host "  + $p" -ForegroundColor Gray
    }
}
$P1 = Get-Id 'PropertyMaster' (TitleWhere 'Dubai Head Office')
$P2 = Get-Id 'PropertyMaster' (TitleWhere 'Abu Dhabi Site')

# -----------------------------------------------------------------------------
# 2. VehicleTypes - the eight types, as data, not hard-coded
# -----------------------------------------------------------------------------

Write-Host "`n=== VehicleTypes ===" -ForegroundColor Cyan
$types = @(
    @{ Title = 'Patrol';           Sort = 1; Cap = 7; For = @('Client or Client Guest', 'Staff') }
    @{ Title = 'Pathfinder';       Sort = 2; Cap = 7; For = @('Client or Client Guest', 'Staff') }
    @{ Title = 'Kicks / Creta';    Sort = 3; Cap = 5; For = @('Client or Client Guest', 'Staff') }
    @{ Title = 'Cargo Van';        Sort = 4; Cap = 2; For = @('Items or Documents') }
    @{ Title = 'Chiller Van';      Sort = 5; Cap = 2; For = @('Items or Documents') }
    @{ Title = 'Hot and Cold Van'; Sort = 6; Cap = 2; For = @('Items or Documents') }
    @{ Title = 'Animal Van';       Sort = 7; Cap = 2; For = @('Items or Documents') }
    # "Others" is the single free-text fallback and is offered for every
    # request type - selecting it reveals the only free-text field in the form.
    @{ Title = 'Others';           Sort = 8; Cap = 0; For = @('Client or Client Guest', 'Staff', 'Items or Documents') }
)
foreach ($t in $types) {
    if (-not (Get-PnPListItem -List 'VehicleTypes' -Query "<View><Query><Where>$(TitleWhere $t.Title)</Where></Query></View>")) {
        Seed 'VehicleTypes' @{
            Title       = $t.Title
            SuitableFor = $t.For
            Capacity    = $t.Cap
            SortOrder   = $t.Sort
            IsActive    = $true
        } | Out-Null
        Write-Host "  + $($t.Title)" -ForegroundColor Gray
    }
}
$VT = @{}
foreach ($t in $types) { $VT[$t.Title] = Get-Id 'VehicleTypes' (TitleWhere $t.Title) }

# -----------------------------------------------------------------------------
# 3. Locations - removes free-text destinations
# -----------------------------------------------------------------------------

Write-Host "`n=== Locations ===" -ForegroundColor Cyan
$locations = @(
    @{ Title = 'Dubai Head Office';          P = $P1; Cat = 'Office';    Addr = 'Sheikh Zayed Road, Dubai' }
    @{ Title = 'Jebel Ali Warehouse';        P = $P1; Cat = 'Warehouse'; Addr = 'Jebel Ali Free Zone, Dubai' }
    @{ Title = 'Dubai International T3';     P = $P1; Cat = 'Airport';   Addr = 'Terminal 3, Dubai International Airport' }
    @{ Title = 'Al Barsha Clinic';           P = $P1; Cat = 'Hospital';  Addr = 'Al Barsha 1, Dubai' }
    @{ Title = 'Marina Staff Residence';     P = $P1; Cat = 'Residence'; Addr = 'Dubai Marina, Dubai' }
    @{ Title = 'Downtown Client Office';     P = $P1; Cat = 'Office';    Addr = 'Burj Khalifa Boulevard, Dubai' }
    @{ Title = 'Abu Dhabi Site Office';      P = $P2; Cat = 'Office';    Addr = 'Corniche Road, Abu Dhabi' }
    @{ Title = 'Yas Island Hotel';           P = $P2; Cat = 'Hotel';     Addr = 'Yas Island, Abu Dhabi' }
    @{ Title = 'Mussafah Depot';             P = $P2; Cat = 'Warehouse'; Addr = 'Mussafah Industrial, Abu Dhabi' }
    @{ Title = 'Zayed International Airport';P = $P2; Cat = 'Airport';   Addr = 'Abu Dhabi' }
)
foreach ($l in $locations) {
    if (-not (Get-PnPListItem -List 'Locations' -Query "<View><Query><Where>$(TitleWhere $l.Title)</Where></Query></View>")) {
        # Latitude and Longitude are deliberately left empty. They exist from
        # day one so phase 2 geocoding needs no schema migration.
        Seed 'Locations' @{
            Title = $l.Title; Property = $l.P; Category = $l.Cat; Address = $l.Addr; IsActive = $true
        } | Out-Null
        Write-Host "  + $($l.Title)" -ForegroundColor Gray
    }
}
$LOC = @{}
foreach ($l in $locations) { $LOC[$l.Title] = Get-Id 'Locations' (TitleWhere $l.Title) }

# -----------------------------------------------------------------------------
# 4. DriversMaster
# -----------------------------------------------------------------------------

Write-Host "`n=== DriversMaster ===" -ForegroundColor Cyan
$passengerTypes = @($VT['Patrol'], $VT['Pathfinder'], $VT['Kicks / Creta'], $VT['Others'])
$vanTypes       = @($VT['Cargo Van'], $VT['Chiller Van'], $VT['Hot and Cold Van'], $VT['Animal Van'], $VT['Others'])

$drivers = @(
    @{ Name = 'Rashid Al Mansoori'; Id = 'DRV-001'; Phone = '+971 50 111 2233'; Lic = 'DXB-884412'
       Exp = $now.AddYears(2);  Types = $passengerTypes; Prop = $P1; Status = 'Available'; Max = 2
       Upn = $TesterUpn }
    @{ Name = 'Suresh Nair';        Id = 'DRV-002'; Phone = '+971 50 222 3344'; Lic = 'DXB-771203'
       Exp = $now.AddYears(1);  Types = $passengerTypes; Prop = $P1; Status = 'Available'; Max = 1
       Upn = $(if ($OtherDriverUpns.Count -gt 0) { $OtherDriverUpns[0] } else { $null }) }
    @{ Name = 'Mohammed Iqbal';     Id = 'DRV-003'; Phone = '+971 50 333 4455'; Lic = 'DXB-662119'
       Exp = $now.AddMonths(8); Types = $vanTypes;       Prop = $P1; Status = 'Available'; Max = 1
       Upn = $(if ($OtherDriverUpns.Count -gt 1) { $OtherDriverUpns[1] } else { $null }) }
    @{ Name = 'Jomon Varghese';     Id = 'DRV-004'; Phone = '+971 50 444 5566'; Lic = 'AUH-553027'
       Exp = $now.AddYears(3);  Types = $passengerTypes; Prop = $P2; Status = 'Available'; Max = 1
       Upn = $null }
    # On leave across today - proves the availability panel and the eligibility
    # filter both exclude them.
    @{ Name = 'Ali Hassan';         Id = 'DRV-005'; Phone = '+971 50 555 6677'; Lic = 'DXB-449981'
       Exp = $now.AddYears(1);  Types = $passengerTypes; Prop = $P1; Status = 'On Leave'; Max = 1
       Upn = $null }
    # Licence expires in 10 days - any trip window past that date must exclude
    # them. Use this one to test the licence check.
    @{ Name = 'Arun Prasad';        Id = 'DRV-006'; Phone = '+971 50 666 7788'; Lic = 'DXB-330876'
       Exp = $now.AddDays(10);  Types = $vanTypes;       Prop = $P1; Status = 'Available'; Max = 1
       Upn = $null }
)

foreach ($d in $drivers) {
    if (Get-PnPListItem -List 'DriversMaster' -Query "<View><Query><Where>$(TitleWhere $d.Name)</Where></Query></View>") { continue }
    $values = @{
        Title                 = $d.Name     # The driver's name IS Title (v2 spec s.12)
        DriverID              = $d.Id
        Phone                 = $d.Phone
        LicenceNumber         = $d.Lic
        LicenceExpiry         = U $d.Exp
        QualifiedVehicleTypes = $d.Types
        HomeProperty          = $d.Prop
        Status                = $d.Status
        MaxConcurrentTrips    = $d.Max
        IsActive              = $true
    }
    if ($d.Upn) { $values['UserAccount'] = $d.Upn }
    Seed 'DriversMaster' $values | Out-Null
    Write-Host "  + $($d.Name)$(if ($d.Upn) { " [$($d.Upn)]" })" -ForegroundColor Gray
}
$DRV = @{}
foreach ($d in $drivers) { $DRV[$d.Name] = Get-Id 'DriversMaster' (TitleWhere $d.Name) }

# -----------------------------------------------------------------------------
# 5. DLClientMaster - driver to property. A driver may appear more than once.
# -----------------------------------------------------------------------------

Write-Host "`n=== DLClientMaster ===" -ForegroundColor Cyan
$links = @(
    @{ D = 'Rashid Al Mansoori'; P = $P1 }
    @{ D = 'Rashid Al Mansoori'; P = $P2 }   # deliberately multi-property
    @{ D = 'Suresh Nair';        P = $P1 }
    @{ D = 'Mohammed Iqbal';     P = $P1 }
    @{ D = 'Jomon Varghese';     P = $P2 }
    @{ D = 'Ali Hassan';         P = $P1 }
    @{ D = 'Arun Prasad';        P = $P1 }
)
foreach ($l in $links) {
    Seed 'DLClientMaster' @{ Title = "$($l.D) allocation"; DriverID = $DRV[$l.D]; PropertyRef = $l.P } | Out-Null
}
Write-Host "  + $($links.Count) allocations (Rashid is on BOTH properties on purpose)" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 6. VehicleMaster
# -----------------------------------------------------------------------------

Write-Host "`n=== VehicleMaster ===" -ForegroundColor Cyan
$vehicles = @(
    @{ Plate = 'A 12345'; Type = 'Patrol';           Make = 'Nissan'; Model = 'Patrol';     Col = 'White';  Cap = 7; P = $P1; Def = 'Rashid Al Mansoori'; St = 'Available' }
    @{ Plate = 'B 23456'; Type = 'Pathfinder';       Make = 'Nissan'; Model = 'Pathfinder'; Col = 'Silver'; Cap = 7; P = $P1; Def = 'Suresh Nair';        St = 'Available' }
    @{ Plate = 'C 34567'; Type = 'Kicks / Creta';    Make = 'Hyundai';Model = 'Creta';      Col = 'Grey';   Cap = 5; P = $P1; Def = $null;                St = 'Available' }
    @{ Plate = 'D 45678'; Type = 'Cargo Van';        Make = 'Toyota'; Model = 'Hiace';      Col = 'White';  Cap = 2; P = $P1; Def = 'Mohammed Iqbal';     St = 'Available' }
    # Under maintenance - the Available vehicles counter must not count it.
    @{ Plate = 'E 56789'; Type = 'Chiller Van';      Make = 'Isuzu';  Model = 'NPR';        Col = 'White';  Cap = 2; P = $P1; Def = $null;                St = 'Under Maintenance' }
    @{ Plate = 'F 67890'; Type = 'Hot and Cold Van'; Make = 'Isuzu';  Model = 'NQR';        Col = 'White';  Cap = 2; P = $P1; Def = $null;                St = 'Available' }
    @{ Plate = 'G 78901'; Type = 'Animal Van';       Make = 'Mitsubishi'; Model = 'Canter'; Col = 'Blue';   Cap = 2; P = $P1; Def = $null;                St = 'Available' }
    @{ Plate = 'H 89012'; Type = 'Patrol';           Make = 'Nissan'; Model = 'Patrol';     Col = 'Black';  Cap = 7; P = $P2; Def = 'Jomon Varghese';     St = 'Available' }
)
foreach ($v in $vehicles) {
    if (Get-PnPListItem -List 'VehicleMaster' -Query "<View><Query><Where>$(TitleWhere $v.Plate)</Where></Query></View>") { continue }
    $values = @{
        Title = $v.Plate            # Title mirrors PlateNo; lookups show it
        PlateNo = $v.Plate; VehicleType = $VT[$v.Type]
        Make = $v.Make; Model = $v.Model; Colour = $v.Col; Capacity = $v.Cap
        PropertyRef = $v.P; VehStatus = $v.St
    }
    if ($v.Def) { $values['DefaultDriver'] = $DRV[$v.Def] }
    Seed 'VehicleMaster' $values | Out-Null
    Write-Host "  + $($v.Plate) $($v.Type)" -ForegroundColor Gray
}
$VEH = @{}
foreach ($v in $vehicles) { $VEH[$v.Plate] = Get-Id 'VehicleMaster' (TitleWhere $v.Plate) }

# -----------------------------------------------------------------------------
# 7. AllocationRules - one per property per request type
# -----------------------------------------------------------------------------

Write-Host "`n=== AllocationRules ===" -ForegroundColor Cyan
$requestTypes = 'Client or Client Guest', 'Staff', 'Items or Documents'
foreach ($pid in @($P1, $P2)) {
    $pname = if ($pid -eq $P1) { 'Dubai Head Office' } else { 'Abu Dhabi Site' }
    foreach ($rt in $requestTypes) {
        $title = "$pname - $rt"
        if (Get-PnPListItem -List 'AllocationRules' -Query "<View><Query><Where>$(TitleWhere $title)</Where></Query></View>") { continue }
        Seed 'AllocationRules' @{
            Title = $title; Property = $pid; RequestType = $rt
            Mode = 'Automatic'
            OfferWindowMinutes = 5; MaxConcurrentTripsPerDriver = 1; AdvanceReleaseHours = 4
            IsActive = $true; NotifyEnRoute = $true; NotifyInTransit = $true
        } | Out-Null
    }
}
Write-Host "  + 6 rules (Automatic, 5 min window, cap 1, release 4 h)" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 8. AvailabilityBlocks
# -----------------------------------------------------------------------------

Write-Host "`n=== AvailabilityBlocks ===" -ForegroundColor Cyan
Seed 'AvailabilityBlocks' @{
    Title = 'Ali Hassan - annual leave'; Driver = $DRV['Ali Hassan']
    FromDateTime = U $todayGst.AddDays(-2); ToDateTime = U $todayGst.AddDays(5)
    BlockType = 'Leave'; Property = $P1; Notes = 'Annual leave. Excluded from every broadcast this week.'
} | Out-Null
Seed 'AvailabilityBlocks' @{
    Title = 'E 56789 - service'; Vehicle = $VEH['E 56789']
    FromDateTime = U $todayGst.AddDays(-1); ToDateTime = U $todayGst.AddDays(3)
    BlockType = 'Maintenance'; Property = $P1; Notes = 'Compressor service.'
} | Out-Null
Write-Host "  + 2 blocks (one driver on leave, one vehicle in maintenance)" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 9. Templates
# -----------------------------------------------------------------------------

Write-Host "`n=== Templates ===" -ForegroundColor Cyan
$templates = @(
    @{ K = 'RequestSubmitted.Requester';   S = 'We have your transport request {RequestId}';                      B = 'We have received your request {RequestId} for {RequiredDateTime}, {PickupLocation} to {DropOffLocation}. We are finding you a driver and will confirm shortly. This is an acknowledgement, not a confirmation.' }
    # The single highest-value message in the system. The subject must be
    # readable from a lock-screen preview: pickup, destination and time.
    @{ K = 'BroadcastOpened.Driver';       S = '{PickupLocation} to {DropOffLocation}, {RequiredTime} - first to accept';  B = 'A new ride is available. {RequestType}, {NoOfPax} passenger(s), {VehicleType}. Open the driver dashboard to accept: {DashboardDeepLink}. This offer closes at {OfferClosesAt}.' }
    @{ K = 'DriverAccepted.Driver';        S = 'You have the trip {RequestId}';                                   B = 'You accepted {RequestId}. Pick up {RequesterName} at {PickupLocation} at {RequiredDateTime}, going to {DropOffLocation}. Contact: {RequesterPhone}.' }
    @{ K = 'DriverAccepted.Requester';     S = 'Your ride is confirmed - {RequestId}';                            B = '{DriverName} is driving you, in a {VehicleColour} {VehicleMake} {VehicleModel}, plate {PlateNo}. Call them on {DriverPhone}. Pick up {RequiredDateTime} at {PickupLocation}.' }
    @{ K = 'WindowExpired.Coordinator';    S = 'Nobody accepted {RequestId} - needs allocating';                  B = '{RequestId} was offered to {OfferedCount} driver(s) and nobody accepted within the window. Declines: {RejectionSummary}. Assign it from the dispatch dashboard.' }
    @{ K = 'ManualAssign.Driver';          S = 'You have been assigned trip {RequestId}';                         B = 'The Transportation Coordinator has assigned you {RequestId}. {PickupLocation} to {DropOffLocation} at {RequiredDateTime}.' }
    @{ K = 'ManualAssign.Requester';       S = 'Your ride is confirmed - {RequestId}';                            B = '{DriverName} is driving you, in a {VehicleColour} {VehicleMake} {VehicleModel}, plate {PlateNo}. Call them on {DriverPhone}.' }
    @{ K = 'StatusEnRoute.Requester';      S = 'Your driver is on the way';                                       B = '{DriverName} is on the way to {PickupLocation}. Plate {PlateNo}.' }
    @{ K = 'StatusInTransit.Requester';    S = 'On the way to {DropOffLocation}';                                 B = 'You are on the way to {DropOffLocation}.' }
    @{ K = 'TripCompleted.Requester';      S = 'Trip {RequestId} completed';                                      B = 'Your trip finished at {ActualDropTime}. Tell us how it went: {FeedbackLink}' }
    @{ K = 'Cancelled.Driver';             S = 'Trip {RequestId} cancelled';                                      B = '{RequestId} has been cancelled. Reason: {Reason}. You are released.' }
    @{ K = 'Cancelled.Requester';          S = 'Your booking {RequestId} was cancelled';                          B = 'Reason: {Reason}. Book again if you still need transport.' }
    @{ K = 'Cancelled.Coordinator';        S = '{RequestId} cancelled by {CancelledBy}';                          B = 'Reason: {Reason}.' }
    @{ K = 'RequestRejected.Requester';    S = 'Your request {RequestId} was declined';                           B = 'Reason: {Reason}. Contact the Transportation Coordinator if you need to discuss it.' }
    @{ K = 'NotificationFailed.Coordinator'; S = 'A notification failed to send';                                 B = '{TemplateKey} to {Recipient} failed after {RetryCount} retries: {ErrorMessage}. A driver may not know they have a trip.' }
)
foreach ($t in $templates) {
    if (Get-PnPListItem -List 'Templates' -Query "<View><Query><Where>$(TitleWhere $t.K)</Where></Query></View>") { continue }
    Seed 'Templates' @{ Title = $t.K; Language = 'en'; Subject = $t.S; Body = $t.B; IsActive = $true } | Out-Null
}
Write-Host "  + $($templates.Count) templates" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 10. TransportRequests - six statuses, one each
# -----------------------------------------------------------------------------

Write-Host "`n=== TransportRequests ===" -ForegroundColor Cyan
Write-Host "  Six requests, deliberately one per status, so every dashboard has" -ForegroundColor DarkGray
Write-Host "  something to show the moment you open it." -ForegroundColor DarkGray

function New-Request {
    param([hashtable] $V, [string] $Note)
    $item = Seed 'TransportRequests' $V
    # RequestId is the prefix plus the SharePoint item ID, zero-padded to four
    # digits, computed on create - exactly what the web part does. A separate
    # counter item would race under concurrent intake.
    $rid = 'FB-{0:D4}' -f $item.Id
    Set-PnPListItem -List 'TransportRequests' -Identity $item.Id -Values @{ Title = $rid; RequestId = $rid } | Out-Null
    Write-Host "  + $rid  $Note" -ForegroundColor Gray
    [pscustomobject]@{ Id = $item.Id; RequestId = $rid }
}

# 1. Submitted - sitting in Pending requests, not yet broadcast
$r1 = New-Request @{
    RequestType = 'Staff'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Dubai Head Office']; DropOffLocation = $LOC['Al Barsha Clinic']
    RequiredDateTime = U $now.AddHours(3); NoOfPax = 1
    VehicleTypeRequested = $VT['Kicks / Creta']; Priority = 'Normal'; ReqStatus = 'Submitted'
    Remarks = 'Routine clinic appointment.'
} 'Submitted      - appears in Pending requests'

# 2. Broadcasting - OPEN OFFERS. This is the one that makes the Driver
#    Dashboard light up. The window is set 30 minutes out rather than the real
#    5 so it does not expire while you are still reading the screen.
$r2 = New-Request @{
    RequestType = 'Client or Client Guest'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Dubai International T3']; DropOffLocation = $LOC['Downtown Client Office']
    RequiredDateTime = U $now.AddHours(2); NoOfPax = 3
    VehicleTypeRequested = $VT['Patrol']; Priority = 'Urgent'; ReqStatus = 'Broadcasting'
    OfferClosesAt = U $now.AddMinutes(30)
    Remarks = 'Client arriving on EK203.'
} 'Broadcasting   - OPEN OFFER, drives the Driver Dashboard'

# 3. Awaiting Coordinator - escalated, with declines to show in the queue
$r3 = New-Request @{
    RequestType = 'Items or Documents'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Jebel Ali Warehouse']; DropOffLocation = $LOC['Mussafah Depot']
    RequiredDateTime = U $now.AddHours(5)
    VehicleTypeRequested = $VT['Cargo Van']; ItemDescription = 'Equipment'
    ItemNotes = 'Two pallets, forklift needed at both ends.'
    Priority = 'Normal'; ReqStatus = 'Awaiting Coordinator'
    OfferClosesAt = U $now.AddMinutes(-20)
} 'Awaiting Coord - ESCALATED, pinned to top with declines'

# 4. Allocated, trip En Route - drives Ongoing trips and Track my ride
$r4 = New-Request @{
    RequestType = 'Staff'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Marina Staff Residence']; DropOffLocation = $LOC['Dubai Head Office']
    RequiredDateTime = U $now.AddMinutes(-40); NoOfPax = 2
    VehicleTypeRequested = $VT['Patrol']; Priority = 'Normal'; ReqStatus = 'Allocated'
} 'Allocated      - trip EN ROUTE, drives Ongoing trips'

# 5. Completed - gives the counters and reports something to total
$r5 = New-Request @{
    RequestType = 'Client or Client Guest'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Downtown Client Office']; DropOffLocation = $LOC['Dubai International T3']
    RequiredDateTime = U $now.AddHours(-6); NoOfPax = 2
    VehicleTypeRequested = $VT['Pathfinder']; Priority = 'Normal'; ReqStatus = 'Completed'
} 'Completed      - feeds counters and every report'

# 6. Cancelled
$r6 = New-Request @{
    RequestType = 'Staff'; Requester = $TesterUpn; Property = $P1
    PickupLocation = $LOC['Dubai Head Office']; DropOffLocation = $LOC['Marina Staff Residence']
    RequiredDateTime = U $now.AddHours(-3); NoOfPax = 1
    VehicleTypeRequested = $VT['Kicks / Creta']; Priority = 'Normal'; ReqStatus = 'Cancelled'
    CancellationReason = 'Meeting moved to Teams.'
} 'Cancelled      - shows the cancelled badge and reason'

# -----------------------------------------------------------------------------
# 11. RideOffers
# -----------------------------------------------------------------------------

Write-Host "`n=== RideOffers ===" -ForegroundColor Cyan

# Request 2: open offers to every eligible passenger-qualified Dubai driver.
# Your own driver record is among them, so the Driver Dashboard shows a card.
foreach ($name in 'Rashid Al Mansoori', 'Suresh Nair') {
    Seed 'RideOffers' @{
        Title = "$($r2.RequestId) - $name"; Request = $r2.Id; Driver = $DRV[$name]
        OfferedOn = U $now.AddMinutes(-1); ResponseStatus = 'Pending'
    } | Out-Null
}
Write-Host "  + 2 PENDING offers on $($r2.RequestId)" -ForegroundColor Gray

# Request 3: everyone declined or let it expire. This is what the Coordinator
# sees in the escalation queue - who was offered it and why they said no.
Seed 'RideOffers' @{
    Title = "$($r3.RequestId) - Mohammed Iqbal"; Request = $r3.Id; Driver = $DRV['Mohammed Iqbal']
    OfferedOn = U $now.AddMinutes(-25); ResponseStatus = 'Rejected'
    RejectionReason = 'Already on a trip'; RespondedOn = U $now.AddMinutes(-24)
} | Out-Null
Seed 'RideOffers' @{
    Title = "$($r3.RequestId) - Arun Prasad"; Request = $r3.Id; Driver = $DRV['Arun Prasad']
    OfferedOn = U $now.AddMinutes(-25); ResponseStatus = 'Rejected'
    RejectionReason = 'Other'; RejectionNotes = 'Forklift licence expired, cannot load pallets.'
    RespondedOn = U $now.AddMinutes(-23)
} | Out-Null
Write-Host "  + 2 REJECTED offers on $($r3.RequestId) (with reasons)" -ForegroundColor Gray

# Requests 4 and 5: accepted, with the other offers closed.
Seed 'RideOffers' @{
    Title = "$($r4.RequestId) - Rashid Al Mansoori"; Request = $r4.Id; Driver = $DRV['Rashid Al Mansoori']
    Vehicle = $VEH['A 12345']; OfferedOn = U $now.AddMinutes(-55)
    ResponseStatus = 'Accepted'; RespondedOn = U $now.AddMinutes(-52)
} | Out-Null
Seed 'RideOffers' @{
    Title = "$($r4.RequestId) - Suresh Nair"; Request = $r4.Id; Driver = $DRV['Suresh Nair']
    OfferedOn = U $now.AddMinutes(-55); ResponseStatus = 'Closed'
} | Out-Null
Seed 'RideOffers' @{
    Title = "$($r5.RequestId) - Suresh Nair"; Request = $r5.Id; Driver = $DRV['Suresh Nair']
    Vehicle = $VEH['B 23456']; OfferedOn = U $now.AddHours(-7)
    ResponseStatus = 'Accepted'; RespondedOn = U $now.AddHours(-7).AddMinutes(2)
} | Out-Null
Write-Host "  + 3 resolved offers on $($r4.RequestId) and $($r5.RequestId)" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 12. TripAssignments
# -----------------------------------------------------------------------------

Write-Host "`n=== TripAssignments ===" -ForegroundColor Cyan

$t4 = Seed 'TripAssignments' @{
    Request = $r4.Id; Driver = $DRV['Rashid Al Mansoori']; Vehicle = $VEH['A 12345']
    AcceptedOn = U $now.AddMinutes(-52); TripStatus = 'En Route'
    EtaMinutes = 12; EtaUpdatedOn = U $now.AddMinutes(-4)
    ActualStartTime = U $now.AddMinutes(-15)
    OpeningReading = 84120
    AllocationMode = 'Automatic'; IsOverride = $false
}
Set-PnPListItem -List 'TripAssignments' -Identity $t4.Id -Values @{
    Title = ('TR-{0:D4}' -f $t4.Id); TripId = ('TR-{0:D4}' -f $t4.Id)
} | Out-Null
Set-PnPListItem -List 'TransportRequests' -Identity $r4.Id -Values @{ AcceptedTrip = $t4.Id } | Out-Null
Write-Host "  + TR-$('{0:D4}' -f $t4.Id)  En Route, ETA 12 min, odometer open at 84120" -ForegroundColor Gray

$t5 = Seed 'TripAssignments' @{
    Request = $r5.Id; Driver = $DRV['Suresh Nair']; Vehicle = $VEH['B 23456']
    AcceptedOn = U $now.AddHours(-7).AddMinutes(2); TripStatus = 'Completed'
    ActualStartTime  = U $now.AddHours(-6).AddMinutes(-20)
    ActualPickupTime = U $now.AddHours(-6)
    ActualDropTime   = U $now.AddHours(-5)
    OpeningReading = 61230; ClosingReading = 61272
    AllocationMode = 'Automatic'; IsOverride = $false
}
Set-PnPListItem -List 'TripAssignments' -Identity $t5.Id -Values @{
    Title = ('TR-{0:D4}' -f $t5.Id); TripId = ('TR-{0:D4}' -f $t5.Id)
} | Out-Null
Set-PnPListItem -List 'TransportRequests' -Identity $r5.Id -Values @{ AcceptedTrip = $t5.Id } | Out-Null
Write-Host "  + TR-$('{0:D4}' -f $t5.Id)  Completed, 42 km" -ForegroundColor Gray

# Driver and vehicle status is DENORMALISED - written by flows, not calculated
# at read time. Set it here to match the trips above or the availability panel
# will disagree with the grid.
Set-PnPListItem -List 'DriversMaster' -Identity $DRV['Rashid Al Mansoori'] -Values @{ Status = 'On Trip' } | Out-Null
Set-PnPListItem -List 'VehicleMaster' -Identity $VEH['A 12345'] -Values @{ VehStatus = 'On Trip' } | Out-Null

# -----------------------------------------------------------------------------
# 13. TripStatusHistory - powers the Track my ride timeline
# -----------------------------------------------------------------------------

Write-Host "`n=== TripStatusHistory ===" -ForegroundColor Cyan
Seed 'TripStatusHistory' @{ Title = 'assigned'; Trip = $t4.Id; FromStatus = ''; ToStatus = 'Assigned';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddMinutes(-52); Comment = 'Accepted from broadcast.' } | Out-Null
Seed 'TripStatusHistory' @{ Title = 'en route'; Trip = $t4.Id; FromStatus = 'Assigned'; ToStatus = 'En Route';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddMinutes(-15); Comment = 'Driver tapped On my way.' } | Out-Null

Seed 'TripStatusHistory' @{ Title = 'assigned'; Trip = $t5.Id; FromStatus = ''; ToStatus = 'Assigned';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddHours(-7); Comment = '' } | Out-Null
Seed 'TripStatusHistory' @{ Title = 'en route'; Trip = $t5.Id; FromStatus = 'Assigned'; ToStatus = 'En Route';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddHours(-6).AddMinutes(-20); Comment = '' } | Out-Null
Seed 'TripStatusHistory' @{ Title = 'in transit'; Trip = $t5.Id; FromStatus = 'En Route'; ToStatus = 'In Transit';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddHours(-6); Comment = '' } | Out-Null
Seed 'TripStatusHistory' @{ Title = 'completed'; Trip = $t5.Id; FromStatus = 'In Transit'; ToStatus = 'Completed';
    ChangedBy = $TesterUpn; ChangedOn = U $now.AddHours(-5); Comment = 'Closing reading 61272.' } | Out-Null
Write-Host "  + 6 transitions across the two trips" -ForegroundColor Gray

# -----------------------------------------------------------------------------
# 14. NotificationLog
# -----------------------------------------------------------------------------

Write-Host "`n=== NotificationLog ===" -ForegroundColor Cyan
Seed 'NotificationLog' @{ Title = 'ack'; RelatedRequest = $r2.Id; Channel = 'Email'; Recipient = $TesterUpn
    TemplateKey = 'RequestSubmitted.Requester'; SentOn = U $now.AddMinutes(-2); Outcome = 'Sent'; RetryCount = 0 } | Out-Null
Seed 'NotificationLog' @{ Title = 'broadcast'; RelatedRequest = $r2.Id; Channel = 'Email'; Recipient = $TesterUpn
    TemplateKey = 'BroadcastOpened.Driver'; SentOn = U $now.AddMinutes(-1); Outcome = 'Sent'; RetryCount = 0 } | Out-Null
# One failure on purpose - a silent failure means a driver never learns they
# have a trip, so the Coordinator alert path needs something to show.
Seed 'NotificationLog' @{ Title = 'failed'; RelatedRequest = $r3.Id; Channel = 'Email'; Recipient = 'arun.prasad@example.com'
    TemplateKey = 'BroadcastOpened.Driver'; SentOn = U $now.AddMinutes(-25); Outcome = 'Failed'; RetryCount = 3
    ErrorMessage = 'SMTP 550 5.1.1 recipient not found' } | Out-Null
Write-Host "  + 3 entries (one FAILED on purpose)" -ForegroundColor Gray

Write-Host "`nSeeding complete." -ForegroundColor Green
Write-Host @"

WHAT YOU SHOULD NOW SEE

  Dispatch Dashboard   Ongoing trips 1, Pending requests 3, Completed 1,
                       Available drivers 3, Available vehicles 5, pairs 1.
                       $($r3.RequestId) pinned to the top in the amber escalation state.

  Driver Dashboard     1 open offer ($($r2.RequestId)) with a live countdown, and
                       1 trip in progress ($($r4.RequestId), En Route) - but only if
                       you signed in as $TesterUpn.

  My Bookings          6 requests, one per status.

  Book a Ride          10 destinations, 8 vehicle types filtered by request type.

  Reports              Run "Trip summary" over this month.

NOTE ON THE OPEN OFFER
  $($r2.RequestId)'s window was set 30 minutes out, not the real 5, so it does not
  expire while you read the screen. Once it passes, re-open it with:
      .\3-Simulate-Flows.ps1 -SiteUrl "$SiteUrl" -Action Broadcast -RequestId $($r2.Id)
"@ -ForegroundColor DarkGray

Disconnect-PnPOnline
