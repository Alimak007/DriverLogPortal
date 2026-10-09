<#
.SYNOPSIS
    Creates every list the Fleet Booking and Dispatch Portal needs.

.DESCRIPTION
    Builds the lists, columns, choice values and indexes described in
    SharePoint-Lists-Specification.txt, plus the five SharePoint groups.

    SAFE TO RE-RUN. Every list and column is created only if it is missing, so
    running it twice changes nothing.

    WHAT IT DOES NOT DO
      - It does not touch DriverLogs, LogEntries, DrvierEntries, FleetData or
        Daily Trip Sheet. Those are left strictly alone.
      - It does not migrate an existing VehicleMaster or DriversMaster. On a
        clean dev site it creates them fresh. If those lists already hold real
        data, STOP and use the migration in section 3 of the specification
        instead - this script will add the new columns but will not re-key
        Title or populate anything.
      - It does not set list permissions. See the note at the end.

.EXAMPLE
    .\1-Provision-Lists.ps1 -SiteUrl "https://pomae.sharepoint.com/sites/travel-dev"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    # Your own account. Added to all five groups so one person can test every
    # role. NEVER do this in production.
    [string] $TesterUpn
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -DisableNameChecking

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

function Ensure-List {
    param([string] $Title)
    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  list exists: $Title" -ForegroundColor DarkGray
        return $existing
    }
    Write-Host "  creating list: $Title" -ForegroundColor Green
    New-PnPList -Title $Title -Template GenericList -OnQuickLaunch:$false
}

function Field-Exists {
    param([string] $List, [string] $Name)
    $null -ne (Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue)
}

function Ensure-Field {
    param(
        [string] $List,
        [string] $Name,
        [string] $Type,                 # Text | Number | Boolean | DateTime | Note | User | Choice | MultiChoice
        [string[]] $Choices,
        [switch] $Required,
        [switch] $Indexed,
        [switch] $DateOnly,
        [string] $Default
    )
    if (Field-Exists -List $List -Name $Name) {
        Write-Host "    field exists: $Name" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    + $Name ($Type)" -ForegroundColor Gray
        $req = if ($Required) { 'TRUE' } else { 'FALSE' }

        switch ($Type) {
            'Choice' {
                $opts = ($Choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
                $def  = if ($Default) { "<Default>$Default</Default>" } else { '' }
                $xml  = "<Field Type='Choice' DisplayName='$Name' Name='$Name' StaticName='$Name' Required='$req' Format='Dropdown'><CHOICES>$opts</CHOICES>$def</Field>"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'MultiChoice' {
                $opts = ($Choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
                $xml  = "<Field Type='MultiChoice' DisplayName='$Name' Name='$Name' StaticName='$Name' Required='$req'><CHOICES>$opts</CHOICES></Field>"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'DateTime' {
                $fmt = if ($DateOnly) { 'DateOnly' } else { 'DateTime' }
                $xml = "<Field Type='DateTime' DisplayName='$Name' Name='$Name' StaticName='$Name' Format='$fmt' Required='$req' />"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'Note' {
                $xml = "<Field Type='Note' DisplayName='$Name' Name='$Name' StaticName='$Name' NumLines='4' RichText='FALSE' Required='$req' />"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'Boolean' {
                $def = if ($Default) { "<Default>$Default</Default>" } else { '<Default>0</Default>' }
                $xml = "<Field Type='Boolean' DisplayName='$Name' Name='$Name' StaticName='$Name'>$def</Field>"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'Number' {
                $def = if ($Default) { "<Default>$Default</Default>" } else { '' }
                $xml = "<Field Type='Number' DisplayName='$Name' Name='$Name' StaticName='$Name' Required='$req' Decimals='0'>$def</Field>"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            'User' {
                $xml = "<Field Type='User' DisplayName='$Name' Name='$Name' StaticName='$Name' Required='$req' UserSelectionMode='PeopleOnly' />"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
            default {
                $xml = "<Field Type='Text' DisplayName='$Name' Name='$Name' StaticName='$Name' Required='$req' MaxLength='255' />"
                Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
            }
        }
    }

    if ($Indexed) {
        try { Set-PnPField -List $List -Identity $Name -Values @{ Indexed = $true } -ErrorAction Stop | Out-Null }
        catch { Write-Warning "    could not index $Name on $List : $($_.Exception.Message)" }
    }
}

function Ensure-Lookup {
    param(
        [string] $List,
        [string] $Name,
        [string] $TargetList,
        [string] $ShowField = 'Title',
        [switch] $Required,
        [switch] $Multi,
        [switch] $Indexed
    )
    if (Field-Exists -List $List -Name $Name) {
        Write-Host "    field exists: $Name" -ForegroundColor DarkGray
    }
    else {
        $target = Get-PnPList -Identity $TargetList
        if (-not $target) { throw "Lookup target list '$TargetList' does not exist. Check the creation order." }

        Write-Host "    + $Name (lookup -> $TargetList)" -ForegroundColor Gray
        $req  = if ($Required) { 'TRUE' } else { 'FALSE' }
        $type = if ($Multi) { 'LookupMulti' } else { 'Lookup' }
        $mult = if ($Multi) { "Mult='TRUE'" } else { '' }
        $xml  = "<Field Type='$type' DisplayName='$Name' Name='$Name' StaticName='$Name' List='{$($target.Id)}' ShowField='$ShowField' Required='$req' $mult />"
        Add-PnPFieldFromXml -List $List -FieldXml $xml | Out-Null
    }

    # A multi-value lookup cannot be indexed; SharePoint refuses it.
    if ($Indexed -and -not $Multi) {
        try { Set-PnPField -List $List -Identity $Name -Values @{ Indexed = $true } -ErrorAction Stop | Out-Null }
        catch { Write-Warning "    could not index $Name on $List : $($_.Exception.Message)" }
    }
}

# -----------------------------------------------------------------------------
# Create in dependency order - section 4 of the specification
# -----------------------------------------------------------------------------

Write-Host "`n=== Master data ===" -ForegroundColor Cyan

# --- PropertyMaster (reused unchanged; created only if the dev site lacks it)
Ensure-List 'PropertyMaster' | Out-Null

# --- DLClientMaster (reused unchanged; driver-to-property allocation)
Ensure-List 'DLClientMaster' | Out-Null
Ensure-Lookup -List 'DLClientMaster' -Name 'PropertyRef' -TargetList 'PropertyMaster' -Indexed

# --- VehicleTypes
Ensure-List 'VehicleTypes' | Out-Null
Ensure-Field  -List 'VehicleTypes' -Name 'SuitableFor' -Type MultiChoice -Required `
              -Choices 'Client or Client Guest', 'Staff', 'Items or Documents'
Ensure-Field  -List 'VehicleTypes' -Name 'Capacity'  -Type Number
Ensure-Field  -List 'VehicleTypes' -Name 'SortOrder' -Type Number
Ensure-Field  -List 'VehicleTypes' -Name 'IsActive'  -Type Boolean -Default 1 -Indexed

# --- Locations
Ensure-List 'Locations' | Out-Null
Ensure-Lookup -List 'Locations' -Name 'Property' -TargetList 'PropertyMaster' -Indexed
Ensure-Field  -List 'Locations' -Name 'Address'   -Type Note
Ensure-Field  -List 'Locations' -Name 'Latitude'  -Type Number
Ensure-Field  -List 'Locations' -Name 'Longitude' -Type Number
Ensure-Field  -List 'Locations' -Name 'Category'  -Type Choice `
              -Choices 'Office', 'Residence', 'Airport', 'Hospital', 'Hotel', 'Warehouse', 'Other'
Ensure-Field  -List 'Locations' -Name 'IsActive'  -Type Boolean -Default 1 -Indexed

# --- DriversMaster
Ensure-List 'DriversMaster' | Out-Null
# The driver's name lives in Title - there is NO DriverName column (v2 spec s.12).
Ensure-Field  -List 'DriversMaster' -Name 'UserAccount'   -Type User -Required -Indexed
Ensure-Field  -List 'DriversMaster' -Name 'DriverID'      -Type Text -Indexed
Ensure-Field  -List 'DriversMaster' -Name 'Phone'         -Type Text
Ensure-Field  -List 'DriversMaster' -Name 'LicenceNumber' -Type Text
Ensure-Field  -List 'DriversMaster' -Name 'LicenceExpiry' -Type DateTime -DateOnly
Ensure-Lookup -List 'DriversMaster' -Name 'QualifiedVehicleTypes' -TargetList 'VehicleTypes' -Multi
Ensure-Lookup -List 'DriversMaster' -Name 'HomeProperty' -TargetList 'PropertyMaster' -Indexed
Ensure-Field  -List 'DriversMaster' -Name 'Status' -Type Choice -Required -Default 'Available' -Indexed `
              -Choices 'Available', 'On Trip', 'On Leave', 'Inactive'
Ensure-Field  -List 'DriversMaster' -Name 'MaxConcurrentTrips' -Type Number -Default 1
Ensure-Field  -List 'DriversMaster' -Name 'IsActive' -Type Boolean -Default 1 -Indexed

# DLClientMaster's driver lookup needs DriversMaster to exist first.
Ensure-Lookup -List 'DLClientMaster' -Name 'DriverID' -TargetList 'DriversMaster' -Indexed

# --- VehicleMaster
Ensure-List 'VehicleMaster' | Out-Null
Ensure-Field  -List 'VehicleMaster' -Name 'PlateNo' -Type Text -Required -Indexed
Ensure-Lookup -List 'VehicleMaster' -Name 'VehicleType' -TargetList 'VehicleTypes' -Required -Indexed
Ensure-Field  -List 'VehicleMaster' -Name 'Make'   -Type Text
Ensure-Field  -List 'VehicleMaster' -Name 'Model'  -Type Text
Ensure-Field  -List 'VehicleMaster' -Name 'Colour' -Type Text
Ensure-Field  -List 'VehicleMaster' -Name 'Capacity' -Type Number
Ensure-Lookup -List 'VehicleMaster' -Name 'DefaultDriver' -TargetList 'DriversMaster'
# PropertyRef, NOT Property: a plain-text Property column already exists and
# the logbook joins on it. SharePoint will not allow a duplicate internal name.
Ensure-Lookup -List 'VehicleMaster' -Name 'PropertyRef' -TargetList 'PropertyMaster' -Indexed
Ensure-Field  -List 'VehicleMaster' -Name 'VehStatus' -Type Choice -Required -Default 'Available' -Indexed `
              -Choices 'Available', 'On Trip', 'Under Maintenance', 'Retired'

# --- AllocationRules
Ensure-List 'AllocationRules' | Out-Null
Ensure-Lookup -List 'AllocationRules' -Name 'Property' -TargetList 'PropertyMaster' -Required -Indexed
Ensure-Field  -List 'AllocationRules' -Name 'RequestType' -Type Choice -Required `
              -Choices 'Client or Client Guest', 'Staff', 'Items or Documents'
Ensure-Field  -List 'AllocationRules' -Name 'Mode' -Type Choice -Required -Default 'Automatic' `
              -Choices 'Automatic', 'Manual'
Ensure-Field  -List 'AllocationRules' -Name 'OfferWindowMinutes'          -Type Number -Required -Default 5
Ensure-Field  -List 'AllocationRules' -Name 'MaxConcurrentTripsPerDriver' -Type Number -Required -Default 1
Ensure-Field  -List 'AllocationRules' -Name 'AdvanceReleaseHours'         -Type Number -Required -Default 4
Ensure-Field  -List 'AllocationRules' -Name 'IsActive'        -Type Boolean -Default 1 -Indexed
Ensure-Field  -List 'AllocationRules' -Name 'NotifyEnRoute'   -Type Boolean -Default 1
Ensure-Field  -List 'AllocationRules' -Name 'NotifyInTransit' -Type Boolean -Default 1

Write-Host "`n=== Dispatch ===" -ForegroundColor Cyan

# --- TransportRequests
Ensure-List 'TransportRequests' | Out-Null
Ensure-Field  -List 'TransportRequests' -Name 'RequestId'   -Type Text -Indexed
Ensure-Field  -List 'TransportRequests' -Name 'RequestType' -Type Choice -Required `
              -Choices 'Client or Client Guest', 'Staff', 'Items or Documents'
Ensure-Field  -List 'TransportRequests' -Name 'Requester' -Type User -Required -Indexed
Ensure-Lookup -List 'TransportRequests' -Name 'Property'             -TargetList 'PropertyMaster' -Required -Indexed
Ensure-Lookup -List 'TransportRequests' -Name 'PickupLocation'       -TargetList 'Locations' -Required
Ensure-Lookup -List 'TransportRequests' -Name 'DropOffLocation'      -TargetList 'Locations' -Required
Ensure-Lookup -List 'TransportRequests' -Name 'VehicleTypeRequested' -TargetList 'VehicleTypes' -Required
Ensure-Field  -List 'TransportRequests' -Name 'RequiredDateTime' -Type DateTime -Required -Indexed
Ensure-Field  -List 'TransportRequests' -Name 'ReturnDateTime'   -Type DateTime
Ensure-Field  -List 'TransportRequests' -Name 'NoOfPax'          -Type Number
Ensure-Field  -List 'TransportRequests' -Name 'OtherVehicleText' -Type Text
Ensure-Field  -List 'TransportRequests' -Name 'ItemDescription'  -Type Choice `
              -Choices 'Documents', 'Parcel', 'Equipment', 'Medical', 'Food', 'Other'
Ensure-Field  -List 'TransportRequests' -Name 'ItemNotes' -Type Note
Ensure-Field  -List 'TransportRequests' -Name 'Priority' -Type Choice -Required -Default 'Normal' `
              -Choices 'Normal', 'Urgent'
# Display name is 'Status'; the INTERNAL name is ReqStatus, to stay distinct
# from TripAssignments.TripStatus. See the v2 spec, section 7.
Ensure-Field  -List 'TransportRequests' -Name 'ReqStatus' -Type Choice -Required -Default 'Submitted' -Indexed `
              -Choices 'Submitted', 'Held', 'Broadcasting', 'Awaiting Coordinator', 'Allocated', 'Completed', 'Rejected', 'Cancelled'
Ensure-Field  -List 'TransportRequests' -Name 'OfferClosesAt' -Type DateTime -Indexed
Ensure-Field  -List 'TransportRequests' -Name 'ReleaseAt'     -Type DateTime -Indexed
Ensure-Field  -List 'TransportRequests' -Name 'Remarks'            -Type Note
Ensure-Field  -List 'TransportRequests' -Name 'RejectionReason'    -Type Note
Ensure-Field  -List 'TransportRequests' -Name 'CancellationReason' -Type Note

# --- TripAssignments (before RideOffers, which has no dependency on it, and
#     before TransportRequests.AcceptedTrip, which does)
Ensure-List 'TripAssignments' | Out-Null
Ensure-Field  -List 'TripAssignments' -Name 'TripId' -Type Text -Indexed
Ensure-Lookup -List 'TripAssignments' -Name 'Request' -TargetList 'TransportRequests' -Required -Indexed
Ensure-Lookup -List 'TripAssignments' -Name 'Driver'  -TargetList 'DriversMaster'     -Required -Indexed
Ensure-Lookup -List 'TripAssignments' -Name 'Vehicle' -TargetList 'VehicleMaster'     -Required -Indexed
Ensure-Field  -List 'TripAssignments' -Name 'AcceptedOn' -Type DateTime -Required
Ensure-Field  -List 'TripAssignments' -Name 'TripStatus' -Type Choice -Required -Default 'Assigned' -Indexed `
              -Choices 'Assigned', 'En Route', 'In Transit', 'Completed', 'Cancelled'
Ensure-Field  -List 'TripAssignments' -Name 'EtaMinutes'       -Type Number
Ensure-Field  -List 'TripAssignments' -Name 'EtaUpdatedOn'     -Type DateTime
Ensure-Field  -List 'TripAssignments' -Name 'ActualStartTime'  -Type DateTime
Ensure-Field  -List 'TripAssignments' -Name 'ActualPickupTime' -Type DateTime
Ensure-Field  -List 'TripAssignments' -Name 'ActualDropTime'   -Type DateTime
Ensure-Field  -List 'TripAssignments' -Name 'OpeningReading'   -Type Number
Ensure-Field  -List 'TripAssignments' -Name 'ClosingReading'   -Type Number
Ensure-Field  -List 'TripAssignments' -Name 'AllocationMode' -Type Choice -Required -Default 'Automatic' `
              -Choices 'Automatic', 'Manual'
Ensure-Field  -List 'TripAssignments' -Name 'AllocatedBy'    -Type User
Ensure-Field  -List 'TripAssignments' -Name 'IsOverride'     -Type Boolean -Default 0
Ensure-Field  -List 'TripAssignments' -Name 'OverrideReason' -Type Note
Ensure-Field  -List 'TripAssignments' -Name 'CancellationReason' -Type Note

Ensure-Lookup -List 'TransportRequests' -Name 'AcceptedTrip' -TargetList 'TripAssignments'

# --- RideOffers
Ensure-List 'RideOffers' | Out-Null
Ensure-Lookup -List 'RideOffers' -Name 'Request' -TargetList 'TransportRequests' -Required -Indexed
Ensure-Lookup -List 'RideOffers' -Name 'Driver'  -TargetList 'DriversMaster'     -Required -Indexed
Ensure-Lookup -List 'RideOffers' -Name 'Vehicle' -TargetList 'VehicleMaster'
Ensure-Field  -List 'RideOffers' -Name 'OfferedOn' -Type DateTime -Required -Indexed
Ensure-Field  -List 'RideOffers' -Name 'ResponseStatus' -Type Choice -Required -Default 'Pending' -Indexed `
              -Choices 'Pending', 'Accepted', 'Rejected', 'Expired', 'Closed', 'Withdrawn'
Ensure-Field  -List 'RideOffers' -Name 'RejectionReason' -Type Choice `
              -Choices 'Already on a trip', 'Vehicle unavailable', 'Outside my shift', 'Personal reason', 'Other'
Ensure-Field  -List 'RideOffers' -Name 'RejectionNotes' -Type Note
Ensure-Field  -List 'RideOffers' -Name 'RespondedOn'    -Type DateTime

# --- TripStatusHistory
Ensure-List 'TripStatusHistory' | Out-Null
Ensure-Lookup -List 'TripStatusHistory' -Name 'Trip' -TargetList 'TripAssignments' -Required -Indexed
Ensure-Field  -List 'TripStatusHistory' -Name 'FromStatus' -Type Text
Ensure-Field  -List 'TripStatusHistory' -Name 'ToStatus'   -Type Text -Required
Ensure-Field  -List 'TripStatusHistory' -Name 'ChangedBy'  -Type User -Required
Ensure-Field  -List 'TripStatusHistory' -Name 'ChangedOn'  -Type DateTime -Required -Indexed
Ensure-Field  -List 'TripStatusHistory' -Name 'Comment'    -Type Note

# --- AvailabilityBlocks
Ensure-List 'AvailabilityBlocks' | Out-Null
Ensure-Lookup -List 'AvailabilityBlocks' -Name 'Driver'   -TargetList 'DriversMaster'  -Indexed
Ensure-Lookup -List 'AvailabilityBlocks' -Name 'Vehicle'  -TargetList 'VehicleMaster'  -Indexed
Ensure-Lookup -List 'AvailabilityBlocks' -Name 'Property' -TargetList 'PropertyMaster'
Ensure-Field  -List 'AvailabilityBlocks' -Name 'FromDateTime' -Type DateTime -Required -Indexed
Ensure-Field  -List 'AvailabilityBlocks' -Name 'ToDateTime'   -Type DateTime -Required -Indexed
Ensure-Field  -List 'AvailabilityBlocks' -Name 'BlockType' -Type Choice -Required `
              -Choices 'Leave', 'Shift', 'Maintenance', 'Other'
Ensure-Field  -List 'AvailabilityBlocks' -Name 'Notes' -Type Note

# --- NotificationLog
Ensure-List 'NotificationLog' | Out-Null
Ensure-Lookup -List 'NotificationLog' -Name 'RelatedRequest' -TargetList 'TransportRequests' -Indexed
Ensure-Lookup -List 'NotificationLog' -Name 'RelatedTrip'    -TargetList 'TripAssignments'   -Indexed
Ensure-Field  -List 'NotificationLog' -Name 'Channel'   -Type Choice -Required -Default 'Email' -Choices 'Email'
Ensure-Field  -List 'NotificationLog' -Name 'Recipient'    -Type Text -Required
Ensure-Field  -List 'NotificationLog' -Name 'TemplateKey'  -Type Text -Required
Ensure-Field  -List 'NotificationLog' -Name 'SentOn'       -Type DateTime -Required -Indexed
Ensure-Field  -List 'NotificationLog' -Name 'Outcome' -Type Choice -Required -Default 'Sent' -Indexed `
              -Choices 'Sent', 'Failed', 'Retrying'
Ensure-Field  -List 'NotificationLog' -Name 'ErrorMessage' -Type Note
Ensure-Field  -List 'NotificationLog' -Name 'RetryCount'   -Type Number -Default 0

# --- Templates
Ensure-List 'Templates' | Out-Null
Ensure-Field -List 'Templates' -Name 'Language' -Type Choice -Required -Default 'en' -Indexed -Choices 'en', 'ar'
Ensure-Field -List 'Templates' -Name 'Subject'  -Type Text -Required
Ensure-Field -List 'Templates' -Name 'Body'     -Type Note -Required
Ensure-Field -List 'Templates' -Name 'IsActive' -Type Boolean -Default 1 -Indexed

# --- ReportSnapshot
Ensure-List 'ReportSnapshot' | Out-Null
Ensure-Field  -List 'ReportSnapshot' -Name 'ReportKey' -Type Choice -Required -Indexed `
              -Choices 'TripSummary', 'VehicleUtilisation', 'DriverUtilisation', 'RequestFulfilment', `
                       'BroadcastEffectiveness', 'DriverResponseBehaviour', `
                       'RejectionCancellationAnalysis', 'DemandByTimeAndRoute'
Ensure-Field  -List 'ReportSnapshot' -Name 'PeriodStart' -Type DateTime -DateOnly -Required -Indexed
Ensure-Field  -List 'ReportSnapshot' -Name 'PeriodEnd'   -Type DateTime -DateOnly -Required
Ensure-Lookup -List 'ReportSnapshot' -Name 'Property' -TargetList 'PropertyMaster' -Indexed
Ensure-Field  -List 'ReportSnapshot' -Name 'DimensionKey'  -Type Text
Ensure-Field  -List 'ReportSnapshot' -Name 'DimensionName' -Type Text
Ensure-Field  -List 'ReportSnapshot' -Name 'MetricsJson'   -Type Note -Required
Ensure-Field  -List 'ReportSnapshot' -Name 'GeneratedOn'   -Type DateTime -Required

# -----------------------------------------------------------------------------
# Groups
# -----------------------------------------------------------------------------

Write-Host "`n=== Groups ===" -ForegroundColor Cyan

$groups = 'FM Requesters', 'FM Drivers', 'FM Coordinators', 'FM Admins', 'FM Viewers'
foreach ($g in $groups) {
    if (Get-PnPGroup -Identity $g -ErrorAction SilentlyContinue) {
        Write-Host "  group exists: $g" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  creating group: $g" -ForegroundColor Green
        New-PnPGroup -Title $g | Out-Null
    }
}

if ($TesterUpn) {
    Write-Host "`n  Adding $TesterUpn to ALL FIVE groups so one account can test every role." -ForegroundColor Yellow
    Write-Host "  This is a TEST-ONLY shortcut. In production nobody holds all five." -ForegroundColor Yellow
    foreach ($g in $groups) {
        try   { Add-PnPGroupMember -LoginName $TesterUpn -Identity $g -ErrorAction Stop | Out-Null }
        catch { Write-Warning "  could not add to $g : $($_.Exception.Message)" }
    }
}

Write-Host "`nProvisioning complete." -ForegroundColor Green
Write-Host @"

NOT DONE BY THIS SCRIPT, on purpose:

  * List permissions. Every group currently inherits the site's permissions.
    The permission table is section 5 of SharePoint-Lists-Specification.txt.
    Apply it before the solution carries real data - until then the web parts
    will show a Coordinator and a Requester the same rows, which is exactly
    the failure mode the specification warns about.

  * Item-level permissions on TransportRequests ("read and edit items created
    by the user"). Set this in list settings -> Advanced settings.

  * Archiving and the nightly reconciliation flow.

NEXT: run 2-Seed-TestData.ps1 to load the dummy data.
"@ -ForegroundColor DarkGray

Disconnect-PnPOnline
