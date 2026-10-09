<#
.SYNOPSIS
    Does by hand what the five Power Automate flows will do, so the web parts
    can be tested before the flows exist.

.DESCRIPTION
    THIS IS A TEST HARNESS, NOT A SUBSTITUTE FOR THE FLOWS.

    It proves the web parts READ and RENDER every state correctly. It does not
    prove the write path, because the write path IS the flows - in particular
    it does not test the optimistic-concurrency claim, which is the single
    most important piece of logic in the solution and can only be tested
    against the real Claim flow with two drivers tapping Accept at once.

    Use it to drive the UI through its states and check each screen.

.PARAMETER Action
    Broadcast  Re-open a request for offers. Sets Status = Broadcasting, writes
               a fresh OfferClosesAt, and creates a Pending RideOffers row for
               every eligible driver.
    Claim      A driver accepts. Creates the TripAssignments row, closes the
               other offers, flips driver and vehicle to On Trip.
    Advance    Move a trip to its next status, writing the timestamp and the
               TripStatusHistory row.
    Expire     Close an open window with nobody accepting - the escalation path.
    Status     Print the current state of everything.

.EXAMPLE
    .\3-Simulate-Flows.ps1 -SiteUrl $url -Action Status
    .\3-Simulate-Flows.ps1 -SiteUrl $url -Action Broadcast -RequestId 2
    .\3-Simulate-Flows.ps1 -SiteUrl $url -Action Claim -RequestId 2 -DriverName "Rashid Al Mansoori"
    .\3-Simulate-Flows.ps1 -SiteUrl $url -Action Advance -TripId 1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Broadcast', 'Claim', 'Advance', 'Expire', 'Status')]
    [string] $Action,

    [int] $RequestId,
    [int] $TripId,
    [string] $DriverName,

    # How long the simulated offer window stays open. The real default is 5
    # minutes; 30 is kinder while you are reading the screen.
    [int] $WindowMinutes = 30,

    [int] $Odometer
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell -DisableNameChecking
Connect-PnPOnline -Url $SiteUrl -Interactive

function U { param([datetime] $d) $d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
$now = (Get-Date).ToUniversalTime()

function Get-Field {
    param($Item, [string] $Name)
    $Item.FieldValues[$Name]
}

switch ($Action) {

    # -------------------------------------------------------------------------
    'Broadcast' {
        if (-not $RequestId) { throw "Broadcast needs -RequestId (the TransportRequests item ID)." }

        $req = Get-PnPListItem -List 'TransportRequests' -Id $RequestId
        $rid = Get-Field $req 'RequestId'
        $typeLookup = Get-Field $req 'VehicleTypeRequested'
        $propLookup = Get-Field $req 'Property'
        $typeId = $typeLookup.LookupId
        $propId = $propLookup.LookupId

        Write-Host "Broadcasting $rid ..." -ForegroundColor Cyan

        # Close anything still open from a previous run.
        $old = Get-PnPListItem -List 'RideOffers' -PageSize 500 | Where-Object {
            (Get-Field $_ 'Request').LookupId -eq $RequestId -and (Get-Field $_ 'ResponseStatus') -eq 'Pending'
        }
        foreach ($o in $old) {
            Set-PnPListItem -List 'RideOffers' -Identity $o.Id -Values @{ ResponseStatus = 'Expired' } | Out-Null
        }

        # Build the eligible driver set - the same checks the Broadcast flow
        # applies: property, vehicle-type qualification, active, not on leave
        # across the window, licence valid through the end of the window, and
        # under the concurrent-trip cap.
        $windowFrom = [datetime](Get-Field $req 'RequiredDateTime')
        $windowTo   = $windowFrom.AddHours(2)

        $blocks = Get-PnPListItem -List 'AvailabilityBlocks' -PageSize 500
        $openTrips = Get-PnPListItem -List 'TripAssignments' -PageSize 500 | Where-Object {
            (Get-Field $_ 'TripStatus') -in 'Assigned', 'En Route', 'In Transit'
        }
        $allocations = Get-PnPListItem -List 'DLClientMaster' -PageSize 500

        $eligible = @()
        foreach ($d in (Get-PnPListItem -List 'DriversMaster' -PageSize 500)) {
            $name   = Get-Field $d 'DriverName'
            $why    = @()

            if (-not (Get-Field $d 'IsActive'))              { $why += 'inactive' }
            if ((Get-Field $d 'Status') -eq 'Inactive')      { $why += 'inactive' }

            $onProperty = $allocations | Where-Object {
                (Get-Field $_ 'Driver').LookupId -eq $d.Id -and (Get-Field $_ 'Property').LookupId -eq $propId
            }
            if (-not $onProperty) { $why += 'not allocated to this property' }

            $quals = @(Get-Field $d 'QualifiedVehicleTypes' | ForEach-Object { $_.LookupId })
            if ($typeId -notin $quals) { $why += 'not qualified for this vehicle type' }

            $blocked = $blocks | Where-Object {
                (Get-Field $_ 'Driver').LookupId -eq $d.Id -and
                (Get-Field $_ 'BlockType') -ne 'Shift' -and
                [datetime](Get-Field $_ 'FromDateTime') -lt $windowTo -and
                [datetime](Get-Field $_ 'ToDateTime')   -gt $windowFrom
            }
            if ($blocked) { $why += 'on leave during the window' }

            $expiry = Get-Field $d 'LicenceExpiry'
            if ($expiry -and [datetime]$expiry -lt $windowTo) { $why += 'licence expires before the trip ends' }

            $held = @($openTrips | Where-Object { (Get-Field $_ 'Driver').LookupId -eq $d.Id }).Count
            $cap  = [int](Get-Field $d 'MaxConcurrentTrips')
            if ($cap -le 0) { $cap = 1 }
            if ($held -ge $cap) { $why += "at the concurrent cap ($held of $cap)" }

            if ($why.Count -eq 0) {
                $eligible += $d
                Write-Host "  eligible: $name" -ForegroundColor Green
            }
            else {
                Write-Host "  excluded: $name - $($why -join '; ')" -ForegroundColor DarkGray
            }
        }

        if ($eligible.Count -eq 0) {
            Write-Warning "No driver is eligible. The real flow would set this straight to Awaiting Coordinator."
        }

        foreach ($d in $eligible) {
            Add-PnPListItem -List 'RideOffers' -Values @{
                Title = "$rid - $(Get-Field $d 'DriverName')"
                Request = $RequestId; Driver = $d.Id
                OfferedOn = U $now; ResponseStatus = 'Pending'
            } | Out-Null
        }

        Set-PnPListItem -List 'TransportRequests' -Identity $RequestId -Values @{
            Status = 'Broadcasting'; OfferClosesAt = U $now.AddMinutes($WindowMinutes)
        } | Out-Null

        Write-Host "$rid is broadcasting to $($eligible.Count) driver(s). Window closes in $WindowMinutes minutes." -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    'Claim' {
        if (-not $RequestId -or -not $DriverName) { throw "Claim needs -RequestId and -DriverName." }

        $req = Get-PnPListItem -List 'TransportRequests' -Id $RequestId
        $rid = Get-Field $req 'RequestId'

        # The real flow refuses here if the status has already moved - that is
        # the whole point of first-accept-wins.
        if ((Get-Field $req 'Status') -ne 'Broadcasting') {
            throw "$rid is '$(Get-Field $req 'Status')', not Broadcasting. The real Claim flow would answer 'this trip has already been taken'."
        }

        $driver = Get-PnPListItem -List 'DriversMaster' -PageSize 500 |
                  Where-Object { (Get-Field $_ 'DriverName') -eq $DriverName } | Select-Object -First 1
        if (-not $driver) { throw "No driver called '$DriverName'." }

        $typeId = (Get-Field $req 'VehicleTypeRequested').LookupId
        $propId = (Get-Field $req 'Property').LookupId

        $busy = @(Get-PnPListItem -List 'TripAssignments' -PageSize 500 | Where-Object {
            (Get-Field $_ 'TripStatus') -in 'Assigned', 'En Route', 'In Transit'
        } | ForEach-Object { (Get-Field $_ 'Vehicle').LookupId })

        # Prefer this driver's default vehicle if it is free, else the first
        # free vehicle of the right type at that property.
        $candidates = Get-PnPListItem -List 'VehicleMaster' -PageSize 500 | Where-Object {
            (Get-Field $_ 'VehicleType').LookupId -eq $typeId -and
            (Get-Field $_ 'Property').LookupId -eq $propId -and
            (Get-Field $_ 'Status') -notin 'Retired', 'Under Maintenance' -and
            $_.Id -notin $busy
        }
        $vehicle = $candidates | Where-Object { (Get-Field $_ 'DefaultDriver').LookupId -eq $driver.Id } | Select-Object -First 1
        if (-not $vehicle) { $vehicle = $candidates | Select-Object -First 1 }
        if (-not $vehicle) { throw "No free vehicle of that type at that property. The real flow would refuse the claim." }

        $trip = Add-PnPListItem -List 'TripAssignments' -Values @{
            Request = $RequestId; Driver = $driver.Id; Vehicle = $vehicle.Id
            AcceptedOn = U $now; TripStatus = 'Assigned'
            AllocationMode = 'Automatic'; IsOverride = $false
        }
        $tid = 'TR-{0:D4}' -f $trip.Id
        Set-PnPListItem -List 'TripAssignments' -Identity $trip.Id -Values @{ Title = $tid; TripId = $tid } | Out-Null

        # The accepting driver's offer is Accepted; every other open offer is
        # Closed, not Expired - they lost the race, the window did not run out.
        foreach ($o in (Get-PnPListItem -List 'RideOffers' -PageSize 500 | Where-Object { (Get-Field $_ 'Request').LookupId -eq $RequestId })) {
            if ((Get-Field $o 'ResponseStatus') -ne 'Pending') { continue }
            if ((Get-Field $o 'Driver').LookupId -eq $driver.Id) {
                Set-PnPListItem -List 'RideOffers' -Identity $o.Id -Values @{
                    ResponseStatus = 'Accepted'; RespondedOn = U $now; Vehicle = $vehicle.Id
                } | Out-Null
            }
            else {
                Set-PnPListItem -List 'RideOffers' -Identity $o.Id -Values @{ ResponseStatus = 'Closed' } | Out-Null
            }
        }

        Set-PnPListItem -List 'TransportRequests' -Identity $RequestId -Values @{
            Status = 'Allocated'; AcceptedTrip = $trip.Id
        } | Out-Null

        Set-PnPListItem -List 'DriversMaster' -Identity $driver.Id -Values @{ Status = 'On Trip' } | Out-Null
        Set-PnPListItem -List 'VehicleMaster' -Identity $vehicle.Id -Values @{ Status = 'On Trip' } | Out-Null

        Add-PnPListItem -List 'TripStatusHistory' -Values @{
            Title = 'assigned'; Trip = $trip.Id; FromStatus = ''; ToStatus = 'Assigned'
            ChangedBy = (Get-PnPContext).Web.CurrentUser.LoginName
            ChangedOn = U $now; Comment = 'Accepted from broadcast (simulated).'
        } | Out-Null

        Write-Host "$DriverName won $rid. Trip $tid created on vehicle $(Get-Field $vehicle 'PlateNo')." -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    'Advance' {
        if (-not $TripId) { throw "Advance needs -TripId (the TripAssignments item ID)." }

        $trip = Get-PnPListItem -List 'TripAssignments' -Id $TripId
        $from = Get-Field $trip 'TripStatus'
        $next = switch ($from) {
            'Assigned'   { 'En Route' }
            'En Route'   { 'In Transit' }
            'In Transit' { 'Completed' }
            default      { $null }
        }
        if (-not $next) { throw "A trip that is '$from' has nowhere to advance to." }

        $values = @{ TripStatus = $next }
        switch ($next) {
            'En Route'   {
                $values['ActualStartTime'] = U $now
                if ($Odometer) { $values['OpeningReading'] = $Odometer }
                else { Write-Warning "No -Odometer given. The real flow requires the opening reading at this tap." }
            }
            'In Transit' { $values['ActualPickupTime'] = U $now }
            'Completed'  {
                $values['ActualDropTime'] = U $now
                if ($Odometer) { $values['ClosingReading'] = $Odometer }
                else { Write-Warning "No -Odometer given. The real flow requires the closing reading at this tap." }
            }
        }
        Set-PnPListItem -List 'TripAssignments' -Identity $TripId -Values $values | Out-Null

        Add-PnPListItem -List 'TripStatusHistory' -Values @{
            Title = $next.ToLower(); Trip = $TripId; FromStatus = $from; ToStatus = $next
            ChangedBy = (Get-PnPContext).Web.CurrentUser.LoginName
            ChangedOn = U $now; Comment = 'Simulated tap.'
        } | Out-Null

        if ($next -eq 'Completed') {
            # The Complete flow releases the driver and the vehicle and writes
            # the Daily Trip Sheet row. The trip sheet write is NOT simulated.
            $requestId = (Get-Field $trip 'Request').LookupId
            Set-PnPListItem -List 'TransportRequests' -Identity $requestId -Values @{ Status = 'Completed' } | Out-Null
            Set-PnPListItem -List 'DriversMaster' -Identity (Get-Field $trip 'Driver').LookupId  -Values @{ Status = 'Available' } | Out-Null
            Set-PnPListItem -List 'VehicleMaster' -Identity (Get-Field $trip 'Vehicle').LookupId -Values @{ Status = 'Available' } | Out-Null
            Write-Host "Driver and vehicle released. NOTE: the Daily Trip Sheet row is not simulated." -ForegroundColor DarkGray
        }

        Write-Host "Trip $(Get-Field $trip 'TripId'): $from -> $next" -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    'Expire' {
        if (-not $RequestId) { throw "Expire needs -RequestId." }

        $req = Get-PnPListItem -List 'TransportRequests' -Id $RequestId
        foreach ($o in (Get-PnPListItem -List 'RideOffers' -PageSize 500 | Where-Object {
            (Get-Field $_ 'Request').LookupId -eq $RequestId -and (Get-Field $_ 'ResponseStatus') -eq 'Pending'
        })) {
            Set-PnPListItem -List 'RideOffers' -Identity $o.Id -Values @{ ResponseStatus = 'Expired' } | Out-Null
        }
        Set-PnPListItem -List 'TransportRequests' -Identity $RequestId -Values @{
            Status = 'Awaiting Coordinator'; OfferClosesAt = U $now.AddMinutes(-1)
        } | Out-Null

        Write-Host "$(Get-Field $req 'RequestId') escalated to the Coordinator. It should now be pinned to the top of Pending requests." -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    'Status' {
        Write-Host "`n=== Requests ===" -ForegroundColor Cyan
        foreach ($r in (Get-PnPListItem -List 'TransportRequests' -PageSize 500)) {
            $closes = Get-Field $r 'OfferClosesAt'
            $left = if ($closes) { [math]::Round(([datetime]$closes - $now).TotalMinutes) } else { $null }
            $suffix = if ($null -ne $left) { "  (window $left min)" } else { '' }
            '{0,-4} {1,-10} {2,-22}{3}' -f $r.Id, (Get-Field $r 'RequestId'), (Get-Field $r 'Status'), $suffix | Write-Host
        }

        Write-Host "`n=== Trips ===" -ForegroundColor Cyan
        foreach ($t in (Get-PnPListItem -List 'TripAssignments' -PageSize 500)) {
            '{0,-4} {1,-10} {2,-12} {3,-22} {4}' -f $t.Id, (Get-Field $t 'TripId'), (Get-Field $t 'TripStatus'),
                (Get-Field $t 'Driver').LookupValue, (Get-Field $t 'Vehicle').LookupValue | Write-Host
        }

        Write-Host "`n=== Open offers ===" -ForegroundColor Cyan
        $open = Get-PnPListItem -List 'RideOffers' -PageSize 500 | Where-Object { (Get-Field $_ 'ResponseStatus') -eq 'Pending' }
        if (-not $open) { Write-Host "  none" -ForegroundColor DarkGray }
        foreach ($o in $open) {
            '{0,-4} {1,-22} {2}' -f $o.Id, (Get-Field $o 'Driver').LookupValue, (Get-Field $o 'Request').LookupValue | Write-Host
        }

        Write-Host "`n=== Drivers ===" -ForegroundColor Cyan
        foreach ($d in (Get-PnPListItem -List 'DriversMaster' -PageSize 500)) {
            '{0,-4} {1,-22} {2,-12} cap {3}' -f $d.Id, (Get-Field $d 'DriverName'), (Get-Field $d 'Status'),
                (Get-Field $d 'MaxConcurrentTrips') | Write-Host
        }

        Write-Host "`n=== Vehicles ===" -ForegroundColor Cyan
        foreach ($v in (Get-PnPListItem -List 'VehicleMaster' -PageSize 500)) {
            '{0,-4} {1,-10} {2,-18} {3}' -f $v.Id, (Get-Field $v 'PlateNo'),
                (Get-Field $v 'VehicleType').LookupValue, (Get-Field $v 'Status') | Write-Host
        }
        Write-Host ""
    }
}

Disconnect-PnPOnline
