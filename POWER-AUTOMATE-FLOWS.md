# Power Automate flows — build guide

Written for: whoever builds the flows. Follow it top to bottom; each flow is self-contained once the shared setup in section 2 is done.

The web parts already call these. The request and response shapes in this document are **taken from the shipped code**, not invented — if a flow's response differs, the web part breaks. Where that matters I've said so explicitly.

---

## 1. The flows at a glance

| # | Flow | Trigger | On what | Priority |
| --- | --- | --- | --- | --- |
| 1 | **Broadcast** | When an item is created | `TransportRequests` | Build first |
| 2 | **Claim** | When an HTTP request is received | Driver taps Accept | Build second — hardest |
| 3 | **Expire** | Recurrence, every 1 minute | `TransportRequests` | Build third |
| 4 | **RespondToOffer** | When an HTTP request is received | Driver taps Reject | |
| 5 | **UpdateTripStatus** | When an HTTP request is received | The four driver taps | |
| 6 | **ManualAssign** | When an HTTP request is received | Coordinator assigns/reassigns | |
| 7 | **Cancel** | When an HTTP request is received | Anyone cancels | |
| 8 | **Reject** | When an HTTP request is received | Coordinator rejects | |
| 9 | **SetEta** | When an HTTP request is received | ETA entered by hand | |
| 10 | **Notify** | Child flow (called by 1–9) | Every material event | Build alongside 1 |
| 11 | **Complete** | Child flow (called by 5) | Trip set to Completed | |
| 12 | **ReleaseAdvanceBookings** | Recurrence, every 15 min | `TransportRequests` (Held) | Before go-live |
| 13 | **NightlyReconciliation** | Recurrence, daily 02:00 | Driver/vehicle status drift | Before go-live |
| 14 | **MonthlyTrfCount** | Recurrence, monthly | `FleetData."No Of TRFs"` | Before go-live |
| 15 | **NightlyReportSnapshot** | Recurrence, daily 03:00 | `ReportSnapshot` | When reports get slow |

The requirements name five flows (Broadcast, Claim, Expire, Notify, Complete). That's the **engine**. Flows 4–9 are the HTTP endpoints the web parts need for the remaining actions, and 12–15 are the scheduled jobs the requirements describe in prose without naming. Nine is the minimum for a working system; the four scheduled ones can follow.

### Seven endpoint URLs go into the web part property pane

| Property pane field | Flow |
| --- | --- |
| Claim flow URL | 2 |
| Respond to offer flow URL | 4 |
| Update trip status flow URL | 5 |
| Manual assign flow URL | 6 |
| Cancel flow URL | 7 |
| Reject flow URL | 8 |
| Set ETA flow URL | 9 |

Paste the same seven into **every** web part that needs them.

---

## 2. Shared setup — do this once

### 2.1 Build inside a solution

**Power Apps → Solutions → New solution** (`Fleet Booking and Dispatch`). Create every flow *inside* it.

Flows edited directly in production are the most common cause of drift in Power Automate projects. A solution is what lets you export dev → UAT → production with environment variables swapped instead of hand-editing URLs in each environment.

### 2.2 Environment variables

**New → More → Environment variable**, type **Text**:

| Name | Dev value |
| --- | --- |
| `FleetSiteUrl` | `https://pomae.sharepoint.com/sites/travel-dev` |
| `FleetPortalPageUrl` | `https://pomae.sharepoint.com/sites/travel-dev/SitePages/Driver.aspx` |
| `CoordinatorEmail` | your address for testing |

Reference as `@{parameters('FleetSiteUrl (new_fleetsiteurl)')}`. Never type a site URL into an action.

### 2.3 Connections

| Connector | Used for | Notes |
| --- | --- | --- |
| **SharePoint** | Every list read and write | **The service identity.** See below |
| **Office 365 Outlook** | Every e-mail | Must be a mailbox that can send |

All standard — no premium licensing, which was one reason SPFx + Flow was chosen over Power Apps.

> **Create the SharePoint connection as a dedicated service account**, not as yourself. Drivers get *no* direct write permission on `RideOffers` or `TripAssignments` — the flow writes on their behalf, so the connection owner needs Contribute on all the dispatch lists. If you use your own account, every flow stops the day you change role or leave.

### 2.4 The HTTP trigger identity problem — read before building flow 2

The "When an HTTP request is received" trigger produces a URL with a SAS signature in the query string. **That URL is in the SPFx bundle, which every authenticated user in the tenant can read.** Anyone who extracts it can call the flow directly.

The requirements ask the flow to "verify that the calling user matches `DriversMaster.UserAccount`". With a plain HTTP trigger **you cannot do this reliably**, because any user id in the request body is supplied by the caller and can be lied about.

Three options, in order of preference:

**A. Restrict the trigger to tenant users (recommended).** Newer Power Automate exposes **Who can trigger the flow** on the Request trigger — set it to *Any user in my tenant*. The caller must then present an OAuth token, and the flow can read the verified caller from the trigger headers. This requires a matching change in the web part: `FlowService` currently uses `context.httpClient`, which sends no token, and would need `context.aadHttpClientFactory`.
> Verify the exact header name in your own tenant before relying on it — inspect a run's trigger outputs and look for `x-ms-client-principal-name` or similar. Tell me if you take this route and I'll make the `FlowService` change.

**B. Queue list pattern.** The driver's action writes to an `ActionQueue` list they have Add-only permission on; a flow triggered on item creation reads `Created By`, which SharePoint sets server-side and cannot be spoofed. Genuinely secure with no Azure — but item-created triggers fire in seconds to minutes, which blows the **3-second claim latency** requirement. Viable for the slower actions (cancel, reject, set ETA), not for Claim.

**C. Accept the risk for phase 1 (what this guide assumes).** Flows take no user id from the body. Each one derives the driver from the *record* — the offer's `Driver`, the trip's `Driver` — so a caller can only act on an offer or trip that already exists. The residual risk is that an authenticated user who extracts the URL could accept or progress a trip belonging to another driver. Compensating controls: `TripStatusHistory` and `NotificationLog` are append-only and record every change, and site audit logging is on.

**This is a decision for the business, not for the person building the flows.** Put it in writing before go-live.

### 2.5 Two conventions every flow follows

**Correlation ID.** Every request body carries `correlationId` (the web part generates it). First action in every HTTP flow:

```
Compose — name: CorrelationId
Input:  @{coalesce(triggerBody()?['correlationId'], guid())}
```

Write it into `NotificationLog` and into every `TripStatusHistory.Comment`. It's what lets one failed booking be traced end to end. **Never log personal data alongside it.**

**Response codes.** `FlowService` treats any non-2xx as a failure and shows the user an error banner. So:

- Success → **200** (or 204 with no body)
- A business outcome that isn't an error → **200 with a body saying so** (see Claim)
- A genuine failure → **4xx/5xx with a plain-text reason**, which gets logged to the browser console

---

## 3. Flow 1 — Broadcast

> **Trigger:** When an item is created · **List:** `TransportRequests`

Builds the eligible driver set, writes one offer row per driver, e-mails them all.

### Trigger

| Setting | Value |
| --- | --- |
| Connector | SharePoint |
| Trigger | **When an item is created** |
| Site Address | `@{parameters('FleetSiteUrl …')}` |
| List Name | `TransportRequests` |
| Settings → Concurrency Control | Off |

### Steps

**1. Guard — only act on new submissions**

```
Condition:  triggerOutputs()?['body/Status/Value']  is equal to  Submitted
```
If no → **Terminate** (status: Succeeded). The flow re-fires on items the web part creates in two steps (create, then write `RequestId` back), so this guard matters.

**2. Initialise variables**

| Name | Type | Value |
| --- | --- | --- |
| `RequestItemId` | Integer | `triggerOutputs()?['body/ID']` |
| `PropertyId` | Integer | `triggerOutputs()?['body/Property/Id']` |
| `VehicleTypeId` | Integer | `triggerOutputs()?['body/VehicleTypeRequested/Id']` |
| `WindowFrom` | String | `triggerOutputs()?['body/RequiredDateTime']` |
| `WindowTo` | String | `coalesce(triggerOutputs()?['body/ReturnDateTime'], addHours(variables('WindowFrom'), 2))` |
| `EligibleDriverIds` | Array | `[]` |

`WindowTo` defaults to +2 hours when no return time is given — the broadcast still has to reserve *something*.

**3. Get the allocation rule**

```
Get items — AllocationRules
Filter Query:  Property/Id eq @{variables('PropertyId')}
               and RequestType eq '@{triggerOutputs()?['body/RequestType/Value']}'
               and IsActive eq 1
Top Count: 1
```

Then set, with defaults for when no rule row exists:

| Variable | Expression |
| --- | --- |
| `OfferWindow` | `coalesce(first(body('Get_items_AllocationRules')?['value'])?['OfferWindowMinutes'], 5)` |
| `MaxTrips` | `coalesce(first(…)?['MaxConcurrentTripsPerDriver'], 1)` |
| `AdvanceHours` | `coalesce(first(…)?['AdvanceReleaseHours'], 4)` |
| `Mode` | `coalesce(first(…)?['Mode/Value'], 'Automatic')` |

**4. Advance booking? Hold it, don't broadcast**

```
Condition:  greater(
              sub(ticks(variables('WindowFrom')), ticks(utcNow())),
              mul(variables('AdvanceHours'), 36000000000)
            )  is equal to  true
```

*Yes:* Update item → `Status = Held`, `ReleaseAt = addHours(variables('WindowFrom'), mul(-1, variables('AdvanceHours')))` → **Terminate (Succeeded)**.

A booking made three weeks out would otherwise sit in drivers' dashboards for weeks. Flow 12 releases it.

**5. Manual mode? Straight to the Coordinator**

```
Condition:  variables('Mode')  is equal to  Manual
```
*Yes:* Update item → `Status = Awaiting Coordinator` → call **Notify** with `WindowExpired.Coordinator` → **Terminate (Succeeded)**.

**6. Send the acknowledgement e-mail immediately**

Call **Notify** (flow 10) with `TemplateKey = RequestSubmitted.Requester`.

> Do this **before** the eligibility work, not after. The client's requirement is that the requester is told *at once* that the request was received. Acknowledgement is not confirmation — the template wording already says so.

**7. Build the eligible driver set**

*7a — drivers allocated to this property*
```
Get items — DLClientMaster
Filter Query:  Property/Id eq @{variables('PropertyId')}
```
```
Compose — AllocatedDriverIds
  union(select(body('Get_items_DLClientMaster')?['value'], item()?['Driver']?['Id']), createArray())
```

*7b — blocks overlapping the window*
```
Get items — AvailabilityBlocks
Filter Query:  FromDateTime lt '@{variables('WindowTo')}'
               and ToDateTime gt '@{variables('WindowFrom')}'
```

*7c — open trips*
```
Get items — TripAssignments
Filter Query:  TripStatus eq 'Assigned' or TripStatus eq 'En Route' or TripStatus eq 'In Transit'
```

*7d — candidate drivers*
```
Get items — DriversMaster
Filter Query:  IsActive eq 1 and Status ne 'Inactive'
Top Count: 500
```

*7e — Apply to each* over `body('Get_items_DriversMaster')?['value']`, **Concurrency off** (it appends to a shared array). Inside, one Condition with all five checks ANDed:

| Check | Expression |
| --- | --- |
| Allocated to the property | `contains(outputs('AllocatedDriverIds'), items('Apply_to_each')?['ID'])` is `true` |
| Qualified for the vehicle type | `contains(select(items('Apply_to_each')?['QualifiedVehicleTypes'], item()?['Id']), variables('VehicleTypeId'))` is `true` |
| Not blocked in the window | `length(filter(body('Get_items_AvailabilityBlocks')?['value'], and(equals(item()?['Driver']?['Id'], items('Apply_to_each')?['ID']), not(equals(item()?['BlockType']?['Value'], 'Shift')))))` is equal to `0` |
| Licence valid to the end of the window | `or(empty(items('Apply_to_each')?['LicenceExpiry']), greaterOrEquals(items('Apply_to_each')?['LicenceExpiry'], variables('WindowTo')))` is `true` |
| Under the concurrent cap | `length(filter(body('Get_items_TripAssignments')?['value'], equals(item()?['Driver']?['Id'], items('Apply_to_each')?['ID'])))` is **less than** `coalesce(items('Apply_to_each')?['MaxConcurrentTrips'], variables('MaxTrips'))` |

If yes → **Append to array variable** `EligibleDriverIds` ← `items('Apply_to_each')?['ID']`.

> The licence must be valid through the **end** of the window, not just the start. A licence expiring mid-trip is the case that catches people out.

**8. Is at least one vehicle of the type free?**

```
Get items — VehicleMaster
Filter Query:  VehicleType/Id eq @{variables('VehicleTypeId')}
               and Property/Id eq @{variables('PropertyId')}
               and Status ne 'Retired' and Status ne 'Under Maintenance'
```
```
Compose — FreeVehicleCount
  length(filter(body('Get_items_VehicleMaster')?['value'],
    not(contains(select(body('Get_items_TripAssignments')?['value'], item()?['Vehicle']?['Id']), item()?['ID']))))
```

The broadcast confirms only that a suitable vehicle **exists**. The vehicle is bound at claim time — binding one to every offer would reserve the whole fleet against offers that will mostly be declined.

**9. Nobody eligible, or no vehicle? Escalate**

```
Condition:  or(equals(length(variables('EligibleDriverIds')), 0),
               equals(outputs('FreeVehicleCount'), 0))  is equal to  true
```
*Yes:* Update item → `Status = Awaiting Coordinator`, `OfferClosesAt = utcNow()` → Notify `WindowExpired.Coordinator` → **Terminate (Succeeded)**.

**10. Open the window**

```
Update item — TransportRequests  (ID = variables('RequestItemId'))
  Status         = Broadcasting
  OfferClosesAt  = addMinutes(utcNow(), variables('OfferWindow'))
```

Set the status **before** creating the offers. A driver who sees a card and taps Accept while the request still reads `Submitted` gets rejected by Claim's guard.

**11. One offer row per eligible driver, and the e-mail**

*Apply to each* over `variables('EligibleDriverIds')`:

```
Create item — RideOffers
  Title           = @{triggerOutputs()?['body/RequestId']} - @{items('Apply_to_each_2')}
  Request Id      = variables('RequestItemId')
  Driver Id       = items('Apply_to_each_2')
  OfferedOn       = utcNow()
  ResponseStatus  = Pending
```
Leave `Vehicle` empty. Then call **Notify** with `BroadcastOpened.Driver`.

> **This is the single highest-value message in the system.** Without Teams it is the only thing that pulls a driver to the dashboard inside the five-minute window. The subject must be readable from a lock-screen preview — pickup, destination, time — and the body must deep-link to `FleetPortalPageUrl`.

---

## 4. Flow 2 — Claim

> **Trigger:** When an HTTP request is received · **Called by:** Driver Dashboard, Accept

The race is the central engineering problem: two drivers can tap Accept in the same second, and SharePoint has no transactions.

### Trigger

Request body JSON schema:
```json
{
  "type": "object",
  "properties": {
    "offerItemId": { "type": "integer" },
    "correlationId": { "type": "string" }
  }
}
```

**Settings → Concurrency Control → ON → Degree of parallelism = 1.**

Claims are then processed strictly in order. At the expected volume this costs nothing and removes an entire class of defect. Do not skip it.

### The response contract — get this right or the UI misbehaves

| Outcome | HTTP | Body |
| --- | --- | --- |
| This driver won | **200** | `{"won": true, "tripItemId": 42, "correlationId": "..."}` |
| Another driver won | **200** | `{"won": false, "message": "Taken by another driver", "correlationId": "..."}` |
| Real failure | 500 | plain text reason |

**Losing is a 200, not a 409.** `FlowService` throws on any non-2xx, so a 409 would show the losing driver a red error banner instead of the clean "Taken by another driver" message, and the card would stay on screen. Losing is the *expected* outcome for every driver but one.

### Steps

**1.** `Compose — CorrelationId` → `@{coalesce(triggerBody()?['correlationId'], guid())}`

**2. Get the offer**
```
Get item — RideOffers,  Id = triggerBody()?['offerItemId']
```

**3. Offer still open?**
```
Condition:  outputs('Get_item_RideOffers')?['body/ResponseStatus/Value']  is equal to  Pending
```
*No* → **Response 200** `{"won": false, "message": "This offer is no longer open"}` → Terminate.

**4. Read the request *with its ETag***

SharePoint's "Get item" doesn't expose the ETag, so use the REST action:

```
Send an HTTP request to SharePoint
  Site Address: @{parameters('FleetSiteUrl …')}
  Method: GET
  Uri: _api/web/lists/getbytitle('TransportRequests')/items(@{outputs('Get_item_RideOffers')?['body/Request/Id']})
  Headers:  Accept: application/json;odata=minimalmetadata
```

```
Compose — ETag
  body('Send_an_HTTP_request_to_SharePoint')?['odata.etag']
```

**5. Still broadcasting?**
```
Condition:  body('Send_an_HTTP_request_to_SharePoint')?['Status']  is equal to  Broadcasting
```
*No* → **Response 200** `{"won": false, "message": "Taken by another driver"}` → Terminate.

This is the cheap check. The ETag below is the authoritative one.

**6. Bind a vehicle, re-verifying availability in the same read**

Get open trips and candidate vehicles as in Broadcast steps 7c/8, then:

```
Compose — FreeVehicles
  filter(body('Get_items_VehicleMaster')?['value'],
    not(contains(select(body('Get_items_TripAssignments')?['value'], item()?['Vehicle']?['Id']), item()?['ID'])))

Compose — PreferredVehicle
  first(filter(outputs('FreeVehicles'),
    equals(item()?['DefaultDriver']?['Id'], outputs('Get_item_RideOffers')?['body/Driver/Id'])))

Compose — BoundVehicleId
  coalesce(outputs('PreferredVehicle')?['ID'], first(outputs('FreeVehicles'))?['ID'])
```

Prefer this driver's default vehicle if it's free, otherwise the first free one of the right type. If `BoundVehicleId` is empty → **Response 200** `{"won": false, "message": "No vehicle is available for this trip"}`.

**7. The check-and-set — this is the whole flow**

Put this in a **Scope** named `Try claim`:

```
Send an HTTP request to SharePoint
  Method: POST
  Uri: _api/web/lists/getbytitle('TransportRequests')/items(@{outputs('Get_item_RideOffers')?['body/Request/Id']})
  Headers:
    Content-Type:  application/json;odata=verbose
    Accept:        application/json;odata=verbose
    IF-MATCH:      @{outputs('ETag')}
    X-HTTP-Method: MERGE
  Body:
    {
      "__metadata": { "type": "SP.Data.TransportRequestsListItem" },
      "Status": "Allocated"
    }
```

An ETag mismatch returns **412 Precondition Failed** and the action fails. That means another driver got there first.

Add a parallel branch after the scope:

- **Scope `Lost the race`** — *Configure run after* `Try claim` → **has failed** ✔ (uncheck Successful)
  → **Response 200** `{"won": false, "message": "Taken by another driver", "correlationId": "..."}` → Terminate (Succeeded)
- **Scope `Won the race`** — runs after `Try claim` is **Successful** → step 8

> Check the entity type name. For a list titled `TransportRequests` it's `SP.Data.TransportRequestsListItem`, but a list created with a different original name keeps that name. Confirm with:
> `_api/web/lists/getbytitle('TransportRequests')?$select=ListItemEntityTypeFullName`

**8. On success, in order**

```
a) Create item — TripAssignments
     Request Id  = offer's Request/Id
     Driver Id   = offer's Driver/Id
     Vehicle Id  = outputs('BoundVehicleId')
     AcceptedOn  = utcNow()
     TripStatus  = Assigned
     AllocationMode = Automatic
     IsOverride  = No

b) Update item — TripAssignments (the row just created)
     Title  = TR-@{formatNumber(outputs('Create_item_TripAssignments')?['body/ID'], '0000')}
     TripId = same

c) Update item — TransportRequests
     AcceptedTrip Id = the new trip ID

d) Update item — RideOffers (the accepting driver's row)
     ResponseStatus = Accepted
     RespondedOn    = utcNow()
     Vehicle Id     = outputs('BoundVehicleId')

e) Get items — RideOffers
     Filter: Request/Id eq <id> and ResponseStatus eq 'Pending'
   Apply to each → Update item → ResponseStatus = Closed

f) Update item — DriversMaster  → Status = On Trip
g) Update item — VehicleMaster  → Status = On Trip

h) Create item — TripStatusHistory
     Trip Id = new trip, FromStatus = (blank), ToStatus = Assigned
     ChangedBy = the driver's UserAccount, ChangedOn = utcNow()
     Comment = Accepted from broadcast. Ref @{outputs('CorrelationId')}

i) Call Notify ×2:  DriverAccepted.Driver  and  DriverAccepted.Requester

j) Response 200
     {"won": true, "tripItemId": <new id>, "correlationId": "@{outputs('CorrelationId')}"}
```

Losing offers become **Closed**, not Expired. Expired means the window ran out; Closed means someone else won. The reports distinguish them.

> Steps (f) and (g) write *denormalised* status. `DriversMaster.Status` and `VehicleMaster.Status` are maintained by flows rather than calculated at read time, so the availability panel and the two Available counters are a single list read. Flows are the only writer — flow 13 corrects drift from manual list edits.

---

## 5. Flow 3 — Expire

> **Trigger:** Recurrence, every **1 minute**

### Steps

```
1. Get items — TransportRequests
     Filter Query:  Status eq 'Broadcasting' and OfferClosesAt lt '@{utcNow()}'
     Top Count: 50

2. Apply to each:
   a) Update item → Status = Awaiting Coordinator
   b) Get items — RideOffers
        Filter: Request/Id eq <id> and ResponseStatus eq 'Pending'
      Apply to each → Update item → ResponseStatus = Expired
   c) Get items — RideOffers
        Filter: Request/Id eq <id> and ResponseStatus eq 'Rejected'
      Compose — RejectionSummary
        join(select(body('…')?['value'],
          concat(item()?['Driver']?['Value'], ': ', coalesce(item()?['RejectionReason']?['Value'], 'no reason'))),
          '; ')
   d) Call Notify → WindowExpired.Coordinator, passing OfferedCount and RejectionSummary
```

The Coordinator alert must carry **who was offered the trip and why they declined**, so they can see why nobody took it before deciding what to do. That list is the whole point of keeping offers in a separate list from trips.

> A one-minute recurrence is 1,440 runs a day against your flow quota. If that's tight, two minutes is a defensible compromise — the window is five minutes, so the worst case is a 7-minute wait before escalation. Don't go beyond that.

---

## 6. Flow 4 — RespondToOffer

> **Trigger:** HTTP · **Called by:** Driver Dashboard, Reject

Schema:
```json
{
  "type": "object",
  "properties": {
    "offerItemId": { "type": "integer" },
    "response": { "type": "string" },
    "reason": { "type": "string" },
    "notes": { "type": "string" },
    "correlationId": { "type": "string" }
  }
}
```

Steps:
```
1. Get item — RideOffers, Id = triggerBody()?['offerItemId']
2. Condition: ResponseStatus is equal to Pending
     No → Response 409 "This offer is no longer open." → Terminate
3. Update item — RideOffers
     ResponseStatus  = Rejected
     RejectionReason = triggerBody()?['reason']
     RejectionNotes  = triggerBody()?['notes']
     RespondedOn     = utcNow()
4. Response 200  {}
```

No e-mail. A rejection is not news to anybody in real time — it surfaces in the escalation alert and in the Driver response behaviour report. The reason is stored because a driver who rejects everything is a management issue the Coordinator should be able to see.

The request **does not** go to Awaiting Coordinator just because one driver declined. Other drivers may still accept; flow 3 escalates it if nobody does.

---

## 7. Flow 5 — UpdateTripStatus

> **Trigger:** HTTP · **Called by:** the four driver taps

Schema:
```json
{
  "type": "object",
  "properties": {
    "tripItemId": { "type": "integer" },
    "toStatus": { "type": "string" },
    "odometer": { "type": "integer" },
    "correlationId": { "type": "string" }
  }
}
```

### Steps

```
1. Get item — TripAssignments, Id = triggerBody()?['tripItemId']
2. Compose — FromStatus  = outputs('Get_item')?['body/TripStatus/Value']

3. Validate the transition SERVER-SIDE
   Compose — Allowed
     if(equals(outputs('FromStatus'),'Assigned'),   createArray('En Route','Cancelled'),
     if(equals(outputs('FromStatus'),'En Route'),   createArray('In Transit','Cancelled'),
     if(equals(outputs('FromStatus'),'In Transit'), createArray('Completed','Cancelled'),
        createArray())))

   Condition: contains(outputs('Allowed'), triggerBody()?['toStatus'])  is equal to  true
     No → Response 409
          "A trip that is @{outputs('FromStatus')} cannot be set to @{triggerBody()?['toStatus']}."
          → Terminate
```

> **The dashboard hiding a button is not the control; this flow is.** The web part offers only legal transitions as a usability measure, but the rule lives here.

```
4. Require the odometer reading where the tap demands one
   Condition:  or(and(equals(toStatus,'En Route'),  empty(triggerBody()?['odometer'])),
                  and(equals(toStatus,'Completed'), empty(triggerBody()?['odometer'])))
     Yes → Response 400 "The odometer reading is required for this step." → Terminate

5. Switch on triggerBody()?['toStatus']:

   En Route    → Update item: TripStatus = En Route,   ActualStartTime  = utcNow(),
                              OpeningReading = triggerBody()?['odometer']
                 Notify StatusEnRoute.Requester   (only if the rule's NotifyEnRoute is Yes)

   In Transit  → Update item: TripStatus = In Transit, ActualPickupTime = utcNow()
                 Notify StatusInTransit.Requester (only if NotifyInTransit is Yes)

   Completed   → Update item: TripStatus = Completed,  ActualDropTime   = utcNow(),
                              ClosingReading = triggerBody()?['odometer']
                 Call the Complete child flow (flow 11)

6. Create item — TripStatusHistory
     Trip Id = tripItemId, FromStatus = outputs('FromStatus'), ToStatus = toStatus
     ChangedBy = the trip's Driver UserAccount, ChangedOn = utcNow()
     Comment = Ref @{outputs('CorrelationId')}

7. Response 200 {}
```

Step 5's notification conditions implement **volume discipline**: a requester who booked a two-minute internal transfer does not need three e-mails about it. Read `NotifyEnRoute` / `NotifyInTransit` from the `AllocationRules` row for that property and request type.

---

## 8. Flow 6 — ManualAssign

> **Trigger:** HTTP · **Called by:** Dispatch Dashboard, Assign and Reassign

**One flow, two shapes.** The web part sends either an assignment or a reassignment:

```json
// assign
{ "requestItemId": 7, "driverItemId": 2, "vehicleItemId": 3,
  "isOverride": true, "overrideReason": "...", "correlationId": "..." }

// reassign
{ "tripItemId": 12, "driverItemId": 4, "vehicleItemId": 5,
  "reason": "...", "reassign": true, "correlationId": "..." }
```

Schema — make every property optional (don't mark any as required):
```json
{
  "type": "object",
  "properties": {
    "requestItemId":  { "type": "integer" },
    "tripItemId":     { "type": "integer" },
    "driverItemId":   { "type": "integer" },
    "vehicleItemId":  { "type": "integer" },
    "isOverride":     { "type": "boolean" },
    "overrideReason": { "type": "string" },
    "reason":         { "type": "string" },
    "reassign":       { "type": "boolean" },
    "correlationId":  { "type": "string" }
  }
}
```

First action: `Condition — equals(triggerBody()?['reassign'], true)`.

### Branch A — fresh assignment

```
1. Get item — TransportRequests, Id = requestItemId
2. Condition: Status is one of Submitted / Held / Broadcasting / Awaiting Coordinator
     No → Response 409 "This request has already been allocated." → Terminate

3. Withdraw the broadcast
   Get items — RideOffers: Request/Id eq <id> and ResponseStatus eq 'Pending'
   Apply to each → Update item → ResponseStatus = Withdrawn

4. Create item — TripAssignments
     Request/Driver/Vehicle Id, AcceptedOn = utcNow(), TripStatus = Assigned
     AllocationMode = Manual
     AllocatedBy    = the Coordinator
     IsOverride     = triggerBody()?['isOverride']
     OverrideReason = triggerBody()?['overrideReason']
5. Update Title and TripId = TR-nnnn
6. Update item — TransportRequests → Status = Allocated, AcceptedTrip Id = new trip
7. Update DriversMaster + VehicleMaster → Status = On Trip
8. Create item — TripStatusHistory → ToStatus = Assigned, Comment = Assigned manually. Ref …
9. Notify ManualAssign.Driver  and  ManualAssign.Requester
10. Response 200 {}
```

**Manual assignment skips the broadcast entirely.** Open offers become **Withdrawn** — a fourth distinct outcome, not Closed or Expired — and the chosen driver is *notified*, not offered a choice. The Driver Dashboard shows them an assignment, with no Accept button.

### Branch B — reassignment

```
1. Get item — TripAssignments, Id = tripItemId
2. Release the outgoing driver and vehicle → Status = Available
3. Update item — TripAssignments → new Driver Id, new Vehicle Id,
     IsOverride = Yes, OverrideReason = triggerBody()?['reason']
4. Set the new driver and vehicle → Status = On Trip
5. Create item — TripStatusHistory
     FromStatus = current, ToStatus = current (status unchanged — the reassignment is the event)
     Comment = Reassigned from <old driver> to <new driver>: <reason>. Ref …
6. Notify: Cancelled.Driver to the OUTGOING driver, ManualAssign.Driver to the incoming one,
   ManualAssign.Requester to the requester (their driver and car have changed)
7. Response 200 {}
```

Don't skip the requester e-mail in step 6. They are standing at a kerbside looking for a specific car.

---

## 9. Flow 7 — Cancel

> **Trigger:** HTTP · **Called by:** requester, driver and Coordinator

**One flow, four shapes:**

| Caller | Body |
| --- | --- |
| Requester, before allocation | `{requestItemId, reason, cancelledBy:"requester"}` |
| Requester, after allocation | `{requestItemId, reason, cancelledBy:"requester", routeToCoordinator:true}` |
| Driver, after accepting | `{tripItemId, reason, cancelledBy:"driver"}` |
| Coordinator | `{requestItemId, reason, cancelledBy:"coordinator"}` |

Schema: `requestItemId`, `tripItemId` (integer), `reason`, `cancelledBy`, `correlationId` (string), `routeToCoordinator` (boolean) — all optional.

### Steps

```
1. Condition: equals(triggerBody()?['routeToCoordinator'], true)
   Yes → This is a REQUEST to cancel, not a cancellation.
         Do NOT change any status.
         Notify Cancelled.Coordinator — "<requester> has asked to cancel <RequestId>: <reason>"
         Notify the assigned driver that cancellation has been requested
         Response 200 {} → Terminate

2. Condition: empty(triggerBody()?['tripItemId'])
```

**2a — request-level cancellation** (requester before allocation, or Coordinator)
```
Get item — TransportRequests
Condition: cancelledBy eq 'requester'
           and Status not in (Submitted, Held, Awaiting Coordinator)
  → Response 409 "This booking can no longer be cancelled directly." → Terminate

Update item → Status = Cancelled, CancellationReason = reason
Withdraw any Pending offers → ResponseStatus = Withdrawn
If a trip exists → TripStatus = Cancelled, release driver + vehicle
Notify Cancelled.Requester (+ Cancelled.Driver if one was assigned)
```

**2b — trip-level cancellation** (driver)
```
Get item — TripAssignments
Update item → TripStatus = Cancelled, CancellationReason = reason
Update TransportRequests → Status = Awaiting Coordinator     ← NOT Cancelled
Release driver + vehicle → Status = Available
Create TripStatusHistory → ToStatus = Cancelled, Comment = reason. Ref …
Notify Cancelled.Coordinator, Cancelled.Requester
```

> **A driver cancelling does not cancel the booking.** The trip returns to the Coordinator for reallocation, and it is **not re-broadcast automatically** — a cancellation close to the required time needs human judgement. Setting the request to `Cancelled` here would silently strand a requester who still needs transport.

```
3. Response 200 {}
```

---

## 10. Flow 8 — Reject

> **Trigger:** HTTP · **Called by:** Dispatch Dashboard

Schema: `requestItemId` (integer), `reason`, `correlationId` (string).

```
1. Get item — TransportRequests
2. Condition: Status is one of Submitted / Held / Broadcasting / Awaiting Coordinator
     No → Response 409 "An allocated request cannot be rejected — cancel it instead." → Terminate
3. Update item → Status = Rejected, RejectionReason = triggerBody()?['reason']
4. Withdraw any Pending offers → ResponseStatus = Withdrawn
5. Notify RequestRejected.Requester
6. Response 200 {}
```

The reason is mandatory in the web part and goes straight into the requester's e-mail. Write it as something they can act on.

---

## 11. Flow 9 — SetEta

> **Trigger:** HTTP · **Called by:** Coordinator or driver

Schema: `tripItemId`, `etaMinutes` (integer), `correlationId` (string).

```
1. Update item — TripAssignments
     EtaMinutes   = triggerBody()?['etaMinutes']
     EtaUpdatedOn = utcNow()
2. Response 200 {}
```

No e-mail. The requester sees it next time Track my ride polls.

**Always write `EtaUpdatedOn`.** The interface shows "approx. 15 min · estimate last updated 09:14" so nobody mistakes a typed-in figure for a live calculation. Phase 2 populates the same two fields from a routing service with no interface change.

---

## 12. Flow 10 — Notify (child flow)

> **Trigger:** When a flow invokes this (child flow) · **Called by:** flows 1–9

Every message is sent by a flow, never from the browser, so a closed tab cannot lose one.

### Input parameters

| Name | Type |
| --- | --- |
| `TemplateKey` | Text |
| `RecipientEmail` | Text |
| `RequestItemId` | Number |
| `TripItemId` | Number |
| `TokensJson` | Text (JSON of placeholder → value) |
| `CorrelationId` | Text |

### Steps

```
1. Get items — Templates
     Filter: Title eq '@{TemplateKey}' and Language eq 'en' and IsActive eq 1
     Top Count: 1
   If none → Create NotificationLog (Outcome = Failed,
             ErrorMessage = "No template for <key>") → Terminate (Failed)

2. Compose — Tokens = json(TokensJson)

3. Compose — Subject / Body
   Replace each {Placeholder}. Nested replace() is unreadable past three tokens;
   prefer an Apply to each over the token object:
     Set variable Subject = replace(variables('Subject'),
       concat('{', items('Apply_to_each')?['key'], '}'),
       string(items('Apply_to_each')?['value']))

4. Scope "Send"
   Send an email (V2) — To: RecipientEmail, Subject, Body (IsHtml = Yes)
   Settings → Retry Policy → Exponential, Count 3, Interval PT10S

5. Create item — NotificationLog
   On success:  Outcome = Sent,   RetryCount = 0
   On failure (run after: has failed):
                Outcome = Failed, ErrorMessage = the error, RetryCount = 3
                + Send an email to CoordinatorEmail using NotificationFailed.Coordinator
   Always:  Channel = Email, Recipient, TemplateKey, SentOn = utcNow(),
            RelatedRequest Id, RelatedTrip Id
```

Templates live in the `Templates` list, **not inside flow definitions**, so wording and language can change without a redeployment. The `Language` column is already there for Arabic.

> A silent send failure means a driver never learns they have a trip. That's why step 5 raises a Coordinator alert rather than just writing a log row. `NotificationLog` is append-only — no group gets delete, including Admins.

---

## 13. Flow 11 — Complete (child flow)

> **Trigger:** When a flow invokes this · **Called by:** flow 5 when a trip reaches Completed

```
1. Get item — TripAssignments → and its Request, Driver, Vehicle
2. Update item — TransportRequests → Status = Completed
3. Release: DriversMaster → Available,  VehicleMaster → Available
4. Create item — Daily Trip Sheet
     Map from the trip: date, driver, vehicle, pickup, destination,
     OpeningReading, ClosingReading, distance = Closing − Opening
5. Notify TripCompleted.Requester
```

> **Step 4 is the only place the new system writes into an existing list, apart from `FleetData."No Of TRFs"` in flow 14.** `Daily Trip Sheet` keeps its schema exactly — it simply stops being filled in by hand. Map to the existing columns; don't add any. Check the real column names on the live list before building this step; they weren't in the requirements.
>
> The odometer readings carried here are what remove a re-keying step the drivers do today.

---

## 14. Flows 12–15 — the scheduled jobs

### 12. ReleaseAdvanceBookings — Recurrence, every 15 minutes

```
Get items — TransportRequests
  Filter: Status eq 'Held' and ReleaseAt le '@{utcNow()}'
Apply to each → Update item → Status = Submitted
```

Setting it back to `Submitted` makes flow 1 fire again — **only if flow 1's trigger is "when an item is created or modified"**. With the created-only trigger in section 3, call the Broadcast logic directly instead: extract steps 7–11 into their own child flow and call it from both places. Cleaner than loosening the trigger, which would re-fire Broadcast on every edit.

### 13. NightlyReconciliation — Recurrence, daily 02:00

```
Get open trips (Assigned / En Route / In Transit)
For every driver:  expected = On Trip if they hold an open trip, else Available
                   (never overwrite On Leave or Inactive)
For every vehicle: expected = On Trip if on an open trip, else Available
                   (never overwrite Under Maintenance or Retired)
Where actual ≠ expected → update, and e-mail a summary to CoordinatorEmail
```

Denormalised status drifts when somebody edits a list by hand. This corrects it. The two exclusions matter — a nightly job that flips a driver back from On Leave to Available would be worse than the drift.

### 14. MonthlyTrfCount — Recurrence, monthly on the 1st

```
Get items — TransportRequests for last month (Created ge … and Created lt …)
Group by employee + vehicle
Update the matching FleetData row → "No Of TRFs" = count
```

Deliberately additive: **no other `FleetData` column is read or written**, and the monthly expense and approval pipeline continues to run exactly as it does today.

### 15. NightlyReportSnapshot — Recurrence, daily 03:00

Pre-aggregate yesterday into `ReportSnapshot`, one row per report key per dimension, measures as JSON in `MetricsJson`.

Build this when reports get slow, not before. The Reports web part currently reads live items; the aggregation guard exists because you must not aggregate thousands of list items in the browser, and `TransportRequests`, `RideOffers`, `TripAssignments` and `TripStatusHistory` will all pass 5,000 items.

---

## 15. Order of build, and how to test each one

| Step | Build | Test |
| --- | --- | --- |
| 1 | **Notify** with one template | Call it manually; check `NotificationLog` gets a Sent row |
| 2 | **Broadcast** | Book a ride in the web part → offers appear on the Driver Dashboard within 20 s |
| 3 | **Claim** | Tap Accept → booking confirmed, driver and vehicle flip to On Trip |
| 4 | **Expire** | Book, don't accept, wait 5 min → request pins to the top of Pending in amber |
| 5 | **UpdateTripStatus** | All four taps; check the timeline and both odometer prompts |
| 6 | **RespondToOffer**, **Cancel**, **Reject**, **SetEta** | Each button in the UI |
| 7 | **ManualAssign** | Assign from the escalation queue; try an out-of-policy driver |
| 8 | Scheduled jobs | |

Paste each URL into the web part property pane as you go. Until a URL is set, that action fails with a message naming exactly which one is missing — so you can ship flows incrementally.

### The test that matters most

Two people, two devices, same open offer, both tap **Accept** on a countdown.

One gets the trip. The other must see **"Taken by another driver"** *immediately*, and the card must disappear. If the loser sees a red error banner, your Claim flow is returning a non-2xx on the lost race — go back to section 4.

Then the load test the requirements ask for: **10 simultaneous claims on one request, zero double allocations.** Trigger the flow 10 times in parallel with the same `offerItemId` set and confirm exactly one `TripAssignments` row exists afterwards. If you see two, concurrency control is off or the `IF-MATCH` header isn't reaching SharePoint.

---

## 16. Things that will bite

- **`IF-MATCH` must be on the MERGE, not the GET.** It's easy to read the ETag and then forget to send it; the flow works perfectly in single-user testing and double-allocates the moment two drivers race.
- **Concurrency control = 1 on Claim only.** Put it on Broadcast or Notify and throughput collapses.
- **Lookup fields in SharePoint actions are `<Name> Id`**, e.g. `Driver Id`, and take an integer. Passing the display name silently writes nothing.
- **Date comparisons need quotes** in OData: `OfferClosesAt lt '@{utcNow()}'`.
- **`Status ne 'Retired'` excludes blanks too** in some SharePoint OData implementations. If vehicles go missing from availability, check for empty `Status` values rather than debugging the flow.
- **Both `RequestId` and `Title` carry the reference.** The web part reads `RequestId` and falls back to `Title`; lookup columns display `Title`. Write both, always.
- **Four distinct offer outcomes, and they are not interchangeable:** `Accepted` (won), `Rejected` (declined with a reason), `Expired` (window ran out), `Closed` (someone else won), `Withdrawn` (Coordinator assigned manually). The Driver response behaviour and Broadcast effectiveness reports count them separately.
- **The `RequestId` format is still an open question** — `FB-0001` or `FB-20260925-0001`. The web part writes it on create, so flows only ever read it. But settle it before the first production request exists.
