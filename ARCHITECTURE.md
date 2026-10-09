# Fleet Booking and Dispatch Portal — how it all works

Written for: you, understanding the system end to end.

This describes the solution **as it stands now**, after dispatch logic moved out of Power Automate and into the web part.

---

## 1. The one sentence version

A requester books a ride in four taps. The browser works out which drivers are eligible and writes an offer to each. The first driver to tap Accept wins, decided by an ETag lock on the request item. Flows run behind it all, sending email and doing the scheduled tidying.

---

## 2. Why the architecture looks like this

The original design put every write in Power Automate. That had to change: **the `When an HTTP request is received` trigger is a premium connector and this tenant has standard only.** There is no cheaper equivalent — polling triggers on this licence wake every five minutes, which is useless for a driver tapping Accept and expecting a response.

So the split is now:

| | Runs where | Latency |
| --- | --- | --- |
| Anything a human waits for | **The browser**, as the signed-in user | Immediate |
| Anything nobody watches happen | **Power Automate**, on polling triggers | Up to ~5 minutes |

That single constraint explains most of what follows — including why the offer window is 15 minutes rather than 5.

---

## 3. The layers

```
┌───────────────────────────────────────────────────────────────┐
│  FIVE WEB PARTS                                               │
│  Book a Ride · My Bookings · Driver Dashboard ·               │
│  Dispatch Dashboard · Reports                                 │
│  React 17 + Fluent 8. No component touches PnPjs directly.    │
└───────────────────────────┬───────────────────────────────────┘
                            │
┌───────────────────────────▼───────────────────────────────────┐
│  SERVICE LAYER            src/common/services/                │
│                                                               │
│   DispatchService   ← EVERY WRITE. The dispatch engine.       │
│   RequestService    ← reads: my bookings, timeline            │
│   OfferService      ← reads: my open offers                   │
│   TripService       ← reads: my trips today                   │
│   MasterDataService ← reads + caches reference data           │
│   ReportService     ← reads: the eight reports                │
│   UserService       ← who am I, what may I see                │
└───────────────────────────┬───────────────────────────────────┘
                            │
┌───────────────────────────▼───────────────────────────────────┐
│  PURE RULES               src/common/rules/                   │
│  dispatchRules · statusTransitions · identifiers · eligibility│
│  No SharePoint, no SPFx imports. This is what the unit tests  │
│  run against.                                                 │
└───────────────────────────┬───────────────────────────────────┘
                            │ PnPjs 4 / SPHttpClient
┌───────────────────────────▼───────────────────────────────────┐
│  SHAREPOINT LISTS on travel-dev                               │
│                                                               │
│  Master:    PropertyMaster · DLClientMaster · DriversMaster · │
│             VehicleMaster · VehicleTypes · Locations ·        │
│             AllocationRules · Templates                       │
│  Dispatch:  TransportRequests · RideOffers · TripAssignments ·│
│             TripStatusHistory · AvailabilityBlocks            │
│  Audit:     NotificationLog · ReportSnapshot                  │
│  Untouched: DriverLogs · LogEntries · DrvierEntries ·         │
│             FleetData · Daily Trip Sheet                      │
└───────────────────────────┬───────────────────────────────────┘
                            │ polling triggers
┌───────────────────────────▼───────────────────────────────────┐
│  POWER AUTOMATE — email and scheduled jobs only                │
└───────────────────────────────────────────────────────────────┘
```

**The rule that keeps this maintainable:** eligibility is implemented once, in `dispatchRules.ts`. No flow re-implements it. If a flow also decided who was eligible, the two would drift and nobody would know which was right.

---

## 4. The five roles

Role comes from SharePoint group membership, read once per page load by `UserService`.

| Group | Role | Sees | Can do |
| --- | --- | --- | --- |
| FM Requesters | `requester` | Book a Ride, My Bookings | Book, track, cancel own (before allocation) |
| FM Drivers | `driver` | Driver Dashboard | Accept/reject offers, four status taps, cancel own trip |
| FM Coordinators | `coordinator` | Dispatch Dashboard, Reports | Everything operational: assign, reassign, reject, cancel, set ETA |
| FM Admins | `admin` | Master data lists, Reports | Own drivers, vehicles, types, locations, rules |
| FM Viewers | `viewer` | Reports | Read only |

Someone in no group still gets `requester` — anyone who can open the portal can at least ask for a ride.

> **Role is a usability measure, not the security boundary.** It decides which buttons render. The real boundary is SharePoint list permissions, because anything the browser can read is reachable through the REST API regardless of what the interface shows.

**Property scoping** comes from `DLClientMaster`, not from groups, because a driver can be allocated to more than one property and a SharePoint group cannot express that. A Management Viewer has an empty property list, which means *all*.

---

## 5. What a REQUESTER does

### Booking — the seven screens

| Screen | They choose | Rule |
| --- | --- | --- |
| 1 | Pickup and destination | Pickup defaults to their property. Destination is a **searchable picker over `Locations`** — typing filters the list, it never becomes the answer |
| 2 | Today or another day, then a time slot | Slots, not a free-text time. Past slots disabled |
| 3 | Client or Client Guest · Staff · Items or Documents | Decides which vehicle tiles appear next |
| 4 | Vehicle **type**, passengers or item description | Filtered by `VehicleTypes.SuitableFor`. **"Others" reveals the only free-text field in the whole form** |
| 5 | — (processing) | Says *received*, not *confirmed* |
| 6 | — (confirmed) | Appears only once a driver accepts |
| 7 | — (track) | Status timeline, driver and vehicle details throughout |

**What the requester never sees:** vehicle availability, driver availability, the vehicle list, who declined, or any way to pick a specific driver or car. The booking web part does not query availability *at all* — not hidden, genuinely not read.

### What happens on Submit

```
createRequest()
  1. Insert TransportRequests   Title="NEW", ReqStatus="Submitted"
  2. Read back the item ID → Title = "FB-0007"
       (derived from the item ID; a counter item would race)
  3. Read AllocationRules for this property + request type
       none found → window 15, cap 1, release 4 h, Automatic

  4. Required time more than AdvanceReleaseHours away?
       → ReqStatus="Held", ReleaseAt set.  STOP. A flow releases it later.

  5. Rule mode = Manual?
       → ReqStatus="Awaiting Coordinator".  STOP.

  6. Otherwise → broadcast()  (section 7)
```

Screen 5 appears the moment step 2 finishes. The **acknowledgement email** comes from a flow and may lag a few minutes — that is why the screen says it in words rather than relying on the mail.

### Cancelling

One button, two very different behaviours, decided by the status the **service** reads — not by what the screen last saw.

| Status | What happens |
| --- | --- |
| Submitted / Held / Awaiting Coordinator | Cancelled outright. Pending offers → `Withdrawn` |
| Allocated (a driver has it) | **Nothing is changed.** A note goes on the trip history, and the requester is told the Coordinator has been asked |

A requester cannot pull a trip out from under a driver who is already on the way.

---

## 6. What a DRIVER does

Drivers have Microsoft 365 accounts but **no Teams**, and there is no push notification. This is the weakest link in the design, and three mechanisms cover it:

1. **The dashboard polls every 20 seconds** and shows an unmissable banner with the count of open offers
2. **An audible alert** the driver can switch on — fires only when the count *rises*
3. **The browser tab title** becomes `(1) New ride offer`, so a backgrounded tab still shows it

Plus the broadcast email, which arrives separately and may be slower (section 9).

A driver who is mid-trip or away from their phone will still miss a window. **That is acceptable by design** — the request escalates to the Coordinator rather than being lost.

### The offer card

Shows pickup, destination, time, request type, passenger count or item description, requester name, and a **live countdown**. It deliberately omits the requester's phone number until the driver accepts.

### Accept — the race

Two drivers can tap Accept in the same second. SharePoint has no transactions, so the request item's **ETag is the lock**.

```
acceptOffer(offerId)
  1. Offer still Pending?                    no → NOT_OPEN
  2. GET the request via SPHttpClient asking for minimalmetadata
     → capture odata.etag          (PnPjs strips it under nometadata)
  3. Status still "Broadcasting"?         no → TAKEN     ← cheap check
  4. Bind a vehicle — prefer this driver's DefaultDriver vehicle if free,
     else the first free one of the right type
     none → NO_VEHICLE
  5. POST MERGE with IF-MATCH: <etag>, body { Status: "Allocated" }
         412 Precondition Failed → TAKEN     ← the authoritative check
  6. WON. Then, in order:
       create TripAssignments (Assigned, Automatic)
       Title = "TR-0012"
       TransportRequests.AcceptedTripId
       winning offer  → Accepted
       other offers   → Closed        (NOT Expired — the window didn't run out)
       driver.Status  → On Trip
       vehicle.VehStatus → On Trip
       TripStatusHistory row
```

**Losing is the normal case.** Every driver but one loses, every time. So `TAKEN` comes back as `{ ok: false }`, not an exception, and the UI shows a plain notice and removes the card — never a red error banner. A driver who taps Accept and sees a failure stops trusting the system within a week.

### Reject

A reason from a fixed dropdown — *Already on a trip · Vehicle unavailable · Outside my shift · Personal reason · Other* — with a notes box only on "Other". Stored on the offer and reported on, because a driver who rejects everything is a management issue the Coordinator should be able to see.

**Rejecting does not change the request status.** Other drivers may still accept.

### The four taps

| Tap | Trip status | Also writes | Requester sees |
| --- | --- | --- | --- |
| Accept | `Assigned` | — | Booking confirmed, driver + car details |
| On my way | `En Route` | `ActualStartTime`, **`OpeningReading`** | "Your driver is on the way" |
| Picked up | `In Transit` | `ActualPickupTime` | "On the way to your destination" |
| Complete | `Completed` | `ActualDropTime`, **`ClosingReading`** | Completion time |

The transition table is enforced in the service, not just by hiding buttons. `Assigned → Completed` is refused. Cancellation is **not** in the table — it goes through `cancel()`, which has different side effects per caller.

The two odometer readings flow into `Daily Trip Sheet`, removing a re-keying step the drivers do today.

> **What tapping Complete does NOT do:** it does not set the request to `Completed`, does not free the driver or vehicle, and does not write `Daily Trip Sheet`. A flow does all three. See section 10 — this has a visible consequence.

### Driver cancels after accepting

Trip → `Cancelled`, driver and vehicle freed, and the **request goes back to `Awaiting Coordinator` — not `Cancelled`**. It is deliberately not re-broadcast automatically: a cancellation close to the required time needs human judgement. Marking the request cancelled here would silently strand a requester who still needs transport.

---

## 7. The allocation engine

Runs in the browser, in `broadcast()`. **Idempotent** — running it twice creates no duplicate offers, which is what makes the sweep in section 8 safe.

### Step 1 — the eligible driver set

Four reads, never a loop of single-item fetches:

| | List | Filter |
| --- | --- | --- |
| A | `DLClientMaster` | `PropertyRef/Id eq {property}` |
| B | `AvailabilityBlocks` | overlapping the trip window |
| C | `TripAssignments` | `Assigned` / `En Route` / `In Transit` |
| D | `DriversMaster` | `IsActive eq 1 and Status ne 'Inactive'` |

The trip window is `RequiredDateTime` → `ReturnDateTime`, or **+2 hours** when no return time is given — the broadcast has to reserve *something*.

A driver is eligible only if **all five** hold:

1. Allocated to this property (in A)
2. `QualifiedVehicleTypes` contains the requested type
3. No non-`Shift` block overlapping the window
4. `LicenceExpiry` empty, **or at or after the END of the window** — a licence expiring mid-trip excludes
5. Open trips **below** their cap (their own `MaxConcurrentTrips`, else the rule's)

### Step 2 — is any vehicle free?

Right type, right property, `VehStatus` not Retired/Under Maintenance, **and** not on an open trip **and** no non-`Shift` block.

> The block check is not optional. A vehicle booked into the garage through a block, with `VehStatus` still `Available`, would otherwise be offered and bound.

The broadcast only confirms a suitable vehicle **exists**. The specific vehicle is bound at claim time — binding one per offer would reserve the whole fleet against offers that mostly get declined.

### Step 3 — outcome

| Result | Action |
| --- | --- |
| No eligible driver, or no free vehicle | `Status = "Awaiting Coordinator"`, `OfferClosesAt = now` |
| Otherwise | `Status = "Broadcasting"`, `OfferClosesAt = now + window`, **then** one `RideOffers` row per driver, batched |

Status is set **before** the offers are created. A driver who sees a card while the request still reads `Submitted` gets rejected by the guard in `acceptOffer` and has no idea why.

---

## 8. What a COORDINATOR does

The Dispatch Dashboard is one batched read behind six counters, two grids and an availability panel. It polls every 60 seconds.

**Counters are filter shortcuts, not static numbers** — clicking one filters the grids below.

| Counter | Counts |
| --- | --- |
| Ongoing trips | `En Route` or `In Transit` |
| Pending requests | `Submitted`, `Broadcasting`, `Awaiting Coordinator` |
| Completed trips | completed in the selected period |
| Available drivers | active, not on leave, under cap |
| Available vehicles | active, no open trip |
| Allocated pairs | distinct driver+vehicle on an open trip |

### The sweep

Every poll also calls `sweepPendingBroadcasts()`, which finds `Submitted` requests and broadcasts them. It covers an advance booking a flow has just released, and a request whose creator closed their browser mid-submit.

> **Known limitation, by design:** a released advance booking is not broadcast until a Coordinator or driver has a dashboard open. Advance bookings release four hours before the trip, inside working hours, so this is acceptable for phase 1. It is deliberately not solved with a flow, because the eligibility rules must live in one place only.

### Manual assignment

Available at any time, whatever the mode. The panel shows every driver with a **verdict** — "not qualified for this vehicle type", "on leave during the trip window", "at the concurrent cap" — so the Coordinator knows they are overriding rather than discovering it later.

An out-of-policy choice is permitted with a **mandatory override reason**, enforced in the service, not just the form.

Manual assignment **skips the broadcast entirely**: open offers become `Withdrawn` — a distinct outcome from `Closed` and `Expired`, counted separately in reports — and the chosen driver is notified rather than offered a choice. They see an assignment with **no Accept button**.

### Escalation queue

Anything `Awaiting Coordinator` is pinned to the top of Pending Requests with an amber edge *and the words*, and the assignment panel shows **who declined and why**.

### Also

- **Reassign** — frees the old driver and vehicle, binds the new, leaves `TripStatus` untouched (the reassignment is the event, not a status change)
- **Reject** — only before allocation; an allocated request must be cancelled
- **Cancel** — cancels the request and any trip, frees everyone, withdraws offers
- **Set ETA** — writes `EtaMinutes` **and** `EtaUpdatedOn`, always both, so the UI can render *"approx. 15 min · updated 09:14"* and nobody mistakes a typed-in figure for a live calculation

---

## 9. Emails — what, when, and why they lag

**The web part never sends mail.** Every message comes from a flow watching a list.

Because HTTP triggers are premium, those flows use **polling triggers** — "when an item is created or modified" — which wake roughly every 5 minutes on this licence.

| Event | To | Why it matters |
| --- | --- | --- |
| Request created | Requester | Acknowledgement. States that confirmation follows — **acknowledgement is not confirmation** |
| `Status → Broadcasting` | Every offered driver | **The highest-value message in the system.** Subject readable from a lock screen: pickup, destination, time. Deep-links into the dashboard |
| `Status → Allocated` | Requester + winning driver | Full confirmation: driver name, phone, make, model, colour, plate |
| `Status → Awaiting Coordinator` | Coordinator | Escalation, with who declined and why |
| `TripStatus → En Route` | Requester | Switchable per property — a two-minute internal transfer doesn't need three emails |
| `TripStatus → In Transit` | Requester | Same |
| `TripStatus → Completed` | Requester | Completion time and feedback link |
| Cancelled (any) | All affected | With the reason |
| `Status → Rejected` | Requester | The reason, written so they can act on it |

Every send is logged to `NotificationLog` with its outcome. A failure raises a Coordinator alert, because a silent failure means a driver never learns they have a trip.

### The consequence you must design around

**The broadcast email can arrive up to five minutes after the offer opens.** That is exactly why `OfferWindowMinutes` defaults to **15, not 5** — the window has to outlast the mail delay or drivers get an offer that has already expired.

A driver with the dashboard **open** sees it in ≤20 seconds regardless. The email is for drivers who don't.

---

## 10. Who writes what — the complete split

### The browser writes (immediate)

| List | When |
| --- | --- |
| `TransportRequests` | create, status through the whole lifecycle except → Completed |
| `RideOffers` | create on broadcast; Accepted / Rejected / Closed / Withdrawn |
| `TripAssignments` | create on claim or manual assign; the four taps; ETA; reassign; cancel |
| `TripStatusHistory` | every transition |
| `DriversMaster.Status` | On Trip / Available |
| `VehicleMaster.VehStatus` | On Trip / Available |

### Flows write (within ~5 minutes)

| What | Why it's a flow |
| --- | --- |
| All email + `NotificationLog` | The browser must not send mail — a closed tab would lose it |
| `Status → Completed`, free driver + vehicle | Follows the `Daily Trip Sheet` write, kept together |
| `Daily Trip Sheet` | Its date columns are **text** in a format the web part must not guess |
| `RideOffers → Expired` | Needs to happen whether or not anyone has a browser open |
| `Held → Submitted` (advance release) | Scheduled; the sweep picks it up |
| Nightly status reconciliation | Corrects drift from hand edits |
| Monthly `FleetData."No Of TRFs"` | The only write into an existing list besides the trip sheet |

### The visible consequence

When a driver taps **Complete**, the trip is immediately `Completed` — but for up to five minutes:

- the **request** still reads `Allocated`
- the **driver and vehicle** still read `On Trip`
- so **Available drivers** and **Available vehicles** are each one low

This is expected, not a bug. If a Coordinator needs that driver immediately they can reassign; the nightly reconciliation corrects anything that drifts permanently.

---

## 11. Status lifecycles

### Request — `TransportRequests.Status`

```
                     ┌──────────────┐
         advance ───►│     Held     │──(flow releases)──┐
                     └──────────────┘                   │
                                                        ▼
   create ──────────────────────────────────────►  Submitted
                                                        │
                        ┌───────────────────────────────┤
                        │ no eligible driver /          │ broadcast()
                        │ no free vehicle / Manual mode │
                        ▼                               ▼
              Awaiting Coordinator ◄──(flow: window Broadcasting
                        │               expired)         │
                        │ manual assign                  │ a driver accepts
                        └──────────────┬─────────────────┘
                                       ▼
                                   Allocated
                                       │ (flow, after the Complete tap)
                                       ▼
                                   Completed

   Rejected / Cancelled — terminal, from any pre-allocation state
```

### Trip — `TripAssignments.TripStatus`

```
Assigned ──► En Route ──► In Transit ──► Completed
   │            │              │
   └────────────┴──────────────┴──► Cancelled   (via cancel(), not a tap)
```

### Offer — `RideOffers.ResponseStatus`

| Value | Set by | Means |
| --- | --- | --- |
| `Pending` | broadcast | Open, window running |
| `Accepted` | claim | This driver won |
| `Rejected` | driver | Declined, with a reason |
| `Closed` | claim | **Another driver won** |
| `Expired` | **flow only** | The window ran out |
| `Withdrawn` | manual assign | Coordinator bypassed the broadcast |

**These five are not interchangeable.** The Broadcast effectiveness and Driver response behaviour reports count them separately, so collapsing any two makes those reports lie.

---

## 12. What an ADMIN and a VIEWER do

**Fleet Admin** works in the SharePoint lists directly — there is no admin web part in phase 1.

- `DriversMaster` — add and retire drivers. `UserAccount` is the blocking field: a driver without it can never be offered a trip
- `VehicleMaster` — add and retire vehicles. `VehicleType` is mandatory for matching
- `AvailabilityBlocks` — leave, shifts, maintenance. Exactly one of Driver or Vehicle per row
- `VehicleTypes`, `Locations` — the pickers
- `AllocationRules` — window, cap, advance release, Automatic/Manual, per property per request type. **A list edit, never a redeployment**
- `Templates` — email wording, so it changes without a deployment

Admins must **not** hand-edit `DriversMaster.Status` or `VehicleMaster.VehStatus`. Those are written by the engine; the nightly reconciliation will overwrite guesses.

**Management Viewer** reads Reports across all properties and writes nothing. *Driver response behaviour* is hidden from them — it exposes individual acceptance rates and is restricted to Admin and Coordinator.

---

## 13. Permissions — what changed, and why it matters

Previously a flow wrote on the driver's behalf under a service identity. **Now the browser writes as the driver**, and there is no service identity in a browser.

Drivers therefore need **Contribute** on `RideOffers`, `TripAssignments` and `TripStatusHistory`.

A 403 on a write returns `NOT_PERMITTED` with a message naming the list — *"You do not have permission to write to the RideOffers list. Ask a Fleet Admin to grant your group Contribute on RideOffers. Nothing was saved."* No retry, no fallback, because there is nothing to fall back to.

**The honest trade-off:** a driver with Contribute on `RideOffers` can, through the REST API, write to another driver's offer row. The interface prevents it; the permission model does not. The compensating controls are that `TripStatusHistory` and `NotificationLog` are append-only and record every change, plus site audit logging. This is a business decision to accept in writing before go-live, not a technical oversight.

---

## 14. Timing summary

| Thing | How fast | Why |
| --- | --- | --- |
| Request submitted → offers exist | Immediate | Browser does it inline |
| Offer appears on an **open** dashboard | ≤20 s | Poll interval |
| Offer email reaches a driver | Up to ~5 min | Polling trigger |
| Accept → confirmed | Immediate | ETag lock, one round trip |
| Status tap → requester's screen | ≤30 s | My Bookings poll |
| Dispatch dashboard refresh | 60 s | |
| Offer expiry | Up to 5 min late | Flow |
| Request → Completed after the tap | Up to 5 min | Flow |
| Advance booking released | 4 h before the trip, **then** the next dashboard poll | Flow + sweep |

---

## 15. Open questions still outstanding

1. **Request ID format** — `FB-0001` or `FB-20260925-0001`? Both implemented; `formatRequestRef()` is the single place to change it. **Settle before the first production request** — it cannot change afterwards without reissuing every identifier.
2. **Are shifts mandatory?** A driver with no `Shift` rows is currently treated as **always available**. The opposite reading is equally defensible and the business has not decided. Flagged in a comment at the eligibility check.
3. **What does `DLClientMaster.DriverID` point at?** (v2 open question 5.) The code joins it to `DriversMaster`. If it points somewhere else, the eligibility filter needs a different join and **no driver is ever eligible** — which looks like a logic bug but is a lookup target.
4. **Localisation** — strings are inline in components, not in resource files. Arabic RTL is designed for but not built.

---

## 16. Where things live

| | |
| --- | --- |
| `src/common/services/DispatchService.ts` | Every write. The engine |
| `src/common/rules/dispatchRules.ts` | Eligibility, free vehicles, transitions — pure, tested |
| `src/common/rules/identifiers.ts` | `FB-nnnn` / `TR-nnnn` |
| `src/common/services/mappers.ts` | List columns → typed models. **Where the real column names live** |
| `src/common/models/Enums.ts` | The exact status strings |
| `src/common/styles/_tokens.scss` | Every colour, font, radius, shadow — the whole re-skin |
| `SharePoint-Lists-Specification.txt` | The list schema |
| `TESTING.md` | How to stand it up and test it |
