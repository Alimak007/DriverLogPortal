# Testing the Fleet Booking and Dispatch Portal

Written for: whoever is standing the solution up on `travel-dev` for the first time.

---

## What you can and cannot test yet

Be clear about this before you start, or you will chase failures that are not bugs.

| | Testable now | Needs the flows |
| --- | --- | --- |
| Every screen renders, in every state | ✅ | |
| Booking wizard, all seven screens | ✅ | |
| Submitting a request (the one direct write) | ✅ | |
| Grids, filters, counters, availability, reports, export | ✅ | |
| Timeline, countdown, offer banner, page title, audible alert | ✅ | |
| Accessibility: keyboard, focus, contrast, screen reader | ✅ | |
| Driver taps **Accept** | | ❌ Claim flow |
| Driver taps **Reject** / the four status taps | | ❌ |
| Coordinator assigns, reassigns, rejects, cancels | | ❌ |
| **First-accept-wins under concurrent claims** | | ❌ — and this is the one that matters most |

`scripts/3-Simulate-Flows.ps1` performs what each flow would do, from PowerShell. That drives the data through every state so you can check the UI. **It does not test the race condition.** Two drivers tapping Accept in the same second can only be tested against the real Claim flow with its ETag check and concurrency control set to 1.

---

## Setup

### 1. Install PnP PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 2. Create the lists

```powershell
cd "C:\Users\Alimak\Documents\Office Docs\AMC\AMC Fleet Dashboard\DriverLogPortal\scripts"

.\1-Provision-Lists.ps1 `
    -SiteUrl   "https://pomae.sharepoint.com/sites/travel-dev" `
    -TesterUpn "aliakbar.makada@queinfotech.com"
```

Safe to re-run — it creates only what is missing.

`-TesterUpn` puts you in **all five** groups so one account can see every role. That is a test-only shortcut; in production nobody holds all five, and the permission table in section 5 of the spec is what actually enforces the rules.

> **If `VehicleMaster` or `DriversMaster` already hold real data, stop.** The script will add the new columns but will not re-key `Title` or populate anything. Use the migration in section 3 of `SharePoint-Lists-Specification.txt` instead, and take a list export first.

### 3. Load the dummy data

```powershell
.\2-Seed-TestData.ps1 `
    -SiteUrl   "https://pomae.sharepoint.com/sites/travel-dev" `
    -TesterUpn "aliakbar.makada@queinfotech.com"
```

Add `-OtherDriverUpns "a@x.com","b@x.com"` if you have colleagues to test the race with later. Add `-Force` to wipe the transactional lists and reseed.

### 4. Deploy and place the web parts

Build a clean package — `heft clean` first, or it ships stale bundles from every previous build:

```bash
npx heft clean
npx heft build --production
npx heft package-solution --production
```

Upload `sharepoint/solution/fleet-booking-dispatch.sppkg` to the tenant App Catalogue and tick **Make this solution available to all sites**.

Then create five pages on `travel-dev` — one web part each. Don't put all five on one page: they each poll, and the Driver Dashboard owns the browser tab title.

| Page | Web part |
| --- | --- |
| `Book.aspx` | Book a Ride |
| `MyBookings.aspx` | My Bookings |
| `Driver.aspx` | Driver Dashboard |
| `Dispatch.aspx` | Dispatch Dashboard |
| `Reports.aspx` | Fleet Reports |

Leave the flow URL properties blank for now. Reads work; writes fail with a message naming the missing URL, which is the designed behaviour, not a crash.

---

## What you should see immediately

Run this any time to check the state of the data:

```powershell
.\3-Simulate-Flows.ps1 -SiteUrl $url -Action Status
```

| Screen | Expected |
| --- | --- |
| **Dispatch** | Ongoing 1 · Pending 3 · Completed 1 · Drivers 3 · Vehicles 5 · Pairs 1 |
| | `FB-0003` pinned top, amber, badge reads "Awaiting Coordinator" |
| **Driver** | Dark banner, "1 open ride offer", live countdown on `FB-0002` |
| | One trip card, `FB-0004`, En Route, prompting for the closing odometer |
| **My Bookings** | 6 rows, one per status |
| **Book a Ride** | 10 destinations, vehicle tiles filtering by request type |
| **Reports** | Trip summary over this month returns 2 trips |

---

## The dummy data, and why each row is there

Nothing in the seed is filler. Each row proves something.

### Properties and people

| | |
| --- | --- |
| **Dubai Head Office**, **Abu Dhabi Site** | Two properties, so you can prove scoping |
| **Rashid Al Mansoori** | Your account. Allocated to **both** properties — the case SharePoint groups cannot express, which is why scoping comes from `DLClientMaster`. Cap of 2, so he can hold two trips |
| Suresh Nair | Second Dubai passenger driver — gives the broadcast someone to race against |
| Mohammed Iqbal | Van driver. Qualified for vans only, so he must *not* appear on a Patrol request |
| Jomon Varghese | Abu Dhabi only. Must *not* appear on a Dubai request |
| **Ali Hassan** | On leave, with an `AvailabilityBlocks` row covering today. Proves the leave check |
| **Arun Prasad** | Licence expires in **10 days**. Proves the licence check — book a trip 3 weeks out and he must drop off the eligible list |

### Vehicles

Eight, one per type. **E 56789 is Under Maintenance** with a matching block, so "Available vehicles" must read 5 and not 6. **A 12345** is Rashid's default vehicle — the claim logic should prefer it when he accepts.

### The six requests

| Ref | Status | Proves |
| --- | --- | --- |
| `FB-0001` | Submitted | Sits in Pending, not yet broadcast |
| `FB-0002` | **Broadcasting** | Open offers → the Driver Dashboard banner, countdown, Accept/Reject |
| `FB-0003` | **Awaiting Coordinator** | Escalation: pinned top, with both declines and their reasons visible in the assignment panel |
| `FB-0004` | Allocated | Trip **En Route**, ETA 12 min → Ongoing grid, Track my ride, odometer prompt |
| `FB-0005` | Completed | 42 km recorded → counters and every report |
| `FB-0006` | Cancelled | Cancelled badge and reason |

`FB-0003` carries two *different* rejection reasons — one from the dropdown, one "Other" with notes — because the escalation queue has to show the Coordinator why nobody took it.

**`FB-0002`'s window is set 30 minutes, not the real 5**, so it doesn't expire while you're reading. Re-open it any time:

```powershell
.\3-Simulate-Flows.ps1 -SiteUrl $url -Action Broadcast -RequestId 2
```

---

## Walkthroughs

### A. Book a ride end to end

1. Open `Book.aspx`. Pickup should default to **Dubai Head Office**.
2. Type "mar" in Destination — the list filters to Marina Staff Residence. **Confirm you cannot type a destination that isn't on the list.**
3. Continue → today's past time slots are disabled.
4. Pick **Items or Documents** → step 4 offers only the four vans and Others. Pick **Client or Client Guest** instead → only Patrol, Pathfinder, Kicks / Creta and Others.
5. Pick **Others** → one free-text box appears. That is the only one in the whole form.
6. Submit. You land on **Processing**, which says *received*, not *confirmed* — these must not be conflated.
7. Check the new row in `TransportRequests`: `RequestId` is `FB-00nn` matching the item ID.

Then make the driver accept it, and watch the page move to Booking confirmed on its own:

```powershell
.\3-Simulate-Flows.ps1 -SiteUrl $url -Action Broadcast -RequestId <newId>
.\3-Simulate-Flows.ps1 -SiteUrl $url -Action Claim -RequestId <newId> -DriverName "Rashid Al Mansoori"
```

### B. Driver dashboard

1. Open `Driver.aspx` as yourself. Banner shows **1**, countdown ticking.
2. **Background the tab.** The browser tab title reads `(1) New ride offer`. That's the third of the three mitigations for having no push notification.
3. Press **Sound: on**, then re-broadcast from PowerShell. A chime fires when the count rises.
4. Confirm the offer card shows the requester's **name but not their phone number** — that is released only on acceptance.
5. Let the window run out, or force it:
   ```powershell
   .\3-Simulate-Flows.ps1 -SiteUrl $url -Action Expire -RequestId 2
   ```
   The card disappears and the request turns up on the Dispatch dashboard, pinned.

### C. Dispatch dashboard

1. `FB-0003` is pinned top with an amber left edge **and** the words "Awaiting Coordinator".
2. Click **Assign** on it. The panel lists every driver with a verdict — Mohammed Iqbal should read *not qualified for this vehicle type* if the request is for a Cargo Van he can drive, Ali Hassan *on leave during the trip window*, Arun Prasad flagged if the window is past his licence expiry.
3. Pick an ineligible driver → an **override reason box appears and Assign stays disabled** until you fill it.
4. Click a counter — the grids filter. Click again — they unfilter.
5. Availability tab → Vehicles. E 56789 reads **Under Maintenance**.

### D. Reports

1. Run **Trip summary** for this month → 2 trips, one with 42 km.
2. Run **Broadcast effectiveness** → escalation rate, median seconds to first accept.
3. **Export to Excel** and **Export to CSV**. Open both. The CSV must not mangle the `/` in "Kicks / Creta".
4. Set the range to 13 months → blocked with a message before anything runs.
5. **Driver response behaviour** should be visible to you only because you're in FM Admins/Coordinators. Remove yourself from both groups, reload, and confirm it disappears.

---

## Accessibility pass

The NFRs ask for WCAG 2.1 AA, verified by an Axe scan plus a manual keyboard and screen-reader pass.

- **Keyboard only, no mouse.** Tab through Dispatch. Every counter, filter header, tab and row action must be reachable, with a visible 3px blue focus ring. Grids are focusable regions.
- **Column filter**: open with Enter, close with Escape.
- **Colour alone**: screenshot a dashboard and desaturate it. Every status must still be readable, because every badge carries its word.
- **360px**: in DevTools device mode. No horizontal page scroll; grids scroll inside themselves.
- **Screen reader**: each grid announces its caption; counters announce "label: value. Show these."

---

## When the flows arrive

1. Paste the seven trigger URLs into each web part's property pane.
2. Re-test every action marked ❌ in the first table — this time through the buttons.
3. **Test the race properly.** Two people, two devices, same open offer, both tap Accept on a countdown. One gets the trip; the other must see **"Taken by another driver"** immediately and the card must vanish. A driver who taps Accept and sees nothing happen stops trusting the system within a week.
4. Then the load test: 10 simultaneous claims on one request, zero double allocations.

---

## Known gaps while testing

- **List permissions are not applied.** Every group inherits the site's. A Requester will currently see other people's bookings. Apply section 5 of the spec plus item-level permissions on `TransportRequests` before this carries anything real.
- **`DLClientMaster` / `PropertyMaster` column names are assumed** (`Driver`, `Property`). The provisioning script creates them that way, so a clean dev site is consistent. On a site where those lists already exist with different names, driver scoping silently falls back to `HomeProperty` and logs a console warning naming the spec file — **check the browser console** for it.
- **No `Daily Trip Sheet` write.** The Complete flow owns that; the simulator says so when it completes a trip.
- **`RequestId` format is still open** — `FB-0001` vs `FB-20260925-0001`. The seed uses the short form. Settle it before any production request exists; it can't be changed afterwards without reissuing every identifier.
