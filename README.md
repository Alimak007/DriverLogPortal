# Fleet Booking and Dispatch Portal

Phase 1 SPFx solution for the `travel-dev` SharePoint site. A requester books a
vehicle in a few taps; the request is broadcast to every eligible driver at
once, and the first to accept owns the trip. If nobody accepts within the offer
window it escalates to the Transportation Coordinator.

- **Platform:** SPFx 1.23.2, React 17, Fluent UI 8, PnPjs 4, Heft (no gulp)
- **Site:** `https://pomae.sharepoint.com/sites/travel-dev`
- **List specification:** [`SharePoint-Lists-Specification.txt`](SharePoint-Lists-Specification.txt)
- **How it all works:** [`ARCHITECTURE.md`](ARCHITECTURE.md) — start here
- **Testing:** [`TESTING.md`](TESTING.md)
- **Flow build guide:** [`POWER-AUTOMATE-FLOWS.md`](POWER-AUTOMATE-FLOWS.md)
  — ⚠️ STALE. Written before dispatch moved into the web part; it still
  describes seven premium HTTP-triggered flows this tenant cannot run.

## The two rules that shape everything

**SPFx reads, flows write.** Web parts query lists directly with PnPjs for
display, but every state change that matters — allocation, acceptance, status
progression, completion — goes through a Power Automate flow. Business rules
live in one place rather than being duplicated across five web parts, and a
closed browser cannot leave a half-committed allocation behind.

**No web part calls PnPjs directly.** All SharePoint access goes through typed
service classes behind interfaces in [`src/common/services`](src/common/services),
with the list schema expressed as TypeScript models. A later move to Power Apps
or Dataverse is a service-layer change, not a rewrite.

## What is in the package

Five web parts, all responsive from 360px:

| Web part | Audience | Purpose |
| --- | --- | --- |
| Book a Ride | Requester | The seven-screen booking wizard |
| My Bookings | Requester | Request history and the status timeline |
| Driver Dashboard | Driver | Open offers, accept and reject, the four status taps |
| Dispatch Dashboard | Coordinator | Counters, grids, availability, manual override |
| Fleet Reports | Coordinator, Admin, Viewer | Eight reports with Excel and CSV export |

## Source layout

```
src/
  common/
    config/AppConfig.ts        List names, font source, ID format, poll intervals
    styles/_tokens.scss        EVERY colour, font, radius and shadow, defined once
    styles/fonts.ts            Both faces declared once for the whole solution
    models/                    The list schema as TypeScript; Enums.ts holds the
                               exact choice strings the flows depend on
    rules/                     Pure functions, no SharePoint, unit-testable:
                               eligibility, status transitions, identifiers
    services/                  Typed services behind interfaces (IServices.ts)
    components/                Shared vocabulary: buttons, badges, DataGrid,
                               Timeline, ConfirmDialog
    utils/                     Dates (UTC in, Gulf time out), export
    webpart/                   FleetWebPartBase - bootstrap + flow endpoints
  webparts/                    The five web parts
```

## Build

```bash
npm install
npx heft build                       # dev build, type-check and lint
npx heft start                       # local workbench against travel-dev

# Packaging - ALWAYS clean first:
npx heft clean
npx heft build --production
npx heft package-solution --production
# -> sharepoint/solution/fleet-booking-dispatch.sppkg
```

**`heft clean` before packaging is not optional.** `ClientSideAssets` accumulates
a hashed bundle per build, and `package-solution` ships whatever it finds there.
Skip the clean and the package carries every stale bundle from every previous
build — it grows without bound and the deployed solution serves bytes nobody
asked for. A clean package is ~400 KB with exactly five `.js` files; if yours is
larger, check for duplicates:

```bash
python -c "import zipfile;z=zipfile.ZipFile('sharepoint/solution/fleet-booking-dispatch.sppkg');print('\n'.join(sorted(n for n in z.namelist() if n.endswith('.js'))))"
```

Production bundle sizes, against the 500 KB gzipped budget:

| Bundle | Gzipped |
| --- | --- |
| book-a-ride | 40 KB |
| my-bookings | 98 KB |
| driver-dashboard | 98 KB |
| dispatch-dashboard | 101 KB |
| reports | 37 KB |

Fluent is imported by path (`@fluentui/react/lib/Dialog`), never from the
package root, and there is no Moment. Keep it that way.

## Before this runs against real data

1. **Provision the lists.** Work through
   [`SharePoint-Lists-Specification.txt`](SharePoint-Lists-Specification.txt).
   The two refactored lists (`VehicleMaster`, `DriversMaster`) need their
   documented, reversible migration, and `DriversMaster.UserAccount` must be
   populated for every driver or nobody can be offered a trip.
2. **Build the five flows** — Broadcast, Claim, Expire, Notify, Complete — and
   paste each HTTP trigger URL into the web part property pane. Until then,
   reads work and every write fails with a message naming which URL is missing.
3. **Confirm the `DLClientMaster` and `PropertyMaster` column internal names.**
   The requirements did not supply them. The service layer assumes `Driver` and
   `Property`; if they differ, driver property scoping silently falls back to
   `HomeProperty` and logs a console warning.
4. **Answer the open questions** in section 8 of the list specification. The
   request ID format in particular cannot be changed after the first production
   request is created.

## Deliberately not built

- **The Power Automate flows.** They are the write path and are built in Power
  Automate, not in this repo. The service layer calls them by name through
  `FlowService`. Step-by-step build instructions, with the exact request and
  response contracts the web parts expect, are in
  [`POWER-AUTOMATE-FLOWS.md`](POWER-AUTOMATE-FLOWS.md).
- **The SPFx application customiser** for the header and SharePoint chrome
  hiding. The requirements specify five web parts and no extension. The two
  body classes it would key off are already added by `FleetWebPartBase`, so
  adding it later needs no web part change.
- **String resource files.** The non-functional requirements ask for all
  strings in resource files with Arabic right-to-left supported from day one.
  User-facing strings are currently inline in the components. This is real
  outstanding work, not an oversight — see the Arabic open question, which
  should be answered before the extraction is done, since it determines whether
  the layout work happens at the same time.
- **Jest unit tests.** The requirements set a 70 percent coverage floor on the
  service and rules layers. `src/common/rules` is written as pure functions
  with no SharePoint dependency precisely so those tests can be written
  against it directly.

## Conventions worth knowing

- **Reference a token, never repeat a literal.** A component that writes
  `#452A24` instead of `$brand-dark` is a defect. A full re-skin is
  `_tokens.scss` alone.
- **Colour never carries meaning on its own.** Every badge carries its label,
  every tile says "Selected", a filtered column shows a funnel mark and a
  count. This is an accessibility requirement and the reason the palette can
  change without the interface becoming ambiguous.
- **No silent catch blocks.** Every failure surfaces a message naming what
  failed and what to do, plus a correlation ID the user can quote.
- **The booking web part never reads availability.** Not driver, not vehicle.
  Anything it read would be reachable by that user through the REST API
  regardless of what the interface displays.
