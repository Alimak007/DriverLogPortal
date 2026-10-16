/**
 * The dispatch rules, tested without SharePoint.
 *
 * These decide who gets offered work and what a driver may do next. They are
 * pure functions precisely so they can be tested like this - no list, no
 * network, no mocking framework.
 */

import {
  distanceKm,
  eligiblePairs,
  explainPairIneligibility,
  IEligibilityContext,
  isLegalTransition,
  isOutsideOperatingArea,
  isPairEligible,
} from './dispatchRules';
import {
  IAvailabilityBlock,
  IDriver,
  ITripAssignment,
  IVehicle,
  TRIP_STATUS,
  TripStatus,
} from '../models';
import { formatRequestRef, formatTripRef } from './identifiers';
import { ActorRole, odometerPromptFor } from './statusTransitions';
import { FleetWebPartKey, mayOpen } from '../webpart/access';

// -----------------------------------------------------------------------------
// Builders - a valid object by default, overridden per test
// -----------------------------------------------------------------------------

const EPOCH = new Date('2026-10-13T06:00:00Z');
const WINDOW_TO = new Date('2026-10-13T10:00:00Z');

function driver(over: Partial<IDriver> = {}): IDriver {
  return {
    id: 1,
    created: EPOCH,
    modified: EPOCH,
    driverName: 'Aabid',
    qualifiedVehicleTypes: [],
    status: 'Available',
    maxConcurrentTrips: 1,
    isActive: true,
    ...over,
  };
}

function vehicle(over: Partial<IVehicle> = {}): IVehicle {
  return {
    id: 300,
    created: EPOCH,
    modified: EPOCH,
    plateNo: 'AUH-12345',
    vehicleType: { id: 10, title: 'Kicks / Creta' },
    defaultDriver: { id: 1, title: 'Aabid' },
    status: 'Available',
    ...over,
  };
}

function block(over: Partial<IAvailabilityBlock> = {}): IAvailabilityBlock {
  return {
    id: 100,
    created: EPOCH,
    modified: EPOCH,
    title: 'block',
    fromDateTime: new Date('2026-10-13T00:00:00Z'),
    toDateTime: new Date('2026-10-14T00:00:00Z'),
    blockType: 'Leave',
    ...over,
  };
}

function trip(over: Partial<ITripAssignment> = {}): ITripAssignment {
  return {
    id: 200,
    created: EPOCH,
    modified: EPOCH,
    tripId: 'TR-0200',
    acceptedOn: EPOCH,
    tripStatus: TRIP_STATUS.Assigned,
    allocationMode: 'Automatic',
    isOverride: false,
    ...over,
  };
}

function ctx(over: Partial<IEligibilityContext> = {}): IEligibilityContext {
  return {
    driversById: new Map<number, IDriver>([[1, driver()]]),
    blocks: [],
    openTrips: [],
    windowTo: WINDOW_TO,
    ruleMaxConcurrent: 1,
    ...over,
  };
}

// -----------------------------------------------------------------------------
// Pair eligibility
// -----------------------------------------------------------------------------

describe('isPairEligible', () => {
  it('accepts a vehicle whose driver passes every check', () => {
    expect(isPairEligible(vehicle(), ctx())).toBe(true);
  });

  // The pairing is the whole model now: no driver means the vehicle can never
  // be offered, because its driver does not exist.
  it('excludes a vehicle with no DefaultDriver', () => {
    const reasons = explainPairIneligibility(vehicle({ defaultDriver: undefined }), ctx());
    expect(reasons.length).toBe(1);
    expect(reasons[0]).toContain('no DefaultDriver');
  });

  it('excludes a vehicle whose driver is not in the active set', () => {
    const reasons = explainPairIneligibility(
      vehicle({ defaultDriver: { id: 99, title: 'Gone' } }),
      ctx(),
    );
    expect(reasons.length).toBe(1);
    expect(reasons[0]).toContain('inactive or missing');
  });

  it('excludes when the DRIVER is blocked in the window', () => {
    const blocks = [block({ driver: { id: 1, title: 'Aabid' }, blockType: 'Leave' })];
    expect(isPairEligible(vehicle(), ctx({ blocks }))).toBe(false);
  });

  // The check that stops a vehicle booked into the garage being offered while
  // its VehStatus still reads Available.
  it('excludes when the VEHICLE is blocked in the window', () => {
    const blocks = [block({ vehicle: { id: 300, title: 'AUH-12345' }, blockType: 'Maintenance' })];
    expect(isPairEligible(vehicle(), ctx({ blocks }))).toBe(false);
  });

  it('does NOT exclude on a Shift block - a shift says when they work', () => {
    const blocks = [block({ driver: { id: 1, title: 'Aabid' }, blockType: 'Shift' })];
    expect(isPairEligible(vehicle(), ctx({ blocks }))).toBe(true);
  });

  it('ignores a block belonging to a different driver or vehicle', () => {
    const blocks = [
      block({ driver: { id: 2, title: 'Someone else' } }),
      block({ vehicle: { id: 999, title: 'Another van' } }),
    ];
    expect(isPairEligible(vehicle(), ctx({ blocks }))).toBe(true);
  });

  // The case that catches people out: valid at pickup, expired by drop-off.
  it('excludes a licence that expires DURING the trip window', () => {
    const d = driver({ licenceExpiry: new Date('2026-10-13T08:00:00Z') });
    expect(isPairEligible(vehicle(), ctx({ driversById: new Map([[1, d]]) }))).toBe(false);
  });

  it('accepts a licence valid to the end of the window', () => {
    const d = driver({ licenceExpiry: new Date('2026-10-13T10:00:00Z') });
    expect(isPairEligible(vehicle(), ctx({ driversById: new Map([[1, d]]) }))).toBe(true);
  });

  it('accepts a driver with no licence expiry recorded', () => {
    const d = driver({ licenceExpiry: undefined });
    expect(isPairEligible(vehicle(), ctx({ driversById: new Map([[1, d]]) }))).toBe(true);
  });

  it('excludes a driver already at their cap', () => {
    const openTrips = [trip({ driver: { id: 1, title: 'Aabid' } })];
    expect(isPairEligible(vehicle(), ctx({ openTrips }))).toBe(false);
  });

  it("honours the driver's own cap above the rule default", () => {
    const d = driver({ maxConcurrentTrips: 2 });
    const openTrips = [trip({ driver: { id: 1, title: 'Aabid' } })];
    expect(
      isPairEligible(vehicle(), ctx({ driversById: new Map([[1, d]]), openTrips, ruleMaxConcurrent: 1 })),
    ).toBe(true);
  });

  it("does not count another driver's trips against this one", () => {
    const openTrips = [trip({ driver: { id: 2, title: 'Other' } })];
    expect(isPairEligible(vehicle(), ctx({ openTrips }))).toBe(true);
  });

  it('reports every reason, not just the first', () => {
    const d = driver({ licenceExpiry: new Date('2026-10-13T08:00:00Z') });
    const blocks = [block({ driver: { id: 1, title: 'Aabid' } })];
    const openTrips = [trip({ driver: { id: 1, title: 'Aabid' } })];
    const reasons = explainPairIneligibility(
      vehicle(),
      ctx({ driversById: new Map([[1, d]]), blocks, openTrips }),
    );
    expect(reasons.length).toBe(3);
  });
});

// -----------------------------------------------------------------------------
// The pair set - this is the vehicle-type routing test
// -----------------------------------------------------------------------------

describe('eligiblePairs', () => {
  it('pairs each vehicle with its own driver', () => {
    const aabid = driver({ id: 1, driverName: 'Aabid' });
    const mahedi = driver({ id: 3, driverName: 'Mahedi' });
    const vehicles = [
      vehicle({ id: 300, plateNo: 'C-1', defaultDriver: { id: 1, title: 'Aabid' } }),
      vehicle({ id: 302, plateNo: 'C-2', defaultDriver: { id: 3, title: 'Mahedi' } }),
    ];
    const pairs = eligiblePairs(
      vehicles,
      ctx({ driversById: new Map([[1, aabid], [3, mahedi]]) }),
    );
    expect(pairs.length).toBe(2);
    expect(pairs.map((p) => p.driver.driverName).sort()).toEqual(['Aabid', 'Mahedi']);
  });

  it('drops only the ineligible pair, keeping the rest', () => {
    const aabid = driver({ id: 1 });
    const mahedi = driver({ id: 3 });
    const vehicles = [
      vehicle({ id: 300, defaultDriver: { id: 1, title: 'Aabid' } }),
      vehicle({ id: 302, defaultDriver: { id: 3, title: 'Mahedi' } }),
    ];
    const blocks = [block({ driver: { id: 3, title: 'Mahedi' } })];
    const pairs = eligiblePairs(
      vehicles,
      ctx({ driversById: new Map([[1, aabid], [3, mahedi]]), blocks }),
    );
    expect(pairs.length).toBe(1);
    expect(pairs[0].driver.id).toBe(1);
  });

  it('returns nothing when no vehicle of the type has a driver', () => {
    const vehicles = [vehicle({ defaultDriver: undefined })];
    expect(eligiblePairs(vehicles, ctx()).length).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// Distance - what replaced property scoping
// -----------------------------------------------------------------------------

describe('distanceKm', () => {
  it('is zero for the same point', () => {
    expect(Math.round(distanceKm(24.4539, 54.3773, 24.4539, 54.3773))).toBe(0);
  });

  // Abu Dhabi to Dubai is about 120 km as the crow flies.
  it('measures Abu Dhabi to Dubai at roughly 120 km', () => {
    const d = distanceKm(24.4539, 54.3773, 25.2048, 55.2708);
    expect(d).toBeGreaterThan(100);
    expect(d).toBeLessThan(140);
  });

  it('is symmetric', () => {
    const a = distanceKm(24.4539, 54.3773, 25.2048, 55.2708);
    const b = distanceKm(25.2048, 55.2708, 24.4539, 54.3773);
    expect(Math.round(a)).toBe(Math.round(b));
  });
});

describe('isOutsideOperatingArea', () => {
  const area = { centreLat: 24.4539, centreLng: 54.3773, radiusKm: 60 };

  it('accepts a pickup inside the radius', () => {
    expect(isOutsideOperatingArea(24.47, 54.4, area)).toBe(false);
  });

  it('rejects a Dubai pickup against an Abu Dhabi centre', () => {
    expect(isOutsideOperatingArea(25.2048, 55.2708, area)).toBe(true);
  });

  // No coordinates is not the same as "too far" - it must not block a booking.
  it('treats missing coordinates as inside', () => {
    expect(isOutsideOperatingArea(undefined, undefined, area)).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Transitions
// -----------------------------------------------------------------------------

describe('isLegalTransition', () => {
  it('allows the four taps in order', () => {
    expect(isLegalTransition(TRIP_STATUS.Assigned, TRIP_STATUS.EnRoute)).toBe(true);
    expect(isLegalTransition(TRIP_STATUS.EnRoute, TRIP_STATUS.InTransit)).toBe(true);
    expect(isLegalTransition(TRIP_STATUS.InTransit, TRIP_STATUS.Completed)).toBe(true);
  });

  it('rejects skipping a step', () => {
    expect(isLegalTransition(TRIP_STATUS.Assigned, TRIP_STATUS.Completed)).toBe(false);
    expect(isLegalTransition(TRIP_STATUS.Assigned, TRIP_STATUS.InTransit)).toBe(false);
    expect(isLegalTransition(TRIP_STATUS.EnRoute, TRIP_STATUS.Completed)).toBe(false);
  });

  it('rejects going backwards', () => {
    expect(isLegalTransition(TRIP_STATUS.InTransit, TRIP_STATUS.EnRoute)).toBe(false);
    expect(isLegalTransition(TRIP_STATUS.Completed, TRIP_STATUS.InTransit)).toBe(false);
  });

  it('allows nothing out of a finished or cancelled trip', () => {
    const every: TripStatus[] = [
      TRIP_STATUS.Assigned,
      TRIP_STATUS.EnRoute,
      TRIP_STATUS.InTransit,
      TRIP_STATUS.Completed,
      TRIP_STATUS.Cancelled,
    ];
    every.forEach((to) => {
      expect(isLegalTransition(TRIP_STATUS.Completed, to)).toBe(false);
      expect(isLegalTransition(TRIP_STATUS.Cancelled, to)).toBe(false);
    });
  });

  // Cancellation goes through cancel(), which has different side effects per
  // caller - it is not a tap.
  it('does not treat Cancelled as a tap from any state', () => {
    expect(isLegalTransition(TRIP_STATUS.Assigned, TRIP_STATUS.Cancelled)).toBe(false);
    expect(isLegalTransition(TRIP_STATUS.EnRoute, TRIP_STATUS.Cancelled)).toBe(false);
    expect(isLegalTransition(TRIP_STATUS.InTransit, TRIP_STATUS.Cancelled)).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Identifiers
// -----------------------------------------------------------------------------

describe('identifiers', () => {
  it('zero-pads to four digits', () => {
    expect(formatRequestRef(1)).toBe('FB-0001');
    expect(formatRequestRef(42)).toBe('FB-0042');
    expect(formatTripRef(7)).toBe('TR-0007');
  });

  it('does not truncate past four digits', () => {
    expect(formatRequestRef(12345)).toBe('FB-12345');
  });
});

// -----------------------------------------------------------------------------
// Odometer prompting
// -----------------------------------------------------------------------------

describe('odometerPromptFor', () => {
  it('offers the opening reading at En Route, without requiring it', () => {
    const prompt = odometerPromptFor(TRIP_STATUS.EnRoute, false);
    expect(prompt).toEqual({ field: 'opening', required: false });
  });

  it('still only offers the opening reading when one somehow exists already', () => {
    const prompt = odometerPromptFor(TRIP_STATUS.EnRoute, true);
    expect(prompt?.required).toBe(false);
  });

  // The whole point of the all-or-nothing rule: a closing reading with no
  // opening one cannot be subtracted into anything, so it is never asked for.
  it('hides the closing reading when the driver skipped the opening one', () => {
    expect(odometerPromptFor(TRIP_STATUS.Completed, false)).toBeUndefined();
  });

  it('requires the closing reading once an opening one was given', () => {
    const prompt = odometerPromptFor(TRIP_STATUS.Completed, true);
    expect(prompt).toEqual({ field: 'closing', required: true });
  });

  it('asks for nothing on the other transitions', () => {
    expect(odometerPromptFor(TRIP_STATUS.InTransit, true)).toBeUndefined();
    expect(odometerPromptFor(TRIP_STATUS.InTransit, false)).toBeUndefined();
    expect(odometerPromptFor(TRIP_STATUS.Assigned, true)).toBeUndefined();
    expect(odometerPromptFor(TRIP_STATUS.Cancelled, true)).toBeUndefined();
  });
});

// -----------------------------------------------------------------------------
// Web part access
// -----------------------------------------------------------------------------

describe('mayOpen', () => {
  const ALL: FleetWebPartKey[] = [
    'bookARide',
    'myBookings',
    'driverDashboard',
    'dispatchDashboard',
    'reports',
  ];

  /** The keys a role may open, asserted against every key so nothing leaks. */
  const expectExactly = (roles: ActorRole[], allowed: FleetWebPartKey[]): void => {
    ALL.forEach((key) => {
      expect({ key, open: mayOpen(roles, key) }).toEqual({
        key,
        open: allowed.indexOf(key) !== -1,
      });
    });
  };

  it('lets FM Admins open everything', () => {
    expectExactly(['admin'], ALL);
  });

  it('limits FM Coordinators to the dashboard and reports', () => {
    expectExactly(['coordinator'], ['dispatchDashboard', 'reports']);
  });

  it('limits FM Drivers to the driver dashboard', () => {
    expectExactly(['driver'], ['driverDashboard']);
  });

  it('limits FM Requesters to booking and my bookings', () => {
    expectExactly(['requester'], ['bookARide', 'myBookings']);
  });

  // Someone in two groups holds two roles and gets the union, not the
  // intersection - a Coordinator who also drives needs both consoles.
  it('unions the roles of someone in several groups', () => {
    expectExactly(
      ['coordinator', 'driver'],
      ['dispatchDashboard', 'reports', 'driverDashboard'],
    );
  });

  it('opens nothing for no roles at all', () => {
    expectExactly([], []);
    ALL.forEach((key) => expect(mayOpen(undefined, key)).toBe(false));
  });
});
