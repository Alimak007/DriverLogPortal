/**
 * The dispatch rules, tested without SharePoint.
 *
 * These are the pieces that decide who gets offered work and what a driver is
 * allowed to do next. They are pure functions precisely so they can be tested
 * like this - no list, no network, no mocking framework.
 */

import {
  freeVehicles,
  IEligibilityContext,
  isDriverEligible,
  isLegalTransition,
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

// -----------------------------------------------------------------------------
// Builders - a valid object by default, overridden per test
// -----------------------------------------------------------------------------

const EPOCH = new Date('2026-10-05T06:00:00Z');
const WINDOW_TO = new Date('2026-10-05T10:00:00Z');

function driver(over: Partial<IDriver> = {}): IDriver {
  return {
    id: 1,
    created: EPOCH,
    modified: EPOCH,
    driverName: 'Rashid Al Mansoori',
    qualifiedVehicleTypes: [{ id: 10, title: 'Patrol' }],
    status: 'Available',
    maxConcurrentTrips: 1,
    isActive: true,
    ...over,
  };
}

function block(over: Partial<IAvailabilityBlock> = {}): IAvailabilityBlock {
  return {
    id: 100,
    created: EPOCH,
    modified: EPOCH,
    title: 'block',
    fromDateTime: new Date('2026-10-05T00:00:00Z'),
    toDateTime: new Date('2026-10-06T00:00:00Z'),
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

function vehicle(over: Partial<IVehicle> = {}): IVehicle {
  return {
    id: 300,
    created: EPOCH,
    modified: EPOCH,
    plateNo: 'A 12345',
    vehicleType: { id: 10, title: 'Patrol' },
    status: 'Available',
    ...over,
  };
}

function ctx(over: Partial<IEligibilityContext> = {}): IEligibilityContext {
  return {
    allocatedDriverIds: new Set<number>([1]),
    vehicleTypeId: 10,
    windowTo: WINDOW_TO,
    blocks: [],
    openTrips: [],
    ruleMaxConcurrent: 1,
    ...over,
  };
}

// -----------------------------------------------------------------------------
// Eligibility
// -----------------------------------------------------------------------------

describe('isDriverEligible', () => {
  it('accepts a driver who passes every check', () => {
    expect(isDriverEligible(driver(), ctx())).toBe(true);
  });

  it('excludes a driver not allocated to the property', () => {
    expect(
      isDriverEligible(driver(), ctx({ allocatedDriverIds: new Set<number>([99]) })),
    ).toBe(false);
  });

  it('excludes a driver not qualified for the requested vehicle type', () => {
    const vanDriver = driver({ qualifiedVehicleTypes: [{ id: 20, title: 'Cargo Van' }] });
    expect(isDriverEligible(vanDriver, ctx({ vehicleTypeId: 10 }))).toBe(false);
  });

  it('excludes a driver on leave across the window', () => {
    expect(
      isDriverEligible(
        driver(),
        ctx({ blocks: [block({ driver: { id: 1, title: 'Rashid' }, blockType: 'Leave' })] }),
      ),
    ).toBe(false);
  });

  it('does NOT exclude on a Shift block - a shift says when they work', () => {
    expect(
      isDriverEligible(
        driver(),
        ctx({ blocks: [block({ driver: { id: 1, title: 'Rashid' }, blockType: 'Shift' })] }),
      ),
    ).toBe(true);
  });

  it('ignores a block belonging to a different driver', () => {
    expect(
      isDriverEligible(
        driver(),
        ctx({ blocks: [block({ driver: { id: 2, title: 'Someone else' } })] }),
      ),
    ).toBe(true);
  });

  // The case that catches people out: a licence valid at pickup but not at
  // drop-off must exclude.
  it('excludes a licence that expires DURING the trip window', () => {
    const expiring = driver({ licenceExpiry: new Date('2026-10-05T08:00:00Z') });
    expect(isDriverEligible(expiring, ctx({ windowTo: WINDOW_TO }))).toBe(false);
  });

  it('accepts a licence that is valid to the end of the window', () => {
    const ok = driver({ licenceExpiry: new Date('2026-10-05T10:00:00Z') });
    expect(isDriverEligible(ok, ctx({ windowTo: WINDOW_TO }))).toBe(true);
  });

  it('accepts a driver with no licence expiry recorded', () => {
    expect(isDriverEligible(driver({ licenceExpiry: undefined }), ctx())).toBe(true);
  });

  it('excludes a driver already at their own concurrent cap', () => {
    expect(
      isDriverEligible(
        driver({ maxConcurrentTrips: 1 }),
        ctx({ openTrips: [trip({ driver: { id: 1, title: 'Rashid' } })] }),
      ),
    ).toBe(false);
  });

  it("honours the driver's own cap above the rule default", () => {
    expect(
      isDriverEligible(
        driver({ maxConcurrentTrips: 2 }),
        ctx({
          ruleMaxConcurrent: 1,
          openTrips: [trip({ driver: { id: 1, title: 'Rashid' } })],
        }),
      ),
    ).toBe(true);
  });

  it('does not count another driver\'s trips against this one', () => {
    expect(
      isDriverEligible(
        driver(),
        ctx({ openTrips: [trip({ driver: { id: 2, title: 'Other' } })] }),
      ),
    ).toBe(true);
  });
});

// -----------------------------------------------------------------------------
// Vehicle availability
// -----------------------------------------------------------------------------

describe('freeVehicles', () => {
  it('returns a vehicle that is on no trip and has no block', () => {
    expect(freeVehicles([vehicle()], [], []).length).toBe(1);
  });

  it('excludes a vehicle bound to an open trip', () => {
    const busy = [trip({ vehicle: { id: 300, title: 'A 12345' } })];
    expect(freeVehicles([vehicle()], busy, []).length).toBe(0);
  });

  // The whole point of checking blocks as well as VehStatus: a vehicle booked
  // into the garage through a block, still reading Available, must not be
  // offered and bound.
  it('excludes a vehicle blocked for maintenance even when its status says Available', () => {
    const blocks = [block({ vehicle: { id: 300, title: 'A 12345' }, blockType: 'Maintenance' })];
    expect(freeVehicles([vehicle({ status: 'Available' })], [], blocks).length).toBe(0);
  });

  it('does not exclude on a Shift block', () => {
    const blocks = [block({ vehicle: { id: 300, title: 'A 12345' }, blockType: 'Shift' })];
    expect(freeVehicles([vehicle()], [], blocks).length).toBe(1);
  });

  it('ignores a block belonging to a different vehicle', () => {
    const blocks = [block({ vehicle: { id: 999, title: 'Z 99999' } })];
    expect(freeVehicles([vehicle()], [], blocks).length).toBe(1);
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

  // Cancellation is NOT a transition - it goes through cancel(), which has
  // different side effects per caller.
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
