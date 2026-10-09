/**
 * The dispatch rules: who may be offered work, which vehicles are free, and
 * what a driver may do next.
 *
 * PURE FUNCTIONS, NO SHAREPOINT AND NO SPFx IMPORTS. That is deliberate and
 * load-bearing - importing anything from @microsoft/sp-* here would drag the
 * whole SPFx runtime into the Jest process and these could not be tested at
 * all. Keep this file dependency-free.
 */

import {
  IAvailabilityBlock,
  IDriver,
  ITripAssignment,
  IVehicle,
  TRIP_STATUS,
  TripStatus,
} from '../models';

export interface IEligibilityContext {
  allocatedDriverIds: Set<number>;
  vehicleTypeId: number;
  windowTo: Date;
  blocks: IAvailabilityBlock[];
  openTrips: ITripAssignment[];
  ruleMaxConcurrent: number;
}

/**
 * Why this driver is NOT eligible. An empty array means they are.
 *
 * Returning reasons rather than a bare boolean is deliberate: when a broadcast
 * finds nobody, "no eligible drivers" is useless and "6 not allocated to this
 * property, 2 not qualified" points straight at the data.
 */
export function explainDriverIneligibility(
  driver: IDriver,
  ctx: IEligibilityContext,
): string[] {
  const reasons: string[] = [];

  if (!ctx.allocatedDriverIds.has(driver.id)) {
    reasons.push('not allocated to this property in DLClientMaster');
  }

  if (!driver.qualifiedVehicleTypes.some((t) => t.id === ctx.vehicleTypeId)) {
    reasons.push(
      driver.qualifiedVehicleTypes.length === 0
        ? 'QualifiedVehicleTypes is empty'
        : 'not qualified for the requested vehicle type',
    );
  }

  // A Shift row says when a driver WORKS, not when they are away, so it never
  // excludes anyone.
  //
  // OPEN QUESTION (v2 spec, question 6): a driver with no Shift rows at all is
  // currently treated as ALWAYS AVAILABLE. The opposite reading - no shift
  // means not working - is equally defensible and the business has not
  // decided. Changing it is a one-line change here.
  if (ctx.blocks.some((b) => b.driver?.id === driver.id && b.blockType !== 'Shift')) {
    reasons.push('on leave or otherwise blocked during the trip window');
  }

  // Against the END of the window. A licence expiring mid-trip must exclude.
  if (driver.licenceExpiry && driver.licenceExpiry < ctx.windowTo) {
    reasons.push('licence expires before the end of the trip window');
  }

  const held = ctx.openTrips.filter((t) => t.driver?.id === driver.id).length;
  const cap = driver.maxConcurrentTrips ?? ctx.ruleMaxConcurrent;
  if (held >= cap) {
    reasons.push(`at the concurrent-trip cap (holding ${held} of ${cap})`);
  }

  return reasons;
}

/** All five checks must hold. */
export function isDriverEligible(driver: IDriver, ctx: IEligibilityContext): boolean {
  return explainDriverIneligibility(driver, ctx).length === 0;
}

/**
 * A vehicle is free only if it is on no open trip AND has no non-Shift
 * availability block.
 *
 * The block check is not optional: a vehicle booked into the garage through a
 * block, with VehStatus still Available, would otherwise be offered and bound.
 */
export function freeVehicles(
  vehicles: IVehicle[],
  openTrips: ITripAssignment[],
  blocks: IAvailabilityBlock[],
): IVehicle[] {
  const busy = new Set<number>();
  openTrips.forEach((t) => { if (t.vehicle) busy.add(t.vehicle.id); });

  return vehicles.filter((v) => {
    if (busy.has(v.id)) return false;
    return !blocks.some((b) => b.vehicle?.id === v.id && b.blockType !== 'Shift');
  });
}

const LEGAL_TRANSITIONS: Record<string, TripStatus[]> = {
  [TRIP_STATUS.Assigned]: [TRIP_STATUS.EnRoute],
  [TRIP_STATUS.EnRoute]: [TRIP_STATUS.InTransit],
  [TRIP_STATUS.InTransit]: [TRIP_STATUS.Completed],
  [TRIP_STATUS.Completed]: [],
  [TRIP_STATUS.Cancelled]: [],
};

/** Cancellation is NOT in this table - it goes through cancel(). */
export function isLegalTransition(from: TripStatus, to: TripStatus): boolean {
  return (LEGAL_TRANSITIONS[from] ?? []).indexOf(to) !== -1;
}
