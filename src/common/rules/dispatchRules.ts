/**
 * The dispatch rules: which vehicle-and-driver pairs may be offered a trip,
 * how far the pickup is from the operating area, and what a driver may do
 * next.
 *
 * PURE FUNCTIONS, NO SHAREPOINT AND NO SPFx IMPORTS. That is deliberate and
 * load-bearing - importing anything from @microsoft/sp-* here would drag the
 * whole SPFx runtime into the Jest process and these could not be tested at
 * all. Keep this file dependency-free.
 *
 * THE PAIRING IS PERMANENT. One driver owns one vehicle, recorded on
 * VehicleMaster.DefaultDriver. That single fact removes three checks that used
 * to live here: the property filter, the vehicle-type qualification, and the
 * search for a free vehicle. The vehicle carries the type, and vehicle
 * availability IS driver availability.
 */

import {
  IAvailabilityBlock,
  IDriver,
  ITripAssignment,
  IVehicle,
  TRIP_STATUS,
  TripStatus,
} from '../models';

// -----------------------------------------------------------------------------
// Distance
// -----------------------------------------------------------------------------

/**
 * Great-circle distance in kilometres.
 *
 * This replaces property scoping. Nothing else stops a requester dropping a
 * pin in another emirate, so without it the engine would cheerfully offer a
 * trip to a driver 140 km away.
 */
export function distanceKm(
  lat1: number,
  lng1: number,
  lat2: number,
  lng2: number,
): number {
  const R = 6371;
  const toRad = (d: number): number => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

export interface IOperatingArea {
  centreLat: number;
  centreLng: number;
  radiusKm: number;
}

/** True when the pickup is too far out to broadcast automatically. */
export function isOutsideOperatingArea(
  pickupLat: number | undefined,
  pickupLng: number | undefined,
  area: IOperatingArea,
): boolean {
  if (pickupLat === undefined || pickupLng === undefined) return false;
  return distanceKm(pickupLat, pickupLng, area.centreLat, area.centreLng) > area.radiusKm;
}

// -----------------------------------------------------------------------------
// Eligibility
// -----------------------------------------------------------------------------

export interface IEligibilityContext {
  /** Every active driver, by id. The vehicle's DefaultDriver is looked up here. */
  driversById: Map<number, IDriver>;
  /** Blocks overlapping the trip window, for drivers AND vehicles. */
  blocks: IAvailabilityBlock[];
  /** Trips that are not yet Completed or Cancelled. */
  openTrips: ITripAssignment[];
  /** End of the trip window. Licences are checked against this, not the start. */
  windowTo: Date;
  /** The rule's cap, used when a driver has no value of their own. */
  ruleMaxConcurrent: number;
}

/** A vehicle and the driver who permanently owns it. */
export interface IEligiblePair {
  vehicle: IVehicle;
  driver: IDriver;
}

/**
 * Why this vehicle cannot be offered. An empty array means it can.
 *
 * Returning reasons rather than a bare boolean is what makes an escalation
 * diagnosable: "no eligible drivers" is useless, "4 vehicles have no assigned
 * driver, 2 licences expired" points straight at the data.
 */
export function explainPairIneligibility(
  vehicle: IVehicle,
  ctx: IEligibilityContext,
): string[] {
  const reasons: string[] = [];

  // 1. The vehicle must have a driver. Without one it can never be offered -
  //    its driver does not exist.
  const driverId = vehicle.defaultDriver?.id;
  if (driverId === undefined) {
    return ['no DefaultDriver assigned - the pairing is required'];
  }

  // 2. That driver must be active.
  const driver = ctx.driversById.get(driverId);
  if (!driver) {
    return [`assigned driver (id ${driverId}) is inactive or missing from DriversMaster`];
  }

  // 3 and 4. A block stops the pair whether it names the driver or the vehicle.
  //
  // A Shift row says when a driver WORKS, not when they are away, so it never
  // excludes anyone.
  //
  // OPEN QUESTION (v2 spec, question 6): a driver with no Shift rows at all is
  // treated as ALWAYS AVAILABLE. The opposite reading - no shift means not
  // working - is equally defensible and the business has still not decided.
  // Changing it is a one-line change here.
  if (ctx.blocks.some((b) => b.driver?.id === driver.id && b.blockType !== 'Shift')) {
    reasons.push('driver is on leave or otherwise blocked during the trip window');
  }
  if (ctx.blocks.some((b) => b.vehicle?.id === vehicle.id && b.blockType !== 'Shift')) {
    reasons.push('vehicle is blocked during the trip window');
  }

  // 5. Against the END of the window. A licence that expires mid-trip excludes.
  if (driver.licenceExpiry && driver.licenceExpiry < ctx.windowTo) {
    reasons.push('licence expires before the end of the trip window');
  }

  // 6. Open trips below the cap.
  //
  // OPEN QUESTION - and a real bug above a cap of 1. This counts the driver's
  // open trips; it does NOT check whether they overlap in time. With a cap of
  // 2 a driver can hold two 09:00 trips, and because the pairing is permanent
  // that means one vehicle in two places at once.
  //
  // Default the cap to 1, which is safe. The proper fix is a window-overlap
  // test alongside the count, which would make a cap above 1 genuinely useful
  // - a driver holding a 09:00 and a 14:00 booking is a real operational need.
  // Not built here: it needs a business decision about pre-booking.
  const held = ctx.openTrips.filter((t) => t.driver?.id === driver.id).length;
  const cap = driver.maxConcurrentTrips ?? ctx.ruleMaxConcurrent;
  if (held >= cap) {
    reasons.push(`driver is at the concurrent-trip cap (holding ${held} of ${cap})`);
  }

  return reasons;
}

/** All five checks must hold. */
export function isPairEligible(vehicle: IVehicle, ctx: IEligibilityContext): boolean {
  return explainPairIneligibility(vehicle, ctx).length === 0;
}

/**
 * The vehicle-and-driver pairs that may be offered this trip.
 *
 * `vehicles` is already filtered to the requested type by the query - the
 * vehicle carries the type now, so there is no qualification check here.
 */
export function eligiblePairs(
  vehicles: IVehicle[],
  ctx: IEligibilityContext,
): IEligiblePair[] {
  const pairs: IEligiblePair[] = [];
  vehicles.forEach((v) => {
    if (!isPairEligible(v, ctx)) return;
    const driver = ctx.driversById.get(v.defaultDriver?.id ?? -1);
    if (driver) pairs.push({ vehicle: v, driver });
  });
  return pairs;
}

// -----------------------------------------------------------------------------
// Transitions
// -----------------------------------------------------------------------------

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
