/**
 * Driver eligibility for a broadcast.
 *
 * PURE FUNCTIONS, NO SHAREPOINT. The requirements call this out specifically:
 * the eligibility and claim logic must be unit-testable without SharePoint.
 * Everything here takes plain data in and returns plain data out.
 *
 * This mirrors what the Broadcast flow does server-side. The flow is the
 * authority; this exists so the Coordinator's manual-assignment panel can show
 * the same eligibility verdict, with the same reasons, without guessing.
 *
 * THE FILTER RUNS ONCE, AT BROADCAST TIME.
 */

import {
  IAllocationRule,
  IAvailabilityBlock,
  IDriver,
  ITransportRequest,
  ITripAssignment,
  IVehicle,
} from '../models';

export type IneligibilityReason =
  | 'not-allocated-to-property'
  | 'not-qualified-for-vehicle-type'
  | 'inactive'
  | 'on-leave-in-window'
  | 'licence-expires-in-window'
  | 'at-concurrent-trip-cap'
  | 'no-vehicle-of-type-free';

export const REASON_TEXT: Record<IneligibilityReason, string> = {
  'not-allocated-to-property': 'Not allocated to this property',
  'not-qualified-for-vehicle-type': 'Not qualified for this vehicle type',
  'inactive': 'Driver record is inactive',
  'on-leave-in-window': 'On leave during the trip window',
  'licence-expires-in-window': 'Licence expires before the trip ends',
  'at-concurrent-trip-cap': 'Already at the concurrent-trip cap',
  'no-vehicle-of-type-free': 'No vehicle of the requested type is free',
};

export interface IEligibilityVerdict {
  driverId: number;
  eligible: boolean;
  reasons: IneligibilityReason[];
}

export interface ITripWindow {
  from: Date;
  to: Date;
}

export interface IEligibilityInput {
  request: ITransportRequest;
  window: ITripWindow;
  drivers: IDriver[];
  /** Driver id -> the property ids that driver is allocated to. */
  driverProperties: Map<number, number[]>;
  blocks: IAvailabilityBlock[];
  /** Trips that are not yet Completed or Cancelled. */
  openTrips: ITripAssignment[];
  vehicles: IVehicle[];
  rule?: IAllocationRule;
}

/** Two half-open intervals overlap. Touching endpoints do not count. */
export function overlaps(a: ITripWindow, b: ITripWindow): boolean {
  return a.from < b.to && b.from < a.to;
}

/**
 * The trip window. Where no return time is given, the window is the required
 * time plus a default duration - the broadcast still has to reserve SOMETHING.
 */
export function tripWindow(
  request: ITransportRequest,
  defaultDurationMinutes: number = 120,
): ITripWindow {
  const from = request.requiredDateTime;
  const to = request.returnDateTime
    ? request.returnDateTime
    : new Date(from.getTime() + defaultDurationMinutes * 60000);
  return { from, to };
}

function isBlocked(
  blocks: IAvailabilityBlock[],
  window: ITripWindow,
  match: (b: IAvailabilityBlock) => boolean,
): boolean {
  return blocks.some(
    (b) =>
      match(b) &&
      b.blockType !== 'Shift' &&
      overlaps(window, { from: b.fromDateTime, to: b.toDateTime }),
  );
}

/** Open trips a driver currently holds. */
export function openTripCount(
  openTrips: ITripAssignment[],
  driverId: number,
): number {
  return openTrips.filter((t) => t.driver?.id === driverId).length;
}

/**
 * At least one vehicle of the requested type is free for the window.
 *
 * The broadcast confirms only that a suitable vehicle EXISTS. The vehicle is
 * bound when a driver accepts. Binding a specific vehicle to every offer would
 * reserve the fleet against offers that will mostly be declined.
 */
export function freeVehiclesOfType(input: IEligibilityInput): IVehicle[] {
  const typeId = input.request.vehicleTypeRequested?.id;
  const propertyId = input.request.property?.id;
  if (typeId === undefined) return [];

  return input.vehicles.filter((v) => {
    if (v.vehicleType?.id !== typeId) return false;
    if (v.status === 'Retired' || v.status === 'Under Maintenance') return false;
    if (propertyId !== undefined && v.property?.id !== propertyId) return false;

    if (isBlocked(input.blocks, input.window, (b) => b.vehicle?.id === v.id)) {
      return false;
    }
    const busy = input.openTrips.some(
      (t) => t.vehicle?.id === v.id && t.tripStatus !== 'Completed' && t.tripStatus !== 'Cancelled',
    );
    return !busy;
  });
}

/** Assess one driver. Returns every reason, not just the first. */
export function assessDriver(
  driver: IDriver,
  input: IEligibilityInput,
  vehicleAvailable: boolean,
): IEligibilityVerdict {
  const reasons: IneligibilityReason[] = [];

  if (!driver.isActive || driver.status === 'Inactive') {
    reasons.push('inactive');
  }

  const propertyId = input.request.property?.id;
  if (propertyId !== undefined) {
    const allocated = input.driverProperties.get(driver.id) ?? [];
    if (allocated.indexOf(propertyId) === -1) {
      reasons.push('not-allocated-to-property');
    }
  }

  const typeId = input.request.vehicleTypeRequested?.id;
  if (typeId !== undefined) {
    const qualified = driver.qualifiedVehicleTypes.some((t) => t.id === typeId);
    if (!qualified) reasons.push('not-qualified-for-vehicle-type');
  }

  if (isBlocked(input.blocks, input.window, (b) => b.driver?.id === driver.id)) {
    reasons.push('on-leave-in-window');
  }

  // The licence must be valid through the END of the trip window, not just
  // at the start.
  if (driver.licenceExpiry && driver.licenceExpiry < input.window.to) {
    reasons.push('licence-expires-in-window');
  }

  const cap =
    input.rule?.maxConcurrentTripsPerDriver ?? driver.maxConcurrentTrips ?? 1;
  if (openTripCount(input.openTrips, driver.id) >= cap) {
    reasons.push('at-concurrent-trip-cap');
  }

  if (!vehicleAvailable) {
    reasons.push('no-vehicle-of-type-free');
  }

  return { driverId: driver.id, eligible: reasons.length === 0, reasons };
}

/**
 * The eligible driver set for a request.
 *
 * A driver receives the offer only if EVERY check holds.
 */
export function eligibleDrivers(input: IEligibilityInput): IEligibilityVerdict[] {
  const vehicleAvailable = freeVehiclesOfType(input).length > 0;
  return input.drivers.map((d) => assessDriver(d, input, vehicleAvailable));
}

/**
 * Vehicle binding at claim time: prefer the accepting driver's default
 * vehicle if it is free, otherwise the first free vehicle of the right type
 * at that property.
 */
export function bindVehicle(
  input: IEligibilityInput,
  driverId: number,
): IVehicle | undefined {
  const free = freeVehiclesOfType(input);
  if (free.length === 0) return undefined;

  const preferred = free.filter((v) => v.defaultDriver?.id === driverId);
  return preferred.length > 0 ? preferred[0] : free[0];
}
