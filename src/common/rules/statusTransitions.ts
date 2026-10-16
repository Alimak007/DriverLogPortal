/**
 * Which status changes are legal, and who may make them.
 *
 * PURE FUNCTIONS, NO SHAREPOINT.
 *
 * IMPORTANT: this is NOT the control. Status transitions are validated
 * SERVER-SIDE by the flow, which rejects an out-of-order change and checks
 * that the caller is the assigned driver or a Coordinator. "The dashboard
 * hiding a button is not the control; the flow is."
 *
 * This exists so the interface does not OFFER an action that the flow will
 * refuse, which is a usability measure, not a security boundary.
 */

import { RequestStatus, TripStatus } from '../models';

export type ActorRole =
  | 'requester'
  | 'driver'
  | 'coordinator'
  | 'admin'
  | 'viewer';

// -----------------------------------------------------------------------------
// Trip status
// -----------------------------------------------------------------------------

/** The four driver taps, in order. One tap per state change. */
export const TRIP_FLOW: TripStatus[] = [
  'Assigned',
  'En Route',
  'In Transit',
  'Completed',
];

const TRIP_NEXT: Record<TripStatus, TripStatus[]> = {
  'Assigned': ['En Route', 'Cancelled'],
  'En Route': ['In Transit', 'Cancelled'],
  'In Transit': ['Completed', 'Cancelled'],
  'Completed': [],
  'Cancelled': [],
};

export function canTransitionTrip(from: TripStatus, to: TripStatus): boolean {
  return TRIP_NEXT[from].indexOf(to) !== -1;
}

/** The next status a driver tap produces, or undefined if the trip is done. */
export function nextTripStatus(from: TripStatus): TripStatus | undefined {
  const i = TRIP_FLOW.indexOf(from);
  return i >= 0 && i < TRIP_FLOW.length - 1 ? TRIP_FLOW[i + 1] : undefined;
}

/** The label on the button that produces that transition. */
export const TAP_LABEL: Partial<Record<TripStatus, string>> = {
  'En Route': 'On my way',
  'In Transit': 'Picked up',
  'Completed': 'Complete',
};

/**
 * Only the ASSIGNED driver or a Coordinator may change a trip's status. A
 * Coordinator may override any status; a driver may only move forward.
 */
export function mayChangeTripStatus(
  role: ActorRole,
  isAssignedDriver: boolean,
  from: TripStatus,
  to: TripStatus,
): boolean {
  if (role === 'coordinator' || role === 'admin') return true;
  if (role === 'driver') {
    return isAssignedDriver && canTransitionTrip(from, to);
  }
  return false;
}

/** What, if anything, to ask the driver for at a given transition. */
export interface IOdometerPrompt {
  field: 'opening' | 'closing';
  /** False means the driver may tap straight through without typing. */
  required: boolean;
}

/**
 * The odometer is ALL OR NOTHING per trip.
 *
 * Since the booking carries a road distance from the map, a driver no longer
 * has to read the dash at the kerbside at all. But a trip with an opening
 * reading and no closing one is worse than a trip with neither: it looks like
 * real data and cannot be subtracted into anything. So:
 *
 *   - Opening, at "On my way": OFFERED, never required.
 *   - Closing, at "Complete" : only EXISTS if an opening reading was given,
 *                              and is then REQUIRED, because the pair is the
 *                              only thing that yields a distance.
 *
 * A trip with neither reading falls back to the planned distance for the
 * monthly trip sheet. A trip with both is measured, and measured always wins.
 *
 * `hasOpeningReading` is read from the saved trip, not from component state,
 * so the rule survives a reload between the two taps.
 */
export function odometerPromptFor(
  to: TripStatus,
  hasOpeningReading: boolean,
): IOdometerPrompt | undefined {
  if (to === 'En Route') return { field: 'opening', required: false };
  if (to === 'Completed') {
    // No opening reading means there is nothing to close. Hide it entirely
    // rather than showing a box that can only produce an orphan number.
    return hasOpeningReading ? { field: 'closing', required: true } : undefined;
  }
  return undefined;
}

// -----------------------------------------------------------------------------
// Request status
// -----------------------------------------------------------------------------

const REQUEST_NEXT: Record<RequestStatus, RequestStatus[]> = {
  'Submitted': ['Broadcasting', 'Held', 'Allocated', 'Rejected', 'Cancelled'],
  'Held': ['Broadcasting', 'Allocated', 'Rejected', 'Cancelled'],
  'Broadcasting': ['Allocated', 'Awaiting Coordinator', 'Cancelled'],
  'Awaiting Coordinator': ['Allocated', 'Broadcasting', 'Rejected', 'Cancelled'],
  'Allocated': ['Completed', 'Cancelled', 'Awaiting Coordinator'],
  'Completed': [],
  'Rejected': [],
  'Cancelled': [],
};

export function canTransitionRequest(
  from: RequestStatus,
  to: RequestStatus,
): boolean {
  return REQUEST_NEXT[from].indexOf(to) !== -1;
}

/**
 * The requester may cancel while the request is Submitted or Awaiting
 * Coordinator. ONCE A DRIVER HAS ACCEPTED, cancellation is routed to the
 * Coordinator, who cancels on their behalf and notifies the driver.
 */
export function requesterMayCancel(status: RequestStatus): boolean {
  return status === 'Submitted' || status === 'Awaiting Coordinator' || status === 'Held';
}

/** After acceptance the requester asks; the Coordinator acts. */
export function cancellationNeedsCoordinator(status: RequestStatus): boolean {
  return status === 'Allocated' || status === 'Broadcasting';
}

/** Counted by the "Pending requests" dashboard counter. */
export const PENDING_STATUSES: RequestStatus[] = [
  'Submitted',
  'Broadcasting',
  'Awaiting Coordinator',
];

/** Counted by the "Ongoing trips" dashboard counter. */
export const ONGOING_TRIP_STATUSES: TripStatus[] = ['En Route', 'In Transit'];

/** A trip that still occupies its driver and vehicle. */
export const OPEN_TRIP_STATUSES: TripStatus[] = [
  'Assigned',
  'En Route',
  'In Transit',
];

export function isTripOpen(s: TripStatus): boolean {
  return OPEN_TRIP_STATUSES.indexOf(s) !== -1;
}
