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

/**
 * The opening odometer reading is captured at the "On my way" tap, the
 * closing reading is prompted for at "Complete". Both flow into the
 * Daily Trip Sheet.
 */
export function odometerPromptFor(to: TripStatus): 'opening' | 'closing' | undefined {
  if (to === 'En Route') return 'opening';
  if (to === 'Completed') return 'closing';
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
