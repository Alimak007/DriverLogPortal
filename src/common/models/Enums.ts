/**
 * Every choice value in the solution, as a string-literal union.
 *
 * THESE STRINGS ARE THE CONTRACT between the web parts, the SharePoint choice
 * columns and the Power Automate flows. Changing one here without changing it
 * in the list and in every flow breaks allocation silently.
 *
 * See SharePoint-Lists-Specification.txt, which carries the same values.
 */

// -----------------------------------------------------------------------------
// Request
// -----------------------------------------------------------------------------

export type RequestType =
  | 'Client or Client Guest'
  | 'Staff'
  | 'Items or Documents';

export const REQUEST_TYPES: RequestType[] = [
  'Client or Client Guest',
  'Staff',
  'Items or Documents',
];

export type RequestStatus =
  | 'Submitted'
  | 'Held'
  | 'Broadcasting'
  | 'Awaiting Coordinator'
  | 'Allocated'
  | 'Completed'
  | 'Rejected'
  | 'Cancelled';

export type Priority = 'Normal' | 'Urgent';

export type ItemDescription =
  | 'Documents'
  | 'Parcel'
  | 'Equipment'
  | 'Medical'
  | 'Food'
  | 'Other';

export const ITEM_DESCRIPTIONS: ItemDescription[] = [
  'Documents',
  'Parcel',
  'Equipment',
  'Medical',
  'Food',
  'Other',
];

// -----------------------------------------------------------------------------
// Offer
// -----------------------------------------------------------------------------

export type OfferResponseStatus =
  | 'Pending'    // open, window running
  | 'Accepted'   // this driver won the race
  | 'Rejected'   // declined with a reason
  | 'Expired'    // window closed with no response
  | 'Closed'     // another driver accepted first
  | 'Withdrawn'; // the Coordinator assigned manually, skipping the broadcast

/**
 * A dropdown, never free text, because a driver who rejects everything is a
 * management issue the Coordinator should be able to see.
 */
export type RejectionReason =
  | 'Already on a trip'
  | 'Vehicle unavailable'
  | 'Outside my shift'
  | 'Personal reason'
  | 'Other';

export const REJECTION_REASONS: RejectionReason[] = [
  'Already on a trip',
  'Vehicle unavailable',
  'Outside my shift',
  'Personal reason',
  'Other',
];

// -----------------------------------------------------------------------------
// Trip
// -----------------------------------------------------------------------------

export type TripStatus =
  | 'Assigned'
  | 'En Route'
  | 'In Transit'
  | 'Completed'
  | 'Cancelled';

export type AllocationMode = 'Automatic' | 'Manual';

// -----------------------------------------------------------------------------
// Master data
// -----------------------------------------------------------------------------

export type DriverStatus = 'Available' | 'On Trip' | 'On Leave' | 'Inactive';

export type VehicleStatus =
  | 'Available'
  | 'On Trip'
  | 'Under Maintenance'
  | 'Retired';

export type BlockType = 'Leave' | 'Shift' | 'Maintenance' | 'Other';

export type LocationCategory =
  | 'Office'
  | 'Residence'
  | 'Airport'
  | 'Hospital'
  | 'Hotel'
  | 'Warehouse'
  | 'Other';

// -----------------------------------------------------------------------------
// Notification
// -----------------------------------------------------------------------------

export type NotificationOutcome = 'Sent' | 'Failed' | 'Retrying';

// -----------------------------------------------------------------------------
// Semantic state
// -----------------------------------------------------------------------------

/**
 * The four meanings the requirements fix for colour coding. Every status in
 * the solution maps to one of these, and the Badge component renders the
 * colour AND the label - status is never conveyed by colour alone.
 */
export type SemanticState =
  | 'healthy'    // available, assigned, completed
  | 'attention'  // awaiting Coordinator, pending, on leave
  | 'progress'   // en route, in transit, on trip
  | 'exception'  // cancelled, rejected, overdue
  | 'neutral';

const REQUEST_STATE: Record<RequestStatus, SemanticState> = {
  'Submitted': 'attention',
  'Held': 'neutral',
  'Broadcasting': 'progress',
  'Awaiting Coordinator': 'attention',
  'Allocated': 'healthy',
  'Completed': 'healthy',
  'Rejected': 'exception',
  'Cancelled': 'exception',
};

const TRIP_STATE: Record<TripStatus, SemanticState> = {
  'Assigned': 'healthy',
  'En Route': 'progress',
  'In Transit': 'progress',
  'Completed': 'healthy',
  'Cancelled': 'exception',
};

const DRIVER_STATE: Record<DriverStatus, SemanticState> = {
  'Available': 'healthy',
  'On Trip': 'progress',
  'On Leave': 'attention',
  'Inactive': 'neutral',
};

const VEHICLE_STATE: Record<VehicleStatus, SemanticState> = {
  'Available': 'healthy',
  'On Trip': 'progress',
  'Under Maintenance': 'attention',
  'Retired': 'neutral',
};

const OFFER_STATE: Record<OfferResponseStatus, SemanticState> = {
  'Pending': 'attention',
  'Accepted': 'healthy',
  'Rejected': 'exception',
  'Expired': 'exception',
  'Closed': 'neutral',
  'Withdrawn': 'neutral',
};

export function requestState(s: RequestStatus): SemanticState {
  return REQUEST_STATE[s] ?? 'neutral';
}
export function tripState(s: TripStatus): SemanticState {
  return TRIP_STATE[s] ?? 'neutral';
}
export function driverState(s: DriverStatus): SemanticState {
  return DRIVER_STATE[s] ?? 'neutral';
}
export function vehicleState(s: VehicleStatus): SemanticState {
  return VEHICLE_STATE[s] ?? 'neutral';
}
export function offerState(s: OfferResponseStatus): SemanticState {
  return OFFER_STATE[s] ?? 'neutral';
}

// -----------------------------------------------------------------------------
// The exact strings, as constants
// -----------------------------------------------------------------------------
// Use these rather than typing a literal. A typo in a status string does not
// fail the build - it silently writes a value no filter will ever match.

export const REQUEST_STATUS = {
  Submitted: 'Submitted',
  Held: 'Held',
  Broadcasting: 'Broadcasting',
  AwaitingCoordinator: 'Awaiting Coordinator',
  Allocated: 'Allocated',
  Completed: 'Completed',
  Rejected: 'Rejected',
  Cancelled: 'Cancelled',
} as const;

export const TRIP_STATUS = {
  Assigned: 'Assigned',
  EnRoute: 'En Route',
  InTransit: 'In Transit',
  Completed: 'Completed',
  Cancelled: 'Cancelled',
} as const;

/**
 * Closed, Expired and Withdrawn are NOT interchangeable. The Broadcast
 * effectiveness and Driver response behaviour reports count them separately,
 * so collapsing any two makes those reports lie.
 */
export const OFFER_STATUS = {
  Pending: 'Pending',
  Accepted: 'Accepted',     // this driver won
  Rejected: 'Rejected',     // driver declined with a reason
  Expired: 'Expired',       // window ran out - set by a flow, never by us
  Closed: 'Closed',         // another driver won
  Withdrawn: 'Withdrawn',   // coordinator assigned manually, bypassing the broadcast
} as const;
