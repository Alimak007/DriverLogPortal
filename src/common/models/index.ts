/**
 * The list schema expressed as TypeScript models.
 *
 * Every field maps to a SharePoint internal name documented in
 * SharePoint-Lists-Specification.txt. When a model here changes, that file
 * changes in the same commit.
 *
 * Dates are Date objects in the model and ISO-8601 UTC on the wire. The
 * service layer converts; nothing above it deals in strings.
 */

import {
  AllocationMode,
  BlockType,
  DriverStatus,
  ItemDescription,
  NotificationOutcome,
  OfferResponseStatus,
  Priority,
  RejectionReason,
  RequestStatus,
  RequestType,
  TripStatus,
  VehicleStatus,
} from './Enums';

export * from './Enums';

/** A SharePoint lookup, flattened to the two fields anything ever needs. */
export interface ILookup {
  id: number;
  title: string;
}

/** A SharePoint person field, flattened. */
export interface IPerson {
  id: number;
  title: string;
  email: string;
  loginName?: string;
}

export interface IListItemBase {
  id: number;
  created: Date;
  modified: Date;
  /** Captured on read so an optimistic update can send If-Match. */
  etag?: string;
}

// -----------------------------------------------------------------------------
// Master data
// -----------------------------------------------------------------------------

export interface IProperty extends IListItemBase {
  title: string;
}

export interface IVehicleType extends IListItemBase {
  title: string;
  /** Which request types this type may serve. Filters the vehicle tiles. */
  suitableFor: RequestType[];
  capacity: number;
  sortOrder: number;
  isActive: boolean;
}

export interface IDriver extends IListItemBase {
  driverName: string;
  /** Identity for login, for offers, and for the Accept check. */
  userAccount?: IPerson;
  driverId?: string;
  /** PERSONAL DATA. Released to a requester only after this driver accepts. */
  phone?: string;
  /** PERSONAL DATA, Fleet Admins only. */
  licenceNumber?: string;
  /** PERSONAL DATA, Fleet Admins only. Expiry inside the window excludes. */
  licenceExpiry?: Date;
  qualifiedVehicleTypes: ILookup[];
  homeProperty?: ILookup;
  status: DriverStatus;
  maxConcurrentTrips: number;
  isActive: boolean;
}

export interface IVehicle extends IListItemBase {
  plateNo: string;
  vehicleType?: ILookup;
  make?: string;
  model?: string;
  colour?: string;
  capacity?: number;
  defaultDriver?: ILookup;
  property?: ILookup;
  status: VehicleStatus;
}

export interface IAllocationRule extends IListItemBase {
  title: string;
  property?: ILookup;
  requestType: RequestType;
  mode: AllocationMode;
  offerWindowMinutes: number;
  maxConcurrentTripsPerDriver: number;
  advanceReleaseHours: number;
  isActive: boolean;
  /** Volume discipline: a two-minute internal transfer needs one e-mail. */
  notifyEnRoute: boolean;
  notifyInTransit: boolean;

  /**
   * The operating area. This is what replaced property scoping: a pickup
   * further than the radius from the centre is not broadcast at all, it goes
   * straight to the Coordinator.
   */
  operatingCentreLat?: number;
  operatingCentreLng?: number;
  operatingRadiusKm?: number;
}

export interface IAvailabilityBlock extends IListItemBase {
  title: string;
  /** Exactly one of driver or vehicle is set on any row. */
  driver?: ILookup;
  vehicle?: ILookup;
  fromDateTime: Date;
  toDateTime: Date;
  blockType: BlockType;
  property?: ILookup;
  notes?: string;
}

// -----------------------------------------------------------------------------
// Dispatch
// -----------------------------------------------------------------------------

export interface ITransportRequest extends IListItemBase {
  requestId: string;
  requestType: RequestType;
  requester?: IPerson;
  /**
   * Still written at booking from the requester's profile, FOR REPORTING ONLY.
   * It no longer affects who is offered the trip - allocation ignores property
   * entirely since the single-pool change.
   */
  property?: ILookup;

  // --- Where the trip runs -------------------------------------------------
  // Free coordinates from a live map search, not a lookup into a list. A
  // requester can pick any place in the UAE.
  pickupText: string;
  pickupPlaceId?: string;
  pickupLat?: number;
  pickupLng?: number;
  /**
   * "Gate 2, by the security hut". NOT an optional extra: a list entry called
   * "Khalifa Plant Gate 2" told a driver which gate; a pin labelled "KIZAD,
   * Abu Dhabi" does not. Shown on the booking form and on both driver cards.
   */
  pickupNotes?: string;

  dropOffText: string;
  dropOffPlaceId?: string;
  dropOffLat?: number;
  dropOffLng?: number;
  dropOffNotes?: string;

  requiredDateTime: Date;
  returnDateTime?: Date;
  noOfPax?: number;
  vehicleTypeRequested?: ILookup;
  /** The only free-text field in the booking form. */
  otherVehicleText?: string;
  itemDescription?: ItemDescription;
  itemNotes?: string;
  priority: Priority;
  status: RequestStatus;
  /** Written by the Broadcast flow. The Expire flow reads it every minute. */
  offerClosesAt?: Date;
  /** Advance bookings only: when the held request is released to broadcast. */
  releaseAt?: Date;
  acceptedTrip?: ILookup;
  remarks?: string;
  rejectionReason?: string;
  cancellationReason?: string;
}

export interface IRideOffer extends IListItemBase {
  title: string;
  request?: ILookup;
  driver?: ILookup;
  /** Usually blank. The vehicle is bound at CLAIM time, not broadcast time. */
  vehicle?: ILookup;
  offeredOn: Date;
  responseStatus: OfferResponseStatus;
  rejectionReason?: RejectionReason;
  rejectionNotes?: string;
  respondedOn?: Date;
}

export interface ITripAssignment extends IListItemBase {
  tripId: string;
  request?: ILookup;
  driver?: ILookup;
  vehicle?: ILookup;
  acceptedOn: Date;
  tripStatus: TripStatus;
  /** PHASE 1: entered by hand. Always displayed with etaUpdatedOn. */
  etaMinutes?: number;
  etaUpdatedOn?: Date;
  actualStartTime?: Date;
  actualPickupTime?: Date;
  actualDropTime?: Date;
  openingReading?: number;
  closingReading?: number;
  allocationMode: AllocationMode;
  allocatedBy?: IPerson;
  isOverride: boolean;
  overrideReason?: string;
  cancellationReason?: string;
}

export interface ITripStatusHistory extends IListItemBase {
  title: string;
  trip?: ILookup;
  fromStatus?: string;
  toStatus: string;
  changedBy?: IPerson;
  changedOn: Date;
  comment?: string;
}

export interface INotificationLogEntry extends IListItemBase {
  title: string;
  relatedRequest?: ILookup;
  relatedTrip?: ILookup;
  channel: string;
  recipient: string;
  templateKey: string;
  sentOn: Date;
  outcome: NotificationOutcome;
  errorMessage?: string;
  retryCount: number;
}

// -----------------------------------------------------------------------------
// Composites - what a view actually needs, assembled by the service layer
// -----------------------------------------------------------------------------

/**
 * A request with its trip and the driver and vehicle details the requester is
 * allowed to see. Driver phone appears ONLY once a driver has accepted.
 */
export interface IRequestWithTrip {
  request: ITransportRequest;
  trip?: ITripAssignment;
  driverName?: string;
  driverPhone?: string;
  vehiclePlate?: string;
  vehicleMake?: string;
  vehicleModel?: string;
  vehicleColour?: string;
  vehicleTypeName?: string;
  timeline: ITripStatusHistory[];
}

/** An open offer as the driver dashboard shows it. */
export interface IOfferCard {
  offer: IRideOffer;
  request: ITransportRequest;
  requesterName: string;
  /** Omitted until the driver accepts. */
  requesterPhone?: string;
  vehicleTypeName: string;
  secondsRemaining: number;
}

/** A trip as the driver dashboard shows it. */
export interface ITripCard {
  trip: ITripAssignment;
  request: ITransportRequest;
  requesterName: string;
  requesterPhone?: string;
  vehiclePlate: string;
}

/** One batched read behind the whole Coordinator dashboard. */
export interface IDispatchSnapshot {
  counters: {
    ongoingTrips: number;
    pendingRequests: number;
    completedTrips: number;
    availableDrivers: number;
    availableVehicles: number;
    allocatedPairs: number;
  };
  ongoingTrips: ITripCard[];
  pendingRequests: ITransportRequest[];
  drivers: IDriver[];
  vehicles: IVehicle[];
  /** Requests whose window closed with nobody accepting. Pinned to the top. */
  escalated: number[];
  loadedAt: Date;
}
