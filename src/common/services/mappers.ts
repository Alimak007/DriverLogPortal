/**
 * SharePoint list item -> typed model.
 *
 * The SELECT and EXPAND strings live here too, next to the mapper that reads
 * them, so a field added to one is never forgotten in the other.
 *
 * Every mapper is defensive: SharePoint returns null for an empty lookup, an
 * empty array for an empty multi-lookup, and omits a field entirely if the
 * caller forgot to select it.
 */

import {
  AllocationMode,
  BlockType,
  DriverStatus,
  IAllocationRule,
  IAvailabilityBlock,
  IDriver,
  IDriverPropertyLink,
  ILocation,
  ILookup,
  INotificationLogEntry,
  IPerson,
  IProperty,
  IRideOffer,
  ITransportRequest,
  ITripAssignment,
  ITripStatusHistory,
  IVehicle,
  IVehicleType,
  ItemDescription,
  LocationCategory,
  NotificationOutcome,
  OfferResponseStatus,
  Priority,
  RejectionReason,
  RequestStatus,
  RequestType,
  TripStatus,
  VehicleStatus,
} from '../models';
import { parseDate } from '../utils/dates';

/* eslint-disable @typescript-eslint/no-explicit-any */

function lookup(raw: any): ILookup | undefined {
  if (!raw || raw.Id === undefined || raw.Id === null) return undefined;
  return { id: raw.Id, title: raw.Title ?? '' };
}

function lookupMany(raw: any): ILookup[] {
  if (!raw) return [];
  const rows = Array.isArray(raw) ? raw : raw.results;
  if (!Array.isArray(rows)) return [];
  return rows.map(lookup).filter((x): x is ILookup => x !== undefined);
}

function person(raw: any): IPerson | undefined {
  if (!raw || raw.Id === undefined || raw.Id === null) return undefined;
  return {
    id: raw.Id,
    title: raw.Title ?? '',
    email: raw.EMail ?? raw.Email ?? '',
    loginName: raw.LoginName ?? raw.Name,
  };
}

function choices(raw: any): string[] {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw;
  if (Array.isArray(raw.results)) return raw.results;
  return [raw];
}

function base(raw: any): { id: number; created: Date; modified: Date; etag?: string } {
  return {
    id: raw.Id,
    created: parseDate(raw.Created) ?? new Date(0),
    modified: parseDate(raw.Modified) ?? new Date(0),
    etag: raw['odata.etag'] ?? raw.__metadata?.etag,
  };
}

// -----------------------------------------------------------------------------
// Master data
// -----------------------------------------------------------------------------

export const SELECT_PROPERTY = ['Id', 'Title', 'Created', 'Modified'];

export function toProperty(raw: any): IProperty {
  return { ...base(raw), title: raw.Title ?? '' };
}

export const SELECT_VEHICLE_TYPE = [
  'Id', 'Title', 'SuitableFor', 'Capacity', 'SortOrder', 'IsActive',
  'Created', 'Modified',
];

export function toVehicleType(raw: any): IVehicleType {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    suitableFor: choices(raw.SuitableFor) as RequestType[],
    capacity: raw.Capacity ?? 0,
    sortOrder: raw.SortOrder ?? 0,
    isActive: raw.IsActive !== false,
  };
}

export const SELECT_LOCATION = [
  'Id', 'Title', 'Address', 'Latitude', 'Longitude', 'Category', 'IsActive',
  'Created', 'Modified', 'Property/Id', 'Property/Title',
];
export const EXPAND_LOCATION = ['Property'];

export function toLocation(raw: any): ILocation {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    property: lookup(raw.Property),
    address: raw.Address ?? undefined,
    latitude: raw.Latitude ?? undefined,
    longitude: raw.Longitude ?? undefined,
    category: (raw.Category ?? undefined) as LocationCategory | undefined,
    isActive: raw.IsActive !== false,
  };
}

export const SELECT_DRIVER = [
  'Id', 'Title', 'DriverID', 'Phone', 'LicenceNumber',
  'LicenceExpiry', 'Status', 'MaxConcurrentTrips', 'IsActive',
  'Created', 'Modified',
  'UserAccount/Id', 'UserAccount/Title', 'UserAccount/EMail',
  'QualifiedVehicleTypes/Id', 'QualifiedVehicleTypes/Title',
  'HomeProperty/Id', 'HomeProperty/Title',
];
export const EXPAND_DRIVER = ['UserAccount', 'QualifiedVehicleTypes', 'HomeProperty'];

export function toDriver(raw: any): IDriver {
  return {
    ...base(raw),
    driverName: raw.Title ?? '',
    userAccount: person(raw.UserAccount),
    driverId: raw.DriverID ?? undefined,
    phone: raw.Phone ?? undefined,
    licenceNumber: raw.LicenceNumber ?? undefined,
    licenceExpiry: parseDate(raw.LicenceExpiry),
    qualifiedVehicleTypes: lookupMany(raw.QualifiedVehicleTypes),
    homeProperty: lookup(raw.HomeProperty),
    status: (raw.Status ?? 'Available') as DriverStatus,
    maxConcurrentTrips: raw.MaxConcurrentTrips ?? 1,
    isActive: raw.IsActive !== false,
  };
}

export const SELECT_VEHICLE = [
  'Id', 'Title', 'PlateNo', 'Make', 'Model', 'Colour', 'Capacity', 'VehStatus',
  'Created', 'Modified',
  'VehicleType/Id', 'VehicleType/Title',
  'DefaultDriver/Id', 'DefaultDriver/Title',
  // PropertyRef, NOT Property. VehicleMaster.Property still exists as a plain
  // TEXT column used by the legacy logbook - never read or write it.
  'PropertyRef/Id', 'PropertyRef/Title',
];
export const EXPAND_VEHICLE = ['VehicleType', 'DefaultDriver', 'PropertyRef'];

export function toVehicle(raw: any): IVehicle {
  return {
    ...base(raw),
    plateNo: raw.PlateNo ?? raw.Title ?? '',
    vehicleType: lookup(raw.VehicleType),
    make: raw.Make ?? undefined,
    model: raw.Model ?? undefined,
    colour: raw.Colour ?? undefined,
    capacity: raw.Capacity ?? undefined,
    defaultDriver: lookup(raw.DefaultDriver),
    property: lookup(raw.PropertyRef),
    status: (raw.VehStatus ?? 'Available') as VehicleStatus,
  };
}

export const SELECT_ALLOCATION_RULE = [
  'Id', 'Title', 'RequestType', 'Mode', 'OfferWindowMinutes',
  'MaxConcurrentTripsPerDriver', 'AdvanceReleaseHours', 'IsActive',
  'NotifyEnRoute', 'NotifyInTransit', 'Created', 'Modified',
  'Property/Id', 'Property/Title',
];
export const EXPAND_ALLOCATION_RULE = ['Property'];

export function toAllocationRule(raw: any): IAllocationRule {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    property: lookup(raw.Property),
    requestType: (raw.RequestType ?? 'Staff') as RequestType,
    mode: (raw.Mode ?? 'Automatic') as AllocationMode,
    offerWindowMinutes: raw.OfferWindowMinutes ?? 5,
    maxConcurrentTripsPerDriver: raw.MaxConcurrentTripsPerDriver ?? 1,
    advanceReleaseHours: raw.AdvanceReleaseHours ?? 4,
    isActive: raw.IsActive !== false,
    notifyEnRoute: raw.NotifyEnRoute !== false,
    notifyInTransit: raw.NotifyInTransit !== false,
  };
}

export const SELECT_BLOCK = [
  'Id', 'Title', 'FromDateTime', 'ToDateTime', 'BlockType', 'Notes',
  'Created', 'Modified',
  'Driver/Id', 'Driver/Title', 'Vehicle/Id', 'Vehicle/Title',
  'Property/Id', 'Property/Title',
];
export const EXPAND_BLOCK = ['Driver', 'Vehicle', 'Property'];

export function toAvailabilityBlock(raw: any): IAvailabilityBlock {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    driver: lookup(raw.Driver),
    vehicle: lookup(raw.Vehicle),
    fromDateTime: parseDate(raw.FromDateTime) ?? new Date(0),
    toDateTime: parseDate(raw.ToDateTime) ?? new Date(0),
    blockType: (raw.BlockType ?? 'Other') as BlockType,
    property: lookup(raw.Property),
    notes: raw.Notes ?? undefined,
  };
}

export const SELECT_DL_CLIENT = [
  'Id', 'Created', 'Modified',
  'DriverID/Id', 'DriverID/Title',
  'PropertyRef/Id', 'PropertyRef/Title',
];
export const EXPAND_DL_CLIENT = ['DriverID', 'PropertyRef'];

/** DLClientMaster uses DriverID and PropertyRef, not Driver and Property. */
export function toDriverPropertyLink(raw: any): IDriverPropertyLink {
  return {
    ...base(raw),
    driver: lookup(raw.DriverID),
    property: lookup(raw.PropertyRef),
  };
}

// -----------------------------------------------------------------------------
// Dispatch
// -----------------------------------------------------------------------------

export const SELECT_REQUEST = [
  // The reference lives in Title. There is no RequestId column.
  'Id', 'Title', 'RequestType', 'RequiredDateTime',
  'ReturnDateTime', 'NoOfPax', 'OtherVehicleText', 'ItemDescription',
  'ItemNotes', 'Priority', 'ReqStatus', 'OfferClosesAt', 'ReleaseAt', 'Remarks',
  'RejectionReason', 'CancellationReason', 'Created', 'Modified',
  'Requester/Id', 'Requester/Title', 'Requester/EMail',
  'Property/Id', 'Property/Title',
  'PickupLocation/Id', 'PickupLocation/Title',
  'DropOffLocation/Id', 'DropOffLocation/Title',
  'VehicleTypeRequested/Id', 'VehicleTypeRequested/Title',
  'AcceptedTrip/Id', 'AcceptedTrip/Title',
];
export const EXPAND_REQUEST = [
  'Requester', 'Property', 'PickupLocation', 'DropOffLocation',
  'VehicleTypeRequested', 'AcceptedTrip',
];

export function toRequest(raw: any): ITransportRequest {
  return {
    ...base(raw),
    requestId: raw.Title ?? '',
    requestType: (raw.RequestType ?? 'Staff') as RequestType,
    requester: person(raw.Requester),
    property: lookup(raw.Property),
    pickupLocation: lookup(raw.PickupLocation),
    dropOffLocation: lookup(raw.DropOffLocation),
    requiredDateTime: parseDate(raw.RequiredDateTime) ?? new Date(0),
    returnDateTime: parseDate(raw.ReturnDateTime),
    noOfPax: raw.NoOfPax ?? undefined,
    vehicleTypeRequested: lookup(raw.VehicleTypeRequested),
    otherVehicleText: raw.OtherVehicleText ?? undefined,
    itemDescription: (raw.ItemDescription ?? undefined) as ItemDescription | undefined,
    itemNotes: raw.ItemNotes ?? undefined,
    priority: (raw.Priority ?? 'Normal') as Priority,
    status: (raw.ReqStatus ?? 'Submitted') as RequestStatus,
    offerClosesAt: parseDate(raw.OfferClosesAt),
    releaseAt: parseDate(raw.ReleaseAt),
    acceptedTrip: lookup(raw.AcceptedTrip),
    remarks: raw.Remarks ?? undefined,
    rejectionReason: raw.RejectionReason ?? undefined,
    cancellationReason: raw.CancellationReason ?? undefined,
  };
}

export const SELECT_OFFER = [
  'Id', 'Title', 'OfferedOn', 'ResponseStatus', 'RejectionReason',
  'RejectionNotes', 'RespondedOn', 'Created', 'Modified',
  'Request/Id', 'Request/Title', 'Driver/Id', 'Driver/Title',
  'Vehicle/Id', 'Vehicle/Title',
];
export const EXPAND_OFFER = ['Request', 'Driver', 'Vehicle'];

export function toOffer(raw: any): IRideOffer {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    request: lookup(raw.Request),
    driver: lookup(raw.Driver),
    vehicle: lookup(raw.Vehicle),
    offeredOn: parseDate(raw.OfferedOn) ?? new Date(0),
    responseStatus: (raw.ResponseStatus ?? 'Pending') as OfferResponseStatus,
    rejectionReason: (raw.RejectionReason ?? undefined) as RejectionReason | undefined,
    rejectionNotes: raw.RejectionNotes ?? undefined,
    respondedOn: parseDate(raw.RespondedOn),
  };
}

export const SELECT_TRIP = [
  // The reference lives in Title. There is no TripId column.
  'Id', 'Title', 'AcceptedOn', 'TripStatus', 'EtaMinutes',
  'EtaUpdatedOn', 'ActualStartTime', 'ActualPickupTime', 'ActualDropTime',
  'OpeningReading', 'ClosingReading', 'AllocationMode', 'IsOverride',
  'OverrideReason', 'CancellationReason', 'Created', 'Modified',
  'Request/Id', 'Request/Title', 'Driver/Id', 'Driver/Title',
  'Vehicle/Id', 'Vehicle/Title',
  'AllocatedBy/Id', 'AllocatedBy/Title', 'AllocatedBy/EMail',
];
export const EXPAND_TRIP = ['Request', 'Driver', 'Vehicle', 'AllocatedBy'];

export function toTrip(raw: any): ITripAssignment {
  return {
    ...base(raw),
    tripId: raw.Title ?? '',
    request: lookup(raw.Request),
    driver: lookup(raw.Driver),
    vehicle: lookup(raw.Vehicle),
    acceptedOn: parseDate(raw.AcceptedOn) ?? new Date(0),
    tripStatus: (raw.TripStatus ?? 'Assigned') as TripStatus,
    etaMinutes: raw.EtaMinutes ?? undefined,
    etaUpdatedOn: parseDate(raw.EtaUpdatedOn),
    actualStartTime: parseDate(raw.ActualStartTime),
    actualPickupTime: parseDate(raw.ActualPickupTime),
    actualDropTime: parseDate(raw.ActualDropTime),
    openingReading: raw.OpeningReading ?? undefined,
    closingReading: raw.ClosingReading ?? undefined,
    allocationMode: (raw.AllocationMode ?? 'Automatic') as AllocationMode,
    allocatedBy: person(raw.AllocatedBy),
    isOverride: raw.IsOverride === true,
    overrideReason: raw.OverrideReason ?? undefined,
    cancellationReason: raw.CancellationReason ?? undefined,
  };
}

export const SELECT_HISTORY = [
  'Id', 'Title', 'FromStatus', 'ToStatus', 'ChangedOn', 'Comment',
  'Created', 'Modified',
  'Trip/Id', 'Trip/Title', 'ChangedBy/Id', 'ChangedBy/Title', 'ChangedBy/EMail',
];
export const EXPAND_HISTORY = ['Trip', 'ChangedBy'];

export function toHistory(raw: any): ITripStatusHistory {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    trip: lookup(raw.Trip),
    fromStatus: raw.FromStatus ?? undefined,
    toStatus: raw.ToStatus ?? '',
    changedBy: person(raw.ChangedBy),
    changedOn: parseDate(raw.ChangedOn) ?? new Date(0),
    comment: raw.Comment ?? undefined,
  };
}

export const SELECT_NOTIFICATION = [
  'Id', 'Title', 'Channel', 'Recipient', 'TemplateKey', 'SentOn', 'Outcome',
  'ErrorMessage', 'RetryCount', 'Created', 'Modified',
  'RelatedRequest/Id', 'RelatedRequest/Title',
  'RelatedTrip/Id', 'RelatedTrip/Title',
];
export const EXPAND_NOTIFICATION = ['RelatedRequest', 'RelatedTrip'];

export function toNotification(raw: any): INotificationLogEntry {
  return {
    ...base(raw),
    title: raw.Title ?? '',
    relatedRequest: lookup(raw.RelatedRequest),
    relatedTrip: lookup(raw.RelatedTrip),
    channel: raw.Channel ?? 'Email',
    recipient: raw.Recipient ?? '',
    templateKey: raw.TemplateKey ?? '',
    sentOn: parseDate(raw.SentOn) ?? new Date(0),
    outcome: (raw.Outcome ?? 'Sent') as NotificationOutcome,
    errorMessage: raw.ErrorMessage ?? undefined,
    retryCount: raw.RetryCount ?? 0,
  };
}
