/**
 * The service contracts.
 *
 * Web parts depend on THESE, never on an implementation and never on PnPjs.
 * Swapping SharePoint for Dataverse means writing new classes behind these
 * interfaces and changing one factory - not touching a single component.
 */

import {
  IAllocationRule,
  IAvailabilityBlock,
  IDriver,
  ILocation,
  IOfferCard,
  IRequestWithTrip,
  ITransportRequest,
  ITripCard,
  ITripStatusHistory,
  IVehicle,
  IVehicleType,
  RequestType,
} from '../models';
import { ActorRole } from '../rules/statusTransitions';
import { IDispatchService } from './DispatchService';

/** Who is looking, and what they are allowed to be shown. */
export interface ICurrentUser {
  id: number;
  displayName: string;
  email: string;
  loginName: string;
  roles: ActorRole[];
  /** Set when this user has a DriversMaster record. */
  driverId?: number;
  /** Property scoping. Empty means every property (Management Viewer). */
  propertyIds: number[];
}

export interface IUserService {
  getCurrentUser(): Promise<ICurrentUser>;
}

/**
 * Reference data. Cached per page load - none of it changes during a session,
 * and re-reading it on every screen would blow the four-call budget.
 */
export interface IMasterDataService {
  getVehicleTypes(): Promise<IVehicleType[]>;
  /** Filtered by SuitableFor against the chosen request type. */
  getVehicleTypesFor(requestType: RequestType): Promise<IVehicleType[]>;
  getLocations(propertyId?: number): Promise<ILocation[]>;
  getDrivers(propertyIds?: number[]): Promise<IDriver[]>;
  getVehicles(propertyIds?: number[]): Promise<IVehicle[]>;
  getAllocationRule(
    propertyId: number,
    requestType: RequestType,
  ): Promise<IAllocationRule | undefined>;
  getBlocks(from: Date, to: Date): Promise<IAvailabilityBlock[]>;
  /** Driver id -> property ids, from DLClientMaster. Multi-property aware. */
  getDriverPropertyMap(): Promise<Map<number, number[]>>;
  clearCache(): void;
}



/**
 * The requester's side.
 *
 * NOTE WHAT IS ABSENT: there is no method here that reads vehicle or driver
 * availability. The booking web part must not query VehicleMaster
 * availability AT ALL, because anything the client reads is reachable through
 * the REST API regardless of what the interface displays.
 */
export interface IRequestService {
  getMyRequests(top?: number): Promise<ITransportRequest[]>;
  getRequestWithTrip(requestItemId: number): Promise<IRequestWithTrip>;
  getTimeline(tripItemId: number): Promise<ITripStatusHistory[]>;
}

export interface IOfferService {
  /** Open offers for the signed-in driver. Polled. */
  getOpenOffers(driverId: number): Promise<IOfferCard[]>;
}


export interface ITripService {
  getMyTripsToday(driverId: number): Promise<ITripCard[]>;
  getCompletedToday(driverId: number): Promise<ITripCard[]>;
}



export interface IReportFilter {
  from: Date;
  to: Date;
  propertyId?: number;
  requestType?: RequestType;
  vehicleTypeId?: number;
  driverId?: number;
}

export interface IReportRow {
  [column: string]: string | number | undefined;
}

export interface IReportResult {
  key: string;
  title: string;
  columns: { key: string; label: string; numeric?: boolean }[];
  rows: IReportRow[];
  /** True when the rows came from a nightly snapshot rather than live items. */
  fromSnapshot: boolean;
  generatedAt: Date;
}

export interface IReportService {
  run(reportKey: string, filter: IReportFilter): Promise<IReportResult>;
}


export interface IServices {
  user: IUserService;
  master: IMasterDataService;
  requests: IRequestService;
  offers: IOfferService;
  trips: ITripService;
  /** Every write in the solution. See IDispatchService in DispatchService.ts. */
  dispatch: IDispatchService;
  reports: IReportService;
}
