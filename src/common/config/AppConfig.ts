/**
 * Solution-wide configuration.
 *
 * Anything here that a business user might want to change belongs in a
 * SharePoint list instead (AllocationRules, VehicleTypes, Locations,
 * Templates). What stays here is deployment shape: list names, font source,
 * identifier format, poll intervals.
 *
 * NO SECRETS. The SPFx package is fully readable by any authenticated user.
 * No keys, connection strings or service credentials may be placed in it.
 */

export type FontSource = 'siteAssets' | 'google';
export type RequestIdFormat = 'short' | 'dated';

/** Internal list titles. These are the SharePoint titles, not display names. */
export const LIST = {
  // Reused unchanged
  propertyMaster: 'PropertyMaster',
  dlClientMaster: 'DLClientMaster',

  // Refactored
  vehicleMaster: 'VehicleMaster',
  driversMaster: 'DriversMaster',

  // New
  transportRequests: 'TransportRequests',
  rideOffers: 'RideOffers',
  tripAssignments: 'TripAssignments',
  tripStatusHistory: 'TripStatusHistory',
  availabilityBlocks: 'AvailabilityBlocks',
  vehicleTypes: 'VehicleTypes',
  locations: 'Locations',
  allocationRules: 'AllocationRules',
  notificationLog: 'NotificationLog',
  templates: 'Templates',
  reportSnapshot: 'ReportSnapshot',
} as const;

export interface IAppConfig {
  /**
   * Where the two faces are served from. Site assets are the safer default
   * where outbound font requests are restricted.
   */
  fonts: {
    source: FontSource;
    siteAssetsPath: string;
  };

  /**
   * RequestId format. OPEN QUESTION in the requirements - the mockups show
   * both FB-001 and FB-20260925-001.
   *   'short' -> FB-0001
   *   'dated' -> FB-20260925-0001
   *
   * Must be settled before the first production request is created. The
   * format cannot change afterwards without reissuing every identifier.
   */
  requestIdFormat: RequestIdFormat;
  requestIdPrefix: string;
  tripIdPrefix: string;

  /**
   * Driver dashboard poll. The non-functional requirement is an open offer
   * visible within 20 seconds of broadcast, at no more than one request per
   * 20 seconds per client.
   */
  driverPollIntervalMs: number;

  /** Dispatch dashboard refresh. Slower; the Coordinator is at a desk. */
  dispatchPollIntervalMs: number;

  /** Page size for every list read. The threshold strategy depends on it. */
  pageSize: number;

  /** Stored UTC, displayed in Gulf Standard Time. Never the browser's zone. */
  displayTimeZone: string;
  displayUtcOffsetMinutes: number;

  /** Reports date range is mandatory and capped. */
  maxReportRangeMonths: number;

  /**
   * Phase 2 switch. The Track My Ride map area is a swappable region and this
   * flag is what activates it. False throughout phase 1.
   */
  liveTrackingEnabled: boolean;

  /**
   * Where the requester's two pages live, RELATIVE TO THE SITE the web part is
   * running on. The web parts link to each other with these, so moving or
   * renaming a page is a change here and nowhere else.
   */
  pages: {
    bookARide: string;
    myBookings: string;
  };

}

export const APP_CONFIG: IAppConfig = {
  fonts: {
    source: 'siteAssets',
    siteAssetsPath: '/SiteAssets/fonts',
  },

  requestIdFormat: 'short',
  requestIdPrefix: 'FB',
  tripIdPrefix: 'TR',

  driverPollIntervalMs: 20000,
  dispatchPollIntervalMs: 60000,

  pageSize: 50,

  displayTimeZone: 'Gulf Standard Time',
  displayUtcOffsetMinutes: 240, // UTC+4, no daylight saving

  maxReportRangeMonths: 12,

  liveTrackingEnabled: false,

  pages: {
    bookARide: '/SitePages/Home.aspx',
    myBookings: '/SitePages/MyBookings.aspx',
  },

};
