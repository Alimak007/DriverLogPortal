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
/**
 * Internal list titles.
 *
 * `Locations` and `DLClientMaster` are deliberately ABSENT. Both lists still
 * exist on the site and keep their data, but nothing in the solution reads
 * them any more: destinations come from a live map search, and property
 * scoping is gone. Do not add them back "just in case" - a fallback path to a
 * list nobody maintains is worse than no path at all.
 */
export const LIST = {
  // Reused unchanged. PropertyMaster is now a REPORTING label only - the
  // allocation engine does not read it.
  propertyMaster: 'PropertyMaster',

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

  /**
   * TomTom, for the map tiles and the place search on the booking form.
   *
   * THIS KEY SHIPS IN THE BUNDLE, which any authenticated user can read and
   * extract. That is accepted for a browser key, but it means two things are
   * not optional:
   *
   *   1. Restrict it by HTTP REFERRER to the SharePoint domain in the TomTom
   *      developer portal. Without that, anyone who reads the bundle can
   *      spend the quota from anywhere.
   *   2. Set a DAILY USAGE ALERT, so exhaustion reaches the key owner rather
   *      than reaching a requester as "Location search is unavailable".
   *
   * The web part property pane overrides this per site. Rotating the key is a
   * change to this one line.
   */
  tomtom: {
    apiKey: string;
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

  tomtom: {
    apiKey: 'ofXMyezRK70KYWilHswlPgWrZHwdLcsC',
  },

};
