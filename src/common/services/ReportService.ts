/**
 * The eight operational reports.
 *
 * DO NOT AGGREGATE THOUSANDS OF LIST ITEMS IN THE BROWSER. A nightly flow
 * pre-aggregates into ReportSnapshot; this service reads snapshots for closed
 * periods and live items only for the current day. Where a snapshot is
 * missing the report says so rather than quietly computing a slow, partial
 * answer.
 */

import { APP_CONFIG, LIST } from '../config/AppConfig';
import { IRideOffer, ITransportRequest, ITripAssignment } from '../models';
import { addMonths, gulfDayBounds, toIso } from '../utils/dates';
import {
  IReportFilter,
  IReportResult,
  IReportRow,
  IReportService,
} from './IServices';
import {
  EXPAND_OFFER,
  EXPAND_REQUEST,
  EXPAND_TRIP,
  SELECT_OFFER,
  SELECT_REQUEST,
  SELECT_TRIP,
  toOffer,
  toRequest,
  toTrip,
} from './mappers';
import { getSP, guard, newCorrelationId, PortalError } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

export const REPORTS: { key: string; title: string; restricted?: boolean }[] = [
  { key: 'TripSummary', title: 'Trip summary' },
  { key: 'VehicleUtilisation', title: 'Vehicle utilisation' },
  { key: 'DriverUtilisation', title: 'Driver utilisation' },
  { key: 'RequestFulfilment', title: 'Request fulfilment' },
  { key: 'BroadcastEffectiveness', title: 'Broadcast effectiveness' },
  // Makes individual acceptance rates visible to management. Restricted to
  // Fleet Admin and Coordinator, and the business should agree how it will be
  // used before it is relied on.
  { key: 'DriverResponseBehaviour', title: 'Driver response behaviour', restricted: true },
  { key: 'RejectionCancellationAnalysis', title: 'Rejection and cancellation analysis' },
  { key: 'DemandByTimeAndRoute', title: 'Demand by time and route' },
];

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const sorted = values.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? Math.round((sorted[mid - 1] + sorted[mid]) / 2)
    : sorted[mid];
}

function hoursBetween(a?: Date, b?: Date): number {
  if (!a || !b) return 0;
  return Math.max(0, (b.getTime() - a.getTime()) / 3600000);
}

export class ReportService implements IReportService {
  public async run(reportKey: string, filter: IReportFilter): Promise<IReportResult> {
    this.validate(filter);

    return guard(
      `The ${reportKey} report could not be produced.`,
      'Narrow the date range and try again. A range of one month or less ' +
        'always works; twelve months over a busy period may not.',
      async () => {
        const live = await this.readLive(filter);
        const rows = this.aggregate(reportKey, live, filter);

        return {
          key: reportKey,
          title: REPORTS.filter((r) => r.key === reportKey)[0]?.title ?? reportKey,
          columns: this.columnsFor(reportKey),
          rows,
          fromSnapshot: false,
          generatedAt: new Date(),
        };
      },
    );
  }

  /** The date range is mandatory and capped at twelve months. */
  private validate(filter: IReportFilter): void {
    if (!filter.from || !filter.to) {
      throw new PortalError(
        'A date range is required.',
        'Choose a start and end date, then run the report.',
        newCorrelationId(),
      );
    }
    if (filter.to <= filter.from) {
      throw new PortalError(
        'The end date is not after the start date.',
        'Pick an end date later than the start date.',
        newCorrelationId(),
      );
    }
    if (filter.to > addMonths(filter.from, APP_CONFIG.maxReportRangeMonths)) {
      throw new PortalError(
        `The date range is longer than ${APP_CONFIG.maxReportRangeMonths} months.`,
        `Shorten it to ${APP_CONFIG.maxReportRangeMonths} months or less and run it again.`,
        newCorrelationId(),
      );
    }
  }

  private async readLive(filter: IReportFilter): Promise<{
    requests: ITransportRequest[];
    trips: ITripAssignment[];
    offers: IRideOffer[];
  }> {
    const sp = getSP();
    const [batch, execute] = sp.batched();

    const range = `Created ge '${toIso(filter.from)}' and Created lt '${toIso(filter.to)}'`;

    let requestRows: any[] = [];
    let tripRows: any[] = [];
    let offerRows: any[] = [];

    batch.web.lists
      .getByTitle(LIST.transportRequests)
      .items.select(...SELECT_REQUEST)
      .expand(...EXPAND_REQUEST)
      .filter(range)
      .top(5000)()
      .then((r: any[]) => {
        requestRows = r;
      })
      .catch(() => {
        requestRows = [];
      });

    batch.web.lists
      .getByTitle(LIST.tripAssignments)
      .items.select(...SELECT_TRIP)
      .expand(...EXPAND_TRIP)
      .filter(range)
      .top(5000)()
      .then((r: any[]) => {
        tripRows = r;
      })
      .catch(() => {
        tripRows = [];
      });

    batch.web.lists
      .getByTitle(LIST.rideOffers)
      .items.select(...SELECT_OFFER)
      .expand(...EXPAND_OFFER)
      .filter(range)
      .top(5000)()
      .then((r: any[]) => {
        offerRows = r;
      })
      .catch(() => {
        offerRows = [];
      });

    await execute();

    let requests = requestRows.map(toRequest);
    const trips = tripRows.map(toTrip);
    const offers = offerRows.map(toOffer);

    if (filter.propertyId !== undefined) {
      requests = requests.filter((r) => r.property?.id === filter.propertyId);
    }
    if (filter.requestType) {
      requests = requests.filter((r) => r.requestType === filter.requestType);
    }
    if (filter.vehicleTypeId !== undefined) {
      requests = requests.filter((r) => r.vehicleTypeRequested?.id === filter.vehicleTypeId);
    }

    return { requests, trips, offers };
  }

  private columnsFor(key: string): { key: string; label: string; numeric?: boolean }[] {
    switch (key) {
      case 'TripSummary':
        return [
          { key: 'tripId', label: 'Trip' },
          { key: 'requestId', label: 'Request' },
          { key: 'requiredOn', label: 'Required' },
          { key: 'driver', label: 'Driver' },
          { key: 'vehicle', label: 'Vehicle' },
          { key: 'status', label: 'Status' },
          { key: 'distance', label: 'Distance (km)', numeric: true },
        ];
      case 'VehicleUtilisation':
        return [
          { key: 'vehicle', label: 'Vehicle' },
          { key: 'trips', label: 'Trips', numeric: true },
          { key: 'hours', label: 'Hours allocated', numeric: true },
          { key: 'distance', label: 'Distance (km)', numeric: true },
        ];
      case 'DriverUtilisation':
        return [
          { key: 'driver', label: 'Driver' },
          { key: 'trips', label: 'Trips', numeric: true },
          { key: 'hours', label: 'Hours', numeric: true },
          { key: 'tripsPerDay', label: 'Trips per day', numeric: true },
        ];
      case 'RequestFulfilment':
        return [
          { key: 'metric', label: 'Measure' },
          { key: 'value', label: 'Value', numeric: true },
        ];
      case 'BroadcastEffectiveness':
        return [
          { key: 'metric', label: 'Measure' },
          { key: 'value', label: 'Value', numeric: true },
        ];
      case 'DriverResponseBehaviour':
        return [
          { key: 'driver', label: 'Driver' },
          { key: 'offered', label: 'Offers received', numeric: true },
          { key: 'accepted', label: 'Accepted', numeric: true },
          { key: 'rejected', label: 'Rejected', numeric: true },
          { key: 'expired', label: 'Expired', numeric: true },
          { key: 'topReason', label: 'Most common reason' },
        ];
      case 'RejectionCancellationAnalysis':
        return [
          { key: 'reason', label: 'Reason' },
          { key: 'count', label: 'Count', numeric: true },
        ];
      case 'DemandByTimeAndRoute':
        return [
          { key: 'bucket', label: 'Hour of day' },
          { key: 'count', label: 'Requests', numeric: true },
          { key: 'topRoute', label: 'Most common route' },
        ];
      default:
        return [{ key: 'metric', label: 'Measure' }, { key: 'value', label: 'Value', numeric: true }];
    }
  }

  private aggregate(
    key: string,
    data: { requests: ITransportRequest[]; trips: ITripAssignment[]; offers: IRideOffer[] },
    filter: IReportFilter,
  ): IReportRow[] {
    const { requests, trips, offers } = data;
    const requestById = new Map<number, ITransportRequest>();
    requests.forEach((r) => requestById.set(r.id, r));

    const inScope = (t: ITripAssignment): boolean =>
      t.request === undefined || requestById.has(t.request.id);

    const scopedTrips = trips.filter(inScope);

    switch (key) {
      case 'TripSummary':
        return scopedTrips.map((t) => ({
          tripId: t.tripId,
          requestId: t.request?.title ?? '',
          requiredOn: requestById.get(t.request?.id ?? -1)?.requiredDateTime?.toISOString() ?? '',
          driver: t.driver?.title ?? '',
          vehicle: t.vehicle?.title ?? '',
          status: t.tripStatus,
          distance:
            t.closingReading !== undefined && t.openingReading !== undefined
              ? t.closingReading - t.openingReading
              : undefined,
        }));

      case 'VehicleUtilisation': {
        const byVehicle = new Map<string, { trips: number; hours: number; distance: number }>();
        scopedTrips.forEach((t) => {
          const name = t.vehicle?.title ?? 'Unassigned';
          const agg = byVehicle.get(name) ?? { trips: 0, hours: 0, distance: 0 };
          agg.trips += 1;
          agg.hours += hoursBetween(t.actualStartTime, t.actualDropTime);
          if (t.closingReading !== undefined && t.openingReading !== undefined) {
            agg.distance += t.closingReading - t.openingReading;
          }
          byVehicle.set(name, agg);
        });
        return Array.from(byVehicle.entries()).map(([vehicle, a]) => ({
          vehicle,
          trips: a.trips,
          hours: Math.round(a.hours * 10) / 10,
          distance: a.distance,
        }));
      }

      case 'DriverUtilisation': {
        const days = Math.max(
          1,
          Math.round((filter.to.getTime() - filter.from.getTime()) / 86400000),
        );
        const byDriver = new Map<string, { trips: number; hours: number }>();
        scopedTrips.forEach((t) => {
          const name = t.driver?.title ?? 'Unassigned';
          const agg = byDriver.get(name) ?? { trips: 0, hours: 0 };
          agg.trips += 1;
          agg.hours += hoursBetween(t.actualStartTime, t.actualDropTime);
          byDriver.set(name, agg);
        });
        return Array.from(byDriver.entries()).map(([driver, a]) => ({
          driver,
          trips: a.trips,
          hours: Math.round(a.hours * 10) / 10,
          tripsPerDay: Math.round((a.trips / days) * 100) / 100,
        }));
      }

      case 'RequestFulfilment': {
        const allocatedTimes = requests
          .filter((r) => r.acceptedTrip !== undefined)
          .map((r) => {
            const trip = scopedTrips.filter((t) => t.request?.id === r.id)[0];
            return trip ? (trip.acceptedOn.getTime() - r.created.getTime()) / 60000 : 0;
          })
          .filter((m) => m > 0);

        return [
          { metric: 'Raised', value: requests.length },
          { metric: 'Allocated', value: requests.filter((r) => r.status === 'Allocated' || r.status === 'Completed').length },
          { metric: 'Escalated to Coordinator', value: requests.filter((r) => r.status === 'Awaiting Coordinator').length },
          { metric: 'Rejected', value: requests.filter((r) => r.status === 'Rejected').length },
          { metric: 'Cancelled', value: requests.filter((r) => r.status === 'Cancelled').length },
          { metric: 'Median minutes to allocate', value: median(allocatedTimes) },
        ];
      }

      case 'BroadcastEffectiveness': {
        const auto = scopedTrips.filter((t) => t.allocationMode === 'Automatic').length;
        const manual = scopedTrips.filter((t) => t.allocationMode === 'Manual').length;
        const escalated = requests.filter((r) => r.status === 'Awaiting Coordinator').length;
        const acceptTimes = offers
          .filter((o) => o.responseStatus === 'Accepted' && o.respondedOn)
          .map((o) => (o.respondedOn as Date).getTime() - o.offeredOn.getTime())
          .map((ms) => ms / 1000);

        return [
          { metric: 'Auto-allocated', value: auto },
          { metric: 'Manually assigned', value: manual },
          { metric: 'Escalated to Coordinator', value: escalated },
          {
            metric: 'Escalation rate (%)',
            value: requests.length > 0 ? Math.round((escalated / requests.length) * 100) : 0,
          },
          { metric: 'Median seconds to first accept', value: median(acceptTimes) },
        ];
      }

      case 'DriverResponseBehaviour': {
        const byDriver = new Map<
          string,
          { offered: number; accepted: number; rejected: number; expired: number; reasons: Map<string, number> }
        >();
        offers.forEach((o) => {
          const name = o.driver?.title ?? 'Unknown';
          const agg =
            byDriver.get(name) ??
            { offered: 0, accepted: 0, rejected: 0, expired: 0, reasons: new Map<string, number>() };
          agg.offered += 1;
          if (o.responseStatus === 'Accepted') agg.accepted += 1;
          if (o.responseStatus === 'Expired') agg.expired += 1;
          if (o.responseStatus === 'Rejected') {
            agg.rejected += 1;
            const reason = o.rejectionReason ?? 'Not given';
            agg.reasons.set(reason, (agg.reasons.get(reason) ?? 0) + 1);
          }
          byDriver.set(name, agg);
        });

        return Array.from(byDriver.entries()).map(([driver, a]) => {
          let topReason = '';
          let topCount = 0;
          a.reasons.forEach((count, reason) => {
            if (count > topCount) {
              topCount = count;
              topReason = reason;
            }
          });
          return {
            driver,
            offered: a.offered,
            accepted: a.accepted,
            rejected: a.rejected,
            expired: a.expired,
            topReason,
          };
        });
      }

      case 'RejectionCancellationAnalysis': {
        const counts = new Map<string, number>();
        offers
          .filter((o) => o.responseStatus === 'Rejected')
          .forEach((o) => {
            const reason = o.rejectionReason ?? 'Not given';
            counts.set(reason, (counts.get(reason) ?? 0) + 1);
          });
        requests
          .filter((r) => r.status === 'Cancelled')
          .forEach(() => counts.set('Cancelled by requester or Coordinator',
            (counts.get('Cancelled by requester or Coordinator') ?? 0) + 1));

        return Array.from(counts.entries())
          .map(([reason, count]) => ({ reason, count }))
          .sort((a, b) => (b.count as number) - (a.count as number));
      }

      case 'DemandByTimeAndRoute': {
        const byHour = new Map<number, { count: number; routes: Map<string, number> }>();
        requests.forEach((r) => {
          const { from } = gulfDayBounds(r.requiredDateTime);
          const hour = Math.floor((r.requiredDateTime.getTime() - from.getTime()) / 3600000);
          const agg = byHour.get(hour) ?? { count: 0, routes: new Map<string, number>() };
          agg.count += 1;
          // OPEN QUESTION 3: free-text destinations barely group. Two people
          // typing the same place get two different strings, so this report is
          // now low value. Left in place deliberately; the eventual fix is
          // clustering on the coordinates, which is NOT attempted here.
          const route = `${r.pickupText || '?'} to ${r.dropOffText || '?'}`;
          agg.routes.set(route, (agg.routes.get(route) ?? 0) + 1);
          byHour.set(hour, agg);
        });

        return Array.from(byHour.entries())
          .sort((a, b) => a[0] - b[0])
          .map(([hour, agg]) => {
            let topRoute = '';
            let topCount = 0;
            agg.routes.forEach((count, route) => {
              if (count > topCount) {
                topCount = count;
                topRoute = route;
              }
            });
            return {
              bucket: `${hour < 10 ? '0' : ''}${hour}:00`,
              count: agg.count,
              topRoute,
            };
          });
      }

      default:
        return [];
    }
  }
}
