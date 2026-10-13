/**
 * A driver's own trips. READS ONLY - the four taps live in
 * DispatchService.updateTripStatus.
 */

import { LIST } from '../config/AppConfig';
import { ITripCard, TRIP_STATUS } from '../models';
import { gulfDayBounds, toIso } from '../utils/dates';
import { ITripService } from './IServices';
import { EXPAND_TRIP, SELECT_TRIP, toRequest, toTrip } from './mappers';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

export class TripService implements ITripService {

  public async getMyTripsToday(driverId: number): Promise<ITripCard[]> {
    return this.readTrips(
      driverId,
      `(TripStatus eq '${TRIP_STATUS.Assigned}' or TripStatus eq '${TRIP_STATUS.EnRoute}' ` +
        `or TripStatus eq '${TRIP_STATUS.InTransit}')`,
      'Your trips could not be loaded.',
    );
  }

  public async getCompletedToday(driverId: number): Promise<ITripCard[]> {
    const { from, to } = gulfDayBounds();
    return this.readTrips(
      driverId,
      `TripStatus eq '${TRIP_STATUS.Completed}' and ActualDropTime ge '${toIso(from)}' ` +
        `and ActualDropTime lt '${toIso(to)}'`,
      "Today's completed trips could not be loaded.",
    );
  }

  private async readTrips(
    driverId: number,
    statusFilter: string,
    whatFailed: string,
  ): Promise<ITripCard[]> {
    return guard(
      whatFailed,
      'Refresh the page. If it keeps failing, contact the Transportation Coordinator.',
      async () => {
        const sp = getSP();

        const tripRows: any[] = await sp.web.lists
          .getByTitle(LIST.tripAssignments)
          .items.select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)
          .filter(`DriverId eq ${driverId} and ${statusFilter}`)
          .top(50)();

        if (tripRows.length === 0) return [];

        const trips = tripRows.map(toTrip);
        const requestIds = trips
          .map((t) => t.request?.id)
          .filter((id): id is number => typeof id === 'number');

        if (requestIds.length === 0) return [];

        const requestRows: any[] = await sp.web.lists
          .getByTitle(LIST.transportRequests)
          .items.select(
            'Id', 'Title', 'RequestType', 'RequiredDateTime',
            'ReturnDateTime', 'NoOfPax', 'ItemDescription', 'Priority',
            'ReqStatus', 'Created', 'Modified',
            'PickupText', 'PickupNotes', 'DropOffText', 'DropOffNotes',
            'Requester/Id', 'Requester/Title', 'Requester/EMail',
            'VehicleTypeRequested/Id', 'VehicleTypeRequested/Title',
          )
          .expand('Requester', 'VehicleTypeRequested')
          .filter(requestIds.map((id) => `Id eq ${id}`).join(' or '))
          .top(50)();

        const byId = new Map<number, any>();
        requestRows.forEach((r) => byId.set(r.Id, r));

        const cards: ITripCard[] = [];
        trips.forEach((trip) => {
          const raw = trip.request ? byId.get(trip.request.id) : undefined;
          if (!raw) return;
          const request = toRequest(raw);
          cards.push({
            trip,
            request,
            requesterName: request.requester?.title ?? '',
            // Released to the driver on the same basis the driver's number is
            // released to the requester: only after acceptance, only for this
            // trip.
            requesterPhone: undefined,
            vehiclePlate: trip.vehicle?.title ?? '',
          });
        });

        // Ordered by required time, as the dashboard shows them.
        cards.sort(
          (a, b) =>
            a.request.requiredDateTime.getTime() - b.request.requiredDateTime.getTime(),
        );
        return cards;
      },
    );
  }
}
