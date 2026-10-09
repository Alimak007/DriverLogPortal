/**
 * The requester's READS. Every write goes through DispatchService.
 *
 * NOTE WHAT IS ABSENT. There is no method in this file that reads vehicle or
 * driver availability, and there is none in the booking web part either. The
 * client's first stated requirement is that vehicle availability is never
 * exposed to the requester, and anything this code reads is reachable by that
 * user through the REST API regardless of what the interface displays. The
 * requester chooses a vehicle TYPE, never a vehicle.
 */

import { APP_CONFIG, LIST } from '../config/AppConfig';
import {
  IRequestWithTrip,
  ITransportRequest,
  ITripStatusHistory,
} from '../models';
import { ICurrentUser, IRequestService } from './IServices';
import {
  EXPAND_HISTORY,
  EXPAND_REQUEST,
  EXPAND_TRIP,
  SELECT_HISTORY,
  SELECT_REQUEST,
  SELECT_TRIP,
  toHistory,
  toRequest,
  toTrip,
} from './mappers';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

export class RequestService implements IRequestService {
  private readonly user: ICurrentUser;

  constructor(user: ICurrentUser) {
    this.user = user;
  }

  public async getMyRequests(top: number = APP_CONFIG.pageSize): Promise<ITransportRequest[]> {
    return guard(
      'Your bookings could not be loaded.',
      'Refresh the page. If it keeps failing, contact the Transportation Coordinator.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.transportRequests)
          .items.select(...SELECT_REQUEST)
          .expand(...EXPAND_REQUEST)
          .filter(`Requester/Id eq ${this.user.id}`)
          .orderBy('Created', false)
          .top(top)();

        return rows.map(toRequest);
      },
    );
  }

  /**
   * One request with everything the requester is allowed to see about it.
   *
   * The driver's name and phone are released ONLY once that driver has
   * accepted, and only for that trip. The same basis releases the requester's
   * phone to the driver.
   */
  public async getRequestWithTrip(requestItemId: number): Promise<IRequestWithTrip> {
    return guard(
      'This booking could not be loaded.',
      'Go back to your bookings and open it again. If it keeps failing, ' +
        'contact the Transportation Coordinator.',
      async () => {
        const sp = getSP();

        const raw: any = await sp.web.lists
          .getByTitle(LIST.transportRequests)
          .items.getById(requestItemId)
          .select(...SELECT_REQUEST)
          .expand(...EXPAND_REQUEST)();

        const request = toRequest(raw);
        const result: IRequestWithTrip = { request, timeline: [] };

        // No trip until a driver accepts. Nothing further is readable, and
        // nothing further should be.
        if (request.status !== 'Allocated' && request.status !== 'Completed') {
          return result;
        }

        const tripRows: any[] = await sp.web.lists
          .getByTitle(LIST.tripAssignments)
          .items.select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)
          .filter(`Request/Id eq ${requestItemId}`)
          .orderBy('Created', false)
          .top(1)();

        if (tripRows.length === 0) return result;

        const trip = toTrip(tripRows[0]);
        result.trip = trip;

        const [driverRows, vehicleRows, timeline] = await Promise.all([
          trip.driver
            ? sp.web.lists
                .getByTitle(LIST.driversMaster)
                .items.select('Id', 'Title', 'Phone')
                .filter(`Id eq ${trip.driver.id}`)
                .top(1)()
            : Promise.resolve([] as any[]),
          trip.vehicle
            ? sp.web.lists
                .getByTitle(LIST.vehicleMaster)
                .items.select('Id', 'PlateNo', 'Make', 'Model', 'Colour', 'VehicleType/Title')
                .expand('VehicleType')
                .filter(`Id eq ${trip.vehicle.id}`)
                .top(1)()
            : Promise.resolve([] as any[]),
          this.getTimeline(trip.id),
        ]);

        if (driverRows.length > 0) {
          result.driverName = driverRows[0].Title;
          result.driverPhone = driverRows[0].Phone;
        }
        if (vehicleRows.length > 0) {
          const v = vehicleRows[0];
          result.vehiclePlate = v.PlateNo;
          result.vehicleMake = v.Make;
          result.vehicleModel = v.Model;
          result.vehicleColour = v.Colour;
          result.vehicleTypeName = v.VehicleType?.Title;
        }
        result.timeline = timeline;

        return result;
      },
    );
  }

  public async getTimeline(tripItemId: number): Promise<ITripStatusHistory[]> {
    return guard(
      'The trip timeline could not be loaded.',
      'Refresh the page to try again.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.tripStatusHistory)
          .items.select(...SELECT_HISTORY)
          .expand(...EXPAND_HISTORY)
          .filter(`Trip/Id eq ${tripItemId}`)
          .orderBy('ChangedOn', true)
          .top(50)();

        return rows.map(toHistory);
      },
    );
  }
}
