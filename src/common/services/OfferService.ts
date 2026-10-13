/**
 * Open offers for the signed-in driver. READS ONLY - the claim race lives in
 * DispatchService.acceptOffer.
 */

import { LIST } from '../config/AppConfig';
import { IOfferCard, OFFER_STATUS, REQUEST_STATUS } from '../models';
import { IOfferService } from './IServices';
import {
  EXPAND_OFFER,
  SELECT_OFFER,
  toOffer,
  toRequest,
} from './mappers';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

export class OfferService implements IOfferService {

  /**
   * Open offers for this driver. Polled every 20 seconds.
   *
   * Reads only the offers made to THIS driver. A driver never sees another
   * driver's trips.
   */
  public async getOpenOffers(driverId: number): Promise<IOfferCard[]> {
    return guard(
      'Your open offers could not be loaded.',
      'They will be retried automatically in a few seconds. If the banner ' +
        'stays empty, refresh the page.',
      async () => {
        const sp = getSP();

        const offerRows: any[] = await sp.web.lists
          .getByTitle(LIST.rideOffers)
          .items.select(...SELECT_OFFER)
          .expand(...EXPAND_OFFER)
          .filter(`DriverId eq ${driverId} and ResponseStatus eq '${OFFER_STATUS.Pending}'`)
          .orderBy('OfferedOn', true)
          .top(50)();

        if (offerRows.length === 0) return [];

        const offers = offerRows.map(toOffer);
        const requestIds = offers
          .map((o) => o.request?.id)
          .filter((id): id is number => typeof id === 'number');

        if (requestIds.length === 0) return [];

        const filter = requestIds.map((id) => `Id eq ${id}`).join(' or ');
        const requestRows: any[] = await sp.web.lists
          .getByTitle(LIST.transportRequests)
          .items.select(
            'Id', 'Title', 'RequestType', 'RequiredDateTime',
            'ReturnDateTime', 'NoOfPax', 'ItemDescription', 'ItemNotes',
            'OtherVehicleText', 'Priority', 'ReqStatus', 'OfferClosesAt',
            'Created', 'Modified',
            // The driver needs the notes as much as the address: a pin reading
            // "KIZAD, Abu Dhabi" does not say which gate.
            'PickupText', 'PickupNotes', 'DropOffText', 'DropOffNotes',
            'Requester/Id', 'Requester/Title', 'Requester/EMail',
            'VehicleTypeRequested/Id', 'VehicleTypeRequested/Title',
          )
          .expand('Requester', 'VehicleTypeRequested')
          .filter(filter)
          .top(50)();

        const byId = new Map<number, any>();
        requestRows.forEach((r) => byId.set(r.Id, r));

        const now = Date.now();
        const cards: IOfferCard[] = [];

        offers.forEach((offer) => {
          const raw = offer.request ? byId.get(offer.request.id) : undefined;
          if (!raw) return;

          const request = toRequest(raw);

          // An offer whose window has passed is not shown, even if the Expire
          // flow has not caught up with it yet.
          const closesAt = request.offerClosesAt?.getTime() ?? 0;
          const secondsRemaining = Math.max(0, Math.round((closesAt - now) / 1000));
          if (request.status !== REQUEST_STATUS.Broadcasting || secondsRemaining <= 0) return;

          cards.push({
            offer,
            request,
            requesterName: request.requester?.title ?? '',
            // THE OFFER CARD DELIBERATELY OMITS THE REQUESTER'S PHONE NUMBER
            // UNTIL THE DRIVER ACCEPTS.
            requesterPhone: undefined,
            vehicleTypeName: request.vehicleTypeRequested?.title ?? '',
            secondsRemaining,
          });
        });

        return cards;
      },
    );
  }
}
