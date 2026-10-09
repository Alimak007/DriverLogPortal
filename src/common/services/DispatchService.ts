/**
 * The dispatch engine. Every write in the solution goes through this file.
 *
 * It used to call Power Automate over HTTP. It cannot: the "When an HTTP
 * request is received" trigger is a PREMIUM connector and this tenant has only
 * standard ones. Polling triggers wake every five minutes, which is useless
 * for a driver tapping Accept. So the eligibility rules, the claim race and
 * every status transition now run here, in the browser, writing the lists as
 * the signed-in user.
 *
 * Flows still exist. They send email and run the scheduled jobs. This file
 * never sends mail, never writes Daily Trip Sheet, never writes
 * NotificationLog and never sets an offer to Expired - see the "owned by
 * flows" notes against the individual methods.
 *
 * NOTHING HERE THROWS FOR AN EXPECTED BUSINESS OUTCOME. Losing the claim race
 * is the normal case for every driver but one, so it comes back as
 * { ok: false, code: 'TAKEN' }, not as an exception.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LIST } from '../config/AppConfig';
import {
  IAllocationRule,
  IDispatchSnapshot,
  IDriver,
  IRideOffer,
  ITransportRequest,
  ITripAssignment,
  ITripCard,
  IVehicle,
  OFFER_STATUS,
  REQUEST_STATUS,
  RejectionReason,
  RequestType,
  TRIP_STATUS,
  TripStatus,
} from '../models';
import {
  explainDriverIneligibility,
  freeVehicles,
  isDriverEligible,
  isLegalTransition,
} from '../rules/dispatchRules';
import { formatRequestRef, formatTripRef } from '../rules/identifiers';
import { OPEN_TRIP_STATUSES, PENDING_STATUSES } from '../rules/statusTransitions';
import { toIso } from '../utils/dates';
import {
  EXPAND_DRIVER,
  EXPAND_OFFER,
  EXPAND_REQUEST,
  EXPAND_TRIP,
  EXPAND_VEHICLE,
  SELECT_DRIVER,
  SELECT_OFFER,
  SELECT_REQUEST,
  SELECT_TRIP,
  SELECT_VEHICLE,
  toAllocationRule,
  toAvailabilityBlock,
  toDriver,
  toOffer,
  toRequest,
  toTrip,
  toVehicle,
  EXPAND_ALLOCATION_RULE,
  EXPAND_BLOCK,
  SELECT_ALLOCATION_RULE,
  SELECT_BLOCK,
} from './mappers';
import { getSP, guard, newCorrelationId } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

// -----------------------------------------------------------------------------
// Result type
// -----------------------------------------------------------------------------

export type DispatchErrorCode =
  | 'TAKEN'               // another driver won the race - EXPECTED, not an error
  | 'NO_VEHICLE'
  | 'NOT_OPEN'
  | 'ILLEGAL_TRANSITION'
  | 'ALREADY_ALLOCATED'
  | 'NOT_PERMITTED'
  | 'UNKNOWN';

export type DispatchResult<T = void> =
  | { ok: true; data: T }
  | { ok: false; code: DispatchErrorCode; message: string };

function ok<T>(data: T): DispatchResult<T> {
  return { ok: true, data };
}
function fail<T = void>(code: DispatchErrorCode, message: string): DispatchResult<T> {
  return { ok: false, code, message };
}

// -----------------------------------------------------------------------------
// Inputs
// -----------------------------------------------------------------------------

/**
 * Permissions are now the browser's problem.
 *
 * A flow used to write on the driver's behalf under a service identity. There
 * is no service identity in a browser, so drivers need Contribute on
 * RideOffers, TripAssignments and TripStatusHistory in their own right.
 *
 * A 403 therefore means the permission grant is missing, not that the user did
 * something wrong. Name the list so whoever reads the message can fix it, and
 * do NOT retry or fall back - there is nothing to fall back to.
 */
export function mapWriteError<T = void>(e: unknown, listName: string): DispatchResult<T> {
  const status =
    (e as { status?: number })?.status ?? (e as { statusCode?: number })?.statusCode;
  const text = String((e as { message?: string })?.message ?? e);

  if (status === 403 || text.indexOf('403') !== -1 || /denied|unauthoriz/i.test(text)) {
    return {
      ok: false,
      code: 'NOT_PERMITTED',
      message:
        `You do not have permission to write to the ${listName} list. ` +
        `Ask a Fleet Admin to grant your group Contribute on ${listName}. ` +
        `Nothing was saved.`,
    };
  }

  console.error(`[FleetPortal] write to ${listName} failed`, e);
  return {
    ok: false,
    code: 'UNKNOWN',
    message: `The change could not be saved to ${listName}. Nothing was saved. Try again.`,
  };
}

export interface INewRequestInput {
  requestType: RequestType;
  propertyId: number;
  pickupLocationId: number;
  dropOffLocationId: number;
  requiredDateTime: Date;
  returnDateTime?: Date;
  noOfPax?: number;
  vehicleTypeId: number;
  otherVehicleText?: string;
  itemDescription?: string;
  itemNotes?: string;
  remarks?: string;
}

export type CancelArgs =
  | { kind: 'requester'; requestItemId: number; reason: string }
  | { kind: 'driver'; tripItemId: number; reason: string }
  | { kind: 'coordinator'; requestItemId: number; reason: string };

/** The rule actually in force, with the documented defaults applied. */
export interface IEffectiveRule {
  offerWindowMinutes: number;
  maxConcurrentTripsPerDriver: number;
  advanceReleaseHours: number;
  mode: 'Automatic' | 'Manual';
}

const RULE_DEFAULTS: IEffectiveRule = {
  // 15, not 5: the offer e-mail is sent by a flow and may take minutes to
  // arrive, so the window has to outlast the delay.
  offerWindowMinutes: 15,
  maxConcurrentTripsPerDriver: 1,
  advanceReleaseHours: 4,
  mode: 'Automatic',
};

export interface IDispatchService {
  createRequest(input: INewRequestInput): Promise<DispatchResult<{ id: number; ref: string }>>;
  broadcast(requestItemId: number): Promise<DispatchResult<{ offerCount: number }>>;
  sweepPendingBroadcasts(): Promise<void>;
  acceptOffer(offerItemId: number): Promise<DispatchResult<{ tripItemId: number }>>;
  rejectOffer(offerItemId: number, reason: RejectionReason, notes?: string): Promise<DispatchResult>;
  updateTripStatus(tripItemId: number, toStatus: TripStatus, odometer?: number): Promise<DispatchResult>;
  manualAssign(
    requestItemId: number,
    driverItemId: number,
    vehicleItemId: number,
    isOverride: boolean,
    overrideReason?: string,
  ): Promise<DispatchResult<{ tripItemId: number }>>;
  reassign(tripItemId: number, newDriverId: number, newVehicleId: number, reason: string): Promise<DispatchResult>;
  cancel(args: CancelArgs): Promise<DispatchResult<{ message: string }>>;
  rejectRequest(requestItemId: number, reason: string): Promise<DispatchResult>;
  setEta(tripItemId: number, etaMinutes: number): Promise<DispatchResult>;

  // Reads behind the Coordinator's screens.
  getSnapshot(propertyIds: number[], completedFrom: Date, completedTo: Date): Promise<IDispatchSnapshot>;
  getOffersForRequest(requestItemId: number): Promise<IRideOffer[]>;
  getTripsInRange(propertyIds: number[], from: Date, to: Date): Promise<ITripAssignment[]>;
}

// -----------------------------------------------------------------------------

function orFilter(field: string, values: (string | number)[]): string {
  if (values.length === 0) return '';
  return values
    .map((v) => (typeof v === 'number' ? `${field} eq ${v}` : `${field} eq '${v}'`))
    .join(' or ');
}

/** Zero-pads to three digits. padStart is not in this project's lib target. */
function pad3(n: number): string {
  const s = String(n);
  return s.length >= 3 ? s : new Array(3 - s.length + 1).join('0') + s;
}

function and(...clauses: string[]): string {
  return clauses.filter((c) => c.length > 0).map((c) => `(${c})`).join(' and ');
}

export class DispatchService implements IDispatchService {
  private readonly ctx: WebPartContext;
  private readonly siteUrl: string;

  constructor(context: WebPartContext) {
    this.ctx = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // ===========================================================================
  // 4.1 Create a request
  // ===========================================================================

  public async createRequest(
    input: INewRequestInput,
  ): Promise<DispatchResult<{ id: number; ref: string }>> {
    try {
      // A request with no property can never be allocated: every eligibility
      // read filters on it. Refuse to create one rather than leaving a booking
      // that silently escalates for ever.
      if (!input.propertyId) {
        return fail(
          'UNKNOWN',
          'This booking has no property, so no driver could be found for it. ' +
            'Choose a pickup point that belongs to a property, or ask a Fleet ' +
            'Admin to set the Property column on the Locations list.',
        );
      }
      if (!input.vehicleTypeId) {
        return fail('UNKNOWN', 'This booking has no vehicle type.');
      }

      const sp = getSP();
      const list = sp.web.lists.getByTitle(LIST.transportRequests);

      const payload: Record<string, unknown> = {
        Title: 'NEW',
        RequestType: input.requestType,
        PropertyId: input.propertyId,
        PickupLocationId: input.pickupLocationId,
        DropOffLocationId: input.dropOffLocationId,
        RequiredDateTime: toIso(input.requiredDateTime),
        ReturnDateTime: toIso(input.returnDateTime),
        VehicleTypeRequestedId: input.vehicleTypeId,
        Priority: 'Normal',
        ReqStatus: REQUEST_STATUS.Submitted,
        RequesterId: this.ctx.pageContext.legacyPageContext.userId,
      };
      if (input.noOfPax !== undefined) payload.NoOfPax = input.noOfPax;
      if (input.otherVehicleText) payload.OtherVehicleText = input.otherVehicleText;
      if (input.itemDescription) payload.ItemDescription = input.itemDescription;
      if (input.itemNotes) payload.ItemNotes = input.itemNotes;
      if (input.remarks) payload.Remarks = input.remarks;

      const added: any = await list.items.add(payload);
      const itemId: number = added?.Id ?? added?.data?.Id;
      if (!itemId) return fail('UNKNOWN', 'The request was not saved. Nothing has been booked.');

      // The reference is derived from the SharePoint item ID. A counter item
      // would race under concurrent intake.
      const ref = formatRequestRef(itemId);
      await list.items.getById(itemId).update({ Title: ref });

      const rule = await this.getEffectiveRule(input.propertyId, input.requestType);

      // Advance booking: hold it. A booking three weeks out would otherwise sit
      // in drivers' dashboards for weeks. A flow releases it back to Submitted.
      const releaseAt = new Date(
        input.requiredDateTime.getTime() - rule.advanceReleaseHours * 3600000,
      );
      if (releaseAt.getTime() > Date.now()) {
        await list.items.getById(itemId).update({
          ReqStatus: REQUEST_STATUS.Held,
          ReleaseAt: toIso(releaseAt),
        });
        return ok({ id: itemId, ref });
      }

      if (rule.mode === 'Manual') {
        await list.items.getById(itemId).update({
          ReqStatus: REQUEST_STATUS.AwaitingCoordinator,
        });
        return ok({ id: itemId, ref });
      }

      await this.broadcast(itemId);
      return ok({ id: itemId, ref });
  
    } catch (e) {
      return mapWriteError(e, 'TransportRequests');
    }
  }

  // ===========================================================================
  // 4.2 Broadcast
  // ===========================================================================

  /**
   * IDEMPOTENT. Running it twice on the same request must not create duplicate
   * offers, because sweepPendingBroadcasts() calls it on every dashboard poll.
   */
  public async broadcast(
    requestItemId: number,
  ): Promise<DispatchResult<{ offerCount: number }>> {
    try {
      const sp = getSP();

      const existing: any[] = await sp.web.lists
        .getByTitle(LIST.rideOffers)
        .items.select('Id')
        .filter(`RequestId eq ${requestItemId} and ResponseStatus eq '${OFFER_STATUS.Pending}'`)
        .top(1)();
      if (existing.length > 0) return ok({ offerCount: 0 });

      const raw: any = await sp.web.lists
        .getByTitle(LIST.transportRequests)
        .items.getById(requestItemId)
        .select(...SELECT_REQUEST)
        .expand(...EXPAND_REQUEST)();
      const request = toRequest(raw);

      const propertyId = request.property?.id;
      const typeId = request.vehicleTypeRequested?.id;
      if (propertyId === undefined || typeId === undefined) {
        return fail('UNKNOWN', 'This request has no property or vehicle type, so it cannot be allocated.');
      }

      const rule = await this.getEffectiveRule(propertyId, request.requestType);

      const windowFrom = request.requiredDateTime;
      const windowTo = request.returnDateTime
        ? request.returnDateTime
        : new Date(windowFrom.getTime() + 2 * 3600000);

      // Four reads. Never a loop of single-item fetches.
      const [links, blocks, openTrips, drivers] = await Promise.all([
        sp.web.lists
          .getByTitle(LIST.dlClientMaster)
          .items.select('Id', 'DriverID/Id', 'PropertyRef/Id')
          .expand('DriverID', 'PropertyRef')
          .filter(`PropertyRef/Id eq ${propertyId}`)
          .top(2000)(),
        sp.web.lists
          .getByTitle(LIST.availabilityBlocks)
          .items.select(...SELECT_BLOCK)
          .expand(...EXPAND_BLOCK)
          .filter(
            `FromDateTime lt '${toIso(windowTo)}' and ToDateTime gt '${toIso(windowFrom)}'`,
          )
          .top(500)(),
        sp.web.lists
          .getByTitle(LIST.tripAssignments)
          .items.select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)
          .filter(orFilter('TripStatus', OPEN_TRIP_STATUSES))
          .top(500)(),
        sp.web.lists
          .getByTitle(LIST.driversMaster)
          .items.select(...SELECT_DRIVER)
          .expand(...EXPAND_DRIVER)
          .filter(`IsActive eq 1 and Status ne 'Inactive'`)
          .top(500)(),
      ]);

      const allocatedDriverIds = new Set<number>(
        (links as any[])
          .map((l) => l.DriverID?.Id)
          .filter((id): id is number => typeof id === 'number'),
      );
      const blockList = (blocks as any[]).map(toAvailabilityBlock);
      const tripList = (openTrips as any[]).map(toTrip);
      const driverList = (drivers as any[]).map(toDriver);

      const eligible = driverList.filter((d) =>
        isDriverEligible(d, {
          allocatedDriverIds,
          vehicleTypeId: typeId,
          windowTo,
          blocks: blockList,
          openTrips: tripList,
          ruleMaxConcurrent: rule.maxConcurrentTripsPerDriver,
        }),
      );

      const vehicles = await this.readCandidateVehicles(typeId, propertyId);
      const free = freeVehicles(vehicles, tripList, blockList);

      const requestList = sp.web.lists.getByTitle(LIST.transportRequests);

      if (eligible.length === 0 || free.length === 0) {
        // Escalation is a legitimate outcome, but escalating SILENTLY is not.
        // "No eligible drivers" is useless on its own; the per-driver reasons
        // point straight at the data that caused it.
        const ctx = {
          allocatedDriverIds,
          vehicleTypeId: typeId,
          windowTo,
          blocks: blockList,
          openTrips: tripList,
          ruleMaxConcurrent: rule.maxConcurrentTripsPerDriver,
        };
        const report: string[] = [
          `[FleetPortal] ${request.requestId} escalated to the Coordinator.`,
          `  property=${propertyId} vehicleType=${typeId}`,
          `  window ${toIso(windowFrom)} -> ${toIso(windowTo)}`,
          `  DLClientMaster rows for this property: ${links.length}`,
          `  allocated driver ids: ${Array.from(allocatedDriverIds).join(', ') || 'NONE'}`,
          `  drivers read: ${driverList.length}, eligible: ${eligible.length}`,
          `  vehicles of this type at this property: ${vehicles.length}, free: ${free.length}`,
        ];
        driverList.forEach((d) => {
          const reasons = explainDriverIneligibility(d, ctx).join('; ');
          report.push(`  - ${d.driverName} (id ${d.id}): ${reasons || 'ELIGIBLE'}`);
        });
        console.warn(report.join(String.fromCharCode(10)));

        await requestList.items.getById(requestItemId).update({
          ReqStatus: REQUEST_STATUS.AwaitingCoordinator,
          OfferClosesAt: toIso(new Date()),
        });
        return ok({ offerCount: 0 });
      }

      // Status BEFORE the offers. A driver who sees a card while the request
      // still reads Submitted is rejected by the guard in acceptOffer and has no
      // idea why.
      await requestList.items.getById(requestItemId).update({
        ReqStatus: REQUEST_STATUS.Broadcasting,
        OfferClosesAt: toIso(new Date(Date.now() + rule.offerWindowMinutes * 60000)),
      });

      const offeredOn = toIso(new Date());
      const [batched, execute] = sp.batched();
      const offerList = batched.web.lists.getByTitle(LIST.rideOffers);

      eligible.forEach((d) => {
        offerList.items
          .add({
            // Spec format: FB-0042-D017. Readable without expanding lookups.
          Title: `${request.requestId}-D${pad3(d.id)}`,
            RequestId: requestItemId,
            DriverId: d.id,
            OfferedOn: offeredOn,
            ResponseStatus: OFFER_STATUS.Pending,
            // Vehicle deliberately left empty - bound at claim time.
          })
          .catch(() => undefined);
      });
      await execute();

      return ok({ offerCount: eligible.length });
  
    } catch (e) {
      return mapWriteError(e, 'RideOffers');
    }
  }

  // ===========================================================================
  // 4.3 Sweep
  // ===========================================================================

  /**
   * Called by the Dispatch Dashboard on every poll.
   *
   * Covers an advance booking a flow has just released back to Submitted, and
   * a request whose creator closed their browser mid-submit.
   *
   * KNOWN LIMITATION, BY DESIGN: a released advance booking is not broadcast
   * until a Coordinator or driver has a dashboard open. Advance bookings
   * release four hours before the trip, inside working hours, so this is
   * acceptable for phase 1. It is deliberately NOT solved with a flow - the
   * eligibility rules must live in one place only.
   */
  public async sweepPendingBroadcasts(): Promise<void> {
    const rows: any[] = await getSP()
      .web.lists.getByTitle(LIST.transportRequests)
      .items.select('Id')
      .filter(`ReqStatus eq '${REQUEST_STATUS.Submitted}'`)
      .top(25)();

    for (const r of rows) {
      try {
        await this.broadcast(r.Id);
      } catch (e) {
        // One bad request must not stop the sweep for the rest.
        console.error('[FleetPortal] sweep failed for request', r.Id, e);
      }
    }
  }

  // ===========================================================================
  // 4.4 Accept - the claim race
  // ===========================================================================

  /**
   * Two drivers can tap Accept in the same second. SharePoint has no
   * transactions, so the request item's ETag is the lock.
   */
  public async acceptOffer(
    offerItemId: number,
  ): Promise<DispatchResult<{ tripItemId: number }>> {
    try {
      const sp = getSP();
      const correlationId = newCorrelationId();

      // 1. The offer must still be open.
      const offerRaw: any = await sp.web.lists
        .getByTitle(LIST.rideOffers)
        .items.getById(offerItemId)
        .select(...SELECT_OFFER)
        .expand(...EXPAND_OFFER)();
      const offer = toOffer(offerRaw);

      if (offer.responseStatus !== OFFER_STATUS.Pending) {
        return fail('NOT_OPEN', 'This offer is no longer open.');
      }
      const requestItemId = offer.request?.id;
      const driverId = offer.driver?.id;
      if (requestItemId === undefined || driverId === undefined) {
        return fail('UNKNOWN', 'This offer is incomplete and cannot be accepted.');
      }

      // 2. Read the request WITH its ETag. PnPjs strips it under the default
      //    nometadata parser, so go through SPHttpClient.
      const itemUrl =
        `${this.siteUrl}/_api/web/lists/getbytitle('${LIST.transportRequests}')` +
        `/items(${requestItemId})`;

      const getRes: SPHttpClientResponse = await this.ctx.spHttpClient.get(
        itemUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=minimalmetadata',
            // Opt out of OData v4, which SPHttpClient sends by default. The
            // POST below already does this, and the two must agree.
            'odata-version': '',
          },
        },
      );
      if (!getRes.ok) {
        return fail('UNKNOWN', 'Could not read this request. Try again.');
      }
      const requestItem: any = await getRes.json();

      // The RESPONSE HEADER is the reliable source. The body key depends on
      // the OData version negotiated: 'odata.etag' under v3 minimalmetadata,
      // '@odata.etag' under v4. Reading only one of them is what produced
      // IF-MATCH: "undefined" and a 412 for every driver, winner included.
      const etag: string | undefined =
        getRes.headers.get('ETag') ??
        requestItem['odata.etag'] ??
        requestItem['@odata.etag'];

      if (!etag) {
        // Without an ETag there is no lock, and proceeding would let two
        // drivers both win. Refuse rather than allocate unsafely.
        console.error(
          '[FleetPortal] No ETag on the request read; the claim lock cannot be ' +
            'armed. Response keys:',
          Object.keys(requestItem ?? {}),
        );
        return fail(
          'UNKNOWN',
          'Could not accept this trip safely. Nothing has been assigned. ' +
            'Try again, and tell support if it keeps happening.',
        );
      }

      // 3. Cheap check - catches most losers without a write.
      if (requestItem.ReqStatus !== REQUEST_STATUS.Broadcasting) {
        return fail('TAKEN', 'Taken by another driver');
      }

      // 4. Bind a vehicle, re-verifying availability in this same read.
      const propertyId: number | undefined = requestItem.PropertyId;
      const typeId: number | undefined = requestItem.VehicleTypeRequestedId;
      if (propertyId === undefined || typeId === undefined) {
        return fail('UNKNOWN', 'This request has no property or vehicle type.');
      }
      const vehicleId = await this.pickFreeVehicle(typeId, propertyId, driverId);
      if (vehicleId === undefined) {
        return fail('NO_VEHICLE', 'No vehicle is available for this trip');
      }

      // 5. THE LOCK. Applies only if nobody has touched the item since step 2.
      const claimRes: SPHttpClientResponse = await this.ctx.spHttpClient.post(
        itemUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': etag, // on the POST, never on the GET
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({ ReqStatus: REQUEST_STATUS.Allocated }),
        },
      );

      if (claimRes.status === 412) {
        // Precondition Failed. Another driver got there first.
        return fail('TAKEN', 'Taken by another driver');
      }
      if (claimRes.status !== 204 && !claimRes.ok) {
        return fail('UNKNOWN', 'Could not accept this trip. Try again.');
      }

      // 6. We won.
      const tripItemId = await this.completeClaim(
        offerItemId,
        requestItemId,
        driverId,
        vehicleId,
        correlationId,
      );
      return ok({ tripItemId });
  
    } catch (e) {
      return mapWriteError(e, 'TripAssignments');
    }
  }

  private async completeClaim(
    offerItemId: number,
    requestItemId: number,
    driverId: number,
    vehicleId: number,
    correlationId: string,
  ): Promise<number> {
    const sp = getSP();
    const now = toIso(new Date());

    const tripList = sp.web.lists.getByTitle(LIST.tripAssignments);
    const added: any = await tripList.items.add({
      Title: 'NEW',
      RequestId: requestItemId,
      DriverId: driverId,
      VehicleId: vehicleId,
      AcceptedOn: now,
      TripStatus: TRIP_STATUS.Assigned,
      AllocationMode: 'Automatic',
      IsOverride: false,
    });
    const tripItemId: number = added?.Id ?? added?.data?.Id;

    await tripList.items.getById(tripItemId).update({ Title: formatTripRef(tripItemId) });

    await sp.web.lists
      .getByTitle(LIST.transportRequests)
      .items.getById(requestItemId)
      .update({ AcceptedTripId: tripItemId });

    const offerList = sp.web.lists.getByTitle(LIST.rideOffers);
    await offerList.items.getById(offerItemId).update({
      ResponseStatus: OFFER_STATUS.Accepted,
      RespondedOn: now,
      VehicleId: vehicleId,
    });

    // Every other open offer is Closed - someone else won. NOT Expired, which
    // means the window ran out.
    const others: any[] = await offerList.items
      .select('Id')
      .filter(
        `RequestId eq ${requestItemId} and ResponseStatus eq '${OFFER_STATUS.Pending}'`,
      )
      .top(200)();

    if (others.length > 0) {
      const [batched, execute] = sp.batched();
      const batchedOffers = batched.web.lists.getByTitle(LIST.rideOffers);
      others.forEach((o) => {
        batchedOffers.items
          .getById(o.Id)
          .update({ ResponseStatus: OFFER_STATUS.Closed })
          .catch(() => undefined);
      });
      await execute();
    }

    await this.setDriverStatus(driverId, 'On Trip');
    await this.setVehicleStatus(vehicleId, 'On Trip');

    await this.writeHistory(tripItemId, '', TRIP_STATUS.Assigned,
      `Accepted from broadcast. Ref ${correlationId}`);

    return tripItemId;
  }

  // ===========================================================================
  // 4.5 Reject an offer
  // ===========================================================================

  public async rejectOffer(
    offerItemId: number,
    reason: RejectionReason,
    notes?: string,
  ): Promise<DispatchResult> {
    try {
      const sp = getSP();
      const list = sp.web.lists.getByTitle(LIST.rideOffers);

      const raw: any = await list.items
        .getById(offerItemId)
        .select('Id', 'ResponseStatus')();

      if (raw.ResponseStatus !== OFFER_STATUS.Pending) {
        return fail('NOT_OPEN', 'This offer is no longer open.');
      }

      await list.items.getById(offerItemId).update({
        ResponseStatus: OFFER_STATUS.Rejected,
        RejectionReason: reason,
        RejectionNotes: notes,
        RespondedOn: toIso(new Date()),
      });

      // The request status is deliberately NOT changed. Other drivers may still
      // accept; a flow escalates it if nobody does.
      return ok(undefined);
  
    } catch (e) {
      return mapWriteError(e, 'RideOffers');
    }
  }

  // ===========================================================================
  // 4.6 The four taps
  // ===========================================================================

  public async updateTripStatus(
    tripItemId: number,
    toStatus: TripStatus,
    odometer?: number,
  ): Promise<DispatchResult> {
    try {
      const sp = getSP();
      const list = sp.web.lists.getByTitle(LIST.tripAssignments);

      const raw: any = await list.items.getById(tripItemId).select('Id', 'TripStatus')();
      const fromStatus: TripStatus = raw.TripStatus;

      if (!isLegalTransition(fromStatus, toStatus)) {
        return fail(
          'ILLEGAL_TRANSITION',
          `A trip that is ${fromStatus} cannot be set to ${toStatus}.`,
        );
      }

      const needsOdometer = toStatus === TRIP_STATUS.EnRoute || toStatus === TRIP_STATUS.Completed;
      if (needsOdometer && (odometer === undefined || isNaN(odometer))) {
        const which = toStatus === TRIP_STATUS.EnRoute ? 'opening' : 'closing';
        return fail('UNKNOWN', `The ${which} odometer reading is required for this step.`);
      }

      const now = toIso(new Date());
      const values: Record<string, unknown> = { TripStatus: toStatus };

      if (toStatus === TRIP_STATUS.EnRoute) {
        values.ActualStartTime = now;
        values.OpeningReading = odometer;
      } else if (toStatus === TRIP_STATUS.InTransit) {
        values.ActualPickupTime = now;
      } else if (toStatus === TRIP_STATUS.Completed) {
        values.ActualDropTime = now;
        values.ClosingReading = odometer;
      }

      await list.items.getById(tripItemId).update(values);
      await this.writeHistory(tripItemId, fromStatus, toStatus);

      // OWNED BY A FLOW on Completed: setting the request to Completed, freeing
      // the driver and vehicle, and writing Daily Trip Sheet. Doing any of it
      // here would double-write.
      return ok(undefined);
  
    } catch (e) {
      return mapWriteError(e, 'TripAssignments');
    }
  }

  // ===========================================================================
  // 4.7 Manual assignment
  // ===========================================================================

  public async manualAssign(
    requestItemId: number,
    driverItemId: number,
    vehicleItemId: number,
    isOverride: boolean,
    overrideReason?: string,
  ): Promise<DispatchResult<{ tripItemId: number }>> {
    try {
      if (isOverride && (!overrideReason || overrideReason.trim().length === 0)) {
        return fail(
          'NOT_PERMITTED',
          'An override reason is required for an out-of-policy assignment.',
        );
      }

      const sp = getSP();
      const correlationId = newCorrelationId();
      const requestList = sp.web.lists.getByTitle(LIST.transportRequests);

      const raw: any = await requestList.items
        .getById(requestItemId)
        .select('Id', 'ReqStatus')();

      if (PENDING_STATUSES.indexOf(raw.ReqStatus) === -1 && raw.ReqStatus !== REQUEST_STATUS.Held) {
        return fail('ALREADY_ALLOCATED', 'This request has already been allocated.');
      }

      // Withdrawn, not Closed. The Coordinator bypassed the broadcast, and the
      // reports count that separately.
      await this.setPendingOffers(requestItemId, OFFER_STATUS.Withdrawn);

      const now = toIso(new Date());
      const tripList = sp.web.lists.getByTitle(LIST.tripAssignments);
      const added: any = await tripList.items.add({
        Title: 'NEW',
        RequestId: requestItemId,
        DriverId: driverItemId,
        VehicleId: vehicleItemId,
        AcceptedOn: now,
        TripStatus: TRIP_STATUS.Assigned,
        AllocationMode: 'Manual',
        AllocatedById: this.ctx.pageContext.legacyPageContext.userId,
        IsOverride: isOverride,
        OverrideReason: overrideReason,
      });
      const tripItemId: number = added?.Id ?? added?.data?.Id;

      await tripList.items.getById(tripItemId).update({ Title: formatTripRef(tripItemId) });
      await requestList.items.getById(requestItemId).update({
        ReqStatus: REQUEST_STATUS.Allocated,
        AcceptedTripId: tripItemId,
      });

      await this.setDriverStatus(driverItemId, 'On Trip');
      await this.setVehicleStatus(vehicleItemId, 'On Trip');

      await this.writeHistory(tripItemId, '', TRIP_STATUS.Assigned,
        `Assigned manually. Ref ${correlationId}`);

      return ok({ tripItemId });
  
    } catch (e) {
      return mapWriteError(e, 'TripAssignments');
    }
  }

  // ===========================================================================
  // 4.8 Reassign
  // ===========================================================================

  public async reassign(
    tripItemId: number,
    newDriverId: number,
    newVehicleId: number,
    reason: string,
  ): Promise<DispatchResult> {
    try {
      if (!reason || reason.trim().length === 0) {
        return fail('NOT_PERMITTED', 'A reason is required to reassign a trip.');
      }

      const sp = getSP();
      const list = sp.web.lists.getByTitle(LIST.tripAssignments);

      const raw: any = await list.items
        .getById(tripItemId)
        .select(...SELECT_TRIP)
        .expand(...EXPAND_TRIP)();
      const trip = toTrip(raw);

      if (trip.driver) await this.setDriverStatus(trip.driver.id, 'Available');
      if (trip.vehicle) await this.setVehicleStatus(trip.vehicle.id, 'Available');

      await list.items.getById(tripItemId).update({
        DriverId: newDriverId,
        VehicleId: newVehicleId,
        IsOverride: true,
        OverrideReason: reason,
        // TripStatus deliberately unchanged - the reassignment is the event,
        // not a status change.
      });

      await this.setDriverStatus(newDriverId, 'On Trip');
      await this.setVehicleStatus(newVehicleId, 'On Trip');

      const driverName = trip.driver?.title ?? 'the previous driver';
      await this.writeHistory(
        tripItemId,
        trip.tripStatus,
        trip.tripStatus,
        `Reassigned from ${driverName} to driver ${newDriverId}: ${reason}`,
      );

      return ok(undefined);
  
    } catch (e) {
      return mapWriteError(e, 'TripAssignments');
    }
  }

  // ===========================================================================
  // 4.9 Cancel
  // ===========================================================================

  public async cancel(args: CancelArgs): Promise<DispatchResult<{ message: string }>> {
    try {
      if (!args.reason || args.reason.trim().length === 0) {
        return fail('NOT_PERMITTED', 'A reason is required to cancel.');
      }

      const sp = getSP();
      const requestList = sp.web.lists.getByTitle(LIST.transportRequests);

      if (args.kind === 'driver') {
        const tripList = sp.web.lists.getByTitle(LIST.tripAssignments);
        const raw: any = await tripList.items
          .getById(args.tripItemId)
          .select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)();
        const trip = toTrip(raw);

        await tripList.items.getById(args.tripItemId).update({
          TripStatus: TRIP_STATUS.Cancelled,
          CancellationReason: args.reason,
        });

        // The booking is NOT cancelled. It goes back to the Coordinator for
        // reallocation, and is deliberately not re-broadcast automatically - a
        // cancellation close to the required time needs human judgement.
        if (trip.request) {
          await requestList.items.getById(trip.request.id).update({
            ReqStatus: REQUEST_STATUS.AwaitingCoordinator,
          });
        }

        if (trip.driver) await this.setDriverStatus(trip.driver.id, 'Available');
        if (trip.vehicle) await this.setVehicleStatus(trip.vehicle.id, 'Available');

        await this.writeHistory(
          args.tripItemId,
          trip.tripStatus,
          TRIP_STATUS.Cancelled,
          `Cancelled by the driver: ${args.reason}`,
        );

        return ok({ message: 'The trip was cancelled and sent back to the Coordinator.' });
      }

      const raw: any = await requestList.items
        .getById(args.requestItemId)
        .select('Id', 'ReqStatus', 'AcceptedTripId')();
      const status: string = raw.ReqStatus;

      if (args.kind === 'requester') {
        const cancellable: string[] = [
          REQUEST_STATUS.Submitted,
          REQUEST_STATUS.Held,
          REQUEST_STATUS.AwaitingCoordinator,
        ];

        if (cancellable.indexOf(status) === -1) {
          // Already allocated: change NOTHING. A flow alerts the Coordinator,
          // who cancels on the requester's behalf and notifies the driver.
          if (raw.AcceptedTripId) {
            await this.writeHistory(
              raw.AcceptedTripId,
              status,
              status,
              `Cancellation requested by the requester: ${args.reason}`,
            );
          }
          return ok({
            message:
              'A driver has already accepted this trip, so the Transportation ' +
              'Coordinator has been asked to cancel it. Your driver stays assigned ' +
              'until they do.',
          });
        }

        await requestList.items.getById(args.requestItemId).update({
          ReqStatus: REQUEST_STATUS.Cancelled,
          CancellationReason: args.reason,
        });
        await this.setPendingOffers(args.requestItemId, OFFER_STATUS.Withdrawn);
        return ok({ message: 'Your booking was cancelled.' });
      }

      // Coordinator
      await requestList.items.getById(args.requestItemId).update({
        ReqStatus: REQUEST_STATUS.Cancelled,
        CancellationReason: args.reason,
      });
      await this.setPendingOffers(args.requestItemId, OFFER_STATUS.Withdrawn);

      if (raw.AcceptedTripId) {
        const tripList = sp.web.lists.getByTitle(LIST.tripAssignments);
        const tripRaw: any = await tripList.items
          .getById(raw.AcceptedTripId)
          .select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)();
        const trip = toTrip(tripRaw);

        await tripList.items.getById(raw.AcceptedTripId).update({
          TripStatus: TRIP_STATUS.Cancelled,
          CancellationReason: args.reason,
        });
        if (trip.driver) await this.setDriverStatus(trip.driver.id, 'Available');
        if (trip.vehicle) await this.setVehicleStatus(trip.vehicle.id, 'Available');
        await this.writeHistory(
          raw.AcceptedTripId,
          trip.tripStatus,
          TRIP_STATUS.Cancelled,
          `Cancelled by the Coordinator: ${args.reason}`,
        );
      }

      return ok({ message: 'The request was cancelled.' });
  
    } catch (e) {
      return mapWriteError(e, 'TransportRequests');
    }
  }

  // ===========================================================================
  // 4.10 Reject a request
  // ===========================================================================

  public async rejectRequest(requestItemId: number, reason: string): Promise<DispatchResult> {
    try {
      if (!reason || reason.trim().length === 0) {
        return fail('NOT_PERMITTED', 'A reason is required to reject a request.');
      }

      const sp = getSP();
      const list = sp.web.lists.getByTitle(LIST.transportRequests);
      const raw: any = await list.items.getById(requestItemId).select('Id', 'ReqStatus')();

      const rejectable: string[] = [
        REQUEST_STATUS.Submitted,
        REQUEST_STATUS.Held,
        REQUEST_STATUS.Broadcasting,
        REQUEST_STATUS.AwaitingCoordinator,
      ];
      if (rejectable.indexOf(raw.ReqStatus) === -1) {
        return fail(
          'ALREADY_ALLOCATED',
          'An allocated request cannot be rejected - cancel it instead.',
        );
      }

      await list.items.getById(requestItemId).update({
        ReqStatus: REQUEST_STATUS.Rejected,
        RejectionReason: reason,
      });
      await this.setPendingOffers(requestItemId, OFFER_STATUS.Withdrawn);

      return ok(undefined);
  
    } catch (e) {
      return mapWriteError(e, 'TransportRequests');
    }
  }

  // ===========================================================================
  // 4.11 ETA
  // ===========================================================================

  public async setEta(tripItemId: number, etaMinutes: number): Promise<DispatchResult> {
    try {
      if (isNaN(etaMinutes) || etaMinutes < 0 || etaMinutes > 600) {
        return fail('UNKNOWN', 'Enter a whole number of minutes between 0 and 600.');
      }

      // ALWAYS write both. The UI renders "approx. 15 min - updated 09:14"; the
      // figure without the timestamp would look like a live calculation, and in
      // phase 1 it is typed in by hand.
      await getSP()
        .web.lists.getByTitle(LIST.tripAssignments)
        .items.getById(tripItemId)
        .update({
          EtaMinutes: Math.round(etaMinutes),
          EtaUpdatedOn: toIso(new Date()),
        });

      return ok(undefined);
  
    } catch (e) {
      return mapWriteError(e, 'TripAssignments');
    }
  }

  // ===========================================================================
  // Reads
  // ===========================================================================

  public async getSnapshot(
    propertyIds: number[],
    completedFrom: Date,
    completedTo: Date,
  ): Promise<IDispatchSnapshot> {
    return guard(
      'The dispatch dashboard could not be loaded.',
      'Refresh the page. If it keeps failing, the counters and grids below ' +
        'may be out of date - do not allocate from them.',
      async () => {
        const sp = getSP();
        const [batch, execute] = sp.batched();

        let requestRows: any[] = [];
        let tripRows: any[] = [];
        let driverRows: any[] = [];
        let vehicleRows: any[] = [];

        const requestFilter = and(
          orFilter('Property/Id', propertyIds),
          `(${orFilter('ReqStatus', [...PENDING_STATUSES, REQUEST_STATUS.Allocated])}) or ` +
            `(ReqStatus eq '${REQUEST_STATUS.Completed}' and Modified ge '${toIso(completedFrom)}' ` +
            `and Modified lt '${toIso(completedTo)}')`,
        );

        batch.web.lists
          .getByTitle(LIST.transportRequests)
          .items.select(...SELECT_REQUEST)
          .expand(...EXPAND_REQUEST)
          .filter(requestFilter)
          .orderBy('RequiredDateTime', true)
          .top(200)()
          .then((r: any[]) => { requestRows = r; })
          .catch(() => { requestRows = []; });

        const tripFilter =
          `(${orFilter('TripStatus', OPEN_TRIP_STATUSES)}) or ` +
          `(TripStatus eq '${TRIP_STATUS.Completed}' and ActualDropTime ge '${toIso(completedFrom)}' ` +
          `and ActualDropTime lt '${toIso(completedTo)}')`;

        batch.web.lists
          .getByTitle(LIST.tripAssignments)
          .items.select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)
          .filter(tripFilter)
          .top(200)()
          .then((r: any[]) => { tripRows = r; })
          .catch(() => { tripRows = []; });

        batch.web.lists
          .getByTitle(LIST.driversMaster)
          .items.select(...SELECT_DRIVER)
          .expand(...EXPAND_DRIVER)
          .filter(and('IsActive eq 1', orFilter('HomeProperty/Id', propertyIds)))
          .orderBy('Title', true)
          .top(500)()
          .then((r: any[]) => { driverRows = r; })
          .catch(() => { driverRows = []; });

        batch.web.lists
          .getByTitle(LIST.vehicleMaster)
          .items.select(...SELECT_VEHICLE)
          .expand(...EXPAND_VEHICLE)
          .filter(and("VehStatus ne 'Retired'", orFilter('PropertyRef/Id', propertyIds)))
          .orderBy('PlateNo', true)
          .top(500)()
          .then((r: any[]) => { vehicleRows = r; })
          .catch(() => { vehicleRows = []; });

        await execute();

        const requests: ITransportRequest[] = requestRows.map(toRequest);
        const trips: ITripAssignment[] = tripRows.map(toTrip);
        const drivers: IDriver[] = driverRows.map(toDriver);
        const vehicles: IVehicle[] = vehicleRows.map(toVehicle);

        const requestById = new Map<number, ITransportRequest>();
        requests.forEach((r) => requestById.set(r.id, r));

        const ongoingTrips: ITripCard[] = trips
          .filter((t) => t.tripStatus === TRIP_STATUS.EnRoute || t.tripStatus === TRIP_STATUS.InTransit)
          .map((trip) => {
            const request = trip.request ? requestById.get(trip.request.id) : undefined;
            return {
              trip,
              request: request as ITransportRequest,
              pickupName: request?.pickupLocation?.title ?? '',
              dropOffName: request?.dropOffLocation?.title ?? '',
              requesterName: request?.requester?.title ?? '',
              vehiclePlate: trip.vehicle?.title ?? '',
            };
          })
          .filter((c) => c.request !== undefined)
          .sort((a, b) => a.request.requiredDateTime.getTime() - b.request.requiredDateTime.getTime());

        const pendingRequests = requests
          .filter((r) => PENDING_STATUSES.indexOf(r.status) !== -1)
          .sort((a, b) => {
            const aEsc = a.status === REQUEST_STATUS.AwaitingCoordinator ? 0 : 1;
            const bEsc = b.status === REQUEST_STATUS.AwaitingCoordinator ? 0 : 1;
            if (aEsc !== bEsc) return aEsc - bEsc;
            return a.created.getTime() - b.created.getTime();
          });

        const openTrips = trips.filter((t) => OPEN_TRIP_STATUSES.indexOf(t.tripStatus) !== -1);

        const tripsPerDriver = new Map<number, number>();
        openTrips.forEach((t) => {
          if (!t.driver) return;
          tripsPerDriver.set(t.driver.id, (tripsPerDriver.get(t.driver.id) ?? 0) + 1);
        });

        const availableDrivers = drivers.filter(
          (d) =>
            d.isActive &&
            d.status !== 'On Leave' &&
            d.status !== 'Inactive' &&
            (tripsPerDriver.get(d.id) ?? 0) < (d.maxConcurrentTrips ?? 1),
        ).length;

        const busyVehicles = new Set<number>();
        openTrips.forEach((t) => { if (t.vehicle) busyVehicles.add(t.vehicle.id); });

        const availableVehicles = vehicles.filter(
          (v) => v.status !== 'Under Maintenance' && v.status !== 'Retired' && !busyVehicles.has(v.id),
        ).length;

        const pairs = new Set<string>();
        openTrips.forEach((t) => {
          if (t.driver && t.vehicle) pairs.add(`${t.driver.id}:${t.vehicle.id}`);
        });

        return {
          counters: {
            ongoingTrips: ongoingTrips.length,
            pendingRequests: pendingRequests.length,
            completedTrips: trips.filter((t) => t.tripStatus === TRIP_STATUS.Completed).length,
            availableDrivers,
            availableVehicles,
            allocatedPairs: pairs.size,
          },
          ongoingTrips,
          pendingRequests,
          drivers,
          vehicles,
          escalated: pendingRequests
            .filter((r) => r.status === REQUEST_STATUS.AwaitingCoordinator)
            .map((r) => r.id),
          loadedAt: new Date(),
        };
      },
    );
  }

  public async getOffersForRequest(requestItemId: number): Promise<IRideOffer[]> {
    return guard(
      'The offer history for this request could not be loaded.',
      'Refresh and open the request again. You can still assign manually without it.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.rideOffers)
          .items.select(...SELECT_OFFER)
          .expand(...EXPAND_OFFER)
          .filter(`RequestId eq ${requestItemId}`)
          .orderBy('OfferedOn', true)
          .top(200)();
        return rows.map(toOffer);
      },
    );
  }

  public async getTripsInRange(
    propertyIds: number[],
    from: Date,
    to: Date,
  ): Promise<ITripAssignment[]> {
    return guard(
      'Trips for that period could not be loaded.',
      'Narrow the date range and try again.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.tripAssignments)
          .items.select(...SELECT_TRIP)
          .expand(...EXPAND_TRIP)
          .filter(`AcceptedOn ge '${toIso(from)}' and AcceptedOn lt '${toIso(to)}'`)
          .orderBy('AcceptedOn', false)
          .top(2000)();
        return rows.map(toTrip);
      },
    );
  }

  // ===========================================================================
  // Internals
  // ===========================================================================

  private async getEffectiveRule(
    propertyId: number,
    requestType: RequestType,
  ): Promise<IEffectiveRule> {
    try {
      const rows: any[] = await getSP()
        .web.lists.getByTitle(LIST.allocationRules)
        .items.select(...SELECT_ALLOCATION_RULE)
        .expand(...EXPAND_ALLOCATION_RULE)
        .filter(
          `Property/Id eq ${propertyId} and RequestType eq '${requestType}' and IsActive eq 1`,
        )
        .top(1)();

      if (rows.length === 0) return { ...RULE_DEFAULTS };
      const rule: IAllocationRule = toAllocationRule(rows[0]);
      return {
        offerWindowMinutes: rule.offerWindowMinutes || RULE_DEFAULTS.offerWindowMinutes,
        maxConcurrentTripsPerDriver:
          rule.maxConcurrentTripsPerDriver || RULE_DEFAULTS.maxConcurrentTripsPerDriver,
        advanceReleaseHours: rule.advanceReleaseHours || RULE_DEFAULTS.advanceReleaseHours,
        mode: rule.mode,
      };
    } catch {
      // A missing or misnamed rules list must not stop a booking. The
      // documented defaults apply.
      console.warn('[FleetPortal] AllocationRules could not be read; using defaults.');
      return { ...RULE_DEFAULTS };
    }
  }

  private async readCandidateVehicles(typeId: number, propertyId: number): Promise<IVehicle[]> {
    const rows: any[] = await getSP()
      .web.lists.getByTitle(LIST.vehicleMaster)
      .items.select(...SELECT_VEHICLE)
      .expand(...EXPAND_VEHICLE)
      .filter(
        `VehicleType/Id eq ${typeId} and PropertyRef/Id eq ${propertyId} ` +
          `and VehStatus ne 'Retired' and VehStatus ne 'Under Maintenance'`,
      )
      .top(200)();
    return rows.map(toVehicle);
  }

  /**
   * Prefer this driver's default vehicle if it is free, otherwise the first
   * free vehicle of the right type.
   */
  private async pickFreeVehicle(
    typeId: number,
    propertyId: number,
    driverId: number,
  ): Promise<number | undefined> {
    const sp = getSP();
    const [vehicles, tripRows, blockRows] = await Promise.all([
      this.readCandidateVehicles(typeId, propertyId),
      sp.web.lists
        .getByTitle(LIST.tripAssignments)
        .items.select(...SELECT_TRIP)
        .expand(...EXPAND_TRIP)
        .filter(orFilter('TripStatus', OPEN_TRIP_STATUSES))
        .top(500)(),
      sp.web.lists
        .getByTitle(LIST.availabilityBlocks)
        .items.select(...SELECT_BLOCK)
        .expand(...EXPAND_BLOCK)
        .filter(`FromDateTime lt '${toIso(new Date(Date.now() + 86400000))}' ` +
                `and ToDateTime gt '${toIso(new Date())}'`)
        .top(500)(),
    ]);

    const free = freeVehicles(
      vehicles,
      (tripRows as any[]).map(toTrip),
      (blockRows as any[]).map(toAvailabilityBlock),
    );
    if (free.length === 0) return undefined;

    const preferred = free.filter((v) => v.defaultDriver?.id === driverId);
    return (preferred.length > 0 ? preferred[0] : free[0]).id;
  }

  private async setPendingOffers(requestItemId: number, to: string): Promise<void> {
    const sp = getSP();
    const rows: any[] = await sp.web.lists
      .getByTitle(LIST.rideOffers)
      .items.select('Id')
      .filter(`RequestId eq ${requestItemId} and ResponseStatus eq '${OFFER_STATUS.Pending}'`)
      .top(200)();
    if (rows.length === 0) return;

    const [batched, execute] = sp.batched();
    const list = batched.web.lists.getByTitle(LIST.rideOffers);
    rows.forEach((r) => {
      list.items.getById(r.Id).update({ ResponseStatus: to }).catch(() => undefined);
    });
    await execute();
  }

  private async setDriverStatus(driverId: number, status: string): Promise<void> {
    await getSP()
      .web.lists.getByTitle(LIST.driversMaster)
      .items.getById(driverId)
      .update({ Status: status }); // DriversMaster keeps the name Status
  }

  private async setVehicleStatus(vehicleId: number, status: string): Promise<void> {
    await getSP()
      .web.lists.getByTitle(LIST.vehicleMaster)
      .items.getById(vehicleId)
      .update({ VehStatus: status }); // VehicleMaster uses VehStatus
  }

  private async writeHistory(
    tripItemId: number,
    fromStatus: string,
    toStatus: string,
    comment?: string,
  ): Promise<void> {
    await getSP()
      .web.lists.getByTitle(LIST.tripStatusHistory)
      .items.add({
        Title: toStatus.toLowerCase(),
        TripId: tripItemId,
        FromStatus: fromStatus,
        ToStatus: toStatus,
        ChangedById: this.ctx.pageContext.legacyPageContext.userId,
        ChangedOn: toIso(new Date()),
        Comment: comment,
      });
  }
}
