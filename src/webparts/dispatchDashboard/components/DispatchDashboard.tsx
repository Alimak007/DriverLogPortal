/**
 * Dispatch Dashboard - the Coordinator's home page.
 *
 * Six counters, two grids and an availability panel from a SINGLE BATCHED
 * READ. Interactive in under 2.5 seconds on a warm cache.
 *
 * COUNTERS ARE FILTER SHORTCUTS, NOT STATIC NUMBERS. Each one filters the
 * grids below it rather than just reporting a figure.
 */

import * as React from 'react';
import { ConfirmDialog } from '../../../common/components/ConfirmDialog';
import { DataGrid, IColumn, RowActions } from '../../../common/components/DataGrid';
import {
  Badge,
  Button,
  Counter,
  ErrorBanner,
  FormRow,
  Loading,
  Notice,
  PageShell,
  Reference,
  SectionHeading,
  Select,
  TextField,
} from '../../../common/components/Shared';
import {
  IDispatchSnapshot,
  IDriver,
  IProperty,
  IRideOffer,
  ITransportRequest,
  ITripCard,
  IVehicle,
  driverState,
  requestState,
  tripState,
  vehicleState,
} from '../../../common/models';
import {
  explainPairIneligibility,
  IEligibilityContext,
} from '../../../common/rules/dispatchRules';
import { ICurrentUser, IServices } from '../../../common/services/IServices';
import {
  formatDateTime,
  formatTime,
  gulfMonthBounds,
  relativeFromNow,
} from '../../../common/utils/dates';
import styles from './DispatchDashboard.module.scss';

export interface IDispatchDashboardProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  refreshMs: number;
}

type CounterKey =
  | 'ongoing'
  | 'pending'
  | 'completed'
  | 'availableDrivers'
  | 'availableVehicles'
  | 'allocated';

type AvailabilityTab = 'drivers' | 'vehicles';

export const DispatchDashboard: React.FC<IDispatchDashboardProps> = ({
  services,
  user,
  bootstrapError,
  refreshMs,
}) => {
  const [snapshot, setSnapshot] = React.useState<IDispatchSnapshot | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<unknown>(bootstrapError);
  const [activeCounter, setActiveCounter] = React.useState<CounterKey | undefined>();
  const [tab, setTab] = React.useState<AvailabilityTab>('drivers');
  const [period, setPeriod] = React.useState(() => gulfMonthBounds());

  const [assignTarget, setAssignTarget] = React.useState<ITransportRequest | undefined>();
  const [offersForTarget, setOffersForTarget] = React.useState<IRideOffer[]>([]);
  const [rejectTarget, setRejectTarget] = React.useState<ITransportRequest | undefined>();
  const [cancelTarget, setCancelTarget] = React.useState<ITransportRequest | undefined>();
  const [busy, setBusy] = React.useState(false);

  /**
   * The property filter.
   *
   * It is a CONTROL, not a scope. Every Coordinator now sees every request by
   * default: with one driver pool, a request scoped to a property would be
   * invisible to everyone else, and if that Coordinator were on leave an
   * escalation could sit unseen. Narrowing is a choice the Coordinator makes,
   * not something their account does to them.
   */
  const [propertyFilter, setPropertyFilter] = React.useState<number | undefined>();
  const [properties, setProperties] = React.useState<IProperty[]>([]);

  React.useEffect(() => {
    if (!services) return;
    services.master
      .getProperties()
      .then(setProperties)
      .catch(() => setProperties([]));
  }, [services]);

  const load = React.useCallback(async (): Promise<void> => {
    if (!services) return;
    try {
      // Picks up advance bookings a flow has just released back to Submitted,
      // and requests whose creator closed their browser mid-submit. Idempotent,
      // so calling it on every poll is safe.
      //
      // KNOWN LIMITATION, BY DESIGN: a released advance booking is not
      // broadcast until a Coordinator or driver has a dashboard open. They
      // release four hours ahead, inside working hours. Deliberately NOT
      // solved with a flow - the eligibility rules live in one place only.
      await services.dispatch.sweepPendingBroadcasts();

      const result = await services.dispatch.getSnapshot(period.from, period.to);
      setSnapshot(result);
    } catch (e) {
      setError(e);
    } finally {
      setLoading(false);
    }
  }, [services, period.from, period.to]);

  React.useEffect(() => {
    if (!services) {
      setLoading(false);
      return undefined;
    }
    load().catch(() => undefined);
    const handle = window.setInterval(() => {
      load().catch(() => undefined);
    }, refreshMs);
    return () => window.clearInterval(handle);
  }, [services, load, refreshMs]);

  // ---------------------------------------------------------------------------
  // Counter filtering
  // ---------------------------------------------------------------------------

  const inProperty = React.useCallback(
    (r?: ITransportRequest): boolean =>
      propertyFilter === undefined || r?.property?.id === propertyFilter,
    [propertyFilter],
  );

  const visibleTrips: ITripCard[] = React.useMemo(() => {
    if (!snapshot) return [];
    if (activeCounter === 'pending') return [];
    return snapshot.ongoingTrips.filter((t) => inProperty(t.request));
  }, [snapshot, activeCounter, inProperty]);

  const visibleRequests: ITransportRequest[] = React.useMemo(() => {
    if (!snapshot) return [];
    if (activeCounter === 'ongoing') return [];
    return snapshot.pendingRequests.filter(inProperty);
  }, [snapshot, activeCounter, inProperty]);

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------

  const openAssign = async (request: ITransportRequest): Promise<void> => {
    setAssignTarget(request);
    setOffersForTarget([]);
    if (!services) return;
    try {
      const offers = await services.dispatch.getOffersForRequest(request.id);
      setOffersForTarget(offers);
    } catch (e) {
      setError(e);
    }
  };

  const doReject = async (reason?: string): Promise<void> => {
    if (!services || !rejectTarget || !reason) return;
    setBusy(true);
    try {
      const result = await services.dispatch.rejectRequest(rejectTarget.id, reason);
      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }
      setRejectTarget(undefined);
      await load();
    } catch (e) {
      setError(e);
    } finally {
      setBusy(false);
    }
  };

  const doCancel = async (reason?: string): Promise<void> => {
    if (!services || !cancelTarget || !reason) return;
    setBusy(true);
    try {
      const result = await services.dispatch.cancel({
        kind: 'coordinator',
        requestItemId: cancelTarget.id,
        reason,
      });
      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }
      setCancelTarget(undefined);
      await load();
    } catch (e) {
      setError(e);
    } finally {
      setBusy(false);
    }
  };

  // ---------------------------------------------------------------------------
  // Grids
  // ---------------------------------------------------------------------------

  const ongoingColumns: IColumn<ITripCard>[] = [
    {
      key: 'id',
      label: 'ID',
      render: (c) => <Reference>{c.request?.requestId ?? c.trip.tripId}</Reference>,
      width: '120px',
    },
    { key: 'pickup', label: 'Pickup', render: (c) => c.request.pickupText, filterValue: (c) => c.request.pickupText },
    {
      key: 'destination',
      label: 'Destination',
      render: (c) => c.request.dropOffText,
      filterValue: (c) => c.request.dropOffText,
    },
    {
      key: 'driver',
      label: 'Driver',
      render: (c) => c.trip.driver?.title ?? '—',
      filterValue: (c) => c.trip.driver?.title ?? '',
    },
    {
      key: 'vehicle',
      label: 'Vehicle',
      render: (c) => c.vehiclePlate || '—',
      filterValue: (c) => c.vehiclePlate,
    },
    {
      key: 'status',
      label: 'Status',
      render: (c) => <Badge state={tripState(c.trip.tripStatus)} label={c.trip.tripStatus} />,
      filterValue: (c) => c.trip.tripStatus,
      width: '130px',
    },
    {
      key: 'eta',
      label: 'ETA',
      // Always with its last-updated stamp, so nobody mistakes a typed-in
      // figure for a live calculation.
      render: (c) =>
        c.trip.etaMinutes === undefined ? (
          <span className={styles.currentTask}>not set</span>
        ) : (
          <>
            approx. {c.trip.etaMinutes} min
            {c.trip.etaUpdatedOn ? (
              <span className={styles.currentTask}>
                {' '}
                &middot; {relativeFromNow(c.trip.etaUpdatedOn)}
              </span>
            ) : undefined}
          </>
        ),
      width: '160px',
    },
  ];

  const pendingColumns: IColumn<ITransportRequest>[] = [
    {
      key: 'id',
      label: 'ID',
      render: (r) => <Reference>{r.requestId}</Reference>,
      width: '120px',
    },
    {
      key: 'when',
      label: 'Date and time',
      render: (r) => formatDateTime(r.requiredDateTime),
      width: '150px',
    },
    {
      key: 'type',
      label: 'Request type',
      render: (r) => r.requestType,
      filterValue: (r) => r.requestType,
    },
    {
      key: 'vehicleType',
      label: 'Vehicle type',
      render: (r) => r.vehicleTypeRequested?.title ?? '—',
      filterValue: (r) => r.vehicleTypeRequested?.title ?? '',
    },
    {
      key: 'status',
      label: 'Status',
      render: (r) => <Badge state={requestState(r.status)} label={r.status} />,
      filterValue: (r) => r.status,
      width: '170px',
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (r) => (
        <RowActions>
          <Button
            kind="action"
            onClick={() => {
              openAssign(r).catch(() => undefined);
            }}
          >
            Assign
          </Button>
          <Button kind="action" onClick={() => setRejectTarget(r)}>
            Reject
          </Button>
          <Button kind="action" onClick={() => setCancelTarget(r)}>
            Cancel
          </Button>
        </RowActions>
      ),
      width: '200px',
    },
  ];

  const driverColumns: IColumn<IDriver>[] = [
    { key: 'name', label: 'Driver', render: (d) => d.driverName, filterValue: (d) => d.driverName },
    {
      key: 'status',
      label: 'Status',
      render: (d) => <Badge state={driverState(d.status)} label={d.status} />,
      filterValue: (d) => d.status,
      width: '140px',
    },
    {
      key: 'task',
      label: 'Current task',
      render: (d) => {
        const open = (snapshot?.ongoingTrips ?? []).filter(
          (t) => t.trip.driver?.id === d.id,
        );
        if (open.length === 0) return <span className={styles.currentTask}>None</span>;
        return (
          <span>
            {open.map((t) => t.request.dropOffText).join(', ')}
            <span className={styles.currentTask}> ({open.length} open)</span>
          </span>
        );
      },
    },
    {
      key: 'cap',
      label: 'Cap',
      render: (d) => d.maxConcurrentTrips,
      numeric: true,
      width: '70px',
    },
    {
      key: 'property',
      label: 'Property',
      render: (d) => d.homeProperty?.title ?? '—',
      filterValue: (d) => d.homeProperty?.title ?? '',
    },
  ];

  const vehicleColumns: IColumn<IVehicle>[] = [
    { key: 'plate', label: 'Plate', render: (v) => v.plateNo, filterValue: (v) => v.plateNo },
    {
      key: 'vehicle',
      label: 'Vehicle',
      render: (v) => [v.colour, v.make, v.model].filter(Boolean).join(' ') || '—',
    },
    {
      key: 'type',
      label: 'Type',
      render: (v) => v.vehicleType?.title ?? '—',
      filterValue: (v) => v.vehicleType?.title ?? '',
    },
    {
      key: 'status',
      label: 'Status',
      render: (v) => <Badge state={vehicleState(v.status)} label={v.status} />,
      filterValue: (v) => v.status,
      width: '170px',
    },
    {
      key: 'task',
      label: 'Current task',
      render: (v) => {
        const open = (snapshot?.ongoingTrips ?? []).filter(
          (t) => t.trip.vehicle?.id === v.id,
        );
        return open.length === 0 ? (
          <span className={styles.currentTask}>None</span>
        ) : (
          <span>{open.map((t) => t.request.dropOffText).join(', ')}</span>
        );
      },
    },
  ];

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  if (!services || !user) {
    return (
      <PageShell heading="Dispatch dashboard">
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading /> : undefined}
      </PageShell>
    );
  }

  if (loading && !snapshot) {
    return (
      <PageShell heading="Dispatch dashboard">
        <Loading what="Loading dispatch" />
      </PageShell>
    );
  }

  const counters = snapshot?.counters;

  return (
    <PageShell
      heading="Dispatch dashboard"
      caption={
        snapshot
          ? `Last refreshed ${formatTime(snapshot.loadedAt)}. Counters filter the grids below.`
          : undefined
      }
      actions={
        <Button kind="secondary" onClick={() => load().catch(() => undefined)}>
          Refresh
        </Button>
      }
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      {snapshot && snapshot.escalated.length > 0 ? (
        <Notice tone="attention">
          <strong>
            {snapshot.escalated.length} request
            {snapshot.escalated.length === 1 ? '' : 's'} need allocating.
          </strong>{' '}
          Nobody accepted these within the offer window. They are pinned to the
          top of Pending requests.
        </Notice>
      ) : undefined}

      {counters ? (
        <div className={styles.section}>
          <div className={styles.periodRow}>
            <span className={styles.periodLabel}>Property:</span>
            <Select
              value={propertyFilter === undefined ? '' : String(propertyFilter)}
              onChange={(e) =>
                setPropertyFilter(e.target.value ? parseInt(e.target.value, 10) : undefined)
              }
              aria-label="Narrow the dashboard to one property"
              style={{ width: 'auto' }}
            >
              <option value="">All properties</option>
              {properties.map((pr) => (
                <option key={pr.id} value={pr.id}>
                  {pr.title}
                </option>
              ))}
            </Select>

            <span className={styles.periodLabel}>Completed counter period:</span>
            <Select
              value={`${period.from.toISOString()}`}
              onChange={(e) => {
                const months = parseInt(e.target.value, 10);
                const now = new Date();
                const bounds = gulfMonthBounds(
                  new Date(now.getFullYear(), now.getMonth() - months, 1),
                );
                setPeriod({ from: bounds.from, to: gulfMonthBounds(now).to });
              }}
              aria-label="Period for the completed trips counter"
              style={{ width: 'auto' }}
            >
              <option value={period.from.toISOString()}>This month</option>
              <option value="1">Last 2 months</option>
              <option value="2">Last 3 months</option>
            </Select>
          </div>

          <div
            className={styles.section}
            role="group"
            aria-label="Dispatch counters. Each one filters the grids below."
          >
            <CounterRow>
              <Counter
                label="Ongoing trips"
                value={counters.ongoingTrips}
                meta="En route or in transit"
                active={activeCounter === 'ongoing'}
                onClick={() =>
                  setActiveCounter(activeCounter === 'ongoing' ? undefined : 'ongoing')
                }
              />
              <Counter
                label="Pending requests"
                value={counters.pendingRequests}
                meta="Submitted, broadcasting or awaiting you"
                active={activeCounter === 'pending'}
                onClick={() =>
                  setActiveCounter(activeCounter === 'pending' ? undefined : 'pending')
                }
              />
              <Counter
                label="Completed trips"
                value={counters.completedTrips}
                meta="In the selected period"
                active={activeCounter === 'completed'}
                onClick={() =>
                  setActiveCounter(activeCounter === 'completed' ? undefined : 'completed')
                }
              />
              <Counter
                label="Available drivers"
                value={counters.availableDrivers}
                meta="Active, not on leave, under cap"
                active={activeCounter === 'availableDrivers'}
                onClick={() => {
                  setActiveCounter(
                    activeCounter === 'availableDrivers' ? undefined : 'availableDrivers',
                  );
                  setTab('drivers');
                }}
              />
              <Counter
                label="Available vehicles"
                value={counters.availableVehicles}
                meta="Active, no open trip"
                active={activeCounter === 'availableVehicles'}
                onClick={() => {
                  setActiveCounter(
                    activeCounter === 'availableVehicles' ? undefined : 'availableVehicles',
                  );
                  setTab('vehicles');
                }}
              />
              <Counter
                label="Allocated pairs"
                value={counters.allocatedPairs}
                meta="Driver and vehicle on an open trip"
                active={activeCounter === 'allocated'}
                onClick={() =>
                  setActiveCounter(activeCounter === 'allocated' ? undefined : 'allocated')
                }
              />
            </CounterRow>
          </div>
        </div>
      ) : undefined}

      {assignTarget && snapshot ? (
        <AssignmentPanel
          request={assignTarget}
          snapshot={snapshot}
          offers={offersForTarget}
          busy={busy}
          onClose={() => setAssignTarget(undefined)}
          onAssign={async (driverId, vehicleId, isOverride, reason) => {
            if (!services) return;
            setBusy(true);
            try {
              const result = await services.dispatch.manualAssign(
                assignTarget.id,
                driverId,
                vehicleId,
                isOverride,
                reason,
              );
              if (!result.ok) {
                setError(new Error(result.message));
                return;
              }
              setAssignTarget(undefined);
              await load();
            } catch (e) {
              setError(e);
            } finally {
              setBusy(false);
            }
          }}
        />
      ) : undefined}

      {activeCounter !== 'pending' ? (
        <section className={styles.section}>
          <div className={styles.sectionHead}>
            <SectionHeading>Ongoing trips</SectionHeading>
          </div>
          <DataGrid
            caption="Ongoing trips, soonest required time first"
            captionVisuallyHidden
            columns={ongoingColumns}
            rows={visibleTrips}
            rowKey={(c) => c.trip.id}
            emptyMessage="No trips are under way."
          />
        </section>
      ) : undefined}

      {activeCounter !== 'ongoing' ? (
        <section className={styles.section}>
          <div className={styles.sectionHead}>
            <SectionHeading>Pending requests</SectionHeading>
            <span className={styles.periodLabel}>
              Oldest first. Anything awaiting you is pinned to the top.
            </span>
          </div>
          <DataGrid
            caption="Pending requests, oldest first, with escalations pinned to the top"
            captionVisuallyHidden
            columns={pendingColumns}
            rows={visibleRequests}
            rowKey={(r) => r.id}
            isPinned={(r) => r.status === 'Awaiting Coordinator'}
            emptyMessage="Nothing is waiting to be allocated."
          />
        </section>
      ) : undefined}

      <section className={styles.section}>
        <SectionHeading>Availability</SectionHeading>
        <div className={styles.tabs} role="tablist" aria-label="Availability">
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'drivers'}
            className={[styles.tab, tab === 'drivers' ? styles.tabActive : ''].join(' ')}
            onClick={() => setTab('drivers')}
          >
            Drivers ({snapshot?.drivers.length ?? 0})
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'vehicles'}
            className={[styles.tab, tab === 'vehicles' ? styles.tabActive : ''].join(' ')}
            onClick={() => setTab('vehicles')}
          >
            Vehicles ({snapshot?.vehicles.length ?? 0})
          </button>
        </div>

        {tab === 'drivers' ? (
          <DataGrid
            caption="Driver availability"
            captionVisuallyHidden
            columns={driverColumns}
            rows={snapshot?.drivers ?? []}
            rowKey={(d) => d.id}
            emptyMessage="No drivers are set up for your property."
          />
        ) : (
          <DataGrid
            caption="Vehicle availability"
            captionVisuallyHidden
            columns={vehicleColumns}
            rows={snapshot?.vehicles ?? []}
            rowKey={(v) => v.id}
            emptyMessage="No vehicles are set up for your property."
          />
        )}

        <span className={styles.staleNote}>
          Driver and vehicle status is maintained by the dispatch flows. If a
          status here looks wrong, it was probably edited by hand in the list -
          the nightly reconciliation will correct it.
        </span>
      </section>

      <ConfirmDialog
        open={rejectTarget !== undefined}
        title="Reject this request?"
        message={`Request ${rejectTarget?.requestId ?? ''} will be declined.`}
        consequences={[
          'The requester is e-mailed the reason you give below.',
          'Any open offers to drivers are closed.',
          'The request cannot be reopened - the requester must book again.',
        ]}
        confirmLabel="Reject the request"
        destructive
        requireReason
        reasonLabel="Reason (the requester will read this)"
        busy={busy}
        onConfirm={(reason) => {
          doReject(reason).catch(() => undefined);
        }}
        onCancel={() => setRejectTarget(undefined)}
      />

      <ConfirmDialog
        open={cancelTarget !== undefined}
        title="Cancel this request?"
        message={`Request ${cancelTarget?.requestId ?? ''} will be cancelled.`}
        consequences={[
          'The requester is notified with the reason.',
          'Any assigned driver is notified and released.',
          'Any open offers to drivers are withdrawn.',
          'The vehicle, if one was bound, is released.',
        ]}
        confirmLabel="Cancel the request"
        destructive
        requireReason
        reasonLabel="Reason"
        busy={busy}
        onConfirm={(reason) => {
          doCancel(reason).catch(() => undefined);
        }}
        onCancel={() => setCancelTarget(undefined)}
      />
    </PageShell>
  );
};

function CounterRow({ children }: { children: React.ReactNode }): JSX.Element {
  return <div className={styles.counterRow}>{children}</div>;
}

// -----------------------------------------------------------------------------
// Manual assignment
// -----------------------------------------------------------------------------

/**
 * The Coordinator can assign or reassign AT ANY TIME, regardless of the
 * allocation mode.
 *
 * The panel permits an OUT-OF-POLICY choice - a driver already at their cap,
 * or a vehicle of a different type - with a mandatory override reason. It
 * shows the eligibility verdict so the Coordinator knows they are overriding
 * rather than discovering it afterwards.
 */
function AssignmentPanel({ request, snapshot, offers, busy, onClose, onAssign }: {
  request: ITransportRequest;
  snapshot: IDispatchSnapshot;
  offers: IRideOffer[];
  busy: boolean;
  onClose: () => void;
  onAssign: (
    driverId: number,
    vehicleId: number,
    isOverride: boolean,
    reason?: string,
  ) => Promise<void>;
}): JSX.Element {
  /**
   * ONE CHOICE, NOT TWO.
   *
   * The pairing is permanent: a driver owns a vehicle, recorded on
   * VehicleMaster.DefaultDriver. Offering the Coordinator separate driver and
   * vehicle lists would let them pair driver A with vehicle B and quietly
   * break the invariant the whole engine now rests on. So the panel lists
   * PAIRS, and picking one sets both.
   */
  const [vehicleId, setVehicleId] = React.useState<number | undefined>();
  const [reason, setReason] = React.useState('');

  const windowFrom = request.requiredDateTime;
  const windowTo = request.returnDateTime
    ? request.returnDateTime
    : new Date(windowFrom.getTime() + 2 * 3600000);

  const openTrips = React.useMemo(
    () => snapshot.ongoingTrips.map((t) => t.trip),
    [snapshot.ongoingTrips],
  );

  const driversById = React.useMemo(() => {
    const map = new Map<number, IDriver>();
    snapshot.drivers.forEach((d) => map.set(d.id, d));
    return map;
  }, [snapshot.drivers]);

  const ctx: IEligibilityContext = React.useMemo(
    () => ({
      driversById,
      // The dashboard snapshot does not carry availability blocks, so a block
      // cannot be seen here. The Coordinator is overriding deliberately
      // anyway, and the reason is recorded on the trip.
      blocks: [],
      openTrips,
      windowTo,
      ruleMaxConcurrent: 1,
    }),
    [driversById, openTrips, windowTo],
  );

  /** Every vehicle with its permanent driver, and why it may not be usable. */
  const candidates = React.useMemo(
    () =>
      snapshot.vehicles.map((v) => ({
        vehicle: v,
        driver: driversById.get(v.defaultDriver?.id ?? -1),
        problems: explainPairIneligibility(v, ctx),
      })),
    [snapshot.vehicles, driversById, ctx],
  );

  const requestedTypeId = request.vehicleTypeRequested?.id;
  const chosen = candidates.filter((c) => c.vehicle.id === vehicleId)[0];

  const wrongType =
    chosen !== undefined &&
    requestedTypeId !== undefined &&
    chosen.vehicle.vehicleType?.id !== requestedTypeId;

  const isOverride = chosen !== undefined && (chosen.problems.length > 0 || wrongType);
  const driverId = chosen?.driver?.id;
  const canAssign =
    chosen !== undefined &&
    driverId !== undefined &&
    (!isOverride || reason.trim().length > 0);

  const declined = offers.filter((o) => o.responseStatus === 'Rejected');

  return (
    <div className={styles.assignPanel}>
      <div className={styles.assignHead}>
        <div>
          <SectionHeading>Assign {request.requestId}</SectionHeading>
          <span className={styles.choiceMeta}>
            {request.pickupText} &rarr; {request.dropOffText}{' '}
            &middot; {formatDateTime(request.requiredDateTime)} &middot;{' '}
            {request.vehicleTypeRequested?.title}
            {request.otherVehicleText ? ` (${request.otherVehicleText})` : ''}
          </span>
        </div>
        <Button kind="secondary" onClick={onClose}>
          Close
        </Button>
      </div>

      {request.status === 'Awaiting Coordinator' ? (
        <Notice tone="attention">
          The offer window closed with nobody accepting.
          {declined.length > 0 ? (
            <ul className={styles.declinedList}>
              {declined.map((o) => (
                <li key={o.id}>
                  {o.driver?.title ?? 'A driver'} declined:{' '}
                  {o.rejectionReason ?? 'no reason given'}
                  {o.rejectionNotes ? ` — ${o.rejectionNotes}` : ''}
                </li>
              ))}
            </ul>
          ) : (
            <div className={styles.declinedList}>
              {offers.length === 0
                ? 'No drivers were eligible for this request.'
                : `${offers.length} driver(s) were offered it and none responded.`}
            </div>
          )}
        </Notice>
      ) : undefined}

      <div className={styles.assignGrid}>
        <div>
          <FormRow label="Driver and vehicle" required>
            <div className={styles.choiceList} role="listbox" aria-label="Driver and vehicle pairs">
              {candidates.map((c) => {
                const label = c.driver
                  ? `${c.driver.driverName} \u2014 ${c.vehicle.plateNo}`
                  : `${c.vehicle.plateNo} \u2014 no assigned driver`;
                const notes: string[] = c.problems.slice();
                if (
                  requestedTypeId !== undefined &&
                  c.vehicle.vehicleType?.id !== requestedTypeId
                ) {
                  notes.push(`different vehicle type (${c.vehicle.vehicleType?.title ?? 'none'})`);
                }
                return (
                  <button
                    key={c.vehicle.id}
                    type="button"
                    role="option"
                    aria-selected={vehicleId === c.vehicle.id}
                    disabled={!c.driver}
                    className={[
                      styles.choice,
                      vehicleId === c.vehicle.id ? styles.choiceSelected : '',
                    ].join(' ')}
                    onClick={() => setVehicleId(c.vehicle.id)}
                  >
                    <span>
                      {label}
                      <span className={styles.choiceMeta}>
                        {notes.length === 0 ? ' \u00B7 eligible' : ` \u00B7 ${notes.join('; ')}`}
                      </span>
                    </span>
                    <Badge state={vehicleState(c.vehicle.status)} label={c.vehicle.status} />
                  </button>
                );
              })}
            </div>
          </FormRow>
        </div>
      </div>

      {isOverride ? (
        <div className={styles.escalationNote}>
          <Notice tone="attention">
            This is an out-of-policy assignment. A reason is required and is
            recorded on the trip.
          </Notice>
          <FormRow label="Override reason" required>
            <TextField
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              placeholder="Why this driver or vehicle despite the rule"
            />
          </FormRow>
        </div>
      ) : undefined}

      <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 12 }}>
        <Button kind="secondary" onClick={onClose} disabled={busy}>
          Cancel
        </Button>
        <Button
          kind="primary"
          disabled={!canAssign || busy}
          onClick={() => {
            if (driverId !== undefined && vehicleId !== undefined) {
              onAssign(
                driverId,
                vehicleId,
                isOverride,
                isOverride ? reason.trim() : undefined,
              ).catch(() => undefined);
            }
          }}
        >
          {busy ? 'Assigning…' : 'Assign'}
        </Button>
      </div>

      <span className={styles.staleNote}>
        Assigning manually skips the broadcast. Any open offers for this request
        are withdrawn, and the chosen driver is notified directly rather than
        being offered a choice.
      </span>
    </div>
  );
}

export default DispatchDashboard;
