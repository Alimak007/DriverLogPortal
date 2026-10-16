/**
 * Dispatch Dashboard - the Coordinator's home page.
 *
 * Six counters, two grids and an availability panel from a SINGLE BATCHED
 * READ. Interactive in under 2.5 seconds on a warm cache.
 *
 * COUNTERS ARE FILTER SHORTCUTS, NOT STATIC NUMBERS. Each one filters the
 * grids below it rather than just reporting a figure. The side menu is the
 * same set of shortcuts under the names a Coordinator looks for.
 *
 * Laid out as the client's "Coordinator - Fleet Management Dashboard" design:
 * a title tab, a side menu, tinted counter cards, two compact tables, the
 * availability panel down the right, and the manual allocation as six
 * numbered steps along the bottom.
 */

import * as React from 'react';
import { ConfirmDialog } from '../../../common/components/ConfirmDialog';
import { DataGrid, IColumn } from '../../../common/components/DataGrid';
import { Icon, IconName } from '../../../common/components/Icons';
import {
  Button,
  ErrorBanner,
  FormRow,
  Loading,
  PageShell,
  Select,
  TextField,
} from '../../../common/components/Shared';
import {
  IDispatchSnapshot,
  IDriver,
  IRideOffer,
  ITransportRequest,
  ITripCard,
  IVehicle,
  RequestType,
  SemanticState,
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
  formatDateLong,
  formatTime,
  gulfMonthBounds,
} from '../../../common/utils/dates';
import { DispatchCalendar } from './DispatchCalendar';
import styles from './DispatchDashboard.module.scss';

export interface IDispatchDashboardProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  refreshMs: number;
  /** The Fleet Reports page, for the side menu. Left out when not given. */
  reportsUrl?: string;
}

type CounterKey =
  | 'ongoing'
  | 'pending'
  | 'completed'
  | 'availableDrivers'
  | 'availableVehicles'
  | 'allocated';

type AvailabilityTab = 'drivers' | 'vehicles';
type Layout = 'list' | 'calendar';

/** Remembers List or Calendar per browser. A convenience only - may be absent. */
const LAYOUT_KEY = 'dispatchDashboard.layout';

function readLayout(): Layout {
  try {
    return window.localStorage.getItem(LAYOUT_KEY) === 'calendar' ? 'calendar' : 'list';
  } catch {
    return 'list';
  }
}

const TITLE = 'Coordinator – Fleet Management Dashboard';

/** The request types, as short as the design writes them. Display only. */
const SHORT_TYPE: Record<RequestType, string> = {
  'Client or Client Guest': 'Client',
  'Staff': 'Staff',
  'Items or Documents': 'Items',
};

/** "25 Oct 2026 10:00" - the design's date and time. Gulf time, as everywhere. */
function when(d?: Date): string {
  return d ? `${formatDateLong(d)} ${formatTime(d)}` : '';
}

function vehicleName(v?: IVehicle): string {
  if (!v) return '';
  const name = [v.make, v.model].filter(Boolean).join(' ');
  return (name || v.plateNo) + (v.colour ? ` (${v.colour})` : '');
}

// -----------------------------------------------------------------------------
// Small parts
// -----------------------------------------------------------------------------

const PILL_CLASS: Record<SemanticState, string> = {
  healthy: styles.pillHealthy,
  attention: styles.pillAttention,
  progress: styles.pillProgress,
  exception: styles.pillException,
  neutral: styles.pillNeutral,
};

/** The two statuses the design draws as SOLID pills: the trip is moving or done. */
const SOLID: string[] = ['In Transit', 'Completed'];

/** A status, as a pill. Always carries its word. */
function Pill({ state, label }: { state: SemanticState; label: string }): JSX.Element {
  return (
    <span
      className={[
        styles.pill,
        PILL_CLASS[state],
        SOLID.indexOf(label) !== -1 ? styles.pillSolid : '',
      ].join(' ')}
    >
      <span className={styles.pillDot} aria-hidden="true" />
      {label}
    </span>
  );
}

/** The round initials in place of a photograph. */
function Avatar({ name, large }: { name: string; large?: boolean }): JSX.Element {
  const initials = name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w.charAt(0).toUpperCase())
    .join('');
  return (
    <span className={[styles.avatar, large ? styles.avatarLarge : ''].join(' ')} aria-hidden="true">
      {initials || '?'}
    </span>
  );
}

/** Long map addresses are cut to one line; the whole text is in the tooltip. */
function Clip({ text }: { text?: string }): JSX.Element {
  return (
    <span className={styles.clip} title={text}>
      {text || '—'}
    </span>
  );
}

/** Closes a popup on Escape and on a click outside it. */
function useDismiss(
  open: boolean,
  close: () => void,
  inside: (React.RefObject<HTMLElement>)[],
): void {
  React.useEffect(() => {
    if (!open) return undefined;
    const onDown = (e: MouseEvent): void => {
      const t = e.target as Node;
      if (inside.some((r) => r.current?.contains(t))) return;
      close();
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') close();
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);
}

interface IMenuItem {
  text: string;
  icon: IconName;
  danger?: boolean;
  href?: string;
  onSelect?: () => void;
}

/**
 * The "..." menu on a row. It is drawn fixed to the window, not inside the
 * grid, so the grid's own scrolling cannot cut it off.
 */
function ActionMenu({ label, items }: { label: string; items: IMenuItem[] }): JSX.Element {
  const [open, setOpen] = React.useState(false);
  const [pos, setPos] = React.useState<{ top: number; left: number } | undefined>();
  const buttonRef = React.useRef<HTMLButtonElement>(null);
  const menuRef = React.useRef<HTMLDivElement>(null);

  useDismiss(open, () => setOpen(false), [buttonRef, menuRef]);

  React.useEffect(() => {
    if (!open) return undefined;
    const close = (): void => setOpen(false);
    window.addEventListener('scroll', close, true);
    window.addEventListener('resize', close);
    return () => {
      window.removeEventListener('scroll', close, true);
      window.removeEventListener('resize', close);
    };
  }, [open]);

  const show = (): void => {
    const r = buttonRef.current?.getBoundingClientRect();
    if (!r) return;
    setPos({ top: r.bottom + 4, left: Math.max(8, r.right - 200) });
    setOpen(true);
  };

  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        className={styles.dotsButton}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => (open ? setOpen(false) : show())}
      >
        <span aria-hidden="true">&bull;&bull;&bull;</span>
      </button>
      {open && pos ? (
        <div ref={menuRef} role="menu" className={styles.menu} style={{ top: pos.top, left: pos.left }}>
          {items.map((it) =>
            it.href ? (
              <a
                key={it.text}
                role="menuitem"
                className={styles.menuItem}
                href={it.href}
                onClick={() => setOpen(false)}
              >
                <Icon name={it.icon} size={16} />
                {it.text}
              </a>
            ) : (
              <button
                key={it.text}
                type="button"
                role="menuitem"
                className={[styles.menuItem, it.danger ? styles.menuItemDanger : ''].join(' ')}
                onClick={() => {
                  setOpen(false);
                  if (it.onSelect) it.onSelect();
                }}
              >
                <Icon name={it.icon} size={16} />
                {it.text}
              </button>
            ),
          )}
        </div>
      ) : undefined}
    </>
  );
}

/** A white card with a title row, as in the design. */
function Box({ title, action, className, children }: {
  title: string;
  action?: React.ReactNode;
  className?: string;
  children: React.ReactNode;
}): JSX.Element {
  return (
    <section className={[styles.box, className ?? ''].join(' ')} aria-label={title}>
      <div className={styles.boxHead}>
        <h2 className={styles.boxTitle}>{title}</h2>
        {action}
      </div>
      {children}
    </section>
  );
}

/** "View All ->", or "Back" while that view is open. */
function ViewAll({ open, onClick }: { open: boolean; onClick: () => void }): JSX.Element {
  return (
    <button type="button" className={styles.viewAll} onClick={onClick} aria-pressed={open}>
      {open ? 'Back' : 'View All'}
      <Icon name={open ? 'chevronLeft' : 'arrowRight'} size={15} />
    </button>
  );
}

/**
 * A pick list drawn as a card with a chevron, as the allocation steps show
 * it. It is a real list box: it opens on the card, and closes on a choice,
 * on Escape and on a click anywhere else.
 */
function Dropdown({ label, value, placeholder, disabled, options, onPick }: {
  label: string;
  value?: React.ReactNode;
  placeholder: string;
  disabled?: boolean;
  options: { key: number; content: React.ReactNode; selected: boolean; disabled?: boolean }[];
  onPick: (key: number) => void;
}): JSX.Element {
  const [open, setOpen] = React.useState(false);
  const wrapRef = React.useRef<HTMLDivElement>(null);
  useDismiss(open, () => setOpen(false), [wrapRef]);

  return (
    <div className={styles.dropdown} ref={wrapRef}>
      <button
        type="button"
        className={[styles.dropdownButton, open ? styles.dropdownOpen : ''].join(' ')}
        aria-haspopup="listbox"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((o) => !o)}
      >
        <span className={styles.srOnly}>{`${label}: `}</span>
        <span className={styles.dropdownValue}>
          {value ?? <span className={styles.placeholder}>{placeholder}</span>}
        </span>
        <span className={styles.dropdownChevron} aria-hidden="true">
          <Icon name="chevronRight" size={18} />
        </span>
      </button>
      {open ? (
        <ul className={styles.dropdownList} role="listbox" aria-label={label}>
          {options.length === 0 ? (
            <li className={styles.dropdownEmpty}>Nothing to choose from.</li>
          ) : (
            options.map((o) => (
              <li key={o.key} role="option" aria-selected={o.selected}>
                <button
                  type="button"
                  className={[styles.dropdownOption, o.selected ? styles.dropdownOptionSelected : ''].join(' ')}
                  disabled={o.disabled}
                  onClick={() => {
                    setOpen(false);
                    onPick(o.key);
                  }}
                >
                  {o.content}
                </button>
              </li>
            ))
          )}
        </ul>
      ) : undefined}
    </div>
  );
}

// -----------------------------------------------------------------------------
// The dashboard
// -----------------------------------------------------------------------------

interface ILastAllocated {
  requestItemId: number;
  requestId: string;
  driverName: string;
  vehicleName: string;
}

export const DispatchDashboard: React.FC<IDispatchDashboardProps> = ({
  services,
  user,
  bootstrapError,
  refreshMs,
  reportsUrl,
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
  const [lastAllocated, setLastAllocated] = React.useState<ILastAllocated | undefined>();

  const stepsRef = React.useRef<HTMLElement>(null);

  const [layout, setLayoutState] = React.useState<Layout>(readLayout);
  const setLayout = (next: Layout): void => {
    setLayoutState(next);
    try {
      window.localStorage.setItem(LAYOUT_KEY, next);
    } catch {
      // Storage blocked: the choice simply is not remembered.
    }
  };

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
      //
      // Housekeeping, so a failure here is logged and must not stop the
      // dashboard loading; the next poll tries the sweep again.
      try {
        await services.dispatch.sweepPendingBroadcasts();
      } catch (sweepError) {
        console.error('[FleetPortal] broadcast sweep failed', sweepError);
      }

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

  const visibleTrips: ITripCard[] = React.useMemo(() => {
    if (!snapshot) return [];
    if (activeCounter === 'pending') return [];
    return snapshot.ongoingTrips;
  }, [snapshot, activeCounter]);

  const visibleRequests: ITransportRequest[] = React.useMemo(() => {
    if (!snapshot) return [];
    if (activeCounter === 'ongoing') return [];
    return snapshot.pendingRequests;
  }, [snapshot, activeCounter]);

  const allPending: ITransportRequest[] = snapshot ? snapshot.pendingRequests : [];

  const toggle = (key: CounterKey | undefined): void =>
    setActiveCounter(activeCounter === key ? undefined : key);

  const vehiclesById = React.useMemo(() => {
    const map = new Map<number, IVehicle>();
    (snapshot?.vehicles ?? []).forEach((v) => map.set(v.id, v));
    return map;
  }, [snapshot]);

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------

  const openAssign = async (request: ITransportRequest): Promise<void> => {
    setAssignTarget(request);
    setOffersForTarget([]);
    window.setTimeout(() => {
      stepsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 50);
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

  const doAssign = async (
    driverId: number,
    vehicleId: number,
    isOverride: boolean,
    reason?: string,
  ): Promise<void> => {
    if (!services || !assignTarget) return;
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
      // Kept so steps 5 and 6 can show what was just done and follow the trip.
      setLastAllocated({
        requestItemId: assignTarget.id,
        requestId: assignTarget.requestId,
        driverName: snapshot?.drivers.filter((d) => d.id === driverId)[0]?.driverName ?? '',
        vehicleName: vehicleName(vehiclesById.get(vehicleId)),
      });
      setAssignTarget(undefined);
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
      render: (c) => <span className={styles.ref}>{c.request?.requestId ?? c.trip.tripId}</span>,
    },
    {
      key: 'pickup',
      label: 'Pickup',
      render: (c) => <Clip text={c.request.pickupText} />,
      filterValue: (c) => c.request.pickupText,
    },
    {
      key: 'destination',
      label: 'Destination',
      render: (c) => <Clip text={c.request.dropOffText} />,
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
      render: (c) => {
        const v = c.trip.vehicle ? vehiclesById.get(c.trip.vehicle.id) : undefined;
        const name = v ? [v.make, v.model].filter(Boolean).join(' ') : '';
        return <span title={c.vehiclePlate}>{name || c.vehiclePlate || '—'}</span>;
      },
      filterValue: (c) => c.vehiclePlate,
    },
    {
      key: 'status',
      label: 'Status',
      render: (c) => <Pill state={tripState(c.trip.tripStatus)} label={c.trip.tripStatus} />,
      filterValue: (c) => c.trip.tripStatus,
    },
    {
      key: 'eta',
      label: 'ETA',
      // A typed-in figure, so the time it was last set is in the tooltip.
      render: (c) =>
        c.trip.etaMinutes === undefined ? (
          <span className={styles.muted}>&ndash;</span>
        ) : (
          <span title={c.trip.etaUpdatedOn ? `Estimate set ${when(c.trip.etaUpdatedOn)}` : undefined}>
            {c.trip.etaMinutes} mins
          </span>
        ),
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (c) => (
        <ActionMenu
          label={`Actions for ${c.request.requestId}`}
          items={[
            { text: 'Cancel trip', icon: 'x', danger: true, onSelect: () => setCancelTarget(c.request) },
          ]}
        />
      ),
    },
  ];

  // The same trip columns, with the drop time in place of the ETA and no
  // actions: a completed trip is history.
  const completedColumns: IColumn<ITripCard>[] = [
    ...ongoingColumns.filter((c) => c.key !== 'eta' && c.key !== 'actions' && c.key !== 'status'),
    {
      key: 'completedOn',
      label: 'Completed On',
      render: (c) => (
        <span className={styles.nowrap}>
          {when(c.trip.actualDropTime ?? c.request.requiredDateTime)}
        </span>
      ),
    },
    {
      key: 'status',
      label: 'Status',
      render: (c) => <Pill state={tripState(c.trip.tripStatus)} label={c.trip.tripStatus} />,
    },
  ];

  const pendingColumns: IColumn<ITransportRequest>[] = [
    { key: 'id', label: 'ID', render: (r) => <span className={styles.ref}>{r.requestId}</span> },
    {
      key: 'when',
      label: 'Date & Time',
      render: (r) => <span className={styles.nowrap}>{when(r.requiredDateTime)}</span>,
    },
    {
      key: 'type',
      label: 'Request Type',
      render: (r) => SHORT_TYPE[r.requestType] ?? r.requestType,
      filterValue: (r) => SHORT_TYPE[r.requestType] ?? r.requestType,
    },
    {
      key: 'vehicleType',
      label: 'Vehicle Type',
      render: (r) => r.vehicleTypeRequested?.title ?? '—',
      filterValue: (r) => r.vehicleTypeRequested?.title ?? '',
    },
    {
      key: 'status',
      label: 'Status',
      render: (r) => <Pill state={requestState(r.status)} label={r.status} />,
      filterValue: (r) => r.status,
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (r) => (
        <ActionMenu
          label={`Actions for ${r.requestId}`}
          items={[
            {
              text: 'Allocate driver & vehicle',
              icon: 'car',
              onSelect: () => {
                openAssign(r).catch(() => undefined);
              },
            },
            { text: 'Reject request', icon: 'x', danger: true, onSelect: () => setRejectTarget(r) },
            { text: 'Cancel request', icon: 'x', danger: true, onSelect: () => setCancelTarget(r) },
          ]}
        />
      ),
    },
  ];

  const taskFor = (match: (t: ITripCard) => boolean): JSX.Element => {
    const open = (snapshot?.ongoingTrips ?? []).filter(match);
    if (open.length === 0) return <span className={styles.muted}>&ndash;</span>;
    return (
      <span title={open.map((t) => t.request.dropOffText).join(', ')}>
        {open.map((t) => t.request.requestId).join(', ')}
      </span>
    );
  };

  const driverColumns: IColumn<IDriver>[] = [
    {
      key: 'name',
      label: 'Driver',
      render: (d) => (
        <span className={styles.person}>
          <Avatar name={d.driverName} />
          <span className={styles.personName}>{d.driverName}</span>
        </span>
      ),
      filterValue: (d) => d.driverName,
    },
    {
      key: 'status',
      label: 'Status',
      render: (d) => <Pill state={driverState(d.status)} label={d.status} />,
      filterValue: (d) => d.status,
    },
    {
      key: 'task',
      label: 'Current Task',
      render: (d) => taskFor((t) => t.trip.driver?.id === d.id),
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (d) =>
        d.phone ? (
          <ActionMenu
            label={`Actions for ${d.driverName}`}
            items={[{ text: `Call ${d.phone}`, icon: 'phone', href: `tel:${d.phone}` }]}
          />
        ) : (
          <span className={styles.muted}>&ndash;</span>
        ),
    },
  ];

  const vehicleColumns: IColumn<IVehicle>[] = [
    {
      key: 'plate',
      label: 'Vehicle',
      render: (v) => (
        <span className={styles.person}>
          <span className={styles.vehicleIcon} aria-hidden="true">
            <Icon name="car" size={16} />
          </span>
          <span className={styles.personName}>
            {vehicleName(v)}
            <span className={styles.subLine}>{v.plateNo}</span>
          </span>
        </span>
      ),
      filterValue: (v) => v.plateNo,
    },
    {
      key: 'status',
      label: 'Status',
      render: (v) => <Pill state={vehicleState(v.status)} label={v.status} />,
      filterValue: (v) => v.status,
    },
    {
      key: 'task',
      label: 'Current Task',
      render: (v) => taskFor((t) => t.trip.vehicle?.id === v.id),
    },
  ];

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  if (!services || !user) {
    return (
      <PageShell heading={TITLE}>
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading /> : undefined}
      </PageShell>
    );
  }

  if (loading && !snapshot) {
    return (
      <PageShell heading={TITLE}>
        <Loading what="Loading dispatch" />
      </PageShell>
    );
  }

  const counters = snapshot?.counters;

  const showDrivers = (): void => {
    toggle('availableDrivers');
    setTab('drivers');
  };
  const showVehicles = (): void => {
    toggle('availableVehicles');
    setTab('vehicles');
  };

  const counterCards: {
    key: CounterKey;
    label: string;
    value: number;
    meta: string;
    icon: IconName;
    tone: string;
    onClick: () => void;
  }[] = counters
    ? [
        { key: 'ongoing', label: 'Ongoing Trips', value: counters.ongoingTrips, meta: 'En route or in transit', icon: 'car', tone: styles.toneBlue, onClick: () => toggle('ongoing') },
        { key: 'pending', label: 'Pending Requests', value: counters.pendingRequests, meta: 'Submitted, broadcasting or awaiting you', icon: 'clock', tone: styles.toneOrange, onClick: () => toggle('pending') },
        { key: 'completed', label: 'Completed Trips', value: counters.completedTrips, meta: 'In the selected period', icon: 'check', tone: styles.toneGreen, onClick: () => toggle('completed') },
        { key: 'availableDrivers', label: 'Available Drivers', value: counters.availableDrivers, meta: 'Active, not on leave, under cap', icon: 'user', tone: styles.tonePurple, onClick: showDrivers },
        { key: 'availableVehicles', label: 'Available Vehicles', value: counters.availableVehicles, meta: 'Active, no open trip', icon: 'car', tone: styles.toneTeal, onClick: showVehicles },
        { key: 'allocated', label: 'Allocated Vehicles/Drivers', value: counters.allocatedPairs, meta: 'Driver and vehicle on an open trip', icon: 'car', tone: styles.toneGrey, onClick: () => toggle('allocated') },
      ]
    : [];

  const ongoingFocused = activeCounter === 'ongoing';
  const pendingFocused = activeCounter === 'pending';
  const availFocused = activeCounter === 'availableDrivers' || activeCounter === 'availableVehicles';
  const completedFocused = activeCounter === 'completed';

  const menu: { text: string; icon: IconName; active: boolean; onClick?: () => void; href?: string }[] = [
    { text: 'Dashboard', icon: 'home', active: activeCounter === undefined, onClick: () => setActiveCounter(undefined) },
    { text: 'Booking Requests', icon: 'doc', active: pendingFocused, onClick: () => toggle('pending') },
    { text: 'Ongoing Trips', icon: 'navigation', active: ongoingFocused, onClick: () => toggle('ongoing') },
    { text: 'Completed Trips', icon: 'check', active: activeCounter === 'completed', onClick: () => toggle('completed') },
    { text: 'Drivers', icon: 'user', active: activeCounter === 'availableDrivers', onClick: showDrivers },
    { text: 'Vehicles', icon: 'car', active: activeCounter === 'availableVehicles', onClick: showVehicles },
  ];
  if (reportsUrl) menu.push({ text: 'Reports', icon: 'chart', active: false, href: reportsUrl });

  // The counters and the side menu filter the LIST grids, so in the calendar
  // they first bring the list back.
  const inList = (fn?: () => void): (() => void) => (): void => {
    if (layout !== 'list') setLayout('list');
    if (fn) fn();
  };

  return (
    <div className={styles.page}>
      <div className={styles.titleRow}>
        <header className={styles.titleBar}>
          <span className={styles.titleIcon} aria-hidden="true">
            <Icon name="user" size={20} />
          </span>
          <h1 className={styles.title}>{TITLE}</h1>
        </header>

        <div className={styles.toolbar}>
          <div className={styles.viewSwitch} role="group" aria-label="Dashboard layout">
            <button
              type="button"
              className={[styles.viewButton, layout === 'list' ? styles.viewButtonActive : ''].join(' ')}
              aria-pressed={layout === 'list'}
              onClick={() => setLayout('list')}
            >
              <Icon name="doc" size={15} />
              List View
            </button>
            <button
              type="button"
              className={[styles.viewButton, layout === 'calendar' ? styles.viewButtonActive : ''].join(' ')}
              aria-pressed={layout === 'calendar'}
              onClick={() => setLayout('calendar')}
            >
              <Icon name="calendar" size={15} />
              Calendar View
            </button>
          </div>
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
          >
            <option value={period.from.toISOString()}>This month</option>
            <option value="1">Last 2 months</option>
            <option value="2">Last 3 months</option>
          </Select>
          <span className={styles.refreshed}>
            {snapshot ? `Refreshed ${formatTime(snapshot.loadedAt)}` : ''}
          </span>
          <Button kind="secondary" onClick={() => load().catch(() => undefined)} aria-label="Refresh">
            <Icon name="refresh" size={16} />
          </Button>
        </div>
      </div>

      <div className={styles.frame}>
        <nav className={styles.sidebar} aria-label="Fleet management">
          <div className={styles.sidebarTitle}>
            <Icon name="car" size={26} />
            Fleet Management
          </div>
          <ul className={styles.menuList}>
            {menu.map((m) => (
              <li key={m.text}>
                {m.href ? (
                  <a className={styles.navItem} href={m.href}>
                    <Icon name={m.icon} size={19} />
                    {m.text}
                  </a>
                ) : (
                  <button
                    type="button"
                    className={[styles.navItem, m.active ? styles.navItemActive : ''].join(' ')}
                    aria-current={m.active ? 'true' : undefined}
                    onClick={inList(m.onClick)}
                  >
                    <Icon name={m.icon} size={19} />
                    {m.text}
                  </button>
                )}
              </li>
            ))}
          </ul>
        </nav>

        <div className={styles.main}>
          <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

          {snapshot && snapshot.escalated.length > 0 ? (
            <div className={styles.escalation} role="status">
              <Icon name="bell" size={16} />
              <span>
                <strong>
                  {snapshot.escalated.length} request
                  {snapshot.escalated.length === 1 ? '' : 's'} need allocating.
                </strong>{' '}
                Nobody accepted these within the offer window. They are pinned to the
                top of Pending Requests.
              </span>
            </div>
          ) : undefined}

          <div className={[styles.grid, layout === 'calendar' ? styles.gridCalendar : ''].join(' ')}>
            {counters ? (
              <div
                className={styles.counters}
                role="group"
                aria-label="Dispatch counters. Each one filters the grids below."
              >
                {counterCards.map((c) => (
                  <button
                    key={c.key}
                    type="button"
                    className={[styles.stat, c.tone, activeCounter === c.key ? styles.statActive : ''].join(' ')}
                    onClick={inList(c.onClick)}
                    aria-pressed={activeCounter === c.key}
                    aria-label={`${c.label}: ${c.value}. Show these.`}
                    title={c.meta}
                  >
                    <span className={styles.statIcon} aria-hidden="true">
                      <Icon name={c.icon} size={24} />
                    </span>
                    <span className={styles.statText}>
                      <span className={styles.statValue}>{c.value}</span>
                      <span className={styles.statLabel}>{c.label}</span>
                    </span>
                    <span className={styles.statChevron} aria-hidden="true">
                      <Icon name="chevronRight" size={16} />
                    </span>
                  </button>
                ))}
              </div>
            ) : undefined}

            {layout === 'calendar' && snapshot ? (
              <div className={styles.calendarArea}>
                <DispatchCalendar
                  services={services}
                  snapshot={snapshot}
                  onSaved={load}
                  onError={setError}
                />
              </div>
            ) : undefined}

            {layout === 'list' && completedFocused ? (
              <Box
                title="Completed Trips"
                className={[styles.ongoing, styles.wideLeft].join(' ')}
                action={<ViewAll open onClick={() => toggle('completed')} />}
              >
                <DataGrid
                  caption="Trips completed in the selected period, latest first"
                  captionVisuallyHidden
                  columns={completedColumns}
                  rows={snapshot?.completedTrips ?? []}
                  rowKey={(c) => c.trip.id}
                  emptyMessage="No trips were completed in this period."
                />
              </Box>
            ) : undefined}

            {layout === 'list' && !pendingFocused && !availFocused && !completedFocused ? (
              <Box
                title="Ongoing Trips"
                className={[styles.ongoing, ongoingFocused ? styles.wideLeft : ''].join(' ')}
                action={<ViewAll open={ongoingFocused} onClick={() => toggle('ongoing')} />}
              >
                <DataGrid
                  caption="Ongoing trips, soonest required time first"
                  captionVisuallyHidden
                  columns={ongoingColumns}
                  rows={visibleTrips}
                  rowKey={(c) => c.trip.id}
                  emptyMessage="No trips are under way."
                />
              </Box>
            ) : undefined}

            {layout === 'list' && !ongoingFocused && !availFocused && !completedFocused ? (
              <Box
                title="Pending Requests"
                className={[styles.pending, pendingFocused ? styles.wideLeft : ''].join(' ')}
                action={<ViewAll open={pendingFocused} onClick={() => toggle('pending')} />}
              >
                <DataGrid
                  caption="Pending requests, oldest first, with escalations pinned to the top"
                  captionVisuallyHidden
                  columns={pendingColumns}
                  rows={visibleRequests}
                  rowKey={(r) => r.id}
                  isPinned={(r) => r.status === 'Awaiting Coordinator'}
                  emptyMessage="Nothing is waiting to be allocated."
                />
              </Box>
            ) : undefined}

            {layout === 'list' ? (
            <Box
              title="Vehicles & Drivers Availability"
              className={[styles.availability, availFocused ? styles.wideAll : ''].join(' ')}
              action={
                <ViewAll
                  open={availFocused}
                  onClick={() => (tab === 'drivers' ? showDrivers() : showVehicles())}
                />
              }
            >
              <div className={styles.tabs} role="tablist" aria-label="Availability">
                <button
                  type="button"
                  role="tab"
                  aria-selected={tab === 'drivers'}
                  className={[styles.tab, tab === 'drivers' ? styles.tabActive : ''].join(' ')}
                  onClick={() => setTab('drivers')}
                >
                  Drivers
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={tab === 'vehicles'}
                  className={[styles.tab, tab === 'vehicles' ? styles.tabActive : ''].join(' ')}
                  onClick={() => setTab('vehicles')}
                >
                  Vehicles
                </button>
              </div>

              {tab === 'drivers' ? (
                <DataGrid
                  caption="Driver availability"
                  captionVisuallyHidden
                  columns={driverColumns}
                  rows={snapshot?.drivers ?? []}
                  rowKey={(d) => d.id}
                  emptyMessage="No drivers are set up."
                />
              ) : (
                <DataGrid
                  caption="Vehicle availability"
                  captionVisuallyHidden
                  columns={vehicleColumns}
                  rows={snapshot?.vehicles ?? []}
                  rowKey={(v) => v.id}
                  emptyMessage="No vehicles are set up."
                />
              )}

              <span className={styles.staleNote}>
                Driver and vehicle status is maintained by the dispatch flows. If a
                status here looks wrong, it was probably edited by hand in the list -
                the nightly reconciliation will correct it.
              </span>
            </Box>
            ) : undefined}
          </div>
        </div>
      </div>

      {snapshot ? (
        <section className={styles.steps} ref={stepsRef} aria-label="Allocate driver and vehicle">
          <header className={styles.stepsBar}>
            <Icon name="gear" size={18} />
            <h2 className={styles.stepsTitle}>Allocate Driver &amp; Vehicle (Steps)</h2>
          </header>
          <AllocationSteps
            key={assignTarget?.id ?? 'none'}
            request={assignTarget}
            pending={allPending}
            snapshot={snapshot}
            offers={offersForTarget}
            busy={busy}
            lastAllocated={lastAllocated}
            onPick={(r) => {
              openAssign(r).catch(() => undefined);
            }}
            onClose={() => setAssignTarget(undefined)}
            onAssign={doAssign}
          />
        </section>
      ) : undefined}

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
        title={cancelTarget?.status === 'Allocated' ? 'Cancel this trip?' : 'Cancel this request?'}
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
    </div>
  );
};

// -----------------------------------------------------------------------------
// Manual assignment, as six steps
// -----------------------------------------------------------------------------

/**
 * The Coordinator can assign or reassign AT ANY TIME, regardless of the
 * allocation mode.
 *
 * The steps permit an OUT-OF-POLICY choice - a driver already at their cap,
 * or a vehicle of a different type - with a mandatory override reason. They
 * show the eligibility verdict so the Coordinator knows they are overriding
 * rather than discovering it afterwards.
 */
function AllocationSteps({
  request, pending, snapshot, offers, busy, lastAllocated, onPick, onClose, onAssign,
}: {
  request?: ITransportRequest;
  pending: ITransportRequest[];
  snapshot: IDispatchSnapshot;
  offers: IRideOffer[];
  busy: boolean;
  lastAllocated?: ILastAllocated;
  onPick: (request: ITransportRequest) => void;
  onClose: () => void;
  onAssign: (
    driverId: number,
    vehicleId: number,
    isOverride: boolean,
    reason?: string,
  ) => Promise<void>;
}): JSX.Element {
  /**
   * ONE PAIR, TWO WAYS IN.
   *
   * The pairing is permanent: a driver owns a vehicle, recorded on
   * VehicleMaster.DefaultDriver. Two independent lists would let the
   * Coordinator pair driver A with vehicle B and quietly break the invariant
   * the whole engine rests on. So step 2 and step 3 both choose the SAME
   * pair: picking a vehicle brings its driver, and picking a driver brings
   * that driver's best vehicle for this request.
   */
  const [vehicleId, setVehicleId] = React.useState<number | undefined>();
  const [reason, setReason] = React.useState('');

  const windowFrom = request?.requiredDateTime ?? new Date();
  const windowTo = request?.returnDateTime
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

  const requestedTypeId = request?.vehicleTypeRequested?.id;

  /** Every vehicle with its permanent driver, and why it may not be usable. */
  const candidates = React.useMemo(
    () =>
      snapshot.vehicles.map((v) => {
        const problems = explainPairIneligibility(v, ctx).slice();
        if (requestedTypeId !== undefined && v.vehicleType?.id !== requestedTypeId) {
          problems.push(`different vehicle type (${v.vehicleType?.title ?? 'none'})`);
        }
        return { vehicle: v, driver: driversById.get(v.defaultDriver?.id ?? -1), problems };
      }),
    [snapshot.vehicles, driversById, ctx, requestedTypeId],
  );

  const chosen = candidates.filter((c) => c.vehicle.id === vehicleId)[0];

  const isOverride = chosen !== undefined && chosen.problems.length > 0;
  const driverId = chosen?.driver?.id;
  const canAssign =
    request !== undefined &&
    chosen !== undefined &&
    driverId !== undefined &&
    (!isOverride || reason.trim().length > 0);

  const declined = offers.filter((o) => o.responseStatus === 'Rejected');

  /** A driver's best vehicle for this request: suitable first, then any. */
  const vehicleForDriver = (id: number): number | undefined => {
    const own = candidates.filter((c) => c.driver?.id === id);
    const best = own.filter((c) => c.problems.length === 0)[0] ?? own[0];
    return best?.vehicle.id;
  };

  const vehicleLine = (v: IVehicle, extra?: React.ReactNode): JSX.Element => (
    <span className={styles.pickRow}>
      <span className={styles.carThumb} aria-hidden="true">
        <Icon name="car" size={22} />
      </span>
      <span className={styles.pickText}>
        <strong>{vehicleName(v)}</strong>
        <span>{`Plate: ${v.plateNo}`}</span>
        {extra}
      </span>
    </span>
  );

  const trackedTrip = lastAllocated
    ? snapshot.ongoingTrips.filter((t) => t.request.id === lastAllocated.requestItemId)[0]
    : undefined;
  const trackedStatus = trackedTrip?.trip.tripStatus;

  const step = (n: number, title: string, caption: string, body: React.ReactNode, done = false): JSX.Element => (
    <li className={[styles.step, done ? styles.stepDone : ''].join(' ')}>
      <span className={styles.stepNumber} aria-hidden="true">
        {n}
      </span>
      <div className={styles.stepHead}>
        <span className={styles.stepTitle}>
          <span className={styles.srOnly}>{`Step ${n}: `}</span>
          {title}
        </span>
        <span className={styles.stepCaption}>{caption}</span>
      </div>
      <div className={styles.stepBody}>{body}</div>
    </li>
  );

  return (
    <ol className={styles.stepList}>
      {step(
        1,
        'Select Request',
        'Click on a pending request from the list.',
        <>
          <Dropdown
            label="Request to allocate"
            placeholder="Choose a pending request"
            value={
              request ? (
                <span className={styles.pickRow}>
                  <span className={styles.docThumb} aria-hidden="true">
                    <Icon name="doc" size={20} />
                  </span>
                  <span className={styles.pickText}>
                    <strong>
                      {request.requestId}{' '}
                      <span className={styles.pickWhen}>{when(request.requiredDateTime)}</span>
                    </strong>
                    <span>{`${SHORT_TYPE[request.requestType] ?? request.requestType} Request`}</span>
                    <span>{request.vehicleTypeRequested?.title ?? '—'}</span>
                  </span>
                </span>
              ) : undefined
            }
            options={pending.map((p) => ({
              key: p.id,
              selected: request?.id === p.id,
              content: (
                <span className={styles.pickText}>
                  <strong>
                    {p.requestId} <span className={styles.pickWhen}>{when(p.requiredDateTime)}</span>
                  </strong>
                  <span>
                    {`${SHORT_TYPE[p.requestType] ?? p.requestType} · ${p.vehicleTypeRequested?.title ?? '—'}`}
                  </span>
                  <span className={styles.subLine}>{p.status}</span>
                </span>
              ),
            }))}
            onPick={(id) => {
              const r = pending.filter((p) => p.id === id)[0];
              if (r) onPick(r);
            }}
          />
          {request ? (
            <span className={styles.subLine}>
              {request.pickupText} &rarr; {request.dropOffText}
              {request.otherVehicleText ? ` · ${request.otherVehicleText}` : ''}
            </span>
          ) : undefined}
          {request && request.status === 'Awaiting Coordinator' ? (
            <div className={styles.declined}>
              The offer window closed with nobody accepting.
              {declined.length > 0 ? (
                <ul>
                  {declined.map((o) => (
                    <li key={o.id}>
                      {o.driver?.title ?? 'A driver'} declined:{' '}
                      {o.rejectionReason ?? 'no reason given'}
                      {o.rejectionNotes ? ` — ${o.rejectionNotes}` : ''}
                    </li>
                  ))}
                </ul>
              ) : (
                <div>
                  {offers.length === 0
                    ? 'No drivers were eligible for this request.'
                    : `${offers.length} driver(s) were offered it and none responded.`}
                </div>
              )}
            </div>
          ) : undefined}
        </>,
        request !== undefined,
      )}

      {step(
        2,
        'Choose Vehicle',
        'Select an available vehicle based on the required type.',
        <Dropdown
          label="Vehicle"
          placeholder={request ? 'Choose a vehicle' : 'Select a request first'}
          disabled={!request}
          value={
            chosen
              ? vehicleLine(
                  chosen.vehicle,
                  <Pill state={vehicleState(chosen.vehicle.status)} label={chosen.vehicle.status} />,
                )
              : undefined
          }
          options={candidates.map((c) => ({
            key: c.vehicle.id,
            selected: c.vehicle.id === vehicleId,
            disabled: !c.driver,
            content: vehicleLine(
              c.vehicle,
              <span className={c.problems.length === 0 ? styles.ok : styles.warn}>
                {c.problems.length === 0
                  ? `Eligible · ${c.vehicle.status}`
                  : c.problems.join('; ')}
              </span>,
            ),
          }))}
          onPick={(id) => setVehicleId(id)}
        />,
        chosen !== undefined,
      )}

      {step(
        3,
        'Choose Driver',
        'Select an available driver based on schedule and location.',
        <Dropdown
          label="Driver"
          placeholder={request ? 'Choose a driver' : 'Select a request first'}
          disabled={!request}
          value={
            chosen && chosen.driver ? (
              <span className={styles.pickRow}>
                <Avatar name={chosen.driver.driverName} large />
                <span className={styles.pickText}>
                  <strong>{chosen.driver.driverName}</strong>
                  <Pill state={driverState(chosen.driver.status)} label={chosen.driver.status} />
                </span>
              </span>
            ) : undefined
          }
          options={snapshot.drivers.map((d) => {
            const vid = vehicleForDriver(d.id);
            const v = vid !== undefined ? candidates.filter((c) => c.vehicle.id === vid)[0] : undefined;
            return {
              key: d.id,
              selected: chosen?.driver?.id === d.id,
              disabled: vid === undefined,
              content: (
                <span className={styles.pickRow}>
                  <Avatar name={d.driverName} />
                  <span className={styles.pickText}>
                    <strong>{d.driverName}</strong>
                    <span className={v && v.problems.length === 0 ? styles.ok : styles.warn}>
                      {v ? `${d.status} · ${v.vehicle.plateNo}` : 'No vehicle assigned'}
                    </span>
                  </span>
                </span>
              ),
            };
          })}
          onPick={(id) => setVehicleId(vehicleForDriver(id))}
        />,
        chosen !== undefined && chosen.driver !== undefined,
      )}

      {step(
        4,
        'Confirm Allocation',
        'Review trip details and confirm the allocation.',
        <>
          <div className={styles.confirmCard}>
            <span className={styles.pickRow}>
              <span className={styles.carThumb} aria-hidden="true">
                <Icon name="car" size={18} />
              </span>
              <span className={styles.pickText}>
                <strong>Vehicle: {chosen ? vehicleName(chosen.vehicle) : '—'}</strong>
                <span>{chosen ? `Plate: ${chosen.vehicle.plateNo}` : 'Not chosen yet'}</span>
              </span>
            </span>
            <span className={styles.pickRow}>
              <Avatar name={chosen?.driver?.driverName ?? '?'} />
              <span className={styles.pickText}>
                <strong>Driver: {chosen?.driver?.driverName ?? '—'}</strong>
              </span>
            </span>
          </div>

          {isOverride ? (
            <div className={styles.override}>
              <span className={styles.warn}>
                Out-of-policy assignment: {chosen?.problems.join('; ')}. A reason is
                required and is recorded on the trip.
              </span>
              <FormRow label="Override reason" required>
                <TextField
                  value={reason}
                  onChange={(e) => setReason(e.target.value)}
                  placeholder="Why this driver or vehicle despite the rule"
                />
              </FormRow>
            </div>
          ) : undefined}

          <Button
            kind="primary"
            className={styles.allocateButton}
            disabled={!canAssign || busy}
            onClick={() => {
              if (driverId !== undefined && vehicleId !== undefined) {
                onAssign(driverId, vehicleId, isOverride, isOverride ? reason.trim() : undefined).catch(
                  () => undefined,
                );
              }
            }}
          >
            {busy ? 'Allocating…' : 'Allocate'}
            {busy ? undefined : <Icon name="chevronRight" size={18} />}
          </Button>
          {request ? (
            <button type="button" className={styles.textButton} onClick={onClose} disabled={busy}>
              Cancel
            </button>
          ) : undefined}
        </>,
      )}

      {step(
        5,
        'Notify Driver & Requester',
        'The driver and the requester are told about the allocation.',
        <div className={styles.notifyList}>
          {[
            {
              who: 'Driver',
              name: lastAllocated && !request ? lastAllocated.driverName : chosen?.driver?.driverName,
              done: 'Task Assigned',
            },
            { who: 'Requester', name: undefined, done: 'Confirmation Queued' },
          ].map((row) => {
            const sent = lastAllocated !== undefined && !request;
            return (
              <div key={row.who} className={[styles.notifyRow, sent ? styles.notifyDone : ''].join(' ')}>
                {row.name ? (
                  <Avatar name={row.name} />
                ) : (
                  <span className={styles.avatar} aria-hidden="true">
                    <Icon name="user" size={15} />
                  </span>
                )}
                <span className={styles.notifyWho}>{row.who}</span>
                <span className={styles.notifyState}>
                  <span className={styles.notifyTick} aria-hidden="true">
                    <Icon name={sent ? 'check' : 'clock'} size={12} />
                  </span>
                  {sent ? row.done : 'Waiting'}
                </span>
              </div>
            );
          })}
          {lastAllocated && !request ? (
            <span className={styles.subLine}>
              {`${lastAllocated.requestId} allocated to ${lastAllocated.driverName} (${lastAllocated.vehicleName}). ` +
                'The requester’s e-mail is sent by the notification flow.'}
            </span>
          ) : undefined}
        </div>,
        lastAllocated !== undefined && !request,
      )}

      {step(
        6,
        'Track & Manage',
        'Monitor the trip status. The driver updates it as the trip goes.',
        <>
          <div className={styles.flow}>
            {[
              { label: 'En Route', icon: 'pin' as IconName, tone: styles.flowGreen },
              { label: 'In Transit', icon: 'car' as IconName, tone: styles.flowBlue },
              { label: 'Completed', icon: 'check' as IconName, tone: styles.flowGreen },
              { label: 'Cancelled', icon: 'x' as IconName, tone: styles.flowRed },
            ].map((f, i) => (
              <React.Fragment key={f.label}>
                {i > 0 ? <span className={styles.flowArrow} aria-hidden="true" /> : undefined}
                <span className={[styles.flowStep, trackedStatus === f.label ? styles.flowNow : ''].join(' ')}>
                  <span className={[styles.flowDot, f.tone].join(' ')} aria-hidden="true">
                    <Icon name={f.icon} size={16} />
                  </span>
                  {f.label}
                </span>
              </React.Fragment>
            ))}
          </div>
          {lastAllocated ? (
            <span className={styles.subLine}>
              {trackedStatus
                ? `${lastAllocated.requestId} is ${trackedStatus}. This updates on every refresh.`
                : `${lastAllocated.requestId} is assigned - waiting for the driver to set off.`}
            </span>
          ) : undefined}
        </>,
      )}
    </ol>
  );
}

export default DispatchDashboard;
