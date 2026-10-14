/**
 * The Coordinator's Calendar View.
 *
 * One row per driver (or per vehicle), with every request placed by its
 * REQUIRED time: a day as an hour-by-hour timeline, a week or a month as day
 * columns. Requests nobody has taken yet sit in a lane of their own at the top.
 *
 * DRAG A TASK TO A RESOURCE; REVIEW BEFORE SAVING. Dropping a trip on another
 * row only STAGES the change. Nothing is written until the Coordinator reviews
 * the list, gives a reason and presses Save. Saving uses the same service
 * calls as the rest of the dashboard: reassign for a trip that already has a
 * driver, manual assignment for a request that has none.
 *
 * The pairing rule holds here too: a driver owns a vehicle, so dropping on a
 * driver brings their vehicle, and dropping on a vehicle brings its driver.
 */

import * as React from 'react';
import { Icon } from '../../../common/components/Icons';
import { Button, TextArea } from '../../../common/components/Shared';
import {
  IAvailabilityBlock,
  ICalendarData,
  IDispatchSnapshot,
  IDriver,
  ITransportRequest,
  ITripAssignment,
  IVehicle,
  RequestType,
} from '../../../common/models';
import {
  explainPairIneligibility,
  IEligibilityContext,
} from '../../../common/rules/dispatchRules';
import { IServices } from '../../../common/services/IServices';
import {
  formatDate,
  formatDateLong,
  formatTime,
  gulfDayBounds,
  gulfMonthBounds,
} from '../../../common/utils/dates';
import styles from './DispatchCalendar.module.scss';

export interface IDispatchCalendarProps {
  services: IServices;
  snapshot: IDispatchSnapshot;
  /** Called after a save, so the dashboard reloads its counters and grids. */
  onSaved: () => Promise<void>;
  onError: (e: unknown) => void;
}

type View = 'day' | 'week' | 'month';
type Mode = 'drivers' | 'vehicles';
type ItemState = 'awaiting' | 'confirmed' | 'road' | 'done';

const DAY_MS = 86400000;
const HOUR_MS = 3600000;
const GULF_OFFSET_MS = 4 * HOUR_MS;
const WEEKDAY = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

const SHORT_TYPE: Record<RequestType, string> = {
  'Client or Client Guest': 'Client',
  'Staff': 'Staff',
  'Items or Documents': 'Items / Documents',
};

const STATE_LABEL: Record<ItemState, string> = {
  awaiting: 'Awaiting acceptance',
  confirmed: 'Confirmed',
  road: 'On the road',
  done: 'Completed',
};

// -----------------------------------------------------------------------------
// Gulf calendar arithmetic. Gulf Standard Time has no daylight saving, so a
// day is always 24 hours and adding whole days is safe.
// -----------------------------------------------------------------------------

const addDays = (d: Date, n: number): Date => new Date(d.getTime() + n * DAY_MS);
const dayStart = (d: Date): Date => gulfDayBounds(d).from;
const gulfWeekday = (d: Date): number => new Date(d.getTime() + GULF_OFFSET_MS).getUTCDay();
const gulfDayOfMonth = (d: Date): number => new Date(d.getTime() + GULF_OFFSET_MS).getUTCDate();

/** The week runs Monday to Sunday. */
function weekStart(d: Date): Date {
  return addDays(dayStart(d), -((gulfWeekday(d) + 6) % 7));
}

/** YYYY-MM-DD for the date field. */
function dateValue(d: Date): string {
  return formatDate(dayStart(d)).split('/').reverse().join('-');
}

function two(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function shortPlace(text?: string): string {
  return (text ?? '').split(',')[0].trim() || '—';
}

function initials(name: string): string {
  return (
    name
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((w) => w.charAt(0).toUpperCase())
      .join('') || '?'
  );
}

// -----------------------------------------------------------------------------
// Model
// -----------------------------------------------------------------------------

interface IItem {
  key: number;
  request: ITransportRequest;
  trip?: ITripAssignment;
  start: Date;
  end: Date;
  state: ItemState;
  /** Where it sits now, counting any change still waiting to be saved. */
  driverId?: number;
  vehicleId?: number;
  staged: boolean;
  /** A finished trip is history; it cannot be moved. */
  movable: boolean;
}

interface IChange {
  requestItemId: number;
  requestRef: string;
  tripItemId?: number;
  fromName: string;
  driverId: number;
  vehicleId: number;
  toName: string;
  problems: string[];
  failed?: string;
}

interface IRow {
  id: number;
  name: string;
  sub: string;
  /** Unavailable all period, from the record's own status. */
  statusBlock?: string;
}

function stateOf(trip?: ITripAssignment): ItemState {
  if (!trip) return 'awaiting';
  if (trip.tripStatus === 'Completed') return 'done';
  if (trip.tripStatus === 'En Route' || trip.tripStatus === 'In Transit') return 'road';
  return 'confirmed';
}

function blockLabel(b: IAvailabilityBlock): string {
  return `${b.blockType === 'Leave' ? 'On leave' : b.blockType} · not available`;
}

// -----------------------------------------------------------------------------
// Component
// -----------------------------------------------------------------------------

export const DispatchCalendar: React.FC<IDispatchCalendarProps> = ({
  services,
  snapshot,
  onSaved,
  onError,
}) => {
  const [view, setView] = React.useState<View>('day');
  const [mode, setMode] = React.useState<Mode>('drivers');
  const [cursor, setCursor] = React.useState<Date>(() => new Date());
  const [data, setData] = React.useState<ICalendarData | undefined>();
  const [loading, setLoading] = React.useState(true);

  const [changes, setChanges] = React.useState<Record<number, IChange>>({});
  const [reason, setReason] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [notice, setNotice] = React.useState<string | undefined>();

  const [dragKey, setDragKey] = React.useState<number | undefined>();
  const [overRow, setOverRow] = React.useState<number | undefined>();
  const [popup, setPopup] = React.useState<{ key: number; top: number; left: number } | undefined>();

  // ---------------------------------------------------------------------------
  // The period
  // ---------------------------------------------------------------------------

  const range = React.useMemo(() => {
    if (view === 'day') {
      const from = dayStart(cursor);
      return { from, to: addDays(from, 1), days: 1 };
    }
    if (view === 'week') {
      const from = weekStart(cursor);
      return { from, to: addDays(from, 7), days: 7 };
    }
    const b = gulfMonthBounds(cursor);
    return { from: b.from, to: b.to, days: Math.round((b.to.getTime() - b.from.getTime()) / DAY_MS) };
  }, [view, cursor]);

  const move = (dir: number): void => {
    if (view === 'day') setCursor(addDays(cursor, dir));
    else if (view === 'week') setCursor(addDays(cursor, dir * 7));
    else {
      // Into the next or previous month: step from the middle of this one.
      setCursor(addDays(range.from, dir > 0 ? range.days + 1 : -1));
    }
  };

  // Reload on a new period, and whenever the dashboard itself refreshes.
  const loadedAt = snapshot.loadedAt.getTime();
  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);
    services.dispatch
      .getCalendar(range.from, range.to)
      .then(
        (d) => {
          if (!cancelled) setData(d);
        },
        (e) => {
          if (!cancelled) onError(e);
        },
      )
      .then(() => {
        if (!cancelled) setLoading(false);
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [services, range.from.getTime(), range.to.getTime(), loadedAt]);

  // ---------------------------------------------------------------------------
  // Lookups and eligibility
  // ---------------------------------------------------------------------------

  const driversById = React.useMemo(() => {
    const map = new Map<number, IDriver>();
    snapshot.drivers.forEach((d) => map.set(d.id, d));
    return map;
  }, [snapshot.drivers]);

  const vehiclesById = React.useMemo(() => {
    const map = new Map<number, IVehicle>();
    snapshot.vehicles.forEach((v) => map.set(v.id, v));
    return map;
  }, [snapshot.vehicles]);

  const vehicleName = (v?: IVehicle): string =>
    v ? [v.make, v.model].filter(Boolean).join(' ') || v.plateNo : '—';

  const items: IItem[] = React.useMemo(() => {
    if (!data) return [];
    return data.entries.map((e) => {
      const start = e.request.requiredDateTime;
      // The same two-hour default the allocation engine reserves.
      const end = e.request.returnDateTime ?? new Date(start.getTime() + 2 * HOUR_MS);
      const change = changes[e.request.id];
      return {
        key: e.request.id,
        request: e.request,
        trip: e.trip,
        start,
        end,
        state: stateOf(e.trip),
        driverId: change ? change.driverId : e.trip?.driver?.id,
        vehicleId: change ? change.vehicleId : e.trip?.vehicle?.id,
        staged: change !== undefined,
        movable: e.trip?.tripStatus !== 'Completed',
      };
    });
  }, [data, changes]);

  const itemByKey = (key?: number): IItem | undefined =>
    key === undefined ? undefined : items.filter((i) => i.key === key)[0];

  /** Why a pairing breaks the rules, for the review list. */
  const problemsFor = (item: IItem, vehicle: IVehicle): string[] => {
    const ctx: IEligibilityContext = {
      driversById,
      blocks: (data?.blocks ?? []).filter(
        (b) => b.fromDateTime < item.end && b.toDateTime > item.start,
      ),
      openTrips: snapshot.ongoingTrips.map((t) => t.trip).filter((t) => t.request?.id !== item.key),
      windowTo: item.end,
      ruleMaxConcurrent: 1,
    };
    const problems = explainPairIneligibility(vehicle, ctx).slice();
    const wanted = item.request.vehicleTypeRequested?.id;
    if (wanted !== undefined && vehicle.vehicleType?.id !== wanted) {
      problems.push(`different vehicle type (${vehicle.vehicleType?.title ?? 'none'})`);
    }
    return problems;
  };

  /** A driver's best vehicle for this request: the right type first. */
  const vehicleOf = (driverId: number, item: IItem): IVehicle | undefined => {
    const own = snapshot.vehicles.filter((v) => v.defaultDriver?.id === driverId);
    const wanted = item.request.vehicleTypeRequested?.id;
    return own.filter((v) => v.vehicleType?.id === wanted)[0] ?? own[0];
  };

  /**
   * Stage a move. Writes nothing - the change waits in the review list.
   * Returns a message when the move is not possible.
   */
  const stage = (item: IItem, rowId: number): string | undefined => {
    if (!item.movable) return `${item.request.requestId} is completed and cannot be moved.`;

    let driverId: number | undefined;
    let vehicle: IVehicle | undefined;
    if (mode === 'drivers') {
      driverId = rowId;
      vehicle = vehicleOf(rowId, item);
      if (!vehicle) {
        return `${driversById.get(rowId)?.driverName ?? 'That driver'} has no vehicle assigned, so the trip cannot go to them.`;
      }
    } else {
      vehicle = vehiclesById.get(rowId);
      driverId = vehicle?.defaultDriver?.id;
      if (!vehicle || driverId === undefined || !driversById.has(driverId)) {
        return `${vehicle?.plateNo ?? 'That vehicle'} has no active driver assigned, so the trip cannot go to it.`;
      }
    }

    const original = { driverId: item.trip?.driver?.id, vehicleId: item.trip?.vehicle?.id };
    const next = { ...changes };
    if (original.driverId === driverId && original.vehicleId === vehicle.id) {
      // Dropped back where it started: nothing to save.
      delete next[item.key];
    } else {
      next[item.key] = {
        requestItemId: item.key,
        requestRef: item.request.requestId,
        tripItemId: item.trip?.id,
        fromName: item.trip?.driver?.title ?? 'Unassigned',
        driverId,
        vehicleId: vehicle.id,
        toName: `${driversById.get(driverId)?.driverName ?? '—'} (${vehicleName(vehicle)} · ${vehicle.plateNo})`,
        problems: problemsFor(item, vehicle),
      };
    }
    setChanges(next);
    return undefined;
  };

  const tryStage = (key: number | undefined, rowId: number): void => {
    const item = itemByKey(key);
    if (!item) return;
    const message = stage(item, rowId);
    setNotice(message);
  };

  // ---------------------------------------------------------------------------
  // Save
  // ---------------------------------------------------------------------------

  const pending = Object.keys(changes).map((k) => changes[Number(k)]);

  const save = async (): Promise<void> => {
    if (reason.trim().length === 0) return;
    setSaving(true);
    setNotice(undefined);
    const left: Record<number, IChange> = {};
    let saved = 0;
    for (const c of pending) {
      try {
        const result = c.tripItemId !== undefined
          ? await services.dispatch.reassign(c.tripItemId, c.driverId, c.vehicleId, reason.trim())
          : await services.dispatch.manualAssign(
              c.requestItemId,
              c.driverId,
              c.vehicleId,
              c.problems.length > 0,
              c.problems.length > 0 ? reason.trim() : undefined,
            );
        if (result.ok) saved += 1;
        else left[c.requestItemId] = { ...c, failed: result.message };
      } catch (e) {
        left[c.requestItemId] = { ...c, failed: String((e as Error)?.message ?? e) };
      }
    }
    setChanges(left);
    setSaving(false);
    if (Object.keys(left).length === 0) setReason('');
    setNotice(
      `${saved} change${saved === 1 ? '' : 's'} saved.` +
        (Object.keys(left).length > 0 ? ' The ones still listed below could not be saved.' : ''),
    );
    await onSaved();
  };

  // ---------------------------------------------------------------------------
  // Rows
  // ---------------------------------------------------------------------------

  const rows: IRow[] = React.useMemo(() => {
    if (mode === 'drivers') {
      return snapshot.drivers
        .slice()
        .sort((a, b) => a.driverName.localeCompare(b.driverName))
        .map((d) => ({
          id: d.id,
          name: d.driverName,
          sub: d.homeProperty?.title ?? '',
          statusBlock:
            d.status === 'On Leave' ? 'On leave · not available'
              : d.status === 'Inactive' ? 'Inactive · not available' : undefined,
        }));
    }
    return snapshot.vehicles
      .slice()
      .sort((a, b) => a.plateNo.localeCompare(b.plateNo))
      .map((v) => ({
        id: v.id,
        name: vehicleName(v),
        sub: `${v.plateNo}${v.defaultDriver ? ` · ${v.defaultDriver.title}` : ''}`,
        statusBlock:
          v.status === 'Under Maintenance' ? 'Under maintenance · not available'
            : v.status === 'Retired' ? 'Retired · not available' : undefined,
      }));
  }, [mode, snapshot.drivers, snapshot.vehicles]);

  const rowOf = (item: IItem): number | undefined => (mode === 'drivers' ? item.driverId : item.vehicleId);
  const unassigned = items.filter((i) => rowOf(i) === undefined);

  const blocksFor = (rowId: number): IAvailabilityBlock[] =>
    (data?.blocks ?? []).filter(
      (b) =>
        b.blockType !== 'Shift' &&
        (mode === 'drivers' ? b.driver?.id === rowId : b.vehicle?.id === rowId),
    );

  // ---------------------------------------------------------------------------
  // Day view geometry
  // ---------------------------------------------------------------------------

  const hours = React.useMemo(() => {
    // A working day by default, widened to fit any trip outside it.
    let first = 7;
    let last = 19;
    if (view === 'day') {
      items.forEach((i) => {
        const s = (i.start.getTime() - range.from.getTime()) / HOUR_MS;
        const e = (i.end.getTime() - range.from.getTime()) / HOUR_MS;
        if (s >= 0 && s < 24) first = Math.min(first, Math.floor(s));
        if (e > 0) last = Math.max(last, Math.min(24, Math.ceil(e)));
      });
    }
    const list: number[] = [];
    for (let h = first; h < last; h++) list.push(h);
    return list;
  }, [view, items, range.from]);

  const dayFrom = range.from.getTime() + hours[0] * HOUR_MS;
  const dayTo = range.from.getTime() + (hours[hours.length - 1] + 1) * HOUR_MS;
  const pct = (t: number): number => Math.max(0, Math.min(100, ((t - dayFrom) / (dayTo - dayFrom)) * 100));

  /** Overlapping trips on one row sit in lanes, one under the other. */
  const laneOf = (list: IItem[]): Map<number, number> => {
    const lanes: number[] = [];
    const out = new Map<number, number>();
    list
      .slice()
      .sort((a, b) => a.start.getTime() - b.start.getTime())
      .forEach((i) => {
        let lane = lanes.findIndex((end) => end <= i.start.getTime());
        if (lane === -1) {
          lane = lanes.length;
          lanes.push(0);
        }
        lanes[lane] = i.end.getTime();
        out.set(i.key, lane);
      });
    return out;
  };

  const days: Date[] = React.useMemo(() => {
    const list: Date[] = [];
    for (let i = 0; i < range.days; i++) list.push(addDays(range.from, i));
    return list;
  }, [range.from, range.days]);

  const inDay = (d: Date, t: Date): boolean => t >= d && t < addDays(d, 1);

  // ---------------------------------------------------------------------------
  // Pieces
  // ---------------------------------------------------------------------------

  const plateOf = (item: IItem): string =>
    item.vehicleId !== undefined ? vehiclesById.get(item.vehicleId)?.plateNo ?? '' : '';

  const openPopup = (e: React.MouseEvent<HTMLElement>, item: IItem): void => {
    const r = e.currentTarget.getBoundingClientRect();
    setPopup({
      key: item.key,
      top: Math.min(r.bottom + 6, window.innerHeight - 320),
      left: Math.max(8, Math.min(r.left, window.innerWidth - 320)),
    });
  };

  const eventButton = (item: IItem, compact: boolean, style?: React.CSSProperties): JSX.Element => (
    <button
      key={item.key}
      type="button"
      className={[
        styles.event,
        styles[`state_${item.state}`],
        compact ? styles.eventCompact : '',
        item.staged ? styles.eventStaged : '',
        dragKey === item.key ? styles.eventDragging : '',
      ].join(' ')}
      style={style}
      draggable={item.movable}
      onDragStart={(e) => {
        e.dataTransfer.setData('text/plain', String(item.key));
        e.dataTransfer.effectAllowed = 'move';
        setDragKey(item.key);
      }}
      onDragEnd={() => {
        setDragKey(undefined);
        setOverRow(undefined);
      }}
      onClick={(e) => openPopup(e, item)}
      title={`${item.request.requestId} · ${formatTime(item.start)} · ${item.request.dropOffText}`}
      aria-label={
        `${item.request.requestId} to ${shortPlace(item.request.dropOffText)}, ` +
        `${formatTime(item.start)} to ${formatTime(item.end)}, ${STATE_LABEL[item.state]}` +
        (item.staged ? ', moved - not saved yet' : '')
      }
    >
      <span className={styles.eventTitle}>
        {/* A month cell has room for the time only; the rest is in the popup. */}
        {compact
          ? formatTime(item.start)
          : `${item.request.requestId} · ${shortPlace(item.request.dropOffText)}`}
      </span>
      {compact ? undefined : (
        <>
          <span className={styles.eventLine}>{`${formatTime(item.start)}–${formatTime(item.end)}`}</span>
          <span className={styles.eventLine}>
            {SHORT_TYPE[item.request.requestType] ?? item.request.requestType}
            {plateOf(item) ? ` · ${plateOf(item)}` : ''}
          </span>
        </>
      )}
      {item.staged ? <span className={styles.stagedMark}>Moved</span> : undefined}
    </button>
  );

  /** A drop target: the track of one resource row. */
  const dropProps = (rowId: number): React.HTMLAttributes<HTMLDivElement> => ({
    // Always accept the drag; the drop itself checks that it was one of ours.
    // Waiting for the dragKey state here races the first dragover.
    onDragOver: (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      if (overRow !== rowId) setOverRow(rowId);
    },
    onDragLeave: () => setOverRow(undefined),
    onDrop: (e) => {
      e.preventDefault();
      const key = parseInt(e.dataTransfer.getData('text/plain'), 10);
      tryStage(isNaN(key) ? dragKey : key, rowId);
      setDragKey(undefined);
      setOverRow(undefined);
    },
  });

  const bandStyle = (from: Date, to: Date): React.CSSProperties => {
    const left = pct(Math.max(from.getTime(), dayFrom));
    const right = pct(Math.min(to.getTime(), dayTo));
    return { left: `${left}%`, width: `${Math.max(0, right - left)}%` };
  };

  const renderDayTrack = (rowItems: IItem[], bands: { from: Date; to: Date; label: string }[]): JSX.Element => {
    const lanes = laneOf(rowItems);
    const laneCount = Math.max(1, ...Array.from(lanes.values()).map((l) => l + 1));
    return (
      <div className={styles.dayTrack} style={{ minHeight: laneCount * 62 + 10 }}>
        {hours.map((h) => (
          <span key={h} className={styles.hourLine} style={{ left: `${pct(range.from.getTime() + h * HOUR_MS)}%` }} />
        ))}
        {bands.map((b, i) => (
          <span key={i} className={styles.band} style={bandStyle(b.from, b.to)}>
            {b.label}
          </span>
        ))}
        {rowItems.map((i) => {
          const left = pct(i.start.getTime());
          const right = pct(i.end.getTime());
          return eventButton(i, false, {
            left: `calc(${left}% + 2px)`,
            width: `calc(${Math.max(right - left, 4)}% - 4px)`,
            top: 5 + (lanes.get(i.key) ?? 0) * 62,
          });
        })}
      </div>
    );
  };

  const renderDayCells = (rowItems: IItem[], bands: { from: Date; to: Date; label: string }[]): JSX.Element => (
    <div className={styles.cellTrack} style={{ gridTemplateColumns: `repeat(${days.length}, minmax(0, 1fr))` }}>
      {days.map((d) => {
        const dayItems = rowItems
          .filter((i) => inDay(d, i.start))
          .sort((a, b) => a.start.getTime() - b.start.getTime());
        const blocked = bands.some((b) => b.from < addDays(d, 1) && b.to > d);
        return (
          <div key={d.getTime()} className={[styles.cell, blocked ? styles.cellBlocked : ''].join(' ')}>
            {dayItems.map((i) => eventButton(i, view === 'month'))}
          </div>
        );
      })}
    </div>
  );

  const bandsForRow = (row: IRow): { from: Date; to: Date; label: string }[] => {
    const list = blocksFor(row.id).map((b) => ({ from: b.fromDateTime, to: b.toDateTime, label: blockLabel(b) }));
    if (row.statusBlock && list.length === 0) {
      list.push({ from: range.from, to: range.to, label: row.statusBlock });
    }
    return list;
  };

  const minWidth =
    view === 'day' ? 210 + hours.length * 66 : view === 'week' ? 210 + 7 * 130 : 210 + range.days * 54;

  const rangeLabel =
    view === 'day'
      ? formatDateLong(range.from)
      : view === 'week'
        ? `${formatDateLong(range.from)} – ${formatDateLong(addDays(range.to, -1))}`
        : formatDateLong(range.from).replace(/^\d+\s/, '');

  const selected = itemByKey(popup?.key);
  const resourceOptions = rows.map((r) => ({ id: r.id, name: r.name }));

  // Close the popup on Escape or a click outside it.
  const popupRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!popup) return undefined;
    const onDown = (e: MouseEvent): void => {
      if (popupRef.current && !popupRef.current.contains(e.target as Node)) setPopup(undefined);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setPopup(undefined);
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [popup]);

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  return (
    <section className={styles.calendar} aria-label="Calendar view">
      <div className={styles.toolbar}>
        <div className={styles.nav}>
          <button type="button" className={styles.navButton} onClick={() => move(-1)} aria-label={`Previous ${view}`}>
            <Icon name="chevronLeft" size={16} />
          </button>
          <button type="button" className={styles.todayButton} onClick={() => setCursor(new Date())}>
            Today
          </button>
          <button type="button" className={styles.navButton} onClick={() => move(1)} aria-label={`Next ${view}`}>
            <Icon name="chevronRight" size={16} />
          </button>
          <input
            type="date"
            className={styles.dateInput}
            value={dateValue(cursor)}
            aria-label="Go to date"
            onChange={(e) => {
              const parsed = new Date(`${e.target.value}T00:00:00Z`);
              if (!isNaN(parsed.getTime())) setCursor(parsed);
            }}
          />
          <span className={styles.rangeLabel}>{rangeLabel}</span>
        </div>

        <div className={styles.segment} role="group" aria-label="Calendar period">
          {(['day', 'week', 'month'] as View[]).map((v) => (
            <button
              key={v}
              type="button"
              className={[styles.segmentButton, view === v ? styles.segmentActive : ''].join(' ')}
              aria-pressed={view === v}
              onClick={() => setView(v)}
            >
              {v.charAt(0).toUpperCase() + v.slice(1)}
            </button>
          ))}
        </div>
      </div>

      <div className={styles.subbar}>
        <div className={styles.segment} role="group" aria-label="Rows">
          <button
            type="button"
            className={[styles.segmentButton, mode === 'drivers' ? styles.segmentActive : ''].join(' ')}
            aria-pressed={mode === 'drivers'}
            onClick={() => setMode('drivers')}
          >
            <Icon name="users" size={15} />
            Drivers
          </button>
          <button
            type="button"
            className={[styles.segmentButton, mode === 'vehicles' ? styles.segmentActive : ''].join(' ')}
            aria-pressed={mode === 'vehicles'}
            onClick={() => setMode('vehicles')}
          >
            <Icon name="car" size={15} />
            Vehicles
          </button>
        </div>
        <span className={styles.hint}>
          {loading ? 'Loading…' : 'Drag a task to a resource; review before saving.'}
        </span>
      </div>

      {notice ? (
        <div className={styles.notice} role="status">
          {notice}
          <button type="button" className={styles.noticeClose} onClick={() => setNotice(undefined)} aria-label="Dismiss">
            <Icon name="x" size={14} />
          </button>
        </div>
      ) : undefined}

      <div className={styles.scroller}>
        <div className={styles.board} style={{ minWidth }}>
          <div className={[styles.row, styles.headRow].join(' ')}>
            <div className={styles.label}>
              <span className={styles.headLabel}>{mode === 'drivers' ? 'Driver / Base property' : 'Vehicle / Driver'}</span>
            </div>
            {view === 'day' ? (
              <div className={styles.dayTrack}>
                {hours.map((h) => (
                  <span key={h} className={styles.hourLabel} style={{ left: `${pct(range.from.getTime() + h * HOUR_MS)}%` }}>
                    {`${two(h)}:00`}
                  </span>
                ))}
              </div>
            ) : (
              <div className={styles.cellTrack} style={{ gridTemplateColumns: `repeat(${days.length}, minmax(0, 1fr))` }}>
                {days.map((d) => (
                  <span
                    key={d.getTime()}
                    className={[styles.dayHead, dayStart(new Date()).getTime() === d.getTime() ? styles.dayHeadToday : ''].join(' ')}
                  >
                    {view === 'week' ? `${WEEKDAY[gulfWeekday(d)]} ${gulfDayOfMonth(d)}` : gulfDayOfMonth(d)}
                  </span>
                ))}
              </div>
            )}
          </div>

          {unassigned.length > 0 ? (
            <div className={[styles.row, styles.unassignedRow].join(' ')}>
              <div className={styles.label}>
                <span className={[styles.avatar, styles.avatarAmber].join(' ')} aria-hidden="true">
                  <Icon name="clock" size={14} />
                </span>
                <span className={styles.labelText}>
                  <strong>Unassigned</strong>
                  <span>Drag onto a {mode === 'drivers' ? 'driver' : 'vehicle'} to allocate</span>
                </span>
              </div>
              {view === 'day' ? renderDayTrack(unassigned, []) : renderDayCells(unassigned, [])}
            </div>
          ) : undefined}

          {rows.map((row) => {
            const rowItems = items.filter((i) => rowOf(i) === row.id);
            const bands = bandsForRow(row);
            return (
              <div
                key={row.id}
                className={[styles.row, overRow === row.id ? styles.rowOver : ''].join(' ')}
                {...dropProps(row.id)}
              >
                <div className={styles.label}>
                  <span className={styles.avatar} aria-hidden="true">
                    {mode === 'drivers' ? initials(row.name) : <Icon name="car" size={14} />}
                  </span>
                  <span className={styles.labelText}>
                    <strong>{row.name}</strong>
                    <span>{row.sub}</span>
                  </span>
                </div>
                {view === 'day' ? renderDayTrack(rowItems, bands) : renderDayCells(rowItems, bands)}
              </div>
            );
          })}

          {rows.length === 0 ? (
            <div className={styles.empty}>No {mode} are set up.</div>
          ) : undefined}
        </div>
      </div>

      <div className={styles.legend}>
        {(['awaiting', 'confirmed', 'road', 'done'] as ItemState[]).map((s) => (
          <span key={s} className={styles.legendItem}>
            <span className={[styles.legendDot, styles[`dot_${s}`]].join(' ')} aria-hidden="true" />
            {STATE_LABEL[s]}
          </span>
        ))}
        <span className={styles.legendItem}>
          <span className={[styles.legendDot, styles.dot_off].join(' ')} aria-hidden="true" />
          Unavailable
        </span>
      </div>

      {pending.length > 0 ? (
        <div className={styles.review} aria-label="Changes waiting to be saved">
          <div className={styles.reviewHead}>
            <strong>{`Review before saving (${pending.length})`}</strong>
            <span>Nothing is written until you save.</span>
          </div>
          <ul className={styles.reviewList}>
            {pending.map((c) => (
              <li key={c.requestItemId} className={c.failed ? styles.reviewFailed : ''}>
                <span>
                  <strong>{c.requestRef}</strong>
                  {`: ${c.fromName} → ${c.toName}`}
                  {c.problems.length > 0 ? (
                    <span className={styles.reviewWarn}>{`Out of policy: ${c.problems.join('; ')}`}</span>
                  ) : undefined}
                  {c.failed ? <span className={styles.reviewWarn}>{`Not saved: ${c.failed}`}</span> : undefined}
                </span>
                <button
                  type="button"
                  className={styles.reviewRemove}
                  aria-label={`Undo the move of ${c.requestRef}`}
                  onClick={() => {
                    const next = { ...changes };
                    delete next[c.requestItemId];
                    setChanges(next);
                  }}
                >
                  <Icon name="x" size={14} />
                </button>
              </li>
            ))}
          </ul>
          <label className={styles.reasonLabel} htmlFor="calendar-reason">
            Reason for these changes (required, recorded on each trip)
          </label>
          <TextArea id="calendar-reason" value={reason} onChange={(e) => setReason(e.target.value)} rows={2} />
          <div className={styles.reviewActions}>
            <Button kind="secondary" onClick={() => setChanges({})} disabled={saving}>
              Discard all
            </Button>
            <Button
              kind="primary"
              onClick={() => {
                save().catch((e) => onError(e));
              }}
              disabled={saving || reason.trim().length === 0}
            >
              {saving ? 'Saving…' : `Save ${pending.length} change${pending.length === 1 ? '' : 's'}`}
            </Button>
          </div>
        </div>
      ) : undefined}

      {popup && selected ? (
        <div
          ref={popupRef}
          className={styles.popup}
          style={{ top: popup.top, left: popup.left }}
          role="dialog"
          aria-label={`${selected.request.requestId} details`}
        >
          <div className={styles.popupHead}>
            <strong>{selected.request.requestId}</strong>
            <span className={[styles.statePill, styles[`state_${selected.state}`]].join(' ')}>
              {STATE_LABEL[selected.state]}
            </span>
            <button type="button" className={styles.noticeClose} onClick={() => setPopup(undefined)} aria-label="Close">
              <Icon name="x" size={14} />
            </button>
          </div>
          <dl className={styles.facts}>
            <dt>When</dt>
            <dd>{`${formatDateLong(selected.start)}, ${formatTime(selected.start)}–${formatTime(selected.end)}`}</dd>
            <dt>From</dt>
            <dd>{selected.request.pickupText}</dd>
            <dt>To</dt>
            <dd>{selected.request.dropOffText}</dd>
            <dt>Type</dt>
            <dd>{`${selected.request.requestType} · ${selected.request.vehicleTypeRequested?.title ?? '—'}`}</dd>
            <dt>Driver</dt>
            <dd>{selected.driverId !== undefined ? driversById.get(selected.driverId)?.driverName ?? '—' : 'Not assigned'}</dd>
            <dt>Vehicle</dt>
            <dd>
              {selected.vehicleId !== undefined
                ? `${vehicleName(vehiclesById.get(selected.vehicleId))} · ${plateOf(selected)}`
                : '—'}
            </dd>
          </dl>
          {selected.movable ? (
            <label className={styles.moveLabel}>
              {`Move to ${mode === 'drivers' ? 'driver' : 'vehicle'}`}
              <select
                className={styles.moveSelect}
                value=""
                onChange={(e) => {
                  const id = parseInt(e.target.value, 10);
                  if (!isNaN(id)) {
                    tryStage(selected.key, id);
                    setPopup(undefined);
                  }
                }}
              >
                <option value="">Choose&hellip;</option>
                {resourceOptions.map((r) => (
                  <option key={r.id} value={r.id}>
                    {r.name}
                  </option>
                ))}
              </select>
            </label>
          ) : (
            <span className={styles.hint}>A completed trip cannot be moved.</span>
          )}
        </div>
      ) : undefined}
    </section>
  );
};

export default DispatchCalendar;
