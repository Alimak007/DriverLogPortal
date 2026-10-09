/**
 * Dates: stored UTC, displayed in Gulf Standard Time, NEVER the browser's zone.
 *
 * A Coordinator in Dubai and a Coordinator travelling must see the same trip
 * at the same clock time. Every display path in the solution goes through
 * here; no component calls toLocaleString itself.
 *
 * Gulf Standard Time is UTC+4 with no daylight saving, which is why the
 * fallback below is a fixed offset and is safe.
 */

const ZONE = 'Asia/Dubai';
const OFFSET_MINUTES = 240; // UTC+4

let intlWorks: boolean | undefined;

function canUseIntl(): boolean {
  if (intlWorks === undefined) {
    try {
      new Intl.DateTimeFormat('en-GB', { timeZone: ZONE }).format(new Date());
      intlWorks = true;
    } catch {
      intlWorks = false;
    }
  }
  return intlWorks;
}

/** Shift a UTC instant into Gulf wall-clock fields. Fallback path only. */
function shifted(d: Date): Date {
  return new Date(d.getTime() + OFFSET_MINUTES * 60000);
}

function fmt(d: Date, options: Intl.DateTimeFormatOptions): string {
  if (canUseIntl()) {
    return new Intl.DateTimeFormat('en-GB', { ...options, timeZone: ZONE }).format(d);
  }
  // Without Intl the offset is applied by hand and the UTC getters are read,
  // which produces the same wall clock.
  const s = shifted(d);
  const two = (n: number): string => (n < 10 ? `0${n}` : String(n));
  const date = `${two(s.getUTCDate())}/${two(s.getUTCMonth() + 1)}/${s.getUTCFullYear()}`;
  const time = `${two(s.getUTCHours())}:${two(s.getUTCMinutes())}`;
  if (options.hour && options.day) return `${date}, ${time}`;
  if (options.hour) return time;
  return date;
}

/** 25/09/2026 */
export function formatDate(d?: Date): string {
  if (!d) return '';
  return fmt(d, { day: '2-digit', month: '2-digit', year: 'numeric' });
}

/** 09:14 */
export function formatTime(d?: Date): string {
  if (!d) return '';
  return fmt(d, { hour: '2-digit', minute: '2-digit', hour12: false });
}

/** 25/09/2026, 09:14 */
export function formatDateTime(d?: Date): string {
  if (!d) return '';
  return fmt(d, {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  });
}

/** 25 Sep 2026 - for headings, where the long form reads better. */
export function formatDateLong(d?: Date): string {
  if (!d) return '';
  return fmt(d, { day: '2-digit', month: 'short', year: 'numeric' });
}

/** Parse what SharePoint returns. Blank and null both become undefined. */
export function parseDate(value?: string): Date | undefined {
  // SharePoint returns null for an empty date; the mapper passes it straight
  // through, so a falsy check covers null, undefined and ''.
  if (!value) return undefined;
  const d = new Date(value);
  return isNaN(d.getTime()) ? undefined : d;
}

/** ISO-8601 UTC, which is what goes on the wire. */
export function toIso(d?: Date): string | undefined {
  return d ? d.toISOString() : undefined;
}

/** "4 min 20 s", for the offer countdown. Never negative. */
export function formatCountdown(totalSeconds: number): string {
  const s = Math.max(0, Math.floor(totalSeconds));
  const m = Math.floor(s / 60);
  const r = s % 60;
  return m > 0 ? `${m} min ${r < 10 ? '0' : ''}${r} s` : `${r} s`;
}

/** "3 minutes ago", for the ETA's last-updated stamp. */
export function relativeFromNow(d: Date, now: Date = new Date()): string {
  const mins = Math.round((now.getTime() - d.getTime()) / 60000);
  if (mins < 1) return 'just now';
  if (mins === 1) return '1 minute ago';
  if (mins < 60) return `${mins} minutes ago`;
  const hrs = Math.round(mins / 60);
  if (hrs === 1) return '1 hour ago';
  if (hrs < 24) return `${hrs} hours ago`;
  return formatDateTime(d);
}

/** Midnight to midnight in Gulf time, expressed as UTC instants. */
export function gulfDayBounds(reference: Date = new Date()): { from: Date; to: Date } {
  const s = shifted(reference);
  const startGulf = Date.UTC(s.getUTCFullYear(), s.getUTCMonth(), s.getUTCDate());
  const from = new Date(startGulf - OFFSET_MINUTES * 60000);
  const to = new Date(from.getTime() + 24 * 3600000);
  return { from, to };
}

/** First and last instant of the month containing the reference date. */
export function gulfMonthBounds(reference: Date = new Date()): { from: Date; to: Date } {
  const s = shifted(reference);
  const startGulf = Date.UTC(s.getUTCFullYear(), s.getUTCMonth(), 1);
  const endGulf = Date.UTC(s.getUTCFullYear(), s.getUTCMonth() + 1, 1);
  return {
    from: new Date(startGulf - OFFSET_MINUTES * 60000),
    to: new Date(endGulf - OFFSET_MINUTES * 60000),
  };
}

export function addMinutes(d: Date, minutes: number): Date {
  return new Date(d.getTime() + minutes * 60000);
}

export function addHours(d: Date, hours: number): Date {
  return new Date(d.getTime() + hours * 3600000);
}

export function addMonths(d: Date, months: number): Date {
  const r = new Date(d.getTime());
  r.setUTCMonth(r.getUTCMonth() + months);
  return r;
}

/**
 * The booking form's time slots. Time is a SLOT LIST, not a free-text time,
 * and past slots are disabled.
 *
 * Returns slots as UTC instants on the given Gulf calendar day.
 */
export interface ITimeSlot {
  label: string;
  value: Date;
  disabled: boolean;
}

export function timeSlots(
  day: Date,
  stepMinutes: number = 30,
  now: Date = new Date(),
  leadMinutes: number = 15,
): ITimeSlot[] {
  const { from } = gulfDayBounds(day);
  const slots: ITimeSlot[] = [];
  const count = Math.floor((24 * 60) / stepMinutes);
  const earliest = addMinutes(now, leadMinutes);

  for (let i = 0; i < count; i++) {
    const value = addMinutes(from, i * stepMinutes);
    slots.push({
      label: formatTime(value),
      value,
      disabled: value < earliest,
    });
  }
  return slots;
}
