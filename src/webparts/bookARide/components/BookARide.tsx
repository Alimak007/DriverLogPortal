/**
 * Book a Ride - the seven-screen booking wizard.
 *
 * ONE FREE-TEXT FIELD IN THE ENTIRE FLOW. The client's requirement is
 * explicit: predefined selections wherever possible, to prevent incorrect
 * information. The single exception is the text revealed when the requester
 * picks the "Others" vehicle type.
 *
 * WHAT THIS COMPONENT NEVER READS: vehicle availability, driver availability,
 * the list of vehicles, or the identity of any driver. It reads VehicleTypes,
 * Locations and PropertyMaster, and nothing else about the fleet. Anything it
 * read would be reachable by this user through the REST API regardless of
 * what the interface displays.
 */

import * as React from 'react';
import { Icon, IconName } from '../../../common/components/Icons';
import {
  Badge,
  Button,
  EmptyState,
  ErrorBanner,
  FormRow,
  LinkButton,
  Loading,
  Notice,
  PageShell,
  Reference,
  Select,
  TextArea,
  TextField,
} from '../../../common/components/Shared';
import { buildStages, EtaDisplay, Timeline } from '../../../common/components/Timeline';
import {
  IRequestWithTrip,
  ITransportRequest,
  IVehicleType,
  ItemDescription,
  ITEM_DESCRIPTIONS,
  REQUEST_TYPES,
  RequestType,
  TripStatus,
  requestState,
} from '../../../common/models';
import { ICurrentUser, IServices } from '../../../common/services/IServices';
import {
  formatDate,
  formatDateLong,
  formatDateTime,
  gulfDayBounds,
  ITimeSlot,
  timeSlots,
} from '../../../common/utils/dates';
import { distanceKm } from '../../../common/rules/dispatchRules';
import MapPicker, { IPlace } from './MapPicker';
import RouteMap from './RouteMap';
import styles from './BookARide.module.scss';

export interface IBookARideProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  /** Property to book against. Defaults to the requester's own property. */
  defaultPropertyId: number;
  /**
   * TomTom key for the map tiles and the place search.
   *
   * The SPFx bundle is readable by any authenticated user, so this key IS
   * extractable. That is expected for a browser key - but it must be
   * restricted by HTTP referrer to the SharePoint domain in the TomTom
   * developer portal, and a daily usage alert must be set. Do not ship
   * without both.
   */
  tomtomApiKey: string;
  /** How long screen 5 waits for a driver before offering to carry on. */
  processingPollMs: number;
  /** The My Bookings page. The link to it is left out when this is not given. */
  myBookingsUrl?: string;
}

type ScreenIndex = 1 | 2 | 3 | 4 | 5 | 6 | 7;

/**
 * What each screen is called. `title` heads the card, `step` and `caption`
 * are what the step indicator says about it. Wording only - nothing here
 * decides behaviour.
 */
const SCREEN_META: Record<
  ScreenIndex,
  { title: string; step: string; caption: string; icon: IconName }
> = {
  1: {
    title: 'Book a Ride',
    step: 'Open Request',
    caption: 'Choose where the trip starts and where it is going.',
    icon: 'car',
  },
  2: {
    title: 'Date & Time',
    step: 'Select Date & Time',
    caption: 'Choose the date and time the trip is needed.',
    icon: 'calendar',
  },
  3: {
    title: 'Type of Request',
    step: 'Type of Request',
    caption: 'Choose the purpose of the request.',
    icon: 'doc',
  },
  4: {
    title: 'Vehicle Type',
    step: 'Select Vehicle Type',
    caption: 'Choose the kind of vehicle you need.',
    icon: 'car',
  },
  5: {
    title: 'Processing Request',
    step: 'Driver & Vehicle Allocation',
    caption: 'A driver and vehicle are being matched to your request.',
    icon: 'gear',
  },
  6: {
    title: 'Booking Confirmed',
    step: 'Booking Confirmation',
    caption: 'Your driver and vehicle details.',
    icon: 'calendar',
  },
  7: {
    title: 'Track My Ride',
    step: 'Track My Ride',
    caption: 'Follow the status of your trip.',
    icon: 'navigation',
  },
};

const REQUEST_TYPE_CAPTION: Record<RequestType, string> = {
  'Client or Client Guest': 'A client or their guest is travelling.',
  'Staff': 'A member of staff is travelling.',
  'Items or Documents': 'Something is being moved, not someone.',
};

interface IDraft {
  pickup?: IPlace;
  dropOff?: IPlace;
  pickupNotes: string;
  dropOffNotes: string;
  isAdvance: boolean;
  day: Date;
  slot?: Date;
  requestType?: RequestType;
  vehicleType?: IVehicleType;
  otherVehicleText: string;
  noOfPax: string;
  itemDescription?: ItemDescription;
  itemNotes: string;
  remarks: string;
}

const emptyDraft = (): IDraft => ({
  pickupNotes: '',
  dropOffNotes: '',
  isAdvance: false,
  day: new Date(),
  otherVehicleText: '',
  noOfPax: '',
  itemNotes: '',
  remarks: '',
});

export const BookARide: React.FC<IBookARideProps> = (props) => {
  const {
    services, user, bootstrapError, defaultPropertyId, processingPollMs,
    myBookingsUrl, tomtomApiKey,
  } = props;

  const [screen, setScreen] = React.useState<ScreenIndex>(1);
  const [draft, setDraft] = React.useState<IDraft>(emptyDraft());
  const [vehicleTypes, setVehicleTypes] = React.useState<IVehicleType[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [submitting, setSubmitting] = React.useState(false);
  const [error, setError] = React.useState<unknown>(bootstrapError);
  const [validation, setValidation] = React.useState<string | undefined>();

  const [requestItemId, setRequestItemId] = React.useState<number | undefined>();
  const [requestId, setRequestId] = React.useState<string>('');
  const [tracked, setTracked] = React.useState<IRequestWithTrip | undefined>();

  /**
   * The operating area, for the out-of-range warning and to bias the search
   * toward local results. Read from AllocationRules - never hard-coded, since
   * the 60 km radius is a figure the business has not confirmed.
   */
  const [area, setArea] = React.useState({
    centre: { lat: 24.4539, lng: 54.3773 },
    radiusKm: 60,
  });

  /**
   * The requester's own property, if the site tells us one. Used ONLY to
   * default the pickup point. It is a stable number, so it never causes the
   * locations below to reload.
   */
  const homePropertyId = React.useMemo(
    () => defaultPropertyId || 0,
    [defaultPropertyId],
  );

  /**
   * Which property is this trip for?
   *
   * Derived from the PICKUP LOCATION, because a trip belongs to the property
   * it starts at - that is where the driver collects, so that is whose drivers
   * and vehicles should be offered it.
   *
   * It is deliberately NOT used to filter the location pickers. A trip may run
   * from any location to any other, including across properties: Abu Dhabi to
   * Dubai airport is an ordinary journey, and narrowing the destination list
   * to the pickup's property would make it impossible to book.
   */
  /**
   * Written on the request FOR REPORTING ONLY. Allocation ignores it entirely
   * now - who gets offered the trip is decided by vehicle type and distance.
   */
  const propertyId = homePropertyId;

  // ---------------------------------------------------------------------------
  // Reference data
  // ---------------------------------------------------------------------------

  React.useEffect(() => {
    if (!services) {
      setLoading(false);
      return;
    }
    let cancelled = false;

    (async () => {
      try {
        const types = await services.master.getVehicleTypes();
        if (cancelled) return;
        setVehicleTypes(types);

        // The operating area comes from AllocationRules. Any request type
        // carries the same centre and radius, so the first active rule will do.
        const rule = await services.master.getAllocationRule('Staff');
        if (cancelled) return;
        if (rule?.operatingCentreLat !== undefined && rule?.operatingCentreLng !== undefined) {
          setArea({
            centre: { lat: rule.operatingCentreLat, lng: rule.operatingCentreLng },
            radiusKm: rule.operatingRadiusKm || 60,
          });
        }
      } catch (e) {
        if (!cancelled) setError(e);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })().catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [services, homePropertyId]);

  // ---------------------------------------------------------------------------
  // Screen 5 -> 6. Confirmation appears ONLY once a driver has accepted.
  // ---------------------------------------------------------------------------

  React.useEffect(() => {
    if (screen !== 5 || !services || requestItemId === undefined) return undefined;

    let stopped = false;

    const poll = async (): Promise<void> => {
      try {
        const result = await services.requests.getRequestWithTrip(requestItemId);
        if (stopped) return;
        setTracked(result);

        if (result.request.status === 'Allocated' && result.trip) {
          setScreen(6);
        }
      } catch (e) {
        if (!stopped) setError(e);
      }
    };

    poll().catch(() => undefined);
    const handle = window.setInterval(() => {
      poll().catch(() => undefined);
    }, processingPollMs);

    return () => {
      stopped = true;
      window.clearInterval(handle);
    };
  }, [screen, services, requestItemId, processingPollMs]);

  // Keep the tracking screen current.
  React.useEffect(() => {
    if (screen !== 7 || !services || requestItemId === undefined) return undefined;

    let stopped = false;
    const poll = async (): Promise<void> => {
      try {
        const result = await services.requests.getRequestWithTrip(requestItemId);
        if (!stopped) setTracked(result);
      } catch (e) {
        if (!stopped) setError(e);
      }
    };

    poll().catch(() => undefined);
    const handle = window.setInterval(() => {
      poll().catch(() => undefined);
    }, processingPollMs);

    return () => {
      stopped = true;
      window.clearInterval(handle);
    };
  }, [screen, services, requestItemId, processingPollMs]);

  // ---------------------------------------------------------------------------
  // Navigation
  // ---------------------------------------------------------------------------

  const validate = (from: ScreenIndex): string | undefined => {
    if (from === 1) {
      if (!draft.pickup) return 'Choose a pickup point.';
      if (!draft.dropOff) return 'Choose a destination.';
      if (
        draft.pickup.lat === draft.dropOff.lat &&
        draft.pickup.lng === draft.dropOff.lng
      ) {
        return 'The pickup point and the destination are the same place.';
      }
    }
    if (from === 2 && !draft.slot) return 'Choose a time.';
    if (from === 3 && !draft.requestType) return 'Choose a request type.';
    if (from === 4) {
      if (!draft.vehicleType) return 'Choose a vehicle type.';
      if (draft.vehicleType.title === 'Others' && draft.otherVehicleText.trim().length === 0) {
        return 'Describe the vehicle you need.';
      }
      if (draft.requestType === 'Items or Documents') {
        if (!draft.itemDescription) return 'Choose what is being moved.';
      } else if (draft.vehicleType.title === 'Others' && draft.noOfPax.trim().length > 0) {
        // The passenger count is asked for only with "Others", and is
        // OPTIONAL there. Blank is fine; a figure that is given must be real.
        const pax = parseInt(draft.noOfPax, 10);
        if (isNaN(pax) || pax < 1) {
          return 'Enter how many people are travelling as a number of 1 or more, or leave it blank.';
        }
      }
    }
    return undefined;
  };

  const submit = async (): Promise<void> => {
    if (!services || !draft.slot || !draft.pickup || !draft.dropOff) return;
    if (!draft.requestType || !draft.vehicleType) return;

    setSubmitting(true);
    setError(undefined);

    try {
      const isItems = draft.requestType === 'Items or Documents';
      // Passengers and remarks are offered only with "Others". For any other
      // vehicle type they are not on screen, so nothing is sent for them.
      const isOthers = draft.vehicleType.title === 'Others';
      const pax = parseInt(draft.noOfPax, 10);
      const result = await services.dispatch.createRequest({
        requestType: draft.requestType,
        propertyId,
        pickupText: draft.pickup.text,
        pickupPlaceId: draft.pickup.placeId,
        pickupLat: draft.pickup.lat,
        pickupLng: draft.pickup.lng,
        pickupNotes: draft.pickupNotes.trim() || undefined,
        dropOffText: draft.dropOff.text,
        dropOffPlaceId: draft.dropOff.placeId,
        dropOffLat: draft.dropOff.lat,
        dropOffLng: draft.dropOff.lng,
        dropOffNotes: draft.dropOffNotes.trim() || undefined,
        requiredDateTime: draft.slot,
        vehicleTypeId: draft.vehicleType.id,
        noOfPax: !isItems && isOthers && !isNaN(pax) ? pax : undefined,
        otherVehicleText: isOthers ? draft.otherVehicleText.trim() : undefined,
        itemDescription: isItems ? draft.itemDescription : undefined,
        itemNotes: isItems && draft.itemNotes.trim().length > 0 ? draft.itemNotes.trim() : undefined,
        remarks:
          isOthers && draft.remarks.trim().length > 0 ? draft.remarks.trim() : undefined,
      });

      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }

      setRequestItemId(result.data.id);
      setRequestId(result.data.ref);
      setScreen(5);
    } catch (e) {
      setError(e);
    } finally {
      setSubmitting(false);
    }
  };

  const next = (): void => {
    const problem = validate(screen);
    setValidation(problem);
    if (problem) return;

    if (screen === 4) {
      submit().catch(() => undefined);
      return;
    }
    setScreen((s) => (s + 1) as ScreenIndex);
  };

  const back = (): void => {
    setValidation(undefined);
    setScreen((s) => Math.max(1, s - 1) as ScreenIndex);
  };

  const startOver = (): void => {
    setDraft(emptyDraft());
    setRequestItemId(undefined);
    setRequestId('');
    setTracked(undefined);
    setValidation(undefined);
    setScreen(1);
  };

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  if (!services || !user) {
    return (
      <PageShell heading="Book a ride">
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading what="Starting the booking portal" /> : undefined}
      </PageShell>
    );
  }

  if (loading) {
    return (
      <PageShell heading="Book a ride">
        <Loading what="Loading destinations and vehicle types" />
      </PageShell>
    );
  }

  const meta = SCREEN_META[screen];

  return (
    <PageShell
      heading="Book a ride"
      caption={
        screen <= 4
          ? 'Four steps. Everything is a choice from a list, so nothing can be mistyped.'
          : undefined
      }
      actions={
        myBookingsUrl ? (
          <LinkButton href={myBookingsUrl}>
            <Icon name="calendar" size={18} />
            My Bookings
          </LinkButton>
        ) : undefined
      }
    >
      <div className={styles.wizard}>
        <StepIndicator current={screen} />
        <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

        <section className={styles.phone} aria-label={meta.title}>
          <header className={styles.phoneHeader}>
            <span className={styles.phoneHeaderIcon}>
              <Icon name={meta.icon} size={22} />
            </span>
            <h2 className={styles.phoneTitle}>{meta.title}</h2>
          </header>

          <div
            className={
              screen === 1
                ? styles.phoneBodyFlush
                : screen === 5
                  ? styles.phoneBodyPlain
                  : styles.phoneBody
            }
          >
        {screen === 1 ? (
          <ScreenWhere
            draft={draft}
            setDraft={setDraft}
            validation={validation}
            apiKey={tomtomApiKey}
            isActive={screen === 1}
            area={area}
          />
        ) : undefined}

        {screen === 2 ? (
          <ScreenWhen draft={draft} setDraft={setDraft} validation={validation} />
        ) : undefined}

        {screen === 3 ? (
          <ScreenRequestType draft={draft} setDraft={setDraft} validation={validation} />
        ) : undefined}

        {screen === 4 ? (
          <ScreenVehicleType
            vehicleTypes={vehicleTypes}
            draft={draft}
            setDraft={setDraft}
            validation={validation}
          />
        ) : undefined}

        {screen === 5 ? (
          <ScreenProcessing
            requestId={requestId}
            request={tracked?.request}
            onTrack={() => setScreen(7)}
          />
        ) : undefined}

        {screen === 6 && tracked ? (
          <ScreenConfirmed
            result={tracked}
            myBookingsUrl={myBookingsUrl}
            onTrack={() => setScreen(7)}
            onBookAnother={startOver}
          />
        ) : undefined}

        {screen === 7 && tracked ? (
          <ScreenTrack result={tracked} apiKey={tomtomApiKey} onBookAnother={startOver} />
        ) : undefined}

          </div>

          {screen <= 4 ? (
            <div className={styles.phoneFooter}>
              {screen > 1 ? (
                <Button
                  kind="secondary"
                  className={styles.navBack}
                  onClick={back}
                  disabled={submitting}
                >
                  <Icon name="chevronLeft" size={18} />
                  Back
                </Button>
              ) : undefined}
              <Button
                kind="primary"
                className={styles.navNext}
                onClick={next}
                disabled={submitting}
              >
                {submitting ? 'Submitting…' : screen === 4 ? 'Submit Request' : 'Next'}
                {submitting ? undefined : <Icon name="arrowRight" size={18} />}
              </Button>
            </div>
          ) : undefined}
        </section>
      </div>
    </PageShell>
  );
};

// -----------------------------------------------------------------------------
// Step indicator
// -----------------------------------------------------------------------------

const ALL_SCREENS: ScreenIndex[] = [1, 2, 3, 4, 5, 6, 7];

/**
 * The numbered chip from the design reference: the step number, its name and
 * one line about it. The rail beneath shows how far through the seven steps
 * this is, and the whole thing is spelled out for screen readers, so progress
 * is never conveyed by colour alone.
 */
function StepIndicator({ current }: { current: ScreenIndex }): JSX.Element {
  const meta = SCREEN_META[current];

  return (
    <nav className={styles.steps} aria-label="Booking progress">
      <div className={styles.stepChip} aria-current="step">
        <span className={styles.stepNumber} aria-hidden="true">
          {current}
        </span>
        <span className={styles.stepChipText}>
          <span className={styles.srOnly}>{`Step ${current} of ${ALL_SCREENS.length}: `}</span>
          <span className={styles.stepNowTitle}>{meta.step}</span>
          <span className={styles.stepNowCaption}>{meta.caption}</span>
        </span>
      </div>
      <div className={styles.stepRail} aria-hidden="true">
        {ALL_SCREENS.map((n) => (
          <span
            key={n}
            className={[styles.stepSegment, n <= current ? styles.stepSegmentOn : ''].join(' ')}
          />
        ))}
      </div>
    </nav>
  );
}

/** The tick on a selected option. The word is there for screen readers. */
function SelectedMark(): JSX.Element {
  return (
    <span className={styles.checkMark}>
      <Icon name="check" size={13} />
      <span className={styles.srOnly}>Selected</span>
    </span>
  );
}

// -----------------------------------------------------------------------------
// Screen 1 - where
// -----------------------------------------------------------------------------

interface IScreenProps {
  draft: IDraft;
  setDraft: React.Dispatch<React.SetStateAction<IDraft>>;
  validation?: string;
}

function ScreenWhere({
  draft,
  setDraft,
  validation,
  apiKey,
  isActive,
  area,
}: IScreenProps & {
  apiKey: string;
  isActive: boolean;
  area: { centre: { lat: number; lng: number }; radiusKm: number };
}): JSX.Element {
  // A warning, never a block. A pickup outside the operating area can still be
  // booked - it simply routes to the Coordinator instead of broadcasting.
  const farOut =
    draft.pickup !== undefined &&
    distanceKm(draft.pickup.lat, draft.pickup.lng, area.centre.lat, area.centre.lng) >
      area.radiusKm;

  return (
    <div className={styles.screen}>
      <MapPicker
        apiKey={apiKey}
        isActive={isActive}
        centre={area.centre}
        radiusKm={area.radiusKm}
        pickup={draft.pickup}
        dropOff={draft.dropOff}
        pickupNotes={draft.pickupNotes}
        dropOffNotes={draft.dropOffNotes}
        onPickup={(p) => setDraft((d) => ({ ...d, pickup: p }))}
        onDropOff={(p) => setDraft((d) => ({ ...d, dropOff: p }))}
        onPickupNotes={(n) => setDraft((d) => ({ ...d, pickupNotes: n }))}
        onDropOffNotes={(n) => setDraft((d) => ({ ...d, dropOffNotes: n }))}
      />

      {farOut ? (
        <Notice tone="attention">
          This pickup is outside the usual operating area. You can still book
          it - the Transportation Coordinator will assign someone by hand
          rather than it going out to drivers automatically.
        </Notice>
      ) : undefined}

      {validation ? <Notice tone="attention">{validation}</Notice> : undefined}
    </div>
  );
}

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function two(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/** A Gulf calendar day as YYYY-MM-DD, which is what the date picker speaks. */
function gulfDateValue(d: Date): string {
  const { from } = gulfDayBounds(d);
  return formatDate(from).split('/').reverse().join('-');
}

/**
 * The month grid. PRESENTATION ONLY: it hands back the same YYYY-MM-DD value a
 * date field would, and the screen turns that into a day exactly as before.
 */
function MonthCalendar({ selected, today, onPick }: {
  selected: string;
  today: string;
  onPick: (value: string) => void;
}): JSX.Element {
  const monthOf = (value: string): { year: number; month: number } => ({
    year: parseInt(value.slice(0, 4), 10),
    month: parseInt(value.slice(5, 7), 10) - 1,
  });

  const [view, setView] = React.useState(() => monthOf(selected));

  // Follow the selection, so pressing Today brings the grid back to this month.
  React.useEffect(() => {
    setView(monthOf(selected));
  }, [selected]);

  const move = (delta: number): void => {
    setView((v) => {
      const d = new Date(Date.UTC(v.year, v.month + delta, 1));
      return { year: d.getUTCFullYear(), month: d.getUTCMonth() };
    });
  };

  const lead = new Date(Date.UTC(view.year, view.month, 1)).getUTCDay();
  const count = new Date(Date.UTC(view.year, view.month + 1, 0)).getUTCDate();

  const cells: (number | undefined)[] = [];
  for (let i = 0; i < lead; i++) cells.push(undefined);
  for (let d = 1; d <= count; d++) cells.push(d);

  return (
    <div className={styles.calendar} role="group" aria-label="Choose a date">
      <div className={styles.calendarHead}>
        <button
          type="button"
          className={styles.calendarNav}
          onClick={() => move(-1)}
          aria-label="Previous month"
        >
          <Icon name="chevronLeft" size={18} />
        </button>
        <span className={styles.calendarMonth} aria-live="polite">
          {`${MONTH_NAMES[view.month]} ${view.year}`}
        </span>
        <button
          type="button"
          className={styles.calendarNav}
          onClick={() => move(1)}
          aria-label="Next month"
        >
          <Icon name="chevronRight" size={18} />
        </button>
      </div>

      <div className={styles.calendarGrid}>
        {WEEKDAYS.map((w) => (
          <span key={w} className={styles.calendarWeekday} aria-hidden="true">
            {w}
          </span>
        ))}
        {cells.map((d, i) => {
          if (d === undefined) return <span key={`blank-${i}`} />;

          const value = `${view.year}-${two(view.month + 1)}-${two(d)}`;
          const isSelected = value === selected;
          return (
            <button
              key={value}
              type="button"
              className={[
                styles.calendarDay,
                value < today ? styles.calendarDayPast : '',
                value === today ? styles.calendarDayToday : '',
                isSelected ? styles.calendarDaySelected : '',
              ].join(' ')}
              aria-pressed={isSelected}
              aria-label={`${d} ${MONTH_NAMES[view.month]} ${view.year}${value === today ? ', today' : ''}`}
              onClick={() => onPick(value)}
            >
              {d}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function ScreenWhen({ draft, setDraft, validation }: IScreenProps): JSX.Element {
  const slots: ITimeSlot[] = React.useMemo(
    () => timeSlots(draft.day, 30, new Date()),
    [draft.day],
  );

  const dayValue = React.useMemo(() => gulfDateValue(draft.day), [draft.day]);
  const todayValue = gulfDateValue(new Date());

  const pickDay = (value: string): void => {
    // Today's cell does exactly what the Today button does.
    if (value === todayValue) {
      setDraft((d) => ({ ...d, isAdvance: false, day: new Date(), slot: undefined }));
      return;
    }
    const parsed = new Date(`${value}T00:00:00Z`);
    if (!isNaN(parsed.getTime())) {
      setDraft((d) => ({ ...d, isAdvance: true, day: parsed, slot: undefined }));
    }
  };

  return (
    <div className={styles.screen}>
      <div className={styles.toggleRow} role="group" aria-label="Today or a later day">
        <Button
          kind={draft.isAdvance ? 'secondary' : 'primary'}
          className={styles.segment}
          onClick={() =>
            setDraft((d) => ({ ...d, isAdvance: false, day: new Date(), slot: undefined }))
          }
          aria-pressed={!draft.isAdvance}
        >
          Today
        </Button>
        <Button
          kind={draft.isAdvance ? 'primary' : 'secondary'}
          className={styles.segment}
          onClick={() => setDraft((d) => ({ ...d, isAdvance: true, slot: undefined }))}
          aria-pressed={draft.isAdvance}
        >
          Advance
        </Button>
      </div>

      <MonthCalendar selected={dayValue} today={todayValue} onPick={pickDay} />

      <FormRow
        label={`Select time on ${formatDateLong(draft.day)}`}
        required
        error={validation}
        hint="Times are in Gulf Standard Time. Slots that have passed cannot be chosen. A booking made well in advance is held and released to drivers shortly before it is needed."
      >
        <div className={styles.slotGrid} role="group" aria-label="Available time slots">
          {slots.map((slot) => (
            <button
              key={slot.value.toISOString()}
              type="button"
              className={[
                styles.slot,
                draft.slot?.getTime() === slot.value.getTime() ? styles.slotSelected : '',
              ].join(' ')}
              disabled={slot.disabled}
              aria-pressed={draft.slot?.getTime() === slot.value.getTime()}
              onClick={() => setDraft((d) => ({ ...d, slot: slot.value }))}
            >
              {slot.label}
            </button>
          ))}
        </div>
      </FormRow>
    </div>
  );
}

// -----------------------------------------------------------------------------
// Screen 3 - request type
// -----------------------------------------------------------------------------

const REQUEST_TYPE_ICON: Record<RequestType, IconName> = {
  'Client or Client Guest': 'users',
  'Staff': 'user',
  'Items or Documents': 'box',
};

function ScreenRequestType({ draft, setDraft, validation }: IScreenProps): JSX.Element {
  return (
    <div className={styles.screen}>
      <p className={styles.screenCaption}>
        This decides which vehicle types are offered on the next screen.
      </p>

      <div className={styles.typeList} role="group" aria-label="Request type">
        {REQUEST_TYPES.map((type) => {
          const selected = draft.requestType === type;
          return (
            <button
              key={type}
              type="button"
              className={[styles.typeCard, selected ? styles.optionSelected : ''].join(' ')}
              aria-pressed={selected}
              onClick={() =>
                setDraft((d) => ({
                  ...d,
                  requestType: type,
                  // The vehicle choice depends on the request type, so it is
                  // cleared rather than left pointing at a type that may no
                  // longer be offered.
                  vehicleType: undefined,
                }))
              }
            >
              <span className={styles.typeIcon}>
                <Icon name={REQUEST_TYPE_ICON[type]} size={26} />
              </span>
              <span className={styles.typeText}>
                <span className={styles.optionTitle}>{type}</span>
                <span className={styles.optionCaption}>{REQUEST_TYPE_CAPTION[type]}</span>
              </span>
              {selected ? <SelectedMark /> : undefined}
            </button>
          );
        })}
      </div>

      {validation ? (
        <div style={{ marginTop: 12 }}>
          <Notice tone="attention">{validation}</Notice>
        </div>
      ) : undefined}
    </div>
  );
}

// -----------------------------------------------------------------------------
// Screen 4 - vehicle type
// -----------------------------------------------------------------------------

type VehicleKind = 'suv' | 'cargo' | 'chiller' | 'hotCold' | 'animal' | 'others';

/** Which drawing goes on a vehicle tile. Decoration only. */
function vehicleKind(title: string): VehicleKind {
  const t = title.toLowerCase();
  if (t === 'others') return 'others';
  if (t.indexOf('chiller') !== -1) return 'chiller';
  if (t.indexOf('hot') !== -1) return 'hotCold';
  if (t.indexOf('animal') !== -1) return 'animal';
  if (t.indexOf('van') !== -1) return 'cargo';
  return 'suv';
}

/**
 * The picture on a vehicle tile. DRAWN, not photographed: the solution holds
 * no vehicle images, and a drawing of a KIND of vehicle cannot be mistaken for
 * the specific car that will turn up. Each van carries a mark for what it is
 * for, and the tile always names the type in words beneath.
 */
function VehicleArt({ kind }: { kind: VehicleKind }): JSX.Element {
  const wheels = (
    <>
      <circle cx="26" cy="43" r="8" fill="#2E2E2E" />
      <circle cx="26" cy="43" r="3.6" fill="#C9CED3" />
      <circle cx="72" cy="43" r="8" fill="#2E2E2E" />
      <circle cx="72" cy="43" r="3.6" fill="#C9CED3" />
    </>
  );

  if (kind === 'suv') {
    return (
      <svg className={styles.vehicleArt} viewBox="0 0 96 56" aria-hidden="true" focusable="false">
        <ellipse cx="48" cy="51" rx="40" ry="3" fill="#000" opacity="0.1" />
        <path
          d="M7 38l2-10q1-4 8-5l12-11q3-3 8-3h27q5 0 8 4l8 10q9 2 10 7l1 8q0 4-4 4H11q-4 0-4-4z"
          fill="#E3E6EA"
          stroke="#8E959D"
          strokeWidth="1.2"
        />
        <path d="M32 13h14v10H22z" fill="#52616F" />
        <path d="M49 13h14q3 0 5 3l5 7H49z" fill="#52616F" />
        <path d="M9 30h78" stroke="#B9BFC6" strokeWidth="1" />
        <rect x="83" y="26" width="6" height="3" rx="1" fill="#F6D58A" />
        <rect x="7" y="27" width="4" height="3" rx="1" fill="#D9534F" />
        <rect x="53" y="26" width="5" height="1.6" rx="0.8" fill="#8E959D" />
        {wheels}
      </svg>
    );
  }

  return (
    <svg className={styles.vehicleArt} viewBox="0 0 96 56" aria-hidden="true" focusable="false">
      <ellipse cx="48" cy="51" rx="40" ry="3" fill="#000" opacity="0.1" />
      <path d="M6 13q0-4 4-4h49v33H6z" fill="#FFFFFF" stroke="#8E959D" strokeWidth="1.2" />
      <path
        d="M59 17h13q4 0 7 4l9 11q2 2 2 5v5H59z"
        fill="#EEF0F2"
        stroke="#8E959D"
        strokeWidth="1.2"
      />
      <path d="M63 20h9l8 11H63z" fill="#52616F" />
      <rect x="85" y="35" width="5" height="3" rx="1" fill="#F6D58A" />

      {kind === 'cargo' ? (
        <>
          <path d="M6 30h53" stroke="#652D16" strokeWidth="3" />
          <path d="M24 16h16v9H24zM32 16v9" fill="none" stroke="#652D16" strokeWidth="1.6" />
        </>
      ) : undefined}

      {kind === 'chiller' ? (
        <>
          <path d="M6 30h53" stroke="#3B8FD9" strokeWidth="3" />
          <path
            d="M32 12v14M26 15.5l12 7M38 15.5l-12 7"
            stroke="#3B8FD9"
            strokeWidth="2"
            strokeLinecap="round"
          />
        </>
      ) : undefined}

      {kind === 'hotCold' ? (
        <>
          <path d="M6 30h26" stroke="#E0662F" strokeWidth="3" />
          <path d="M32 30h27" stroke="#3B8FD9" strokeWidth="3" />
          <path d="M32 12a7 7 0 0 0 0 14z" fill="#E0662F" />
          <path d="M32 12a7 7 0 0 1 0 14z" fill="#3B8FD9" />
        </>
      ) : undefined}

      {kind === 'animal' ? (
        <>
          <path d="M6 30h53" stroke="#E08A2F" strokeWidth="3" />
          <ellipse cx="32" cy="21.5" rx="4.2" ry="3.4" fill="#E08A2F" />
          <circle cx="26.5" cy="17" r="1.9" fill="#E08A2F" />
          <circle cx="30.5" cy="14.2" r="1.9" fill="#E08A2F" />
          <circle cx="34.5" cy="14.2" r="1.9" fill="#E08A2F" />
          <circle cx="38" cy="17" r="1.9" fill="#E08A2F" />
        </>
      ) : undefined}

      {kind === 'others' ? (
        <>
          <circle cx="32" cy="21" r="9" fill="#FFFFFF" stroke="#E0662F" strokeWidth="2" />
          <path d="M32 16v10M27 21h10" stroke="#E0662F" strokeWidth="2.6" strokeLinecap="round" />
        </>
      ) : undefined}

      {wheels}
    </svg>
  );
}

function ScreenVehicleType({
  vehicleTypes,
  draft,
  setDraft,
  validation,
}: IScreenProps & { vehicleTypes: IVehicleType[] }): JSX.Element {
  const offered = React.useMemo(() => {
    if (!draft.requestType) return [];
    return vehicleTypes
      .filter((t) => t.suitableFor.indexOf(draft.requestType as RequestType) !== -1)
      .sort((a, b) => a.sortOrder - b.sortOrder);
  }, [vehicleTypes, draft.requestType]);

  const isOthers = draft.vehicleType?.title === 'Others';
  const isItems = draft.requestType === 'Items or Documents';

  return (
    <div className={styles.screen}>
      {offered.length === 0 ? (
        <EmptyState title="No vehicle types are set up for this request type">
          Ask a Fleet Admin to check the VehicleTypes list.
        </EmptyState>
      ) : (
        <div className={styles.vehicleGrid} role="group" aria-label="Vehicle type">
          {offered.map((t) => {
            const selected = draft.vehicleType?.id === t.id;
            return (
              <button
                key={t.id}
                type="button"
                className={[styles.vehicleTile, selected ? styles.optionSelected : ''].join(' ')}
                aria-pressed={selected}
                onClick={() => setDraft((d) => ({ ...d, vehicleType: t }))}
              >
                <VehicleArt kind={vehicleKind(t.title)} />
                <span className={styles.vehicleName}>{t.title}</span>
                {t.capacity > 0 ? (
                  <span className={styles.srOnly}>{`, up to ${t.capacity} people`}</span>
                ) : undefined}
                {selected ? <SelectedMark /> : undefined}
              </button>
            );
          })}
        </div>
      )}

      {/* The only required free-text field in the entire form. */}
      {isOthers ? (
        <div style={{ marginTop: 16 }}>
          <FormRow
            label="What kind of vehicle do you need?"
            required
            hint="This is the only free-text field in the form. Be specific - the Coordinator reads it."
          >
            <TextField
              value={draft.otherVehicleText}
              maxLength={255}
              onChange={(e) => setDraft((d) => ({ ...d, otherVehicleText: e.target.value }))}
            />
          </FormRow>
        </div>
      ) : undefined}

      {/*
        What is being moved is always asked of an Items or Documents request.
        The passenger count and the remarks are asked ONLY with "Others", and
        both are optional.
      */}
      {isItems || isOthers ? (
      <div className={styles.vehicleDetails}>
        {isItems ? (
          <>
            <FormRow label="What is being moved?" required>
              <Select
                value={draft.itemDescription ?? ''}
                onChange={(e) =>
                  setDraft((d) => ({
                    ...d,
                    itemDescription: (e.target.value || undefined) as ItemDescription | undefined,
                  }))
                }
              >
                <option value="">Choose one</option>
                {ITEM_DESCRIPTIONS.map((i) => (
                  <option key={i} value={i}>
                    {i}
                  </option>
                ))}
              </Select>
            </FormRow>
            <FormRow label="Anything the driver should know?">
              <TextArea
                value={draft.itemNotes}
                maxLength={500}
                onChange={(e) => setDraft((d) => ({ ...d, itemNotes: e.target.value }))}
              />
            </FormRow>
          </>
        ) : (
          <FormRow label="How many people are travelling?" hint="Optional.">
            <TextField
              type="number"
              min={1}
              max={60}
              value={draft.noOfPax}
              onChange={(e) => setDraft((d) => ({ ...d, noOfPax: e.target.value }))}
            />
          </FormRow>
        )}

        {isOthers ? (
          <FormRow label="Remarks" hint="Optional.">
            <TextArea
              value={draft.remarks}
              maxLength={500}
              onChange={(e) => setDraft((d) => ({ ...d, remarks: e.target.value }))}
            />
          </FormRow>
        ) : undefined}
      </div>
      ) : undefined}

      {validation ? <Notice tone="attention">{validation}</Notice> : undefined}
    </div>
  );
}

// -----------------------------------------------------------------------------
// Screen 5 - processing
// -----------------------------------------------------------------------------

/**
 * ACKNOWLEDGEMENT IS NOT CONFIRMATION, and this screen must not conflate the
 * two. It says the request was received and that confirmation will follow. An
 * e-mail saying the same thing goes out on submit, before any driver has
 * responded.
 */
function ScreenProcessing({ requestId, request, onTrack }: {
  requestId: string;
  request?: ITransportRequest;
  onTrack: () => void;
}): JSX.Element {
  const status = request?.status;
  const stage = status === 'Broadcasting' ? 2 : status === 'Awaiting Coordinator' ? 3 : 1;

  // The fourth line is the same `stage` read one step further on: once the
  // request is with the Coordinator, assigning it is what is under way.
  const steps = [
    { label: 'Request Submitted', done: true, active: false },
    { label: 'Finding Available Vehicle', done: stage > 1, active: stage === 1 },
    { label: 'Matching Driver', done: stage > 2, active: stage === 2 },
    { label: 'Allocating Assignment', done: false, active: stage === 3 },
  ];

  return (
    <div className={styles.screen}>
      <ProcessingArt />

      <ol className={styles.processingSteps}>
        {steps.map((s) => (
          <li
            key={s.label}
            className={[
              styles.processingStep,
              s.active ? styles.processingStepActive : '',
            ].join(' ')}
          >
            <span
              className={[
                styles.processingMark,
                s.done ? styles.processingMarkDone : '',
                s.active ? styles.processingMarkActive : '',
              ].join(' ')}
              aria-hidden="true"
            >
              {s.done ? <Icon name="check" size={12} /> : undefined}
            </span>
            <span className={styles.processingLabel}>{s.label}</span>
            <span className={styles.srOnly}>
              {s.done ? ' — done' : s.active ? ' — in progress' : ' — waiting'}
            </span>
          </li>
        ))}
      </ol>

      <div className={styles.infoBox}>
        <span className={styles.infoBoxIcon} aria-hidden="true">
          <Icon name="clock" size={30} />
        </span>
        <span>
          Your request is being processed. You will receive a confirmation
          shortly.
          <span className={styles.infoBoxRef}>
            Reference <Reference>{requestId}</Reference> &middot; this page
            will update on its own.
          </span>
        </span>
      </div>

      {status === 'Awaiting Coordinator' ? (
        <div style={{ marginTop: 12 }}>
          <Notice tone="attention">
            No driver was free to take this straight away, so it has gone to
            the Transportation Coordinator, who will assign someone. Your
            booking has <strong>not</strong> been lost.
          </Notice>
        </div>
      ) : undefined}

      <div className={styles.actions}>
        <Button kind="secondary" className={styles.actionWide} onClick={onTrack}>
          Track this request
          <Icon name="arrowRight" size={18} />
        </Button>
      </div>
    </div>
  );
}

/**
 * The picture on the processing screen: the request, in the middle, being
 * matched with vehicles and drivers. Decoration only.
 */
function ProcessingArt(): JSX.Element {
  const person = (cx: number, cy: number, colour: string): JSX.Element => (
    <g fill={colour}>
      <circle cx={cx} cy={cy - 4} r="4.4" />
      <path d={`M${cx - 8} ${cy + 9}a8 7 0 0 1 16 0z`} />
    </g>
  );
  const car = (cx: number, cy: number): JSX.Element => (
    <g fill="#1F4E8C">
      <path
        d={`M${cx - 9} ${cy + 1}l2-5q1-2 3-2h8q2 0 3 2l2 5q2 0 2 2v4h-22v-4q0-2 2-2z`}
      />
      <circle cx={cx - 6} cy={cy + 8} r="2.6" fill="#16365F" />
      <circle cx={cx + 6} cy={cy + 8} r="2.6" fill="#16365F" />
      <rect x={cx - 5} y={cy - 4} width="10" height="4" rx="1" fill="#DCEBFA" />
    </g>
  );

  return (
    <svg
      className={styles.processingArt}
      viewBox="0 0 320 232"
      aria-hidden="true"
      focusable="false"
    >
      {/* Links */}
      <g stroke="#9DB7D6" strokeWidth="1.4" strokeDasharray="3 4" fill="none">
        <path d="M160 46V70" />
        <path d="M86 56L144 82" />
        <path d="M234 56L176 82" />
        <path d="M70 100L142 92" />
        <path d="M250 100L178 92" />
        <path d="M160 108V124" />
      </g>

      {/* Cloud */}
      <path
        d="M134 44a13 13 0 0 1 3-25 19 19 0 0 1 36-4 15 15 0 0 1 14 29z"
        fill="#2D8FE3"
      />

      {/* Vehicles and drivers around the request */}
      <circle cx="70" cy="48" r="19" fill="#DCEBFA" />
      {car(70, 46)}
      <circle cx="250" cy="48" r="19" fill="#DCEBFA" />
      {car(250, 46)}
      <circle cx="52" cy="102" r="19" fill="#DCEBFA" />
      {person(52, 102, '#2D8FE3')}
      <circle cx="268" cy="102" r="19" fill="#DCEBFA" />
      {person(268, 102, '#2D8FE3')}

      {/* The request */}
      <circle cx="160" cy="90" r="18" fill="#1F3B63" />
      {person(160, 91, '#FFFFFF')}

      {/* Leaves */}
      <g fill="#7FA9CF">
        <path d="M84 206q-14-22-4-48 12 22 4 48z" />
        <path d="M94 206q-2-22 14-38 4 22-14 38z" opacity="0.75" />
        <path d="M236 206q14-22 4-48-12 22-4 48z" />
        <path d="M226 206q2-22-14-38-4 22 14 38z" opacity="0.75" />
      </g>

      {/* Laptop */}
      <rect x="104" y="124" width="112" height="72" rx="7" fill="#1F3B63" />
      <rect x="110" y="130" width="100" height="60" rx="3" fill="#FFFFFF" />
      <rect x="110" y="130" width="100" height="9" rx="3" fill="#DCEBFA" />
      <g stroke="#C9D9EC" strokeWidth="3" strokeLinecap="round">
        <path d="M118 150h22M118 160h14M118 170h20M182 150h20M188 160h14M184 170h18" />
      </g>
      <path d="M90 196h140l-8 12H98z" fill="#2C4F7F" />
      <rect x="146" y="198" width="28" height="4" rx="2" fill="#9DB7D6" />

      {/* Done */}
      <circle cx="160" cy="163" r="17" fill="#FFFFFF" />
      <circle cx="160" cy="163" r="14" fill="#21B25B" />
      <path
        d="M153 163.5l5 5 9.5-10"
        fill="none"
        stroke="#FFFFFF"
        strokeWidth="3.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />

      {/* Gear */}
      <g transform="translate(226 126)">
        <path
          d="M0-15l3 1 1 4 3 2 4-2 3 3-2 4 2 3 4 1v4l-4 1-2 3 2 4-3 3-4-2-3 2-1 4h-4l-1-4-3-2-4 2-3-3 2-4-2-3-4-1v-4l4-1 2-3-2-4 3-3 4 2 3-2 1-4z"
          transform="translate(0 2) scale(0.9)"
          fill="#1F3B63"
        />
        <circle cx="0" cy="2" r="5" fill="#FFFFFF" />
      </g>
    </svg>
  );
}

// -----------------------------------------------------------------------------
// Screens 6 and 7 - shared parts
// -----------------------------------------------------------------------------

/** One line of the confirmation: an icon, a label, and the value. */
function InfoRow({ icon, label, action, children }: {
  icon: IconName;
  label: string;
  action?: React.ReactNode;
  children: React.ReactNode;
}): JSX.Element {
  return (
    <div className={styles.infoRow}>
      <span className={styles.infoIcon}>
        <Icon name={icon} size={18} />
      </span>
      <span className={styles.infoLabel}>{label}</span>
      <span className={styles.infoValue}>{children}</span>
      {action}
    </div>
  );
}

/** The round call button beside a driver's number. */
function CallButton({ phone, name }: { phone: string; name?: string }): JSX.Element {
  return (
    <a
      className={styles.callButton}
      href={`tel:${phone}`}
      aria-label={`Call ${name ?? 'the driver'}`}
    >
      <Icon name="phone" size={16} />
    </a>
  );
}

// -----------------------------------------------------------------------------
// Screen 6 - confirmed
// -----------------------------------------------------------------------------

/**
 * Appears ONLY once a driver has accepted. Carries everything a requester at a
 * kerbside needs to identify the car that pulls up.
 */
function ScreenConfirmed({ result, myBookingsUrl, onTrack, onBookAnother }: {
  result: IRequestWithTrip;
  myBookingsUrl?: string;
  onTrack: () => void;
  onBookAnother: () => void;
}): JSX.Element {
  const { request } = result;
  const vehicle = [result.vehicleColour, result.vehicleMake, result.vehicleModel]
    .filter(Boolean)
    .join(' ');

  return (
    <div className={styles.screen}>
      <div className={styles.confirmHero}>
        <span className={styles.confirmTick} aria-hidden="true">
          <Icon name="check" size={34} />
        </span>
        <p className={styles.heroHeadline}>Your booking has been confirmed!</p>
        <p className={styles.screenCaption}>
          A driver has accepted. Here is who is coming and what they are driving.
        </p>
      </div>

      <div className={styles.infoList}>
        <InfoRow icon="hash" label="Request ID">
          <Reference>{request.requestId}</Reference>
        </InfoRow>
        <InfoRow icon="calendar" label="Date & Time">
          {formatDateTime(request.requiredDateTime)}
        </InfoRow>
        <InfoRow icon="target" label="From">
          {request.pickupText || '—'}
        </InfoRow>
        <InfoRow icon="pin" label="To">
          {request.dropOffText || '—'}
        </InfoRow>
        <InfoRow icon="doc" label="Type">
          {request.requestType}
        </InfoRow>
        <InfoRow
          icon="user"
          label="Driver"
          action={
            result.driverPhone ? (
              <CallButton phone={result.driverPhone} name={result.driverName} />
            ) : undefined
          }
        >
          {result.driverName ?? '—'}
          <span className={styles.infoSub}>
            {result.driverPhone ? (
              <a href={`tel:${result.driverPhone}`}>{result.driverPhone}</a>
            ) : (
              '—'
            )}
          </span>
        </InfoRow>
        <InfoRow icon="car" label="Vehicle">
          {vehicle || '—'}
          <span className={styles.infoSub}>{`Plate: ${result.vehiclePlate ?? '—'}`}</span>
        </InfoRow>
        <InfoRow icon="car" label="Vehicle type">
          {result.vehicleTypeName ?? '—'}
        </InfoRow>
      </div>

      <div className={styles.actions}>
        <Button kind="primary" className={styles.actionWide} onClick={onTrack}>
          Track my ride
          <Icon name="arrowRight" size={18} />
        </Button>
        {myBookingsUrl ? (
          <LinkButton href={myBookingsUrl} className={styles.actionWide}>
            View My Bookings
            <Icon name="arrowRight" size={18} />
          </LinkButton>
        ) : undefined}
        <Button kind="secondary" className={styles.actionWide} onClick={onBookAnother}>
          Book another ride
        </Button>
      </div>
    </div>
  );
}

// -----------------------------------------------------------------------------
// Screen 7 - track my ride
// -----------------------------------------------------------------------------

/**
 * Phase 1 is status-based. The driver's name, phone and the vehicle's details
 * remain on screen THROUGHOUT, because that is what a requester at a kerbside
 * actually needs in order to identify the car that pulls up.
 */
function ScreenTrack({ result, apiKey, onBookAnother }: {
  result: IRequestWithTrip;
  apiKey: string;
  onBookAnother: () => void;
}): JSX.Element {
  const { request, trip } = result;
  const [showProgress, setShowProgress] = React.useState(false);

  const timestamps: Partial<Record<TripStatus, Date>> = {};
  if (trip) {
    timestamps.Assigned = trip.acceptedOn;
    timestamps['En Route'] = trip.actualStartTime;
    timestamps['In Transit'] = trip.actualPickupTime;
    timestamps.Completed = trip.actualDropTime;
  }
  result.timeline.forEach((h) => {
    const d = h.changedOn;
    if (h.toStatus === 'En Route' && !timestamps['En Route']) timestamps['En Route'] = d;
    if (h.toStatus === 'In Transit' && !timestamps['In Transit']) timestamps['In Transit'] = d;
    if (h.toStatus === 'Completed' && !timestamps.Completed) timestamps.Completed = d;
  });

  const stages = trip ? buildStages(trip.tripStatus, timestamps) : [];
  const now = trip ? stages.filter((s) => s.status === trip.tripStatus)[0] : undefined;
  const vehicle = [result.vehicleMake, result.vehicleModel].filter(Boolean).join(' ');

  // The addresses under the two place names. Looked up from the locations the
  // wizard already holds; nothing further is read.

  return (
    <div className={styles.screen}>
      {!trip ? (
        <EmptyState title="No driver has accepted yet">
          You will be e-mailed as soon as one does.
        </EmptyState>
      ) : (
        <>
          <button
            type="button"
            className={styles.statusBanner}
            aria-expanded={showProgress}
            onClick={() => setShowProgress((v) => !v)}
          >
            <span className={styles.statusBannerIcon} aria-hidden="true">
              <Icon name="car" size={24} />
            </span>
            <span className={styles.statusBannerText}>
              <strong>{now ? now.label : trip.tripStatus}</strong>
              {trip.etaMinutes === undefined ? (
                <span className={styles.statusBannerSub}>
                  {now ? now.description : ''}
                </span>
              ) : (
                <EtaDisplay minutes={trip.etaMinutes} updatedOn={trip.etaUpdatedOn} />
              )}
            </span>
            <span
              className={[styles.statusBannerChevron, showProgress ? styles.chevronOpen : ''].join(' ')}
              aria-hidden="true"
            >
              <Icon name="chevronRight" size={20} />
            </span>
            <span className={styles.srOnly}>
              {showProgress ? 'Hide trip progress' : 'Show trip progress'}
            </span>
          </button>

          {showProgress ? (
            <div className={styles.whiteCard}>
              <Timeline stages={stages} current={trip.tripStatus} />
            </div>
          ) : undefined}

          {/*
            A real map where we have coordinates, the illustration where we do
            not - requests booked before the map change have no lat/lng, and a
            map centred on nothing is worse than a drawing.
          */}
          {request.pickupLat !== undefined &&
          request.pickupLng !== undefined &&
          request.dropOffLat !== undefined &&
          request.dropOffLng !== undefined ? (
            <RouteMap
              apiKey={apiKey}
              isActive
              pickupLat={request.pickupLat}
              pickupLng={request.pickupLng}
              pickupText={request.pickupText || '—'}
              dropOffLat={request.dropOffLat}
              dropOffLng={request.dropOffLng}
              dropOffText={request.dropOffText || '—'}
            />
          ) : (
            <TrackMap
              tripStatus={trip.tripStatus}
              fromTitle={request.pickupText || '—'}
              fromAddress={request.pickupNotes}
              toTitle={request.dropOffText || '—'}
              toAddress={request.dropOffNotes}
            />
          )}

          <div className={styles.rideCard}>
            <div className={styles.rideRow}>
              <span className={styles.driverAvatar} aria-hidden="true">
                <Icon name="user" size={28} />
              </span>
              <span className={styles.driverText}>
                <strong>{result.driverName ?? '—'}</strong>
                <span className={styles.infoSub}>
                  {result.driverPhone ? (
                    <a href={`tel:${result.driverPhone}`}>{result.driverPhone}</a>
                  ) : (
                    '—'
                  )}
                </span>
              </span>
              {result.driverPhone ? (
                <CallButton phone={result.driverPhone} name={result.driverName} />
              ) : undefined}
            </div>
            <div className={styles.rideRow}>
              <span className={styles.rideVehicleArt}>
                <VehicleArt kind={vehicleKind(result.vehicleTypeName ?? '')} />
              </span>
              <span className={styles.driverText}>
                <strong>
                  {vehicle || '—'}
                  {result.vehicleColour ? ` (${result.vehicleColour})` : ''}
                </strong>
                <span className={styles.infoSub}>{`Plate: ${result.vehiclePlate ?? '—'}`}</span>
              </span>
            </div>
          </div>
        </>
      )}

      <div className={styles.trackMeta}>
        <span>
          Reference <Reference>{request.requestId}</Reference>
        </span>
        <Badge state={requestState(request.status)} label={request.status} />
      </div>

      <div className={styles.actions}>
        <Button kind="secondary" className={styles.actionWide} onClick={onBookAnother}>
          Book another ride
        </Button>
      </div>
    </div>
  );
}

/**
 * The route picture: where the trip starts, where it ends, and a car placed by
 * the trip's STATUS. It is a diagram, not a live map - phase 1 has no
 * coordinates, so the car shows how far the trip has got, not where the
 * vehicle is. Phase 2 drops a live map into this region.
 */
function TrackMap({ tripStatus, fromTitle, fromAddress, toTitle, toAddress }: {
  tripStatus: TripStatus;
  fromTitle: string;
  fromAddress?: string;
  toTitle: string;
  toAddress?: string;
}): JSX.Element {
  // Start and end of the drawn route, as percentages of the picture.
  const from = { x: 16, y: 80 };
  const to = { x: 40, y: 22 };

  // How far along the line the car sits, by status.
  const progress =
    tripStatus === 'Completed' ? 1 : tripStatus === 'In Transit' ? 0.5 : tripStatus === 'En Route' ? 0.12 : 0;
  const carX = from.x + (to.x - from.x) * progress;
  const carY = from.y + (to.y - from.y) * progress;

  return (
    <div className={styles.trackMap}>
      <svg
        className={styles.trackRoute}
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        aria-hidden="true"
        focusable="false"
      >
        <line
          x1={from.x}
          y1={from.y}
          x2={to.x}
          y2={to.y}
          stroke="#2F6FE4"
          strokeWidth="5"
          strokeLinecap="round"
          vectorEffect="non-scaling-stroke"
        />
      </svg>

      <span className={styles.trackPin} style={{ left: `${to.x}%`, top: `${to.y}%` }} aria-hidden="true">
        <Icon name="pin" size={34} />
      </span>
      <span
        className={styles.trackOrigin}
        style={{ left: `${from.x}%`, top: `${from.y}%` }}
        aria-hidden="true"
      />
      <span
        className={styles.trackCar}
        style={{ left: `${carX}%`, top: `${carY}%` }}
        aria-hidden="true"
      >
        <Icon name="car" size={20} />
      </span>

      <div className={styles.trackLabel} style={{ left: `${to.x + 9}%`, top: `${to.y - 10}%` }}>
        <span className={styles.srOnly}>Going to: </span>
        <strong>{toTitle}</strong>
        {toAddress ? <span>{toAddress}</span> : undefined}
      </div>
      <div className={styles.trackLabel} style={{ left: `${from.x + 9}%`, top: `${from.y - 9}%` }}>
        <span className={styles.srOnly}>From: </span>
        <strong>{fromTitle}</strong>
        {fromAddress ? <span>{fromAddress}</span> : undefined}
      </div>
    </div>
  );
}

export default BookARide;
