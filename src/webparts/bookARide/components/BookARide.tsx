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
import {
  Badge,
  Button,
  Card,
  CardTitleRow,
  EmptyState,
  ErrorBanner,
  FieldRow,
  FormRow,
  Loading,
  Notice,
  PageShell,
  Panel,
  Reference,
  Select,
  TextArea,
  TextField,
  Tile,
  TileGrid,
} from '../../../common/components/Shared';
import { buildStages, EtaDisplay, Timeline } from '../../../common/components/Timeline';
import {
  ILocation,
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
import styles from './BookARide.module.scss';

export interface IBookARideProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  /** Property to book against. Defaults to the requester's own property. */
  defaultPropertyId: number;
  /** How long screen 5 waits for a driver before offering to carry on. */
  processingPollMs: number;
}

type ScreenIndex = 1 | 2 | 3 | 4 | 5 | 6 | 7;

const SCREEN_NAMES: Record<ScreenIndex, string> = {
  1: 'Where',
  2: 'When',
  3: 'Request type',
  4: 'Vehicle',
  5: 'Processing',
  6: 'Confirmed',
  7: 'Track',
};

const REQUEST_TYPE_CAPTION: Record<RequestType, string> = {
  'Client or Client Guest': 'A client or their guest is travelling.',
  'Staff': 'A member of staff is travelling.',
  'Items or Documents': 'Something is being moved, not someone.',
};

interface IDraft {
  pickup?: ILocation;
  dropOff?: ILocation;
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
  isAdvance: false,
  day: new Date(),
  otherVehicleText: '',
  noOfPax: '1',
  itemNotes: '',
  remarks: '',
});

export const BookARide: React.FC<IBookARideProps> = (props) => {
  const { services, user, bootstrapError, defaultPropertyId, processingPollMs } = props;

  const [screen, setScreen] = React.useState<ScreenIndex>(1);
  const [draft, setDraft] = React.useState<IDraft>(emptyDraft());
  const [locations, setLocations] = React.useState<ILocation[]>([]);
  const [vehicleTypes, setVehicleTypes] = React.useState<IVehicleType[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [submitting, setSubmitting] = React.useState(false);
  const [error, setError] = React.useState<unknown>(bootstrapError);
  const [validation, setValidation] = React.useState<string | undefined>();

  const [requestItemId, setRequestItemId] = React.useState<number | undefined>();
  const [requestId, setRequestId] = React.useState<string>('');
  const [tracked, setTracked] = React.useState<IRequestWithTrip | undefined>();

  /**
   * The requester's own property, if the site tells us one. Used ONLY to
   * default the pickup point. It is a stable number, so it never causes the
   * locations below to reload.
   */
  const homePropertyId = React.useMemo(
    () => defaultPropertyId || user?.propertyIds[0] || 0,
    [defaultPropertyId, user],
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
  const propertyId = draft.pickup?.property?.id ?? homePropertyId;

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
        const [locs, types] = await Promise.all([
          // EVERY location, unfiltered. Both pickers read this one list.
          services.master.getLocations(),
          services.master.getVehicleTypes(),
        ]);
        if (cancelled) return;
        setLocations(locs);
        setVehicleTypes(types);

        // Pickup DEFAULTS to the requester's own property - a starting point,
        // not a restriction. They can pick any location in the list.
        if (homePropertyId) {
          const home = locs.filter((l) => l.property?.id === homePropertyId)[0];
          if (home) setDraft((d) => ({ ...d, pickup: d.pickup ?? home }));
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
      if (draft.pickup.id === draft.dropOff.id) {
        return 'The pickup point and the destination are the same.';
      }
    }
    if (from === 2 && !draft.slot) return 'Choose a time.';
    if (from === 3 && !draft.requestType) return 'Choose a request type.';
    if (from === 4) {
      if (!draft.vehicleType) return 'Choose a vehicle type.';
      if (draft.vehicleType.title === 'Others' && draft.otherVehicleText.trim().length === 0) {
        return 'Describe the vehicle you need.';
      }
      if (draft.requestType !== 'Items or Documents') {
        const pax = parseInt(draft.noOfPax, 10);
        if (isNaN(pax) || pax < 1) return 'Enter how many people are travelling.';
      } else if (!draft.itemDescription) {
        return 'Choose what is being moved.';
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
      const result = await services.dispatch.createRequest({
        requestType: draft.requestType,
        propertyId,
        pickupLocationId: draft.pickup.id,
        dropOffLocationId: draft.dropOff.id,
        requiredDateTime: draft.slot,
        vehicleTypeId: draft.vehicleType.id,
        noOfPax: isItems ? undefined : parseInt(draft.noOfPax, 10),
        otherVehicleText:
          draft.vehicleType.title === 'Others' ? draft.otherVehicleText.trim() : undefined,
        itemDescription: isItems ? draft.itemDescription : undefined,
        itemNotes: isItems && draft.itemNotes.trim().length > 0 ? draft.itemNotes.trim() : undefined,
        remarks: draft.remarks.trim().length > 0 ? draft.remarks.trim() : undefined,
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

  return (
    <PageShell
      heading="Book a ride"
      caption={
        screen <= 4
          ? 'Four steps. Everything is a choice from a list, so nothing can be mistyped.'
          : undefined
      }
    >
      <div className={styles.wizard}>
        <StepIndicator current={screen} />
        <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

        {screen === 1 ? (
          <ScreenWhere
            locations={locations}
            draft={draft}
            setDraft={setDraft}
            validation={validation}
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
            onTrack={() => setScreen(7)}
            onBookAnother={startOver}
          />
        ) : undefined}

        {screen === 7 && tracked ? (
          <ScreenTrack result={tracked} onBookAnother={startOver} />
        ) : undefined}

        {screen <= 4 ? (
          <div className={styles.navRow}>
            <Button kind="secondary" onClick={back} disabled={screen === 1 || submitting}>
              Back
            </Button>
            <Button kind="primary" onClick={next} disabled={submitting}>
              {submitting ? 'Submitting…' : screen === 4 ? 'Submit request' : 'Continue'}
            </Button>
          </div>
        ) : undefined}
      </div>
    </PageShell>
  );
};

// -----------------------------------------------------------------------------
// Step indicator
// -----------------------------------------------------------------------------

function StepIndicator({ current }: { current: ScreenIndex }): JSX.Element {
  return (
    <nav className={styles.steps} aria-label="Booking progress">
      {([1, 2, 3, 4] as ScreenIndex[]).map((n, i) => {
        const done = current > n;
        const isCurrent = current === n;
        return (
          <React.Fragment key={n}>
            {i > 0 ? (
              <span className={styles.stepSeparator} aria-hidden="true">
                &mdash;
              </span>
            ) : undefined}
            <span
              className={[
                styles.step,
                isCurrent ? styles.stepCurrent : '',
                done ? styles.stepDone : '',
              ].join(' ')}
              aria-current={isCurrent ? 'step' : undefined}
            >
              <span className={styles.stepDot} aria-hidden="true">
                {done ? '✓' : n}
              </span>
              {`Step ${n}: ${SCREEN_NAMES[n]}`}
              {done ? ' (done)' : ''}
            </span>
          </React.Fragment>
        );
      })}
    </nav>
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
  locations,
  draft,
  setDraft,
  validation,
}: IScreenProps & { locations: ILocation[] }): JSX.Element {
  return (
    <div className={styles.screen}>
      <h2 className={styles.screenHeading}>Where are you going?</h2>
      <span className={styles.screenCaption}>
        Pickup starts at your own property. Both points are chosen from the list
        of known locations.
      </span>

      <FormRow label="Pickup point" required>
        <LocationPicker
          locations={locations}
          selected={draft.pickup}
          onSelect={(l) => setDraft((d) => ({ ...d, pickup: l }))}
          label="pickup point"
        />
      </FormRow>

      <FormRow
        label="Destination"
        required
        error={validation}
        hint="Type to search. If your destination is not listed, ask a Fleet Admin to add it."
      >
        <LocationPicker
          locations={locations}
          selected={draft.dropOff}
          onSelect={(l) => setDraft((d) => ({ ...d, dropOff: l }))}
          label="destination"
        />
      </FormRow>
    </div>
  );
}

/**
 * A searchable picker over Locations, NEVER a text box. What the requester
 * types filters the list; it never becomes the answer.
 */
function LocationPicker({ locations, selected, onSelect, label }: {
  locations: ILocation[];
  selected?: ILocation;
  onSelect: (location: ILocation) => void;
  label: string;
}): JSX.Element {
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState(false);

  const matches = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    const pool = locations.filter((l) => l.isActive);
    if (q.length === 0) return pool.slice(0, 20);
    return pool
      .filter(
        (l) =>
          l.title.toLowerCase().indexOf(q) !== -1 ||
          (l.address ?? '').toLowerCase().indexOf(q) !== -1,
      )
      .slice(0, 20);
  }, [locations, query]);

  return (
    <div className={styles.pickerWrap}>
      <TextField
        value={query}
        placeholder={selected ? selected.title : `Search for a ${label}`}
        onChange={(e) => {
          setQuery(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        aria-label={`Search for a ${label}`}
        aria-expanded={open}
      />

      {selected && !open ? (
        <div className={styles.pickerMeta} style={{ marginTop: 6 }}>
          Selected: <strong>{selected.title}</strong>
          {selected.address ? ` · ${selected.address}` : ''}
        </div>
      ) : undefined}

      {open ? (
        <ul className={styles.pickerList} aria-label={`Matching ${label}s`}>
          {matches.length === 0 ? (
            <li className={styles.pickerEmpty}>
              No location matches that. Try a shorter search.
            </li>
          ) : (
            matches.map((l) => (
              <li key={l.id}>
                <button
                  type="button"
                  className={[
                    styles.pickerOption,
                    selected?.id === l.id ? styles.pickerOptionSelected : '',
                  ].join(' ')}
                  onClick={() => {
                    onSelect(l);
                    setQuery('');
                    setOpen(false);
                  }}
                >
                  {l.title}
                  {l.address ? <span className={styles.pickerMeta}>{l.address}</span> : undefined}
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
// Screen 2 - when
// -----------------------------------------------------------------------------

function ScreenWhen({ draft, setDraft, validation }: IScreenProps): JSX.Element {
  const slots: ITimeSlot[] = React.useMemo(
    () => timeSlots(draft.day, 30, new Date()),
    [draft.day],
  );

  const dayValue = React.useMemo(() => {
    const { from } = gulfDayBounds(draft.day);
    return formatDate(from).split('/').reverse().join('-');
  }, [draft.day]);

  return (
    <div className={styles.screen}>
      <h2 className={styles.screenHeading}>When do you need it?</h2>
      <span className={styles.screenCaption}>
        Times are shown in Gulf Standard Time. Slots that have already passed
        cannot be chosen.
      </span>

      <div className={styles.toggleRow}>
        <Button
          kind={draft.isAdvance ? 'secondary' : 'primary'}
          onClick={() =>
            setDraft((d) => ({ ...d, isAdvance: false, day: new Date(), slot: undefined }))
          }
          aria-pressed={!draft.isAdvance}
        >
          Today
        </Button>
        <Button
          kind={draft.isAdvance ? 'primary' : 'secondary'}
          onClick={() => setDraft((d) => ({ ...d, isAdvance: true, slot: undefined }))}
          aria-pressed={draft.isAdvance}
        >
          Another day
        </Button>
      </div>

      {draft.isAdvance ? (
        <FormRow label="Date" required htmlFor="booking-date">
          <TextField
            id="booking-date"
            type="date"
            value={dayValue}
            onChange={(e) => {
              const parsed = new Date(`${e.target.value}T00:00:00Z`);
              if (!isNaN(parsed.getTime())) {
                setDraft((d) => ({ ...d, day: parsed, slot: undefined }));
              }
            }}
          />
        </FormRow>
      ) : undefined}

      <FormRow
        label={`Time on ${formatDateLong(draft.day)}`}
        required
        error={validation}
        hint="A booking made well in advance is held and released to drivers shortly before it is needed."
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

function ScreenRequestType({ draft, setDraft, validation }: IScreenProps): JSX.Element {
  return (
    <div className={styles.screen}>
      <h2 className={styles.screenHeading}>What kind of request is this?</h2>
      <span className={styles.screenCaption}>
        This decides which vehicle types are offered on the next screen.
      </span>

      <TileGrid label="Request type">
        {REQUEST_TYPES.map((type) => (
          <Tile
            key={type}
            title={type}
            caption={REQUEST_TYPE_CAPTION[type]}
            selected={draft.requestType === type}
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
          />
        ))}
      </TileGrid>

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
      <h2 className={styles.screenHeading}>What kind of vehicle?</h2>
      <span className={styles.screenCaption}>
        You choose a vehicle type. The specific vehicle and driver are assigned
        for you.
      </span>

      {offered.length === 0 ? (
        <EmptyState title="No vehicle types are set up for this request type">
          Ask a Fleet Admin to check the VehicleTypes list.
        </EmptyState>
      ) : (
        <TileGrid label="Vehicle type">
          {offered.map((t) => (
            <Tile
              key={t.id}
              title={t.title}
              caption={t.capacity > 0 ? `Up to ${t.capacity} people` : undefined}
              selected={draft.vehicleType?.id === t.id}
              onClick={() => setDraft((d) => ({ ...d, vehicleType: t }))}
            />
          ))}
        </TileGrid>
      )}

      {/* The only free-text field in the entire form. */}
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

      <div style={{ marginTop: 16 }}>
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
          <FormRow
            label="How many people are travelling?"
            required
            hint="Checked against the capacity of the vehicle that is assigned."
          >
            <TextField
              type="number"
              min={1}
              max={60}
              value={draft.noOfPax}
              onChange={(e) => setDraft((d) => ({ ...d, noOfPax: e.target.value }))}
            />
          </FormRow>
        )}

        <FormRow label="Remarks">
          <TextArea
            value={draft.remarks}
            maxLength={500}
            onChange={(e) => setDraft((d) => ({ ...d, remarks: e.target.value }))}
          />
        </FormRow>
      </div>

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

  const steps = [
    { label: 'Request received', done: true },
    { label: 'Finding a vehicle', done: stage > 1, active: stage === 1 },
    { label: 'Matching a driver', done: stage > 2, active: stage === 2 },
  ];

  return (
    <div className={styles.screen}>
      <div className={styles.processing}>
        <h2 className={styles.screenHeading}>Your request has been received</h2>
        <span className={styles.screenCaption}>
          Reference <Reference>{requestId}</Reference>. We have e-mailed you to
          confirm we have it. You will get a second e-mail once a driver has
          accepted &mdash; this page will update on its own.
        </span>

        <ol className={styles.processingSteps}>
          {steps.map((s) => (
            <li
              key={s.label}
              className={[
                styles.processingStep,
                s.done ? styles.processingStepDone : '',
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
                {s.done ? '✓' : ''}
              </span>
              {s.label}
              {s.done ? ' — done' : s.active ? ' — in progress' : ' — waiting'}
            </li>
          ))}
        </ol>

        {status === 'Awaiting Coordinator' ? (
          <div style={{ marginTop: 20 }}>
            <Notice tone="attention">
              No driver was free to take this straight away, so it has gone to
              the Transportation Coordinator, who will assign someone. Your
              booking has <strong>not</strong> been lost.
            </Notice>
          </div>
        ) : undefined}

        <div style={{ marginTop: 20 }}>
          <Button kind="secondary" onClick={onTrack}>
            Track this request
          </Button>
        </div>
      </div>
    </div>
  );
}

// -----------------------------------------------------------------------------
// Screen 6 - confirmed
// -----------------------------------------------------------------------------

/**
 * Appears ONLY once a driver has accepted. Carries everything a requester at a
 * kerbside needs to identify the car that pulls up.
 */
function ScreenConfirmed({ result, onTrack, onBookAnother }: {
  result: IRequestWithTrip;
  onTrack: () => void;
  onBookAnother: () => void;
}): JSX.Element {
  const { request } = result;

  return (
    <div className={styles.screen}>
      <h2 className={styles.screenHeading}>Booking confirmed</h2>
      <span className={styles.screenCaption}>
        A driver has accepted. Here is who is coming and what they are driving.
      </span>

      <div className={styles.confirmGrid}>
        <Card>
          <CardTitleRow title="Your trip">
            <Reference>{request.requestId}</Reference>
          </CardTitleRow>
          <FieldRow label="When">{formatDateTime(request.requiredDateTime)}</FieldRow>
          <FieldRow label="From">{request.pickupLocation?.title ?? '—'}</FieldRow>
          <FieldRow label="To">{request.dropOffLocation?.title ?? '—'}</FieldRow>
          <FieldRow label="Type">{request.requestType}</FieldRow>
        </Card>

        <Card>
          <CardTitleRow title="Your driver" />
          <FieldRow label="Name">{result.driverName ?? '—'}</FieldRow>
          <FieldRow label="Phone">
            {result.driverPhone ? (
              <a href={`tel:${result.driverPhone}`}>{result.driverPhone}</a>
            ) : (
              '—'
            )}
          </FieldRow>
          <FieldRow label="Vehicle">
            {[result.vehicleColour, result.vehicleMake, result.vehicleModel]
              .filter(Boolean)
              .join(' ') || '—'}
          </FieldRow>
          <FieldRow label="Plate">{result.vehiclePlate ?? '—'}</FieldRow>
          <FieldRow label="Vehicle type">{result.vehicleTypeName ?? '—'}</FieldRow>
        </Card>
      </div>

      <div className={styles.navRow}>
        <Button kind="secondary" onClick={onBookAnother}>
          Book another ride
        </Button>
        <Button kind="primary" onClick={onTrack}>
          Track my ride
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
function ScreenTrack({ result, onBookAnother }: {
  result: IRequestWithTrip;
  onBookAnother: () => void;
}): JSX.Element {
  const { request, trip } = result;

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

  return (
    <div className={styles.screen}>
      <h2 className={styles.screenHeading}>Track my ride</h2>
      <span className={styles.screenCaption}>
        Reference <Reference>{request.requestId}</Reference> &middot;{' '}
        <Badge state={requestState(request.status)} label={request.status} />
      </span>

      {!trip ? (
        <EmptyState title="No driver has accepted yet">
          You will be e-mailed as soon as one does.
        </EmptyState>
      ) : (
        <>
          <div className={styles.confirmGrid}>
            <Panel plain>
              <Timeline
                stages={buildStages(trip.tripStatus, timestamps)}
                current={trip.tripStatus}
              />
              <EtaDisplay minutes={trip.etaMinutes} updatedOn={trip.etaUpdatedOn} />
            </Panel>

            <Card>
              <CardTitleRow title="Your driver and vehicle" />
              <FieldRow label="Driver">{result.driverName ?? '—'}</FieldRow>
              <FieldRow label="Phone">
                {result.driverPhone ? (
                  <a href={`tel:${result.driverPhone}`}>{result.driverPhone}</a>
                ) : (
                  '—'
                )}
              </FieldRow>
              <FieldRow label="Vehicle">
                {[result.vehicleColour, result.vehicleMake, result.vehicleModel]
                  .filter(Boolean)
                  .join(' ') || '—'}
              </FieldRow>
              <FieldRow label="Plate">{result.vehiclePlate ?? '—'}</FieldRow>
              <FieldRow label="Going to">{request.dropOffLocation?.title ?? '—'}</FieldRow>
            </Card>
          </div>

          {/*
            Phase 2 drops a map into this region. The LiveTrackingEnabled flag
            and the Latitude/Longitude columns are already in place, so no
            interface change is needed when it arrives.
          */}
        </>
      )}

      <div className={styles.navRow}>
        <span />
        <Button kind="secondary" onClick={onBookAnother}>
          Book another ride
        </Button>
      </div>
    </div>
  );
}

export default BookARide;
