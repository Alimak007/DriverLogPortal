/**
 * Driver Dashboard.
 *
 * THIS IS THE DESIGN'S WEAKEST LINK. Drivers have Microsoft 365 accounts but
 * no Teams, so there is no push notification. A broadcast offer is only useful
 * if the driver sees it inside the five-minute window. Three mechanisms work
 * together, and ALL THREE live in this component:
 *
 *   1. Poll every 20 seconds, with an unmissable banner carrying a count of
 *      open offers, plus an audible alert the driver can enable.
 *   2. (The e-mail with a deep link is sent by the Broadcast flow.)
 *   3. The PAGE TITLE updates with the open-offer count, so a backgrounded
 *      browser tab still shows it.
 *
 * A driver who is mid-trip or away from their phone will still miss the
 * window. That is acceptable by design, because the escalation path sends the
 * request to the Coordinator rather than losing it.
 */

import * as React from 'react';
import { ConfirmDialog } from '../../../common/components/ConfirmDialog';
import {
  Badge,
  Button,
  EmptyState,
  ErrorBanner,
  FieldRow,
  FormRow,
  Loading,
  Notice,
  PageShell,
  Reference,
  SectionHeading,
  TextField,
} from '../../../common/components/Shared';
import {
  IOfferCard,
  ITripCard,
  REJECTION_REASONS,
  RejectionReason,
  TripStatus,
  tripState,
} from '../../../common/models';
import {
  nextTripStatus,
  odometerPromptFor,
  TAP_LABEL,
} from '../../../common/rules/statusTransitions';
import { ICurrentUser, IServices } from '../../../common/services/IServices';
import { formatCountdown, formatDateTime, formatTime } from '../../../common/utils/dates';
import styles from './DriverDashboard.module.scss';

export interface IDriverDashboardProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  /** One request per 20 seconds per client is the stated ceiling. */
  pollMs: number;
}

/** A short two-tone chime, synthesised so the bundle carries no audio file. */
function playAlert(): void {
  try {
    const Ctor =
      (window as unknown as { AudioContext?: typeof AudioContext }).AudioContext ??
      (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
    if (!Ctor) return;

    const ctx = new Ctor();
    const beep = (at: number, frequency: number): void => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.frequency.value = frequency;
      osc.type = 'sine';
      gain.gain.setValueAtTime(0.0001, ctx.currentTime + at);
      gain.gain.exponentialRampToValueAtTime(0.25, ctx.currentTime + at + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + at + 0.28);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(ctx.currentTime + at);
      osc.stop(ctx.currentTime + at + 0.3);
    };
    beep(0, 880);
    beep(0.32, 1180);
    window.setTimeout(() => {
      ctx.close().catch(() => undefined);
    }, 1200);
  } catch {
    // An alert that cannot sound is not a reason to break the dashboard.
  }
}

const ALERT_KEY = 'fleet-portal-audible-alert';

export const DriverDashboard: React.FC<IDriverDashboardProps> = ({
  services,
  user,
  bootstrapError,
  pollMs,
}) => {
  const [offers, setOffers] = React.useState<IOfferCard[]>([]);
  const [trips, setTrips] = React.useState<ITripCard[]>([]);
  const [completed, setCompleted] = React.useState<ITripCard[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<unknown>(bootstrapError);
  const [busyId, setBusyId] = React.useState<number | undefined>();
  const [takenMessage, setTakenMessage] = React.useState<string | undefined>();
  const [tick, setTick] = React.useState(0);

  const [alertOn, setAlertOn] = React.useState<boolean>(() => {
    try {
      return window.localStorage.getItem(ALERT_KEY) === 'on';
    } catch {
      return false;
    }
  });

  const [rejectTarget, setRejectTarget] = React.useState<IOfferCard | undefined>();
  const [cancelTarget, setCancelTarget] = React.useState<ITripCard | undefined>();
  const [odometer, setOdometer] = React.useState<Record<number, string>>({});

  const driverId = user?.driverId;
  const previousOfferCount = React.useRef(0);
  const originalTitle = React.useRef<string>('');

  // ---------------------------------------------------------------------------
  // Poll
  // ---------------------------------------------------------------------------

  const load = React.useCallback(async (): Promise<void> => {
    if (!services || driverId === undefined) return;
    try {
      const [open, mine, done] = await Promise.all([
        services.offers.getOpenOffers(driverId),
        services.trips.getMyTripsToday(driverId),
        services.trips.getCompletedToday(driverId),
      ]);
      setOffers(open);
      setTrips(mine);
      setCompleted(done);
    } catch (e) {
      setError(e);
    } finally {
      setLoading(false);
    }
  }, [services, driverId]);

  React.useEffect(() => {
    if (!services || driverId === undefined) {
      setLoading(false);
      return undefined;
    }
    load().catch(() => undefined);
    const handle = window.setInterval(() => {
      load().catch(() => undefined);
    }, pollMs);
    return () => window.clearInterval(handle);
  }, [services, driverId, load, pollMs]);

  // One-second tick so the countdown on each offer card is live between polls.
  React.useEffect(() => {
    if (offers.length === 0) return undefined;
    const handle = window.setInterval(() => setTick((t) => t + 1), 1000);
    return () => window.clearInterval(handle);
  }, [offers.length]);

  // Mechanism 3: the page title carries the count, so a backgrounded tab still
  // shows it.
  React.useEffect(() => {
    if (!originalTitle.current) originalTitle.current = document.title;
    document.title =
      offers.length > 0
        ? `(${offers.length}) New ride offer${offers.length === 1 ? '' : 's'}`
        : originalTitle.current;

    return () => {
      if (originalTitle.current) document.title = originalTitle.current;
    };
  }, [offers.length]);

  // Mechanism 1: the audible alert, on a rising edge only, and only if the
  // driver has switched it on.
  React.useEffect(() => {
    if (alertOn && offers.length > previousOfferCount.current) playAlert();
    previousOfferCount.current = offers.length;
  }, [offers.length, alertOn]);

  const toggleAlert = (): void => {
    const next = !alertOn;
    setAlertOn(next);
    try {
      window.localStorage.setItem(ALERT_KEY, next ? 'on' : 'off');
    } catch {
      // A browser that refuses storage still gets the alert for this session.
    }
    if (next) playAlert();
  };

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------

  const accept = async (card: IOfferCard): Promise<void> => {
    if (!services) return;
    setBusyId(card.offer.id);
    setError(undefined);
    setTakenMessage(undefined);

    try {
      const result = await services.dispatch.acceptOffer(card.offer.id);

      if (result.ok) {
        await load();
        return;
      }

      // LOSING IS EXPECTED, NOT AN ERROR. Every driver but one loses, every
      // time. Show a plain notice and take the card away - never a red error
      // banner, or drivers stop trusting the system within a week.
      if (result.code === 'TAKEN' || result.code === 'NOT_OPEN' || result.code === 'NO_VEHICLE') {
        setTakenMessage(result.message);
        setOffers((current) => current.filter((o) => o.offer.id !== card.offer.id));
        return;
      }

      setError(new Error(result.message));
    } catch (e) {
      setError(e);
    } finally {
      setBusyId(undefined);
    }
  };

  const reject = async (reason?: string): Promise<void> => {
    if (!services || !rejectTarget || !reason) return;
    const [choice, notes] = reason.split(':');
    setBusyId(rejectTarget.offer.id);
    try {
      const result = await services.dispatch.rejectOffer(
        rejectTarget.offer.id,
        choice.trim() as RejectionReason,
        notes ? notes.trim() : undefined,
      );
      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }
      setRejectTarget(undefined);
      await load();
    } catch (e) {
      setError(e);
    } finally {
      setBusyId(undefined);
    }
  };

  const advance = async (card: ITripCard, to: TripStatus): Promise<void> => {
    if (!services) return;
    const prompt = odometerPromptFor(to);
    const raw = odometer[card.trip.id];
    const reading = prompt ? parseInt(raw ?? '', 10) : undefined;

    if (prompt && (raw === undefined || raw.trim().length === 0 || isNaN(reading as number))) {
      setError(
        new Error(
          `Enter the ${prompt} odometer reading before tapping ${TAP_LABEL[to] ?? to}.`,
        ),
      );
      return;
    }

    setBusyId(card.trip.id);
    setError(undefined);
    try {
      const result = await services.dispatch.updateTripStatus(card.trip.id, to, reading);
      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }
      setOdometer((o) => ({ ...o, [card.trip.id]: '' }));
      await load();
    } catch (e) {
      setError(e);
    } finally {
      setBusyId(undefined);
    }
  };

  const cancelTrip = async (reason?: string): Promise<void> => {
    if (!services || !cancelTarget || !reason) return;
    setBusyId(cancelTarget.trip.id);
    try {
      const result = await services.dispatch.cancel({
        kind: 'driver',
        tripItemId: cancelTarget.trip.id,
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
      setBusyId(undefined);
    }
  };

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  if (!services || !user) {
    return (
      <PageShell heading="Driver dashboard">
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading /> : undefined}
      </PageShell>
    );
  }

  if (driverId === undefined) {
    return (
      <PageShell heading="Driver dashboard">
        <EmptyState title="You are not set up as a driver">
          Your account is not linked to a record in DriversMaster, so no offers
          can be sent to you. Ask a Fleet Admin to add your account to the
          driver list.
        </EmptyState>
      </PageShell>
    );
  }

  return (
    <PageShell
      heading="Driver dashboard"
      caption={`Signed in as ${user.displayName}. Offers refresh automatically.`}
      actions={
        <>
          <Button kind="secondary" onClick={toggleAlert} aria-pressed={alertOn}>
            {alertOn ? 'Sound: on' : 'Sound: off'}
          </Button>
          <Button kind="secondary" onClick={() => load().catch(() => undefined)}>
            Refresh now
          </Button>
        </>
      }
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      {takenMessage ? (
        <div className={styles.takenNotice} role="alert">
          <strong>{takenMessage}.</strong> Another driver accepted that trip
          first. Nothing has been assigned to you.
          <div style={{ marginTop: 8 }}>
            <Button kind="action" onClick={() => setTakenMessage(undefined)}>
              Dismiss
            </Button>
          </div>
        </div>
      ) : undefined}

      {/* Mechanism 1: the unmissable banner. */}
      {offers.length > 0 ? (
        <div className={styles.offerBanner} role="status" aria-live="assertive">
          <div>
            <div className={styles.offerBannerCount}>{offers.length}</div>
            <div className={styles.offerBannerText}>
              open ride offer{offers.length === 1 ? '' : 's'} waiting for you.
              First to accept gets the trip.
            </div>
          </div>
          <div className={styles.offerBannerActions}>
            {!alertOn ? (
              <span className={styles.bannerQuiet}>
                Turn sound on to be alerted next time.
              </span>
            ) : undefined}
          </div>
        </div>
      ) : undefined}

      {loading ? (
        <Loading what="Loading your offers and trips" />
      ) : (
        <>
          <section className={styles.section}>
            <SectionHeading>Open offers</SectionHeading>
            {offers.length === 0 ? (
              <EmptyState title="No open offers">
                You will see new offers here within {Math.round(pollMs / 1000)}{' '}
                seconds of them being sent, and you will get an e-mail too.
              </EmptyState>
            ) : (
              <div className={styles.cardGrid}>
                {offers.map((card) => (
                  <OfferCardView
                    key={card.offer.id}
                    card={card}
                    tick={tick}
                    busy={busyId === card.offer.id}
                    onAccept={() => {
                      accept(card).catch(() => undefined);
                    }}
                    onReject={() => setRejectTarget(card)}
                  />
                ))}
              </div>
            )}
          </section>

          <section className={styles.section}>
            <SectionHeading>My trips today</SectionHeading>
            {trips.length === 0 ? (
              <EmptyState title="No trips assigned to you right now" />
            ) : (
              <div className={styles.cardGrid}>
                {trips.map((card) => (
                  <TripCardView
                    key={card.trip.id}
                    card={card}
                    busy={busyId === card.trip.id}
                    odometer={odometer[card.trip.id] ?? ''}
                    onOdometer={(v) =>
                      setOdometer((o) => ({ ...o, [card.trip.id]: v }))
                    }
                    onAdvance={(to) => {
                      advance(card, to).catch(() => undefined);
                    }}
                    onCancel={() => setCancelTarget(card)}
                  />
                ))}
              </div>
            )}
          </section>

          <section className={styles.section}>
            <SectionHeading>Completed today</SectionHeading>
            {completed.length === 0 ? (
              <EmptyState title="Nothing finished yet today" />
            ) : (
              <div className={styles.cardGrid}>
                {completed.map((card) => (
                  <div className={styles.doneCard} key={card.trip.id}>
                    <div className={styles.cardHead}>
                      <Reference>{card.request.requestId}</Reference>
                      <Badge state="healthy" label="Completed" />
                    </div>
                    <div className={styles.route}>
                      {card.request.dropOffText}
                      <span className={styles.routeLeg}>from {card.request.pickupText}</span>
                    </div>
                    <FieldRow label="Finished">
                      {formatTime(card.trip.actualDropTime)}
                    </FieldRow>
                    {card.trip.openingReading !== undefined &&
                    card.trip.closingReading !== undefined ? (
                      <FieldRow label="Distance">
                        {card.trip.closingReading - card.trip.openingReading} km
                      </FieldRow>
                    ) : undefined}
                  </div>
                ))}
              </div>
            )}
          </section>

          <span className={styles.pollNote}>
            This page checks for new offers every {Math.round(pollMs / 1000)}{' '}
            seconds. If you miss an offer, the request goes to the
            Transportation Coordinator &mdash; it is never lost.
          </span>
        </>
      )}

      <ConfirmDialog
        open={rejectTarget !== undefined}
        title="Reject this offer?"
        message="Tell us why. The reason is recorded and reported on."
        consequences={[
          'The offer is closed for you and you will not see it again.',
          'Other drivers can still accept it.',
          'If nobody accepts, it goes to the Transportation Coordinator.',
        ]}
        confirmLabel="Reject the offer"
        destructive
        requireReason
        reasonLabel="Reason"
        reasonOptions={REJECTION_REASONS as unknown as string[]}
        busy={busyId !== undefined}
        onConfirm={(reason) => {
          reject(reason).catch(() => undefined);
        }}
        onCancel={() => setRejectTarget(undefined)}
      />

      <ConfirmDialog
        open={cancelTarget !== undefined}
        title="Cancel this trip?"
        message="You have already accepted this trip. Cancelling it now needs a reason."
        consequences={[
          'The Transportation Coordinator is notified immediately.',
          'The requester is notified immediately.',
          'The trip goes back to the Coordinator for reallocation - it is NOT ' +
            'broadcast again automatically.',
          'Your vehicle is released.',
        ]}
        confirmLabel="Cancel the trip"
        destructive
        requireReason
        reasonLabel="Why are you cancelling?"
        busy={busyId !== undefined}
        onConfirm={(reason) => {
          cancelTrip(reason).catch(() => undefined);
        }}
        onCancel={() => setCancelTarget(undefined)}
      />
    </PageShell>
  );
};

// -----------------------------------------------------------------------------
// Offer card
// -----------------------------------------------------------------------------

function OfferCardView({ card, busy, onAccept, onReject }: {
  card: IOfferCard;
  tick: number;
  busy: boolean;
  onAccept: () => void;
  onReject: () => void;
}): JSX.Element {
  const closesAt = card.request.offerClosesAt?.getTime() ?? 0;
  const seconds = Math.max(0, Math.round((closesAt - Date.now()) / 1000));
  const urgent = seconds <= 60;

  return (
    <div className={styles.offerCard}>
      <div className={styles.cardHead}>
        <Reference>{card.request.requestId}</Reference>
        <span
          className={[styles.countdown, urgent ? styles.countdownUrgent : ''].join(' ')}
        >
          {formatCountdown(seconds)} left
        </span>
      </div>

      <div className={styles.route}>
        {card.request.dropOffText}
        <span className={styles.routeLeg}>from {card.request.pickupText}</span>
      </div>

      <FieldRow label="When">{formatDateTime(card.request.requiredDateTime)}</FieldRow>
      <FieldRow label="Type">{card.request.requestType}</FieldRow>
      <FieldRow label="Vehicle">{card.vehicleTypeName}</FieldRow>
      {card.request.requestType === 'Items or Documents' ? (
        <FieldRow label="Items">{card.request.itemDescription ?? '—'}</FieldRow>
      ) : (
        <FieldRow label="Passengers">{card.request.noOfPax ?? '—'}</FieldRow>
      )}
      <FieldRow label="Requested by">{card.requesterName}</FieldRow>

      {/* The card deliberately omits the requester's phone number until the
          driver accepts. */}

      <div className={styles.cardActions}>
        <Button kind="primary" onClick={onAccept} disabled={busy || seconds === 0}>
          {busy ? 'Accepting…' : 'Accept'}
        </Button>
        <Button kind="secondary" onClick={onReject} disabled={busy}>
          Reject
        </Button>
      </div>
    </div>
  );
}

// -----------------------------------------------------------------------------
// Trip card - the four taps
// -----------------------------------------------------------------------------

function TripCardView({ card, busy, odometer, onOdometer, onAdvance, onCancel }: {
  card: ITripCard;
  busy: boolean;
  odometer: string;
  onOdometer: (value: string) => void;
  onAdvance: (to: TripStatus) => void;
  onCancel: () => void;
}): JSX.Element {
  const next = nextTripStatus(card.trip.tripStatus);
  const prompt = next ? odometerPromptFor(next) : undefined;

  return (
    <div className={styles.tripCard}>
      <div className={styles.cardHead}>
        <Reference>{card.request.requestId}</Reference>
        <Badge state={tripState(card.trip.tripStatus)} label={card.trip.tripStatus} />
      </div>

      <div className={styles.route}>
        {card.request.dropOffText}
        <span className={styles.routeLeg}>from {card.request.pickupText}</span>
      </div>

      <FieldRow label="When">{formatDateTime(card.request.requiredDateTime)}</FieldRow>
      <FieldRow label="Vehicle">{card.vehiclePlate || '—'}</FieldRow>
      <FieldRow label="Passenger">{card.requesterName}</FieldRow>
      {card.request.requestType !== 'Items or Documents' ? (
        <FieldRow label="Passengers">{card.request.noOfPax ?? '—'}</FieldRow>
      ) : undefined}

      {prompt ? (
        <div className={styles.odometerRow}>
          <div className={styles.odometerField}>
            <FormRow
              label={`${prompt === 'opening' ? 'Opening' : 'Closing'} odometer reading`}
              required
              hint="Goes straight into the trip sheet, so you do not re-key it later."
            >
              <TextField
                type="number"
                inputMode="numeric"
                value={odometer}
                onChange={(e) => onOdometer(e.target.value)}
              />
            </FormRow>
          </div>
        </div>
      ) : undefined}

      <div className={styles.cardActions}>
        {next ? (
          <Button kind="primary" onClick={() => onAdvance(next)} disabled={busy}>
            {busy ? 'Saving…' : (TAP_LABEL[next] ?? next)}
          </Button>
        ) : undefined}
        <Button kind="destructive" onClick={onCancel} disabled={busy}>
          Cancel trip
        </Button>
      </div>

      {card.trip.tripStatus === 'Assigned' ? (
        <Notice>
          The requester has been told you are coming. Tap{' '}
          <strong>On my way</strong> when you set off.
        </Notice>
      ) : undefined}
    </div>
  );
}

export default DriverDashboard;
