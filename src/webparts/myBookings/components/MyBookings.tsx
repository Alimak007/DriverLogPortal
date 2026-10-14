/**
 * My Bookings - request history, and the status timeline for an active trip.
 *
 * The cancellation rule is the thing to get right here. The requester may
 * cancel while the request is Submitted, Held or Awaiting Coordinator. ONCE A
 * DRIVER HAS ACCEPTED, cancellation is routed to the Coordinator, who cancels
 * on their behalf and notifies the driver. The interface must not pretend
 * otherwise - it offers a different action, with different wording, in that
 * case.
 */

import * as React from 'react';
import { ConfirmDialog } from '../../../common/components/ConfirmDialog';
import { DataGrid, IColumn, RowActions } from '../../../common/components/DataGrid';
import { Icon, IconName } from '../../../common/components/Icons';
import {
  Badge,
  Button,
  EmptyState,
  ErrorBanner,
  HeadedCard,
  LinkButton,
  Loading,
  Notice,
  PageShell,
  Reference,
} from '../../../common/components/Shared';
import { buildStages, EtaDisplay, Timeline } from '../../../common/components/Timeline';
import {
  IRequestWithTrip,
  ITransportRequest,
  TripStatus,
  requestState,
} from '../../../common/models';
import {
  cancellationNeedsCoordinator,
  requesterMayCancel,
} from '../../../common/rules/statusTransitions';
import { ICurrentUser, IServices } from '../../../common/services/IServices';
import { formatDateTime } from '../../../common/utils/dates';
import styles from './MyBookings.module.scss';

export interface IMyBookingsProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
  historyCount: number;
  refreshMs: number;
  /** The Book a Ride page. The button to it is left out when this is not given. */
  bookARideUrl?: string;
}

/** One line of the booking detail: an icon, a label, and the value. */
function InfoRow({ icon, label, action, children }: {
  icon: IconName;
  label: string;
  action?: React.ReactNode;
  children: React.ReactNode;
}): JSX.Element {
  return (
    <div className={styles.infoRow}>
      <span className={styles.infoIcon} aria-hidden="true">
        <Icon name={icon} size={18} />
      </span>
      <span className={styles.infoLabel}>{label}</span>
      <span className={styles.infoValue}>{children}</span>
      {action}
    </div>
  );
}

export const MyBookings: React.FC<IMyBookingsProps> = ({
  services,
  user,
  bootstrapError,
  historyCount,
  refreshMs,
  bookARideUrl,
}) => {
  const [requests, setRequests] = React.useState<ITransportRequest[]>([]);
  const [selectedId, setSelectedId] = React.useState<number | undefined>();
  const [detail, setDetail] = React.useState<IRequestWithTrip | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [detailLoading, setDetailLoading] = React.useState(false);
  const [error, setError] = React.useState<unknown>(bootstrapError);
  const [cancelling, setCancelling] = React.useState(false);
  const [dialogOpen, setDialogOpen] = React.useState(false);
  const [notice, setNotice] = React.useState<string | undefined>();

  const load = React.useCallback(async (): Promise<void> => {
    if (!services) return;
    try {
      const rows = await services.requests.getMyRequests(historyCount);
      setRequests(rows);
      setSelectedId((current) => current ?? rows[0]?.id);
    } catch (e) {
      setError(e);
    } finally {
      setLoading(false);
    }
  }, [services, historyCount]);

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

  React.useEffect(() => {
    if (!services || selectedId === undefined) {
      setDetail(undefined);
      return undefined;
    }

    let cancelled = false;
    setDetailLoading(true);

    const read = async (): Promise<void> => {
      try {
        const result = await services.requests.getRequestWithTrip(selectedId);
        if (!cancelled) setDetail(result);
      } catch (e) {
        if (!cancelled) setError(e);
      } finally {
        if (!cancelled) setDetailLoading(false);
      }
    };

    read().catch(() => undefined);
    const handle = window.setInterval(() => {
      read().catch(() => undefined);
    }, refreshMs);

    return () => {
      cancelled = true;
      window.clearInterval(handle);
    };
  }, [services, selectedId, refreshMs]);

  const selected = detail?.request;
  const mayCancelDirectly = selected ? requesterMayCancel(selected.status) : false;
  const needsCoordinator = selected ? cancellationNeedsCoordinator(selected.status) : false;

  const doCancel = async (reason?: string): Promise<void> => {
    if (!services || !selected || !reason) return;
    setCancelling(true);
    setError(undefined);
    setNotice(undefined);
    try {
      // One call either way. The service decides whether this is a
      // cancellation or a request for the Coordinator to cancel, because that
      // depends on the status it reads, not on what the screen last saw.
      const result = await services.dispatch.cancel({
        kind: 'requester',
        requestItemId: selected.id,
        reason,
      });

      if (!result.ok) {
        setError(new Error(result.message));
        return;
      }

      setDialogOpen(false);
      setNotice(result.data.message);
      await load();
    } catch (e) {
      setError(e);
    } finally {
      setCancelling(false);
    }
  };

  const columns: IColumn<ITransportRequest>[] = [
    {
      key: 'requestId',
      label: 'Reference',
      render: (r) => <Reference>{r.requestId}</Reference>,
      width: '96px',
    },
    {
      key: 'required',
      label: 'Required',
      render: (r) => formatDateTime(r.requiredDateTime),
      width: '128px',
    },
    {
      key: 'route',
      label: 'From and to',
      render: (r) =>
        `${r.pickupText || '—'} → ${r.dropOffText || '—'}`,
    },
    {
      key: 'type',
      label: 'Request type',
      render: (r) => r.requestType,
      filterValue: (r) => r.requestType,
      width: '128px',
    },
    {
      key: 'status',
      label: 'Status',
      // Colour AND the word, always.
      render: (r) => <Badge state={requestState(r.status)} label={r.status} />,
      filterValue: (r) => r.status,
      width: '168px',
    },
    {
      key: 'actions',
      label: 'Actions',
      render: (r) => (
        <RowActions>
          <Button kind="action" onClick={() => setSelectedId(r.id)}>
            {selectedId === r.id ? 'Showing' : 'View'}
          </Button>
        </RowActions>
      ),
      width: '86px',
    },
  ];

  if (!services || !user) {
    return (
      <PageShell heading="My bookings">
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading /> : undefined}
      </PageShell>
    );
  }

  const bookARide = bookARideUrl ? (
    <LinkButton href={bookARideUrl} kind="primary">
      <Icon name="car" size={18} />
      Book a Ride
    </LinkButton>
  ) : undefined;

  const nowStage = detail?.trip
    ? buildStages(detail.trip.tripStatus, stageTimestamps(detail)).filter(
        (st) => st.status === detail.trip?.tripStatus,
      )[0]
    : undefined;

  return (
    <PageShell
      heading="My bookings"
      caption="Your requests, newest first. Select one to see where it has got to."
      actions={
        <>
          {bookARide}
          <Button kind="secondary" onClick={() => load().catch(() => undefined)}>
            Refresh
          </Button>
        </>
      }
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      {notice ? <Notice tone="healthy">{notice}</Notice> : undefined}

      {loading ? (
        <Loading what="Loading your bookings" />
      ) : requests.length === 0 ? (
        <EmptyState title="You have not booked anything yet">
          Use the Book a ride page to make your first request.
          {bookARide ? <div className={styles.emptyAction}>{bookARide}</div> : undefined}
        </EmptyState>
      ) : (
        <div className={styles.split}>
          <HeadedCard icon="calendar" title="Request History">
            <DataGrid
              caption="Your transport requests, newest first"
              captionVisuallyHidden
              columns={columns}
              rows={requests}
              rowKey={(r) => r.id}
              emptyMessage="No requests match the filters you have set."
            />
          </HeadedCard>

          <HeadedCard icon="doc" title="Booking Detail" className={styles.detailCard}>
            {detailLoading && !detail ? (
              <Loading what="Loading the booking" />
            ) : !detail || !selected ? (
              <EmptyState title="Select a booking">
                Choose a row to see its status and your driver&apos;s details.
              </EmptyState>
            ) : (
              <>
                <div className={styles.detailHeader}>
                  <span>
                    Reference <Reference>{selected.requestId}</Reference>
                  </span>
                  <Badge state={requestState(selected.status)} label={selected.status} />
                </div>

                {detail.trip ? (
                  <div className={styles.statusBanner}>
                    <span className={styles.statusBannerIcon} aria-hidden="true">
                      <Icon name="car" size={22} />
                    </span>
                    <div className={styles.statusBannerText}>
                      <strong>{nowStage ? nowStage.label : detail.trip.tripStatus}</strong>
                      <EtaDisplay
                        minutes={detail.trip.etaMinutes}
                        updatedOn={detail.trip.etaUpdatedOn}
                      />
                    </div>
                  </div>
                ) : undefined}

                <div className={styles.infoList}>
                  <InfoRow icon="calendar" label="Required">
                    {formatDateTime(selected.requiredDateTime)}
                  </InfoRow>
                  <InfoRow icon="target" label="From">
                    {selected.pickupText || '—'}
                  </InfoRow>
                  <InfoRow icon="pin" label="To">
                    {selected.dropOffText || '—'}
                  </InfoRow>
                  <InfoRow icon="doc" label="Request type">
                    {selected.requestType}
                  </InfoRow>
                  <InfoRow icon="car" label="Vehicle type">
                    {selected.vehicleTypeRequested?.title ?? '—'}
                    {selected.otherVehicleText ? ` (${selected.otherVehicleText})` : ''}
                  </InfoRow>
                  {selected.requestType === 'Items or Documents' ? (
                    <InfoRow icon="box" label="Items">
                      {selected.itemDescription ?? '—'}
                    </InfoRow>
                  ) : (
                    <InfoRow icon="users" label="Passengers">
                      {selected.noOfPax ?? '—'}
                    </InfoRow>
                  )}
                </div>

                {selected.status === 'Submitted' || selected.status === 'Broadcasting' ? (
                  <Notice>
                    We have your request and are finding you a driver. We will
                    e-mail you as soon as one accepts.
                  </Notice>
                ) : undefined}

                {selected.status === 'Awaiting Coordinator' ? (
                  <Notice tone="attention">
                    No driver was free to take this straight away, so the
                    Transportation Coordinator is assigning someone. Your
                    booking has not been lost.
                  </Notice>
                ) : undefined}

                {selected.status === 'Rejected' && selected.rejectionReason ? (
                  <Notice tone="attention">
                    This request was declined: {selected.rejectionReason}
                  </Notice>
                ) : undefined}

                {selected.status === 'Cancelled' && selected.cancellationReason ? (
                  <Notice tone="attention">
                    Cancelled: {selected.cancellationReason}
                  </Notice>
                ) : undefined}

                {detail.trip ? (
                  <>
                    <h3 className={styles.subHeading}>Your driver and vehicle</h3>
                    <div className={styles.infoList}>
                      <InfoRow
                        icon="user"
                        label="Driver"
                        action={
                          detail.driverPhone ? (
                            <a
                              className={styles.callButton}
                              href={`tel:${detail.driverPhone}`}
                              aria-label={`Call ${detail.driverName ?? 'the driver'}`}
                            >
                              <Icon name="phone" size={16} />
                            </a>
                          ) : undefined
                        }
                      >
                        {detail.driverName ?? '—'}
                        <span className={styles.infoSub}>
                          {detail.driverPhone ? (
                            <a href={`tel:${detail.driverPhone}`}>{detail.driverPhone}</a>
                          ) : (
                            '—'
                          )}
                        </span>
                      </InfoRow>
                      <InfoRow icon="car" label="Vehicle">
                        {[detail.vehicleColour, detail.vehicleMake, detail.vehicleModel]
                          .filter(Boolean)
                          .join(' ') || '—'}
                        <span className={styles.infoSub}>
                          {`Plate: ${detail.vehiclePlate ?? '—'}`}
                        </span>
                      </InfoRow>
                    </div>

                    <h3 className={styles.subHeading}>Progress</h3>
                    <div className={styles.progressCard}>
                      <Timeline
                        stages={buildStages(
                          detail.trip.tripStatus,
                          stageTimestamps(detail),
                        )}
                        current={detail.trip.tripStatus}
                      />
                    </div>
                  </>
                ) : undefined}

                {mayCancelDirectly || needsCoordinator ? (
                  <div className={styles.detailActions}>
                    <Button
                      kind="destructive"
                      className={styles.cancelButton}
                      onClick={() => setDialogOpen(true)}
                    >
                      {mayCancelDirectly ? 'Cancel this booking' : 'Ask to cancel'}
                    </Button>
                  </div>
                ) : undefined}
              </>
            )}
          </HeadedCard>
        </div>
      )}

      <ConfirmDialog
        open={dialogOpen}
        title={mayCancelDirectly ? 'Cancel this booking?' : 'Ask the Coordinator to cancel?'}
        message={
          mayCancelDirectly
            ? 'This cancels the request before any driver has taken it.'
            : 'A driver has already accepted this trip, so only the Transportation ' +
              'Coordinator can cancel it. Your request goes to them.'
        }
        consequences={
          mayCancelDirectly
            ? [
                'The request is closed and will not be offered to any driver.',
                'Any open offers to drivers are withdrawn.',
                'You will need to book again if you still need transport.',
              ]
            : [
                'The Transportation Coordinator is notified immediately.',
                'Your assigned driver is notified that cancellation has been requested.',
                'The trip stays assigned until the Coordinator acts on it.',
              ]
        }
        confirmLabel={mayCancelDirectly ? 'Cancel the booking' : 'Send the request'}
        destructive
        requireReason
        reasonLabel="Why are you cancelling?"
        busy={cancelling}
        onConfirm={(reason) => {
          doCancel(reason).catch(() => undefined);
        }}
        onCancel={() => setDialogOpen(false)}
      />
    </PageShell>
  );
};

function stageTimestamps(detail: IRequestWithTrip): Partial<Record<TripStatus, Date>> {
  const stamps: Partial<Record<TripStatus, Date>> = {};
  if (detail.trip) {
    stamps.Assigned = detail.trip.acceptedOn;
    stamps['En Route'] = detail.trip.actualStartTime;
    stamps['In Transit'] = detail.trip.actualPickupTime;
    stamps.Completed = detail.trip.actualDropTime;
  }
  detail.timeline.forEach((h) => {
    const status = h.toStatus as TripStatus;
    if (!stamps[status]) stamps[status] = h.changedOn;
  });
  return stamps;
}

export default MyBookings;
