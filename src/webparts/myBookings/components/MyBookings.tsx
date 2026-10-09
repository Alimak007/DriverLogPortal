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
import {
  Badge,
  Button,
  EmptyState,
  ErrorBanner,
  FieldRow,
  Loading,
  Notice,
  PageShell,
  Reference,
  SectionHeading,
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
}

export const MyBookings: React.FC<IMyBookingsProps> = ({
  services,
  user,
  bootstrapError,
  historyCount,
  refreshMs,
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
      width: '130px',
    },
    {
      key: 'required',
      label: 'Required',
      render: (r) => formatDateTime(r.requiredDateTime),
      width: '150px',
    },
    {
      key: 'route',
      label: 'From and to',
      render: (r) =>
        `${r.pickupLocation?.title ?? '—'} → ${r.dropOffLocation?.title ?? '—'}`,
    },
    {
      key: 'type',
      label: 'Request type',
      render: (r) => r.requestType,
      filterValue: (r) => r.requestType,
      width: '160px',
    },
    {
      key: 'status',
      label: 'Status',
      // Colour AND the word, always.
      render: (r) => <Badge state={requestState(r.status)} label={r.status} />,
      filterValue: (r) => r.status,
      width: '150px',
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
      width: '100px',
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

  return (
    <PageShell
      heading="My bookings"
      caption="Your requests, newest first. Select one to see where it has got to."
      actions={
        <Button kind="secondary" onClick={() => load().catch(() => undefined)}>
          Refresh
        </Button>
      }
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      {notice ? <Notice tone="healthy">{notice}</Notice> : undefined}

      {loading ? (
        <Loading what="Loading your bookings" />
      ) : requests.length === 0 ? (
        <EmptyState title="You have not booked anything yet">
          Use the Book a ride page to make your first request.
        </EmptyState>
      ) : (
        <div className={styles.split}>
          <div>
            <SectionHeading>Request history</SectionHeading>
            <DataGrid
              caption="Your transport requests, newest first"
              captionVisuallyHidden
              columns={columns}
              rows={requests}
              rowKey={(r) => r.id}
              emptyMessage="No requests match the filters you have set."
            />
          </div>

          <div>
            <SectionHeading>Booking detail</SectionHeading>
            {detailLoading && !detail ? (
              <Loading what="Loading the booking" />
            ) : !detail || !selected ? (
              <EmptyState title="Select a booking">
                Choose a row to see its status and your driver&apos;s details.
              </EmptyState>
            ) : (
              <div className={styles.detailPanel}>
                <div className={styles.detailHeader}>
                  <Reference>{selected.requestId}</Reference>
                  <Badge state={requestState(selected.status)} label={selected.status} />
                </div>

                <div className={styles.detailGrid}>
                  <FieldRow label="Required">
                    {formatDateTime(selected.requiredDateTime)}
                  </FieldRow>
                  <FieldRow label="From">
                    {selected.pickupLocation?.title ?? '—'}
                  </FieldRow>
                  <FieldRow label="To">
                    {selected.dropOffLocation?.title ?? '—'}
                  </FieldRow>
                  <FieldRow label="Request type">{selected.requestType}</FieldRow>
                  <FieldRow label="Vehicle type">
                    {selected.vehicleTypeRequested?.title ?? '—'}
                    {selected.otherVehicleText ? ` (${selected.otherVehicleText})` : ''}
                  </FieldRow>
                  {selected.requestType === 'Items or Documents' ? (
                    <FieldRow label="Items">
                      {selected.itemDescription ?? '—'}
                    </FieldRow>
                  ) : (
                    <FieldRow label="Passengers">{selected.noOfPax ?? '—'}</FieldRow>
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
                    <hr className={styles.divider} />
                    <SectionHeading>Your driver and vehicle</SectionHeading>
                    <div className={styles.detailGrid}>
                      <FieldRow label="Driver">{detail.driverName ?? '—'}</FieldRow>
                      <FieldRow label="Phone">
                        {detail.driverPhone ? (
                          <a href={`tel:${detail.driverPhone}`}>{detail.driverPhone}</a>
                        ) : (
                          '—'
                        )}
                      </FieldRow>
                      <FieldRow label="Vehicle">
                        {[detail.vehicleColour, detail.vehicleMake, detail.vehicleModel]
                          .filter(Boolean)
                          .join(' ') || '—'}
                      </FieldRow>
                      <FieldRow label="Plate">{detail.vehiclePlate ?? '—'}</FieldRow>
                    </div>

                    <hr className={styles.divider} />
                    <SectionHeading>Progress</SectionHeading>
                    <Timeline
                      stages={buildStages(
                        detail.trip.tripStatus,
                        stageTimestamps(detail),
                      )}
                      current={detail.trip.tripStatus}
                    />
                    <EtaDisplay
                      minutes={detail.trip.etaMinutes}
                      updatedOn={detail.trip.etaUpdatedOn}
                    />
                  </>
                ) : undefined}

                {mayCancelDirectly || needsCoordinator ? (
                  <div className={styles.detailActions}>
                    <Button kind="destructive" onClick={() => setDialogOpen(true)}>
                      {mayCancelDirectly ? 'Cancel this booking' : 'Ask to cancel'}
                    </Button>
                  </div>
                ) : undefined}
              </div>
            )}
          </div>
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
