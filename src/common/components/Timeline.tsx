/**
 * What the requester sees instead of a map.
 *
 * Phase 1 has no GPS. Trip visibility comes from status transitions, each
 * stamped with a real timestamp, surfaced as a vertical timeline with a
 * completed, current and pending marker per stage.
 *
 * Each stage spells out its state in words - "Done", "Now", "To come" - so the
 * marker colour is never the only thing carrying the meaning.
 */

import * as React from 'react';
import { TripStatus } from '../models';
import { TRIP_FLOW } from '../rules/statusTransitions';
import { formatDateTime } from '../utils/dates';
import styles from './Shared.module.scss';

export interface ITimelineStage {
  status: TripStatus;
  label: string;
  /** What the requester is told at this stage. */
  description: string;
  reachedAt?: Date;
}

const STAGE_TEXT: Record<string, { label: string; description: string }> = {
  'Assigned': {
    label: 'Booking confirmed',
    description: 'A driver has accepted your booking.',
  },
  'En Route': {
    label: 'Driver on the way',
    description: 'Your driver is on the way to the pickup point.',
  },
  'In Transit': {
    label: 'On the way to your destination',
    description: 'You have been picked up.',
  },
  'Completed': {
    label: 'Trip completed',
    description: 'The trip has finished.',
  },
};

export function buildStages(
  current: TripStatus,
  timestamps: Partial<Record<TripStatus, Date>>,
): ITimelineStage[] {
  return TRIP_FLOW.map((status) => ({
    status,
    label: STAGE_TEXT[status].label,
    description: STAGE_TEXT[status].description,
    reachedAt: timestamps[status],
  }));
}

export const Timeline: React.FC<{
  stages: ITimelineStage[];
  current: TripStatus;
}> = ({ stages, current }) => {
  const currentIndex = TRIP_FLOW.indexOf(current);

  return (
    <ol className={styles.timeline}>
      {stages.map((stage, index) => {
        const isDone = index < currentIndex || stage.reachedAt !== undefined;
        const isCurrent = index === currentIndex;

        const markerClass = isCurrent
          ? styles.markerCurrent
          : isDone
            ? styles.markerDone
            : styles.markerPending;

        const stateWord = isCurrent ? 'Now' : isDone ? 'Done' : 'To come';

        return (
          <li className={styles.timelineItem} key={stage.status}>
            <span
              className={[styles.timelineMarker, markerClass].join(' ')}
              aria-hidden="true"
            >
              {isDone && !isCurrent ? '✓' : ''}
            </span>
            <div>
              <div className={styles.timelineLabel}>{stage.label}</div>
              <div className={styles.timelineState}>
                {stateWord}
                {stage.reachedAt ? ` · ${formatDateTime(stage.reachedAt)}` : ''}
              </div>
              {isCurrent ? (
                <div className={styles.timelineTime}>{stage.description}</div>
              ) : undefined}
            </div>
          </li>
        );
      })}
    </ol>
  );
};

/**
 * The manual ETA.
 *
 * Always shown WITH the time it was last updated, so nobody mistakes it for a
 * live calculation. Phase 2 populates the same field from a routing service
 * with no change to this component.
 */
export const EtaDisplay: React.FC<{
  minutes?: number;
  updatedOn?: Date;
}> = ({ minutes, updatedOn }) => {
  if (minutes === undefined) return <></>;
  return (
    <div className={styles.timelineTime}>
      approx. {minutes} min
      {updatedOn ? ` · estimate last updated ${formatDateTime(updatedOn)}` : ''}
    </div>
  );
};
