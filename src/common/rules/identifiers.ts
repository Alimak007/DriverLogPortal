/**
 * Request and trip identifiers.
 *
 * PURE FUNCTIONS, NO SHAREPOINT.
 *
 * The rule from the requirements: a prefix plus the SharePoint item ID,
 * zero-padded to four digits, computed ON CREATE.
 *
 * A SEPARATE COUNTER ITEM WOULD RACE UNDER CONCURRENT INTAKE AND MUST NOT BE
 * USED. The item ID is allocated by SharePoint itself and is therefore the
 * only collision-free source available without a transaction.
 *
 * OPEN QUESTION: the mockups show both FB-001 and FB-20260925-001. Both are
 * implemented; APP_CONFIG.requestIdFormat chooses. This must be settled before
 * the first production request is created - the format cannot change
 * afterwards without reissuing every identifier.
 */

import { APP_CONFIG, RequestIdFormat } from '../config/AppConfig';

const PAD = 4;

function pad(itemId: number): string {
  const s = String(itemId);
  return s.length >= PAD ? s : new Array(PAD - s.length + 1).join('0') + s;
}

function datePart(d: Date): string {
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  return `${y}${m < 10 ? '0' : ''}${m}${day < 10 ? '0' : ''}${day}`;
}

export function formatIdentifier(
  prefix: string,
  itemId: number,
  createdOn: Date,
  format: RequestIdFormat,
): string {
  return format === 'dated'
    ? `${prefix}-${datePart(createdOn)}-${pad(itemId)}`
    : `${prefix}-${pad(itemId)}`;
}

/**
 * FB-0001, or FB-20260925-0001.
 *
 * OPEN QUESTION: which format. Both are implemented and APP_CONFIG chooses, so
 * this is the ONE place to change it. Settle it before the first production
 * request exists - the format cannot change afterwards without reissuing every
 * identifier.
 */
export function formatRequestRef(itemId: number, createdOn: Date = new Date()): string {
  return formatIdentifier(
    APP_CONFIG.requestIdPrefix,
    itemId,
    createdOn,
    APP_CONFIG.requestIdFormat,
  );
}

/** TR-0001, or TR-20260925-0001. */
export function formatTripRef(itemId: number, createdOn: Date = new Date()): string {
  return formatIdentifier(
    APP_CONFIG.tripIdPrefix,
    itemId,
    createdOn,
    APP_CONFIG.requestIdFormat,
  );
}

/** Both formats parse back to the item id, so either can be searched for. */
export function itemIdFromIdentifier(identifier: string): number | undefined {
  const parts = identifier.split('-');
  if (parts.length < 2) return undefined;
  const n = parseInt(parts[parts.length - 1], 10);
  return isNaN(n) ? undefined : n;
}
