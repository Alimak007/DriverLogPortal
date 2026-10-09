/**
 * The one place PnPjs is configured.
 *
 * Everything above this file deals in typed services behind interfaces. That
 * is what makes a later move to Power Apps, Dataverse or any other store a
 * service-layer change rather than a rewrite.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/batching';

let instance: SPFI | undefined;
let webUrl = '';

export function initSP(context: WebPartContext): SPFI {
  if (!instance) {
    instance = spfi().using(SPFx(context));
    webUrl = context.pageContext.web.absoluteUrl;
  }
  return instance;
}

export function getSP(): SPFI {
  if (!instance) {
    throw new Error(
      'The SharePoint context has not been initialised. ' +
        'Call initSP(context) from the web part onInit before using any service.',
    );
  }
  return instance;
}

export function getWebUrl(): string {
  return webUrl;
}

/**
 * A correlation ID on every write path, as the non-functional requirements
 * ask. It travels to the flow, into the flow run history and into
 * NotificationLog, so one failed booking can be traced end to end.
 *
 * NO PERSONAL DATA IN LOGS - this is an opaque identifier and nothing else.
 */
export function newCorrelationId(): string {
  const rand = (): string => Math.random().toString(16).slice(2, 10);
  return `${Date.now().toString(16)}-${rand()}-${rand()}`;
}

/**
 * A failure that a human has to be told about.
 *
 * Every failed write surfaces a message naming WHAT FAILED and WHAT TO DO.
 * There are no silent catch blocks in this solution.
 */
export class PortalError extends Error {
  public readonly correlationId: string;
  public readonly whatFailed: string;
  public readonly whatToDo: string;
  public readonly cause?: unknown;

  constructor(
    whatFailed: string,
    whatToDo: string,
    correlationId: string,
    cause?: unknown,
  ) {
    super(`${whatFailed} ${whatToDo}`);
    this.name = 'PortalError';
    this.whatFailed = whatFailed;
    this.whatToDo = whatToDo;
    this.correlationId = correlationId;
    this.cause = cause;
  }
}

/**
 * Wrap a read or a write so that whatever PnPjs throws becomes a message the
 * interface can show. Never swallow - always rethrow as a PortalError.
 */
export async function guard<T>(
  whatFailed: string,
  whatToDo: string,
  work: (correlationId: string) => Promise<T>,
  correlationId: string = newCorrelationId(),
): Promise<T> {
  try {
    return await work(correlationId);
  } catch (e) {
    console.error(`[FleetPortal ${correlationId}] ${whatFailed}`, e);
    throw new PortalError(whatFailed, whatToDo, correlationId, e);
  }
}
