/**
 * The one place the concrete services are chosen.
 *
 * Swapping SharePoint for Dataverse is this file plus new classes behind the
 * same interfaces. No component changes.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ensureFonts } from '../styles/fonts';
import { DispatchService } from './DispatchService';
import { ICurrentUser, IServices } from './IServices';
import { MasterDataService } from './MasterDataService';
import { OfferService } from './OfferService';
import { ReportService } from './ReportService';
import { RequestService } from './RequestService';
import { initSP } from './SPContext';
import { TripService } from './TripService';
import { UserService } from './UserService';

export interface IBootstrapResult {
  services: IServices;
  user: ICurrentUser;
}

/**
 * The PROMISE is cached, not the result.
 *
 * Five web parts on one page call onInit at the same moment. Caching the
 * result would let all five get past the check before any of them finished,
 * and they would each do their own user lookup. Caching the promise means the
 * first call does the work and the other four await it.
 */
let pending: Promise<IBootstrapResult> | undefined;

/** The web part context the cached services were built with. */
let owner: WebPartContext | undefined;

/**
 * Called from each web part's onInit. Idempotent for the same context.
 *
 * A NEW context means SharePoint has navigated to a page without reloading
 * (this module stays loaded, the previous page's web part is disposed). The
 * cached services hold the old context, so they are rebuilt rather than
 * reused - reusing them is what made a page fail until it was refreshed.
 */
export async function bootstrap(context: WebPartContext): Promise<IBootstrapResult> {
  if (!pending || owner !== context) {
    owner = context;
    pending = (async (): Promise<IBootstrapResult> => {
      initSP(context);
      ensureFonts(context.pageContext.web.absoluteUrl);

      const userService = new UserService(context);
      const user = await userService.getCurrentUser();

      return {
        user,
        services: {
          user: userService,
          master: new MasterDataService(),
          requests: new RequestService(user),
          offers: new OfferService(),
          trips: new TripService(),
          dispatch: new DispatchService(context),
          reports: new ReportService(),
        },
      };
    })();

    // A failed bootstrap must not be cached, or every later web part on the
    // page inherits the failure with no way to retry.
    const attempt = pending;
    attempt.catch(() => {
      // Only forget it if a newer bootstrap has not already replaced it.
      if (pending === attempt) pending = undefined;
    });
  }

  return pending;
}

/** Forces the next bootstrap to re-read. */
export function resetServices(): void {
  pending = undefined;
  owner = undefined;
}
