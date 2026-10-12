/**
 * What all five web parts share.
 *
 * Each one bootstraps the same services, declares the same flow-endpoint
 * properties, and adds the two body classes that hide SharePoint's own chrome.
 * Putting that here means a new web part cannot forget any of it.
 */

import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IServices } from '../services/IServices';
import { ICurrentUser } from '../services/IServices';
import { bootstrap } from '../services/ServiceFactory';
import { applyPageChrome, removePageChrome } from '../styles/pageChrome';

/**
 * Common web part properties.
 *
 * The seven Power Automate flow URLs that used to live here are gone. The
 * "When an HTTP request is received" trigger is a PREMIUM connector and this
 * tenant has standard only, so every dispatch action now runs in the browser
 * through DispatchService. Flows still send e-mail and run the scheduled jobs,
 * but the web part never calls one.
 */
export interface IFleetWebPartProps {
  /** Reserved. Every web part declares the same shape. */
  readonly _reserved?: never;
}

/**
 * Hiding SharePoint's chrome.
 *
 * The brand guideline hides the suite bar, the site header and the page-level
 * margins so the application reads as itself rather than as a page inside a
 * site. Those elements are outside the web part's DOM, so the rules live in a
 * global stylesheet - see common/styles/pageChrome.ts.
 *
 * NOT applied in page EDIT mode. An author needs the suite bar and the header
 * to do their job, and a web part that hides the toolbar you are editing with
 * is its own kind of bug. Switching between edit and view re-renders but does
 * not re-run onInit, so reload the page after publishing to see the final
 * result.
 */

export abstract class FleetWebPartBase<
  TProps extends IFleetWebPartProps,
> extends BaseClientSideWebPart<TProps> {
  protected services!: IServices;
  protected currentUser!: ICurrentUser;
  protected bootstrapError?: unknown;

  protected async onInit(): Promise<void> {
    await super.onInit();

    try {
      const result = await bootstrap(this.context);
      this.services = result.services;
      this.currentUser = result.user;
    } catch (e) {
      // Captured rather than thrown, so the web part can render the reason
      // instead of leaving an empty canvas. No silent catch: the component
      // shows this.
      this.bootstrapError = e;
    }

    applyPageChrome(this.displayMode !== DisplayMode.Edit);
  }

  protected onDispose(): void {
    removePageChrome();
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected abstract settingsGroups(): IPropertyPaneGroup[];

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Fleet Booking and Dispatch Portal' },
          groups: this.settingsGroups(),
        },
      ],
    };
  }
}
