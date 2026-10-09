/**
 * What all five web parts share.
 *
 * Each one bootstraps the same services, declares the same flow-endpoint
 * properties, and adds the two body classes that hide SharePoint's own chrome.
 * Putting that here means a new web part cannot forget any of it.
 */

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IServices } from '../services/IServices';
import { ICurrentUser } from '../services/IServices';
import { bootstrap } from '../services/ServiceFactory';

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
 * The web part's half of the chrome contract.
 *
 * The brand guideline hides SharePoint's suite bar, site header and page
 * margins with a stylesheet INJECTED BY AN SPFX APPLICATION CUSTOMISER, keyed
 * off two body classes that the web part adds and removes. This base class
 * adds the classes; everything is reversible.
 *
 * NOTE: that application customiser is NOT part of this package. The
 * requirements specify five web parts and no extension, so until one is added
 * these classes are inert - they are here so the contract is already in place
 * and adding the extension needs no web part change. Do not remove them
 * because nothing appears to read them yet.
 */
const BODY_CLASS = 'fleet-portal-host';
const FULL_WIDTH_CLASS = 'fleet-portal-full-width';

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

    document.body.classList.add(BODY_CLASS, FULL_WIDTH_CLASS);
  }

  protected onDispose(): void {
    document.body.classList.remove(BODY_CLASS, FULL_WIDTH_CLASS);
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
