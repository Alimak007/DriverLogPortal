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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { AccessNotice } from '../components/Shared';
import { IServices } from '../services/IServices';
import { ICurrentUser } from '../services/IServices';
import { bootstrap } from '../services/ServiceFactory';
import { applyPageChrome, removePageChrome } from '../styles/pageChrome';
import { FleetWebPartKey, mayOpen, WEB_PART_ACCESS } from './access';

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

  /**
   * Unmounting belongs here now that rendering does.
   *
   * It used to be repeated in all five web parts. Once renderElement moved the
   * render into this class, leaving the unmount behind meant the pair was
   * split across files - which is both a lint failure and the shape a leak
   * hides in. Subclasses no longer override this at all.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    removePageChrome();
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected abstract settingsGroups(): IPropertyPaneGroup[];

  /** Which entry in WEB_PART_ACCESS governs this web part. */
  protected abstract accessKey(): FleetWebPartKey;

  /**
   * Render the web part, or the access notice in its place.
   *
   * Every web part calls this instead of ReactDom.render, so the check cannot
   * be forgotten by a new one - the abstract accessKey() above makes it a
   * compile error to add a web part without declaring who it is for.
   *
   * THREE DELIBERATE EXCEPTIONS, each of which fails OPEN:
   *
   *   1. A failed bootstrap. currentUser is undefined because the lookup
   *      broke, not because the person lacks a role. The component already
   *      renders that error, and replacing it with "not available to you"
   *      would send someone chasing a permission problem that does not exist.
   *   2. A site collection administrator. Whoever builds these pages has to
   *      be able to see what they are placing, and locking the owner out of
   *      their own site to enforce a navigation aid is a poor trade.
   *   3. Page edit mode is NOT an exception - an author sees exactly what
   *      their audience will see. The property pane still opens either way,
   *      so a page can be configured from behind the notice.
   */
  protected renderElement(element: React.ReactElement): void {
    const key = this.accessKey();

    const bootstrapFailed = this.bootstrapError !== undefined || !this.currentUser;
    const legacy = this.context.pageContext.legacyPageContext as
      | { isSiteAdmin?: boolean }
      | undefined;
    const isSiteAdmin = legacy?.isSiteAdmin === true;

    const allowed =
      bootstrapFailed || isSiteAdmin || mayOpen(this.currentUser?.roles, key);
    const access = WEB_PART_ACCESS[key];

    // One render call, not one per branch: the lint rule pairs render against
    // unmount, and two renders against a single unmount is the shape a leak
    // hides in even when this particular pair is fine.
    ReactDom.render(
      allowed
        ? element
        : React.createElement(AccessNotice, {
            title: access.title,
            groups: access.groups,
          }),
      this.domElement,
    );
  }

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
