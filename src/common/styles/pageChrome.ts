/**
 * Hides SharePoint's own chrome so the page reads as the application rather
 * than as a page inside a site.
 *
 * These elements live OUTSIDE the web part's DOM, so a .module.scss file
 * cannot reach them - CSS modules are scoped by design. The only way to style
 * them is a global stylesheet injected at runtime, which is what this does.
 *
 * REVERSIBLE two ways over: every rule is prefixed with the body class, so
 * removing the class disables all of them at once, and the <style> element is
 * removed when the last web part on the page disposes.
 *
 * SELECTOR STABILITY - read before editing.
 * SharePoint's generated class names are renumbered between renders and
 * service updates: overlay-56 became overlay-60, and mainHeader-247 became
 * mainHeader-193, within a single day. Anything that depends on those numbers
 * breaks silently - no error, the chrome simply reappears.
 *
 * So the header rules below key off data-automation attributes, which are
 * what Microsoft's own test suites hook into and are far more durable. The
 * numbered classes are kept only as a prefix-matched fallback.
 *
 * The Fluent `___jdtuxv0` hashes in the second block have no such attribute
 * available, so they remain the weak point. If the gap under the command bar
 * comes back, that is why.
 */

const STYLE_ID = 'fleet-portal-page-chrome';
export const HOST_CLASS = 'fleet-portal-host';
export const FULL_WIDTH_CLASS = 'fleet-portal-full-width';

/**
 * Reference count, because five web parts can share one page. Without it the
 * first one to dispose would strip the styling from the other four.
 */
let hostCount = 0;

const CSS = `
/* ---------------------------------------------------------------------------
   Hide the suite bar and the whole site header.

   TARGETED BY data-automation ATTRIBUTES, NOT BY CLASS. SharePoint renumbers
   its generated classes between renders and service updates - overlay-56
   became overlay-60 and mainHeader-247 became mainHeader-193 within a day.
   The automation ids are what the Microsoft test suites hook into, so they
   are the most stable handle available from CSS.

   The numbered classes are kept underneath as a second line of defence, by
   prefix. Both prefixes are SharePoint-specific and cannot collide with an
   SPFx module class, which is always suffixed with a content hash.
   --------------------------------------------------------------------------- */
body.${HOST_CLASS} #SuiteNavWrapper,
body.${HOST_CLASS} [data-automation-id="SiteHeaderOverlay"],
body.${HOST_CLASS} [data-automationid="SiteHeader"],
body.${HOST_CLASS} div[role="region"][aria-label="SharePoint Site Header"],
body.${HOST_CLASS} [class*="headerRow-"],
body.${HOST_CLASS} [class*="mainHeader-"] {
  display: none !important;
}

/* ---------------------------------------------------------------------------
   Close the gap the hidden chrome leaves behind.

   The Fluent provider classes are atomic and regenerated on every Fluent
   release, so the stable "fui-FluentProviderXX" token is matched by prefix as
   well as by the exact hash seen on the page today. These two are layout
   wrappers, so a prefix match here is safe in a way it is not for overlays.
   --------------------------------------------------------------------------- */
body.${HOST_CLASS} #spCommandBar,
body.${HOST_CLASS} .fui-FluentProvider.___jdtuxv0,
body.${HOST_CLASS} .fui-FluentProvider.___13yoiqc,
body.${HOST_CLASS} [class*="fui-FluentProviderrk"],
body.${HOST_CLASS} [class*="fui-FluentProviderr0"] {
  margin-top: 0 !important;
}
`;

/**
 * Add the stylesheet and the body classes. Safe to call from every web part on
 * the page; only the first call injects anything.
 *
 * @param apply pass false in page EDIT mode, so an author can still reach the
 *              suite bar. The host is still counted, to keep dispose balanced.
 */
export function applyPageChrome(apply: boolean): void {
  hostCount += 1;
  if (!apply) return;

  if (!document.getElementById(STYLE_ID)) {
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.appendChild(document.createTextNode(CSS));
    document.head.appendChild(style);
  }

  document.body.classList.add(HOST_CLASS, FULL_WIDTH_CLASS);
}

/**
 * Undo it, but only once the LAST web part on the page has gone. Removing the
 * body class on the first dispose would un-style the other four.
 */
export function removePageChrome(): void {
  hostCount = Math.max(0, hostCount - 1);
  if (hostCount > 0) return;

  document.body.classList.remove(HOST_CLASS, FULL_WIDTH_CLASS);

  const style = document.getElementById(STYLE_ID);
  if (style && style.parentNode) style.parentNode.removeChild(style);
}
