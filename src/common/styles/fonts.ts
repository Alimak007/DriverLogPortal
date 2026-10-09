/**
 * Font loading for the whole solution.
 *
 * Both faces are declared ONCE, here, and never per component. Call
 * ensureFonts() from each web part's onInit; it is idempotent.
 *
 * They can be served from the site's own SiteAssets/fonts folder or from
 * Google Fonts - a setting in APP_CONFIG.fonts. Site assets are the safer
 * default where outbound font requests are restricted.
 *
 * If neither loads, the fallback chains in _tokens.scss degrade the interface
 * to Georgia and Segoe UI rather than to browser defaults.
 */

import { APP_CONFIG } from '../config/AppConfig';

const STYLE_ELEMENT_ID = 'fm-portal-fonts';

const GOOGLE_HREF =
  'https://fonts.googleapis.com/css2' +
  '?family=Playfair+Display:wght@700' +
  '&family=Roboto:wght@400;500;600;700' +
  '&display=swap';

/** @font-face rules for fonts served from the site's own SiteAssets folder. */
function siteAssetsCss(webAbsoluteUrl: string): string {
  const base = `${webAbsoluteUrl}${APP_CONFIG.fonts.siteAssetsPath}`;

  const face = (
    family: string,
    file: string,
    weight: number,
  ): string =>
    `@font-face{` +
    `font-family:'${family}';` +
    `src:url('${base}/${file}.woff2') format('woff2');` +
    `font-weight:${weight};font-style:normal;font-display:swap;}`;

  return [
    face('Playfair Display', 'PlayfairDisplay-Bold', 700),
    face('Roboto', 'Roboto-Regular', 400),
    face('Roboto', 'Roboto-Medium', 500),
    face('Roboto', 'Roboto-SemiBold', 600),
    face('Roboto', 'Roboto-Bold', 700),
  ].join('');
}

let loaded = false;

/**
 * Inject the font declarations. Safe to call from every web part on a page;
 * only the first call does anything.
 */
export function ensureFonts(webAbsoluteUrl: string): void {
  if (loaded || document.getElementById(STYLE_ELEMENT_ID)) {
    loaded = true;
    return;
  }

  if (APP_CONFIG.fonts.source === 'google') {
    const link = document.createElement('link');
    link.id = STYLE_ELEMENT_ID;
    link.rel = 'stylesheet';
    link.href = GOOGLE_HREF;
    document.head.appendChild(link);
  } else {
    const style = document.createElement('style');
    style.id = STYLE_ELEMENT_ID;
    style.appendChild(document.createTextNode(siteAssetsCss(webAbsoluteUrl)));
    document.head.appendChild(style);
  }

  loaded = true;
}
