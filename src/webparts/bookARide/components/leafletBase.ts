/**
 * The Leaflet bits both maps share: the stylesheet, the pins, and the tile
 * layer with its fallback.
 *
 * Two components draw maps - the booking picker and the route view on the
 * confirmation screen - and the tile fallback in particular must behave
 * identically in both. Duplicating it would mean one of them silently keeping
 * the old behaviour the next time it changes.
 */

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as L from 'leaflet';

/**
 * Leaflet's stylesheet, loaded as a plain <link>.
 *
 * Deliberately NOT imported as a module. The SPFx bundler runs CSS through
 * css-loader with module scoping, which hashes every class name - and Leaflet
 * positions its tiles, panes and controls by class name. Hash them and you get
 * a map-shaped pile of mis-positioned tiles.
 *
 * If the tenant blocks unpkg, copy leaflet.css into SiteAssets and change this
 * one URL.
 */
const LEAFLET_CSS = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
let cssRequested = false;

export function ensureLeafletCss(): void {
  if (cssRequested) return;
  cssRequested = true;
  SPComponentLoader.loadCss(LEAFLET_CSS);
}

export const PICKUP_FILL = '#2E7D32';
export const DROPOFF_FILL = '#C62828';

/**
 * A map pin, drawn as inline SVG.
 *
 * NOT the default marker images. Leaflet resolves those by relative URL, which
 * the SPFx bundler breaks - you get a working map with invisible markers. A
 * divIcon sidesteps the asset pipeline completely, and two visually distinct
 * pins on one map were required anyway.
 *
 * Passing our own className also drops Leaflet's `leaflet-div-icon` class,
 * which is what would otherwise draw a white box behind each pin.
 */
export function makePin(fill: string, label: string): L.DivIcon {
  return L.divIcon({
    className: 'fleet-pin',
    html:
      `<svg width="26" height="36" viewBox="0 0 26 36" xmlns="http://www.w3.org/2000/svg" ` +
      `role="img" aria-label="${label}">` +
      `<path d="M13 0C5.8 0 0 5.8 0 13c0 9.8 13 23 13 23s13-13.2 13-23C26 5.8 20.2 0 13 0z" ` +
      `fill="${fill}"/><circle cx="13" cy="13" r="5" fill="#fff"/></svg>`,
    iconSize: [26, 36],
    iconAnchor: [13, 36],
    popupAnchor: [0, -36],
  });
}

export interface ITileCallbacks {
  /** TomTom refused; OpenStreetMap is being used instead. */
  onFallback?: () => void;
  /** Neither provider would serve a tile. */
  onFailure?: () => void;
}

/**
 * Add the tile layer, falling back to OpenStreetMap if TomTom refuses.
 *
 * A TomTom key is entitled per product: a key can be authorised for Search and
 * NOT for Maps, which returns 403 on every tile and leaves an empty map while
 * the search works perfectly. Rather than show a blank panel, swap the layer
 * once and carry on.
 *
 * Search, pins and coordinates stay TomTom - only the backdrop changes, so
 * nothing about the stored data is affected. Once the Maps product is enabled
 * on the key this never fires.
 *
 * OSM's tile policy expects light, attributed use. If the entitlement is not
 * coming, move to a paid provider rather than leaning on it permanently.
 */
export function addTileLayer(map: L.Map, apiKey: string, cb: ITileCallbacks = {}): void {
  const tomtom = L.tileLayer(
    `https://api.tomtom.com/map/1/tile/basic/main/{z}/{x}/{y}.png?key=${apiKey}`,
    {
      // A licence condition. Leave it visible.
      attribution: '&copy; TomTom',
      maxZoom: 18,
    },
  );

  let swapped = false;
  tomtom.on('tileerror', () => {
    if (swapped) return;
    swapped = true;
    console.warn(
      '[FleetPortal] TomTom tiles returned an error - the key is probably not ' +
        'entitled to the Maps product. Falling back to OpenStreetMap.',
    );
    map.removeLayer(tomtom);
    if (cb.onFallback) cb.onFallback();

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors',
      maxZoom: 18,
    })
      .on('tileerror', () => {
        if (cb.onFailure) cb.onFailure();
      })
      .addTo(map);
  });

  tomtom.addTo(map);
}
