/**
 * The Google Maps bits both maps share: the script loader, the pins, and the
 * marker wrapper.
 *
 * Two components draw maps - the booking picker and the route view on the
 * confirmation screen - and the Maps JavaScript API must be loaded EXACTLY
 * ONCE per page. Including it twice logs "You have included the Google Maps
 * JavaScript API multiple times" and can leave the second map half-built, so
 * the loader below is a singleton shared by both.
 *
 * Replaces the previous Leaflet + TomTom pairing. One thing is genuinely lost
 * in that swap: Leaflet could fall back to OpenStreetMap tiles when TomTom
 * refused. There is no equivalent here and there cannot be - Google's terms
 * require Google content to be shown on a Google map, so a fallback basemap
 * from another provider is not an option. If the script does not load, the
 * maps degrade to the text addresses, which both components already render.
 *
 * EVERYTHING HERE GOES THROUGH THE JS API, never Google's REST endpoints. That
 * is not a style preference: a referrer-restricted browser key is REFUSED
 * outright by the Geocoding web service ("API keys with referer restrictions
 * cannot be used with this API"). The JS API authenticates by referrer, which
 * is exactly what this key is set up for.
 */

export const PICKUP_FILL = '#2E7D32';
export const DROPOFF_FILL = '#C62828';

/**
 * The stable release channel, not `weekly`.
 *
 * `weekly` picks up Google's changes the week they ship, which is the wrong
 * trade for a line-of-business form: a breaking change arrives without anyone
 * having deployed anything. `quarterly` lags but is predictable.
 *
 * If a class used here turns out to be missing at runtime, this is the line to
 * bump - see the guard in `hasNewPlaces()`.
 */
const API_VERSION = 'quarterly';

/**
 * `places` for the address search, `marker` for AdvancedMarkerElement, and
 * `geocoding` for the reverse lookup after a pin drag. Listing them here means
 * they are ready when the callback fires, so components never await
 * importLibrary() mid-render.
 */
const LIBRARIES = 'places,marker,geocoding';

const CALLBACK_NAME = '__fleetPortalGoogleMapsReady';
const SCRIPT_ID = 'fleet-portal-google-maps';

let loadPromise: Promise<void> | undefined;

/** True once google.maps is usable. */
export function isGoogleMapsReady(): boolean {
  return typeof google !== 'undefined' && typeof google.maps !== 'undefined';
}

/**
 * Is the new Places API present?
 *
 * Google runs two generations of Places side by side. The legacy
 * AutocompleteService is closed to new Cloud projects, so this code targets
 * the new AutocompleteSuggestion - but if the pinned release channel ever
 * predates it, the failure would otherwise be an undefined-property crash
 * inside a keystroke handler. Check explicitly and report it instead.
 */
export function hasNewPlaces(): boolean {
  return (
    isGoogleMapsReady() &&
    typeof google.maps.places !== 'undefined' &&
    typeof google.maps.places.AutocompleteSuggestion !== 'undefined'
  );
}

/** True when Advanced Markers are available at all. */
export function hasAdvancedMarkers(): boolean {
  return (
    isGoogleMapsReady() &&
    typeof google.maps.marker !== 'undefined' &&
    typeof google.maps.marker.AdvancedMarkerElement !== 'undefined'
  );
}

/**
 * Load the Maps JavaScript API once, and resolve when it is usable.
 *
 * NOTE: a plain script tag, deliberately - the same reasoning that applied to
 * the old TomTom calls. SPHttpClient attaches SharePoint bearer tokens and
 * request digests; pointing anything SharePoint-aware at a Google host would
 * send the tenant's auth material to a third party. That is a security
 * incident, not a config mistake. A script tag carries no tenant credentials.
 *
 * On failure the cached promise is cleared and the tag removed, so a later
 * mount can retry rather than inheriting a permanently rejected promise.
 */
export function ensureGoogleMaps(apiKey: string): Promise<void> {
  if (loadPromise) return loadPromise;

  loadPromise = new Promise<void>((resolve, reject) => {
    if (isGoogleMapsReady()) {
      resolve();
      return;
    }

    if (!apiKey) {
      reject(new Error('No Google Maps API key is configured.'));
      return;
    }

    const globals = window as unknown as Record<string, unknown>;

    const cleanUp = (): void => {
      delete globals[CALLBACK_NAME];
      loadPromise = undefined;
      const tag = document.getElementById(SCRIPT_ID);
      if (tag && tag.parentNode) tag.parentNode.removeChild(tag);
    };

    globals[CALLBACK_NAME] = () => {
      delete globals[CALLBACK_NAME];
      if (!hasNewPlaces()) {
        console.warn(
          '[FleetPortal] Google Maps loaded but the new Places API is missing. ' +
            'Confirm "Places API (New)" is enabled on the key, or raise ' +
            'API_VERSION in googleMapsBase.ts.',
        );
      }
      resolve();
    };

    const script = document.createElement('script');
    script.id = SCRIPT_ID;
    // `loading=async` is what suppresses Google's own performance warning.
    script.src =
      'https://maps.googleapis.com/maps/api/js' +
      `?key=${encodeURIComponent(apiKey)}` +
      `&libraries=${LIBRARIES}` +
      `&v=${API_VERSION}` +
      `&callback=${CALLBACK_NAME}` +
      '&loading=async' +
      '&language=en' +
      // Biases geocoding and address formatting to the UAE.
      '&region=AE';
    script.async = true;
    script.onerror = () => {
      cleanUp();
      reject(
        new Error(
          'The Google Maps script could not be loaded. The key may be invalid, ' +
            'or maps.googleapis.com may be blocked on this network.',
        ),
      );
    };

    document.head.appendChild(script);
  });

  return loadPromise;
}

// -----------------------------------------------------------------------------
// Pins
// -----------------------------------------------------------------------------

/**
 * The pin, as inline SVG - the same artwork the Leaflet build used, so the
 * booking form looks unchanged.
 *
 * Two visually distinct pins on one map were a requirement, and Google's
 * default red teardrop cannot express that. Inline SVG also sidesteps the SPFx
 * asset pipeline, which rewrites relative image URLs and would otherwise leave
 * working markers that render nothing.
 */
function pinSvg(fill: string, label: string): string {
  return (
    `<svg width="26" height="36" viewBox="0 0 26 36" xmlns="http://www.w3.org/2000/svg" ` +
    `role="img" aria-label="${label}">` +
    `<path d="M13 0C5.8 0 0 5.8 0 13c0 9.8 13 23 13 23s13-13.2 13-23C26 5.8 20.2 0 13 0z" ` +
    `fill="${fill}"/><circle cx="13" cy="13" r="5" fill="#fff"/></svg>`
  );
}

/** The pin as a DOM node, for AdvancedMarkerElement's `content`. */
function makePinElement(fill: string, label: string): HTMLElement {
  const host = document.createElement('div');
  host.className = 'fleet-pin';
  host.innerHTML = pinSvg(fill, label);
  return host;
}

/** The pin as a data URI, for the classic Marker's `icon`. */
function pinDataUri(fill: string, label: string): string {
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(pinSvg(fill, label))}`;
}

// -----------------------------------------------------------------------------
// Markers
// -----------------------------------------------------------------------------

/**
 * Which maps were built with a Map ID, and so can carry Advanced Markers.
 *
 * A WeakSet rather than a flag on the map: it holds no strong reference, so a
 * map discarded when its wizard step unmounts is still collectable.
 *
 * Declared ABOVE addMarker on purpose - the lint rule in this repo rejects a
 * reference to a const declared further down the file.
 */
const mapsWithMapId = new WeakSet<google.maps.Map>();

/**
 * One marker, whichever generation drew it.
 *
 * The two components should not care which path was taken, so the branch lives
 * here and they get the same two operations either way.
 */
export interface IMapMarker {
  setPosition(lat: number, lng: number): void;
  remove(): void;
}

export interface IMarkerOptions {
  map: google.maps.Map;
  lat: number;
  lng: number;
  fill: string;
  label: string;
  draggable?: boolean;
  /** Suppresses the click target on a read-only map. */
  clickable?: boolean;
  onDragEnd?: (lat: number, lng: number) => void;
}

/**
 * Normalise the several shapes Google hands back for a position.
 *
 * Typed as `unknown` rather than the real union on purpose. Google returns
 * LatLng, LatLngLiteral, LatLngAltitude or null depending on which API and
 * which release produced it, and on LatLng `lat` is a METHOD while on the
 * literal it is a NUMBER. Narrowing here once is what keeps that distinction
 * out of both components.
 */
function readLatLng(value: unknown): { lat: number; lng: number } | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const raw = value as { lat?: number | (() => number); lng?: number | (() => number) };
  const lat = typeof raw.lat === 'function' ? raw.lat() : raw.lat;
  const lng = typeof raw.lng === 'function' ? raw.lng() : raw.lng;
  if (typeof lat !== 'number' || typeof lng !== 'number') return undefined;
  return { lat, lng };
}

/**
 * Add a marker, preferring Advanced Markers and falling back to the classic
 * one.
 *
 * WHY THE FALLBACK: AdvancedMarkerElement renders nothing unless the map was
 * created with a Map ID. A missing Map ID gives a console message and an empty
 * map - exactly the invisible-marker failure the Leaflet build hit by a
 * different route. Rather than depend on that one console setting having been
 * done, fall back to the classic Marker, which needs no Map ID and takes the
 * same SVG as a data URI. The map still works; only the marker API differs.
 */
export function addMarker(options: IMarkerOptions): IMapMarker {
  const { map, lat, lng, fill, label, draggable, clickable, onDragEnd } = options;
  // Not `map.get('mapId')`: a Map's constructor options are not all exposed as
  // MVCObject properties, so that read can come back undefined on a map that
  // definitely has a Map ID - which would silently downgrade every marker.
  // createMap() records the answer instead.
  const useAdvanced = hasAdvancedMarkers() && mapsWithMapId.has(map);

  if (useAdvanced) {
    const marker = new google.maps.marker.AdvancedMarkerElement({
      map,
      position: { lat, lng },
      content: makePinElement(fill, label),
      title: label,
      gmpDraggable: Boolean(draggable),
      gmpClickable: clickable === undefined ? true : clickable,
    });

    if (onDragEnd) {
      marker.addListener('dragend', () => {
        // Read it off the marker rather than the event: the event payload has
        // changed shape between releases, the marker's own position has not.
        const at = readLatLng(marker.position);
        if (at) onDragEnd(at.lat, at.lng);
      });
    }

    return {
      setPosition: (newLat: number, newLng: number) => {
        marker.position = { lat: newLat, lng: newLng };
      },
      remove: () => {
        google.maps.event.clearInstanceListeners(marker);
        marker.map = null;
      },
    };
  }

  const marker = new google.maps.Marker({
    map,
    position: { lat, lng },
    title: label,
    draggable: Boolean(draggable),
    clickable: clickable === undefined ? true : clickable,
    icon: {
      url: pinDataUri(fill, label),
      scaledSize: new google.maps.Size(26, 36),
      anchor: new google.maps.Point(13, 36),
    },
  });

  if (onDragEnd) {
    marker.addListener('dragend', () => {
      const at = readLatLng(marker.getPosition());
      if (at) onDragEnd(at.lat, at.lng);
    });
  }

  return {
    setPosition: (newLat: number, newLng: number) =>
      marker.setPosition({ lat: newLat, lng: newLng }),
    remove: () => {
      google.maps.event.clearInstanceListeners(marker);
      marker.setMap(null);
    },
  };
}

// -----------------------------------------------------------------------------
// The map itself
// -----------------------------------------------------------------------------

export interface IMapOptions {
  container: HTMLElement;
  centre: { lat: number; lng: number };
  zoom: number;
  mapId?: string;
  /** False on the read-only route view, so the page still scrolls over it. */
  scrollwheel?: boolean;
}

/** Both maps on a page would otherwise log the same warning twice. */
let warnedAboutMapId = false;

/**
 * Create a map with the chrome this solution wants and nothing else.
 *
 * `mapTypeControl` and `streetViewControl` are off: neither means anything on
 * a form whose job is to confirm two points, and both add controls a keyboard
 * user has to tab through to reach the fields below.
 */
export function createMap(options: IMapOptions): google.maps.Map {
  const { container, centre, zoom, mapId, scrollwheel } = options;

  if (!mapId && !warnedAboutMapId) {
    warnedAboutMapId = true;
    console.warn(
      '[FleetPortal] No Google Maps Map ID is configured, so the classic ' +
        'marker is used instead of Advanced Markers. Set one in the web part ' +
        'property pane to silence this.',
    );
  }

  const map = new google.maps.Map(container, {
    center: centre,
    zoom,
    mapId: mapId || undefined,
    zoomControl: true,
    mapTypeControl: false,
    streetViewControl: false,
    fullscreenControl: false,
    scrollwheel: scrollwheel === undefined ? true : scrollwheel,
    // Lets a touch user scroll the page without the map swallowing the gesture.
    gestureHandling: 'cooperative',
  });

  if (mapId) mapsWithMapId.add(map);
  return map;
}

/**
 * Frame two points, without zooming to street level when they are close.
 *
 * Google's fitBounds has no maxZoom, so two pins a hundred metres apart zoom
 * in until the map is a single junction and reads as broken. Leaflet took a
 * maxZoom argument; here the ceiling has to be applied after the fact, once
 * the map has settled.
 */
export function fitTwoPoints(
  map: google.maps.Map,
  a: { lat: number; lng: number },
  b: { lat: number; lng: number },
  paddingPx: number,
  maxZoom: number,
): void {
  const bounds = new google.maps.LatLngBounds();
  bounds.extend(a);
  bounds.extend(b);
  map.fitBounds(bounds, paddingPx);

  google.maps.event.addListenerOnce(map, 'idle', () => {
    const zoom = map.getZoom();
    if (zoom !== undefined && zoom > maxZoom) map.setZoom(maxZoom);
  });
}

// -----------------------------------------------------------------------------
// Road distance
// -----------------------------------------------------------------------------

/** What the Routes API tells us about a planned pickup -> drop-off trip. */
export interface IRouteEstimate {
  distanceKm: number;
  durationMinutes: number;
}

/**
 * Road distance and driving time between two points.
 *
 * ROUTES API OVER fetch, NOT google.maps.DirectionsService.
 *
 * DirectionsService calls the LEGACY Directions API, and Google has closed the
 * legacy APIs to projects created from 2025 onwards - this project included.
 * It fails with "You're calling a legacy API, which is not enabled for your
 * project", and no amount of enabling fixes it because the API cannot be
 * switched on. The Routes API is its replacement and must be called over REST.
 *
 * A plain fetch is correct here for the same reason it was for Places: an
 * SPHttpClient would attach SharePoint bearer tokens and request digests to a
 * Google host. Note also that this is a `connect-src` request, not a script,
 * so it is NOT governed by SharePoint's Trusted Script Sources list.
 *
 * TRAFFIC_UNAWARE is deliberate. Traffic-aware routing bills at a higher tier
 * and models traffic AS IT IS NOW - which says nothing useful about a trip
 * booked for tomorrow morning. The cheaper, stabler number is also the more
 * honest one.
 *
 * Returns undefined on any failure. A booking must never be blocked because
 * Google could not price a route.
 */
export async function routeDistance(
  apiKey: string,
  origin: { lat: number; lng: number },
  destination: { lat: number; lng: number },
): Promise<IRouteEstimate | undefined> {
  if (!apiKey) return undefined;

  try {
    const res = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': apiKey,
        // Without a field mask the Routes API refuses the request outright.
        // Asking for only these two also keeps it on the cheapest SKU.
        'X-Goog-FieldMask': 'routes.distanceMeters,routes.duration',
      },
      body: JSON.stringify({
        origin: { location: { latLng: { latitude: origin.lat, longitude: origin.lng } } },
        destination: {
          location: { latLng: { latitude: destination.lat, longitude: destination.lng } },
        },
        travelMode: 'DRIVE',
        routingPreference: 'TRAFFIC_UNAWARE',
        regionCode: 'AE',
        units: 'METRIC',
      }),
    });

    if (!res.ok) {
      const body = await res.text();
      console.warn(
        '[FleetPortal] Routes API returned ' +
          `${res.status}. If it says the API is disabled, enable "Routes API" ` +
          'in the Google Cloud console for this key.',
        body.slice(0, 300),
      );
      return undefined;
    }

    const body: {
      routes?: { distanceMeters?: number; duration?: string }[];
    } = await res.json();

    const route = body.routes?.[0];
    if (!route || route.distanceMeters === undefined) return undefined;

    // `duration` comes back as a protobuf duration string - "1239s".
    const seconds = parseInt(String(route.duration ?? '0').replace('s', ''), 10);

    return {
      distanceKm: Math.round((route.distanceMeters / 1000) * 10) / 10,
      durationMinutes: isNaN(seconds) ? 0 : Math.round(seconds / 60),
    };
  } catch (e) {
    console.warn('[FleetPortal] Routes API unreachable', e);
    return undefined;
  }
}

/**
 * Nudge a map that was created while its container was hidden.
 *
 * A map built inside a wizard step that is not on screen has no size, and
 * comes back grey when the step is shown. Google re-measures on a resize
 * event; re-centring afterwards stops the nudge from leaving it off-target.
 */
export function refreshMapSize(map: google.maps.Map): void {
  const centre = map.getCenter();
  google.maps.event.trigger(map, 'resize');
  if (centre) map.setCenter(centre);
}
