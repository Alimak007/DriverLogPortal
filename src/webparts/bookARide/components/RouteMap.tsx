/**
 * The route on the confirmation and tracking screens.
 *
 * TWO PINS AND A LINE. NOTHING ELSE.
 *
 * No vehicle position, no moving car, no "driver is here" marker - and the
 * heading says "Your route", not "Track my ride". There is no live tracking in
 * phase 1, and a map that looks like it is tracking something WILL be read as
 * tracking something. The requester would then watch a stationary car and
 * conclude the driver had stopped, which is worse than showing no map at all.
 *
 * Trip progress is conveyed by the status timeline above this, in words.
 *
 * The straight line is deliberate too: it is a connector between two points,
 * not a driving route. Google's Directions API could return the real road path
 * now that the solution is on Google, but drawing one would imply an ETA the
 * solution does not calculate and a distance nothing stores. That is a product
 * decision, not a missing feature.
 */

import * as React from 'react';
import {
  addMarker,
  createMap,
  DROPOFF_FILL,
  ensureGoogleMaps,
  fitTwoPoints,
  IMapMarker,
  PICKUP_FILL,
  refreshMapSize,
} from './googleMapsBase';
import styles from './BookARide.module.scss';

export interface IRouteMapProps {
  apiKey: string;
  /** Cloud Map ID. Optional - without it the classic marker is used. */
  mapId?: string;
  /** True when the screen showing this map is actually visible. */
  isActive: boolean;
  pickupLat: number;
  pickupLng: number;
  pickupText: string;
  dropOffLat: number;
  dropOffLng: number;
  dropOffText: string;
}

export function RouteMap(props: IRouteMapProps): JSX.Element {
  const {
    apiKey, mapId, isActive,
    pickupLat, pickupLng, pickupText,
    dropOffLat, dropOffLng, dropOffText,
  } = props;

  const containerRef = React.useRef<HTMLDivElement>(null);
  const mapRef = React.useRef<google.maps.Map | undefined>(undefined);
  const markersRef = React.useRef<IMapMarker[]>([]);
  const lineRef = React.useRef<google.maps.Polyline | undefined>(undefined);

  const [ready, setReady] = React.useState(false);
  const [mapFailed, setMapFailed] = React.useState(false);

  // --- load the API once (shared with the picker) --------------------------
  React.useEffect(() => {
    let cancelled = false;
    ensureGoogleMaps(apiKey)
      .then(() => {
        if (!cancelled) setReady(true);
      })
      .catch((e: Error) => {
        console.error('[FleetPortal] Google Maps could not load', e);
        if (!cancelled) setMapFailed(true);
      });
    return () => {
      cancelled = true;
    };
  }, [apiKey]);

  // --- draw once, and tear down on unmount ---------------------------------
  React.useEffect(() => {
    if (!ready || !containerRef.current || mapRef.current) return undefined;

    const map = createMap({
      container: containerRef.current,
      centre: { lat: pickupLat, lng: pickupLng },
      zoom: 12,
      mapId,
      // Read-only: this is a picture of the trip, not a control.
      scrollwheel: false,
    });

    markersRef.current = [
      addMarker({
        map,
        lat: pickupLat,
        lng: pickupLng,
        fill: PICKUP_FILL,
        label: 'Pickup point',
        clickable: false,
      }),
      addMarker({
        map,
        lat: dropOffLat,
        lng: dropOffLng,
        fill: DROPOFF_FILL,
        label: 'Destination',
        clickable: false,
      }),
    ];

    /**
     * Dashed, via a repeating symbol.
     *
     * Google has no strokeDasharray: a dashed line is a fully transparent
     * stroke plus a dash glyph repeated along it. The dashes carry the meaning
     * here - they are what stops this reading as a driving route.
     *
     * The numbers reproduce Leaflet's `dashArray: '6 6'` exactly. The glyph
     * path spans 2 units, so `scale: 3` draws a 6px dash, and `repeat: 12px`
     * leaves a 6px gap after it. Changing one without the other turns the line
     * solid or scatters it into dots.
     */
    lineRef.current = new google.maps.Polyline({
      map,
      path: [
        { lat: pickupLat, lng: pickupLng },
        { lat: dropOffLat, lng: dropOffLng },
      ],
      strokeOpacity: 0,
      clickable: false,
      icons: [
        {
          icon: {
            path: 'M 0,-1 0,1',
            strokeColor: '#2F5D9F',
            strokeOpacity: 0.85,
            strokeWeight: 3,
            scale: 3,
          },
          offset: '0',
          repeat: '12px',
        },
      ],
    });

    fitTwoPoints(
      map,
      { lat: pickupLat, lng: pickupLng },
      { lat: dropOffLat, lng: dropOffLng },
      34,
      14,
    );

    mapRef.current = map;

    return () => {
      // Google has no map.remove(). Detach everything that holds a listener,
      // or returning to this screen a second time leaves the old handlers
      // bound to a node React has already thrown away.
      markersRef.current.forEach((m) => m.remove());
      markersRef.current = [];
      if (lineRef.current) {
        lineRef.current.setMap(null);
        lineRef.current = undefined;
      }
      google.maps.event.clearInstanceListeners(map);
      mapRef.current = undefined;
    };
    // The trip's endpoints do not change while this is on screen.
  }, [ready]);

  // --- the map has no size until the screen is actually visible ------------
  React.useEffect(() => {
    if (!isActive || !mapRef.current) return undefined;
    const map = mapRef.current;
    const handle = window.setTimeout(() => refreshMapSize(map), 0);
    return () => window.clearTimeout(handle);
  }, [isActive, ready]);

  return (
    <section className={styles.routeMapCard} aria-label="Your route">
      <h3 className={styles.routeMapHeading}>Your route</h3>

      <div className={styles.routeMapPane}>
        <div
          ref={containerRef}
          className={styles.routeMapCanvas}
          role="img"
          aria-label={`Route from ${pickupText} to ${dropOffText}`}
        />
        {mapFailed ? (
          <div className={styles.mapFallback}>
            The map could not load. Your pickup and destination are below.
          </div>
        ) : !ready ? (
          <div className={styles.mapFallback}>Loading the map&hellip;</div>
        ) : undefined}
      </div>

      {/*
        The addresses in words, always - not only when the map fails. A map
        alone is unreadable to a screen reader and useless on a bad connection,
        and the driver's notes matter more than the picture does.
      */}
      <ul className={styles.routeLegend}>
        <li>
          <span className={styles.routeDot} style={{ background: PICKUP_FILL }} aria-hidden="true" />
          <span>
            <strong>From</strong> {pickupText}
          </span>
        </li>
        <li>
          <span className={styles.routeDot} style={{ background: DROPOFF_FILL }} aria-hidden="true" />
          <span>
            <strong>To</strong> {dropOffText}
          </span>
        </li>
      </ul>
    </section>
  );
}

export default RouteMap;
