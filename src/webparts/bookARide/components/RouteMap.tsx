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
 * not a driving route. Drawing a road-following route would need a Routing
 * API call and would imply an ETA the solution does not calculate.
 */

import * as L from 'leaflet';
import * as React from 'react';
import {
  addTileLayer,
  DROPOFF_FILL,
  ensureLeafletCss,
  makePin,
  PICKUP_FILL,
} from './leafletBase';
import styles from './BookARide.module.scss';

export interface IRouteMapProps {
  apiKey: string;
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
    apiKey, isActive,
    pickupLat, pickupLng, pickupText,
    dropOffLat, dropOffLng, dropOffText,
  } = props;

  const containerRef = React.useRef<HTMLDivElement>(null);
  const mapRef = React.useRef<L.Map | undefined>(undefined);
  const [tilesFailed, setTilesFailed] = React.useState(false);
  const [usingFallbackTiles, setUsingFallbackTiles] = React.useState(false);

  // --- create once, and DESTROY ON UNMOUNT ---------------------------------
  React.useEffect(() => {
    ensureLeafletCss();
    if (!containerRef.current || mapRef.current) return undefined;

    const map = L.map(containerRef.current, {
      center: [pickupLat, pickupLng],
      zoom: 12,
      zoomControl: true,
      // Read-only: this is a picture of the trip, not a control.
      scrollWheelZoom: false,
    });

    addTileLayer(map, apiKey, {
      onFallback: () => setUsingFallbackTiles(true),
      onFailure: () => setTilesFailed(true),
    });

    L.marker([pickupLat, pickupLng], {
      icon: makePin(PICKUP_FILL, 'Pickup point'),
      interactive: false,
      keyboard: false,
    })
      .addTo(map)
      .bindPopup(pickupText);

    L.marker([dropOffLat, dropOffLng], {
      icon: makePin(DROPOFF_FILL, 'Destination'),
      interactive: false,
      keyboard: false,
    })
      .addTo(map)
      .bindPopup(dropOffText);

    L.polyline(
      [
        [pickupLat, pickupLng],
        [dropOffLat, dropOffLng],
      ],
      { color: '#2F5D9F', weight: 3, opacity: 0.85, dashArray: '6 6' },
    ).addTo(map);

    map.fitBounds(
      L.latLngBounds([
        [pickupLat, pickupLng],
        [dropOffLat, dropOffLng],
      ]),
      { padding: [34, 34], maxZoom: 14 },
    );

    mapRef.current = map;

    return () => {
      // Without this, returning to the screen a second time throws
      // "Map container is already initialized" and it renders blank.
      map.remove();
      mapRef.current = undefined;
    };
    // The trip's endpoints do not change while this is on screen.
  }, []);

  // --- the map has no size until the screen is actually visible ------------
  React.useEffect(() => {
    if (!isActive || !mapRef.current) return undefined;
    const map = mapRef.current;
    const handle = window.setTimeout(() => map.invalidateSize(), 0);
    return () => window.clearTimeout(handle);
  }, [isActive]);

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
        {tilesFailed ? (
          <div className={styles.mapFallback}>
            The map could not load. Your pickup and destination are below.
          </div>
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

      {usingFallbackTiles ? (
        <p className={styles.routeMapNote}>
          Showing OpenStreetMap &mdash; the TomTom key is not enabled for maps.
        </p>
      ) : undefined}
    </section>
  );
}

export default RouteMap;
