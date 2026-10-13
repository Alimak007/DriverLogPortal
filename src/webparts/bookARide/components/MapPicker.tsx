/**
 * Pick a pickup and a destination anywhere in the UAE.
 *
 * Replaces the Locations list entirely: the requester searches live with
 * TomTom and drops a pin, and what we store is free text plus coordinates.
 *
 * TWO SEARCHES, TWO PINS, ONE MAP. Pickup is green, drop-off is red, and both
 * are draggable.
 */

import * as L from 'leaflet';
import * as React from 'react';
import { FormRow, Notice, TextArea, TextField } from '../../../common/components/Shared';
import { Icon } from '../../../common/components/Icons';
import {
  addTileLayer,
  DROPOFF_FILL,
  ensureLeafletCss,
  makePin,
  PICKUP_FILL,
} from './leafletBase';
import styles from './BookARide.module.scss';

/** What a chosen place looks like once it reaches the request. */
export interface IPlace {
  text: string;
  placeId?: string;
  lat: number;
  lng: number;
}

export interface IMapPickerProps {
  apiKey: string;
  /** True when the wizard is actually showing this step. Drives invalidateSize. */
  isActive: boolean;
  centre: { lat: number; lng: number };
  radiusKm: number;

  pickup?: IPlace;
  dropOff?: IPlace;
  pickupNotes: string;
  dropOffNotes: string;

  onPickup: (place: IPlace | undefined) => void;
  onDropOff: (place: IPlace | undefined) => void;
  onPickupNotes: (notes: string) => void;
  onDropOffNotes: (notes: string) => void;
}

// -----------------------------------------------------------------------------
// Leaflet setup
// -----------------------------------------------------------------------------


// -----------------------------------------------------------------------------
// TomTom
// -----------------------------------------------------------------------------

interface ISuggestion {
  text: string;
  placeId?: string;
  lat: number;
  lng: number;
}

type SearchState = 'idle' | 'searching' | 'none' | 'forbidden' | 'quota' | 'error';

/**
 * Fuzzy Search, not Geocoding.
 *
 * Geocoding expects a well-formed address; fuzzy search copes with the
 * partial, misspelled input people actually type.
 *
 * NOTE: a plain fetch, deliberately. SPHttpClient attaches SharePoint bearer
 * tokens and request digests - pointing it at api.tomtom.com would send the
 * tenant's auth material to a third party. That is a security incident, not a
 * config mistake. Never use SPHttpClient for an external host.
 */
async function tomtomSearch(
  apiKey: string,
  query: string,
  centre: { lat: number; lng: number },
  radiusKm: number,
): Promise<{ state: SearchState; results: ISuggestion[] }> {
  const url =
    `https://api.tomtom.com/search/2/search/${encodeURIComponent(query)}.json` +
    `?key=${encodeURIComponent(apiKey)}` +
    `&countrySet=AE` +
    `&lat=${centre.lat}&lon=${centre.lng}` +
    `&radius=${Math.round(radiusKm * 1000)}` +
    `&limit=5&typeahead=true`;

  let res: Response;
  try {
    res = await fetch(url);
  } catch (e) {
    console.error('[FleetPortal] TomTom search unreachable', e);
    return { state: 'error', results: [] };
  }

  if (res.status === 403) {
    console.error('[FleetPortal] TomTom returned 403. Check the key and its referrer restriction.');
    return { state: 'forbidden', results: [] };
  }
  if (res.status === 429) {
    console.error('[FleetPortal] TomTom returned 429 - the daily quota is spent.');
    return { state: 'quota', results: [] };
  }
  if (!res.ok) {
    console.error('[FleetPortal] TomTom search failed', res.status);
    return { state: 'error', results: [] };
  }

  const body: { results?: { id?: string; address?: { freeformAddress?: string }; position?: { lat: number; lon: number } }[] } =
    await res.json();

  const results: ISuggestion[] = (body.results ?? [])
    .filter((r) => r.position && r.address?.freeformAddress)
    .map((r) => ({
      text: r.address?.freeformAddress as string,
      placeId: r.id,
      lat: (r.position as { lat: number }).lat,
      lng: (r.position as { lon: number }).lon,
    }));

  return { state: results.length === 0 ? 'none' : 'idle', results };
}

/** Make the text match the pin after a drag. */
async function reverseGeocode(apiKey: string, lat: number, lng: number): Promise<string | undefined> {
  try {
    const res = await fetch(
      `https://api.tomtom.com/search/2/reverseGeocode/${lat},${lng}.json?key=${encodeURIComponent(apiKey)}`,
    );
    if (!res.ok) return undefined;
    const body: { addresses?: { address?: { freeformAddress?: string } }[] } = await res.json();
    return body.addresses?.[0]?.address?.freeformAddress;
  } catch {
    return undefined;
  }
}

// -----------------------------------------------------------------------------
// One search field
// -----------------------------------------------------------------------------

function PlaceSearch({
  id,
  label,
  placeholder,
  apiKey,
  centre,
  radiusKm,
  value,
  onPick,
  notes,
  onNotes,
  notesPlaceholder,
}: {
  id: string;
  label: string;
  placeholder: string;
  apiKey: string;
  centre: { lat: number; lng: number };
  radiusKm: number;
  value?: IPlace;
  onPick: (p: IPlace | undefined) => void;
  notes: string;
  onNotes: (n: string) => void;
  notesPlaceholder: string;
}): JSX.Element {
  const [query, setQuery] = React.useState('');
  const [results, setResults] = React.useState<ISuggestion[]>([]);
  const [state, setState] = React.useState<SearchState>('idle');
  const [open, setOpen] = React.useState(false);
  const seq = React.useRef(0);

  // DEBOUNCED AT 300 ms, and nothing under 3 characters. Without both you fire
  // a request per keystroke and burn the free tier on prefixes.
  React.useEffect(() => {
    const q = query.trim();
    if (q.length < 3) {
      setResults([]);
      setState('idle');
      return undefined;
    }
    if (!apiKey) {
      setState('forbidden');
      return undefined;
    }

    const mine = ++seq.current;
    setState('searching');
    const handle = window.setTimeout(() => {
      tomtomSearch(apiKey, q, centre, radiusKm)
        .then((r) => {
          // A slow earlier request must not overwrite a newer one.
          if (mine !== seq.current) return;
          setResults(r.results);
          setState(r.state);
          setOpen(true);
        })
        .catch(() => undefined);
    }, 300);

    return () => window.clearTimeout(handle);
  }, [query, apiKey, centre, radiusKm]);

  const message =
    state === 'none'
      ? 'No match. Try a landmark or a nearby road.'
      : state === 'forbidden'
        ? 'Location search is unavailable. Contact IT.'
        : state === 'quota'
          ? 'Location search is unavailable. Contact IT.'
          : state === 'error'
            ? 'Location search could not be reached. You can still drag the pin on the map.'
            : undefined;

  return (
    <div className={styles.placeField}>
      <FormRow label={label} required htmlFor={id} error={undefined}>
        <div className={styles.placeInputWrap}>
          <TextField
            id={id}
            value={query}
            placeholder={value ? value.text : placeholder}
            autoComplete="off"
            onChange={(e) => setQuery(e.target.value)}
            onFocus={() => setOpen(true)}
          />
          {state === 'searching' ? (
            <span className={styles.placeSpinner} aria-hidden="true" />
          ) : undefined}
        </div>
      </FormRow>

      {value && !open ? (
        <div className={styles.placeChosen}>
          <Icon name="pin" />
          <span>{value.text}</span>
          <button
            type="button"
            className={styles.placeClear}
            onClick={() => {
              onPick(undefined);
              setQuery('');
            }}
          >
            Change
          </button>
        </div>
      ) : undefined}

      {open && (results.length > 0 || message) ? (
        <ul className={styles.placeList} aria-label={`Matches for ${label}`}>
          {message ? <li className={styles.placeEmpty}>{message}</li> : undefined}
          {results.map((r) => (
            <li key={`${r.placeId ?? ''}${r.lat},${r.lng}`}>
              <button
                type="button"
                className={styles.placeOption}
                onClick={() => {
                  onPick({ text: r.text, placeId: r.placeId, lat: r.lat, lng: r.lng });
                  setQuery('');
                  setOpen(false);
                }}
              >
                <Icon name="pin" />
                <span>{r.text}</span>
              </button>
            </li>
          ))}
        </ul>
      ) : undefined}

      {/*
        NOT an optional extra. A list entry called "Khalifa Plant Gate 2" told
        a driver which gate; a pin reading "KIZAD, Abu Dhabi" does not. Visible
        without expanding anything, here and on the driver's cards.
      */}
      <FormRow label="Notes for the driver" htmlFor={`${id}-notes`}>
        <TextArea
          id={`${id}-notes`}
          rows={2}
          maxLength={255}
          value={notes}
          placeholder={notesPlaceholder}
          onChange={(e) => onNotes(e.target.value)}
        />
      </FormRow>
    </div>
  );
}

// -----------------------------------------------------------------------------
// The picker
// -----------------------------------------------------------------------------

export function MapPicker(props: IMapPickerProps): JSX.Element {
  const {
    apiKey, isActive, centre, radiusKm,
    pickup, dropOff, pickupNotes, dropOffNotes,
    onPickup, onDropOff, onPickupNotes, onDropOffNotes,
  } = props;

  const containerRef = React.useRef<HTMLDivElement>(null);
  const mapRef = React.useRef<L.Map | undefined>(undefined);
  const pickupMarker = React.useRef<L.Marker | undefined>(undefined);
  const dropOffMarker = React.useRef<L.Marker | undefined>(undefined);
  const [tilesFailed, setTilesFailed] = React.useState(false);
  const [usingFallbackTiles, setUsingFallbackTiles] = React.useState(false);

  // Keep the latest callbacks reachable from the Leaflet event handlers, which
  // are bound once and would otherwise close over the first render's props.
  const onPickupRef = React.useRef(onPickup);
  const onDropOffRef = React.useRef(onDropOff);
  onPickupRef.current = onPickup;
  onDropOffRef.current = onDropOff;

  // --- create once, and DESTROY ON UNMOUNT ---------------------------------
  React.useEffect(() => {
    ensureLeafletCss();
    if (!containerRef.current || mapRef.current) return undefined;

    const map = L.map(containerRef.current, {
      center: [centre.lat, centre.lng],
      zoom: 11,
      zoomControl: true,
    });

    addTileLayer(map, apiKey, {
      onFallback: () => setUsingFallbackTiles(true),
      onFailure: () => setTilesFailed(true),
    });

    mapRef.current = map;

    return () => {
      // Without this, coming back to the step a second time throws
      // "Map container is already initialized" and the step renders blank.
      // Grey tiles look broken; this one looks like a crash.
      map.remove();
      mapRef.current = undefined;
      pickupMarker.current = undefined;
      dropOffMarker.current = undefined;
    };
    // Created once. The centre and key do not change within a booking.
  }, []);

  // --- the map has no size until the step is actually visible --------------
  React.useEffect(() => {
    if (!isActive || !mapRef.current) return undefined;
    const map = mapRef.current;
    const handle = window.setTimeout(() => map.invalidateSize(), 0);
    return () => window.clearTimeout(handle);
  }, [isActive]);

  // --- keep the pins in step with the chosen places ------------------------
  React.useEffect(() => {
    const map = mapRef.current;
    if (!map) return;

    const sync = (
      ref: React.MutableRefObject<L.Marker | undefined>,
      place: IPlace | undefined,
      fill: string,
      label: string,
      commit: (p: IPlace) => void,
    ): void => {
      if (!place) {
        if (ref.current) {
          map.removeLayer(ref.current);
          ref.current = undefined;
        }
        return;
      }
      if (!ref.current) {
        const marker = L.marker([place.lat, place.lng], {
          icon: makePin(fill, label),
          draggable: true,
          title: label,
        }).addTo(map);

        marker.on('dragend', () => {
          const pos = marker.getLatLng();
          // Reverse geocode so the text matches where the pin now is. There is
          // no id on a reverse geocode, so PlaceId is left empty - expected,
          // and the column is optional.
          reverseGeocode(apiKey, pos.lat, pos.lng)
            .then((text) =>
              commit({ text: text ?? `${pos.lat.toFixed(5)}, ${pos.lng.toFixed(5)}`, lat: pos.lat, lng: pos.lng }),
            )
            .catch(() => undefined);
        });
        ref.current = marker;
      } else {
        ref.current.setLatLng([place.lat, place.lng]);
      }
    };

    sync(pickupMarker, pickup, PICKUP_FILL, 'Pickup point', (p) => onPickupRef.current(p));
    sync(dropOffMarker, dropOff, DROPOFF_FILL, 'Destination', (p) => onDropOffRef.current(p));

    if (pickup && dropOff) {
      map.fitBounds(
        L.latLngBounds([
          [pickup.lat, pickup.lng],
          [dropOff.lat, dropOff.lng],
        ]),
        { padding: [40, 40], maxZoom: 15 },
      );
    } else if (pickup) {
      map.setView([pickup.lat, pickup.lng], 14);
    } else if (dropOff) {
      map.setView([dropOff.lat, dropOff.lng], 14);
    }
  }, [pickup, dropOff, apiKey]);

  return (
    <div className={styles.mapPicker}>
      <div className={styles.mapPane}>
        <div ref={containerRef} className={styles.mapCanvas} role="application" aria-label="Pickup and destination map" />
        {tilesFailed ? (
          // The map is an aid; the coordinates are the data. A tile failure
          // must never block a booking that already has both pins.
          <div className={styles.mapFallback}>
            The map could not load, but your chosen places are still saved.
          </div>
        ) : undefined}
        <p className={styles.mapHint}>
          Search below, or drag a pin to adjust it. Green is the pickup, red is the destination.
          {usingFallbackTiles ? ' Showing OpenStreetMap - the TomTom key is not enabled for maps.' : ''}
        </p>
      </div>

      <div className={styles.placeFields}>
        <PlaceSearch
          id="pickup-search"
          label="Pickup point"
          placeholder="Search a place, building or road"
          apiKey={apiKey}
          centre={centre}
          radiusKm={radiusKm}
          value={pickup}
          onPick={onPickup}
          notes={pickupNotes}
          onNotes={onPickupNotes}
          notesPlaceholder="Which gate, entrance or floor?"
        />
        <PlaceSearch
          id="dropoff-search"
          label="Destination"
          placeholder="Search a place, building or road"
          apiKey={apiKey}
          centre={centre}
          radiusKm={radiusKm}
          value={dropOff}
          onPick={onDropOff}
          notes={dropOffNotes}
          onNotes={onDropOffNotes}
          notesPlaceholder="Which gate, entrance or floor?"
        />
      </div>

      {!apiKey ? (
        <Notice tone="attention">
          No TomTom key is configured, so the map and search are unavailable.
          Set it in the web part property pane.
        </Notice>
      ) : undefined}
    </div>
  );
}

export default MapPicker;
