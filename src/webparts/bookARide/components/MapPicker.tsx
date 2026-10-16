/**
 * Pick a pickup and a destination anywhere in the UAE.
 *
 * Replaces the Locations list entirely: the requester searches live with
 * Google Places and drops a pin, and what we store is free text plus
 * coordinates.
 *
 * TWO SEARCHES, TWO PINS, ONE MAP. Pickup is green, drop-off is red, and both
 * are draggable.
 */

import * as React from 'react';
import { FormRow, Notice, TextArea, TextField } from '../../../common/components/Shared';
import { Icon } from '../../../common/components/Icons';
import {
  addMarker,
  createMap,
  DROPOFF_FILL,
  ensureGoogleMaps,
  fitTwoPoints,
  hasNewPlaces,
  IMapMarker,
  PICKUP_FILL,
  refreshMapSize,
} from './googleMapsBase';
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
  /** Cloud Map ID. Optional - without it the classic marker is used. */
  mapId?: string;
  /** True when the wizard is actually showing this step. Drives the resize. */
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
// Google Places
// -----------------------------------------------------------------------------

interface ISuggestion {
  key: string;
  primary: string;
  secondary?: string;
  /** Held so the coordinates can be fetched only if this one is chosen. */
  prediction: google.maps.places.PlacePrediction;
}

type SearchState =
  | 'idle'
  | 'searching'
  | 'resolving'
  | 'none'
  | 'forbidden'
  | 'quota'
  | 'error';

/**
 * A locationBias circle is capped at 50 km by the API.
 *
 * AllocationRules can legitimately hold a wider operating radius than that,
 * and passing it through unclamped makes the whole request fail rather than
 * widening the bias. Clamp, and let the bias be approximate - it only ranks
 * results, it does not restrict them. `includedRegionCodes` does the
 * restricting.
 */
const MAX_BIAS_RADIUS_M = 50000;

/** Longest place text we will store, matching the list column. */
const MAX_TEXT = 255;

/**
 * Turn whatever Google knows about a place into one line a driver can act on.
 *
 * `formattedAddress` alone loses the name - "Marina Mall" becomes a street and
 * a district, which is worse for a driver, not better. `displayName` alone
 * loses the location. Where the name is not already inside the address, both
 * are kept.
 */
function placeText(place: google.maps.places.Place, fallback: string): string {
  const name = place.displayName ?? undefined;
  const address = place.formattedAddress ?? undefined;

  let text: string;
  if (name && address) {
    text = address.indexOf(name) === -1 ? `${name}, ${address}` : address;
  } else {
    text = address ?? name ?? fallback;
  }

  return text.length > MAX_TEXT ? text.slice(0, MAX_TEXT) : text;
}

/** Map a thrown Places error onto the banner states the field already has. */
function classifyError(e: unknown): SearchState {
  const message = (e instanceof Error ? e.message : String(e)).toLowerCase();
  if (
    message.indexOf('denied') !== -1 ||
    message.indexOf('not authorized') !== -1 ||
    message.indexOf('referer') !== -1
  ) {
    return 'forbidden';
  }
  if (message.indexOf('quota') !== -1 || message.indexOf('over_query_limit') !== -1) {
    return 'quota';
  }
  return 'error';
}

/**
 * Autocomplete, not a text search.
 *
 * Predictions are cheap and carry NO COORDINATES - that is the design, not an
 * omission. The coordinates cost a separate Place Details call, which is why
 * it is deferred to `resolvePlace` below and only ever runs for the one
 * suggestion the requester actually picks. The old TomTom call returned
 * positions for all five results whether they were wanted or not.
 */
async function searchPlaces(
  query: string,
  centre: { lat: number; lng: number },
  radiusKm: number,
  sessionToken: google.maps.places.AutocompleteSessionToken,
): Promise<{ state: SearchState; results: ISuggestion[] }> {
  if (!hasNewPlaces()) return { state: 'forbidden', results: [] };

  try {
    const { suggestions } =
      await google.maps.places.AutocompleteSuggestion.fetchAutocompleteSuggestions({
        input: query,
        sessionToken,
        includedRegionCodes: ['ae'],
        locationBias: {
          center: centre,
          radius: Math.min(Math.round(radiusKm * 1000), MAX_BIAS_RADIUS_M),
        },
        language: 'en',
      });

    const results: ISuggestion[] = [];
    suggestions.forEach((s) => {
      const prediction = s.placePrediction;
      if (!prediction) return;
      results.push({
        key: prediction.placeId,
        primary: prediction.mainText?.text ?? prediction.text.text,
        secondary: prediction.secondaryText?.text,
        prediction,
      });
    });

    return { state: results.length === 0 ? 'none' : 'idle', results };
  } catch (e) {
    console.error('[FleetPortal] Google Places autocomplete failed', e);
    return { state: classifyError(e), results: [] };
  }
}

/**
 * Turn the chosen prediction into coordinates.
 *
 * This is the call that closes the autocomplete session for billing, which is
 * why the session token must be discarded by the caller afterwards.
 */
async function resolvePlace(
  suggestion: ISuggestion,
): Promise<{ place?: IPlace; state: SearchState }> {
  try {
    const place = suggestion.prediction.toPlace();
    await place.fetchFields({ fields: ['location', 'formattedAddress', 'displayName'] });

    const location = place.location;
    if (!location) return { state: 'error' };

    return {
      state: 'idle',
      place: {
        text: placeText(place, suggestion.primary),
        placeId: place.id,
        lat: location.lat(),
        lng: location.lng(),
      },
    };
  } catch (e) {
    console.error('[FleetPortal] Google Place details failed', e);
    return { state: classifyError(e) };
  }
}

/**
 * Make the text match the pin after a drag.
 *
 * Uses the JS Geocoder, NOT the Geocoding web service - a referrer-restricted
 * key is refused outright by the REST endpoint. If the Geocoding API is not
 * enabled on the key this returns undefined and the caller falls back to the
 * raw coordinates, so a dragged pin still books; it just reads as numbers.
 */
async function reverseGeocode(lat: number, lng: number): Promise<string | undefined> {
  try {
    const geocoder = new google.maps.Geocoder();
    const { results } = await geocoder.geocode({ location: { lat, lng } });
    const address = results[0]?.formatted_address;
    return address && address.length > MAX_TEXT ? address.slice(0, MAX_TEXT) : address;
  } catch (e) {
    console.warn(
      '[FleetPortal] Reverse geocode failed - the pin will read as coordinates. ' +
        'Check that the Geocoding API is enabled on the key.',
      e,
    );
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
  ready,
  failed,
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
  ready: boolean;
  /** The API will not arrive at all - as opposed to not having arrived yet. */
  failed: boolean;
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

  /**
   * One session token spans the whole typing run and the Place Details call
   * that ends it. Reusing a token across sessions, or omitting it, makes
   * Google bill every keystroke as its own request - which is the single
   * easiest way to run up a Places bill on a debounced field like this one.
   */
  const session = React.useRef<google.maps.places.AutocompleteSessionToken | undefined>(
    undefined,
  );

  const centreLat = centre.lat;
  const centreLng = centre.lng;

  // DEBOUNCED AT 300 ms, and nothing under 3 characters. Without both you fire
  // a request per keystroke and burn the quota on prefixes.
  React.useEffect(() => {
    const q = query.trim();
    if (q.length < 3) {
      setResults([]);
      setState('idle');
      return undefined;
    }
    if (failed) {
      setState('forbidden');
      return undefined;
    }
    /**
     * Still loading, not broken.
     *
     * Leaflet was bundled, so the search was usable the instant the step
     * rendered. Google arrives over the network, and someone who types
     * straight away would otherwise be told location search is unavailable
     * when it is merely half a second away. Hold the spinner instead - `ready`
     * is in the dependency list, so this effect re-runs and searches for real
     * the moment the API lands.
     */
    if (!ready) {
      setState('searching');
      return undefined;
    }

    const mine = ++seq.current;
    setState('searching');
    const handle = window.setTimeout(() => {
      if (!session.current) {
        session.current = new google.maps.places.AutocompleteSessionToken();
      }
      searchPlaces(q, { lat: centreLat, lng: centreLng }, radiusKm, session.current)
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
  }, [query, ready, failed, centreLat, centreLng, radiusKm]);

  const choose = (suggestion: ISuggestion): void => {
    setState('resolving');
    resolvePlace(suggestion)
      .then((r) => {
        // The session ends with the details call, win or lose.
        session.current = undefined;
        if (!r.place) {
          setState(r.state);
          return;
        }
        onPick(r.place);
        setState('idle');
        setQuery('');
        setResults([]);
        setOpen(false);
      })
      .catch(() => setState('error'));
  };

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

  const busy = state === 'searching' || state === 'resolving';

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
          {busy ? <span className={styles.placeSpinner} aria-hidden="true" /> : undefined}
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
              // Abandoning a choice ends that autocomplete session too.
              // Carrying the token into the next one makes Google bill every
              // keystroke separately instead of as one session.
              session.current = undefined;
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
            <li key={r.key}>
              <button
                type="button"
                className={styles.placeOption}
                disabled={state === 'resolving'}
                onClick={() => choose(r)}
              >
                <Icon name="pin" />
                <span>
                  {r.primary}
                  {r.secondary ? (
                    <span className={styles.placeOptionMeta}>{r.secondary}</span>
                  ) : undefined}
                </span>
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
    apiKey, mapId, isActive, centre, radiusKm,
    pickup, dropOff, pickupNotes, dropOffNotes,
    onPickup, onDropOff, onPickupNotes, onDropOffNotes,
  } = props;

  const containerRef = React.useRef<HTMLDivElement>(null);
  const mapRef = React.useRef<google.maps.Map | undefined>(undefined);
  const pickupMarker = React.useRef<IMapMarker | undefined>(undefined);
  const dropOffMarker = React.useRef<IMapMarker | undefined>(undefined);

  /**
   * Leaflet was bundled, so the map existed the moment the step rendered.
   * Google loads over the network, so there is now a real waiting state and
   * the step has to say so rather than showing an empty grey box.
   */
  const [ready, setReady] = React.useState(false);
  const [mapFailed, setMapFailed] = React.useState(false);

  // Keep the latest callbacks reachable from the marker handlers, which are
  // bound once and would otherwise close over the first render's props.
  const onPickupRef = React.useRef(onPickup);
  const onDropOffRef = React.useRef(onDropOff);
  onPickupRef.current = onPickup;
  onDropOffRef.current = onDropOff;

  // --- load the API once ---------------------------------------------------
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

  // --- create the map once the API is there, and tear it down on unmount ---
  React.useEffect(() => {
    if (!ready || !containerRef.current || mapRef.current) return undefined;

    mapRef.current = createMap({
      container: containerRef.current,
      centre,
      zoom: 11,
      mapId,
    });

    return () => {
      // Google has no map.remove(). Detach the markers and their listeners,
      // then let React discard the container - leaving listeners bound to a
      // removed node is how these leak across wizard steps.
      if (pickupMarker.current) pickupMarker.current.remove();
      if (dropOffMarker.current) dropOffMarker.current.remove();
      pickupMarker.current = undefined;
      dropOffMarker.current = undefined;
      if (mapRef.current) google.maps.event.clearInstanceListeners(mapRef.current);
      mapRef.current = undefined;
    };
    // Created once. The centre and key do not change within a booking.
  }, [ready]);

  // --- the map has no size until the step is actually visible --------------
  React.useEffect(() => {
    if (!isActive || !mapRef.current) return undefined;
    const map = mapRef.current;
    const handle = window.setTimeout(() => refreshMapSize(map), 0);
    return () => window.clearTimeout(handle);
  }, [isActive, ready]);

  // --- keep the pins in step with the chosen places ------------------------
  React.useEffect(() => {
    const map = mapRef.current;
    if (!map) return;

    const sync = (
      ref: React.MutableRefObject<IMapMarker | undefined>,
      place: IPlace | undefined,
      fill: string,
      label: string,
      commit: (p: IPlace) => void,
    ): void => {
      if (!place) {
        if (ref.current) {
          ref.current.remove();
          ref.current = undefined;
        }
        return;
      }
      if (!ref.current) {
        ref.current = addMarker({
          map,
          lat: place.lat,
          lng: place.lng,
          fill,
          label,
          draggable: true,
          onDragEnd: (lat, lng) => {
            // Reverse geocode so the text matches where the pin now is. There
            // is no place id on a reverse geocode, so PlaceId is left empty -
            // expected, and the column is optional.
            reverseGeocode(lat, lng)
              .then((text) =>
                commit({
                  text: text ?? `${lat.toFixed(5)}, ${lng.toFixed(5)}`,
                  lat,
                  lng,
                }),
              )
              .catch(() => undefined);
          },
        });
      } else {
        ref.current.setPosition(place.lat, place.lng);
      }
    };

    sync(pickupMarker, pickup, PICKUP_FILL, 'Pickup point', (p) => onPickupRef.current(p));
    sync(dropOffMarker, dropOff, DROPOFF_FILL, 'Destination', (p) => onDropOffRef.current(p));

    if (pickup && dropOff) {
      fitTwoPoints(
        map,
        { lat: pickup.lat, lng: pickup.lng },
        { lat: dropOff.lat, lng: dropOff.lng },
        40,
        15,
      );
    } else if (pickup) {
      map.setCenter({ lat: pickup.lat, lng: pickup.lng });
      map.setZoom(14);
    } else if (dropOff) {
      map.setCenter({ lat: dropOff.lat, lng: dropOff.lng });
      map.setZoom(14);
    }
  }, [pickup, dropOff, ready]);

  return (
    <div className={styles.mapPicker}>
      <div className={styles.mapPane}>
        <div
          ref={containerRef}
          className={styles.mapCanvas}
          role="application"
          aria-label="Pickup and destination map"
        />
        {mapFailed ? (
          // The map is an aid; the coordinates are the data. A map failure
          // must never block a booking that already has both pins.
          <div className={styles.mapFallback}>
            The map could not load, but your chosen places are still saved.
          </div>
        ) : !ready ? (
          <div className={styles.mapFallback}>Loading the map&hellip;</div>
        ) : undefined}
        <p className={styles.mapHint}>
          Search below, or drag a pin to adjust it. Green is the pickup, red is the
          destination.
        </p>
      </div>

      <div className={styles.placeFields}>
        <PlaceSearch
          id="pickup-search"
          label="Pickup point"
          placeholder="Search a place, building or road"
          ready={ready}
          failed={mapFailed}
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
          ready={ready}
          failed={mapFailed}
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
          No Google Maps key is configured, so the map and search are unavailable.
          Set it in the web part property pane.
        </Notice>
      ) : undefined}
    </div>
  );
}

export default MapPicker;
