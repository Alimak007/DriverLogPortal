/**
 * Reference data, cached for the life of the page.
 *
 * None of it changes during a session, and re-reading it on every screen
 * would blow the four-batched-calls budget the non-functional requirements
 * set for a dashboard load.
 */

import { APP_CONFIG, LIST } from '../config/AppConfig';
import {
  IAllocationRule,
  IAvailabilityBlock,
  IDriver,
  ILocation,
  IVehicle,
  IVehicleType,
  RequestType,
} from '../models';
import { toIso } from '../utils/dates';
import { IMasterDataService } from './IServices';
import {
  EXPAND_ALLOCATION_RULE,
  EXPAND_BLOCK,
  EXPAND_DRIVER,
  EXPAND_LOCATION,
  EXPAND_VEHICLE,
  SELECT_ALLOCATION_RULE,
  SELECT_BLOCK,
  SELECT_DRIVER,
  SELECT_LOCATION,
  SELECT_VEHICLE,
  SELECT_VEHICLE_TYPE,
  toAllocationRule,
  toAvailabilityBlock,
  toDriver,
  toLocation,
  toVehicle,
  toVehicleType,
} from './mappers';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** `Property/Id eq 1 or Property/Id eq 2`, or '' for no scoping. */
function scopeFilter(field: string, propertyIds?: number[]): string {
  if (!propertyIds || propertyIds.length === 0) return '';
  return propertyIds.map((id) => `${field} eq ${id}`).join(' or ');
}

function and(...clauses: string[]): string {
  const parts = clauses.filter((c) => c.length > 0).map((c) => `(${c})`);
  return parts.join(' and ');
}

export class MasterDataService implements IMasterDataService {
  private vehicleTypes?: IVehicleType[];
  private locations = new Map<string, ILocation[]>();
  private drivers = new Map<string, IDriver[]>();
  private vehicles = new Map<string, IVehicle[]>();
  private rules?: IAllocationRule[];
  private driverProperties?: Map<number, number[]>;

  public clearCache(): void {
    this.vehicleTypes = undefined;
    this.locations.clear();
    this.drivers.clear();
    this.vehicles.clear();
    this.rules = undefined;
    this.driverProperties = undefined;
  }

  public async getVehicleTypes(): Promise<IVehicleType[]> {
    if (this.vehicleTypes) return this.vehicleTypes;

    return guard(
      'The list of vehicle types could not be loaded.',
      'Refresh the page. If the list is empty, ask a Fleet Admin to check VehicleTypes.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.vehicleTypes)
          .items.select(...SELECT_VEHICLE_TYPE)
          .filter('IsActive eq 1')
          .orderBy('SortOrder', true)
          .top(APP_CONFIG.pageSize)();

        this.vehicleTypes = rows.map(toVehicleType);
        return this.vehicleTypes;
      },
    );
  }

  /**
   * The vehicle tiles the requester sees, filtered by SuitableFor against the
   * chosen request type. "Others" is included for every type - it is the
   * single free-text fallback.
   */
  public async getVehicleTypesFor(requestType: RequestType): Promise<IVehicleType[]> {
    const all = await this.getVehicleTypes();
    return all.filter((t) => t.suitableFor.indexOf(requestType) !== -1);
  }

  public async getLocations(propertyId?: number): Promise<ILocation[]> {
    const key = String(propertyId ?? 'all');
    const hit = this.locations.get(key);
    if (hit) return hit;

    return guard(
      'The list of destinations could not be loaded.',
      'Refresh the page. If no destinations appear, ask a Fleet Admin to check Locations.',
      async () => {
        const filter = and(
          'IsActive eq 1',
          propertyId !== undefined ? `Property/Id eq ${propertyId}` : '',
        );

        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.locations)
          .items.select(...SELECT_LOCATION)
          .expand(...EXPAND_LOCATION)
          .filter(filter)
          .orderBy('Title', true)
          .top(500)();

        const mapped = rows.map(toLocation);
        this.locations.set(key, mapped);
        return mapped;
      },
    );
  }

  public async getDrivers(propertyIds?: number[]): Promise<IDriver[]> {
    const key = (propertyIds ?? []).join(',') || 'all';
    const hit = this.drivers.get(key);
    if (hit) return hit;

    return guard(
      'The driver list could not be loaded.',
      'Refresh the page. If it keeps failing, contact a Fleet Admin.',
      async () => {
        const filter = and(
          'IsActive eq 1',
          scopeFilter('HomeProperty/Id', propertyIds),
        );

        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.driversMaster)
          .items.select(...SELECT_DRIVER)
          .expand(...EXPAND_DRIVER)
          .filter(filter)
          .orderBy('Title', true)
          .top(500)();

        const mapped = rows.map(toDriver);
        this.drivers.set(key, mapped);
        return mapped;
      },
    );
  }

  public async getVehicles(propertyIds?: number[]): Promise<IVehicle[]> {
    const key = (propertyIds ?? []).join(',') || 'all';
    const hit = this.vehicles.get(key);
    if (hit) return hit;

    return guard(
      'The vehicle list could not be loaded.',
      'Refresh the page. If it keeps failing, contact a Fleet Admin.',
      async () => {
        const filter = and(
          "VehStatus ne 'Retired'",
          scopeFilter('PropertyRef/Id', propertyIds),
        );

        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.vehicleMaster)
          .items.select(...SELECT_VEHICLE)
          .expand(...EXPAND_VEHICLE)
          .filter(filter)
          .orderBy('PlateNo', true)
          .top(500)();

        const mapped = rows.map(toVehicle);
        this.vehicles.set(key, mapped);
        return mapped;
      },
    );
  }

  public async getAllocationRule(
    propertyId: number,
    requestType: RequestType,
  ): Promise<IAllocationRule | undefined> {
    if (!this.rules) {
      this.rules = await guard(
        'The allocation rules could not be loaded.',
        'Refresh the page. Until they load, the default five-minute window applies.',
        async () => {
          const rows: any[] = await getSP()
            .web.lists.getByTitle(LIST.allocationRules)
            .items.select(...SELECT_ALLOCATION_RULE)
            .expand(...EXPAND_ALLOCATION_RULE)
            .filter('IsActive eq 1')
            .top(200)();
          return rows.map(toAllocationRule);
        },
      );
    }

    return this.rules.filter(
      (r) => r.property?.id === propertyId && r.requestType === requestType,
    )[0];
  }

  public async getBlocks(from: Date, to: Date): Promise<IAvailabilityBlock[]> {
    return guard(
      'Leave and maintenance records could not be loaded.',
      'Refresh the page. Availability shown may be out of date until it loads.',
      async () => {
        // Any block that overlaps the window at all.
        const filter = `FromDateTime lt '${toIso(to)}' and ToDateTime gt '${toIso(from)}'`;

        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.availabilityBlocks)
          .items.select(...SELECT_BLOCK)
          .expand(...EXPAND_BLOCK)
          .filter(filter)
          .top(500)();

        return rows.map(toAvailabilityBlock);
      },
    );
  }

  /**
   * Driver id -> property ids, from DLClientMaster.
   *
   * A driver may be allocated across more than one property, which is exactly
   * why scoping is derived from this list rather than from SharePoint groups.
   */
  public async getDriverPropertyMap(): Promise<Map<number, number[]>> {
    if (this.driverProperties) return this.driverProperties;

    const map = new Map<number, number[]>();
    try {
      const rows: any[] = await getSP()
        .web.lists.getByTitle(LIST.dlClientMaster)
        .items.select('Id', 'DriverID/Id', 'PropertyRef/Id')
        .expand('DriverID', 'PropertyRef')
        .top(2000)();

      rows.forEach((r) => {
        const d = r.DriverID?.Id;
        const p = r.PropertyRef?.Id;
        if (typeof d !== 'number' || typeof p !== 'number') return;
        const existing = map.get(d) ?? [];
        if (existing.indexOf(p) === -1) existing.push(p);
        map.set(d, existing);
      });
    } catch {
      console.warn(
        '[FleetPortal] DLClientMaster could not be read with the assumed ' +
          'column names (Driver, Property). Confirm the internal names before ' +
          'go-live - see SharePoint-Lists-Specification.txt section 2.',
      );
    }

    this.driverProperties = map;
    return map;
  }
}
