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
  IProperty,
  IVehicle,
  IVehicleType,
  RequestType,
} from '../models';
import { toIso } from '../utils/dates';
import { IMasterDataService } from './IServices';
import {
  EXPAND_BLOCK,
  EXPAND_DRIVER,
  EXPAND_VEHICLE,
  SELECT_ALLOCATION_RULE,
  SELECT_PROPERTY,
  SELECT_BLOCK,
  SELECT_DRIVER,
  SELECT_VEHICLE,
  SELECT_VEHICLE_TYPE,
  toAllocationRule,
  toAvailabilityBlock,
  toDriver,
  toProperty,
  toVehicle,
  toVehicleType,
} from './mappers';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */


export class MasterDataService implements IMasterDataService {
  private vehicleTypes?: IVehicleType[];
  private properties?: IProperty[];
  private drivers = new Map<string, IDriver[]>();
  private vehicles = new Map<string, IVehicle[]>();
  private rules?: IAllocationRule[];
  private driverProperties?: Map<number, number[]>;

  public clearCache(): void {
    this.vehicleTypes = undefined;
    this.properties = undefined;
    this.drivers.clear();
    this.vehicles.clear();
    this.rules = undefined;
    this.driverProperties = undefined;
  }

  /**
   * Every property, for the Dispatch Dashboard's filter control.
   *
   * This is the ONLY remaining use of PropertyMaster in the solution: a label
   * the Coordinator can choose to narrow the view by. It does not scope
   * anything on its own.
   */
  public async getProperties(): Promise<IProperty[]> {
    if (this.properties) return this.properties;

    return guard(
      'The property list could not be loaded.',
      'Refresh the page. The dashboard still works without it; you just cannot filter by property.',
      async () => {
        const rows: any[] = await getSP()
          .web.lists.getByTitle(LIST.propertyMaster)
          .items.select(...SELECT_PROPERTY)
          .orderBy('Title', true)
          .top(200)();
        this.properties = rows.map(toProperty);
        return this.properties;
      },
    );
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

  public async getDrivers(): Promise<IDriver[]> {
    const key = 'all';
    const hit = this.drivers.get(key);
    if (hit) return hit;

    return guard(
      'The driver list could not be loaded.',
      'Refresh the page. If it keeps failing, contact a Fleet Admin.',
      async () => {
        // No property filter. There is one driver pool now - HomeProperty is
        // informational only.
        const filter = 'IsActive eq 1';

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

  public async getVehicles(): Promise<IVehicle[]> {
    const key = 'all';
    const hit = this.vehicles.get(key);
    if (hit) return hit;

    return guard(
      'The vehicle list could not be loaded.',
      'Refresh the page. If it keeps failing, contact a Fleet Admin.',
      async () => {
        // PropertyRef is informational now; it no longer narrows the pool.
        const filter = "VehStatus ne 'Retired'";

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
                .filter('IsActive eq 1')
            .top(200)();
          return rows.map(toAllocationRule);
        },
      );
    }

    // Keyed on request type alone now that property scoping is gone.
    return this.rules.filter((r) => r.requestType === requestType)[0];
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

}
