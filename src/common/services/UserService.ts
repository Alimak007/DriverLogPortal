/**
 * Who is signed in, what role they hold, and which properties they may see.
 *
 * ROLE HERE IS A USABILITY MEASURE, NOT THE SECURITY BOUNDARY. The boundary is
 * the SharePoint group permission table. This decides which web part sections
 * to render; it decides nothing about what the server will allow.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LIST } from '../config/AppConfig';
import { ActorRole } from '../rules/statusTransitions';
import { ICurrentUser, IUserService } from './IServices';
import { getSP, guard } from './SPContext';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** SharePoint group title -> role. Groups are the authority; this is the map. */
const GROUP_ROLE: { group: string; role: ActorRole }[] = [
  { group: 'FM Coordinators', role: 'coordinator' },
  { group: 'FM Admins', role: 'admin' },
  { group: 'FM Drivers', role: 'driver' },
  { group: 'FM Viewers', role: 'viewer' },
  { group: 'FM Requesters', role: 'requester' },
];

export class UserService implements IUserService {
  private readonly context: WebPartContext;
  private cached?: ICurrentUser;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getCurrentUser(): Promise<ICurrentUser> {
    if (this.cached) return this.cached;

    return guard(
      'Your account details could not be loaded.',
      'Refresh the page. If it happens again, contact the Transportation Coordinator.',
      async () => {
        const sp = getSP();
        const page = this.context.pageContext.user;

        const groups: any[] = await sp.web.currentUser.groups();
        const titles = groups.map((g) => String(g.Title));

        const roles: ActorRole[] = GROUP_ROLE.filter(
          (m) => titles.indexOf(m.group) !== -1,
        ).map((m) => m.role);

        // Everyone who can open the portal can at least raise a request.
        if (roles.length === 0) roles.push('requester');

        const siteUser = await sp.web.currentUser();

        let driverId: number | undefined;
        let propertyIds: number[] = [];

        if (roles.indexOf('driver') !== -1) {
          const driverRows: any[] = await sp.web.lists
            .getByTitle(LIST.driversMaster)
            .items.select('Id', 'HomeProperty/Id')
            .expand('HomeProperty')
            // UserAccountId is the hidden lookup-id column behind the
            // UserAccount person field. Filtering on it needs no $expand,
            // whereas "UserAccount/Id" does and throws without one.
            .filter(`UserAccountId eq ${siteUser.Id}`)
            .top(1)();

          if (driverRows.length > 0) {
            driverId = driverRows[0].Id;
            if (driverRows[0].HomeProperty?.Id) {
              propertyIds.push(driverRows[0].HomeProperty.Id);
            }
          }
        }

        // A driver may be allocated across more than one property, which
        // SharePoint groups cannot express - so scoping comes from
        // DLClientMaster, not from permissions alone.
        if (driverId !== undefined) {
          const links = await this.readDriverProperties(driverId);
          propertyIds = Array.from(new Set(propertyIds.concat(links)));
        }

        // Management Viewer sees every property. An empty array means "all".
        if (roles.indexOf('viewer') !== -1 && roles.length === 1) {
          propertyIds = [];
        }

        this.cached = {
          id: siteUser.Id,
          displayName: page.displayName,
          email: page.email,
          loginName: page.loginName,
          roles,
          driverId,
          propertyIds,
        };
        return this.cached;
      },
    );
  }

  private async readDriverProperties(driverId: number): Promise<number[]> {
    try {
      const rows: any[] = await getSP()
        .web.lists.getByTitle(LIST.dlClientMaster)
        .items.select('Id', 'PropertyRef/Id', 'DriverID/Id')
        .expand('PropertyRef', 'DriverID')
        .filter(`DriverIDId eq ${driverId}`)();
      return rows
        .map((r) => r.PropertyRef?.Id)
        .filter((id): id is number => typeof id === 'number');
    } catch {
      // If DriverID / PropertyRef are not the real internal names, a driver
      // falls back to HomeProperty rather than the page failing to render -
      // but they will not see a second property's requests. Confirm before
      // go-live.
      console.warn(
        '[FleetPortal] DLClientMaster could not be read with the assumed ' +
          'column names (Driver, Property). Confirm the internal names - see ' +
          'SharePoint-Lists-Specification.txt section 2. Falling back to ' +
          'HomeProperty for scoping.',
      );
      return [];
    }
  }
}

export function hasRole(user: ICurrentUser, ...roles: ActorRole[]): boolean {
  return roles.some((r) => user.roles.indexOf(r) !== -1);
}
