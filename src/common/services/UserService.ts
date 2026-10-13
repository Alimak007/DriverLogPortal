/**
 * Who is signed in, what role they hold, and which properties they may see.
 *
 * THERE IS NO propertyIds ANY MORE, deliberately. Every Coordinator sees every
 * request. With a single driver pool, an escalated request scoped to one
 * property would be invisible to everyone else - and if that Coordinator is on
 * leave, nobody picks it up. Scoping would put a blind spot in exactly the
 * path that exists to catch failures. The Dispatch Dashboard offers a property
 * filter the Coordinator chooses instead.
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

        if (roles.indexOf('driver') !== -1) {
          const driverRows: any[] = await sp.web.lists
            .getByTitle(LIST.driversMaster)
            .items.select('Id')
            // UserAccountId is the hidden lookup-id column behind the
            // UserAccount person field. Filtering on it needs no $expand,
            // whereas "UserAccount/Id" does and throws without one.
            .filter(`UserAccountId eq ${siteUser.Id}`)
            .top(1)();

          if (driverRows.length > 0) {
            driverId = driverRows[0].Id;
          }
        }

        const resolved: ICurrentUser = {
          id: siteUser.Id,
          displayName: page.displayName,
          email: page.email,
          loginName: page.loginName,
          roles,
          driverId,
        };
        this.cached = resolved;
        return resolved;
      },
    );
  }

}

export function hasRole(user: ICurrentUser, ...roles: ActorRole[]): boolean {
  return roles.some((r) => user.roles.indexOf(r) !== -1);
}
