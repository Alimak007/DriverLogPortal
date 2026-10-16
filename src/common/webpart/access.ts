/**
 * Which SharePoint group may open which web part.
 *
 * THIS IS NAVIGATION, NOT SECURITY. Read that again before relying on it.
 *
 * A web part that declines to render still ran in the browser, and the person
 * looking at it still holds whatever permission SharePoint granted them on the
 * lists behind it. Drivers have Edit on this site, so a driver who opens the
 * REST API directly can read and write TransportRequests whatever this file
 * says. The real boundary is the SharePoint permission table on the lists and
 * the site - this only stops people landing on a page that is not theirs and
 * wondering what it is for.
 *
 * It is worth having anyway: a Requester who opens the Dispatch Dashboard sees
 * a sentence telling them where to go, rather than a Coordinator's console
 * full of other people's trips.
 *
 * Changing who sees what is a change to this one table.
 */

import { ActorRole } from '../rules/statusTransitions';

export type FleetWebPartKey =
  | 'bookARide'
  | 'myBookings'
  | 'driverDashboard'
  | 'dispatchDashboard'
  | 'reports';

export interface IWebPartAccess {
  /** Shown in the notice, so the reader knows which page they are on. */
  title: string;
  /** Any one of these is enough. */
  roles: ActorRole[];
  /** The SharePoint groups behind those roles, named so the notice can say. */
  groups: string[];
}

/**
 * FM Admins hold every role deliberately - an administrator who cannot open a
 * page cannot diagnose a report about it.
 */
export const WEB_PART_ACCESS: Record<FleetWebPartKey, IWebPartAccess> = {
  bookARide: {
    title: 'Book a Ride',
    roles: ['requester', 'admin'],
    groups: ['FM Requesters', 'FM Admins'],
  },
  myBookings: {
    title: 'My Bookings',
    roles: ['requester', 'admin'],
    groups: ['FM Requesters', 'FM Admins'],
  },
  driverDashboard: {
    title: 'Driver Dashboard',
    roles: ['driver', 'admin'],
    groups: ['FM Drivers', 'FM Admins'],
  },
  dispatchDashboard: {
    title: 'Dispatch Dashboard',
    roles: ['coordinator', 'admin'],
    groups: ['FM Coordinators', 'FM Admins'],
  },
  reports: {
    title: 'Reports',
    roles: ['coordinator', 'admin'],
    groups: ['FM Coordinators', 'FM Admins'],
  },
};

/**
 * May this user open this web part?
 *
 * `roles` comes from UserService, which reads the signed-in user's SharePoint
 * group memberships. A user in several groups holds several roles and needs
 * only one of them to match.
 */
export function mayOpen(roles: ActorRole[] | undefined, key: FleetWebPartKey): boolean {
  if (!roles || roles.length === 0) return false;
  const allowed = WEB_PART_ACCESS[key].roles;
  return roles.some((r) => allowed.indexOf(r) !== -1);
}
