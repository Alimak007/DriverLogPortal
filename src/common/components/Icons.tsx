/**
 * The icon set.
 *
 * Drawn inline as SVG so the bundle carries no icon font and no image files.
 * Every icon is DECORATIVE: it is hidden from assistive technology and always
 * sits beside the words that carry the meaning.
 */

import * as React from 'react';

export type IconName =
  | 'arrowRight'
  | 'bell'
  | 'box'
  | 'building'
  | 'calendar'
  | 'car'
  | 'chart'
  | 'check'
  | 'chevronLeft'
  | 'chevronRight'
  | 'clock'
  | 'doc'
  | 'download'
  | 'flag'
  | 'gear'
  | 'hash'
  | 'home'
  | 'lock'
  | 'navigation'
  | 'phone'
  | 'pin'
  | 'plane'
  | 'plus'
  | 'refresh'
  | 'search'
  | 'target'
  | 'user'
  | 'users'
  | 'van'
  | 'volume'
  | 'volumeOff'
  | 'x';

const PATHS: Record<IconName, React.ReactNode> = {
  arrowRight: <path d="M5 12h14M13 6l6 6-6 6" />,
  bell: <path d="M6 16v-5a6 6 0 0 1 12 0v5l2 2H4zM10 20a2 2 0 0 0 4 0" />,
  box: <path d="M12 3l8 4.5v9L12 21l-8-4.5v-9zM4 7.5l8 4.5 8-4.5M12 12v9" />,
  building: (
    <>
      <path d="M4 21V5a1 1 0 0 1 1-1h8a1 1 0 0 1 1 1v16M14 10h5a1 1 0 0 1 1 1v10M2 21h20" />
      <path d="M8 8h2M8 12h2M8 16h2" />
    </>
  ),
  calendar: (
    <>
      <rect x="3" y="5" width="18" height="16" rx="2" />
      <path d="M3 10h18M8 3v4M16 3v4" />
    </>
  ),
  car: (
    <>
      <path d="M5 11l1.6-4.2A2 2 0 0 1 8.5 5.5h7a2 2 0 0 1 1.9 1.3L19 11" />
      <rect x="3" y="11" width="18" height="6" rx="2" />
      <path d="M6.5 17v2M17.5 17v2M7 14h.01M17 14h.01" />
    </>
  ),
  chart: <path d="M5 20V11M11 20V5M17 20v-7M3 20h18" />,
  check: <path d="M5 12.5l4.5 4.5L19 7.5" />,
  chevronLeft: <path d="M15 6l-6 6 6 6" />,
  chevronRight: <path d="M9 6l6 6-6 6" />,
  clock: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 7v5l3 2" />
    </>
  ),
  doc: (
    <>
      <path d="M14 3H7a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7z" />
      <path d="M14 3v4h4M9 13h6M9 17h6" />
    </>
  ),
  download: <path d="M12 4v11M7 10l5 5 5-5M5 20h14" />,
  flag: <path d="M5 21V4h11l-2 4 2 4H5" />,
  gear: (
    <>
      <circle cx="12" cy="12" r="4" />
      <path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9L7 7M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1" />
    </>
  ),
  hash: <path d="M4 9h16M4 15h16M10 3L8 21M16 3l-2 18" />,
  home: <path d="M3 11l9-7 9 7M5 10v10h14V10M10 20v-6h4v6" />,
  lock: (
    <>
      <rect x="5" y="11" width="14" height="10" rx="2" />
      <path d="M8 11V8a4 4 0 0 1 8 0v3" />
    </>
  ),
  navigation: <path d="M21 3L3 10.5l7.5 3 3 7.5z" />,
  phone: (
    <path d="M6.5 3h3l1.5 4.5-2 1.5a12 12 0 0 0 6 6l1.5-2 4.5 1.5v3a2 2 0 0 1-2.2 2A17 17 0 0 1 4.5 5.2 2 2 0 0 1 6.5 3z" />
  ),
  pin: (
    <>
      <path d="M12 21s7-6.2 7-11.5A7 7 0 0 0 5 9.5C5 14.8 12 21 12 21z" />
      <circle cx="12" cy="9.5" r="2.5" />
    </>
  ),
  plane: <path d="M2 16l20-6-2-3-6 1-5-5H7l3 6-5 1-2-2H1z" />,
  plus: <path d="M12 5v14M5 12h14" />,
  refresh: (
    <path d="M20 11a8 8 0 0 0-14.6-4.5L4 8M4 4v4h4M4 13a8 8 0 0 0 14.6 4.5L20 16M20 20v-4h-4" />
  ),
  search: (
    <>
      <circle cx="11" cy="11" r="7" />
      <path d="M21 21l-4.3-4.3" />
    </>
  ),
  target: (
    <>
      <circle cx="12" cy="12" r="7" />
      <circle cx="12" cy="12" r="2.5" />
      <path d="M12 2v3M12 19v3M2 12h3M19 12h3" />
    </>
  ),
  user: (
    <>
      <circle cx="12" cy="8" r="4" />
      <path d="M4 21c0-4 3.6-6 8-6s8 2 8 6" />
    </>
  ),
  users: (
    <>
      <circle cx="9" cy="8" r="3.5" />
      <path d="M2.5 20c0-3.5 3-5.5 6.5-5.5s6.5 2 6.5 5.5" />
      <circle cx="17" cy="9" r="2.5" />
      <path d="M17.5 14.6c2.4.3 4 1.9 4 4.4" />
    </>
  ),
  van: (
    <>
      <path d="M2 7h11v10H2zM13 10h4.5l3.5 3.5V17h-8z" />
      <circle cx="6.5" cy="17.5" r="2" />
      <circle cx="16.5" cy="17.5" r="2" />
    </>
  ),
  volume: <path d="M4 9h4l5-4v14l-5-4H4zM16.5 9a4 4 0 0 1 0 6M19 6.5a8 8 0 0 1 0 11" />,
  volumeOff: <path d="M4 9h4l5-4v14l-5-4H4zM17 9.5l5 5M22 9.5l-5 5" />,
  x: <path d="M6 6l12 12M18 6L6 18" />,
};

export const Icon: React.FC<{
  name: IconName;
  size?: number;
  className?: string;
}> = ({ name, size = 20, className }) => (
  <svg
    className={className}
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={2}
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
    focusable="false"
  >
    {PATHS[name]}
  </svg>
);
