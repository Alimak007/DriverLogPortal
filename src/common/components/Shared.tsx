/**
 * The shared component vocabulary.
 *
 * Every button in the solution comes from here, so focus and disabled
 * behaviour cannot drift apart. Every badge comes from here, so no badge can
 * be rendered without its label.
 */

import * as React from 'react';
import { SemanticState } from '../models';
import { PortalError } from '../services/SPContext';
import styles from './Shared.module.scss';

// -----------------------------------------------------------------------------
// Page shell
// -----------------------------------------------------------------------------

export interface IPageShellProps {
  heading: string;
  caption?: string;
  actions?: React.ReactNode;
  children?: React.ReactNode;
}

export const PageShell: React.FC<IPageShellProps> = ({
  heading,
  caption,
  actions,
  children,
}) => (
  <div className={styles.shell}>
    <div className={styles.shellHeader}>
      <div>
        <h1 className={styles.pageHeading}>{heading}</h1>
        {caption ? <div className={styles.caption}>{caption}</div> : undefined}
      </div>
      {actions ? <div className={styles.headerActions}>{actions}</div> : undefined}
    </div>
    {children}
  </div>
);

export const SectionHeading: React.FC<{ children: React.ReactNode; id?: string }> = ({
  children,
  id,
}) => (
  <h2 className={styles.sectionHeading} id={id}>
    {children}
  </h2>
);

// -----------------------------------------------------------------------------
// Buttons
// -----------------------------------------------------------------------------

export type ButtonKind = 'primary' | 'secondary' | 'destructive' | 'action';

export interface IButtonProps
  extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  kind?: ButtonKind;
  block?: boolean;
}

const KIND_CLASS: Record<ButtonKind, string> = {
  primary: styles.btnPrimary,
  secondary: styles.btnSecondary,
  destructive: styles.btnDestructive,
  action: styles.btnAction,
};

export const Button: React.FC<IButtonProps> = ({
  kind = 'secondary',
  block,
  className,
  children,
  ...rest
}) => (
  <button
    type="button"
    className={[KIND_CLASS[kind], block ? styles.btnBlock : '', className ?? '']
      .filter(Boolean)
      .join(' ')}
    {...rest}
  >
    {children}
  </button>
);

/**
 * A link that looks like a button. For going to ANOTHER PAGE - it is a real
 * anchor, so it can be opened in a new tab and is announced as a link.
 */
export const LinkButton: React.FC<{
  href: string;
  kind?: ButtonKind;
  className?: string;
  children: React.ReactNode;
}> = ({ href, kind = 'secondary', className, children }) => (
  <a
    href={href}
    className={[KIND_CLASS[kind], styles.btnLink, className ?? ''].filter(Boolean).join(' ')}
  >
    {children}
  </a>
);

// -----------------------------------------------------------------------------
// Badges
// -----------------------------------------------------------------------------

const STATE_CLASS: Record<SemanticState, string> = {
  healthy: styles.badgeHealthy,
  attention: styles.badgeAttention,
  progress: styles.badgeProgress,
  exception: styles.badgeException,
  neutral: styles.badgeNeutral,
};

/**
 * A status badge. The label is required, not optional - colour never carries
 * meaning on its own anywhere in this interface.
 */
export const Badge: React.FC<{ state: SemanticState; label: string }> = ({
  state,
  label,
}) => <span className={STATE_CLASS[state]}>{label}</span>;

// -----------------------------------------------------------------------------
// Panels and cards
// -----------------------------------------------------------------------------

export const Panel: React.FC<{
  children: React.ReactNode;
  plain?: boolean;
  className?: string;
}> = ({ children, plain, className }) => (
  <div className={[plain ? styles.panelPlain : styles.panel, className ?? ''].join(' ')}>
    {children}
  </div>
);

export const Card: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className={styles.card}>{children}</div>
);

export const CardTitleRow: React.FC<{
  title: React.ReactNode;
  children?: React.ReactNode;
}> = ({ title, children }) => (
  <div className={styles.cardTitleRow}>
    <span className={styles.cardTitle}>{title}</span>
    {children}
  </div>
);

/** A reference number, set in the heading face as the brand rule requires. */
export const Reference: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <span className={styles.reference}>{children}</span>
);

export const FieldRow: React.FC<{ label: string; children: React.ReactNode }> = ({
  label,
  children,
}) => (
  <div className={styles.fieldRow}>
    <span className={styles.fieldLabel}>{label}</span>
    <span className={styles.fieldValue}>{children}</span>
  </div>
);

// -----------------------------------------------------------------------------
// Counters
// -----------------------------------------------------------------------------

/**
 * Counters are FILTER SHORTCUTS, not static numbers. Each navigates to the
 * matching filtered view, so each is a button with an accessible name that
 * says so.
 */
export const Counter: React.FC<{
  label: string;
  value: number;
  meta?: string;
  active?: boolean;
  onClick?: () => void;
}> = ({ label, value, meta, active, onClick }) => (
  <button
    type="button"
    className={[styles.counter, active ? styles.counterActive : ''].join(' ')}
    onClick={onClick}
    aria-pressed={active === true}
    aria-label={`${label}: ${value}. Show these.`}
  >
    <span className={styles.counterValue}>{value}</span>
    <span className={styles.counterLabel}>{label}</span>
    {meta ? <span className={styles.counterMeta}>{meta}</span> : undefined}
  </button>
);

// -----------------------------------------------------------------------------
// Form parts
// -----------------------------------------------------------------------------

export const FormRow: React.FC<{
  label: string;
  htmlFor?: string;
  required?: boolean;
  hint?: string;
  error?: string;
  children: React.ReactNode;
}> = ({ label, htmlFor, required, hint, error, children }) => (
  <div className={styles.formRow}>
    <label
      className={[styles.label, required ? styles.required : ''].join(' ')}
      htmlFor={htmlFor}
    >
      {label}
    </label>
    {children}
    {hint ? <span className={styles.hint}>{hint}</span> : undefined}
    {error ? (
      <span className={styles.validation} role="alert">
        {error}
      </span>
    ) : undefined}
  </div>
);

export const TextField: React.FC<
  React.InputHTMLAttributes<HTMLInputElement>
> = (props) => <input className={styles.field} {...props} />;

export const TextArea: React.FC<
  React.TextareaHTMLAttributes<HTMLTextAreaElement>
> = (props) => <textarea className={styles.field} rows={3} {...props} />;

export const Select: React.FC<
  React.SelectHTMLAttributes<HTMLSelectElement>
> = ({ children, ...rest }) => (
  <select className={styles.field} {...rest}>
    {children}
  </select>
);

/** The sticky command bar: Submit and Cancel stay in reach on a long form. */
export const CommandBar: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className={styles.commandBar}>{children}</div>
);

// -----------------------------------------------------------------------------
// Selectable tiles
// -----------------------------------------------------------------------------

/**
 * Selection is carried by a word ("Selected") as well as by the colour change
 * and by aria-pressed, so nothing here is distinguished by colour alone.
 */
export const Tile: React.FC<{
  title: string;
  caption?: string;
  selected: boolean;
  disabled?: boolean;
  onClick: () => void;
}> = ({ title, caption, selected, disabled, onClick }) => (
  <button
    type="button"
    className={[styles.tile, selected ? styles.tileSelected : ''].join(' ')}
    onClick={onClick}
    disabled={disabled}
    aria-pressed={selected}
  >
    <span className={styles.tileTitle}>{title}</span>
    {caption ? <span className={styles.tileCaption}>{caption}</span> : undefined}
    {selected ? <span className={styles.tileMark}>Selected</span> : undefined}
  </button>
);

export const TileGrid: React.FC<{ children: React.ReactNode; label: string }> = ({
  children,
  label,
}) => (
  <div className={styles.tileGrid} role="group" aria-label={label}>
    {children}
  </div>
);

// -----------------------------------------------------------------------------
// Messages
// -----------------------------------------------------------------------------

/**
 * Every failed write surfaces a message naming WHAT FAILED and WHAT TO DO.
 * The correlation reference is shown so a user can quote it to support.
 */
export const ErrorBanner: React.FC<{ error: unknown; onDismiss?: () => void }> = ({
  error,
  onDismiss,
}) => {
  if (!error) return <></>;

  const portal = error instanceof PortalError ? error : undefined;
  const whatFailed = portal
    ? portal.whatFailed
    : 'Something went wrong.';
  const whatToDo = portal
    ? portal.whatToDo
    : 'Refresh the page and try again. If it keeps failing, contact the ' +
      'Transportation Coordinator.';

  return (
    <div className={styles.errorBanner} role="alert">
      <div className={styles.errorTitle}>{whatFailed}</div>
      <div className={styles.errorAction}>{whatToDo}</div>
      {portal ? (
        <div className={styles.errorRef}>Reference: {portal.correlationId}</div>
      ) : undefined}
      {onDismiss ? (
        <div style={{ marginTop: 8 }}>
          <Button kind="action" onClick={onDismiss}>
            Dismiss
          </Button>
        </div>
      ) : undefined}
    </div>
  );
};

export const Notice: React.FC<{
  tone?: 'progress' | 'attention' | 'healthy';
  children: React.ReactNode;
}> = ({ tone = 'progress', children }) => {
  const toneClass =
    tone === 'attention'
      ? styles.noticeAttention
      : tone === 'healthy'
        ? styles.noticeHealthy
        : '';
  return (
    <div className={[styles.notice, toneClass].join(' ')} role="status">
      {children}
    </div>
  );
};

export const EmptyState: React.FC<{ title: string; children?: React.ReactNode }> = ({
  title,
  children,
}) => (
  <div className={styles.emptyState}>
    <div className={styles.emptyTitle}>{title}</div>
    {children}
  </div>
);

export const Loading: React.FC<{ what?: string }> = ({ what = 'Loading' }) => (
  <div className={styles.loading} role="status" aria-live="polite">
    {what}&hellip;
  </div>
);

export const VisuallyHidden: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => <span className={styles.srOnly}>{children}</span>;
