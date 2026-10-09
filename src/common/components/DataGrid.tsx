/**
 * The grid every view uses.
 *
 * It is a REAL <table> with a <caption> and scoped headers, scrolling inside
 * its own focusable, labelled region. The header row sticks to the top while
 * the body scrolls.
 *
 * A filterable header becomes a button showing a caret, and a FUNNEL MARK AND
 * A COUNT when a filter is set - not just a tint. The filter list is a popup
 * of the values that column actually holds, NEVER a free-text box.
 */

import * as React from 'react';
import styles from './Shared.module.scss';
import { Button } from './Shared';

export interface IColumn<T> {
  key: string;
  label: string;
  /** The cell. Return a node; the grid does no formatting of its own. */
  render: (row: T) => React.ReactNode;
  /** The value a filter matches on, and what a filter list shows. */
  filterValue?: (row: T) => string;
  /** Right-aligned, tabular figures. */
  numeric?: boolean;
  width?: string;
}

export interface IDataGridProps<T> {
  /** Names the grid for a screen reader. Required - every grid has one. */
  caption: string;
  /** The caption is read but not shown where a heading already says it. */
  captionVisuallyHidden?: boolean;
  columns: IColumn<T>[];
  rows: T[];
  rowKey: (row: T) => string | number;
  /** Rows pinned to the top carry an amber edge AND their status in words. */
  isPinned?: (row: T) => boolean;
  emptyMessage?: string;
}

export function DataGrid<T>(props: IDataGridProps<T>): React.ReactElement {
  const {
    caption,
    captionVisuallyHidden,
    columns,
    rows,
    rowKey,
    isPinned,
    emptyMessage,
  } = props;

  const [openFilter, setOpenFilter] = React.useState<string | undefined>(undefined);
  const [selected, setSelected] = React.useState<Record<string, string[]>>({});

  const regionRef = React.useRef<HTMLDivElement>(null);

  // Close the popup on Escape, and on a click outside it.
  React.useEffect(() => {
    if (!openFilter) return undefined;

    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setOpenFilter(undefined);
    };
    const onClick = (e: MouseEvent): void => {
      if (regionRef.current && !regionRef.current.contains(e.target as Node)) {
        setOpenFilter(undefined);
      }
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('mousedown', onClick);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('mousedown', onClick);
    };
  }, [openFilter]);

  const filtered = React.useMemo(() => {
    const active = Object.keys(selected).filter((k) => selected[k].length > 0);
    if (active.length === 0) return rows;

    return rows.filter((row) =>
      active.every((key) => {
        const column = columns.filter((c) => c.key === key)[0];
        if (!column || !column.filterValue) return true;
        return selected[key].indexOf(column.filterValue(row)) !== -1;
      }),
    );
  }, [rows, selected, columns]);

  const ordered = React.useMemo(() => {
    if (!isPinned) return filtered;
    const pinned = filtered.filter((r) => isPinned(r));
    const rest = filtered.filter((r) => !isPinned(r));
    return pinned.concat(rest);
  }, [filtered, isPinned]);

  const distinctValues = (column: IColumn<T>): string[] => {
    if (!column.filterValue) return [];
    const set = new Set<string>();
    rows.forEach((r) => {
      const v = column.filterValue as (row: T) => string;
      set.add(v(r));
    });
    return Array.from(set).sort();
  };

  const toggle = (columnKey: string, value: string): void => {
    setSelected((prev) => {
      const current = prev[columnKey] ?? [];
      const next =
        current.indexOf(value) === -1
          ? current.concat([value])
          : current.filter((v) => v !== value);
      return { ...prev, [columnKey]: next };
    });
  };

  const clear = (columnKey: string): void => {
    setSelected((prev) => ({ ...prev, [columnKey]: [] }));
    setOpenFilter(undefined);
  };

  const captionClass = captionVisuallyHidden
    ? styles.srOnly
    : styles['caption-visible'];

  return (
    <div
      className={styles.gridRegion}
      ref={regionRef}
      tabIndex={0}
      role="region"
      aria-label={caption}
    >
      <table className={styles.table}>
        <caption className={captionClass}>
          {caption}
          {ordered.length !== rows.length
            ? ` — showing ${ordered.length} of ${rows.length} rows`
            : ''}
        </caption>
        <thead>
          <tr>
            {columns.map((column) => {
              const activeCount = (selected[column.key] ?? []).length;
              const filterable = column.filterValue !== undefined;

              return (
                <th
                  key={column.key}
                  scope="col"
                  style={column.width ? { width: column.width } : undefined}
                  className={column.numeric ? styles.numeric : undefined}
                >
                  {filterable ? (
                    <button
                      type="button"
                      className={styles.filterButton}
                      onClick={() =>
                        setOpenFilter(openFilter === column.key ? undefined : column.key)
                      }
                      aria-expanded={openFilter === column.key}
                      aria-label={
                        activeCount > 0
                          ? `${column.label}, filtered to ${activeCount} values. Change the filter.`
                          : `${column.label}. Filter this column.`
                      }
                    >
                      {column.label}
                      {activeCount > 0 ? (
                        // A funnel mark AND a count, never just a tint.
                        <span className={styles.filterMark}>
                          {'▼'} filtered ({activeCount})
                        </span>
                      ) : (
                        <span className={styles.filterMark} aria-hidden="true">
                          {'▾'}
                        </span>
                      )}
                    </button>
                  ) : (
                    column.label
                  )}

                  {openFilter === column.key ? (
                    <div className={styles.filterPopup} role="group" aria-label={`Filter ${column.label}`}>
                      {distinctValues(column).map((value) => (
                        <label className={styles.filterOption} key={value}>
                          <input
                            type="checkbox"
                            checked={(selected[column.key] ?? []).indexOf(value) !== -1}
                            onChange={() => toggle(column.key, value)}
                          />
                          {value === '' ? '(blank)' : value}
                        </label>
                      ))}
                      <div className={styles.filterActions}>
                        <Button kind="action" onClick={() => clear(column.key)}>
                          Clear
                        </Button>
                        <Button kind="action" onClick={() => setOpenFilter(undefined)}>
                          Close
                        </Button>
                      </div>
                    </div>
                  ) : undefined}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {ordered.length === 0 ? (
            <tr>
              <td colSpan={columns.length}>
                {emptyMessage ?? 'Nothing to show.'}
              </td>
            </tr>
          ) : (
            ordered.map((row) => (
              <tr
                key={rowKey(row)}
                className={isPinned && isPinned(row) ? styles.rowPinned : undefined}
              >
                {columns.map((column) => (
                  <td
                    key={column.key}
                    className={column.numeric ? styles.numeric : undefined}
                  >
                    {column.render(row)}
                  </td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

export const RowActions: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className={styles.rowActions}>{children}</div>
);
