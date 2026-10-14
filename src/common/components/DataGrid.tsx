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
import { Icon } from './Icons';

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
  /** Rows per page. Defaults to 10; 0 shows every row on one page. */
  pageSize?: number;
}

/** The page numbers to offer: the first, the last, and two either side of the current one. */
function pageList(current: number, count: number): (number | 'gap')[] {
  const out: (number | 'gap')[] = [];
  for (let p = 1; p <= count; p++) {
    if (p === 1 || p === count || Math.abs(p - current) <= 1) {
      out.push(p);
    } else if (out[out.length - 1] !== 'gap') {
      out.push('gap');
    }
  }
  return out;
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
    pageSize = 10,
  } = props;

  const [openFilter, setOpenFilter] = React.useState<string | undefined>(undefined);
  const [selected, setSelected] = React.useState<Record<string, string[]>>({});
  const [page, setPage] = React.useState(1);

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

  // Paging. A new filter starts again at page 1. A refresh that shortens the
  // list keeps the reader where they were, or on the last page that remains.
  const paged = pageSize > 0;
  const pageCount = paged ? Math.max(1, Math.ceil(ordered.length / pageSize)) : 1;
  const current = Math.min(page, pageCount);

  React.useEffect(() => {
    setPage(1);
  }, [selected]);

  const visible = paged
    ? ordered.slice((current - 1) * pageSize, current * pageSize)
    : ordered;
  const firstShown = ordered.length === 0 ? 0 : (current - 1) * pageSize + 1;
  const lastShown = (current - 1) * pageSize + visible.length;

  const goTo = (p: number): void => {
    setPage(Math.max(1, Math.min(pageCount, p)));
    // Back to the top of the table, so the new page is read from its start.
    if (regionRef.current) regionRef.current.scrollTop = 0;
  };

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
    <>
    <div
      className={[styles.gridRegion, openFilter ? styles.gridRegionFiltering : '']
        .filter(Boolean)
        .join(' ')}
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
          {pageCount > 1 ? `, page ${current} of ${pageCount}` : ''}
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
            visible.map((row) => (
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

    {pageCount > 1 ? (
      <nav className={styles.pager} aria-label={`${caption}: pages`}>
        <span className={styles.pagerInfo} aria-live="polite">
          {`Showing ${firstShown}–${lastShown} of ${ordered.length}`}
        </span>
        <span className={styles.pagerButtons}>
          <button
            type="button"
            className={styles.pagerButton}
            onClick={() => goTo(current - 1)}
            disabled={current === 1}
            aria-label="Previous page"
          >
            <Icon name="chevronLeft" size={14} />
          </button>
          {pageList(current, pageCount).map((p, i) =>
            p === 'gap' ? (
              <span key={`gap${i}`} className={styles.pagerGap} aria-hidden="true">
                &hellip;
              </span>
            ) : (
              <button
                key={p}
                type="button"
                className={[styles.pagerButton, p === current ? styles.pagerButtonActive : '']
                  .filter(Boolean)
                  .join(' ')}
                onClick={() => goTo(p)}
                aria-current={p === current ? 'page' : undefined}
                aria-label={`Page ${p}`}
              >
                {p}
              </button>
            ),
          )}
          <button
            type="button"
            className={styles.pagerButton}
            onClick={() => goTo(current + 1)}
            disabled={current === pageCount}
            aria-label="Next page"
          >
            <Icon name="chevronRight" size={14} />
          </button>
        </span>
      </nav>
    ) : undefined}
    </>
  );
}

export const RowActions: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className={styles.rowActions}>{children}</div>
);
