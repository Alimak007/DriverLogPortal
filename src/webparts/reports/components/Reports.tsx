/**
 * Reports - eight reports sharing one filter bar.
 *
 * The date range is MANDATORY and capped at twelve months. Every report
 * exports to Excel and CSV.
 *
 * Driver response behaviour is restricted to Fleet Admin and Coordinator. It
 * makes individual acceptance rates visible to management, and a report that
 * quietly becomes a performance-management tool without anyone agreeing to it
 * is a problem - so it is both permission-gated and labelled as sensitive
 * where it appears.
 */

import * as React from 'react';
import { DataGrid, IColumn } from '../../../common/components/DataGrid';
import { Icon, IconName } from '../../../common/components/Icons';
import {
  Button,
  EmptyState,
  ErrorBanner,
  FormRow,
  HeadedCard,
  Loading,
  Notice,
  PageShell,
  Select,
  TextField,
} from '../../../common/components/Shared';
import { IVehicleType, REQUEST_TYPES, RequestType } from '../../../common/models';
import { REPORTS } from '../../../common/services/ReportService';
import {
  ICurrentUser,
  IReportFilter,
  IReportResult,
  IReportRow,
  IServices,
} from '../../../common/services/IServices';
import { hasRole } from '../../../common/services/UserService';
import { addMonths, formatDateTime, gulfMonthBounds } from '../../../common/utils/dates';
import { exportCsv, exportExcel } from '../../../common/utils/export';
import styles from './Reports.module.scss';

export interface IReportsProps {
  services?: IServices;
  user?: ICurrentUser;
  bootstrapError?: unknown;
}

/** The drawing on each report's tile. Decoration only. */
const REPORT_ICON: Record<string, IconName> = {
  TripSummary: 'doc',
  VehicleUtilisation: 'car',
  DriverUtilisation: 'user',
  RequestFulfilment: 'check',
  BroadcastEffectiveness: 'chart',
  DriverResponseBehaviour: 'users',
  RejectionCancellationAnalysis: 'x',
  DemandByTimeAndRoute: 'clock',
};

function toDateInput(d: Date): string {
  const two = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${d.getUTCFullYear()}-${two(d.getUTCMonth() + 1)}-${two(d.getUTCDate())}`;
}

export const Reports: React.FC<IReportsProps> = ({ services, user, bootstrapError }) => {
  const initial = React.useMemo(() => gulfMonthBounds(), []);

  const [from, setFrom] = React.useState(toDateInput(initial.from));
  const [to, setTo] = React.useState(toDateInput(initial.to));
  const [requestType, setRequestType] = React.useState<string>('');
  const [vehicleTypeId, setVehicleTypeId] = React.useState<string>('');

  const [reportKey, setReportKey] = React.useState<string>('TripSummary');
  const [result, setResult] = React.useState<IReportResult | undefined>();
  const [running, setRunning] = React.useState(false);
  const [error, setError] = React.useState<unknown>(bootstrapError);

  const [vehicleTypes, setVehicleTypes] = React.useState<IVehicleType[]>([]);

  const maySeeRestricted = user ? hasRole(user, 'admin', 'coordinator') : false;

  const available = React.useMemo(
    () => REPORTS.filter((r) => !r.restricted || maySeeRestricted),
    [maySeeRestricted],
  );

  React.useEffect(() => {
    if (!services) return;
    let cancelled = false;

    (async () => {
      try {
        const types = await services.master.getVehicleTypes();
        if (cancelled) return;
        setVehicleTypes(types);
      } catch (e) {
        if (!cancelled) setError(e);
      }
    })().catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [services]);


  const run = async (key: string): Promise<void> => {
    if (!services) return;
    setRunning(true);
    setError(undefined);
    setReportKey(key);

    try {
      const filter: IReportFilter = {
        from: new Date(`${from}T00:00:00Z`),
        to: new Date(`${to}T23:59:59Z`),
        requestType: (requestType || undefined) as RequestType | undefined,
        vehicleTypeId: vehicleTypeId ? parseInt(vehicleTypeId, 10) : undefined,
      };
      const r = await services.reports.run(key, filter);
      setResult(r);
    } catch (e) {
      setError(e);
      setResult(undefined);
    } finally {
      setRunning(false);
    }
  };

  const rangeTooLong = React.useMemo(() => {
    const f = new Date(`${from}T00:00:00Z`);
    const t = new Date(`${to}T00:00:00Z`);
    if (isNaN(f.getTime()) || isNaN(t.getTime())) return false;
    return t > addMonths(f, 12);
  }, [from, to]);

  const columns: IColumn<IReportRow>[] = React.useMemo(() => {
    if (!result) return [];
    return result.columns.map((c) => ({
      key: c.key,
      label: c.label,
      numeric: c.numeric,
      render: (row: IReportRow) => {
        const v = row[c.key];
        return v === undefined || v === null || v === '' ? '—' : String(v);
      },
      // A filter list is only offered where the column is categorical. A
      // measure column would produce a filter list of every distinct number,
      // which helps nobody.
      filterValue: c.numeric ? undefined : (row: IReportRow) => String(row[c.key] ?? ''),
    }));
  }, [result]);

  if (!services || !user) {
    return (
      <PageShell heading="Fleet reports">
        <ErrorBanner error={bootstrapError} />
        {!bootstrapError ? <Loading /> : undefined}
      </PageShell>
    );
  }

  const exports = result ? (
    <>
      <Button kind="secondary" onClick={() => exportExcel(result)}>
        <Icon name="download" size={16} />
        Excel
      </Button>
      <Button kind="secondary" onClick={() => exportCsv(result)}>
        <Icon name="download" size={16} />
        CSV
      </Button>
    </>
  ) : undefined;

  return (
    <PageShell
      heading="Fleet reports"
      caption="Choose a period and a report. Every report exports to Excel and CSV."
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      <div className={styles.layout}>
        <HeadedCard icon="calendar" title="Filters" className={styles.filtersCard}>
          <div className={styles.filterGrid}>
            <FormRow label="From" required htmlFor="report-from">
              <TextField
                id="report-from"
                type="date"
                value={from}
                onChange={(e) => setFrom(e.target.value)}
              />
            </FormRow>

            <FormRow
              label="To"
              required
              htmlFor="report-to"
              error={rangeTooLong ? 'The range cannot be longer than 12 months.' : undefined}
            >
              <TextField
                id="report-to"
                type="date"
                value={to}
                onChange={(e) => setTo(e.target.value)}
              />
            </FormRow>

            {/*
              NO PROPERTY FILTER. Allocation stopped being scoped by property
              when the solution moved to a single driver pool, so filtering a
              report by it narrowed the rows without narrowing anything the
              reader cares about. The column is still written on the request
              for reference; it is simply not a lens on these reports.
            */}

            <FormRow label="Request type" htmlFor="report-request-type">
              <Select
                id="report-request-type"
                value={requestType}
                onChange={(e) => setRequestType(e.target.value)}
              >
                <option value="">All request types</option>
                {REQUEST_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </Select>
            </FormRow>

            <FormRow label="Vehicle type" htmlFor="report-vehicle-type">
              <Select
                id="report-vehicle-type"
                value={vehicleTypeId}
                onChange={(e) => setVehicleTypeId(e.target.value)}
              >
                <option value="">All vehicle types</option>
                {vehicleTypes.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.title}
                  </option>
                ))}
              </Select>
            </FormRow>
          </div>

          <Button
            kind="primary"
            className={styles.runButton}
            disabled={running || rangeTooLong}
            onClick={() => {
              run(reportKey).catch(() => undefined);
            }}
          >
            {running ? 'Running…' : 'Run report'}
            {running ? undefined : <Icon name="arrowRight" size={18} />}
          </Button>
        </HeadedCard>

        <HeadedCard icon="chart" title="Reports" className={styles.reportsCard}>
          <div className={styles.reportList} role="group" aria-label="Available reports">
            {available.map((r) => {
              const active = reportKey === r.key;
              return (
                <button
                  key={r.key}
                  type="button"
                  className={[styles.reportTile, active ? styles.reportTileActive : ''].join(' ')}
                  aria-pressed={active}
                  onClick={() => {
                    run(r.key).catch(() => undefined);
                  }}
                >
                  <span className={styles.reportIcon} aria-hidden="true">
                    <Icon name={REPORT_ICON[r.key] ?? 'chart'} size={22} />
                  </span>
                  <span className={styles.reportText}>
                    <span className={styles.reportTitle}>{r.title}</span>
                    {r.restricted ? (
                      <span className={styles.restrictedMark}>
                        <Icon name="lock" size={12} />
                        Restricted &middot; shows individual acceptance rates
                      </span>
                    ) : undefined}
                  </span>
                  {active ? (
                    <span className={styles.checkMark}>
                      <Icon name="check" size={13} />
                      <span className={styles.srOnly}>Selected</span>
                    </span>
                  ) : undefined}
                </button>
              );
            })}
          </div>
        </HeadedCard>
      </div>

      {reportKey === 'DriverResponseBehaviour' ? (
        <Notice tone="attention">
          This report makes individual drivers&apos; acceptance rates visible.
          Agree with the business how it will be used before relying on it for
          anything other than tuning the eligibility rules.
        </Notice>
      ) : undefined}

      {reportKey === 'BroadcastEffectiveness' ? (
        <Notice>
          This is the report that proves the business case. If the escalation
          rate is high, automatic allocation is not working and the eligibility
          rules, the offer window or driver engagement need attention. Review it
          weekly for the first two months after go-live.
        </Notice>
      ) : undefined}

      <HeadedCard
        icon="doc"
        title={result && !running ? result.title : 'Results'}
        actions={running ? undefined : exports}
      >
        {running ? (
          <Loading what="Running the report" />
        ) : !result ? (
          <EmptyState title="No report has been run yet">
            Set a period above and choose a report.
          </EmptyState>
        ) : (
          <>
            <div className={styles.resultMeta}>
              <span className={styles.metaPill}>
                <Icon name="hash" size={14} />
                {result.rows.length} row{result.rows.length === 1 ? '' : 's'}
              </span>
              <span className={styles.metaPill}>
                <Icon name="clock" size={14} />
                generated {formatDateTime(result.generatedAt)}
              </span>
              <span className={styles.metaPill}>
                {result.fromSnapshot ? 'from nightly snapshot' : 'live data'}
              </span>
            </div>

            <DataGrid
              caption={`${result.title}, ${from} to ${to}`}
              captionVisuallyHidden
              columns={columns}
              rows={result.rows}
              rowKey={(row) => JSON.stringify(row).slice(0, 120)}
              emptyMessage="No data for that period and those filters."
            />
          </>
        )}
      </HeadedCard>
    </PageShell>
  );
};

export default Reports;
