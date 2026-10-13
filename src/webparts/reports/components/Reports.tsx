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
import {
  Button,
  EmptyState,
  ErrorBanner,
  FormRow,
  Loading,
  Notice,
  PageShell,
  SectionHeading,
  Select,
  TextField,
} from '../../../common/components/Shared';
import { IProperty, IVehicleType, REQUEST_TYPES, RequestType } from '../../../common/models';
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
  const [propertyId, setPropertyId] = React.useState<string>('');

  const [reportKey, setReportKey] = React.useState<string>('TripSummary');
  const [result, setResult] = React.useState<IReportResult | undefined>();
  const [running, setRunning] = React.useState(false);
  const [error, setError] = React.useState<unknown>(bootstrapError);

  const [vehicleTypes, setVehicleTypes] = React.useState<IVehicleType[]>([]);
  const [properties, setProperties] = React.useState<IProperty[]>([]);

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
        const [types, props] = await Promise.all([
          services.master.getVehicleTypes(),
          services.master.getProperties(),
        ]);
        if (cancelled) return;
        setVehicleTypes(types);
        setProperties(props);
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
        propertyId: propertyId ? parseInt(propertyId, 10) : undefined,
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

  return (
    <PageShell
      heading="Fleet reports"
      caption="Choose a period and a report. Every report exports to Excel and CSV."
    >
      <ErrorBanner error={error} onDismiss={() => setError(undefined)} />

      <div className={styles.filterBar}>
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

        <FormRow label="Property" htmlFor="report-property">
          <Select
            id="report-property"
            value={propertyId}
            onChange={(e) => setPropertyId(e.target.value)}
          >
            <option value="">All properties</option>
            {properties.map((p) => (
              <option key={p.id} value={p.id}>
                {p.title}
              </option>
            ))}
          </Select>
        </FormRow>

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

        <div className={styles.filterActions}>
          <Button
            kind="primary"
            disabled={running || rangeTooLong}
            onClick={() => {
              run(reportKey).catch(() => undefined);
            }}
          >
            {running ? 'Running…' : 'Run report'}
          </Button>
        </div>
      </div>

      <SectionHeading>Reports</SectionHeading>
      <div className={styles.reportList} role="group" aria-label="Available reports">
        {available.map((r) => (
          <button
            key={r.key}
            type="button"
            className={[
              styles.reportButton,
              reportKey === r.key ? styles.reportButtonActive : '',
            ].join(' ')}
            aria-pressed={reportKey === r.key}
            onClick={() => {
              run(r.key).catch(() => undefined);
            }}
          >
            {r.title}
            {r.restricted ? (
              <span className={styles.restrictedMark}>
                Restricted &middot; shows individual acceptance rates
              </span>
            ) : undefined}
          </button>
        ))}
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

      {running ? (
        <Loading what="Running the report" />
      ) : !result ? (
        <EmptyState title="No report has been run yet">
          Set a period above and choose a report.
        </EmptyState>
      ) : (
        <>
          <div className={styles.resultHead}>
            <SectionHeading>{result.title}</SectionHeading>
            <div className={styles.filterActions}>
              <span className={styles.generated}>
                {result.rows.length} row{result.rows.length === 1 ? '' : 's'}
                {' · '}generated {formatDateTime(result.generatedAt)}
                {result.fromSnapshot ? ' · from nightly snapshot' : ' · live data'}
              </span>
              <Button kind="secondary" onClick={() => exportExcel(result)}>
                Export to Excel
              </Button>
              <Button kind="secondary" onClick={() => exportCsv(result)}>
                Export to CSV
              </Button>
            </div>
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
    </PageShell>
  );
};

export default Reports;
