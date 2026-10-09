/**
 * Export a report to CSV and to Excel.
 *
 * NO LIBRARY. The bundle budget is 500 KB gzipped per web part, and a
 * spreadsheet library costs more than that on its own. CSV is written by hand;
 * the Excel file is a SpreadsheetML workbook, which Excel opens natively and
 * which keeps column types and the sheet name - something a .csv cannot do.
 */

import { IReportResult } from '../services/IServices';

function escapeCsv(value: string): string {
  if (value.indexOf('"') === -1 && value.indexOf(',') === -1 && value.indexOf('\n') === -1) {
    return value;
  }
  return `"${value.replace(/"/g, '""')}"`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function cellText(value: string | number | undefined): string {
  if (value === undefined || value === null) return '';
  return String(value);
}

function download(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  // Revoked on the next tick so the click has taken effect first.
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function stamp(): string {
  const d = new Date();
  const two = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${d.getFullYear()}${two(d.getMonth() + 1)}${two(d.getDate())}-${two(d.getHours())}${two(d.getMinutes())}`;
}

export function exportCsv(report: IReportResult): void {
  const header = report.columns.map((c) => escapeCsv(c.label)).join(',');
  const body = report.rows
    .map((row) =>
      report.columns.map((c) => escapeCsv(cellText(row[c.key]))).join(','),
    )
    .join('\r\n');

  // The BOM is what makes Excel read UTF-8 correctly on a double-click.
  const blob = new Blob([`﻿${header}\r\n${body}`], {
    type: 'text/csv;charset=utf-8;',
  });
  download(blob, `${report.key}-${stamp()}.csv`);
}

export function exportExcel(report: IReportResult): void {
  const rows = report.rows
    .map((row) => {
      const cells = report.columns
        .map((c) => {
          const raw = row[c.key];
          const isNumber = c.numeric === true && typeof raw === 'number';
          const type = isNumber ? 'Number' : 'String';
          return `<Cell><Data ss:Type="${type}">${escapeXml(cellText(raw))}</Data></Cell>`;
        })
        .join('');
      return `<Row>${cells}</Row>`;
    })
    .join('');

  const header = report.columns
    .map((c) => `<Cell ss:StyleID="h"><Data ss:Type="String">${escapeXml(c.label)}</Data></Cell>`)
    .join('');

  const xml =
    '<?xml version="1.0"?>' +
    '<?mso-application progid="Excel.Sheet"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' +
    'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Styles><Style ss:ID="h"><Font ss:Bold="1"/></Style></Styles>' +
    `<Worksheet ss:Name="${escapeXml(report.title.slice(0, 31))}">` +
    `<Table><Row>${header}</Row>${rows}</Table>` +
    '</Worksheet></Workbook>';

  const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  download(blob, `${report.key}-${stamp()}.xls`);
}
