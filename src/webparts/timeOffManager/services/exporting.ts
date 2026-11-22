import { IConfig, ILeaveRequest } from '../models/types';
import { addDays } from './dates';

/** CSV text; cells that start with = + - @ are prefixed so spreadsheets do not run them as formulas. */
export function toCsv(headers: string[], rows: Array<Array<string | number | undefined>>): string {
  const cell = (v: string | number | undefined): string => {
    let s = v === undefined || v === null ? '' : String(v);
    if (/^[=+\-@\t\r]/.test(s) && !/^-?\d+(\.\d+)?$/.test(s)) s = "'" + s;
    return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return '\ufeff' + [headers].concat(rows as unknown as string[][]).map((r) => r.map(cell).join(',')).join('\r\n');
}

export function download(name: string, text: string, type: string = 'text/csv;charset=utf-8'): void {
  const blob = new Blob([text], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(a.href);
    a.remove();
  }, 500);
}

function icsText(s: string): string {
  return (s || '').replace(/\\/g, '\\\\').replace(/;/g, '\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
}

const compact = (iso: string): string => iso.replace(/-/g, '');

/** iCalendar file of approved leave (Outlook, Teams, Google and Apple calendars all import it). */
export function toIcs(config: IConfig, requests: ILeaveRequest[], typeName: (key: string) => string): string {
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Brillienta//Time Off Manager//EN', 'CALSCALE:GREGORIAN', `X-WR-CALNAME:${icsText(config.appName)}`];
  requests.forEach((r) => {
    lines.push(
      'BEGIN:VEVENT',
      `UID:to-${r.id}@brillienta-timeoff`,
      `DTSTAMP:${stamp}`,
      `DTSTART;VALUE=DATE:${compact(r.start)}`,
      `DTEND;VALUE=DATE:${compact(addDays(r.end || r.start, 1))}`,
      `SUMMARY:${icsText(`${r.employeeName} – ${typeName(r.leaveType)}${r.unit === 'half' ? ' (half day)' : r.unit === 'hours' ? ` (${r.hours} h)` : ''}`)}`,
      'TRANSP:TRANSPARENT',
      'X-MICROSOFT-CDO-BUSYSTATUS:OOF',
      'END:VEVENT'
    );
  });
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

/** Parse a simple CSV (quoted cells, commas or semicolons). */
export function parseCsv(text: string): string[][] {
  const src = text.replace(/^\ufeff/, '');
  const sep = (src.split('\n')[0] || '').split(';').length > (src.split('\n')[0] || '').split(',').length ? ';' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else cell += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === sep) {
      row.push(cell.trim());
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      row.push(cell.trim());
      if (row.some((c) => c !== '')) rows.push(row);
      row = [];
      cell = '';
    } else cell += ch;
  }
  row.push(cell.trim());
  if (row.some((c) => c !== '')) rows.push(row);
  return rows;
}
