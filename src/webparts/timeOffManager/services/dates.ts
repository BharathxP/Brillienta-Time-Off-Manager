// Calendar-day arithmetic on "yyyy-mm-dd" strings (UTC based, so no time-zone or daylight-saving drift).
import { IConfig, IEmployee, IWorkCalendar } from '../models/types';

const DAY = 86400000;

export function toIso(d: Date): string {
  return d.toISOString().substring(0, 10);
}

export function parseIso(iso: string): Date {
  const [y, m, d] = (iso || '').substring(0, 10).split('-').map(Number);
  return new Date(Date.UTC(y || 1970, (m || 1) - 1, d || 1));
}

/** Today in the browser's time zone, as a calendar day. */
export function today(now: Date = new Date()): string {
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

export function pad(n: number): string {
  return (n < 10 ? '0' : '') + n;
}

export function addDays(iso: string, n: number): string {
  return toIso(new Date(parseIso(iso).getTime() + n * DAY));
}

export function addMonths(iso: string, n: number): string {
  const d = parseIso(iso);
  const day = d.getUTCDate();
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n, 1));
  const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
  t.setUTCDate(Math.min(day, last));
  return toIso(t);
}

export function weekday(iso: string): number {
  return parseIso(iso).getUTCDay();
}

export function daysBetween(a: string, b: string): number {
  return Math.round((parseIso(b).getTime() - parseIso(a).getTime()) / DAY);
}

export function isIsoDate(s: string | undefined): boolean {
  return !!s && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(parseIso(s).getTime());
}

/** Every day from start to end inclusive. */
export function eachDay(start: string, end: string, max: number = 1000): string[] {
  const out: string[] = [];
  let d = start;
  while (d <= end && out.length < max) {
    out.push(d);
    d = addDays(d, 1);
  }
  return out;
}

export function startOfWeek(iso: string, weekStartsOn: number): string {
  const wd = weekday(iso);
  return addDays(iso, -((wd - weekStartsOn + 7) % 7));
}

export function monthStart(iso: string): string {
  return iso.substring(0, 8) + '01';
}

export function monthEnd(iso: string): string {
  return addDays(addMonths(monthStart(iso), 1), -1);
}

// ------------------------------------------------------------------ leave years

/** The leave year a day falls in, named after the calendar year it starts in. */
export function leaveYearOf(iso: string, startMonth: number): number {
  const [y, m] = iso.split('-').map(Number);
  return m >= startMonth ? y : y - 1;
}

export function leaveYearStart(year: number, startMonth: number): string {
  return `${year}-${pad(startMonth)}-01`;
}

export function leaveYearEnd(year: number, startMonth: number): string {
  return addDays(leaveYearStart(year + 1, startMonth), -1);
}

export function leaveYearLabel(year: number, startMonth: number): string {
  return startMonth === 1 ? String(year) : `${year}–${String(year + 1).substring(2)}`;
}

// ------------------------------------------------------------------ work calendars

export function calendarFor(config: IConfig, employee?: IEmployee): IWorkCalendar {
  const key = (employee && employee.calendar) || config.defaultCalendar;
  return config.calendars.filter((c) => c.key === key)[0] || config.calendars.filter((c) => c.key === config.defaultCalendar)[0] || config.calendars[0] || FALLBACK_CALENDAR;
}

export const FALLBACK_CALENDAR: IWorkCalendar = { key: 'default', name: 'Mon–Fri', workDays: [1, 2, 3, 4, 5], hoursPerDay: 8, dayStart: '09:00', holidays: [] };

export function holidayOn(cal: IWorkCalendar, iso: string): string | undefined {
  const h = cal.holidays.filter((x) => x.date === iso && !x.optional)[0];
  return h ? h.name : undefined;
}

export function isWorkingDay(cal: IWorkCalendar, iso: string): boolean {
  return cal.workDays.indexOf(weekday(iso)) >= 0 && !holidayOn(cal, iso);
}

/** The first working day on or after a day (the day itself when no working day is found within a month). */
export function nextWorkingDay(cal: IWorkCalendar, from: string): string {
  for (let i = 0; i < 31; i++) {
    const d = addDays(from, i);
    if (isWorkingDay(cal, d)) return d;
  }
  return from;
}

export interface IDayCount {
  /** Days that count against the balance. */
  days: number;
  /** Days skipped (weekends, holidays) with the reason. */
  excluded: Array<{ date: string; why: string }>;
  /** The counted days. */
  counted: string[];
}

const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
export const WEEKDAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

export function weekdayName(n: number): string {
  return WEEKDAY_NAMES[n] || '';
}

/** Working days between two dates, leaving out non-working weekdays and holidays unless calendar days count. */
export function countDays(cal: IWorkCalendar, start: string, end: string, calendarDays: boolean): IDayCount {
  const out: IDayCount = { days: 0, excluded: [], counted: [] };
  if (!isIsoDate(start) || !isIsoDate(end) || end < start) return out;
  eachDay(start, end).forEach((d) => {
    if (!calendarDays) {
      const h = holidayOn(cal, d);
      if (h) {
        out.excluded.push({ date: d, why: h });
        return;
      }
      if (cal.workDays.indexOf(weekday(d)) < 0) {
        out.excluded.push({ date: d, why: `${weekdayName(weekday(d))} (non-working day)` });
        return;
      }
    }
    out.counted.push(d);
    out.days += 1;
  });
  return out;
}

// ------------------------------------------------------------------ formatting

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

export function fmtDate(iso?: string, withWeekday: boolean = false): string {
  if (!iso || !isIsoDate(iso.substring(0, 10))) return '';
  const d = parseIso(iso);
  const s = `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  return withWeekday ? `${WEEKDAY_SHORT[d.getUTCDay()]}, ${s}` : s;
}

export function fmtRange(start: string, end: string): string {
  if (!start) return '';
  if (!end || end === start) return fmtDate(start, true);
  const a = parseIso(start);
  const b = parseIso(end);
  if (a.getUTCFullYear() === b.getUTCFullYear()) {
    const left = a.getUTCMonth() === b.getUTCMonth() ? `${a.getUTCDate()}` : `${a.getUTCDate()} ${MONTHS[a.getUTCMonth()]}`;
    return `${left} – ${b.getUTCDate()} ${MONTHS[b.getUTCMonth()]} ${b.getUTCFullYear()}`;
  }
  return `${fmtDate(start)} – ${fmtDate(end)}`;
}

export function fmtDays(n: number): string {
  const r = Math.round(n * 100) / 100;
  return `${r} day${r === 1 ? '' : 's'}`;
}

export function fmtNum(n: number): string {
  return String(Math.round(n * 100) / 100);
}

export function relativeTime(iso?: string, now: number = Date.now()): string {
  if (!iso) return '';
  const t = Date.parse(iso);
  if (isNaN(t)) return '';
  const s = Math.round((now - t) / 1000);
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  if (s < 86400 * 7) return `${Math.floor(s / 86400)} d ago`;
  return fmtDate(new Date(t).toISOString());
}

/** "09:00" + hours → "11:00". */
export function addTime(hhmm: string, hours: number): string {
  const [h, m] = (hhmm || '09:00').split(':').map(Number);
  const total = Math.round(((h || 0) * 60 + (m || 0) + hours * 60) % (24 * 60));
  return `${pad(Math.floor(total / 60))}:${pad(total % 60)}`;
}
