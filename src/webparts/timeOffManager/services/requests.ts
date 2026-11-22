// Request drafts: how long they are, whether they are allowed, and recurring series. Pure functions
// (unit tested in tests/requests.test.cjs).

import { IConfig, IEmployee, ILeaveRequest, ILeaveType, IRecurrence, LeaveUnit } from '../models/types';
import { IBalance } from './balances';
import { addDays, addMonths, calendarFor, countDays, daysBetween, eachDay, fmtDate, isIsoDate, isWorkingDay, leaveYearOf, weekday, WEEKDAY_SHORT } from './dates';
import { lc } from './workflow';

export interface IDraft {
  kind: 'leave' | 'compoff';
  employeeEmail: string;
  leaveType: string;
  start: string;
  end: string;
  unit: LeaveUnit;
  halfDayPart?: 'am' | 'pm';
  startTime?: string;
  hours?: number;
  reason: string;
  recurrence?: IRecurrence;
  hasAttachment?: boolean;
}

export interface IDuration {
  days: number;
  hours: number;
  excluded: Array<{ date: string; why: string }>;
  counted: string[];
}

/** Days (and hours) a draft takes from the balance. */
export function duration(config: IConfig, type: ILeaveType | undefined, e: IEmployee | undefined, d: Pick<IDraft, 'start' | 'end' | 'unit' | 'hours' | 'kind'>): IDuration {
  const cal = calendarFor(config, e);
  const calendarDays = !!type && type.countCalendarDays;
  if (d.kind === 'compoff') {
    // A comp-off claim credits the days worked (a non-working day is the point).
    const n = d.unit === 'day' ? Math.max(0, daysBetween(d.start, d.end) + 1) : 0;
    if (d.unit === 'half') return { days: 0.5, hours: cal.hoursPerDay / 2, excluded: [], counted: [d.start] };
    if (d.unit === 'hours') return { days: round2((d.hours || 0) / cal.hoursPerDay), hours: d.hours || 0, excluded: [], counted: [d.start] };
    return { days: n, hours: n * cal.hoursPerDay, excluded: [], counted: isIsoDate(d.start) && n ? eachDay(d.start, d.end) : [] };
  }
  if (d.unit === 'half' || d.unit === 'hours') {
    const c = countDays(cal, d.start, d.start, calendarDays);
    if (!c.days) return { days: 0, hours: 0, excluded: c.excluded, counted: [] };
    if (d.unit === 'half') return { days: 0.5, hours: cal.hoursPerDay / 2, excluded: [], counted: c.counted };
    const h = d.hours || 0;
    return { days: round2(h / cal.hoursPerDay), hours: h, excluded: [], counted: c.counted };
  }
  const c = countDays(cal, d.start, d.end, calendarDays);
  return { days: c.days, hours: c.days * cal.hoursPerDay, excluded: c.excluded, counted: c.counted };
}

const round2 = (n: number): number => Math.round(n * 100) / 100;

export interface IValidation {
  errors: string[];
  warnings: string[];
}

/** Rules a draft must meet before it can be submitted. */
export function validateDraft(
  config: IConfig,
  type: ILeaveType | undefined,
  e: IEmployee | undefined,
  d: IDraft,
  dur: IDuration,
  balance: IBalance | undefined,
  existing: ILeaveRequest[],
  today: string,
  privileged: boolean
): IValidation {
  const errors: string[] = [];
  const warnings: string[] = [];
  if (!e) errors.push('Choose who the request is for.');
  if (!type) {
    errors.push('Choose a leave type.');
    return { errors, warnings };
  }
  if (!isIsoDate(d.start) || (d.unit === 'day' && !isIsoDate(d.end))) {
    errors.push('Choose the dates.');
    return { errors, warnings };
  }
  if (d.unit === 'day' && d.end < d.start) errors.push('The end date is before the start date.');
  if (d.kind === 'compoff') {
    const cal = calendarFor(config, e);
    if (d.start > today) errors.push('A comp-off claim is for a day already worked.');
    if (config.compOff.requireNonWorkingDay) {
      const days = d.unit === 'day' ? eachDay(d.start, d.end) : [d.start];
      const working = days.filter((x) => isWorkingDay(cal, x));
      if (working.length) errors.push(`Comp off is earned on weekends and holidays; ${working.map((x) => fmtDate(x)).join(', ')} ${working.length === 1 ? 'is a working day' : 'are working days'}.`);
    }
    if (config.compOff.maxDaysPerClaim > 0 && dur.days > config.compOff.maxDaysPerClaim) errors.push(`A claim can be at most ${config.compOff.maxDaysPerClaim} days.`);
    if (!d.reason.trim()) errors.push('Say what you worked on.');
    return { errors, warnings };
  }
  const allowed = type.units[d.unit];
  if (!allowed) errors.push(`${type.name} cannot be taken ${d.unit === 'half' ? 'as a half day' : d.unit === 'hours' ? 'by the hour' : 'in full days'}.`);
  if (d.unit === 'hours' && !(d.hours && d.hours > 0)) errors.push('Choose how many hours.');
  if (dur.days <= 0) errors.push('These dates are all weekends or holidays, so no leave is needed.');
  const el = balance ? { ok: balance.eligible, why: balance.why } : { ok: true, why: undefined };
  if (!el.ok) errors.push(el.why || 'Not eligible.');
  const lyStart = leaveYearOf(d.start, config.yearStartMonth);
  const lyEnd = leaveYearOf(d.unit === 'day' ? d.end : d.start, config.yearStartMonth);
  if (lyStart !== lyEnd) errors.push('The dates run into the next leave year. Make one request for each leave year.');
  if (type.maxConsecutiveDays > 0 && dur.days > type.maxConsecutiveDays) errors.push(`${type.name} can be at most ${type.maxConsecutiveDays} days at a time.`);
  if (type.minNoticeDays > 0 && daysBetween(today, d.start) < type.minNoticeDays) {
    const msg = `${type.name} needs ${type.minNoticeDays} days' notice.`;
    if (privileged) warnings.push(msg + ' (You may submit anyway.)');
    else errors.push(msg);
  }
  if (d.start < today && !config.cancellation.allowPast && !privileged && type.key !== 'sick') warnings.push('This request is for days in the past.');
  if (type.attachmentOverDays > 0 && dur.days > type.attachmentOverDays && !d.hasAttachment) {
    errors.push(`Attach a supporting document (needed for ${type.name} over ${type.attachmentOverDays} day${type.attachmentOverDays === 1 ? '' : 's'}).`);
  }
  if (balance && !balance.unlimited) {
    const after = balance.available - dur.days;
    if (after < -type.allowNegative - 1e-9) {
      errors.push(`Not enough ${type.name}: ${round2(balance.available)} available, ${dur.days} requested.`);
    } else if (after < 0) warnings.push(`This takes the balance to ${round2(after)} days.`);
  }
  const clash = overlapping(existing, d);
  if (clash.length) errors.push(`You already have time off on these days (${clash.map((r) => fmtDate(r.start)).join(', ')}).`);
  return { errors, warnings };
}

/** Open requests of the same person whose days overlap the draft's. Hour leaves on one day may share it. */
export function overlapping(existing: ILeaveRequest[], d: Pick<IDraft, 'employeeEmail' | 'start' | 'end' | 'unit' | 'halfDayPart'>): ILeaveRequest[] {
  const end = d.unit === 'day' ? d.end : d.start;
  return existing.filter((r) => {
    if (r.kind !== 'leave' || lc(r.employeeEmail) !== lc(d.employeeEmail)) return false;
    if (r.status !== 'Pending' && r.status !== 'Approved' && r.status !== 'CancelRequested') return false;
    if (r.start > end || r.end < d.start) return false;
    if (d.unit === 'hours' && r.unit === 'hours') return false;
    if (d.unit === 'half' && r.unit === 'half' && r.halfDayPart !== d.halfDayPart) return false;
    return true;
  });
}

// ------------------------------------------------------------------ recurrence (weekly, bi-weekly, monthly)

export const MAX_OCCURRENCES = 52;

/** Start days of a recurring series (the first occurrence included). */
export function occurrences(start: string, rule: IRecurrence | undefined, horizon?: string): string[] {
  if (!rule || rule.freq === 'none' || !isIsoDate(start)) return [start];
  const until = rule.until && isIsoDate(rule.until) ? rule.until : horizon || addMonths(start, 12);
  const count = rule.count && rule.count > 0 ? Math.min(rule.count, MAX_OCCURRENCES) : MAX_OCCURRENCES;
  const out: string[] = [];
  if (rule.freq === 'monthly') {
    for (let i = 0; out.length < count; i++) {
      const d = addMonths(start, i);
      if (d > until) break;
      out.push(d);
    }
    return out;
  }
  const step = rule.freq === 'biweekly' ? 14 : 7;
  const days = rule.weekdays && rule.weekdays.length ? rule.weekdays.slice().sort() : [weekday(start)];
  // Weeks are counted from the week of the start date (Sunday based), so "every other week" keeps its rhythm.
  const week0 = addDays(start, -weekday(start));
  for (let w = 0; out.length < count && w < 400; w += step) {
    const ws = addDays(week0, w);
    if (ws > until) break;
    for (const wd of days) {
      const d = addDays(ws, wd);
      if (d < start || d > until) continue;
      out.push(d);
      if (out.length >= count) break;
    }
  }
  return out;
}

export function describeRecurrence(rule: IRecurrence | undefined): string {
  if (!rule || rule.freq === 'none') return '';
  const days = rule.weekdays && rule.weekdays.length ? ' on ' + rule.weekdays.map((d) => WEEKDAY_SHORT[d]).join(', ') : '';
  const base = rule.freq === 'weekly' ? `Every week${days}` : rule.freq === 'biweekly' ? `Every other week${days}` : 'Every month';
  const end = rule.count ? `, ${rule.count} times` : rule.until ? `, until ${fmtDate(rule.until)}` : '';
  return base + end;
}

export function newSeriesId(): string {
  return 'S' + Date.now().toString(36) + Math.random().toString(36).substring(2, 6);
}
