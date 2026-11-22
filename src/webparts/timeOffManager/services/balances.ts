// Entitlements and balances. Pure functions (unit tested in tests/balances.test.cjs).
//
// A balance is never stored: it is worked out from the policy (grade, tenure, accrual, proration), the ledger
// (adjustments, carry-forward, comp-off credits) and the requests. So a policy change applies at once and the
// numbers can always be explained line by line.

import { IConfig, IEmployee, ILeaveRequest, ILeaveType, ILedgerEntry } from '../models/types';
import { addDays, addMonths, daysBetween, leaveYearEnd, leaveYearStart, parseIso } from './dates';

export interface IBalance {
  leaveType: string;
  year: number;
  unlimited: boolean;
  /** Days for the whole year from the policy (accrual: what the full year accrues). */
  annual: number;
  /** Policy days available now (fixed: the prorated year; accrual: accrued so far). */
  entitled: number;
  tenureExtra: number;
  carried: number;
  adjusted: number;
  /** Comp off: credits earned and not lapsed. */
  earned: number;
  expired: number;
  used: number;
  pending: number;
  available: number;
  eligible: boolean;
  why?: string;
}

const round2 = (n: number): number => Math.round(n * 100) / 100;
const roundHalf = (n: number): number => Math.round(n * 2) / 2;

/** Completed years of service on a day. */
export function serviceYears(hireDate: string | undefined, on: string): number {
  if (!hireDate) return 0;
  const h = parseIso(hireDate);
  const d = parseIso(on);
  let y = d.getUTCFullYear() - h.getUTCFullYear();
  if (d.getUTCMonth() < h.getUTCMonth() || (d.getUTCMonth() === h.getUTCMonth() && d.getUTCDate() < h.getUTCDate())) y--;
  return Math.max(0, y);
}

export function serviceMonths(hireDate: string | undefined, on: string): number {
  if (!hireDate) return 999;
  const h = parseIso(hireDate);
  const d = parseIso(on);
  let m = (d.getUTCFullYear() - h.getUTCFullYear()) * 12 + d.getUTCMonth() - h.getUTCMonth();
  if (d.getUTCDate() < h.getUTCDate()) m--;
  return Math.max(0, m);
}

/** Whether a person may take a leave type (grade, gender, length of service). */
export function eligibility(type: ILeaveType, e: IEmployee | undefined, on: string): { ok: boolean; why?: string } {
  if (!type.active) return { ok: false, why: `${type.name} is switched off.` };
  const el = type.eligibility;
  if (!e) return { ok: true };
  if (el.grades.length && el.grades.indexOf(e.grade || '') < 0) return { ok: false, why: `${type.name} is for grades ${el.grades.join(', ')}.` };
  if (el.genders.length && e.gender && el.genders.indexOf(e.gender) < 0) return { ok: false, why: `${type.name} is not available for this employee.` };
  if (el.minServiceMonths > 0 && serviceMonths(e.hireDate, on) < el.minServiceMonths) {
    return { ok: false, why: `${type.name} needs ${el.minServiceMonths} months of service.` };
  }
  return { ok: true };
}

/** Annual days from the policy: the grade's days (or the standard days) plus the tenure tier reached on `on`. */
export function annualDays(type: ILeaveType, e: IEmployee | undefined, on: string): { days: number; tenureExtra: number } {
  const en = type.entitlement;
  const grade = e && e.grade;
  const base = grade && en.byGrade && typeof en.byGrade[grade] === 'number' ? en.byGrade[grade] : en.days;
  const years = serviceYears(e && e.hireDate, on);
  let extra = 0;
  (en.tenure || []).forEach((t) => {
    if (years >= t.years && t.extraDays > extra) extra = t.extraDays;
  });
  return { days: base + extra, tenureExtra: extra };
}

export const PERIODS: Record<string, number> = { monthly: 12, quarterly: 4, biweekly: 26, yearly: 1 };

/** Accrual periods of a leave year: [start, end] pairs. */
export function accrualPeriods(year: number, startMonth: number, frequency: string): Array<[string, string]> {
  const yStart = leaveYearStart(year, startMonth);
  const yEnd = leaveYearEnd(year, startMonth);
  const out: Array<[string, string]> = [];
  if (frequency === 'biweekly') {
    let s = yStart;
    for (let i = 0; i < 26; i++) {
      const e = i === 25 ? yEnd : addDays(s, 13);
      out.push([s, e]);
      s = addDays(e, 1);
    }
    return out;
  }
  const months = 12 / (PERIODS[frequency] || 12);
  let s = yStart;
  while (s <= yEnd) {
    const next = addMonths(s, months);
    out.push([s, addDays(next, -1)]);
    s = next;
  }
  return out;
}

/** Policy days in force on `asOf` for the leave year (before ledger and requests). */
export function policyEntitlement(type: ILeaveType, e: IEmployee | undefined, year: number, startMonth: number, asOf: string): { annual: number; entitled: number; tenureExtra: number } {
  const en = type.entitlement;
  if (en.mode === 'unlimited' || en.mode === 'earned') return { annual: 0, entitled: 0, tenureExtra: 0 };
  const yStart = leaveYearStart(year, startMonth);
  const yEnd = leaveYearEnd(year, startMonth);
  const hire = e && e.hireDate;
  const leaving = e && e.leavingDate;
  if (hire && hire > yEnd) return { annual: 0, entitled: 0, tenureExtra: 0 };
  if (leaving && leaving < yStart) return { annual: 0, entitled: 0, tenureExtra: 0 };
  if (en.mode === 'fixed') {
    const a = annualDays(type, e, yStart);
    let entitled = a.days;
    // Joiners (when the type prorates) and leavers (always) get the share of the year they work.
    const from = en.prorate && hire && hire > yStart ? hire : yStart;
    const to = leaving && leaving < yEnd ? leaving : yEnd;
    if (from !== yStart || to !== yEnd) {
      const span = daysBetween(yStart, yEnd) + 1;
      entitled = roundHalf((a.days * (daysBetween(from, to) + 1)) / span);
    }
    return { annual: a.days, entitled, tenureExtra: a.tenureExtra };
  }
  // Accrual: each period adds annual / periods, once the person works there at the period's start or end.
  const a = annualDays(type, e, asOf < yStart ? yStart : asOf > yEnd ? yEnd : asOf);
  const periods = accrualPeriods(year, startMonth, en.accrual.frequency);
  const per = a.days / periods.length;
  let accrued = 0;
  let full = 0;
  periods.forEach(([ps, pe]) => {
    const point = en.accrual.timing === 'start' ? ps : pe;
    if (hire && en.prorate && hire > point) return;
    if (leaving && leaving < ps) return;
    full += per;
    if (point <= asOf) accrued += per;
  });
  return { annual: round2(full), entitled: round2(accrued), tenureExtra: a.tenureExtra };
}

/** Requests that use a balance: approved (and approved ones waiting for a cancellation decision). */
export function isUsing(r: ILeaveRequest): boolean {
  return r.kind === 'leave' && (r.status === 'Approved' || r.status === 'CancelRequested');
}

export function isPending(r: ILeaveRequest): boolean {
  return r.kind === 'leave' && r.status === 'Pending';
}

interface ICredit {
  left: number;
  effective: string;
  expiresOn?: string;
}

/**
 * Earned credits (comp off) used first-expiring-first. Each use takes from credits valid on its start day; what is
 * left of a credit after its expiry date has lapsed.
 */
export function earnedBalance(credits: ILedgerEntry[], uses: ILeaveRequest[], asOf: string): { earned: number; expired: number; used: number; pending: number; available: number } {
  const pool: ICredit[] = credits
    .filter((c) => c.days > 0)
    .map((c) => ({ left: c.days, effective: c.effective, expiresOn: c.expiresOn }))
    .sort((a, b) => (a.expiresOn || '9999') < (b.expiresOn || '9999') ? -1 : (a.expiresOn || '9999') > (b.expiresOn || '9999') ? 1 : 0);
  const manual = credits.filter((c) => c.days < 0).reduce((s, c) => s + c.days, 0);
  let used = 0;
  let pending = 0;
  let unmet = 0;
  uses
    .slice()
    .sort((a, b) => (a.start < b.start ? -1 : a.start > b.start ? 1 : 0))
    .forEach((u) => {
      if (isUsing(u)) used += u.days;
      else if (isPending(u)) pending += u.days;
      else return;
      let need = u.days;
      pool.forEach((c) => {
        if (need <= 0 || c.left <= 0) return;
        if (c.effective > u.start || (c.expiresOn && c.expiresOn < u.start)) return;
        const take = Math.min(c.left, need);
        c.left -= take;
        need -= take;
      });
      unmet += need;
    });
  let expired = 0;
  let left = 0;
  pool.forEach((c) => {
    if (c.expiresOn && c.expiresOn < asOf) expired += c.left;
    else left += c.left;
  });
  const earned = credits.filter((c) => c.days > 0).reduce((s, c) => s + c.days, 0);
  return { earned: round2(earned), expired: round2(expired), used: round2(used), pending: round2(pending), available: round2(left - unmet + manual) };
}

/** One person's balance of one leave type for a leave year. */
export function computeBalance(
  config: IConfig,
  type: ILeaveType,
  e: IEmployee | undefined,
  year: number,
  ledger: ILedgerEntry[],
  requests: ILeaveRequest[],
  asOf: string
): IBalance {
  const email = ((e && e.email) || '').toLowerCase();
  const mine = (x: { employeeEmail: string }): boolean => x.employeeEmail.toLowerCase() === email;
  const el = eligibility(type, e, asOf);
  const en = type.entitlement;
  const base: IBalance = {
    leaveType: type.key,
    year,
    unlimited: en.mode === 'unlimited',
    annual: 0,
    entitled: 0,
    tenureExtra: 0,
    carried: 0,
    adjusted: 0,
    earned: 0,
    expired: 0,
    used: 0,
    pending: 0,
    available: 0,
    eligible: el.ok,
    why: el.why
  };
  const myRequests = requests.filter((r) => mine(r) && r.leaveType === type.key && r.kind === 'leave');
  if (en.mode === 'earned') {
    const credits = ledger.filter((l) => mine(l) && l.leaveType === type.key);
    const eb = earnedBalance(credits, myRequests, asOf);
    return { ...base, earned: eb.earned, expired: eb.expired, used: eb.used, pending: eb.pending, available: eb.available };
  }
  const yearRequests = myRequests.filter((r) => r.leaveYear === year);
  const used = round2(yearRequests.filter(isUsing).reduce((s, r) => s + r.days, 0));
  const pending = round2(yearRequests.filter(isPending).reduce((s, r) => s + r.days, 0));
  if (en.mode === 'unlimited') return { ...base, used, pending };
  const p = policyEntitlement(type, e, year, config.yearStartMonth, asOf);
  const entries = ledger.filter((l) => mine(l) && l.leaveType === type.key && l.leaveYear === year);
  const carried = round2(entries.filter((l) => l.kind === 'carry').reduce((s, l) => s + l.days, 0));
  const adjusted = round2(entries.filter((l) => l.kind !== 'carry').reduce((s, l) => s + l.days, 0));
  let pool = p.entitled + carried + adjusted;
  if (en.maxBalance > 0) pool = Math.min(pool, en.maxBalance);
  return { ...base, annual: p.annual, entitled: p.entitled, tenureExtra: p.tenureExtra, carried, adjusted, used, pending, available: round2(pool - used - pending) };
}

/** Days to carry from a finished year into the next: what was left, up to the type's limit. */
export function carryForwardDays(type: ILeaveType, closing: IBalance): number {
  const max = type.entitlement.carryForwardMax;
  if (type.entitlement.mode === 'unlimited' || type.entitlement.mode === 'earned' || max === 0) return 0;
  // Pending requests of the finished year still count as taken.
  const left = Math.max(0, closing.available);
  return round2(max < 0 ? left : Math.min(left, max));
}

/** Balances of every active type for one person. */
export function allBalances(config: IConfig, e: IEmployee | undefined, year: number, ledger: ILedgerEntry[], requests: ILeaveRequest[], asOf: string): IBalance[] {
  return config.leaveTypes.filter((t) => t.active).map((t) => computeBalance(config, t, e, year, ledger, requests, asOf));
}

/**
 * Bradford factor of absences in a window: S² × D, where S = number of separate spells and D = total days.
 * Many short absences score higher than one long one. Spells on consecutive working days are one spell
 * (requests that touch or overlap are merged).
 */
export function bradfordFactor(requests: Array<{ start: string; end: string; days: number }>): { spells: number; days: number; score: number } {
  const sorted = requests.slice().sort((a, b) => (a.start < b.start ? -1 : 1));
  let spells = 0;
  let lastEnd = '';
  let days = 0;
  sorted.forEach((r) => {
    days += r.days;
    if (!lastEnd || r.start > addDays(lastEnd, 3)) spells++;
    if (r.end > lastEnd) lastEnd = r.end;
  });
  return { spells, days: round2(days), score: Math.round(spells * spells * days) };
}
