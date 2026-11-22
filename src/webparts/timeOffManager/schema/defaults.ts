import { IConfig, IHoliday, IWorkCalendar, IWorkflow } from '../models/types';
import { defaultLeaveTypes } from './leaveTypes';

export const DEFAULT_BRAND = '#038f62';
export const DEFAULT_ACCENT = '#caf578';

function fixedHolidays(years: number[], list: Array<[string, string]>): IHoliday[] {
  const out: IHoliday[] = [];
  years.forEach((y) => list.forEach(([md, name]) => out.push({ date: `${y}-${md}`, name })));
  return out;
}

export function defaultCalendars(): IWorkCalendar[] {
  const years = [2026, 2027];
  return [
    {
      key: 'default',
      name: 'Head office (Mon–Fri)',
      workDays: [1, 2, 3, 4, 5],
      hoursPerDay: 8,
      dayStart: '09:00',
      holidays: fixedHolidays(years, [
        ['01-01', "New Year's Day"],
        ['01-26', 'Republic Day'],
        ['05-01', 'Labour Day'],
        ['08-15', 'Independence Day'],
        ['10-02', 'Gandhi Jayanti'],
        ['12-25', 'Christmas Day']
      ]).concat([
        { date: '2026-03-04', name: 'Holi' },
        { date: '2026-11-09', name: 'Diwali' },
        { date: '2027-03-22', name: 'Holi' },
        { date: '2027-10-29', name: 'Diwali' }
      ])
    },
    {
      key: 'six-day',
      name: 'Operations (Mon–Sat)',
      workDays: [1, 2, 3, 4, 5, 6],
      hoursPerDay: 8,
      dayStart: '08:00',
      holidays: fixedHolidays(years, [
        ['01-01', "New Year's Day"],
        ['08-15', 'Independence Day'],
        ['12-25', 'Christmas Day']
      ])
    },
    {
      key: 'us',
      name: 'United States office',
      workDays: [1, 2, 3, 4, 5],
      hoursPerDay: 8,
      dayStart: '09:00',
      holidays: fixedHolidays(years, [
        ['01-01', "New Year's Day"],
        ['07-04', 'Independence Day'],
        ['12-25', 'Christmas Day']
      ]).concat([
        { date: '2026-05-25', name: 'Memorial Day' },
        { date: '2026-09-07', name: 'Labor Day' },
        { date: '2026-11-26', name: 'Thanksgiving' },
        { date: '2027-05-31', name: 'Memorial Day' },
        { date: '2027-09-06', name: 'Labor Day' },
        { date: '2027-11-25', name: 'Thanksgiving' }
      ])
    }
  ];
}

export function defaultWorkflows(): IWorkflow[] {
  return [
    { key: 'standard', name: 'Manager approval', levels: [{ kind: 'manager', label: 'Manager', minDays: 0 }] },
    {
      key: 'two-level',
      name: 'Manager, then HR for 5+ days',
      levels: [
        { kind: 'manager', label: 'Manager', minDays: 0 },
        { kind: 'hr', label: 'HR', minDays: 5 }
      ]
    },
    {
      key: 'extended',
      name: 'Manager, department head, HR',
      levels: [
        { kind: 'manager', label: 'Manager', minDays: 0 },
        { kind: 'departmentHead', label: 'Department head', minDays: 0 },
        { kind: 'hr', label: 'HR', minDays: 0 }
      ]
    },
    { key: 'hr-only', name: 'HR only', levels: [{ kind: 'hr', label: 'HR', minDays: 0 }] }
  ];
}

export function defaultConfig(orgName: string = ''): IConfig {
  return {
    appName: 'Time Off Manager',
    orgName,
    brandColor: DEFAULT_BRAND,
    accentColor: DEFAULT_ACCENT,
    headline: 'Plan time off, together.',
    yearStartMonth: 1,
    weekStartsOn: 1,
    leaveTypes: defaultLeaveTypes(),
    workflows: defaultWorkflows(),
    calendars: defaultCalendars(),
    defaultCalendar: 'default',
    templates: [
      { key: 'fri-half', name: 'Friday afternoon off', leaveType: 'casual', unit: 'half', halfDayPart: 'pm', reason: 'Personal time' },
      { key: 'doctor', name: 'Doctor appointment (2 h)', leaveType: 'medical', unit: 'hours', hours: 2, startTime: '10:00', reason: 'Medical appointment' },
      { key: 'long-weekend', name: 'Long weekend (Fri + Mon)', leaveType: 'annual', unit: 'day', days: 4, reason: 'Long weekend' },
      { key: 'wfh-biweekly', name: 'WFH every other Friday', leaveType: 'wfh', unit: 'day', reason: 'Work from home', recurrence: { freq: 'biweekly', weekdays: [5], count: 6 } }
    ],
    grades: ['L1', 'L2', 'L3', 'L4', 'L5'],
    departmentHeads: {},
    admins: [],
    hr: [],
    hourOptions: [1, 2, 3, 4],
    onBehalf: { managers: true, hr: true },
    cancellation: { pendingDirect: true, approvedNeedsApproval: true, allowPast: false },
    reminders: { enabled: true, afterHours: 24, repeatHours: 24, maxReminders: 3, escalateAfterHours: 72 },
    autoApproveAfterHours: 0,
    notifications: { email: true, inApp: true, notifyHr: false, ccEmails: '' },
    integrations: { webhookUrl: '', outbox: true },
    calendarSync: { outlook: true, showAs: 'oof' },
    teamCalendar: { scope: 'team', showLeaveType: true },
    compOff: { enabled: true, leaveType: 'compoff', requireNonWorkingDay: true, maxDaysPerClaim: 2 },
    blackouts: [],
    staffing: { enabled: false, maxAwayPercent: 50, maxAwayCount: 0, mode: 'warn', perDepartment: {} },
    groupCalendar: { enabled: false, groupId: '', groupName: '' },
    autoReply: {
      enabled: true,
      message: 'Hi, I am on leave from {start} to {end} and will reply when I am back. For anything urgent, please contact {manager}.'
    }
  };
}

function isObj(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

/** Saved settings over the defaults: new settings added by an update get their default, arrays are taken as saved. */
export function mergeConfig(saved: unknown, orgName: string = ''): IConfig {
  const base = defaultConfig(orgName) as unknown as Record<string, unknown>;
  const deep = (into: Record<string, unknown>, from: Record<string, unknown>): Record<string, unknown> => {
    const out: Record<string, unknown> = { ...into };
    Object.keys(from).forEach((k) => {
      const v = from[k];
      if (v === undefined || v === null) return;
      out[k] = isObj(v) && isObj(into[k]) && k !== 'departmentHeads' && k !== 'perDepartment' ? deep(into[k] as Record<string, unknown>, v) : v;
    });
    return out;
  };
  const merged = (isObj(saved) ? deep(base, saved) : base) as unknown as IConfig;
  // Leave types saved by an earlier version get any policy field added since.
  const template = defaultLeaveTypes()[1];
  merged.leaveTypes = (merged.leaveTypes || []).map((t) => ({
    ...template,
    ...t,
    units: { ...template.units, ...(t.units || {}) },
    entitlement: { ...template.entitlement, ...(t.entitlement || {}), accrual: { ...template.entitlement.accrual, ...((t.entitlement && t.entitlement.accrual) || {}) } },
    eligibility: { ...template.eligibility, ...(t.eligibility || {}) },
    autoApprove: { ...template.autoApprove, ...(t.autoApprove || {}) }
  }));
  return merged;
}
