// Sample organisation for demos and the local preview: twelve people in three teams with grades, start dates and
// calendars, and requests in every state. Sample people have @example.com addresses.

import { IConfig, IEmployee, ILeaveRequest, ILedgerEntry } from '../models/types';
import { addDays, leaveYearOf, startOfWeek, today, weekday } from './dates';
import { buildSteps, decide, indexEmployees, startWorkflow } from './workflow';
import { duration } from './requests';

export const SAMPLE_DOMAIN = 'example.com';

export function sampleEmployees(me: { email: string; name: string }): IEmployee[] {
  const e = (name: string, mgr: string, dept: string, title: string, grade: string, hire: string, extra: Partial<IEmployee> = {}): IEmployee => ({
    email: name.toLowerCase().replace(/[^a-z]+/g, '.') + '@' + SAMPLE_DOMAIN,
    name,
    managerEmail: mgr ? mgr.toLowerCase().replace(/[^a-z]+/g, '.') + '@' + SAMPLE_DOMAIN : '',
    department: dept,
    jobTitle: title,
    grade,
    hireDate: hire,
    active: true,
    ...extra
  });
  const people = [
    e('Elena Fischer', '', 'Leadership', 'Chief Executive Officer', 'L5', '2014-02-03', { gender: 'Female' }),
    e('Rahul Mehta', 'Elena Fischer', 'Engineering', 'Head of Engineering', 'L5', '2016-06-13', { gender: 'Male' }),
    e('Sofia Alvarez', 'Rahul Mehta', 'Engineering', 'Senior Developer', 'L3', '2019-09-02', { gender: 'Female' }),
    e('Kenji Watanabe', 'Rahul Mehta', 'Engineering', 'Developer', 'L2', '2023-01-16', { gender: 'Male' }),
    e('Amara Okafor', 'Rahul Mehta', 'Engineering', 'QA Engineer', 'L2', '2024-03-04', { gender: 'Female' }),
    e('Lucas Martin', 'Elena Fischer', 'Sales', 'Sales Director', 'L4', '2017-11-20', { gender: 'Male', calendar: 'us' }),
    e('Hannah Becker', 'Lucas Martin', 'Sales', 'Account Executive', 'L2', '2022-05-09', { gender: 'Female', calendar: 'us' }),
    e('Diego Rossi', 'Lucas Martin', 'Sales', 'Sales Associate', 'L1', '2025-08-18', { gender: 'Male', calendar: 'us' }),
    e('Grace Kim', 'Elena Fischer', 'Operations', 'Operations Manager', 'L4', '2018-04-23', { gender: 'Female', calendar: 'six-day' }),
    e('Omar Haddad', 'Grace Kim', 'Operations', 'Support Lead', 'L3', '2020-10-05', { gender: 'Male', calendar: 'six-day' }),
    e('Mei Lin', 'Grace Kim', 'Operations', 'Support Specialist', 'L1', '2026-02-02', { gender: 'Female', calendar: 'six-day' })
  ];
  // The signed-in person joins Engineering under Rahul, with a report of their own, so every role can be tried.
  people.push({ email: me.email, name: me.name, managerEmail: 'rahul.mehta@' + SAMPLE_DOMAIN, department: 'Engineering', jobTitle: 'Engineering Manager', grade: 'L4', hireDate: '2021-07-01', active: true, gender: 'Female' });
  people.filter((p) => p.name === 'Kenji Watanabe' || p.name === 'Amara Okafor').forEach((p) => (p.managerEmail = me.email));
  return people;
}

/** Next weekday on or after a day. */
function workday(iso: string): string {
  let d = iso;
  while (weekday(d) === 0 || weekday(d) === 6) d = addDays(d, 1);
  return d;
}

export function sampleRequests(config: IConfig, people: IEmployee[], me: { email: string; name: string }): Array<Omit<ILeaveRequest, 'id' | 'created'>> {
  const dir = indexEmployees(people);
  const t = today();
  const monday = startOfWeek(t, 1);
  const now = new Date().toISOString();
  const hoursAgo = (h: number): string => new Date(Date.now() - h * 3600000).toISOString();
  const byName = (n: string): IEmployee => people.filter((p) => p.name === n)[0] || people.filter((p) => p.email === me.email)[0];
  const mk = (who: string, type: string, start: string, end: string, opts: Partial<ILeaveRequest> & { decide?: Array<'approve' | 'reject'>; waited?: number } = {}): Omit<ILeaveRequest, 'id' | 'created'> => {
    const e = byName(who);
    const ty = config.leaveTypes.filter((x) => x.key === type)[0];
    const unit = opts.unit || 'day';
    const dur = duration(config, ty, e, { start, end, unit, hours: opts.hours, kind: opts.kind || 'leave' });
    const steps = buildSteps(config, dir, [], ty, e, dur.days, t);
    const wf = opts.kind === 'compoff' ? { ...ty, autoApprove: { enabled: false, maxDays: 0 } } : ty;
    const s = startWorkflow(wf, steps, dur.days, opts.waited ? hoursAgo(opts.waited) : now, e.name);
    let r: Omit<ILeaveRequest, 'id' | 'created'> = {
      kind: opts.kind || 'leave',
      employeeEmail: e.email,
      employeeName: e.name,
      leaveType: type,
      start,
      end,
      unit,
      halfDayPart: opts.halfDayPart,
      startTime: opts.startTime,
      hours: dur.hours,
      days: dur.days,
      leaveYear: leaveYearOf(start, config.yearStartMonth),
      status: s.status,
      currentStep: s.currentStep,
      steps: s.steps,
      reason: opts.reason || '',
      history: s.history,
      seriesId: opts.seriesId,
      seriesRule: opts.seriesRule
    };
    (opts.decide || []).forEach((d) => {
      const step = r.steps[r.currentStep];
      if (!step || r.status !== 'Pending') return;
      const a = step.approvers[0];
      const res = decide({ ...r, id: 0, created: now }, d, { email: a.email, name: a.name, isHr: false, isAdmin: false }, d === 'reject' ? 'Release week — can you move it?' : '', false, now);
      r = { ...r, status: res.status, steps: res.steps, currentStep: res.currentStep, history: res.history };
    });
    return r;
  };
  const out = [
    mk('Sofia Alvarez', 'annual', workday(addDays(monday, 14)), addDays(workday(addDays(monday, 14)), 4), { decide: ['approve'], reason: 'Family trip' }),
    mk('Kenji Watanabe', 'annual', workday(addDays(monday, 21)), addDays(workday(addDays(monday, 21)), 2), { reason: 'Moving house', waited: 30 }),
    mk('Amara Okafor', 'sick', workday(addDays(monday, -3)), workday(addDays(monday, -3)), { decide: ['approve'], reason: 'Flu' }),
    mk('Amara Okafor', 'medical', workday(addDays(t, 2)), workday(addDays(t, 2)), { unit: 'hours', hours: 2, startTime: '10:00', reason: 'Dentist' }),
    mk('Kenji Watanabe', 'compoff', addDays(monday, -2), addDays(monday, -2), { kind: 'compoff', reason: 'Weekend release support' }),
    mk('Hannah Becker', 'casual', workday(t), workday(t), { unit: 'half', halfDayPart: 'pm', reason: 'School event' }),
    mk('Diego Rossi', 'annual', workday(addDays(monday, 7)), addDays(workday(addDays(monday, 7)), 1), { decide: ['reject'], reason: 'Concert' }),
    mk('Omar Haddad', 'annual', workday(addDays(monday, 1)), addDays(workday(addDays(monday, 1)), 2), { decide: ['approve'], reason: 'Wedding anniversary' }),
    mk('Grace Kim', 'earned', workday(addDays(monday, 35)), addDays(workday(addDays(monday, 35)), 9), { decide: ['approve'], reason: 'Summer holiday' }),
    mk(me.name, 'annual', workday(addDays(monday, 28)), addDays(workday(addDays(monday, 28)), 4), { reason: 'Holiday in Lisbon', waited: 50 }),
    mk('Rahul Mehta', 'wfh', workday(addDays(monday, 3)), workday(addDays(monday, 3)), { reason: 'Work from home' })
  ];
  // A bi-weekly series: WFH every other Friday.
  const fri = addDays(monday, 4);
  for (let i = 0; i < 4; i++) {
    const d = addDays(fri, i * 14);
    out.push(mk('Sofia Alvarez', 'wfh', d, d, { reason: 'Focus day', seriesId: 'Ssample', seriesRule: 'Every other week on Fri, 4 times' }));
  }
  return out;
}

export function sampleLedger(config: IConfig, people: IEmployee[]): ILedgerEntry[] {
  const y = leaveYearOf(today(), config.yearStartMonth);
  const sofia = people.filter((p) => p.name === 'Sofia Alvarez')[0];
  const omar = people.filter((p) => p.name === 'Omar Haddad')[0];
  return [
    { employeeEmail: sofia.email, leaveType: 'annual', leaveYear: y, days: 3, kind: 'carry', note: `Carried forward from ${y - 1}`, effective: `${y}-01-01`, uniqueKey: `carry:${y - 1}:${sofia.email}:annual` },
    { employeeEmail: omar.email, leaveType: 'compoff', leaveYear: y, days: 1, kind: 'compoff', note: 'Comp off for weekend migration', effective: addDays(today(), -20), expiresOn: addDays(today(), 70) }
  ];
}
