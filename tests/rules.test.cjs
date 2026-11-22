const test = require('node:test');
const assert = require('node:assert');
const { lib } = require('./helpers.cjs');
const { defaultConfig } = lib('schema/defaults');
const dates = lib('services/dates');
const bal = lib('services/balances');
const wf = lib('services/workflow');
const req = lib('services/requests');
const lic = lib('services/licensing/license');

const cfg = defaultConfig('Acme');
const type = (k) => cfg.leaveTypes.find((t) => t.key === k);
const emp = (o = {}) => ({ email: 'a@x.com', name: 'A', active: true, managerEmail: 'm@x.com', grade: 'L2', hireDate: '2020-01-01', department: 'Eng', ...o });
const mgr = { email: 'm@x.com', name: 'M', active: true, managerEmail: 'b@x.com', department: 'Eng' };
const boss = { email: 'b@x.com', name: 'B', active: true, department: 'Lead' };

test('30 leave types ship by default', () => {
  assert.strictEqual(cfg.leaveTypes.length, 30);
  assert.strictEqual(new Set(cfg.leaveTypes.map((t) => t.key)).size, 30);
});

test('weekends and holidays are excluded from the day count', () => {
  const cal = dates.calendarFor(cfg, emp());
  // Thu 31 Dec 2026 – Mon 4 Jan 2027: Thu, (Fri 1 Jan holiday), Sat, Sun, Mon
  const c = dates.countDays(cal, '2026-12-31', '2027-01-04', false);
  assert.strictEqual(c.days, 2);
  assert.strictEqual(c.excluded.length, 3);
  assert.strictEqual(dates.countDays(cal, '2026-12-31', '2027-01-04', true).days, 5);
});

test('half-day and two-hour leave', () => {
  const t = type('casual');
  assert.strictEqual(req.duration(cfg, t, emp(), { kind: 'leave', start: '2026-11-24', end: '2026-11-24', unit: 'half' }).days, 0.5);
  const h = req.duration(cfg, t, emp(), { kind: 'leave', start: '2026-11-24', end: '2026-11-24', unit: 'hours', hours: 2 });
  assert.strictEqual(h.hours, 2);
  assert.strictEqual(h.days, 0.25);
});

test('grade and tenure based entitlement', () => {
  const t = { ...type('casual'), entitlement: { ...type('casual').entitlement, days: 10, byGrade: { L3: 14 }, tenure: [{ years: 5, extraDays: 3 }], prorate: false } };
  assert.strictEqual(bal.policyEntitlement(t, emp({ grade: 'L1', hireDate: '2024-01-01' }), 2026, 1, '2026-06-01').entitled, 10);
  assert.strictEqual(bal.policyEntitlement(t, emp({ grade: 'L3', hireDate: '2024-01-01' }), 2026, 1, '2026-06-01').entitled, 14);
  assert.strictEqual(bal.policyEntitlement(t, emp({ grade: 'L3', hireDate: '2019-01-01' }), 2026, 1, '2026-06-01').entitled, 17);
});

test('fixed entitlement is prorated for new joiners', () => {
  const t = { ...type('casual'), entitlement: { ...type('casual').entitlement, days: 12, prorate: true } };
  assert.strictEqual(bal.policyEntitlement(t, emp({ hireDate: '2026-07-01' }), 2026, 1, '2026-08-01').entitled, 6);
});

test('monthly accrual credits per completed month', () => {
  const t = { ...type('annual'), entitlement: { ...type('annual').entitlement, days: 24, byGrade: {}, tenure: [], accrual: { frequency: 'monthly', timing: 'end' } } };
  const p = bal.policyEntitlement(t, emp(), 2026, 1, '2026-03-31');
  assert.strictEqual(p.entitled, 6);
  assert.strictEqual(p.annual, 24);
  assert.strictEqual(bal.accrualPeriods(2026, 1, 'biweekly').length, 26);
});

test('balance = entitlement + ledger − taken − pending; carry forward respects the cap', () => {
  const t = { ...type('casual'), entitlement: { ...type('casual').entitlement, days: 8, carryForwardMax: 3, prorate: false } };
  const c = { ...cfg, leaveTypes: [t] };
  const r = (status, days) => ({ kind: 'leave', employeeEmail: 'a@x.com', leaveType: 'casual', leaveYear: 2026, status, days, start: '2026-05-01', end: '2026-05-01' });
  const ledger = [{ employeeEmail: 'a@x.com', leaveType: 'casual', leaveYear: 2026, days: 2, kind: 'adjustment', effective: '2026-01-02' }];
  const b = bal.computeBalance(c, t, emp(), 2026, ledger, [r('Approved', 2), r('Pending', 1), r('Rejected', 4)], '2026-06-01');
  assert.strictEqual(b.available, 7);
  assert.strictEqual(bal.carryForwardDays(t, b), 3);
});

test('comp off credits are used first-expiring-first and lapse', () => {
  const credits = [
    { days: 1, effective: '2026-01-10', expiresOn: '2026-04-10' },
    { days: 1, effective: '2026-02-10', expiresOn: '2026-05-10' }
  ];
  const use = [{ kind: 'leave', status: 'Approved', days: 1, start: '2026-03-01' }];
  const e = bal.earnedBalance(credits, use, '2026-06-01');
  assert.strictEqual(e.used, 1);
  assert.strictEqual(e.expired, 1);
  assert.strictEqual(e.available, 0);
});

test('multilevel workflow: manager, then HR from 5 days; delegate added', () => {
  const c = { ...cfg, hr: [{ email: 'hr@x.com', name: 'HR' }] };
  const dir = wf.indexEmployees([emp(), mgr, boss]);
  const t = { ...type('annual'), workflowKey: 'two-level' };
  assert.strictEqual(wf.buildSteps(c, dir, [], t, emp(), 2, '2026-06-01').length, 1);
  const deleg = [{ active: true, delegator: { email: 'm@x.com', name: 'M' }, delegate: { email: 'd@x.com', name: 'D' }, from: '2026-05-01', to: '2026-06-30' }];
  const steps = wf.buildSteps(c, dir, deleg, t, emp(), 5, '2026-06-01');
  assert.deepStrictEqual(steps.map((s) => s.kind), ['manager', 'hr']);
  assert.deepStrictEqual(steps[0].approvers.map((a) => a.email), ['m@x.com', 'd@x.com']);
});

test('approve moves to the next level, reject ends it, on-behalf is recorded', () => {
  const c = { ...cfg, hr: [{ email: 'hr@x.com', name: 'HR' }] };
  const dir = wf.indexEmployees([emp(), mgr, boss]);
  const t = { ...type('annual'), workflowKey: 'two-level', autoApprove: { enabled: false, maxDays: 0 } };
  const steps = wf.buildSteps(c, dir, [], t, emp(), 6, '2026-06-01');
  const s = wf.startWorkflow(t, steps, 6, '2026-06-01T00:00:00Z', 'A');
  let r = { id: 1, employeeEmail: 'a@x.com', status: s.status, steps: s.steps, currentStep: s.currentStep, history: s.history };
  assert.strictEqual(r.status, 'Pending');
  const m = { email: 'm@x.com', name: 'M', isHr: false, isAdmin: false };
  assert.ok(wf.canDecide(r, m, c).ok);
  let d = wf.decide(r, 'approve', m, '', false, '2026-06-02T00:00:00Z');
  assert.strictEqual(d.status, 'Pending');
  assert.strictEqual(d.currentStep, 1);
  r = { ...r, ...d };
  const admin = { email: 'z@x.com', name: 'Z', isHr: true, isAdmin: true };
  assert.deepStrictEqual(wf.canDecide(r, admin, c), { ok: true, onBehalf: true });
  d = wf.decide(r, 'reject', admin, 'no', true, '2026-06-03T00:00:00Z');
  assert.strictEqual(d.status, 'Rejected');
  assert.ok(d.steps[1].onBehalf);
});

test('auto approval within the type limit', () => {
  const t = type('casual'); // auto-approves up to 1 day
  assert.strictEqual(wf.startWorkflow(t, [{ approvers: [], status: 'waiting' }], 1, 'x', 'A').status, 'Approved');
  assert.strictEqual(wf.startWorkflow(t, [{ approvers: [], status: 'waiting' }], 2, 'x', 'A').status, 'Pending');
});

test('reminders, escalation and timeout auto-approval', () => {
  const c = { ...cfg, hr: [{ email: 'hr@x.com', name: 'HR' }] };
  const now = Date.parse('2026-06-10T00:00:00Z');
  const r = (h, extra = {}) => ({ status: 'Pending', currentStep: 0, steps: [{ kind: 'manager', approvers: [{ email: 'm@x.com' }], startedAt: new Date(now - h * 3600000).toISOString(), ...extra }] });
  assert.strictEqual(wf.reminderAction(r(2), c, now), 'none');
  assert.strictEqual(wf.reminderAction(r(30), c, now), 'remind');
  assert.strictEqual(wf.reminderAction(r(30, { remindedAt: new Date(now - 3600000).toISOString(), reminders: 1 }), c, now), 'none');
  assert.strictEqual(wf.reminderAction(r(80), c, now), 'escalate');
  assert.strictEqual(wf.reminderAction(r(80), { ...c, autoApproveAfterHours: 48 }, now), 'autoApprove');
});

test('weekly, bi-weekly and monthly recurrence', () => {
  assert.deepStrictEqual(req.occurrences('2026-11-27', { freq: 'biweekly', weekdays: [5], count: 3 }), ['2026-11-27', '2026-12-11', '2026-12-25']);
  assert.deepStrictEqual(req.occurrences('2026-11-23', { freq: 'weekly', weekdays: [1, 3], count: 3 }), ['2026-11-23', '2026-11-25', '2026-11-30']);
  assert.deepStrictEqual(req.occurrences('2026-01-31', { freq: 'monthly', weekdays: [], count: 2 }), ['2026-01-31', '2026-02-28']);
});

test('validation: balance, notice, overlap, comp off on a working day', () => {
  const t = type('earned'); // 7 days notice
  const e = emp();
  const b = { available: 2, unlimited: false, eligible: true };
  const draft = { kind: 'leave', employeeEmail: e.email, leaveType: 'earned', start: '2026-11-25', end: '2026-11-27', unit: 'day', reason: '' };
  const dur = req.duration(cfg, t, e, draft);
  const v = req.validateDraft(cfg, t, e, draft, dur, b, [], '2026-11-22', false);
  assert.ok(v.errors.some((x) => /notice/.test(x)));
  assert.ok(v.errors.some((x) => /Not enough/.test(x)));
  const clash = req.overlapping([{ kind: 'leave', employeeEmail: e.email, status: 'Approved', start: '2026-11-26', end: '2026-11-26', unit: 'day' }], draft);
  assert.strictEqual(clash.length, 1);
  const co = { ...draft, kind: 'compoff', start: '2026-11-20', end: '2026-11-20', reason: 'release' };
  assert.ok(req.validateDraft(cfg, type('compoff'), e, co, req.duration(cfg, type('compoff'), e, co), undefined, [], '2026-11-22', false).errors.some((x) => /working day/.test(x)));
});

test('licence pricing quote and user limit', () => {
  assert.deepStrictEqual(lic.quote(120), { extra: 70, packs: 2, usd: 2, inr: 7000 });
  const set = lic.licensedEmails([{ email: 'A@x', id: 2 }, { email: 'b@x', id: 1 }, { email: 'c@x', id: 3 }], 2);
  assert.ok(set.has('a@x') && set.has('b@x') && !set.has('c@x'));
  const s = lic.licenseStatus(undefined, new Date('2026-11-01'), new Date('2026-11-10'));
  assert.strictEqual(s.state, 'trial');
  assert.strictEqual(lic.licenseStatus(undefined, new Date('2026-10-01'), new Date('2026-11-10')).active, false);
});
