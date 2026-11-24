// Regression tests for bugs found in review and end-to-end testing.
const test = require('node:test');
const assert = require('node:assert');
const { lib } = require('./helpers.cjs');
const { defaultConfig } = lib('schema/defaults');
const wf = lib('services/workflow');
const bal = lib('services/balances');

const cfg = { ...defaultConfig('Acme'), hr: [{ email: 'hr@x.com', name: 'HR' }] };
const type = (k) => cfg.leaveTypes.find((t) => t.key === k);
const people = [
  { email: 'eli@x.com', name: 'Eli', active: true, managerEmail: 'maya@x.com', department: 'Eng' },
  { email: 'maya@x.com', name: 'Maya', active: true, managerEmail: 'dan@x.com', department: 'Eng' },
  { email: 'dan@x.com', name: 'Dan', active: true, department: 'Eng' }
];
const dir = wf.indexEmployees(people);

test('a requester is never an approver of their own request, even as a delegate', () => {
  const deleg = [{ active: true, delegator: { email: 'maya@x.com', name: 'Maya' }, delegate: { email: 'eli@x.com', name: 'Eli' }, from: '2026-01-01', to: '2026-12-31' }];
  const steps = wf.buildSteps(cfg, dir, deleg, type('annual'), people[0], 1, '2026-06-01');
  assert.deepStrictEqual(steps[0].approvers.map((a) => a.email), ['maya@x.com']);
  const r = { status: 'Pending', currentStep: 0, steps: [{ ...steps[0], status: 'pending' }], employeeEmail: 'eli@x.com' };
  assert.deepStrictEqual(wf.refreshDelegates(r, deleg, '2026-06-01'), undefined);
  assert.strictEqual(wf.waitingFor([r], 'eli@x.com').length, 0);
});

test('admins do not approve their own leave', () => {
  const r = { status: 'Pending', currentStep: 0, steps: [{ status: 'pending', approvers: [{ email: 'maya@x.com' }] }], employeeEmail: 'boss@x.com', history: [] };
  assert.strictEqual(wf.canDecide(r, { email: 'boss@x.com', name: 'Boss', isHr: true, isAdmin: true }, cfg).ok, false);
});

test('declining the cancellation of a pending request puts it back to Pending, not Approved', () => {
  const r = { id: 1, status: 'CancelRequested', currentStep: 0, steps: [{ status: 'pending', label: 'Manager', approvers: [{ email: 'maya@x.com' }] }], employeeEmail: 'eli@x.com', history: [] };
  const d = wf.decide(r, 'reject', { email: 'maya@x.com', name: 'Maya', isHr: false, isAdmin: false }, '', false, '2026-06-01T00:00:00Z');
  assert.strictEqual(d.status, 'Pending');
  const approved = { ...r, steps: [{ ...r.steps[0], status: 'approved' }], currentStep: 1 };
  assert.strictEqual(wf.decide(approved, 'reject', { email: 'maya@x.com', name: 'Maya', isHr: false, isAdmin: false }, '', false, 'x').status, 'Approved');
});

test('a pending request waiting for a cancellation decision counts as pending, an approved one as taken', () => {
  const pendingCancel = { kind: 'leave', status: 'CancelRequested', steps: [{ status: 'pending' }], days: 2 };
  const approvedCancel = { kind: 'leave', status: 'CancelRequested', steps: [{ status: 'approved' }], days: 2 };
  assert.strictEqual(bal.isUsing(pendingCancel), false);
  assert.strictEqual(bal.isPending(pendingCancel), true);
  assert.strictEqual(bal.isUsing(approvedCancel), true);
  assert.strictEqual(bal.isPending(approvedCancel), false);
});

test('the request form opens on the next working day (never a weekend or holiday)', () => {
  const d = lib('services/dates');
  const cal = d.calendarFor(cfg, undefined);
  assert.strictEqual(d.nextWorkingDay(cal, '2026-11-28'), '2026-11-30'); // Saturday → Monday
  assert.strictEqual(d.nextWorkingDay(cal, '2026-12-25'), '2026-12-28'); // Christmas (Fri) → Monday
  assert.strictEqual(d.nextWorkingDay(cal, '2026-12-01'), '2026-12-01');
});
