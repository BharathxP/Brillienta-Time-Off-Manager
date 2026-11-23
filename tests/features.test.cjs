const test = require('node:test');
const assert = require('node:assert');
const { lib } = require('./helpers.cjs');
const { defaultConfig } = lib('schema/defaults');
const bal = lib('services/balances');
const req = lib('services/requests');
const { xlsxWorkbook } = lib('services/xlsx');
const { flowPackageZip } = lib('services/flowPackage');

const cfg = defaultConfig('Acme');
const type = (k) => cfg.leaveTypes.find((t) => t.key === k);
const emp = (o = {}) => ({ email: 'a@x.com', name: 'A', active: true, department: 'Eng', grade: 'L2', hireDate: '2020-01-01', ...o });

test('sandwich rule counts the weekend between Friday and Monday', () => {
  const plain = type('casual');
  const sandwich = { ...plain, sandwich: true };
  const d = { kind: 'leave', start: '2026-11-27', end: '2026-11-30', unit: 'day' }; // Fri–Mon
  assert.strictEqual(req.duration(cfg, plain, emp(), d).days, 2);
  assert.strictEqual(req.duration(cfg, sandwich, emp(), d).days, 4);
});

test('partial first and last day take half a day off each end', () => {
  const t = type('annual');
  const d = { kind: 'leave', start: '2026-11-23', end: '2026-11-27', unit: 'day' }; // Mon–Fri
  assert.strictEqual(req.duration(cfg, t, emp(), { ...d, partial: 'start' }).days, 4.5);
  assert.strictEqual(req.duration(cfg, t, emp(), { ...d, partial: 'both' }).days, 4);
});

test('blackouts block by department and leave type; HR is only warned', () => {
  const c = { ...cfg, blackouts: [{ key: 'b', name: 'Year-end close', from: '2026-12-28', to: '2026-12-31', departments: ['Finance'], leaveTypes: [], mode: 'block' }] };
  const t = type('casual');
  const d = { kind: 'leave', employeeEmail: 'a@x.com', leaveType: 'casual', start: '2026-12-29', end: '2026-12-29', unit: 'day', reason: '' };
  const dur = req.duration(c, t, emp(), d);
  const b = { available: 5, unlimited: false, eligible: true };
  assert.strictEqual(req.validateDraft(c, t, emp(), d, dur, b, [], '2026-11-22', false).errors.length, 0);
  const fin = emp({ department: 'Finance' });
  assert.ok(req.validateDraft(c, t, fin, d, dur, b, [], '2026-11-22', false).errors.some((x) => /Year-end close/.test(x)));
  assert.ok(req.validateDraft(c, t, fin, d, dur, b, [], '2026-11-22', true).warnings.some((x) => /Year-end close/.test(x)));
});

test('minimum staffing: a third person away in a team of four goes over a 50% limit', () => {
  const c = { ...cfg, staffing: { enabled: true, maxAwayPercent: 50, maxAwayCount: 0, mode: 'block', perDepartment: {} } };
  const people = ['a', 'b', 'c', 'd'].map((n) => emp({ email: `${n}@x.com`, name: n.toUpperCase() }));
  const away = (n) => ({ id: n.charCodeAt(0), kind: 'leave', employeeEmail: `${n}@x.com`, employeeName: n, status: 'Approved', start: '2026-12-01', end: '2026-12-02', unit: 'day' });
  const clash = req.teamClash(c, people, [away('b'), away('c')], people[0], '2026-12-01', '2026-12-01');
  assert.strictEqual(clash.away.length, 2);
  assert.strictEqual(clash.limit, 2);
  assert.strictEqual(clash.peak.count, 3);
  const t = type('casual');
  const d = { kind: 'leave', employeeEmail: 'a@x.com', leaveType: 'casual', start: '2026-12-01', end: '2026-12-01', unit: 'day', reason: '' };
  const v = req.validateDraft(c, t, people[0], d, req.duration(c, t, people[0], d), { available: 5, unlimited: false, eligible: true }, [away('b'), away('c')], '2026-11-22', false, { employees: people });
  assert.ok(v.errors.some((x) => /Staffing/.test(x)));
});

test('leaving date prorates fixed leave and blocks later requests', () => {
  const t = { ...type('casual'), entitlement: { ...type('casual').entitlement, days: 12, prorate: true } };
  assert.strictEqual(bal.policyEntitlement(t, emp({ leavingDate: '2026-06-30' }), 2026, 1, '2026-03-01').entitled, 6);
  const leaver = emp({ leavingDate: '2026-12-15' });
  const d = { kind: 'leave', employeeEmail: 'a@x.com', leaveType: 'casual', start: '2026-12-16', end: '2026-12-16', unit: 'day', reason: '' };
  assert.ok(req.validateDraft(cfg, type('casual'), leaver, d, req.duration(cfg, type('casual'), leaver, d), { available: 5, unlimited: false, eligible: true }, [], '2026-11-22', false).errors.some((x) => /last working day/.test(x)));
});

test('Bradford factor: spells squared times days', () => {
  const r = (s, e, d) => ({ start: s, end: e, days: d });
  assert.deepStrictEqual(bal.bradfordFactor([r('2026-01-05', '2026-01-05', 1), r('2026-03-02', '2026-03-02', 1), r('2026-06-01', '2026-06-02', 2)]), { spells: 3, days: 4, score: 36 });
  assert.strictEqual(bal.bradfordFactor([r('2026-01-05', '2026-01-09', 5)]).score, 5);
});

test('Excel workbook and Power Automate flow package are valid zips', () => {
  const x = xlsxWorkbook([{ name: 'Leave', columns: ['A', 'B'], rows: [['x', 1]] }]);
  assert.strictEqual(String.fromCharCode(x[0], x[1]), 'PK');
  const z = flowPackageZip({ siteUrl: 'https://contoso.sharepoint.com/sites/hr', listId: '00000000-0000-0000-0000-000000000000', listTitle: 'BrillTM_EmailQueue', displayName: 'Send', sharedMailbox: '', allowedDomains: [], useReplyTo: false });
  assert.strictEqual(String.fromCharCode(z[0], z[1]), 'PK');
});

test('book ahead: accrual types may use what accrues later in the leave year', () => {
  const t = type('annual');
  const b = { available: 0, annual: 20, entitled: 0, unlimited: false, eligible: true };
  assert.deepStrictEqual(bal.bookable(t, b), { available: 20, ahead: 20 });
  const off = { ...t, entitlement: { ...t.entitlement, bookAhead: false } };
  assert.deepStrictEqual(bal.bookable(off, b), { available: 0, ahead: 0 });
  assert.strictEqual(bal.bookable(type('casual'), { ...b, available: 3 }).available, 3);
});

test('list prefix: BrillTM_ on new sites, earlier TO_ sites keep their lists, a custom prefix wins', async () => {
  const { resolveListPrefix, cleanPrefix } = lib('services/listPrefix');
  const site = (...lists) => async (p) => lists.indexOf(p + 'Settings') >= 0;
  assert.strictEqual(await resolveListPrefix('BrillTM_', site()), 'BrillTM_');
  assert.strictEqual(await resolveListPrefix('BrillTM_', site('TO_Settings')), 'TO_');
  assert.strictEqual(await resolveListPrefix('TO_', site('TO_Settings', 'BrillTM_Settings')), 'BrillTM_');
  assert.strictEqual(await resolveListPrefix(undefined, site('TO_Settings')), 'TO_');
  assert.strictEqual(await resolveListPrefix('Acme HR_', site('BrillTM_Settings')), 'AcmeHR_');
  assert.strictEqual(cleanPrefix("x'; drop"), 'xdrop');
});
