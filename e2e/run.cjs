// End-to-end tests: the real web part in Chromium, several people at once, against the strict SharePoint emulator.
//   npm run e2e            (builds lib/ and the e2e page first)
//   node e2e/run.cjs       (reuse the last build)
// Each scenario drives the UI like a person would and then checks what was written to the SharePoint lists.

const { chromium } = require('playwright');
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { start } = require('./spEmulator.cjs');

const PORT = 4399;
const BASE = `http://localhost:${PORT}`;
const USERS = [
  { Id: 1, Title: 'Alex Admin', Email: 'admin@contoso.com', IsSiteAdmin: true },
  { Id: 2, Title: 'Maya Manager', Email: 'maya.manager@contoso.com' },
  { Id: 3, Title: 'Dan Director', Email: 'dan.director@contoso.com' },
  { Id: 4, Title: 'Eli Employee', Email: 'eli.employee@contoso.com' },
  { Id: 5, Title: 'Ria Rao', Email: 'ria.rao@contoso.com' },
  { Id: 6, Title: 'Hana HR', Email: 'hana.hr@contoso.com' },
  { Id: 7, Title: 'Dave Delegate', Email: 'dave.delegate@contoso.com' }
];
const EMPLOYEES_CSV = [
  'Name,Email,Manager email,Department,Job title,Grade,Start date,Work calendar,Location,Gender,Leaving date,Active',
  'Alex Admin,admin@contoso.com,dan.director@contoso.com,People,HR Director,L5,2015-01-05,default,Pune,Male,,Yes',
  'Dan Director,dan.director@contoso.com,,Engineering,Engineering Director,L5,2016-02-01,default,Pune,Male,,Yes',
  'Maya Manager,maya.manager@contoso.com,dan.director@contoso.com,Engineering,Engineering Manager,L4,2018-04-02,default,Pune,Female,,Yes',
  'Eli Employee,eli.employee@contoso.com,maya.manager@contoso.com,Engineering,Developer,L2,2022-03-01,default,Pune,Male,,Yes',
  'Ria Rao,ria.rao@contoso.com,maya.manager@contoso.com,Engineering,Graduate Developer,L1,2026-06-01,default,Pune,Female,,Yes',
  'Hana HR,hana.hr@contoso.com,admin@contoso.com,People,HR Partner,L3,2019-09-09,default,Pune,Female,,Yes',
  'Dave Delegate,dave.delegate@contoso.com,dan.director@contoso.com,Engineering,Tech Lead,L3,2017-05-15,default,Pune,Male,,Yes'
].join('\n');

// ------------------------------------------------------------------ dates (local calendar, like the app)
const pad = (n) => String(n).padStart(2, '0');
const isoLocal = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const addDays = (iso, n) => {
  const [y, m, d] = iso.split('-').map(Number);
  return isoLocal(new Date(y, m - 1, d + n));
};
const weekday = (iso) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d).getDay();
};
const TODAY = isoLocal(new Date());
const nextWeekday = (from, wd) => {
  let d = addDays(from, 1);
  while (weekday(d) !== wd) d = addDays(d, 1);
  return d;
};
// Far enough ahead to clear notice periods, and away from holidays in the default calendar.
let MON = nextWeekday(addDays(TODAY, 14), 1);
const HOLIDAYS = ['01-01', '01-26', '05-01', '08-15', '10-02', '12-25', '03-04', '11-09', '03-22', '10-29'];
const clear = (start, n) => {
  for (let i = 0; i < n; i++) if (HOLIDAYS.indexOf(addDays(start, i).substring(5)) >= 0) return false;
  return true;
};
while (!clear(MON, 21)) MON = addDays(MON, 7);
const LAST_SAT = (() => {
  let d = addDays(TODAY, -1);
  while (weekday(d) !== 6) d = addDays(d, -1);
  return d;
})();

// ------------------------------------------------------------------ helpers
let browser;
let emu;
const results = [];
const findings = [];
const shots = path.join(__dirname, '..', 'e2e-results');
fs.mkdirSync(shots, { recursive: true });

async function state() {
  return emu.getState();
}
function items(st, list) {
  return st.lists['BrillTM_' + list] ? st.lists['BrillTM_' + list].items : [];
}
function config(st) {
  const row = items(st, 'Settings').filter((r) => r.Title === 'Config')[0];
  return row ? JSON.parse(row.SettingValue) : undefined;
}
async function seedLists(mutate) {
  const st = await state();
  mutate(st.lists);
}

const sessions = {};
async function as(email) {
  if (sessions[email]) return sessions[email];
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(String(e)));
  page.on('console', (m) => {
    if (m.type() === 'error' && !/ERR_TUNNEL|Failed to load resource/.test(m.text())) page.errors.push(m.text());
  });
  sessions[email] = page;
  return page;
}
async function open(email, hash, owner) {
  const page = await as(email);
  const user = USERS.filter((u) => u.Email === email)[0];
  await page.goto('about:blank');
  await page.goto(`${BASE}/?user=${encodeURIComponent(email)}${owner || user.IsSiteAdmin ? '&owner=1' : ''}#/${hash || 'home'}`);
  await page.waitForSelector('.to-app, .to-setup, .to-locked', { timeout: 20000 });
  await page.waitForTimeout(600);
  return page;
}
async function go(page, hash) {
  await page.evaluate((h) => (window.location.hash = '#/' + h), hash);
  await page.waitForTimeout(500);
}
async function toast(page) {
  await page.waitForSelector('.to-toasts .ms-MessageBar', { timeout: 10000 });
  const t = (await page.textContent('.to-toasts')).trim();
  await page.evaluate(() => document.querySelectorAll('.to-toasts .ms-MessageBar button').forEach((b) => b.click()));
  return t;
}
async function shot(page, name) {
  await page.screenshot({ path: path.join(shots, name + '.png') });
}

// Request form
async function openForm(page) {
  await page.click('.to-nav-cta button');
  await page.waitForSelector('.ms-Panel .to-request-form');
}
async function pickDropdown(page, label, option) {
  const dd = page.locator('.ms-Panel .ms-Dropdown-container', { has: page.locator(`label:text-is("${label}")`) }).locator('.ms-Dropdown');
  await dd.click();
  await page.click(`.ms-Dropdown-item:has-text("${option}")`);
}
async function pickType(page, name) {
  const dd = page.locator('.ms-Panel .ms-Dropdown-container', { has: page.locator('label:has-text("Leave type")') }).locator('.ms-Dropdown');
  await dd.click();
  await page.locator('.ms-Dropdown-item', { hasText: name }).first().click();
}
async function setDate(page, label, iso) {
  const input = page.locator('.ms-Panel .ms-DatePicker', { has: page.locator(`label:text-is("${label}")`) }).locator('input');
  await input.click();
  await input.fill(iso);
  await input.press('Enter');
  await page.keyboard.press('Escape').catch(() => undefined);
  await page.waitForTimeout(150);
}
async function pickEmployee(page, name) {
  const box = page.locator('.ms-Panel .ms-ComboBox-container', { has: page.locator('label:text-is("For")') });
  await box.locator('button').click();
  await page.locator('[role=option]', { hasText: name }).first().click();
}
async function summary(page) {
  return (await page.textContent('.ms-Panel .to-summary')).replace(/\s+/g, ' ');
}
async function submit(page) {
  const btn = page.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary');
  if (await btn.isDisabled()) throw new Error('Submit is disabled: ' + (await summary(page)));
  await btn.click();
  return toast(page);
}
async function pickYear(page, iso) {
  const y = Number(iso.substring(0, 4));
  await page.locator('.ms-Dropdown[aria-label="Leave year"]').click();
  await page.locator('.ms-Dropdown-item', { hasText: `Leave year ${y}` }).click();
  await page.waitForTimeout(200);
}
const annualCard = async (page) => (await page.locator('.to-balance', { hasText: 'Annual leave' }).first().textContent()).replace(/\s+/g, ' ');
async function openRow(page, text) {
  await page.locator('.to-table tbody tr', { hasText: text }).first().click();
  await page.waitForSelector('.ms-Panel');
  await page.waitForTimeout(300);
}
async function panelButton(page, text) {
  await page.locator('.ms-Panel button', { hasText: text }).first().click();
}
async function confirmDialog(page, input) {
  await page.waitForSelector('.ms-Dialog-main');
  if (input !== undefined) await page.fill('.ms-Dialog-main textarea', input);
  await page.locator('.ms-Dialog-main .ms-Button--primary').click();
}
async function closePanel(page) {
  const close = page.locator('.ms-Panel-closeButton');
  if (await close.count()) await close.first().click();
  await page.waitForTimeout(200);
}

async function scenario(name, fn) {
  const t0 = Date.now();
  try {
    await fn();
    results.push({ name, ok: true, ms: Date.now() - t0 });
    console.log(`  ✓ ${name}`);
  } catch (e) {
    results.push({ name, ok: false, ms: Date.now() - t0, error: e.message.split('\n')[0] });
    console.log(`  ✗ ${name}\n      ${e.message.split('\n').slice(0, 3).join('\n      ')}`);
    for (const [email, page] of Object.entries(sessions)) await shot(page, `FAIL-${results.length}-${email.split('@')[0]}`).catch(() => undefined);
  }
}

const findReq = (st, pred) => items(st, 'Requests').filter(pred);
const eli = 'eli.employee@contoso.com';
const ria = 'ria.rao@contoso.com';
const maya = 'maya.manager@contoso.com';
const hana = 'hana.hr@contoso.com';
const admin = 'admin@contoso.com';
const dave = 'dave.delegate@contoso.com';

// ------------------------------------------------------------------ scenarios
async function main() {
  emu = await start(PORT, USERS, path.join(__dirname, '..', 'e2e-dist'));
  browser = await chromium.launch();
  console.log(`Today ${TODAY}; leave week starts ${MON}; last Saturday ${LAST_SAT}\n`);

  await scenario('Setup: a site owner installs on an empty site (lists and columns created)', async () => {
    const p = await open(admin, 'home');
    await p.waitForSelector('text=Set up Brillienta');
    await p.click('label:has-text("Later (CSV import or one by one)")');
    await p.click('button:has-text("Set up Time Off Manager")');
    await p.waitForSelector('.to-app', { timeout: 30000 });
    const st = await state();
    const lists = Object.keys(st.lists).sort();
    assert.deepStrictEqual(lists, ['BrillTM_Audit', 'BrillTM_Delegations', 'BrillTM_EmailQueue', 'BrillTM_Employees', 'BrillTM_Events', 'BrillTM_Ledger', 'BrillTM_Notifications', 'BrillTM_Requests', 'BrillTM_Settings']);
    assert.ok(st.lists.BrillTM_Requests.fields.PartialDays && st.lists.BrillTM_Employees.fields.LeavingDate, 'schema v2 columns');
    assert.strictEqual(items(st, 'Settings').filter((r) => r.Title === 'SchemaVersion')[0].SettingValue, '2');
    const cfg = config(st);
    assert.strictEqual(cfg.leaveTypes.length, 30);
    assert.strictEqual(cfg.hr[0].email, admin);
  });

  await scenario('Before being added, an employee sees that they are not in the employee list', async () => {
    const p = await open(eli, 'home');
    await p.waitForSelector('text=You are not in the employee list yet');
  });

  await scenario('HR imports employees from CSV (managers, grades, start dates)', async () => {
    const p = await open(admin, 'people');
    await p.setInputFiles('input[type=file][accept=".csv,text/csv"] >> nth=0', { name: 'employees.csv', mimeType: 'text/csv', buffer: Buffer.from(EMPLOYEES_CSV) });
    const t = await toast(p);
    assert.match(t, /CSV imported/);
    const st = await state();
    const emps = items(st, 'Employees');
    assert.strictEqual(emps.length, 7);
    assert.strictEqual(emps.filter((e) => e.Email === eli)[0].ManagerEmail, maya);
    assert.strictEqual(emps.filter((e) => e.Email === ria)[0].HireDate, '2026-06-01');
  });

  await scenario('Admin adds Hana to HR and sets annual leave to "Manager, then HR for 5+ days"', async () => {
    const p = await open(admin, 'admin/access');
    const hr = p.locator('.to-card', { hasText: 'Roles' }).locator('.ms-ComboBox-container', { has: p.locator('label:text-is("HR")') });
    await hr.locator('button').click();
    await p.locator('[role=option]', { hasText: 'Hana HR' }).first().click();
    await p.click('button:has-text("Save changes")');
    assert.match(await toast(p), /Settings saved/);
    await go(p, 'admin/types');
    await p.locator('.to-table tbody tr', { hasText: 'Annual leave' }).locator('button:has-text("Edit")').click();
    await p.click('.ms-Panel button[role=tab]:has-text("Approval")');
    await pickDropdown(p, 'Approval workflow', 'Manager, then HR for 5+ days');
    await p.click('.ms-Panel button:has-text("Done")');
    await p.click('button:has-text("Save changes")');
    assert.match(await toast(p), /Settings saved/);
    const cfg = config(await state());
    assert.deepStrictEqual(cfg.hr.map((x) => x.email).sort(), [admin, hana].sort());
    assert.strictEqual(cfg.leaveTypes.filter((t) => t.key === 'annual')[0].workflowKey, 'two-level');
  });

  let annualId;
  await scenario('Employee requests 6 days of annual leave; the weekend is not counted; route = Maya then HR', async () => {
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', MON);
    await setDate(p, 'To', addDays(MON, 7));
    const s = await summary(p);
    assert.match(s, /Takes6 days/, s);
    assert.match(s, /Manager.*Maya Manager.*HR/, s);
    assert.match(await submit(p), /Sent to Maya Manager/);
    const st = await state();
    const r = findReq(st, (x) => x.EmployeeEmail === eli && x.LeaveType === 'annual')[0];
    annualId = r.Id;
    assert.strictEqual(r.Status, 'Pending');
    assert.strictEqual(r.Days, 6);
    assert.strictEqual(r.CurrentApprovers, `;${maya};`);
    assert.ok(items(st, 'Notifications').some((n) => n.Recipient === maya && /waiting for your approval/.test(n.Title)), 'bell notification for Maya');
    assert.ok(items(st, 'EmailQueue').some((m) => m.MailTo === maya && m.MailStatus === 'Queued'), 'email queued for Maya');
    assert.ok(items(st, 'Events').some((e) => e.EventType === 'submitted' && e.RequestId === r.Id), 'outbox event');
  });

  await scenario('Manager (another session) sees it in the bell and approves; it moves on to HR', async () => {
    const p = await open(maya, 'home');
    const bell = await p.textContent('.to-topbar');
    assert.match(bell, /1/);
    await go(p, 'approvals');
    await openRow(p, 'Eli Employee');
    await p.fill('.ms-Panel textarea', 'Enjoy!');
    await panelButton(p, 'Approve');
    assert.match(await toast(p), /Approved/);
    const r = findReq(await state(), (x) => x.Id === annualId)[0];
    assert.strictEqual(r.Status, 'Pending');
    assert.strictEqual(r.CurrentStep, 1);
    assert.ok(r.CurrentApprovers.indexOf(hana) >= 0 && r.CurrentApprovers.indexOf(admin) >= 0, r.CurrentApprovers);
    const steps = JSON.parse(r.Steps);
    assert.strictEqual(steps[0].comment, 'Enjoy!');
  });

  await scenario('HR approves the final level; the employee sees it approved and the balance drops by 6', async () => {
    const before = await (async () => {
      const p = await open(eli, 'requests');
      await pickYear(p, MON);
      return annualCard(p);
    })();
    const h = await open(hana, 'approvals');
    await openRow(h, 'Eli Employee');
    await panelButton(h, 'Approve');
    assert.match(await toast(h), /Approved/);
    const st = await state();
    const r = findReq(st, (x) => x.Id === annualId)[0];
    assert.strictEqual(r.Status, 'Approved');
    assert.ok(items(st, 'EmailQueue').some((m) => m.MailTo === eli && /approved/i.test(m.Title)), 'approval email to Eli');
    const p = await open(eli, 'requests');
    await pickYear(p, MON);
    const after = await annualCard(p);
    const num = (t) => Number(/([\d.]+)\s*days left/.exec(t)[1]);
    assert.strictEqual(Math.round((num(before) - num(after)) * 100) / 100, 0, `pending already counted: ${before} → ${after}`);
    assert.match(after, /6 taken/, after);
    assert.match(await p.textContent('.to-table'), /Approved/);
  });

  await scenario('Half-day casual leave is approved automatically by policy', async () => {
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    await p.click('.ms-Panel label:has-text("Half day")');
    await setDate(p, 'Date', addDays(MON, 9));
    await pickDropdown(p, 'Which half', 'Afternoon');
    assert.match(await summary(p), /Takes0.5 days.*Approved automatically/);
    assert.match(await submit(p), /Approved straight away/);
    const r = findReq(await state(), (x) => x.EmployeeEmail === eli && x.LeaveType === 'casual')[0];
    assert.strictEqual(r.Status, 'Approved');
    assert.strictEqual(r.Days, 0.5);
    assert.strictEqual(r.HalfDayPart, 'pm');
  });

  await scenario('Cancelling approved leave needs the approvers; manager confirms; days come back', async () => {
    const p = await open(eli, 'requests');
    await openRow(p, String(Number(MON.substring(8))));
    await panelButton(p, 'Cancel leave');
    await confirmDialog(p, 'Plans changed');
    await toast(p);
    let r = findReq(await state(), (x) => x.Id === annualId)[0];
    assert.strictEqual(r.Status, 'CancelRequested');
    const m = await open(maya, 'approvals');
    await openRow(m, 'Eli Employee');
    await panelButton(m, 'Approve cancellation');
    await toast(m);
    r = findReq(await state(), (x) => x.Id === annualId)[0];
    assert.strictEqual(r.Status, 'Cancelled');
    const p2 = await open(eli, 'requests');
    await pickYear(p2, MON);
    assert.match(await annualCard(p2), /0 taken/);
  });

  await scenario('Comp off: claim for a Saturday worked, manager approves, a credit with expiry appears', async () => {
    const p = await open(eli, 'home');
    await p.click('button:has-text("Claim comp off")');
    await p.waitForSelector('.ms-Panel .to-request-form');
    await setDate(p, 'From', LAST_SAT);
    await setDate(p, 'To', LAST_SAT);
    await p.fill('.ms-Panel textarea', 'Weekend release support');
    assert.match(await submit(p), /Comp-off claim sent/);
    const m = await open(maya, 'approvals');
    await openRow(m, 'Comp off claim');
    await panelButton(m, 'Approve');
    await toast(m);
    const st = await state();
    const credit = items(st, 'Ledger').filter((l) => l.EmployeeEmail === eli && l.EntryKind === 'compoff')[0];
    assert.ok(credit, 'ledger credit');
    assert.strictEqual(credit.Days, 1);
    assert.strictEqual(credit.ExpiresOn, addDays(LAST_SAT, 90));
    const p2 = await open(eli, 'requests');
    assert.match((await p2.locator('.to-balance', { hasText: 'Comp off' }).first().textContent()).replace(/\s+/g, ' '), /^.*1\s*days left/);
  });

  await scenario('A comp-off claim for a working day is refused', async () => {
    const p = await open(eli, 'home');
    await p.click('button:has-text("Claim comp off")');
    const lastFri = addDays(LAST_SAT, -1);
    await setDate(p, 'From', lastFri);
    await setDate(p, 'To', lastFri);
    await p.fill('.ms-Panel textarea', 'x');
    assert.match(await summary(p), /is a working day/);
    assert.ok(await p.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').isDisabled());
    await closePanel(p);
  });

  await scenario('Delegation: Maya delegates to Dave; a new request goes to both; Dave approves', async () => {
    const m = await open(maya, 'approvals');
    const card = m.locator('.to-card', { hasText: 'Approval delegation' });
    await card.locator('.ms-ComboBox-container', { has: m.locator('label:text-is("Delegate to")') }).locator('button').click();
    await m.locator('[role=option]', { hasText: 'Dave Delegate' }).first().click();
    await card.locator('button:has-text("Delegate")').click();
    await toast(m);
    assert.strictEqual(items(await state(), 'Delegations').length, 1);
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', addDays(MON, 14));
    await setDate(p, 'To', addDays(MON, 14));
    assert.match(await summary(p), /Dave Delegate \(for Maya Manager\)/);
    await submit(p);
    const r = findReq(await state(), (x) => x.EmployeeEmail === eli && x.StartDate === addDays(MON, 14))[0];
    assert.ok(r.CurrentApprovers.indexOf(dave) >= 0 && r.CurrentApprovers.indexOf(maya) >= 0);
    const d = await open(dave, 'approvals');
    await openRow(d, 'Eli Employee');
    await panelButton(d, 'Approve');
    await toast(d);
    assert.strictEqual(findReq(await state(), (x) => x.Id === r.Id)[0].Status, 'Approved');
  });

  await scenario('Decline with a reason; the employee sees "Declined" and the reason', async () => {
    const p = await open(ria, 'home');
    await openForm(p);
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', addDays(MON, 1));
    await setDate(p, 'To', addDays(MON, 2));
    await submit(p);
    const m = await open(maya, 'approvals');
    await openRow(m, 'Ria Rao');
    await panelButton(m, 'Decline');
    await confirmDialog(m, 'Release week — please move it');
    await toast(m);
    const r = findReq(await state(), (x) => x.EmployeeEmail === ria && x.LeaveType === 'annual')[0];
    assert.strictEqual(r.Status, 'Rejected');
    const p2 = await open(ria, 'requests');
    await openRow(p2, 'Declined');
    assert.match(await p2.textContent('.ms-Panel'), /Release week — please move it/);
  });

  let sickId;
  await scenario('HR submits sick leave on behalf of an employee; the employee is told', async () => {
    const h = await open(hana, 'home');
    await openForm(h);
    await pickEmployee(h, 'Ria Rao');
    await pickType(h, 'Sick leave');
    await setDate(h, 'From', addDays(MON, 3));
    await setDate(h, 'To', addDays(MON, 3));
    await submit(h);
    const st = await state();
    const r = findReq(st, (x) => x.EmployeeEmail === ria && x.LeaveType === 'sick')[0];
    sickId = r.Id;
    assert.strictEqual(r.SubmittedByEmail, hana);
    assert.match(JSON.parse(r.History)[0].action, /on behalf of Ria Rao/);
    assert.ok(items(st, 'Notifications').some((n) => n.Recipient === ria), 'Ria notified');
  });

  await scenario('The employee changes the pending request to 2 days; approval restarts', async () => {
    const p = await open(ria, 'requests');
    await openRow(p, 'Sick leave');
    await panelButton(p, 'Change');
    await p.waitForSelector('text=Change request');
    await setDate(p, 'To', addDays(MON, 4));
    assert.match(await summary(p), /Takes2 days/);
    await p.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').click();
    assert.match(await toast(p), /approval starts again/);
    const r = findReq(await state(), (x) => x.Id === sickId)[0];
    assert.strictEqual(r.Days, 2);
    assert.ok(JSON.parse(r.History).some((h) => /Changed — approval restarted/.test(h.action)));
  });

  await scenario('Admin approves on behalf of the manager; it is recorded and audited', async () => {
    const a = await open(admin, 'approvals');
    await a.click('.to-seg button:has-text("All open")');
    await openRow(a, 'Ria Rao');
    assert.match(await a.textContent('.ms-Panel'), /deciding on behalf/);
    await panelButton(a, 'Approve');
    await toast(a);
    const st = await state();
    const r = findReq(st, (x) => x.Id === sickId)[0];
    assert.strictEqual(r.Status, 'Approved');
    assert.strictEqual(JSON.parse(r.Steps)[0].onBehalf, true);
    assert.ok(items(st, 'Audit').some((x) => /on behalf/.test(x.Action)));
  });

  await scenario('Bi-weekly recurring WFH (3 Fridays) creates a linked series', async () => {
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Work from home');
    const fri = addDays(MON, 4);
    await setDate(p, 'From', fri);
    await setDate(p, 'To', fri);
    await pickDropdown(p, 'Repeat', 'Every two weeks');
    const times = p.locator('.ms-Panel input[role=spinbutton]').first();
    await times.fill('3');
    await times.press('Tab');
    assert.match(await summary(p), /× 3 occurrences/);
    assert.match(await submit(p), /3 requests created/);
    const series = findReq(await state(), (x) => x.EmployeeEmail === eli && x.LeaveType === 'wfh');
    assert.strictEqual(series.length, 3);
    assert.deepStrictEqual(series.map((x) => x.StartDate).sort(), [fri, addDays(fri, 14), addDays(fri, 28)]);
    assert.strictEqual(new Set(series.map((x) => x.SeriesId)).size, 1);
  });

  await scenario('Blackout period blocks employees (HR only gets a warning)', async () => {
    const start = addDays(MON, 21);
    await seedLists((lists) => {
      const row = lists.BrillTM_Settings.items.filter((r) => r.Title === 'Config')[0];
      const cfg = JSON.parse(row.SettingValue);
      cfg.blackouts = [{ key: 'b1', name: 'Year-end close', from: start, to: addDays(start, 4), departments: ['Engineering'], leaveTypes: [], mode: 'block' }];
      row.SettingValue = JSON.stringify(cfg);
    });
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    await setDate(p, 'From', start);
    await setDate(p, 'To', start);
    assert.match(await summary(p), /Year-end close: leave is not allowed/);
    assert.ok(await p.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').isDisabled());
    await closePanel(p);
    const h = await open(admin, 'home');
    await openForm(h);
    await pickEmployee(h, 'Eli Employee');
    await pickType(h, 'Casual leave');
    await setDate(h, 'From', start);
    await setDate(h, 'To', start);
    assert.match(await summary(h), /Year-end close: leave is not allowed/);
    assert.ok(!(await h.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').isDisabled()), 'HR may still submit');
    await closePanel(h);
  });

  await scenario('Opening balances import sets the available balance exactly', async () => {
    const p = await open(admin, 'people');
    const csv = `Email,Leave type,Available\n${ria},Annual leave,7.5\n${eli},SL,4\nnobody@contoso.com,Annual leave,1`;
    await p.setInputFiles('input[type=file][accept=".csv,text/csv"] >> nth=1', { name: 'opening.csv', mimeType: 'text/csv', buffer: Buffer.from(csv) });
    await toast(p);
    assert.match(await p.textContent('.ms-MessageBar'), /2 set/);
    const r = await open(ria, 'requests');
    assert.match((await r.locator('.to-balance', { hasText: 'Annual leave' }).first().textContent()).replace(/\s+/g, ' '), /7.5\s*days left/);
  });

  await scenario('Bulk adjustment gives everyone in Engineering one extra casual day', async () => {
    const p = await open(admin, 'people');
    await p.click('button:has-text("Bulk adjust")');
    await pickDropdown(p, 'Department', 'Engineering');
    await pickDropdown(p, 'Leave type', 'Casual leave');
    await p.fill('.ms-Panel input[placeholder="e.g. Company anniversary day"]', 'Founders day');
    await p.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').click();
    await toast(p);
    const grants = items(await state(), 'Ledger').filter((l) => l.Note === 'Founders day');
    assert.strictEqual(grants.length, 5);
  });

  await scenario('Reports: Excel export downloads a real workbook; leave register lists approved leave', async () => {
    const p = await open(admin, 'reports');
    const [dl] = await Promise.all([p.waitForEvent('download'), p.locator('button:has-text("Export balances")').first().click()]);
    const file = path.join(shots, 'balances.xlsx');
    await dl.saveAs(file);
    const head = fs.readFileSync(file).subarray(0, 2).toString();
    assert.strictEqual(head, 'PK');
    assert.match(await p.textContent('.to-card:has-text("Leave register")'), /approved leave/);
  });

  await scenario('Roles: an employee sees no admin pages, even by typing the address', async () => {
    const p = await open(eli, 'admin');
    const nav = await p.textContent('.to-nav');
    assert.ok(!/Admin|Reports|Employees/.test(nav), nav);
    assert.match(await p.textContent('.to-crumb'), /Overview/);
    const m = await open(maya, 'home');
    assert.match(await m.textContent('.to-nav'), /Approvals.*My team/);
  });

  await scenario('Team calendar shows approved leave of the team', async () => {
    const p = await open(maya, 'calendar');
    await p.click('button:has-text("Today")');
    // the delegated 1-day annual leave is 2 weeks after MON; move forward as needed
    for (let i = 0; i < 3 && !(await p.locator('.to-tc-leave').count()); i++) await p.click('button[aria-label="Next month"]');
    assert.ok((await p.locator('.to-tc-leave').count()) > 0, 'leave bars on the calendar');
    await shot(p, 'team-calendar');
  });

  await scenario('Two approvers decide the same request at once: the second is told, nothing is overwritten', async () => {
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', addDays(MON, 17));
    await setDate(p, 'To', addDays(MON, 17));
    await submit(p);
    const m = await open(maya, 'approvals');
    const d = await open(dave, 'approvals');
    await openRow(m, 'Eli Employee');
    await openRow(d, 'Eli Employee');
    await panelButton(d, 'Approve');
    assert.match(await toast(d), /Approved/);
    await panelButton(m, 'Decline');
    await confirmDialog(m, 'Too late?');
    assert.match(await toast(m), /Someone else has just decided/);
    const r = findReq(await state(), (x) => x.EmployeeEmail === eli && x.StartDate === addDays(MON, 17))[0];
    assert.strictEqual(r.Status, 'Approved');
    assert.strictEqual(JSON.parse(r.Steps)[0].actedBy.email, dave);
  });

  await scenario('Partial first/last day and team cover: 4 days counted, approver sees who else is away', async () => {
    const p = await open(ria, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    await setDate(p, 'From', addDays(MON, 7));
    await setDate(p, 'To', addDays(MON, 11));
    await pickDropdown(p, 'First and last day', 'Both');
    assert.match(await summary(p), /Takes4 days/);
    assert.match(await summary(p), /Also away in Engineering/);
    await submit(p);
    const r = findReq(await state(), (x) => x.EmployeeEmail === ria && x.StartDate === addDays(MON, 7))[0];
    assert.strictEqual(r.PartialDays, 'both');
    assert.strictEqual(r.Days, 4);
    const m = await open(maya, 'approvals');
    await openRow(m, 'Casual leave');
    assert.match(await m.textContent('.ms-Panel'), /Team cover · Engineering/);
    await closePanel(m);
  });

  await scenario('Reminders: a request waiting 30 h reminds its approver once', async () => {
    const p0 = await open(ria, 'home');
    await openForm(p0);
    await pickType(p0, 'Annual leave');
    await setDate(p0, 'From', addDays(MON, 15));
    await setDate(p0, 'To', addDays(MON, 15));
    await submit(p0);
    await seedLists((lists) => {
      const r = lists.BrillTM_Requests.items.filter((x) => x.EmployeeEmail === ria && x.StartDate === addDays(MON, 15))[0];
      const steps = JSON.parse(r.Steps);
      steps[0].startedAt = new Date(Date.now() - 30 * 3600000).toISOString();
      r.Steps = JSON.stringify(steps);
      lists.BrillTM_Settings.items.filter((s) => s.Title === 'LastReminderRun').forEach((s) => (s.SettingValue = new Date(Date.now() - 2 * 3600000).toISOString()));
    });
    sessions[eli] && (await sessions[eli].close(), delete sessions[eli]);
    await open(eli, 'home');
    await new Promise((r) => setTimeout(r, 2500));
    const st = await state();
    const r = findReq(st, (x) => x.EmployeeEmail === ria && x.StartDate === addDays(MON, 15))[0];
    const steps = JSON.parse(r.Steps);
    assert.strictEqual(steps[0].reminders, 1, JSON.stringify(steps[0]));
    assert.ok(items(st, 'Notifications').some((n) => n.Recipient === maya && /Reminder/.test(n.Title)));
  });

  await scenario('Access: apply recommended permissions breaks inheritance on every list', async () => {
    const p = await open(admin, 'admin/access');
    await p.click('button:has-text("Apply recommended permissions")');
    assert.match(await toast(p), /Permissions applied/);
    const st = await state();
    Object.values(st.lists).forEach((l) => assert.ok(l.unique && l.assignments.length >= 1, l.title));
  });

  await scenario('Upgrade: a 1.0 site keeps working for employees, and a site owner visit adds the new columns', async () => {
    await seedLists((lists) => {
      delete lists.BrillTM_Requests.fields.PartialDays;
      delete lists.BrillTM_Requests.fields.AutoReply;
      delete lists.BrillTM_Requests.fields.GroupEventId;
      delete lists.BrillTM_Employees.fields.LeavingDate;
      lists.BrillTM_Settings.items.filter((s) => s.Title === 'SchemaVersion')[0].SettingValue = '1';
    });
    for (const k of Object.keys(sessions)) {
      await sessions[k].context().close();
      delete sessions[k];
    }
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    await setDate(p, 'From', addDays(MON, 16));
    await setDate(p, 'To', addDays(MON, 16));
    assert.match(await submit(p), /Approved straight away|Sent to/);
    let st = await state();
    assert.ok(!st.lists.BrillTM_Requests.fields.PartialDays, 'employee did not change the schema');
    await open(admin, 'home');
    await new Promise((r) => setTimeout(r, 1500));
    st = await state();
    assert.ok(st.lists.BrillTM_Requests.fields.PartialDays && st.lists.BrillTM_Employees.fields.LeavingDate, 'columns added');
    assert.strictEqual(items(st, 'Settings').filter((s) => s.Title === 'SchemaVersion')[0].SettingValue, '2');
  });

  await scenario('Licence: when the trial has ended the app locks (data kept) and shows the licence page to admins', async () => {
    await seedLists((lists) => {
      lists.BrillTM_Settings.created = new Date(Date.now() - 30 * 86400000).toISOString();
      lists.BrillTM_Settings.items.filter((s) => s.Title === 'InstalledAt')[0].SettingValue = new Date(Date.now() - 30 * 86400000).toISOString();
    });
    for (const k of Object.keys(sessions)) {
      await sessions[k].context().close();
      delete sessions[k];
    }
    const p = await open(eli, 'home');
    await p.waitForSelector('.to-locked');
    assert.match(await p.textContent('.to-locked'), /trial has ended.*Ask your Time Off Manager admin/s);
    const a = await open(admin, 'home');
    assert.match(await a.textContent('.to-locked'), /Activate a licence key/);
    await seedLists((lists) => {
      lists.BrillTM_Settings.created = new Date().toISOString();
      lists.BrillTM_Settings.items.filter((s) => s.Title === 'InstalledAt')[0].SettingValue = new Date().toISOString();
    });
  });

  await scenario('Sample organisation can be installed and removed again', async () => {
    await fetch(`${BASE}/__reset`, { method: 'POST' });
    for (const k of Object.keys(sessions)) {
      await sessions[k].context().close();
      delete sessions[k];
    }
    const p = await open(admin, 'home');
    await p.click('label:has-text("Add a sample organisation")');
    await p.click('button:has-text("Set up Time Off Manager")');
    await p.waitForSelector('.to-app', { timeout: 40000 });
    let st = await state();
    assert.ok(items(st, 'Employees').length >= 12 && items(st, 'Requests').length > 5);
    await go(p, 'admin/access');
    await p.click('button:has-text("Remove the sample data")');
    await confirmDialog(p);
    console.log('      toast: ' + (await toast(p)));
    st = await state();
    assert.strictEqual(items(st, 'Employees').filter((e) => /@example\.com$/.test(e.Email)).length, 0);
    assert.strictEqual(items(st, 'Requests').filter((e) => /@example\.com$/.test(e.EmployeeEmail)).length, 0);
  });

  // Page errors in any session are failures too.
  const errors = Object.entries(sessions).flatMap(([u, pg]) => pg.errors.map((e) => `${u}: ${e}`));
  await browser.close();
  emu.server.close();
  const passed = results.filter((r) => r.ok).length;
  console.log(`\n${passed}/${results.length} scenarios passed${errors.length ? `; page errors:\n  ${errors.join('\n  ')}` : '; no page errors'}`);
  fs.writeFileSync(path.join(shots, 'results.json'), JSON.stringify({ today: TODAY, results, errors, findings }, null, 2));
  process.exit(passed === results.length && !errors.length ? 0 : 1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
