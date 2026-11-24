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
/** Open the app the way Teams does: Teams' host page on another site (the web part finds the Time Off Manager site). */
async function openTeams(email, hostSite, extra) {
  const page = await as(email);
  await page.goto('about:blank');
  await page.goto(`${BASE}${hostSite}/_layouts/15/teamshostedapp.aspx?teams=1&user=${encodeURIComponent(email)}&host=${encodeURIComponent(hostSite)}${extra || ''}#/home`);
  await page.waitForSelector('.to-app, .to-setup, .to-locked, .to-card', { timeout: 20000 });
  await page.waitForTimeout(600);
  return page;
}
async function fresh(email) {
  if (sessions[email]) {
    await sessions[email].context().close();
    delete sessions[email];
  }
}
/** A licence key signed with the test bundle's key pair (e2e/build.cjs). */
function licenceKey(payload) {
  const crypto = require('crypto');
  const b64url = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const key = crypto.createPrivateKey(fs.readFileSync(path.join(__dirname, '..', 'e2e-dist', 'test-license-key.pem')));
  const signed = 'TO1.' + b64url(JSON.stringify({ v: 1, id: 'TO-TEST', product: 'timeoff', customer: 'Contoso', issued: TODAY, ...payload }));
  return signed + '.' + b64url(crypto.sign('sha256', Buffer.from(signed), { key, dsaEncoding: 'ieee-p1363' }));
}
const TENANT = '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13';
const setConfig = (mutate) =>
  seedLists((lists) => {
    const row = lists.BrillTM_Settings.items.filter((r) => r.Title === 'Config')[0];
    const cfg = JSON.parse(row.SettingValue);
    mutate(cfg);
    row.SettingValue = JSON.stringify(cfg);
  });
const cardText = async (page, name) => (await page.locator('.to-balance', { hasText: name }).first().textContent()).replace(/\s+/g, ' ');
const daysLeft = (t) => Number(/(-?[\d.]+)\s*days left/.exec(t)[1]);

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
    assert.ok(st.lists.BrillTM_Requests.fields.PartialDays && st.lists.BrillTM_Employees.fields.LeavingDate, 'all columns');
    assert.strictEqual(items(st, 'Settings').filter((r) => r.Title === 'SchemaVersion')[0].SettingValue, '1');
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

  await scenario('Access: BrillTM groups are created, filled from Admin › Access and granted on every list', async () => {
    const p = await open(admin, 'admin/access');
    assert.match(await p.locator('.to-admin').innerText(), /BrillTM Admins[\s\S]*not created yet/);
    await p.click('button:has-text("Apply recommended permissions")');
    assert.match(await toast(p), /Permissions applied/);
    let st = await state();
    const g = (name) => Object.values(st.groups).filter((x) => x.Title === name)[0];
    const emails = (name) => g(name).members.map((m) => m.Email).sort();
    assert.ok(g('BrillTM Admins') && g('BrillTM HR') && g('BrillTM Users'), Object.values(st.groups).map((x) => x.Title).join(', '));
    assert.deepStrictEqual(emails('BrillTM Admins'), [admin]);
    assert.deepStrictEqual(emails('BrillTM HR'), [admin, hana].sort());
    assert.ok(g('BrillTM Users').members.some((m) => /spo-grid-all-users/.test(m.LoginName)), 'Everyone except external users in BrillTM Users');
    const READ = 1073741826;
    const CONTRIBUTE = 1073741827;
    Object.values(st.lists).forEach((l) => {
      assert.ok(l.unique, l.title + ' stops inheriting');
      const role = (id) => l.assignments.filter((a) => a.principal === id).map((a) => a.role);
      const usersRole = ['BrillTM_Settings', 'BrillTM_Employees', 'BrillTM_Audit'].indexOf(l.title) >= 0 ? READ : CONTRIBUTE;
      assert.deepStrictEqual(role(g('BrillTM Users').Id), [usersRole], l.title + ' users');
      assert.deepStrictEqual(role(g('BrillTM HR').Id), [CONTRIBUTE], l.title + ' HR');
      assert.deepStrictEqual(role(g('BrillTM Admins').Id), [CONTRIBUTE], l.title + ' admins');
      assert.ok(l.assignments.every((a) => a.principal >= 10), l.title + ': only the BrillTM groups are granted directly');
    });
    // Adding an app admin under Access and saving updates the BrillTM Admins group.
    const admins = p.locator('.to-card', { hasText: 'Roles' }).locator('.ms-ComboBox-container', { has: p.locator('label:text-is("App admins")') });
    await admins.locator('button').click();
    await p.locator('[role=option]', { hasText: 'Dan Director' }).first().click();
    await p.click('button:has-text("Save changes")');
    await p.waitForTimeout(800);
    st = await state();
    assert.deepStrictEqual(emails('BrillTM Admins'), [admin, 'dan.director@contoso.com']);
    // Removing Hana from HR takes her out of BrillTM HR.
    await p.locator('.to-card', { hasText: 'Roles' }).locator('.to-chip', { hasText: 'Hana HR' }).locator('button').click();
    await p.click('button:has-text("Save changes")');
    await p.waitForTimeout(800);
    st = await state();
    assert.deepStrictEqual(emails('BrillTM HR'), [admin]);
    await go(p, 'people');
    await go(p, 'admin/access');
    assert.match(await p.locator('.to-admin').innerText(), /BrillTM HR\s*1 member/);
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

  await scenario('List names: every list and every write uses the BrillTM_ prefix', async () => {
    const st = await state();
    const names = Object.keys(st.lists);
    assert.ok(names.length === 9 && names.every((n) => n.indexOf('BrillTM_') === 0), names.join(', '));
    const writes = st.log.filter((l) => / (POST|MERGE|DELETE) /.test(l) && /getbytitle\('/.test(l));
    const elsewhere = writes.filter((l) => !/getbytitle\('BrillTM_/.test(l));
    assert.ok(writes.length > 100, `${writes.length} writes`);
    assert.deepStrictEqual(elsewhere, []);
  });

  await scenario('Admin screens show the BrillTM_ list names (Power Automate, API, Power BI, flow package)', async () => {
    const p = await open(admin, 'admin/integrations');
    const integ = await p.locator('.to-admin').innerText();
    const fields = await p.$$eval('.to-admin input', (els) => els.map((e) => e.value).join(' '));
    assert.match(integ + fields, /BrillTM_Events/);
    assert.match(fields, /getbytitle\('BrillTM_Requests'\)/);
    await go(p, 'admin/notifications');
    assert.match(await p.locator('.to-admin').innerText(), /BrillTM_EmailQueue/);
    const [dl] = await Promise.all([p.waitForEvent('download'), p.click('button:has-text("Download the flow package")')]);
    const file = path.join(shots, 'flow.zip');
    await dl.saveAs(file);
    const zip = fs.readFileSync(file).toString('latin1');
    assert.strictEqual(zip.substring(0, 2), 'PK');
    assert.ok(zip.indexOf('BrillTM_EmailQueue') >= 0, 'flow package targets BrillTM_EmailQueue');
    await go(p, 'reports');
    const bi = await p.$$eval('input[readonly]', (els) => els.map((e) => e.value).join(' '));
    assert.match(bi, /BrillTM_Requests/);
    assert.match(bi, /BrillTM_Ledger/);
  });

  // ---------------------------------------------------------------- Microsoft Teams

  await scenario('Teams: opened from a team site, the app finds the HR site with SharePoint search', async () => {
    const p = await openTeams(eli, '/sites/eng');
    const site = await p.evaluate(() => window.__site);
    assert.strictEqual(site.source, 'found');
    assert.strictEqual(site.url, `${BASE}/sites/hr`);
    await go(p, 'requests');
    await p.waitForSelector('.to-balance');
    assert.match(await p.textContent('.to-page'), /Annual leave/);
    // Outside Teams the same team site has no Time Off Manager: an employee is told it is not set up.
    await fresh(eli);
    const q = await (await as(eli)).goto(`${BASE}/?user=${eli}&host=/sites/eng#/home`).then(() => sessions[eli]);
    await q.waitForSelector('.to-setup');
    assert.match(await q.textContent('.to-setup'), /not set up on this site yet/);
  });

  await scenario('Teams: links in emails sent from Teams point at the SharePoint page, not Teams’ host page', async () => {
    await open(admin, 'home'); // an HR admin on the real page records it
    await new Promise((r) => setTimeout(r, 800));
    const stored = items(await state(), 'Settings').filter((r) => r.Title === 'PageUrl')[0];
    assert.ok(stored && stored.SettingValue.indexOf('teamshostedapp') < 0, JSON.stringify(stored));
    await fresh(eli);
    const p = await openTeams(eli, '/sites/eng');
    await p.click('.to-nav-cta button');
    await p.waitForSelector('.ms-Panel .to-request-form');
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', addDays(MON, 43));
    await setDate(p, 'To', addDays(MON, 43));
    await submit(p);
    const mail = items(await state(), 'EmailQueue').slice(-1)[0];
    const link = ((/href="([^"]+)"/.exec(mail.MailBody) || [])[1] || '').replace(/&amp;/g, '&');
    assert.strictEqual(stored.SettingValue, `${BASE}/`, 'recorded without query string');
    assert.match(link, new RegExp(`^${BASE}/#/approvals\\?id=\\d+$`), link);
    assert.ok(mail.MailBody.indexOf('teamshostedapp') < 0, 'no Teams host page link');
  });

  await scenario('Teams: several Time Off Manager sites → the person picks once; an admin can mark the default', async () => {
    const st = await state();
    const copy = JSON.parse(JSON.stringify(st.lists.BrillTM_Settings));
    copy.items = copy.items.filter((i) => i.Title !== 'TeamsDefault');
    st.sites['/sites/emea'] = { title: 'EMEA HR', lists: { BrillTM_Settings: copy }, groups: {}, nextGroup: 10 };
    await fresh(ria);
    const p = await openTeams(ria, '/sites/eng');
    await p.waitForSelector('text=Which Time Off Manager?');
    const text = await p.textContent('.to-setup');
    assert.match(text, /\/sites\/hr/);
    assert.match(text, /\/sites\/emea/);
    await p.locator('.to-deleg-row', { hasText: '/sites/hr' }).locator('button').click();
    await p.waitForSelector('.to-app');
    assert.strictEqual((await p.evaluate(() => window.__site)).url, `${BASE}/sites/hr`);
    const again = await openTeams(ria, '/sites/eng'); // remembered on this device
    assert.strictEqual((await again.evaluate(() => window.__site)).url, `${BASE}/sites/hr`);
    // Admin marks the HR site for Teams: someone new opens it directly.
    const a = await open(admin, 'admin/m365');
    await a.locator('.ms-Toggle', { hasText: 'Use this site in Teams' }).locator('button').click();
    assert.match(await toast(a), /Teams opens this site/);
    await fresh(dave);
    const d = await openTeams(dave, '/sites/eng');
    assert.strictEqual((await d.evaluate(() => window.__site)).url, `${BASE}/sites/hr`);
    assert.ok(!(await d.locator('text=Which Time Off Manager?').count()));
    delete (await state()).sites['/sites/emea'];
  });

  await scenario('Teams tab setting: a site address on the tab opens that site; a wrong address is explained', async () => {
    await fresh(eli);
    const p = await (await as(eli)).goto(`${BASE}/?user=${eli}&host=/sites/eng&site=${encodeURIComponent(BASE + '/sites/hr/SitePages/Leave.aspx')}#/home`).then(() => sessions[eli]);
    await p.waitForSelector('.to-app');
    const site = await p.evaluate(() => window.__site);
    assert.deepStrictEqual([site.source, site.url], ['property', `${BASE}/sites/hr`]);
    await fresh(eli);
    const q = await (await as(eli)).goto(`${BASE}/?user=${eli}&host=/sites/eng&site=${encodeURIComponent('https://evil.example.com/sites/hr')}#/home`).then(() => sessions[eli]);
    await q.waitForSelector('text=is not a SharePoint site address on this tenant');
  });

  // ---------------------------------------------------------------- regressions found in review

  await scenario('A delegate who is also the requester never approves their own leave', async () => {
    await seedLists((lists) => lists.BrillTM_Delegations.items.forEach((d) => (d.IsActive = false)));
    const m = await open(maya, 'approvals');
    const card = m.locator('.to-card', { hasText: 'Approval delegation' });
    await card.locator('.ms-ComboBox-container', { has: m.locator('label:text-is("Delegate to")') }).locator('button').click();
    await m.locator('[role=option]', { hasText: 'Eli Employee' }).first().click();
    await card.locator('button:has-text("Delegate")').click();
    await toast(m);
    await fresh(eli);
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Annual leave');
    await setDate(p, 'From', addDays(MON, 44));
    await setDate(p, 'To', addDays(MON, 44));
    const s1 = await summary(p);
    assert.match(s1, /ManagerMaya Manager/);
    assert.ok(!/for Maya Manager/.test(s1), s1);
    await submit(p);
    const r = findReq(await state(), (x) => x.EmployeeEmail === eli && x.StartDate === addDays(MON, 44))[0];
    assert.strictEqual(r.CurrentApprovers, `;${maya};`);
    await go(p, 'approvals');
    const who = await p.$$eval('.to-table tbody tr td:nth-child(2) .to-person-name', (els) => els.map((e) => e.textContent));
    assert.ok(who.length >= 1 && who.every((n) => n === 'Ria Rao'), 'as delegate Eli sees Ria’s requests but never his own: ' + who.join(', '));
    await seedLists((lists) => lists.BrillTM_Delegations.items.forEach((d) => (d.IsActive = false)));
  });

  await scenario('Withdrawing a pending request with approval: declining it puts it back to Pending (not Approved)', async () => {
    await setConfig((c) => (c.cancellation.pendingDirect = false));
    await fresh(eli);
    const p = await open(eli, 'requests');
    await openRow(p, String(Number(addDays(MON, 44).substring(8))) + ' ');
    await panelButton(p, 'Withdraw');
    await confirmDialog(p, 'Maybe not');
    await toast(p);
    const id = findReq(await state(), (x) => x.EmployeeEmail === eli && x.StartDate === addDays(MON, 44))[0].Id;
    assert.strictEqual(findReq(await state(), (x) => x.Id === id)[0].Status, 'CancelRequested');
    const m = await open(maya, 'approvals');
    await openRow(m, 'Cancellation pending');
    await panelButton(m, 'Keep the leave');
    await confirmDialog(m, 'Let us talk first');
    await toast(m);
    const st = await state();
    assert.strictEqual(findReq(st, (x) => x.Id === id)[0].Status, 'Pending');
    assert.ok(items(st, 'Notifications').some((n) => n.Recipient === eli && /cancellation was declined/i.test(n.Title)));
    assert.ok(!items(st, 'EmailQueue').some((mm) => mm.MailTo === eli && /approved/i.test(mm.Title) && mm.Id > 0 && mm.Template === 'approved' && mm.Title.indexOf(addDays(MON, 44)) >= 0));
    await setConfig((c) => (c.cancellation.pendingDirect = true));
  });

  await scenario('Admins cannot approve their own leave; it goes to their manager', async () => {
    const a = await open(admin, 'home');
    await openForm(a);
    await pickType(a, 'Casual leave');
    await setDate(a, 'From', addDays(MON, 45));
    await setDate(a, 'To', addDays(MON, 46));
    assert.match(await summary(a), /ManagerDan Director/);
    await submit(a);
    await go(a, 'approvals');
    await a.click('.to-seg button:has-text("All open")');
    await openRow(a, 'Alex Admin');
    assert.strictEqual(await a.locator('.ms-Panel button', { hasText: /^Approve$/ }).count(), 0, 'no approve button on own request');
    assert.ok((await a.locator('.ms-Panel button', { hasText: 'Remind approver' }).count()) === 1);
    await closePanel(a);
  });

  // ---------------------------------------------------------------- more coverage

  await scenario('Leave template: “Doctor appointment (2 h)” fills the form; 2-hour leave is approved by policy', async () => {
    const p = await open(eli, 'home');
    await p.click('.to-hero-templates button:has-text("Doctor appointment")');
    await p.waitForSelector('.ms-Panel .to-request-form');
    await setDate(p, 'Date', addDays(MON, 50));
    assert.match(await summary(p), /Takes2 h.*Approved automatically/);
    assert.match(await submit(p), /Approved straight away/);
    const r = findReq(await state(), (x) => x.EmployeeEmail === eli && x.LeaveType === 'medical')[0];
    assert.deepStrictEqual([r.Unit, r.Hours, r.StartTime, r.Status], ['hours', 2, '10:00', 'Approved']);
  });

  await scenario('Changing a policy applies at once: casual leave 8 → 10 days raises the balance by 2', async () => {
    let p = await open(eli, 'requests');
    const before = daysLeft(await cardText(p, 'Casual leave'));
    const a = await open(admin, 'admin/types');
    await a.locator('.to-table tbody tr', { hasText: 'Casual leave' }).locator('button:has-text("Edit")').click();
    await a.click('.ms-Panel button[role=tab]:has-text("Entitlement")');
    const days = a.locator('.ms-Panel input[role=spinbutton]').first();
    await days.fill('10');
    await days.press('Tab');
    await a.click('.ms-Panel button:has-text("Done")');
    await a.click('button:has-text("Save changes")');
    await toast(a);
    p = await open(eli, 'requests');
    assert.strictEqual(daysLeft(await cardText(p, 'Casual leave')), before + 2);
  });

  await scenario('A new company holiday is left out of leave counts straight away', async () => {
    const day = addDays(MON, 51); // a Wednesday
    const a = await open(admin, 'admin/calendars');
    const card = a.locator('.to-card', { hasText: 'Head office' }).first();
    const date = card.locator('.ms-DatePicker', { has: a.locator('label:text-is("Date")') }).locator('input');
    await date.click();
    await date.fill(day);
    await date.press('Enter');
    await a.keyboard.press('Escape');
    await card.locator('input[type=text]').last().fill('Company day');
    await card.locator('button:has-text("Add")').click();
    await a.click('button:has-text("Save changes")');
    await toast(a);
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    await setDate(p, 'From', addDays(MON, 49));
    await setDate(p, 'To', addDays(MON, 53));
    const s1 = await summary(p);
    assert.match(s1, /Takes4 days/, s1);
    assert.match(s1, /Company day/);
    await closePanel(p);
  });

  await scenario('Bulk approval: a manager approves everything waiting in one go', async () => {
    const m = await open(maya, 'approvals');
    const waiting = await m.locator('.to-table tbody tr').count();
    assert.ok(waiting >= 2, `${waiting} waiting`);
    await m.locator('.to-table thead .ms-Checkbox').click();
    await m.click(`button:has-text("Approve ${waiting}")`);
    assert.match(await toast(m), new RegExp(`${waiting} requests? approved`));
    const st = await state();
    assert.strictEqual(findReq(st, (x) => x.Status === 'Pending' && (x.CurrentApprovers || '').indexOf(`;${maya};`) >= 0).length, 0);
  });

  await scenario('My team: the manager sees both reports and exports their balances to Excel', async () => {
    const m = await open(maya, 'team');
    const text = await m.textContent('.to-page');
    assert.match(text, /Eli Employee/);
    assert.match(text, /Ria Rao/);
    const [dl] = await Promise.all([m.waitForEvent('download'), m.click('button:has-text("Export balances")')]);
    const file = path.join(shots, 'team.xlsx');
    await dl.saveAs(file);
    assert.strictEqual(fs.readFileSync(file).subarray(0, 2).toString(), 'PK');
  });

  await scenario('Notifications: opening one from the bell marks it read', async () => {
    const p = await open(eli, 'home');
    await p.click('button[aria-label^="Notifications"]');
    await p.waitForSelector('.ms-Callout');
    const unreadBefore = items(await state(), 'Notifications').filter((n) => n.Recipient === eli && !n.IsRead).length;
    assert.ok(unreadBefore > 0);
    await p.locator('.ms-Callout [role=button]').first().click();
    await new Promise((r) => setTimeout(r, 500));
    const unreadAfter = items(await state(), 'Notifications').filter((n) => n.Recipient === eli && !n.IsRead).length;
    assert.strictEqual(unreadAfter, unreadBefore - 1);
  });

  await scenario('Phone: an employee requests leave on a phone-sized screen without sideways scrolling', async () => {
    await fresh(ria);
    const { devices } = require('playwright');
    const ctx = await browser.newContext({ ...devices['iPhone 13'] });
    const p = await ctx.newPage();
    p.errors = [];
    p.on('pageerror', (e) => p.errors.push(String(e)));
    sessions[ria] = p;
    await p.goto(`${BASE}/?user=${ria}#/home`);
    await p.waitForSelector('.to-app');
    await p.click('.to-hero button:has-text("Request time off")');
    await p.waitForSelector('.ms-Panel .to-request-form');
    await pickType(p, 'Casual leave');
    await setDate(p, 'From', addDays(MON, 56));
    await setDate(p, 'To', addDays(MON, 56));
    const wide = await p.evaluate(() => [document.documentElement.scrollWidth, window.innerWidth]);
    assert.ok(wide[0] <= wide[1] + 1, `page ${wide[0]}px on a ${wide[1]}px screen`);
    await submit(p);
    assert.ok(findReq(await state(), (x) => x.EmployeeEmail === ria && x.StartDate === addDays(MON, 56)).length === 1);
  });

  await scenario('New leave year: the first HR visit in January carries balances forward once', async () => {
    const st0 = await state();
    const year = Number(TODAY.substring(0, 4));
    assert.ok(!items(st0, 'Ledger').some((l) => (l.UniqueKey || '').indexOf(`carry:${year}:`) === 0));
    await fresh(admin);
    const page = await as(admin);
    await page.clock.setFixedTime(new Date(`${year + 1}-01-02T09:00:00`));
    await page.goto(`${BASE}/?user=${admin}&owner=1#/home`);
    await page.waitForSelector('.to-app');
    await new Promise((r) => setTimeout(r, 2500));
    const st = await state();
    const carried = items(st, 'Ledger').filter((l) => (l.UniqueKey || '').indexOf(`carry:${year}:`) === 0);
    assert.ok(carried.length > 0, 'carry-forward entries');
    assert.ok(carried.every((l) => l.LeaveYear === year + 1 && l.EntryKind === 'carry'));
    assert.ok(carried.filter((l) => l.LeaveType === 'annual').every((l) => l.Days <= 5), 'annual limited to 5 days');
    assert.match(items(st, 'Settings').filter((r) => r.Title === 'CarriedYears')[0].SettingValue, new RegExp(String(year)));
    await page.goto('about:blank');
    await page.goto(`${BASE}/?user=${admin}&owner=1#/home`);
    await new Promise((r) => setTimeout(r, 1500));
    assert.strictEqual(items(await state(), 'Ledger').filter((l) => (l.UniqueKey || '').indexOf(`carry:${year}:`) === 0).length, carried.length, 'only once');
    await fresh(admin);
  });

  await scenario('Licence: a key for another tenant is refused; a 3-user key limits who can request leave', async () => {
    const a = await open(admin, 'admin/licence');
    await a.fill('textarea[placeholder="TO1.…"]', licenceKey({ tenant: '11111111-2222-3333-4444-555555555555', users: 50 }));
    await a.click('button:has-text("Activate")');
    assert.match(await toast(a), /another Microsoft 365 organisation/);
    await a.fill('textarea[placeholder="TO1.…"]', licenceKey({ tenant: TENANT, users: 3, expires: addDays(TODAY, 365) }));
    await a.click('button:has-text("Activate")');
    assert.match(await toast(a), /Licence activated/);
    assert.match(await a.textContent('.to-admin'), /Licensed for 3 users/);
    await go(a, 'people');
    assert.strictEqual(await a.locator('.to-badge.is-warn', { hasText: 'Over licence' }).count(), 4);
    await fresh(eli);
    const p = await open(eli, 'home');
    await openForm(p);
    await pickType(p, 'Casual leave');
    assert.match(await summary(p), /above the number of users in the licence/);
    assert.ok(await p.locator('.ms-Panel .ms-Panel-footer button.ms-Button--primary').isDisabled());
    await closePanel(p);
    await seedLists((lists) => (lists.BrillTM_Settings.items = lists.BrillTM_Settings.items.filter((r) => r.Title !== 'LicenseKey')));
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
