// Local preview: the real app (compiled lib/) against an in-memory SharePoint seeded with the sample organisation.
// ?role=admin (default) | hr | manager | employee   ?fresh=1 shows the setup wizard.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from '../lib/webparts/timeOffManager/components/App';
import { defaultConfig } from '../lib/webparts/timeOffManager/schema/defaults';
import { SCHEMA_VERSION } from '../lib/webparts/timeOffManager/schema/schema';
import { requestToItem, employeeToItem } from '../lib/webparts/timeOffManager/services/LeaveService';
import { sampleEmployees, sampleRequests, sampleLedger } from '../lib/webparts/timeOffManager/services/sample';
import { createFakeHttp, seed } from './fakeSharePoint';
import { version as APP_VERSION } from '../package.json';

const params = new URLSearchParams(window.location.search);
const role = params.get('role') || 'admin';
const WEB = 'https://contoso.sharepoint.com/sites/hr';
const people = { admin: ['Priya Nair', 'priya.nair@example.com'], hr: ['Priya Nair', 'priya.nair@example.com'], manager: ['Priya Nair', 'priya.nair@example.com'], employee: ['Kenji Watanabe', 'kenji.watanabe@example.com'] }[role] || ['Priya Nair', 'priya.nair@example.com'];
const user = { Id: 7, Title: people[0], Email: people[1], LoginName: 'i:0#.f|membership|' + people[1], IsSiteAdmin: role === 'admin' };
const author = { Id: 7, Title: user.Title, EMail: user.Email };
const http = createFakeHttp(user);

function render() {
  initializeIcons(undefined, { disableWarnings: true });
  ReactDom.render(
    React.createElement(App, { http, webUrl: WEB, siteTitle: 'Human Resources', permissions: { manageLists: role === 'admin', manageWeb: role === 'admin' }, tenantId: '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13', prefix: 'TO_', appVersion: APP_VERSION, host: { inTeams: false, view: params.get('view') || 'app' } }),
    document.getElementById('root')
  );
}

if (!params.get('fresh')) {
  const cfg = defaultConfig('Acme');
  const me = { email: 'priya.nair@example.com', name: 'Priya Nair' };
  if (role === 'hr' || role === 'admin') cfg.hr = [me];
  const iso = (d) => d.toISOString().substring(0, 10);
  cfg.blackouts = [{ key: 'ye', name: 'Year-end close', from: iso(new Date(Date.now() + 35 * 864e5)), to: iso(new Date(Date.now() + 40 * 864e5)), departments: ['Engineering'], leaveTypes: [], mode: 'block' }];
  cfg.staffing = { enabled: true, maxAwayPercent: 50, maxAwayCount: 0, mode: 'warn', perDepartment: {} };
  const emps = sampleEmployees(me);
  seed('TO_Settings', [
    { Title: 'SchemaVersion', SettingValue: String(SCHEMA_VERSION) },
    { Title: 'AppVersion', SettingValue: APP_VERSION },
    { Title: 'InstalledAt', SettingValue: new Date(Date.now() - 3 * 86400000).toISOString() },
    { Title: 'LastReminderRun', SettingValue: new Date().toISOString() },
    { Title: 'CarriedYears', SettingValue: '[' + (new Date().getFullYear() - 1) + ']' },
    { Title: 'Config', SettingValue: JSON.stringify(cfg) }
  ], author);
  seed('TO_Employees', emps.map(employeeToItem), author);
  seed('TO_Requests', sampleRequests(cfg, emps, me).map((r) => requestToItem(r, (cfg.leaveTypes.find((t) => t.key === r.leaveType) || {}).name || r.leaveType, (emps.find((e) => e.email === r.employeeEmail) || {}).department || '')), author);
  seed('TO_Ledger', sampleLedger(cfg, emps).map((l) => ({ Title: l.kind, EmployeeEmail: l.employeeEmail, LeaveType: l.leaveType, LeaveYear: l.leaveYear, Days: l.days, EntryKind: l.kind, Note: l.note, Effective: l.effective, ExpiresOn: l.expiresOn || '', UniqueKey: l.uniqueKey || '' })), author);
  seed('TO_Notifications', [{ Title: 'Leave request waiting for your approval', Recipient: 'priya.nair@example.com', Body: 'Kenji Watanabe: Annual leave', Link: 'approvals', IsRead: false }], author);
  seed('TO_Delegations', [], author);
}
render();
