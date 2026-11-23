import * as React from 'react';
import { DefaultButton, Dropdown, Icon, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IConfig } from '../../models/types';
import { Card, PageHeader, PeopleField, useBusy, useConfirm, NumberBox } from '../common/Ui';
import { LeaveTypesAdmin } from './LeaveTypesAdmin';
import { WorkflowsAdmin } from './WorkflowsAdmin';
import { CalendarsAdmin } from './CalendarsAdmin';
import { TemplatesAdmin } from './TemplatesAdmin';
import { LicenseAdmin } from './LicenseAdmin';
import { M365Admin, RulesAdmin } from './RulesAdmin';
import { flowPackageZip } from '../../services/flowPackage';
import { downloadBytes } from '../../services/exporting';
import { DelegationCard } from '../approvals/Approvals';
import { MONTH_NAMES, relativeTime } from '../../services/dates';
import { Provisioner } from '../../services/Provisioner';
import { SETTING_KEYS } from '../../services/SettingsStore';
import { listTitle, LISTS, SCHEMA_VERSION } from '../../schema/schema';
import { DEFAULT_ACCENT, DEFAULT_BRAND } from '../../schema/defaults';

export interface ISectionProps {
  draft: IConfig;
  setDraft: (c: IConfig) => void;
}

const SECTIONS: Array<{ key: string; label: string; icon: string; adminOnly?: boolean }> = [
  { key: 'general', label: 'General & branding', icon: 'Color' },
  { key: 'types', label: 'Leave types & policies', icon: 'Airplane' },
  { key: 'workflows', label: 'Approval workflows', icon: 'Flow' },
  { key: 'calendars', label: 'Work weeks & holidays', icon: 'Calendar' },
  { key: 'templates', label: 'Leave templates', icon: 'Lightning' },
  { key: 'rules', label: 'Blackouts & staffing', icon: 'Blocked' },
  { key: 'm365', label: 'Teams & Outlook', icon: 'TeamsLogo' },
  { key: 'delegation', label: 'Delegations', icon: 'Switch' },
  { key: 'notifications', label: 'Notifications & reminders', icon: 'Ringer' },
  { key: 'integrations', label: 'Power Automate & API', icon: 'PlugConnected' },
  { key: 'access', label: 'Access & security', icon: 'Permissions', adminOnly: true },
  { key: 'licence', label: 'Licence & pricing', icon: 'Certificate', adminOnly: true },
  { key: 'audit', label: 'Audit log', icon: 'History' }
];

export function AdminPage(props: { section?: string }): JSX.Element {
  const { config, saveConfig, navigate, isAdmin } = useApp();
  const sections = SECTIONS.filter((s) => isAdmin || !s.adminOnly);
  const section = sections.filter((s) => s.key === props.section)[0] || sections[0];
  const [draft, setDraft] = React.useState<IConfig>(config);
  const [busy, run] = useBusy();
  React.useEffect(() => setDraft(config), [config]);
  const dirty = JSON.stringify(draft) !== JSON.stringify(config);
  const editsConfig = ['general', 'types', 'workflows', 'calendars', 'templates', 'rules', 'm365', 'notifications', 'integrations', 'access'].indexOf(section.key) >= 0;
  const sp: ISectionProps = { draft, setDraft };
  let body: JSX.Element;
  switch (section.key) {
    case 'types':
      body = <LeaveTypesAdmin {...sp} />;
      break;
    case 'workflows':
      body = <WorkflowsAdmin {...sp} />;
      break;
    case 'calendars':
      body = <CalendarsAdmin {...sp} />;
      break;
    case 'templates':
      body = <TemplatesAdmin {...sp} />;
      break;
    case 'rules':
      body = <RulesAdmin {...sp} />;
      break;
    case 'm365':
      body = <M365Admin {...sp} />;
      break;
    case 'delegation':
      body = <DelegationCard forAll />;
      break;
    case 'notifications':
      body = <NotificationsAdmin {...sp} />;
      break;
    case 'integrations':
      body = <IntegrationsAdmin {...sp} />;
      break;
    case 'access':
      body = <AccessAdmin {...sp} />;
      break;
    case 'licence':
      body = <LicenseAdmin />;
      break;
    case 'audit':
      body = <AuditLog />;
      break;
    default:
      body = <GeneralAdmin {...sp} />;
  }
  return (
    <>
      <PageHeader
        eyebrow={isAdmin ? 'Admin' : 'HR settings'}
        title={section.label}
        actions={
          editsConfig ? (
            <>
              <DefaultButton text="Discard" disabled={!dirty || busy} onClick={() => setDraft(config)} />
              <PrimaryButton text={busy ? 'Saving…' : 'Save changes'} disabled={!dirty || busy} onClick={() => void run(() => saveConfig(draft, section.label), 'Settings saved.')} />
            </>
          ) : undefined
        }
      />
      <div className="to-admin">
        <div className="to-admin-nav">
          {sections.map((s) => (
            <button key={s.key} className={'to-nav-item' + (s.key === section.key ? ' is-active' : '')} onClick={() => navigate(`admin/${s.key}`)}>
              <Icon iconName={s.icon} />
              <span className="to-nav-label">{s.label}</span>
            </button>
          ))}
        </div>
        <div className="to-stack" style={{ minWidth: 0 }}>
          {dirty && <MessageBar messageBarType={MessageBarType.warning}>You have unsaved changes.</MessageBar>}
          {body}
        </div>
      </div>
    </>
  );
}

function GeneralAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { branding, saveLogo, toast } = useApp();
  const set = (patch: Partial<IConfig>): void => setDraft({ ...draft, ...patch });
  const onLogo = (file: File): void => {
    if (file.size > 300 * 1024) {
      toast('Choose an image under 300 KB.', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => void saveLogo(String(reader.result)).then(() => toast('Logo saved.', 'success'));
    reader.readAsDataURL(file);
  };
  return (
    <>
      <Card title="Product & organisation" icon="CityNext">
        <div className="to-form-grid">
          <TextField label="App name" value={draft.appName} onChange={(_, v) => set({ appName: v || '' })} description="Shown in the navigation, emails and notifications." />
          <TextField label="Organisation" value={draft.orgName} onChange={(_, v) => set({ orgName: v || '' })} />
          <TextField className="is-wide" label="Home headline" value={draft.headline} onChange={(_, v) => set({ headline: v || '' })} description="The last word is highlighted." />
          <Dropdown label="Leave year starts in" selectedKey={draft.yearStartMonth} options={MONTH_NAMES.map((m, i) => ({ key: i + 1, text: m }))} onChange={(_, o) => o && set({ yearStartMonth: Number(o.key) })} />
          <Dropdown label="Week starts on" selectedKey={draft.weekStartsOn} options={[{ key: 1, text: 'Monday' }, { key: 0, text: 'Sunday' }]} onChange={(_, o) => o && set({ weekStartsOn: Number(o.key) as 0 | 1 })} />
          <TextField label="Levels / grades" className="is-wide" value={draft.grades.join(', ')} onChange={(_, v) => set({ grades: (v || '').split(',').map((g) => g.trim()).filter((g) => !!g) })} description="Comma separated. Used by grade-based entitlements and eligibility." />
        </div>
      </Card>
      <Card title="Theme, brand & logo" icon="Color">
        <div className="to-form-grid">
          <div>
            <div className="to-label">Brand colour</div>
            <div className="to-row">
              <input type="color" value={draft.brandColor} onChange={(e) => set({ brandColor: e.target.value })} aria-label="Brand colour" className="to-color-input" />
              <TextField value={draft.brandColor} onChange={(_, v) => set({ brandColor: v || DEFAULT_BRAND })} styles={{ root: { width: 120 } }} />
              <DefaultButton text="Reset" onClick={() => set({ brandColor: DEFAULT_BRAND })} />
            </div>
          </div>
          <div>
            <div className="to-label">Accent colour (highlights)</div>
            <div className="to-row">
              <input type="color" value={draft.accentColor} onChange={(e) => set({ accentColor: e.target.value })} aria-label="Accent colour" className="to-color-input" />
              <TextField value={draft.accentColor} onChange={(_, v) => set({ accentColor: v || DEFAULT_ACCENT })} styles={{ root: { width: 120 } }} />
              <DefaultButton text="Reset" onClick={() => set({ accentColor: DEFAULT_ACCENT })} />
            </div>
          </div>
          <div className="is-wide">
            <div className="to-label">Logo</div>
            <div className="to-row">
              <img src={branding.logo} alt="" className="to-logo-preview" />
              <label className="ms-Button ms-Button--default to-upload">
                <input type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" style={{ display: 'none' }} onChange={(e) => e.target.files && e.target.files[0] && onLogo(e.target.files[0])} />
                Upload logo
              </label>
              {branding.custom && <DefaultButton text="Use the default" onClick={() => void saveLogo('').then(() => toast('Default logo restored.', 'success'))} />}
            </div>
            <div className="to-small to-muted" style={{ marginTop: 4 }}>A square PNG or SVG works best (shown at 30 px). Saved at once.</div>
          </div>
        </div>
        <div className="to-brand-preview" style={{ '--to-brand': draft.brandColor, '--to-lime': draft.accentColor } as unknown as React.CSSProperties}>
          <span className="to-brand-preview-bar">
            <img src={branding.logo} alt="" /> {draft.appName}
          </span>
          <span className="to-brand-preview-btn">Request time off</span>
          <span className="to-underline">highlight</span>
        </div>
      </Card>
      <Card title="Requests" icon="Settings">
        <div className="to-stack">
          <Toggle label="Managers may submit leave on behalf of the people below them" checked={draft.onBehalf.managers} onChange={(_, c) => set({ onBehalf: { ...draft.onBehalf, managers: !!c } })} />
          <Toggle label="HR may submit leave and decide approvals on behalf of anyone" checked={draft.onBehalf.hr} onChange={(_, c) => set({ onBehalf: { ...draft.onBehalf, hr: !!c } })} />
          <Toggle label="Pending requests can be withdrawn without approval" checked={draft.cancellation.pendingDirect} onChange={(_, c) => set({ cancellation: { ...draft.cancellation, pendingDirect: !!c } })} />
          <Toggle label="Cancelling approved leave needs the approvers' confirmation" checked={draft.cancellation.approvedNeedsApproval} onChange={(_, c) => set({ cancellation: { ...draft.cancellation, approvedNeedsApproval: !!c } })} />
          <Toggle label="Employees may cancel leave that has already started" checked={draft.cancellation.allowPast} onChange={(_, c) => set({ cancellation: { ...draft.cancellation, allowPast: !!c } })} />
          <TextField label="Hour-based leave lengths (hours)" value={draft.hourOptions.join(', ')} onChange={(_, v) => set({ hourOptions: (v || '').split(',').map((x) => Number(x.trim())).filter((x) => x > 0 && x < 24) })} description="e.g. 1, 2, 3, 4 — offered for hourly and two-hour leave." />
        </div>
      </Card>
      <Card title="Comp off / in lieu" icon="Switch">
        <div className="to-stack">
          <Toggle label="Employees may claim comp off for extra days worked" checked={draft.compOff.enabled} onChange={(_, c) => set({ compOff: { ...draft.compOff, enabled: !!c } })} />
          <Dropdown label="Credited to" selectedKey={draft.compOff.leaveType} options={draft.leaveTypes.filter((t) => t.entitlement.mode === 'earned').map((t) => ({ key: t.key, text: t.name }))} onChange={(_, o) => o && set({ compOff: { ...draft.compOff, leaveType: String(o.key) } })} />
          <Toggle label="Only for work on weekends, non-working days and holidays" checked={draft.compOff.requireNonWorkingDay} onChange={(_, c) => set({ compOff: { ...draft.compOff, requireNonWorkingDay: !!c } })} />
          <NumberBox label="Most days per claim (0 = no limit)" min={0} max={30} value={String(draft.compOff.maxDaysPerClaim)} onChange={(_, v) => set({ compOff: { ...draft.compOff, maxDaysPerClaim: Number(v) || 0 } })} />
          <div className="to-small to-muted">How long credits stay valid is set on the comp-off leave type (Leave types › expiry).</div>
        </div>
      </Card>
      <Card title="Team calendar & Outlook" icon="CalendarWeek">
        <div className="to-stack">
          <Dropdown label="Employees see" selectedKey={draft.teamCalendar.scope} options={[{ key: 'team', text: 'Their team' }, { key: 'department', text: 'Their department' }, { key: 'everyone', text: 'Everyone' }]} onChange={(_, o) => o && set({ teamCalendar: { ...draft.teamCalendar, scope: o.key as 'team' } })} />
          <Toggle label="Employees see the leave type of colleagues (otherwise “Away”)" checked={draft.teamCalendar.showLeaveType} onChange={(_, c) => set({ teamCalendar: { ...draft.teamCalendar, showLeaveType: !!c } })} />
          <Toggle label="Add approved leave to each person's Outlook / Teams calendar" checked={draft.calendarSync.outlook} onChange={(_, c) => set({ calendarSync: { ...draft.calendarSync, outlook: !!c } })} />
          <Dropdown label="Show as" selectedKey={draft.calendarSync.showAs} options={[{ key: 'oof', text: 'Out of office' }, { key: 'busy', text: 'Busy' }, { key: 'free', text: 'Free' }]} onChange={(_, o) => o && set({ calendarSync: { ...draft.calendarSync, showAs: o.key as 'oof' } })} />
          <div className="to-small to-muted">Needs the Microsoft Graph permission Calendars.ReadWrite (SharePoint admin center › API access). The event is added the next time the person opens the app.</div>
        </div>
      </Card>
    </>
  );
}

function NotificationsAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { svc } = useApp();
  const [stats, setStats] = React.useState<{ queued: number; sent: number; failed: number } | undefined>();
  React.useEffect(() => {
    svc.emailQueueStats().then(setStats).catch(() => undefined);
  }, [svc]);
  const r = draft.reminders;
  const setR = (patch: Partial<IConfig['reminders']>): void => setDraft({ ...draft, reminders: { ...r, ...patch } });
  return (
    <>
      <Card title="Notifications" icon="Ringer">
        <div className="to-stack">
          <Toggle label="In-app notifications (the bell)" checked={draft.notifications.inApp} onChange={(_, c) => setDraft({ ...draft, notifications: { ...draft.notifications, inApp: !!c } })} />
          <Toggle label="Email notifications (sent by a Power Automate flow)" checked={draft.notifications.email} onChange={(_, c) => setDraft({ ...draft, notifications: { ...draft.notifications, email: !!c } })} />
          <Toggle label="Tell HR about every approved leave" checked={draft.notifications.notifyHr} onChange={(_, c) => setDraft({ ...draft, notifications: { ...draft.notifications, notifyHr: !!c } })} />
          <TextField label="Copy approved-leave emails to" value={draft.notifications.ccEmails} onChange={(_, v) => setDraft({ ...draft, notifications: { ...draft.notifications, ccEmails: v || '' } })} placeholder="payroll@contoso.com; team-leads@contoso.com" />
          {stats && (
            <div className="to-small to-muted">
              Email queue: {stats.queued} waiting, {stats.sent} sent, {stats.failed} failed.
            </div>
          )}
        </div>
      </Card>
      <Card title="Reminders to approvers & escalation" icon="Clock">
        <div className="to-stack">
          <Toggle label="Remind approvers about requests waiting for them" checked={r.enabled} onChange={(_, c) => setR({ enabled: !!c })} />
          <div className="to-form-grid">
            <NumberBox label="First reminder after (hours)" min={1} max={720} value={String(r.afterHours)} onChange={(_, v) => setR({ afterHours: Number(v) || 24 })} />
            <NumberBox label="Repeat every (hours)" min={1} max={720} value={String(r.repeatHours)} onChange={(_, v) => setR({ repeatHours: Number(v) || 24 })} />
            <NumberBox label="At most (reminders)" min={1} max={20} value={String(r.maxReminders)} onChange={(_, v) => setR({ maxReminders: Number(v) || 3 })} />
            <NumberBox label="Escalate to HR after (hours, 0 = never)" min={0} max={2000} value={String(r.escalateAfterHours)} onChange={(_, v) => setR({ escalateAfterHours: Number(v) || 0 })} />
            <NumberBox label="Auto-approve after (hours without a decision, 0 = never)" min={0} max={2000} value={String(draft.autoApproveAfterHours)} onChange={(_, v) => setDraft({ ...draft, autoApproveAfterHours: Number(v) || 0 })} />
          </div>
          <div className="to-small to-muted">
            Reminders run at most once an hour, whenever someone opens the app. For reminders on a fixed schedule, add the scheduled flow described under Power Automate & API.
          </div>
        </div>
      </Card>
      <Card title="Send the emails with Power Automate" icon="Mail">
        <FlowPackage />
        <div className="to-small to-muted" style={{ margin: '12px 0 8px' }}>Or build it by hand:</div>
        <FlowSteps
          steps={[
            <>
              Create an <b>Automated cloud flow</b> with the trigger <b>SharePoint › When an item is created</b>, site <code>{svc.sp.webUrl}</code>, list <code>{svc.title('emailQueue')}</code>.
            </>,
            <>
              Add <b>Office 365 Outlook › Send an email (V2)</b> (or “Send an email from a shared mailbox”): To = <code>MailTo</code>, CC = <code>MailCc</code>, Subject = <code>Title</code>, Body = <code>MailBody</code> (switch the body to HTML).
            </>,
            <>
              Add <b>SharePoint › Update item</b> on the same list: Id = <code>ID</code>, MailStatus = <code>Sent</code>. In the step’s settings, configure “run after has failed” to set <code>Failed</code> and <code>MailError</code>.
            </>
          ]}
        />
      </Card>
    </>
  );
}

/** A ready-made flow for the email queue: Power Automate › My flows › Import › Import package (legacy). */
function FlowPackage(): JSX.Element {
  const { sp, svc, config, toast } = useApp();
  const [shared, setShared] = React.useState('');
  const [busy, run] = useBusy();
  const make = (): Promise<boolean> =>
    run(async () => {
      const list = await sp.get<{ Id: string }>(`${sp.list(svc.title('emailQueue'))}?$select=Id`);
      const zip = flowPackageZip({ siteUrl: sp.webUrl, listId: list.Id, listTitle: svc.title('emailQueue'), displayName: `${config.appName} – send emails`, sharedMailbox: shared.trim(), allowedDomains: [], useReplyTo: false });
      downloadBytes('time-off-manager-email-flow.zip', zip, 'application/zip');
      toast('Flow package downloaded. Import it in Power Automate and pick your connections.', 'success');
    });
  return (
    <div className="to-row" style={{ alignItems: 'flex-end' }}>
      <TextField label="Send from a shared mailbox (optional)" value={shared} onChange={(_, v) => setShared(v || '')} placeholder="hr@contoso.com" styles={{ root: { minWidth: 260 } }} />
      <PrimaryButton iconProps={{ iconName: 'Download' }} text="Download the flow package" disabled={busy} onClick={() => void make()} />
    </div>
  );
}

function FlowSteps(props: { steps: React.ReactNode[] }): JSX.Element {
  return (
    <div className="to-stack">
      {props.steps.map((s, i) => (
        <div key={i} className="to-flow-step">
          <span className="to-flow-num">{i + 1}</span>
          <div>{s}</div>
        </div>
      ))}
    </div>
  );
}

function IntegrationsAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { svc, sp } = useApp();
  const api = (key: Parameters<typeof svc.title>[0]): string => `${sp.webUrl}/_api/web/lists/getbytitle('${svc.title(key)}')/items`;
  return (
    <>
      <Card title="Events for Power Automate" icon="Flow">
        <div className="to-stack">
          <Toggle label="Write every change of a request to the Events list (outbox)" checked={draft.integrations.outbox} onChange={(_, c) => setDraft({ ...draft, integrations: { ...draft.integrations, outbox: !!c } })} />
          <p className="to-muted" style={{ margin: 0 }}>
            Each row has <code>EventType</code> (submitted, stepApproved, approved, rejected, cancelled) and a JSON <code>Payload</code> with the request. Start a flow with <b>When an item is created</b> on <code>{svc.title('events')}</code>,
            then <b>Parse JSON</b> on Payload — for example to post to a Teams channel, update payroll, or create a shift in Shifts.
          </p>
          <FlowSteps
            steps={[
              <>Scheduled reminders: a <b>Recurrence</b> flow (e.g. 09:00 on weekdays) → <b>Get items</b> on <code>{svc.title('requests')}</code> with filter <code>{"Status eq 'Pending'"}</code> → email the addresses in <code>CurrentApprovers</code>.</>,
              <>Payroll export: on <code>EventType = approved</code> where the leave type is unpaid, add a row to an Excel table or call your payroll API.</>,
              <>Teams post: on <code>approved</code>, <b>Post message in a chat or channel</b> with the employee name and dates.</>
            ]}
          />
        </div>
      </Card>
      <Card title="Webhook / API connectivity" icon="PlugConnected">
        <div className="to-stack">
          <TextField
            label="Webhook address (HTTPS)"
            value={draft.integrations.webhookUrl}
            onChange={(_, v) => setDraft({ ...draft, integrations: { ...draft.integrations, webhookUrl: (v || '').trim() } })}
            placeholder="https://prod-00.westeurope.logic.azure.com/workflows/…"
            description="Every event is POSTed here as JSON (Content-Type text/plain). Works with Power Automate “When an HTTP request is received”, Azure Functions, Logic Apps and HR systems."
            errorMessage={draft.integrations.webhookUrl && !/^https:\/\//i.test(draft.integrations.webhookUrl) ? 'Use an https:// address.' : undefined}
          />
          <div className="to-label">SharePoint REST API (read and write with Microsoft Entra ID authentication)</div>
          {(['requests', 'employees', 'ledger', 'delegations', 'events'] as const).map((k) => (
            <TextField key={k} readOnly value={api(k)} prefix={LISTS.filter((l) => l.key === k)[0].name} />
          ))}
          <div className="to-small to-muted">Microsoft Graph works too: <code>/sites/{'{site-id}'}/lists/{svc.title('requests')}/items?expand=fields</code>.</div>
        </div>
      </Card>
    </>
  );
}

function AccessAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { sp, svc, tenantId, settings } = useApp();
  const [busy, run] = useBusy();
  const [log, setLog] = React.useState<string[]>([]);
  const applied = settings.get(SETTING_KEYS.securityAppliedAt);
  const prov = new Provisioner(sp, svc.prefix);
  return (
    <>
      <Card title="Roles" icon="People">
        <div className="to-stack">
          <PeopleField label="App admins" value={draft.admins} onChange={(v) => setDraft({ ...draft, admins: v })} description="Everything, including access and the licence. Site owners are always admins." />
          <PeopleField label="HR" value={draft.hr} onChange={(v) => setDraft({ ...draft, hr: v })} description="Manage employees, balances, leave types and reports; approve at HR levels and on behalf of others." />
          <div className="to-small to-muted">Managers are found from the employee list (people whose manager they are) and from approval workflows.</div>
        </div>
      </Card>
      <Card title="Department heads" icon="Org">
        <div className="to-stack">
          {Object.keys(draft.departmentHeads).length === 0 && <div className="to-small to-muted">Not set: the most senior manager of each department is used.</div>}
          {svc && <DeptHeads {...props} />}
        </div>
      </Card>
      <Card title="SharePoint lists" icon="Permissions">
        <p className="to-muted" style={{ marginTop: 0 }}>
          The app keeps its data in {LISTS.length} lists on this site ({LISTS.map((l) => listTitle(svc.prefix, l.key)).join(', ')}). Recommended permissions: everyone may read the settings, employees and audit lists and contribute to the rest; HR and app admins contribute everywhere.
        </p>
        <div className="to-row">
          <DefaultButton text="Validate & repair lists" disabled={busy} onClick={() => void run(async () => {
            const out: string[] = [];
            const res = await prov.ensureSchema((m) => out.push(m));
            if (res.ok) {
              await settings.set(SETTING_KEYS.schemaVersion, String(SCHEMA_VERSION));
            }
            setLog(out.concat(res.errors));
          }, 'Lists checked.')} />
          <PrimaryButton text="Apply recommended permissions" disabled={busy} onClick={() => void run(async () => {
            const out: string[] = [];
            const errors = await prov.applySecurity(tenantId, draft.admins.concat(draft.hr), (m) => out.push(m));
            await settings.set(SETTING_KEYS.securityAppliedAt, new Date().toISOString());
            setLog(out.concat(errors));
          }, 'Permissions applied.')} />
          {applied && <span className="to-small to-muted">Last applied {relativeTime(applied)}</span>}
        </div>
        {log.length > 0 && <pre className="to-log">{log.join('\n')}</pre>}
      </Card>
      <SampleData />
    </>
  );
}

function SampleData(): JSX.Element | null {
  const { world, engine, reload, me } = useApp();
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const sample = world.employees.filter((e) => /@example\.com$/i.test(e.email)).length;
  if (!sample) return null;
  return (
    <Card title="Sample organisation" icon="Delete">
      {dialog}
      <p className="to-muted" style={{ marginTop: 0 }}>{sample} sample people (@example.com) with their requests and balances are in the lists.</p>
      <DefaultButton
        text="Remove the sample data"
        disabled={busy}
        onClick={async () => {
          if (!(await ask('Remove the sample organisation?', 'Sample people, their requests and ledger entries are deleted. Your own data stays.', 'Remove'))) return;
          await run(async () => {
            await engine.removeSampleData(world, { email: me.email, name: me.title });
            await reload();
          }, 'Sample data removed.');
        }}
      />
    </Card>
  );
}

function DeptHeads(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { world } = useApp();
  const departments = world.employees.map((e) => e.department || '').filter((d, i, all) => d && all.indexOf(d) === i).sort();
  return (
    <div className="to-stack">
      {departments.map((d) => (
        <PeopleField
          key={d}
          single
          label={d}
          value={draft.departmentHeads[d] ? [draft.departmentHeads[d]] : []}
          onChange={(v) => {
            const next = { ...draft.departmentHeads };
            if (v[0]) next[d] = v[0];
            else delete next[d];
            setDraft({ ...draft, departmentHeads: next });
          }}
        />
      ))}
    </div>
  );
}

function AuditLog(): JSX.Element {
  const { svc } = useApp();
  const [rows, setRows] = React.useState<Array<{ id: number; action: string; actor: string; target: string; details: string; created: string }>>([]);
  React.useEffect(() => {
    svc.loadAudit().then(setRows).catch(() => undefined);
  }, [svc]);
  return (
    <Card title="Latest 200 changes" icon="History">
      <div className="to-table-wrap">
        <table className="to-table">
          <thead>
            <tr>
              <th>When</th>
              <th>What</th>
              <th>Who</th>
              <th>Target</th>
              <th>Details</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id}>
                <td className="to-small">{relativeTime(r.created)}</td>
                <td>{r.action}</td>
                <td>{r.actor}</td>
                <td>{r.target}</td>
                <td className="to-small to-muted">{r.details}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}
