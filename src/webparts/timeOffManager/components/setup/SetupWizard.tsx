import * as React from 'react';
import { Checkbox, ChoiceGroup, Dropdown, Icon, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { ICurrentUser } from '../../models/types';
import { SPClient } from '../../services/SPClient';
import { SETTING_KEYS, SettingsStore } from '../../services/SettingsStore';
import { LeaveService } from '../../services/LeaveService';
import { GraphService } from '../../services/GraphService';
import { LogLevel, Provisioner } from '../../services/Provisioner';
import { SCHEMA_VERSION } from '../../schema/schema';
import { defaultConfig } from '../../schema/defaults';
import { MONTH_NAMES } from '../../services/dates';
import { sampleEmployees, sampleLedger, sampleRequests } from '../../services/sample';
import { DEFAULT_LOGO } from '../../assets/defaultLogo';

interface IProps {
  sp: SPClient;
  settings: SettingsStore;
  svc: LeaveService;
  graph?: GraphService;
  prefix: string;
  me: ICurrentUser;
  siteTitle: string;
  canInstall: boolean;
  appVersion: string;
  onDone: () => void;
}

/** First run: creates the lists, the configuration (30 leave types, workflows, calendars) and the employee list. */
export function SetupWizard(props: IProps): JSX.Element {
  const [org, setOrg] = React.useState(props.siteTitle);
  const [month, setMonth] = React.useState(1);
  const [people, setPeople] = React.useState<'graph' | 'sample' | 'later'>(props.graph ? 'graph' : 'sample');
  const [meHr, setMeHr] = React.useState(true);
  const [log, setLog] = React.useState<Array<{ m: string; l: LogLevel }>>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const add = (m: string, l: LogLevel = 'info'): void => setLog((x) => x.concat({ m, l }));

  if (!props.canInstall) {
    return (
      <div className="to-setup">
        <img src={DEFAULT_LOGO} alt="" className="to-setup-logo" />
        <h1>Brillienta Time Off Manager</h1>
        <p className="to-muted">Time Off Manager is not set up on this site yet. A site owner opens this page once to set it up.</p>
      </div>
    );
  }

  const install = async (): Promise<void> => {
    setBusy(true);
    setError('');
    try {
      const prov = new Provisioner(props.sp, props.prefix);
      add('Creating the SharePoint lists…');
      const res = await prov.ensureSchema(add);
      if (!res.ok) throw new Error(res.errors.join('; '));
      await props.settings.load();
      const cfg = defaultConfig(org);
      cfg.yearStartMonth = month;
      if (meHr) cfg.hr = [{ email: props.me.email, name: props.me.title }];
      cfg.admins = [{ email: props.me.email, name: props.me.title }];
      if (people === 'graph' && props.graph) {
        add('Importing people from Microsoft Entra ID…');
        try {
          const users = await props.graph.loadUsers((n) => n % 500 === 0 && add(`${n} people read`));
          const plan = GraphService.merge([], users);
          for (const e of plan.add) await props.svc.saveEmployee(e);
          await props.settings.set(SETTING_KEYS.lastImportAt, new Date().toISOString());
          add(`${plan.add.length} employees imported`, 'success');
        } catch (e) {
          add(`Microsoft Graph import skipped (${(e as Error).message}). Approve the API permissions, then import from Employees & balances.`, 'warning');
        }
      } else if (people === 'sample') {
        add('Adding the sample organisation…');
        const me = { email: props.me.email, name: props.me.title };
        const emps = sampleEmployees(me);
        for (const e of emps) await props.svc.saveEmployee(e);
        for (const r of sampleRequests(cfg, emps, me)) await props.svc.createRequest(r, (cfg.leaveTypes.filter((t) => t.key === r.leaveType)[0] || { name: r.leaveType }).name, (emps.filter((x) => x.email === r.employeeEmail)[0] || { department: '' }).department || '');
        for (const l of sampleLedger(cfg, emps)) await props.svc.addLedger(l);
        add(`${emps.length} sample employees with requests`, 'success');
      }
      await props.settings.setJson(SETTING_KEYS.config, cfg);
      // The year before installation has nothing to carry forward.
      const thisYear = new Date().getMonth() + 1 >= month ? new Date().getFullYear() : new Date().getFullYear() - 1;
      if (!props.settings.get(SETTING_KEYS.carriedYears)) await props.settings.setJson(SETTING_KEYS.carriedYears, [thisYear - 1]);
      if (!props.settings.get(SETTING_KEYS.installedAt)) await props.settings.set(SETTING_KEYS.installedAt, new Date().toISOString());
      await props.settings.set(SETTING_KEYS.installedBy, props.me.email);
      await props.settings.set(SETTING_KEYS.appVersion, props.appVersion);
      await props.settings.set(SETTING_KEYS.schemaVersion, String(SCHEMA_VERSION));
      add('Ready.', 'success');
      props.onDone();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="to-setup">
      <img src={DEFAULT_LOGO} alt="" className="to-setup-logo" />
      <p className="to-eyebrow">
        <span className="to-star">✳</span> Welcome
      </p>
      <h1>
        Set up Brillienta <span className="to-underline">Time Off Manager</span>
      </h1>
      <p className="to-muted">
        Everything stays on this SharePoint site: nine lists are created, with 30 leave types, approval workflows and holiday calendars you can change later. It takes about a minute. The 14-day free trial starts now.
      </p>
      <div className="to-card to-stack" style={{ textAlign: 'left' }}>
        <TextField label="Organisation name" value={org} onChange={(_, v) => setOrg(v || '')} />
        <Dropdown label="Leave year starts in" selectedKey={month} options={MONTH_NAMES.map((m, i) => ({ key: i + 1, text: m }))} onChange={(_, o) => o && setMonth(Number(o.key))} />
        <ChoiceGroup
          label="Employees"
          selectedKey={people}
          onChange={(_, o) => o && setPeople(o.key as 'graph')}
          options={[
            { key: 'graph', text: 'Import from Microsoft 365 (names, managers, departments, start dates)', disabled: !props.graph },
            { key: 'sample', text: 'Add a sample organisation to try it out (removable)' },
            { key: 'later', text: 'Later (CSV import or one by one)' }
          ]}
        />
        <Checkbox label={`Make me (${props.me.title}) HR`} checked={meHr} onChange={(_, c) => setMeHr(!!c)} />
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        <PrimaryButton text={busy ? 'Setting up…' : 'Set up Time Off Manager'} disabled={busy} onClick={() => void install()} />
        {log.length > 0 && (
          <div className="to-log">
            {log.map((x, i) => (
              <div key={i} className={`is-${x.l}`}>
                <Icon iconName={x.l === 'success' ? 'CheckMark' : x.l === 'error' ? 'Error' : x.l === 'warning' ? 'Warning' : 'Info'} /> {x.m}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
