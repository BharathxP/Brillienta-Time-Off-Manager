import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, PrimaryButton, SpinButton, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Card, Kpi, useBusy } from '../common/Ui';
import { PRICING, checkLicenseKey, describeStatus, quote } from '../../services/licensing/license';
import { LICENSE_PUBLIC_KEY } from '../../services/licensing/publicKey';
import { SETTING_KEYS } from '../../services/SettingsStore';
import { fmtDate } from '../../services/dates';

const INCLUDED = [
  'Comp off / in lieu of leaves',
  'Level / grade based leaves',
  'Tenure based leaves',
  'Accrual based leaves',
  'Custom theme, brand & logo',
  'Non-working days exclusion',
  'Submit leaves on behalf of users',
  'Half-day & two-hour leaves',
  'Hourly leave approval',
  'Approval delegation',
  'Reminder notifications to approvers',
  'Leave cancellation',
  'Auto approvals',
  'Approve on behalf of manager',
  'Leave templates',
  'Recurrence & bi-weekly leaves',
  'Teams calendar',
  'Multilevel approval workflow',
  '30 types of leave',
  'Customize with Power Automate',
  'API connectivity',
  'Connect with Power BI',
  'Free updates with support',
  'Unlimited support'
];

export function LicenseAdmin(): JSX.Element {
  const { license, tenantId, settings, refreshLicense, world, licensed, me, svc } = useApp();
  const [key, setKey] = React.useState('');
  const [busy, run] = useBusy();
  const active = world.employees.filter((e) => e.active).length;
  const [users, setUsers] = React.useState(Math.max(PRICING.baseUsers, active));
  const q = quote(users);
  const activate = (): Promise<boolean> =>
    run(async () => {
      const payload = await checkLicenseKey(key, tenantId, LICENSE_PUBLIC_KEY);
      if (payload.expires && Date.parse(payload.expires + 'T23:59:59Z') < Date.now()) throw new Error(`This licence ended on ${payload.expires}.`);
      await settings.set(SETTING_KEYS.licenseKey, key.trim());
      await svc.audit('Licence activated', { email: me.email, name: me.title }, payload.id, `${payload.users} users`);
      await refreshLicense();
      setKey('');
    }, 'Licence activated.');
  const mail = `mailto:${PRICING.contact}?subject=${encodeURIComponent('Brillienta Time Off Manager licence')}&body=${encodeURIComponent(`Tenant id: ${tenantId || ''}\nUsers: ${users}\n`)}`;
  return (
    <>
      <MessageBar messageBarType={license.state === 'licensed' ? MessageBarType.success : license.state === 'trial' ? MessageBarType.info : MessageBarType.warning}>{describeStatus(license)}</MessageBar>
      {license.keyProblem && <MessageBar messageBarType={MessageBarType.error}>The stored licence key was not accepted: {license.keyProblem}</MessageBar>}
      <div className="to-kpis">
        <Kpi label="Active employees" value={active} icon="People" tint={1} />
        <Kpi label="Licensed users" value={license.users || 'Unlimited (trial)'} icon="Permissions" tint={2} />
        <Kpi label="Over the licence" value={Math.max(0, active - licensed.size)} icon="Warning" tint={3} tone={active > licensed.size ? 'danger' : undefined} />
        {license.license && <Kpi label="Customer" value={license.license.customer} icon="CityNext" tint={4} />}
      </div>
      <div className="to-grid to-grid-2">
        <Card title="Activate a licence key" icon="Certificate">
          <div className="to-stack">
            <TextField label="Your Microsoft 365 tenant id (send it with your order)" readOnly value={tenantId || 'Not available'} />
            <TextField label="Licence key" multiline rows={4} value={key} onChange={(_, v) => setKey(v || '')} placeholder="TO1.…" />
            <PrimaryButton text="Activate" disabled={busy || !key.trim()} onClick={() => void activate()} />
            {license.license && (
              <div className="to-small to-muted">
                Licence {license.license.id} · issued {fmtDate(license.license.issued)} · {license.license.expires ? `valid until ${fmtDate(license.license.expires)}` : 'perpetual'}
              </div>
            )}
          </div>
        </Card>
        <Card title="Pricing" icon="Money">
          <div className="to-stack">
            <div className="to-price-lines">
              <div>
                <b>{PRICING.baseUsers} users</b> included in the base plan
              </div>
              <div>
                <b>Add-on ${PRICING.addOnPackUsd}</b> for every {PRICING.addOnPackUsers} users
              </div>
              <div>
                <b>Additional users at ₹{PRICING.extraUserInr}/user</b>
              </div>
              <div>Free updates with support · Unlimited support</div>
            </div>
            <SpinButton label="Users you need" min={1} max={100000} step={10} value={String(users)} onChange={(_, v) => setUsers(Number(v) || 1)} />
            <div className="to-calc">
              <div>
                <span>Users above the base plan</span>
                <b>{q.extra}</b>
              </div>
              <div>
                <span>Add-on packs of {PRICING.addOnPackUsers} (USD)</span>
                <b>
                  {q.packs} × ${PRICING.addOnPackUsd} = ${q.usd}
                </b>
              </div>
              <div>
                <span>Or per additional user (INR)</span>
                <b>₹{q.inr.toLocaleString('en-IN')}</b>
              </div>
            </div>
            <DefaultButton iconProps={{ iconName: 'Mail' }} text="Request a quote" href={mail} />
          </div>
        </Card>
      </div>
      <Card title="Every plan includes" icon="CheckList">
        <div className="to-included">
          {INCLUDED.map((f) => (
            <span key={f}>
              <Icon iconName="CheckMark" /> {f}
            </span>
          ))}
        </div>
      </Card>
    </>
  );
}

/** Shown instead of the app when the trial or licence has ended. */
export function LicenseLocked(): JSX.Element {
  const { license, isAdmin, config, branding } = useApp();
  return (
    <div className="to-locked">
      <img src={branding.logo} alt="" className="to-logo-preview" />
      <h1>{config.appName}</h1>
      <p className="to-muted">{describeStatus(license)}. Your data is safe in SharePoint and comes back as soon as a licence is activated.</p>
      {isAdmin ? <LicenseAdmin /> : <p>Ask your Time Off Manager admin to renew the licence.</p>}
    </div>
  );
}
