import * as React from 'react';
import { Customizer, ILayerStyles, MessageBar, MessageBarType, PrimaryButton, ThemeProvider } from '@fluentui/react';
import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import styles from './TimeOff.module.scss';
import { AppCtx, IAppContext, IHostInfo, IRequestFormOptions, IRoute, ToastKind, errorText, parseRoute } from './AppContext';
import { IConfig, ICurrentUser, Role } from '../models/types';
import { SPClient } from '../services/SPClient';
import { SETTING_KEYS, SettingsStore } from '../services/SettingsStore';
import { LeaveService } from '../services/LeaveService';
import { Notifier } from '../services/Notifier';
import { GraphService } from '../services/GraphService';
import { Engine, IWorld } from '../services/Engine';
import { mergeConfig } from '../schema/defaults';
import { SCHEMA_VERSION, listTitle } from '../schema/schema';
import { indexEmployees, lc } from '../services/workflow';
import { leaveYearOf, today } from '../services/dates';
import { ILicenseStatus, checkLicenseKey, licenseStatus, licensedEmails } from '../services/licensing/license';
import { LICENSE_PUBLIC_KEY } from '../services/licensing/publicKey';
import { DEFAULT_LOGO } from '../assets/defaultLogo';
import { Shell } from './Shell';
import { Provisioner } from '../services/Provisioner';
import { SetupWizard } from './setup/SetupWizard';
import { LicenseLocked } from './admin/LicenseAdmin';
import { Loading } from './common/Ui';
import { NAVY, brandTheme, ensureFonts } from './theme';

export interface IAppProps {
  http: SPHttpClient;
  webUrl: string;
  siteTitle: string;
  permissions: { manageLists: boolean; manageWeb: boolean };
  tenantId?: string;
  prefix: string;
  appVersion: string;
  graphFactory?: MSGraphClientFactory;
  host?: IHostInfo;
}

type Phase = { kind: 'loading' } | { kind: 'error'; message: string } | { kind: 'setup' } | { kind: 'notInstalled' } | { kind: 'ready' };

interface IToast {
  id: number;
  message: string;
  kind: ToastKind;
}

const EMPTY_WORLD = (config: IConfig): IWorld => ({ config, employees: [], delegations: [], requests: [], ledger: [] });

function isSafeImage(src: string): boolean {
  return /^data:image\/(png|jpe?g|gif|svg\+xml|webp);base64,[A-Za-z0-9+/=]+$/.test(src) || /^https:\/\//.test(src);
}

export default function App(props: IAppProps): JSX.Element {
  const host: IHostInfo = props.host || { inTeams: false, view: 'app' };
  const sp = React.useMemo(() => new SPClient(props.http, props.webUrl), [props.http, props.webUrl]);
  const svc = React.useMemo(() => new LeaveService(sp, props.prefix), [sp, props.prefix]);
  const settings = React.useMemo(() => new SettingsStore(sp, props.prefix), [sp, props.prefix]);
  const graph = React.useMemo(() => (props.graphFactory ? new GraphService(props.graphFactory) : undefined), [props.graphFactory]);
  const configRef = React.useRef<IConfig>(mergeConfig(undefined));
  const pageUrl = window.location.href.split('#')[0];
  const engine = React.useMemo(() => new Engine(svc, new Notifier(svc, () => configRef.current, pageUrl), graph), [svc, graph, pageUrl]);

  const [phase, setPhase] = React.useState<Phase>({ kind: 'loading' });
  const [me, setMe] = React.useState<ICurrentUser | undefined>();
  const [config, setConfig] = React.useState<IConfig>(() => mergeConfig(undefined));
  const [world, setWorld] = React.useState<IWorld>(() => EMPTY_WORLD(mergeConfig(undefined)));
  const [license, setLicense] = React.useState<ILicenseStatus | undefined>();
  const [logo, setLogo] = React.useState('');
  const [route, setRoute] = React.useState<IRoute>(() => parseRoute(window.location.hash));
  const [toasts, setToasts] = React.useState<IToast[]>([]);
  const [requestForm, setRequestForm] = React.useState<IRequestFormOptions | undefined>();
  const [notificationsTick, setNotificationsTick] = React.useState(0);
  const [boot, setBoot] = React.useState(0);
  const housekeeping = React.useRef(false);

  configRef.current = config;
  React.useEffect(() => ensureFonts(), []);

  const toast = React.useCallback((message: string, kind: ToastKind = 'info') => {
    const id = Date.now() + Math.random();
    setToasts((t) => t.concat({ id, message, kind }).slice(-4));
    window.setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), kind === 'error' ? 8000 : 4000);
  }, []);

  React.useEffect(() => {
    const onHash = (): void => setRoute(parseRoute(window.location.hash));
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);

  const navigate = React.useCallback((path: string) => {
    window.location.hash = '#/' + path.replace(/^#?\/?/, '');
    const main = document.querySelector('[data-to-main]');
    if (main) main.scrollTop = 0;
  }, []);

  const loadLicense = React.useCallback(async (): Promise<ILicenseStatus> => {
    const now = await sp.serverTime();
    let created: string | undefined;
    try {
      created = (await sp.get<{ Created: string }>(`${sp.list(listTitle(props.prefix, 'settings'))}?$select=Created`)).Created;
    } catch {
      created = undefined;
    }
    const candidates = [settings.get(SETTING_KEYS.installedAt), created].map((d) => (d ? Date.parse(d) : NaN)).filter((t) => !isNaN(t));
    const trialStart = new Date(candidates.length ? Math.min(...candidates) : now.getTime());
    const key = settings.get(SETTING_KEYS.licenseKey);
    let payload;
    let problem: string | undefined;
    if (key) {
      try {
        payload = await checkLicenseKey(key, props.tenantId, LICENSE_PUBLIC_KEY);
      } catch (e) {
        problem = errorText(e);
      }
    }
    return licenseStatus(payload, trialStart, now, problem);
  }, [sp, settings, props.prefix, props.tenantId]);

  const loadWorld = React.useCallback(
    async (cfg: IConfig): Promise<IWorld> => {
      const year = leaveYearOf(today(), cfg.yearStartMonth);
      const [employees, requests, ledger, delegations] = await Promise.all([
        svc.loadEmployees(),
        svc.loadRequests(year - 1),
        svc.loadLedger(year - 1).catch(() => []),
        svc.loadDelegations().catch(() => [])
      ]);
      return { config: cfg, employees, requests, ledger, delegations };
    },
    [svc]
  );

  React.useEffect(() => {
    let alive = true;
    (async () => {
      setPhase({ kind: 'loading' });
      const u = await sp.get<{ Id: number; Title: string; Email: string; LoginName: string; IsSiteAdmin: boolean }>('web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin');
      const user: ICurrentUser = { id: u.Id, title: u.Title, email: u.Email, loginName: u.LoginName, isSiteAdmin: u.IsSiteAdmin, canManageWeb: props.permissions.manageWeb || u.IsSiteAdmin };
      if (!alive) return;
      setMe(user);
      const installed = await settings.load();
      const version = Number(settings.get(SETTING_KEYS.schemaVersion) || 0);
      if (!installed || version < 1) {
        setPhase({ kind: props.permissions.manageLists ? 'setup' : 'notInstalled' });
        return;
      }
      svc.schemaVersion = version;
      if (version < SCHEMA_VERSION && props.permissions.manageLists) {
        // An update added columns: a site owner's visit adds them (until then the app writes without them).
        try {
          const res = await new Provisioner(sp, props.prefix).ensureSchema(() => undefined);
          if (res.ok) {
            await settings.set(SETTING_KEYS.schemaVersion, String(SCHEMA_VERSION));
            await settings.set(SETTING_KEYS.appVersion, props.appVersion);
            svc.schemaVersion = SCHEMA_VERSION;
          }
        } catch {
          // stays on the earlier columns; Admin › Access › Validate & repair can retry
        }
      }
      const cfg = mergeConfig(settings.getJson(SETTING_KEYS.config), props.siteTitle);
      const [w, lic] = await Promise.all([loadWorld(cfg), loadLicense()]);
      if (!alive) return;
      configRef.current = cfg;
      setConfig(cfg);
      setWorld(w);
      setLicense(lic);
      setLogo(settings.get(SETTING_KEYS.brandLogo) || '');
      setPhase({ kind: 'ready' });
    })().catch((e) => alive && setPhase({ kind: 'error', message: errorText(e) }));
    return () => {
      alive = false;
    };
  }, [sp, settings, boot]);

  const reload = React.useCallback(async () => {
    const w = await loadWorld(configRef.current);
    setWorld(w);
  }, [loadWorld]);

  const saveConfig = React.useCallback(
    async (next: IConfig, what?: string) => {
      await settings.setJson(SETTING_KEYS.config, next);
      configRef.current = next;
      setConfig(next);
      setWorld((w) => ({ ...w, config: next }));
      if (me) void svc.audit('Settings changed', { email: me.email, name: me.title }, 'Config', what || 'Settings saved');
    },
    [settings, svc, me]
  );

  const saveLogo = React.useCallback(
    async (value: string) => {
      await settings.set(SETTING_KEYS.brandLogo, value);
      setLogo(value);
    },
    [settings]
  );

  const refreshLicense = React.useCallback(async () => {
    await settings.load();
    setLicense(await loadLicense());
  }, [settings, loadLicense]);

  // Roles: admins (site owners or listed), HR (listed), managers (people report to them or they approve).
  const myEmployee = me ? world.employees.filter((e) => lc(e.email) === lc(me.email))[0] : undefined;
  const isListed = (list: Array<{ email: string }>): boolean => !!me && list.some((p) => lc(p.email) === lc(me.email));
  const isAdmin = !!me && (me.canManageWeb || isListed(config.admins));
  const isHr = isAdmin || isListed(config.hr);
  const isManager =
    !!me &&
    (world.employees.some((e) => lc(e.managerEmail) === lc(me.email)) ||
      world.requests.some((r) => r.steps.some((s) => s.approvers.some((a) => lc(a.email) === lc(me.email)))) ||
      world.delegations.some((d) => d.active && lc(d.delegate.email) === lc(me.email)));
  const role: Role = isAdmin ? 'admin' : isHr ? 'hr' : isManager ? 'manager' : 'employee';

  // Housekeeping once per load, by whoever opens the app: approver reminders, escalation, auto-approval and
  // delegation changes (at most hourly across all browsers), the new year's carry-forward, and the person's own
  // Outlook calendar.
  React.useEffect(() => {
    if (phase.kind !== 'ready' || !me || housekeeping.current || !license || !license.active) return;
    housekeeping.current = true;
    const actor = { email: me.email, name: me.title, isHr, isAdmin };
    (async () => {
      const last = Date.parse((await settings.fresh(SETTING_KEYS.lastReminderRun).catch(() => '')) || '');
      if (isNaN(last) || Date.now() - last > 3600000) {
        await settings.set(SETTING_KEYS.lastReminderRun, new Date().toISOString());
        const res = await engine.runReminders(world, actor);
        if (res.autoApproved || res.escalated || res.redelegated) await reload();
      }
      const year = leaveYearOf(today(), config.yearStartMonth);
      const carried = settings.getJson<number[]>(SETTING_KEYS.carriedYears) || [];
      if (isHr && carried.indexOf(year - 1) < 0 && world.employees.length) {
        await settings.setJson(SETTING_KEYS.carriedYears, carried.concat(year - 1));
        const n = await engine.carryForward(world, year - 1, { email: me.email, name: me.title }, () => undefined);
        if (n) toast(`${n} balance${n === 1 ? '' : 's'} carried forward into ${year}.`, 'info');
      }
      const synced = await engine.syncMyCalendar(world, myEmployee);
      if (synced) toast(`Your Outlook calendar was updated (${synced} leave${synced === 1 ? '' : 's'}).`, 'success');
    })().catch(() => undefined);
  }, [phase.kind, me, license]);

  const brand = config.brandColor || NAVY;
  const theme = React.useMemo(() => brandTheme(brand), [brand]);
  const accent = config.accentColor || '#caf578';
  const frameStyle = React.useMemo(() => ({ '--to-brand': brand, '--to-lime': accent }) as unknown as React.CSSProperties, [brand, accent]);
  // Panels, dialogs and dropdowns render in Fluent layers at the end of <body>, outside the styled root: give every
  // layer the root class and brand variables so the to-* styles reach them too (as in the Employee Directory).
  const layerSettings = React.useMemo(() => ({ Layer: { styles: { root: [styles.root, frameStyle] } as unknown as ILayerStyles } }), [frameStyle]);

  if (phase.kind === 'loading' || !me) {
    return (
      <div className={styles.root}>
        <div className="to-app" style={{ minHeight: 320 }}>
          <Loading label="Loading Time Off Manager…" />
        </div>
      </div>
    );
  }
  if (phase.kind === 'error') {
    return (
      <div className={styles.root}>
        <div className="to-card" style={{ margin: 16 }}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            Time Off Manager could not load: {phase.message}
          </MessageBar>
          <PrimaryButton style={{ marginTop: 12 }} text="Try again" onClick={() => setBoot((b) => b + 1)} />
        </div>
      </div>
    );
  }
  if (phase.kind === 'setup' || phase.kind === 'notInstalled') {
    return (
      <ThemeProvider theme={theme}>
        <Customizer scopedSettings={layerSettings}>
        <div className={styles.root} style={frameStyle}>
          <SetupWizard sp={sp} settings={settings} svc={svc} graph={graph} prefix={props.prefix} me={me} siteTitle={props.siteTitle} canInstall={phase.kind === 'setup'} appVersion={props.appVersion} onDone={() => setBoot((b) => b + 1)} />
        </div>
        </Customizer>
      </ThemeProvider>
    );
  }

  const lic = license as ILicenseStatus;
  const active = world.employees.filter((e) => e.active);
  const ctx: IAppContext = {
    sp,
    svc,
    settings,
    engine,
    graph,
    config,
    me,
    myEmployee,
    role,
    isAdmin,
    isHr,
    isManager,
    world: { ...world, config },
    dir: indexEmployees(world.employees),
    year: leaveYearOf(today(), config.yearStartMonth),
    license: lic,
    licensed: licensedEmails(active, lic.users),
    branding: { logo: logo && isSafeImage(logo) ? logo : DEFAULT_LOGO, custom: !!logo },
    tenantId: props.tenantId,
    appVersion: props.appVersion,
    host,
    pageUrl,
    route,
    navigate,
    toast,
    reload,
    saveConfig,
    saveLogo,
    refreshLicense,
    openRequest: (o) => setRequestForm(o || {}),
    notificationsTick,
    bumpNotifications: () => setNotificationsTick((t) => t + 1)
  };

  return (
    <ThemeProvider theme={theme}>
      <Customizer scopedSettings={layerSettings}>
      <AppCtx.Provider value={ctx}>
        <div className={styles.root} style={frameStyle}>
          {!lic.active ? (
            <div className="to-app" style={{ display: 'block' }}>
              <LicenseLocked />
            </div>
          ) : (
            <Shell requestForm={requestForm} closeRequestForm={() => setRequestForm(undefined)} siteTitle={props.siteTitle} />
          )}
          <div className="to-toasts" aria-live="polite">
            {toasts.map((t) => (
              <MessageBar key={t.id} messageBarType={t.kind === 'success' ? MessageBarType.success : t.kind === 'error' ? MessageBarType.error : t.kind === 'warning' ? MessageBarType.warning : MessageBarType.info} onDismiss={() => setToasts((x) => x.filter((y) => y.id !== t.id))}>
                {t.message}
              </MessageBar>
            ))}
          </div>
        </div>
      </AppCtx.Provider>
      </Customizer>
    </ThemeProvider>
  );
}
