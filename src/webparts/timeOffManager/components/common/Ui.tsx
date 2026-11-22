import * as React from 'react';
import { ComboBox, DefaultButton, Dialog, DialogFooter, DialogType, IComboBoxOption, Icon, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import { IEmployee, ILeaveType, IPerson, RequestStatus } from '../../models/types';
import { useApp } from '../AppContext';
import { lc } from '../../services/workflow';

function colorIndex(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return (Math.abs(h) % 6) + 1;
}

export function initials(name: string): string {
  const parts = (name || '?').replace(/\(.*?\)/g, '').trim().split(/\s+/);
  return ((parts[0] || '?')[0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}

/** Profile photo from SharePoint (the Microsoft 365 photo), initials while loading or when there is none. */
export function Avatar(props: { email?: string; name: string; size?: 'sm' | 'md' | 'lg' }): JSX.Element {
  const { sp } = useApp();
  const [state, setState] = React.useState<'loading' | 'ok' | 'none'>('loading');
  const size = props.size || 'md';
  const src = props.email && !/@example\.(com|org)$/i.test(props.email) ? `${sp.webUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(props.email)}` : '';
  return (
    <span className={`to-avatar to-av-${colorIndex(props.name || '?')}` + (size === 'md' ? '' : ` is-${size}`)} title={props.name}>
      <span className="to-avatar-inner">
        {state !== 'ok' && initials(props.name)}
        {src && state !== 'none' && (
          <img
            src={src}
            alt=""
            style={state === 'ok' ? undefined : { position: 'absolute', opacity: 0, width: 1, height: 1 }}
            onLoad={(e) => setState(e.currentTarget.naturalWidth > 0 && e.currentTarget.naturalWidth < 20 ? 'none' : 'ok')}
            onError={() => setState('none')}
          />
        )}
      </span>
    </span>
  );
}

export function PersonLine(props: { email: string; name: string; sub?: React.ReactNode; size?: 'sm' | 'md' | 'lg' }): JSX.Element {
  return (
    <span className="to-person">
      <Avatar email={props.email} name={props.name} size={props.size} />
      <span style={{ minWidth: 0 }}>
        <span className="to-person-name">{props.name}</span>
        {props.sub && <span className="to-person-sub">{props.sub}</span>}
      </span>
    </span>
  );
}

export function Empty(props: { icon?: string; title: string; children?: React.ReactNode }): JSX.Element {
  return (
    <div className="to-empty">
      <Icon iconName={props.icon || 'Calendar'} />
      <div style={{ fontWeight: 650, color: 'var(--to-text)' }}>{props.title}</div>
      {props.children && <div style={{ marginTop: 6 }}>{props.children}</div>}
    </div>
  );
}

export function Loading(props: { label?: string }): JSX.Element {
  return (
    <div className="to-center" style={{ minHeight: 200 }}>
      <Spinner size={SpinnerSize.large} label={props.label || 'Loading…'} />
    </div>
  );
}

export function ErrorBox(props: { error?: string }): JSX.Element | null {
  if (!props.error) return null;
  return (
    <MessageBar messageBarType={MessageBarType.error} isMultiline>
      {props.error}
    </MessageBar>
  );
}

export function PageHeader(props: { title: React.ReactNode; sub?: React.ReactNode; actions?: React.ReactNode; eyebrow?: string }): JSX.Element {
  return (
    <div className="to-page-header">
      <div>
        {props.eyebrow && (
          <p className="to-eyebrow">
            <span className="to-star">✳</span> {props.eyebrow}
          </p>
        )}
        <h1>{props.title}</h1>
        {props.sub && <div className="to-sub">{props.sub}</div>}
      </div>
      {props.actions && <div className="to-page-actions">{props.actions}</div>}
    </div>
  );
}

export function Kpi(props: { label: string; value: React.ReactNode; hint?: string; icon?: string; tint?: number; tone?: 'danger' | 'warning' | 'success'; onClick?: () => void }): JSX.Element {
  return (
    <div
      className={'to-kpi' + (props.onClick ? ' is-link' : '') + (props.tone ? ' is-' + props.tone : '')}
      onClick={props.onClick}
      onKeyDown={(e) => props.onClick && (e.key === 'Enter' || e.key === ' ') && props.onClick()}
      role={props.onClick ? 'button' : undefined}
      tabIndex={props.onClick ? 0 : undefined}
    >
      <div className="to-kpi-icon">
        <div className="to-kpi-label">{props.label}</div>
        {props.icon && (
          <span className={`to-icon-tile is-sm to-tint-${props.tint || 1}`}>
            <Icon iconName={props.icon} />
          </span>
        )}
      </div>
      <div className="to-kpi-value">{props.value}</div>
      {props.hint && <div className="to-kpi-hint">{props.hint}</div>}
    </div>
  );
}

export function Card(props: { title?: React.ReactNode; icon?: string; actions?: React.ReactNode; children: React.ReactNode; className?: string; style?: React.CSSProperties }): JSX.Element {
  return (
    <div className={'to-card ' + (props.className || '')} style={props.style}>
      {props.title && (
        <h3 className="to-card-title">
          {props.icon && <Icon iconName={props.icon} />}
          {props.title}
          {props.actions && <span className="to-card-actions">{props.actions}</span>}
        </h3>
      )}
      {props.children}
    </div>
  );
}

const STATUS_LABEL: Record<RequestStatus, string> = {
  Pending: 'Pending',
  Approved: 'Approved',
  Rejected: 'Declined',
  Cancelled: 'Cancelled',
  CancelRequested: 'Cancellation pending'
};

export function StatusBadge(props: { status: RequestStatus }): JSX.Element {
  return <span className={`to-status is-${props.status.toLowerCase()}`}>{STATUS_LABEL[props.status]}</span>;
}

export function TypeChip(props: { type?: ILeaveType; label?: string }): JSX.Element {
  const t = props.type;
  return (
    <span className="to-type-chip" style={{ '--to-type': t ? t.color : '#8a959c' } as unknown as React.CSSProperties}>
      <Icon iconName={(t && t.icon) || 'Calendar'} />
      {props.label || (t ? t.name : 'Leave')}
    </span>
  );
}

/** Segmented control. */
export function Segmented<T extends string>(props: { value: T; options: Array<{ key: T; text: string; icon?: string }>; onChange: (v: T) => void; label?: string }): JSX.Element {
  return (
    <div className="to-seg" role="group" aria-label={props.label}>
      {props.options.map((o) => (
        <button key={o.key} type="button" className={o.key === props.value ? 'is-active' : ''} aria-pressed={o.key === props.value} onClick={() => props.onChange(o.key)} title={o.text}>
          {o.icon && <Icon iconName={o.icon} />}
          <span className="to-seg-text">{o.text}</span>
        </button>
      ))}
    </div>
  );
}

/** Searchable pick of an employee from the employee list. */
export function EmployeePicker(props: { label?: string; value?: string; onChange: (e: IEmployee | undefined) => void; filter?: (e: IEmployee) => boolean; placeholder?: string; required?: boolean; disabled?: boolean }): JSX.Element {
  const { world } = useApp();
  const options: IComboBoxOption[] = world.employees
    .filter((e) => e.active && (!props.filter || props.filter(e)))
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((e) => ({ key: lc(e.email), text: e.name, data: e, title: [e.jobTitle, e.department].filter((x) => !!x).join(' · ') }));
  return (
    <ComboBox
      label={props.label}
      required={props.required}
      disabled={props.disabled}
      placeholder={props.placeholder || 'Type a name'}
      allowFreeform={false}
      autoComplete="on"
      options={options}
      selectedKey={props.value ? lc(props.value) : null}
      onChange={(_, o) => props.onChange(o ? (o.data as IEmployee) : undefined)}
      useComboBoxAsMenuWidth
      onRenderOption={(o) => (
        <span className="to-person">
          <span style={{ minWidth: 0 }}>
            <span className="to-person-name">{o ? o.text : ''}</span>
            <span className="to-person-sub">{o && o.title}</span>
          </span>
        </span>
      )}
    />
  );
}

/** Pick people (for HR, admins, workflow approvers): from the employee list, or any typed email. */
export function PeopleField(props: { label: string; value: IPerson[]; onChange: (v: IPerson[]) => void; single?: boolean; description?: string }): JSX.Element {
  const [typed, setTyped] = React.useState('');
  const add = (p: IPerson): void => {
    if (props.value.some((x) => lc(x.email) === lc(p.email))) return;
    props.onChange(props.single ? [p] : props.value.concat(p));
  };
  return (
    <div>
      <EmployeePicker label={props.label} onChange={(e) => e && add({ email: e.email, name: e.name })} placeholder="Add a person" />
      <div className="to-row" style={{ marginTop: 6 }}>
        <TextField
          placeholder="…or type an email and press Enter"
          value={typed}
          onChange={(_, v) => setTyped(v || '')}
          onKeyDown={(ev) => {
            if (ev.key === 'Enter' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(typed.trim())) {
              add({ email: typed.trim(), name: typed.trim() });
              setTyped('');
            }
          }}
          styles={{ root: { flex: 1 } }}
        />
      </div>
      {props.description && <div className="to-small to-muted" style={{ marginTop: 4 }}>{props.description}</div>}
      <div className="to-chips" style={{ marginTop: 8 }}>
        {props.value.map((p) => (
          <span key={p.email} className="to-chip">
            {p.name || p.email}
            <button type="button" aria-label={`Remove ${p.name}`} onClick={() => props.onChange(props.value.filter((x) => x.email !== p.email))}>
              ×
            </button>
          </span>
        ))}
        {!props.value.length && <span className="to-small to-muted">Nobody yet.</span>}
      </div>
    </div>
  );
}

interface IConfirmState {
  title: string;
  message: string;
  confirmText: string;
  withInput?: string;
  resolve: (value: string | false) => void;
}

/** Promise-based confirm / prompt dialog. */
export function useConfirm(): [(title: string, message: string, confirmText?: string, withInput?: string) => Promise<string | false>, JSX.Element | null] {
  const [state, setState] = React.useState<IConfirmState | undefined>();
  const [text, setText] = React.useState('');
  const ask = React.useCallback(
    (title: string, message: string, confirmText: string = 'OK', withInput?: string): Promise<string | false> =>
      new Promise((resolve) => {
        setText('');
        setState({ title, message, confirmText, withInput, resolve });
      }),
    []
  );
  const close = (v: string | false): void => {
    if (state) state.resolve(v);
    setState(undefined);
  };
  const dialog = state ? (
    <Dialog hidden={false} onDismiss={() => close(false)} dialogContentProps={{ type: DialogType.normal, title: state.title, subText: state.message }} modalProps={{ isBlocking: true }}>
      {state.withInput !== undefined && <TextField label={state.withInput} multiline rows={3} value={text} onChange={(_, v) => setText(v || '')} />}
      <DialogFooter>
        <PrimaryButton text={state.confirmText} onClick={() => close(text || 'ok')} />
        <DefaultButton text="Cancel" onClick={() => close(false)} />
      </DialogFooter>
    </Dialog>
  ) : null;
  return [ask, dialog];
}

/** Run an async action with a busy flag and toasts. */
export function useBusy(): [boolean, (work: () => Promise<unknown>, success?: string) => Promise<boolean>] {
  const { toast } = useApp();
  const [busy, setBusy] = React.useState(false);
  const run = React.useCallback(
    async (work: () => Promise<unknown>, success?: string): Promise<boolean> => {
      setBusy(true);
      try {
        await work();
        if (success) toast(success, 'success');
        return true;
      } catch (e) {
        toast(e instanceof Error ? e.message : String(e), 'error');
        return false;
      } finally {
        setBusy(false);
      }
    },
    [toast]
  );
  return [busy, run];
}
