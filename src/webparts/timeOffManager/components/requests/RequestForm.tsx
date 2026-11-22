import * as React from 'react';
import { Checkbox, ChoiceGroup, DatePicker, DayOfWeek, DefaultButton, Dropdown, Icon, IDropdownOption, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, SpinButton, TextField } from '@fluentui/react';
import { IRequestFormOptions, useApp } from '../AppContext';
import { ILeaveTemplate, IRecurrence, LeaveUnit, RecurrenceFreq } from '../../models/types';
import { EmployeePicker, TypeChip } from '../common/Ui';
import { addDays, fmtDate, fmtDays, fmtNum, leaveYearOf, today, WEEKDAY_SHORT, calendarFor, addTime } from '../../services/dates';
import { IDraft, duration, occurrences, validateDraft, describeRecurrence } from '../../services/requests';
import { computeBalance } from '../../services/balances';
import { buildSteps, lc } from '../../services/workflow';

export function isoToDate(iso: string): Date | undefined {
  if (!iso) return undefined;
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d);
}

export function dateToIso(d: Date | null | undefined): string {
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function IsoDatePicker(props: { label: string; value: string; onChange: (v: string) => void; minDate?: string; maxDate?: string; required?: boolean }): JSX.Element {
  const { config } = useApp();
  return (
    <DatePicker
      label={props.label}
      isRequired={props.required}
      value={isoToDate(props.value)}
      onSelectDate={(d) => props.onChange(dateToIso(d))}
      formatDate={(d) => (d ? fmtDate(dateToIso(d), true) : '')}
      firstDayOfWeek={config.weekStartsOn === 0 ? DayOfWeek.Sunday : DayOfWeek.Monday}
      minDate={props.minDate ? isoToDate(props.minDate) : undefined}
      maxDate={props.maxDate ? isoToDate(props.maxDate) : undefined}
      allowTextInput={false}
    />
  );
}

/** Who someone may submit for: themselves, HR anyone, managers the people below them. */
export function useSubmitFor(): (email: string) => boolean {
  const { me, isHr, config, dir } = useApp();
  return React.useCallback(
    (email: string) => {
      if (lc(email) === lc(me.email)) return true;
      if (isHr && config.onBehalf.hr) return true;
      if (!config.onBehalf.managers) return false;
      let e = dir.byEmail[lc(email)];
      for (let i = 0; e && i < 12; i++) {
        if (lc(e.managerEmail) === lc(me.email)) return true;
        e = dir.byEmail[lc(e.managerEmail)];
      }
      return false;
    },
    [me.email, isHr, config.onBehalf, dir]
  );
}

const FREQ_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Does not repeat' },
  { key: 'weekly', text: 'Every week' },
  { key: 'biweekly', text: 'Every two weeks (bi-weekly)' },
  { key: 'monthly', text: 'Every month' }
];

export function RequestForm(props: { options: IRequestFormOptions; onClose: () => void }): JSX.Element {
  const app = useApp();
  const { config, me, world, dir, engine, isHr, isAdmin, toast, licensed } = app;
  const canSubmitFor = useSubmitFor();
  const o = props.options;
  const [kind, setKind] = React.useState<'leave' | 'compoff'>(o.kind || 'leave');
  const [email, setEmail] = React.useState(o.employeeEmail || me.email);
  const employee = dir.byEmail[lc(email)];
  const firstType = config.leaveTypes.filter((t) => t.active && !t.isCompOff)[0];
  const [draft, setDraft] = React.useState<IDraft>(() => ({
    kind: o.kind || 'leave',
    employeeEmail: o.employeeEmail || me.email,
    leaveType: o.leaveType || (firstType ? firstType.key : ''),
    start: o.start || addDays(today(), 1),
    end: o.start || addDays(today(), 1),
    unit: 'day',
    halfDayPart: 'am',
    startTime: calendarFor(config, employee).dayStart,
    hours: 2,
    reason: '',
    recurrence: { freq: 'none', weekdays: [], count: 4 }
  }));
  const [files, setFiles] = React.useState<File[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const set = (patch: Partial<IDraft>): void => setDraft((d) => ({ ...d, ...patch }));

  const applyTemplate = (t: ILeaveTemplate): void => {
    const start = draft.start;
    const span = t.unit === 'day' && t.days && t.days > 1 ? t.days - 1 : 0;
    set({
      leaveType: t.leaveType,
      unit: t.unit,
      halfDayPart: t.halfDayPart || 'am',
      startTime: t.startTime || draft.startTime,
      hours: t.hours || draft.hours,
      reason: t.reason,
      end: addDays(start, span),
      recurrence: t.recurrence ? { ...t.recurrence } : { freq: 'none', weekdays: [], count: 4 }
    });
  };
  React.useEffect(() => {
    if (o.template) applyTemplate(o.template);
  }, []);

  const typeKey = kind === 'compoff' ? config.compOff.leaveType : draft.leaveType;
  const type = config.leaveTypes.filter((t) => t.key === typeKey)[0];
  const full: IDraft = { ...draft, kind, employeeEmail: email, leaveType: typeKey, end: draft.unit === 'day' ? draft.end : draft.start, hasAttachment: files.length > 0 };
  const dur = duration(config, type, employee, full);
  const year = leaveYearOf(full.start || today(), config.yearStartMonth);
  const balance = type && employee ? computeBalance(config, type, employee, year, world.ledger, world.requests, today()) : undefined;
  const privileged = isHr || isAdmin || lc(email) !== lc(me.email);
  const v = validateDraft(config, type, employee, full, dur, kind === 'compoff' ? undefined : balance, world.requests, today(), privileged);
  if (employee && !licensed.has(lc(employee.email))) v.errors.push('This person is above the number of users in the licence. Ask your admin to add users.');
  const recurring = kind === 'leave' && !!type && type.allowRecurrence && !!draft.recurrence && draft.recurrence.freq !== 'none';
  const starts = recurring ? occurrences(full.start, draft.recurrence) : [full.start];
  const steps = type && employee ? buildSteps(config, dir, world.delegations, type, employee, dur.days, today()) : [];
  const auto = kind === 'leave' && !!type && type.autoApprove.enabled && (type.autoApprove.maxDays <= 0 || dur.days <= type.autoApprove.maxDays);
  const cal = calendarFor(config, employee);

  const typeOptions: IDropdownOption[] = config.leaveTypes
    .filter((t) => t.active && !t.isCompOff)
    .concat(config.leaveTypes.filter((t) => t.active && t.isCompOff))
    .map((t) => {
      const b = employee ? computeBalance(config, t, employee, year, world.ledger, world.requests, today()) : undefined;
      const left = !b ? '' : b.unlimited ? 'no limit' : `${fmtNum(b.available)} left`;
      return { key: t.key, text: t.name, data: { type: t, left, ok: !b || b.eligible }, disabled: !!b && !b.eligible };
    });
  const unitOptions = type
    ? [
        { key: 'day', text: 'Full day(s)', disabled: !type.units.day },
        { key: 'half', text: 'Half day', disabled: !type.units.half },
        { key: 'hours', text: 'Hours', disabled: !type.units.hours }
      ]
    : [];
  React.useEffect(() => {
    if (type && !type.units[draft.unit]) set({ unit: type.units.day ? 'day' : type.units.half ? 'half' : 'hours' });
  }, [typeKey]);

  const submit = async (): Promise<void> => {
    setError('');
    setBusy(true);
    try {
      const created = await engine.submit(world, full, { email: me.email, name: me.title, isHr, isAdmin }, files);
      await app.reload();
      app.bumpNotifications();
      const approved = created.filter((r) => r.status === 'Approved').length;
      toast(
        kind === 'compoff'
          ? 'Comp-off claim sent for approval.'
          : created.length > 1
            ? `${created.length} requests created (${describeRecurrence(draft.recurrence)}).`
            : approved
              ? 'Approved straight away by the leave policy.'
              : `Sent to ${created[0] && created[0].steps[0] ? created[0].steps[0].approvers.map((a) => a.name).join(', ') : 'your approver'}.`,
        'success'
      );
      props.onClose();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const rec = draft.recurrence as IRecurrence;
  const setRec = (patch: Partial<IRecurrence>): void => set({ recurrence: { ...rec, ...patch } });
  const myTemplates = config.templates.filter((t) => config.leaveTypes.some((x) => x.key === t.leaveType && x.active));

  return (
    <Panel
      isOpen
      type={PanelType.custom}
      customWidth="760px"
      onDismiss={props.onClose}
      headerText={kind === 'compoff' ? 'Claim comp off' : 'Request time off'}
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="to-row">
          <PrimaryButton text={busy ? 'Submitting…' : kind === 'compoff' ? 'Submit claim' : recurring ? `Submit ${starts.length} requests` : auto ? 'Submit (auto-approved)' : 'Submit request'} disabled={busy || v.errors.length > 0} onClick={() => void submit()} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
          {v.errors.length > 0 && <span className="to-small to-muted">{v.errors.length} thing{v.errors.length === 1 ? '' : 's'} to fix</span>}
        </div>
      )}
    >
      <div className="to-request-form">
        <div className="to-stack">
          {config.compOff.enabled && (
            <ChoiceGroup
              selectedKey={kind}
              options={[
                { key: 'leave', text: 'Take time off', iconProps: { iconName: 'Airplane' } },
                { key: 'compoff', text: 'Claim comp off for extra work', iconProps: { iconName: 'Switch' } }
              ]}
              onChange={(_, op) => op && setKind(op.key as 'leave' | 'compoff')}
              styles={{ flexContainer: { display: 'flex', gap: 8 } }}
            />
          )}
          {(isHr || app.isManager) && (
            <EmployeePicker label="For" value={email} onChange={(e) => e && setEmail(e.email)} filter={(e) => canSubmitFor(e.email)} />
          )}
          {lc(email) !== lc(me.email) && (
            <MessageBar messageBarType={MessageBarType.info}>
              You are submitting on behalf of {employee ? employee.name : email}. They are told, and the history shows it was you.
            </MessageBar>
          )}
          {kind === 'leave' && myTemplates.length > 0 && (
            <div>
              <div className="to-label">Start from a template</div>
              <div className="to-chips">
                {myTemplates.map((t) => (
                  <button key={t.key} type="button" className="to-chip is-button" onClick={() => applyTemplate(t)}>
                    <Icon iconName="Lightning" /> {t.name}
                  </button>
                ))}
              </div>
            </div>
          )}
          {kind === 'leave' && (
            <Dropdown
              label="Leave type"
              required
              selectedKey={draft.leaveType}
              options={typeOptions}
              onChange={(_, op) => op && set({ leaveType: String(op.key) })}
              onRenderOption={(op) => (
                <span className="to-row" style={{ width: '100%', flexWrap: 'nowrap' }}>
                  <TypeChip type={op && op.data.type} />
                  <span className="to-spacer" />
                  <span className="to-small to-muted">{op && op.data.left}</span>
                </span>
              )}
              onRenderTitle={(ops) => (ops && ops[0] ? <TypeChip type={ops[0].data.type} /> : null)}
              calloutProps={{ calloutMaxHeight: 420 }}
            />
          )}
          {type && type.description && <div className="to-small to-muted">{type.description}</div>}
          <ChoiceGroup label={kind === 'compoff' ? 'You worked' : 'Length'} selectedKey={draft.unit} options={unitOptions} onChange={(_, op) => op && set({ unit: op.key as LeaveUnit })} styles={{ flexContainer: { display: 'flex', gap: 16 } }} />
          <div className="to-form-grid">
            <IsoDatePicker label={draft.unit === 'day' ? 'From' : 'Date'} required value={draft.start} onChange={(d) => set({ start: d, end: draft.end < d || draft.unit !== 'day' ? d : draft.end })} maxDate={kind === 'compoff' ? today() : undefined} />
            {draft.unit === 'day' && <IsoDatePicker label="To" required value={draft.end} onChange={(d) => set({ end: d })} minDate={draft.start} maxDate={kind === 'compoff' ? today() : undefined} />}
            {draft.unit === 'half' && (
              <Dropdown
                label="Which half"
                selectedKey={draft.halfDayPart}
                options={[
                  { key: 'am', text: `Morning (${cal.dayStart}–${addTime(cal.dayStart, cal.hoursPerDay / 2)})` },
                  { key: 'pm', text: `Afternoon (${addTime(cal.dayStart, cal.hoursPerDay / 2)}–${addTime(cal.dayStart, cal.hoursPerDay)})` }
                ]}
                onChange={(_, op) => op && set({ halfDayPart: op.key as 'am' | 'pm' })}
              />
            )}
            {draft.unit === 'hours' && (
              <>
                <TextField label="Start time" type="time" value={draft.startTime} onChange={(_, val) => set({ startTime: val || '' })} />
                <Dropdown
                  label="Hours"
                  selectedKey={draft.hours}
                  options={config.hourOptions.map((h) => ({ key: h, text: `${h} hour${h === 1 ? '' : 's'}${draft.startTime ? ` (until ${addTime(draft.startTime, h)})` : ''}` }))}
                  onChange={(_, op) => op && set({ hours: Number(op.key) })}
                />
              </>
            )}
          </div>
          {kind === 'leave' && type && type.allowRecurrence && (
            <div className="to-recur">
              <Dropdown label="Repeat" selectedKey={rec.freq} options={FREQ_OPTIONS} onChange={(_, op) => op && setRec({ freq: op.key as RecurrenceFreq })} />
              {(rec.freq === 'weekly' || rec.freq === 'biweekly') && (
                <div>
                  <div className="to-label">On</div>
                  <div className="to-row">
                    {[1, 2, 3, 4, 5, 6, 0].map((d) => (
                      <Checkbox
                        key={d}
                        label={WEEKDAY_SHORT[d]}
                        checked={rec.weekdays.length ? rec.weekdays.indexOf(d) >= 0 : new Date(draft.start + 'T00:00:00Z').getUTCDay() === d}
                        onChange={(_, c) => {
                          const base = rec.weekdays.length ? rec.weekdays : [new Date(draft.start + 'T00:00:00Z').getUTCDay()];
                          setRec({ weekdays: c ? base.concat(d) : base.filter((x) => x !== d) });
                        }}
                      />
                    ))}
                  </div>
                </div>
              )}
              {rec.freq !== 'none' && (
                <div className="to-form-grid">
                  <SpinButton label="Times" min={1} max={52} value={String(rec.count || '')} onChange={(_, val) => setRec({ count: Number(val) || undefined, until: undefined })} />
                  <IsoDatePicker label="…or until" value={rec.until || ''} onChange={(d) => setRec({ until: d, count: undefined })} minDate={draft.start} />
                </div>
              )}
            </div>
          )}
          <TextField label={kind === 'compoff' ? 'What did you work on?' : 'Reason'} multiline rows={3} value={draft.reason} onChange={(_, val) => set({ reason: val || '' })} required={kind === 'compoff'} />
          <div>
            <div className="to-label">Supporting document{type && type.attachmentOverDays > 0 ? ` (needed over ${type.attachmentOverDays} day${type.attachmentOverDays === 1 ? '' : 's'})` : ' (optional)'}</div>
            <input type="file" multiple onChange={(e) => setFiles(Array.from(e.target.files || []))} aria-label="Attach files" />
          </div>
          {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        </div>
        <aside className="to-summary">
          <div className="to-summary-big">
            <span>{kind === 'compoff' ? 'Credit' : 'Takes'}</span>
            <strong>{draft.unit === 'hours' ? `${dur.hours} h` : fmtDays(dur.days)}</strong>
            {draft.unit === 'hours' && <span className="to-small to-muted">= {fmtNum(dur.days)} day at {cal.hoursPerDay} h/day</span>}
            {recurring && (
              <span className="to-small">
                × {starts.length} occurrences = {fmtDays(dur.days * starts.length)}
              </span>
            )}
          </div>
          {kind === 'leave' && balance && (
            <div className="to-summary-row">
              <span>Balance</span>
              {balance.unlimited ? (
                <strong>No limit</strong>
              ) : (
                <strong>
                  {fmtNum(balance.available)} → <span className={balance.available - dur.days < 0 ? 'to-danger' : ''}>{fmtNum(balance.available - dur.days * starts.length)}</span>
                </strong>
              )}
            </div>
          )}
          {dur.excluded.length > 0 && (
            <div className="to-summary-block">
              <div className="to-label">Not counted</div>
              {dur.excluded.slice(0, 8).map((x) => (
                <div key={x.date} className="to-small">
                  {fmtDate(x.date, true)} · <span className="to-muted">{x.why}</span>
                </div>
              ))}
              {dur.excluded.length > 8 && <div className="to-small to-muted">and {dur.excluded.length - 8} more</div>}
            </div>
          )}
          {recurring && (
            <div className="to-summary-block">
              <div className="to-label">{describeRecurrence(rec)}</div>
              <div className="to-small">{starts.slice(0, 10).map((s) => fmtDate(s)).join(' · ')}{starts.length > 10 ? ' …' : ''}</div>
            </div>
          )}
          <div className="to-summary-block">
            <div className="to-label">Approval</div>
            {auto ? (
              <div className="to-small">
                <Icon iconName="Lightning" /> Approved automatically ({type && type.autoApprove.maxDays > 0 ? `up to ${type.autoApprove.maxDays} day${type.autoApprove.maxDays === 1 ? '' : 's'}` : 'always'})
              </div>
            ) : steps.length ? (
              <ol className="to-route">
                {steps.map((s, i) => (
                  <li key={i}>
                    <strong>{s.label}</strong>
                    <span className="to-small to-muted">
                      {s.approvers.map((a) => a.name + (a.delegatedFrom ? ` (for ${a.delegatedFrom})` : '')).join(', ')}
                    </span>
                  </li>
                ))}
              </ol>
            ) : (
              <div className="to-small to-muted">No approver is set up, so it is approved straight away.</div>
            )}
          </div>
          {v.errors.map((x) => (
            <MessageBar key={x} messageBarType={MessageBarType.blocked}>
              {x}
            </MessageBar>
          ))}
          {v.warnings.map((x) => (
            <MessageBar key={x} messageBarType={MessageBarType.warning}>
              {x}
            </MessageBar>
          ))}
        </aside>
      </div>
    </Panel>
  );
}
