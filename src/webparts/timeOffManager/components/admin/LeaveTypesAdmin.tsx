import * as React from 'react';
import { Checkbox, DefaultButton, Dropdown, Icon, IconButton, Panel, PanelType, Pivot, PivotItem, PrimaryButton, SearchBox, TextField, Toggle } from '@fluentui/react';
import { ILeaveType } from '../../models/types';
import { ISectionProps } from './AdminPage';
import { blankLeaveType } from '../../schema/leaveTypes';
import { Card, TypeChip, NumberBox } from '../common/Ui';

const MODES = [
  { key: 'fixed', text: 'Fixed days per leave year' },
  { key: 'accrual', text: 'Accrues over the year' },
  { key: 'earned', text: 'Earned (comp off / in lieu)' },
  { key: 'unlimited', text: 'No limit (tracked only)' }
];

function policySummary(t: ILeaveType): string {
  const e = t.entitlement;
  if (e.mode === 'unlimited') return 'No limit';
  if (e.mode === 'earned') return `Earned${e.expiryDays ? `, lapses after ${e.expiryDays} d` : ''}`;
  const grades = Object.keys(e.byGrade || {}).length;
  return `${e.days} d/yr${e.mode === 'accrual' ? `, ${e.accrual.frequency}` : ''}${grades ? `, ${grades} grades` : ''}${e.tenure.length ? `, ${e.tenure.length} tenure tiers` : ''}${e.carryForwardMax ? `, carry ${e.carryForwardMax < 0 ? 'all' : e.carryForwardMax}` : ''}`;
}

/** The leave types (30 to start with) and every policy setting of each. */
export function LeaveTypesAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const [edit, setEdit] = React.useState<string | undefined>();
  const [q, setQ] = React.useState('');
  const types = draft.leaveTypes;
  const shown = types.filter((t) => !q || t.name.toLowerCase().indexOf(q.toLowerCase()) >= 0);
  const update = (t: ILeaveType): void => setDraft({ ...draft, leaveTypes: types.map((x) => (x.key === t.key ? t : x)) });
  const add = (): void => {
    let key = 'custom';
    for (let i = 1; types.some((t) => t.key === key); i++) key = `custom${i}`;
    setDraft({ ...draft, leaveTypes: types.concat(blankLeaveType(key)) });
    setEdit(key);
  };
  const move = (i: number, d: number): void => {
    const next = types.slice();
    const [x] = next.splice(i, 1);
    next.splice(Math.max(0, Math.min(next.length, i + d)), 0, x);
    setDraft({ ...draft, leaveTypes: next });
  };
  const current = types.filter((t) => t.key === edit)[0];
  return (
    <>
      <Card
        title={`${types.filter((t) => t.active).length} of ${types.length} leave types active`}
        icon="Airplane"
        actions={<PrimaryButton iconProps={{ iconName: 'Add' }} text="New leave type" onClick={add} />}
      >
        <SearchBox placeholder="Find a leave type" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { maxWidth: 320, marginBottom: 12 } }} />
        <div className="to-table-wrap">
          <table className="to-table">
            <thead>
              <tr>
                <th>Leave type</th>
                <th>Policy</th>
                <th>Units</th>
                <th>Approval</th>
                <th>Active</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {shown.map((t) => {
                const i = types.indexOf(t);
                const wf = draft.workflows.filter((w) => w.key === t.workflowKey)[0];
                return (
                  <tr key={t.key} className={t.active ? '' : 'is-inactive'}>
                    <td>
                      <TypeChip type={t} />
                      <div className="to-small to-muted">
                        {t.code} · {t.paid ? 'paid' : 'unpaid'}
                      </div>
                    </td>
                    <td className="to-small">{policySummary(t)}</td>
                    <td className="to-small">{[t.units.day && 'days', t.units.half && 'half days', t.units.hours && 'hours'].filter((x) => !!x).join(', ')}</td>
                    <td className="to-small">
                      {t.autoApprove.enabled ? <><Icon iconName="LightningBolt" /> auto{t.autoApprove.maxDays ? ` ≤ ${t.autoApprove.maxDays} d` : ''}, else </> : null}
                      {wf ? wf.name : '—'}
                    </td>
                    <td>
                      <Toggle checked={t.active} onChange={(_, c) => update({ ...t, active: !!c })} ariaLabel={`${t.name} active`} />
                    </td>
                    <td style={{ whiteSpace: 'nowrap' }}>
                      <IconButton iconProps={{ iconName: 'Up' }} ariaLabel="Move up" disabled={i === 0} onClick={() => move(i, -1)} />
                      <IconButton iconProps={{ iconName: 'Down' }} ariaLabel="Move down" disabled={i === types.length - 1} onClick={() => move(i, 1)} />
                      <DefaultButton text="Edit" onClick={() => setEdit(t.key)} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
      {current && <LeaveTypeEditor type={current} props={props} onChange={update} onClose={() => setEdit(undefined)} />}
    </>
  );
}

function LeaveTypeEditor(p: { type: ILeaveType; props: ISectionProps; onChange: (t: ILeaveType) => void; onClose: () => void }): JSX.Element {
  const t = p.type;
  const { draft, setDraft } = p.props;
  const e = t.entitlement;
  const set = (patch: Partial<ILeaveType>): void => p.onChange({ ...t, ...patch });
  const setE = (patch: Partial<ILeaveType['entitlement']>): void => set({ entitlement: { ...e, ...patch } });
  const num = (v: string | undefined): number => Number(v) || 0;
  return (
    <Panel
      isOpen
      type={PanelType.medium}
      onDismiss={p.onClose}
      headerText={t.name}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="to-row">
          <PrimaryButton text="Done" onClick={p.onClose} />
          <span className="to-small to-muted">Changes are kept in the page; press Save changes to store them.</span>
          <span className="to-spacer" />
          {t.key.indexOf('custom') === 0 && (
            <DefaultButton
              text="Delete"
              onClick={() => {
                setDraft({ ...draft, leaveTypes: draft.leaveTypes.filter((x) => x.key !== t.key) });
                p.onClose();
              }}
            />
          )}
        </div>
      )}
    >
      <Pivot>
        <PivotItem headerText="Basics">
          <div className="to-stack" style={{ marginTop: 12 }}>
            <div className="to-form-grid">
              <TextField label="Name" value={t.name} onChange={(_, v) => set({ name: v || '' })} />
              <TextField label="Code" value={t.code} onChange={(_, v) => set({ code: (v || '').toUpperCase() })} />
              <div>
                <div className="to-label">Colour</div>
                <input type="color" className="to-color-input" value={t.color} onChange={(ev) => set({ color: ev.target.value })} aria-label="Colour" />
              </div>
              <TextField label="Icon (Fluent UI name)" value={t.icon} onChange={(_, v) => set({ icon: v || 'Calendar' })} onRenderSuffix={() => <Icon iconName={t.icon} />} />
            </div>
            <TextField label="Description (shown when requesting)" multiline rows={2} value={t.description} onChange={(_, v) => set({ description: v || '' })} />
            <Toggle label="Active" checked={t.active} onChange={(_, c) => set({ active: !!c })} />
            <Toggle label="Paid leave" checked={t.paid} onChange={(_, c) => set({ paid: !!c })} />
            <div>
              <div className="to-label">Can be taken as</div>
              <div className="to-row">
                <Checkbox label="Full days" checked={t.units.day} onChange={(_, c) => set({ units: { ...t.units, day: !!c } })} />
                <Checkbox label="Half days" checked={t.units.half} onChange={(_, c) => set({ units: { ...t.units, half: !!c } })} />
                <Checkbox label="Hours (e.g. two-hour leave)" checked={t.units.hours} onChange={(_, c) => set({ units: { ...t.units, hours: !!c } })} />
              </div>
            </div>
            <Toggle label="Count every calendar day (weekends and holidays included)" checked={t.countCalendarDays} onChange={(_, c) => set({ countCalendarDays: !!c })} onText="Calendar days" offText="Working days only (non-working days excluded)" />
            <Toggle label="Sandwich rule: weekends and holidays between leave days count" checked={!!t.sandwich} onChange={(_, c) => set({ sandwich: !!c })} disabled={t.countCalendarDays} />
            <Toggle label="Can repeat (weekly, bi-weekly, monthly)" checked={t.allowRecurrence} onChange={(_, c) => set({ allowRecurrence: !!c })} />
          </div>
        </PivotItem>
        <PivotItem headerText="Entitlement">
          <div className="to-stack" style={{ marginTop: 12 }}>
            <Dropdown label="How the balance is built" selectedKey={e.mode} options={MODES} onChange={(_, o) => o && setE({ mode: o.key as ILeaveType['entitlement']['mode'] })} />
            {(e.mode === 'fixed' || e.mode === 'accrual') && (
              <>
                <NumberBox label="Days per leave year (standard)" min={0} max={400} step={0.5} value={String(e.days)} onChange={(_, v) => setE({ days: num(v) })} />
                {e.mode === 'accrual' && (
                  <div className="to-form-grid">
                    <Dropdown label="Accrues" selectedKey={e.accrual.frequency} options={[{ key: 'monthly', text: 'Monthly' }, { key: 'biweekly', text: 'Bi-weekly (26 periods)' }, { key: 'quarterly', text: 'Quarterly' }, { key: 'yearly', text: 'Yearly' }]} onChange={(_, o) => o && setE({ accrual: { ...e.accrual, frequency: o.key as 'monthly' } })} />
                    <Dropdown label="Credited at the" selectedKey={e.accrual.timing} options={[{ key: 'start', text: 'Start of each period' }, { key: 'end', text: 'End of each period' }]} onChange={(_, o) => o && setE({ accrual: { ...e.accrual, timing: o.key as 'start' } })} />
                  </div>
                )}
                {e.mode === 'accrual' && <Toggle label="Employees may book against what accrues later in the leave year" checked={e.bookAhead !== false} onChange={(_, c) => setE({ bookAhead: !!c })} />}
                <Toggle label={e.mode === 'accrual' ? 'New joiners accrue from their start date' : 'Prorate the first year by the start date'} checked={e.prorate} onChange={(_, c) => setE({ prorate: !!c })} />
                <div>
                  <div className="to-label">Level / grade based days (empty = the standard days)</div>
                  <div className="to-grade-grid">
                    {draft.grades.map((g) => (
                      <TextField
                        key={g}
                        prefix={g}
                        type="number"
                        value={typeof e.byGrade[g] === 'number' ? String(e.byGrade[g]) : ''}
                        onChange={(_, v) => {
                          const next = { ...e.byGrade };
                          if (v === '' || v === undefined) delete next[g];
                          else next[g] = Number(v);
                          setE({ byGrade: next });
                        }}
                      />
                    ))}
                  </div>
                </div>
                <div>
                  <div className="to-label">Tenure based extra days</div>
                  {e.tenure.map((tier, i) => (
                    <div key={i} className="to-row" style={{ marginBottom: 6 }}>
                      <span>After</span>
                      <NumberBox min={0} max={50} value={String(tier.years)} onChange={(_, v) => setE({ tenure: e.tenure.map((x, j) => (j === i ? { ...x, years: num(v) } : x)) })} styles={{ root: { width: 90 } }} />
                      <span>years:</span>
                      <NumberBox min={0} max={100} step={0.5} value={String(tier.extraDays)} onChange={(_, v) => setE({ tenure: e.tenure.map((x, j) => (j === i ? { ...x, extraDays: num(v) } : x)) })} styles={{ root: { width: 90 } }} />
                      <span>extra days</span>
                      <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove tier" onClick={() => setE({ tenure: e.tenure.filter((_, j) => j !== i) })} />
                    </div>
                  ))}
                  <DefaultButton iconProps={{ iconName: 'Add' }} text="Add tier" onClick={() => setE({ tenure: e.tenure.concat({ years: (e.tenure.length ? e.tenure[e.tenure.length - 1].years : 0) + 2, extraDays: 1 }) })} />
                </div>
                <div className="to-form-grid">
                  <NumberBox label="Carry forward at most (days, −1 = all)" min={-1} max={400} value={String(e.carryForwardMax)} onChange={(_, v) => setE({ carryForwardMax: Number(v) })} />
                  <NumberBox label="Balance cap (0 = none)" min={0} max={400} value={String(e.maxBalance)} onChange={(_, v) => setE({ maxBalance: num(v) })} />
                </div>
              </>
            )}
            {e.mode === 'earned' && (
              <>
                <NumberBox label="Credits lapse after (days, 0 = never)" min={0} max={730} value={String(e.expiryDays)} onChange={(_, v) => setE({ expiryDays: num(v) })} />
                <Toggle label="This is the comp-off type employees claim into" checked={!!t.isCompOff} onChange={(_, c) => set({ isCompOff: !!c })} />
              </>
            )}
          </div>
        </PivotItem>
        <PivotItem headerText="Rules">
          <div className="to-stack" style={{ marginTop: 12 }}>
            <div className="to-form-grid">
              <NumberBox label="Notice needed (days)" min={0} max={365} value={String(t.minNoticeDays)} onChange={(_, v) => set({ minNoticeDays: num(v) })} />
              <NumberBox label="Longest request (days, 0 = any)" min={0} max={400} value={String(t.maxConsecutiveDays)} onChange={(_, v) => set({ maxConsecutiveDays: num(v) })} />
              <NumberBox label="May go negative by (days)" min={0} max={60} step={0.5} value={String(t.allowNegative)} onChange={(_, v) => set({ allowNegative: num(v) })} />
              <NumberBox label="Document needed over (days, 0 = never)" min={0} max={60} value={String(t.attachmentOverDays)} onChange={(_, v) => set({ attachmentOverDays: num(v) })} />
              <NumberBox label="Months of service needed" min={0} max={240} value={String(t.eligibility.minServiceMonths)} onChange={(_, v) => set({ eligibility: { ...t.eligibility, minServiceMonths: num(v) } })} />
            </div>
            <div>
              <div className="to-label">Only for grades (none ticked = everyone)</div>
              <div className="to-row">
                {draft.grades.map((g) => (
                  <Checkbox key={g} label={g} checked={t.eligibility.grades.indexOf(g) >= 0} onChange={(_, c) => set({ eligibility: { ...t.eligibility, grades: c ? t.eligibility.grades.concat(g) : t.eligibility.grades.filter((x) => x !== g) } })} />
                ))}
              </div>
            </div>
            <div>
              <div className="to-label">Only for (none ticked = everyone)</div>
              <div className="to-row">
                {['Female', 'Male', 'Other'].map((g) => (
                  <Checkbox key={g} label={g} checked={t.eligibility.genders.indexOf(g) >= 0} onChange={(_, c) => set({ eligibility: { ...t.eligibility, genders: c ? t.eligibility.genders.concat(g) : t.eligibility.genders.filter((x) => x !== g) } })} />
                ))}
              </div>
            </div>
          </div>
        </PivotItem>
        <PivotItem headerText="Approval">
          <div className="to-stack" style={{ marginTop: 12 }}>
            <Dropdown label="Approval workflow" selectedKey={t.workflowKey} options={draft.workflows.map((w) => ({ key: w.key, text: w.name }))} onChange={(_, o) => o && set({ workflowKey: String(o.key) })} />
            <Toggle label="Auto-approve" checked={t.autoApprove.enabled} onChange={(_, c) => set({ autoApprove: { ...t.autoApprove, enabled: !!c } })} />
            {t.autoApprove.enabled && <NumberBox label="…for requests up to (days, 0 = always)" min={0} max={60} step={0.5} value={String(t.autoApprove.maxDays)} onChange={(_, v) => set({ autoApprove: { ...t.autoApprove, maxDays: num(v) } })} />}
          </div>
        </PivotItem>
      </Pivot>
    </Panel>
  );
}
