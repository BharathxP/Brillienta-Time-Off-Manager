import * as React from 'react';
import { Checkbox, DefaultButton, Dropdown, IconButton, PrimaryButton, SpinButton, TextField, Toggle } from '@fluentui/react';
import { IBlackout } from '../../models/types';
import { ISectionProps } from './AdminPage';
import { useApp } from '../AppContext';
import { Card } from '../common/Ui';
import { IsoDatePicker } from '../requests/RequestForm';
import { addDays, fmtDate, today } from '../../services/dates';

/** Blackout periods and minimum staffing. */
export function RulesAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { world } = useApp();
  const departments = world.employees.map((e) => e.department || '').filter((d, i, all) => d && all.indexOf(d) === i).sort();
  const st = draft.staffing;
  const setSt = (patch: Partial<typeof st>): void => setDraft({ ...draft, staffing: { ...st, ...patch } });
  const update = (b: IBlackout): void => setDraft({ ...draft, blackouts: draft.blackouts.map((x) => (x.key === b.key ? b : x)) });
  const add = (): void =>
    setDraft({ ...draft, blackouts: draft.blackouts.concat({ key: 'b' + Date.now().toString(36), name: 'Year-end close', from: addDays(today(), 30), to: addDays(today(), 36), departments: [], leaveTypes: [], mode: 'block' }) });
  return (
    <>
      <Card title="Blackout periods" icon="Blocked" actions={<PrimaryButton iconProps={{ iconName: 'Add' }} text="New blackout" onClick={add} />}>
        <p className="to-muted" style={{ marginTop: 0 }}>
          Dates when leave is not allowed (or only with a warning) — e.g. year-end close, audits, peak season. HR and admins can still submit with a warning.
        </p>
        {!draft.blackouts.length && <div className="to-small to-muted">No blackout periods.</div>}
        {draft.blackouts.map((b) => (
          <div key={b.key} className="to-level">
            <div className="to-level-body">
              <div className="to-form-grid">
                <TextField label="Name" value={b.name} onChange={(_, v) => update({ ...b, name: v || '' })} />
                <Dropdown label="Effect" selectedKey={b.mode} options={[{ key: 'block', text: 'Block requests' }, { key: 'warn', text: 'Warn only' }]} onChange={(_, o) => o && update({ ...b, mode: o.key as 'block' })} />
                <IsoDatePicker label="From" value={b.from} onChange={(d) => update({ ...b, from: d, to: b.to < d ? d : b.to })} />
                <IsoDatePicker label="To" value={b.to} onChange={(d) => update({ ...b, to: d })} minDate={b.from} />
                <Dropdown
                  label="Departments (none = everyone)"
                  multiSelect
                  selectedKeys={b.departments}
                  options={departments.map((d) => ({ key: d, text: d }))}
                  onChange={(_, o) => o && update({ ...b, departments: o.selected ? b.departments.concat(String(o.key)) : b.departments.filter((x) => x !== o.key) })}
                />
                <Dropdown
                  label="Leave types (none = all)"
                  multiSelect
                  selectedKeys={b.leaveTypes}
                  options={draft.leaveTypes.filter((t) => t.active).map((t) => ({ key: t.key, text: t.name }))}
                  onChange={(_, o) => o && update({ ...b, leaveTypes: o.selected ? b.leaveTypes.concat(String(o.key)) : b.leaveTypes.filter((x) => x !== o.key) })}
                />
              </div>
              <div className="to-small to-muted" style={{ marginTop: 6 }}>
                {fmtDate(b.from)} – {fmtDate(b.to)} · {b.departments.length ? b.departments.join(', ') : 'all departments'} · {b.leaveTypes.length ? `${b.leaveTypes.length} leave types` : 'all leave types'}
              </div>
            </div>
            <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove blackout" onClick={() => setDraft({ ...draft, blackouts: draft.blackouts.filter((x) => x.key !== b.key) })} />
          </div>
        ))}
      </Card>
      <Card title="Minimum staffing" icon="Group">
        <div className="to-stack">
          <Toggle label="Limit how many people of a department may be away on the same day" checked={st.enabled} onChange={(_, c) => setSt({ enabled: !!c })} />
          <div className="to-form-grid">
            <SpinButton label="At most (% of the department, 0 = no % limit)" min={0} max={100} step={5} value={String(st.maxAwayPercent)} onChange={(_, v) => setSt({ maxAwayPercent: Number(v) || 0 })} />
            <SpinButton label="At most (people, 0 = no count limit)" min={0} max={500} value={String(st.maxAwayCount)} onChange={(_, v) => setSt({ maxAwayCount: Number(v) || 0 })} />
            <Dropdown label="When over the limit" selectedKey={st.mode} options={[{ key: 'warn', text: 'Warn the employee and the approver' }, { key: 'block', text: 'Block the request' }]} onChange={(_, o) => o && setSt({ mode: o.key as 'warn' })} />
          </div>
          <div>
            <div className="to-label">Per department (overrides the limits above; empty = use them)</div>
            <div className="to-grade-grid">
              {departments.map((d) => (
                <TextField
                  key={d}
                  prefix={d}
                  type="number"
                  value={st.perDepartment[d] ? String(st.perDepartment[d]) : ''}
                  onChange={(_, v) => {
                    const next = { ...st.perDepartment };
                    if (!v || Number(v) <= 0) delete next[d];
                    else next[d] = Number(v);
                    setSt({ perDepartment: next });
                  }}
                />
              ))}
            </div>
          </div>
          <div className="to-small to-muted">Pending and approved leave both count. Approvers see the team cover on every request.</div>
        </div>
      </Card>
    </>
  );
}

/** Microsoft 365: group (Teams channel) calendar and Outlook automatic replies. */
export function M365Admin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const { graph, toast } = useApp();
  const [q, setQ] = React.useState('');
  const [found, setFound] = React.useState<Array<{ id: string; name: string; mail: string }>>([]);
  const g = draft.groupCalendar;
  const search = async (): Promise<void> => {
    if (!graph) {
      toast('Microsoft Graph is not available here.', 'error');
      return;
    }
    try {
      setFound(await graph.findGroups(q));
    } catch (e) {
      toast(`Could not search groups (approve Group.Read.All): ${(e as Error).message}`, 'error');
    }
  };
  return (
    <>
      <Card title="Teams channel calendar (Microsoft 365 group)" icon="TeamsLogo">
        <div className="to-stack">
          <Toggle label="Add approved leave to a Microsoft 365 group calendar" checked={g.enabled} onChange={(_, c) => setDraft({ ...draft, groupCalendar: { ...g, enabled: !!c } })} />
          <div className="to-row" style={{ alignItems: 'flex-end' }}>
            <TextField label="Find the team / group" value={q} onChange={(_, v) => setQ(v || '')} placeholder="Start of the group name" styles={{ root: { flex: 1 } }} />
            <DefaultButton text="Search" disabled={!q.trim()} onClick={() => void search()} />
          </div>
          {found.map((x) => (
            <div key={x.id} className="to-deleg-row">
              <span>
                <strong>{x.name}</strong> <span className="to-muted">{x.mail}</span>
              </span>
              <DefaultButton text={g.groupId === x.id ? 'Selected' : 'Use'} disabled={g.groupId === x.id} onClick={() => setDraft({ ...draft, groupCalendar: { enabled: true, groupId: x.id, groupName: x.name } })} />
            </div>
          ))}
          <TextField label="Group id" value={g.groupId} onChange={(_, v) => setDraft({ ...draft, groupCalendar: { ...g, groupId: (v || '').trim() } })} description={g.groupName ? `Selected: ${g.groupName}` : 'Or paste the group (team) id.'} />
          <div className="to-small to-muted">
            The event (“Name – Leave type”, shown as free) is posted by whoever gives the final approval and removed when the leave is cancelled. Needs the Graph permission Group.ReadWrite.All, and approvers must be members of the group.
          </div>
        </div>
      </Card>
      <Card title="Outlook automatic replies" icon="Mail">
        <div className="to-stack">
          <Toggle label="Offer an automatic reply during leave (employees tick it on the request)" checked={draft.autoReply.enabled} onChange={(_, c) => setDraft({ ...draft, autoReply: { ...draft.autoReply, enabled: !!c } })} />
          <TextField label="Message" multiline rows={4} value={draft.autoReply.message} onChange={(_, v) => setDraft({ ...draft, autoReply: { ...draft.autoReply, message: v || '' } })} description="Placeholders: {name}, {start}, {end}, {manager}." />
          <Checkbox label="Uses the Graph permission MailboxSettings.ReadWrite; the reply is scheduled from the employee's own browser the next time they open the app after approval." checked disabled />
        </div>
      </Card>
    </>
  );
}
