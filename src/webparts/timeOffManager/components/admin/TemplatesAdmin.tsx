import * as React from 'react';
import { DefaultButton, Dropdown, IconButton, PrimaryButton, SpinButton, TextField } from '@fluentui/react';
import { ILeaveTemplate, LeaveUnit, RecurrenceFreq } from '../../models/types';
import { ISectionProps } from './AdminPage';
import { Card, TypeChip } from '../common/Ui';
import { describeRecurrence } from '../../services/requests';

/** Leave templates: one-click requests (e.g. "Friday afternoon off", "Doctor appointment (2 h)", "WFH every other Friday"). */
export function TemplatesAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const update = (t: ILeaveTemplate): void => setDraft({ ...draft, templates: draft.templates.map((x) => (x.key === t.key ? t : x)) });
  const add = (): void => {
    const key = 't' + Date.now().toString(36);
    setDraft({ ...draft, templates: draft.templates.concat({ key, name: 'New template', leaveType: draft.leaveTypes[0].key, unit: 'day', days: 1, reason: '' }) });
  };
  return (
    <>
      <Card title="Templates fill in the request form" icon="Lightning" actions={<PrimaryButton iconProps={{ iconName: 'Add' }} text="New template" onClick={add} />}>
        <p className="to-muted" style={{ marginTop: 0 }}>Employees pick a template on the home page or in the request form; they only choose the date. Templates may repeat (weekly, bi-weekly, monthly) when the leave type allows it.</p>
      </Card>
      {draft.templates.map((t) => {
        const type = draft.leaveTypes.filter((x) => x.key === t.leaveType)[0];
        const freq = (t.recurrence && t.recurrence.freq) || 'none';
        return (
          <Card key={t.key} title={<>{t.name} <TypeChip type={type} /></>} actions={<IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete template" onClick={() => setDraft({ ...draft, templates: draft.templates.filter((x) => x.key !== t.key) })} />}>
            <div className="to-form-grid">
              <TextField label="Name" value={t.name} onChange={(_, v) => update({ ...t, name: v || '' })} />
              <Dropdown label="Leave type" selectedKey={t.leaveType} options={draft.leaveTypes.filter((x) => x.active).map((x) => ({ key: x.key, text: x.name }))} onChange={(_, o) => o && update({ ...t, leaveType: String(o.key) })} />
              <Dropdown label="Length" selectedKey={t.unit} options={[{ key: 'day', text: 'Full day(s)' }, { key: 'half', text: 'Half day' }, { key: 'hours', text: 'Hours' }]} onChange={(_, o) => o && update({ ...t, unit: o.key as LeaveUnit })} />
              {t.unit === 'day' && <SpinButton label="Calendar days from the start date" min={1} max={60} value={String(t.days || 1)} onChange={(_, v) => update({ ...t, days: Number(v) || 1 })} />}
              {t.unit === 'half' && <Dropdown label="Half" selectedKey={t.halfDayPart || 'am'} options={[{ key: 'am', text: 'Morning' }, { key: 'pm', text: 'Afternoon' }]} onChange={(_, o) => o && update({ ...t, halfDayPart: o.key as 'am' })} />}
              {t.unit === 'hours' && (
                <>
                  <SpinButton label="Hours" min={0.5} max={12} step={0.5} value={String(t.hours || 2)} onChange={(_, v) => update({ ...t, hours: Number(v) || 2 })} />
                  <TextField label="Start time" type="time" value={t.startTime || '09:00'} onChange={(_, v) => update({ ...t, startTime: v || '09:00' })} />
                </>
              )}
              <TextField label="Reason" value={t.reason} onChange={(_, v) => update({ ...t, reason: v || '' })} />
              <Dropdown
                label="Repeat"
                disabled={!type || !type.allowRecurrence}
                selectedKey={freq}
                options={[{ key: 'none', text: 'Does not repeat' }, { key: 'weekly', text: 'Weekly' }, { key: 'biweekly', text: 'Bi-weekly' }, { key: 'monthly', text: 'Monthly' }]}
                onChange={(_, o) => o && update({ ...t, recurrence: o.key === 'none' ? undefined : { freq: o.key as RecurrenceFreq, weekdays: (t.recurrence && t.recurrence.weekdays) || [], count: (t.recurrence && t.recurrence.count) || 4 } })}
              />
              {t.recurrence && <SpinButton label="Times" min={1} max={52} value={String(t.recurrence.count || 4)} onChange={(_, v) => update({ ...t, recurrence: { ...(t.recurrence as NonNullable<ILeaveTemplate['recurrence']>), count: Number(v) || 4 } })} />}
            </div>
            {t.recurrence && <div className="to-small to-muted" style={{ marginTop: 6 }}>{describeRecurrence(t.recurrence)}</div>}
          </Card>
        );
      })}
      {!draft.templates.length && <DefaultButton text="Add a template" onClick={add} />}
    </>
  );
}
