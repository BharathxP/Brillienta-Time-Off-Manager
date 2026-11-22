import * as React from 'react';
import { Checkbox, DefaultButton, IconButton, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { IHoliday, IWorkCalendar } from '../../models/types';
import { ISectionProps } from './AdminPage';
import { Card, NumberBox } from '../common/Ui';
import { IsoDatePicker } from '../requests/RequestForm';
import { fmtDate, isIsoDate, WEEKDAY_SHORT } from '../../services/dates';
import { parseCsv } from '../../services/exporting';

/** Work weeks and holiday calendars: what counts as a working day for each group of employees. */
export function CalendarsAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const update = (c: IWorkCalendar): void => setDraft({ ...draft, calendars: draft.calendars.map((x) => (x.key === c.key ? c : x)) });
  const add = (): void => {
    let key = 'calendar';
    for (let i = 1; draft.calendars.some((c) => c.key === key); i++) key = `calendar${i}`;
    setDraft({ ...draft, calendars: draft.calendars.concat({ key, name: 'New calendar', workDays: [1, 2, 3, 4, 5], hoursPerDay: 8, dayStart: '09:00', holidays: [] }) });
  };
  return (
    <>
      <Card title="Non-working days are excluded" icon="Info">
        <p className="to-muted" style={{ margin: 0 }}>
          A request counts only the working days of the employee’s calendar: non-working weekdays and holidays are left out (unless the leave type counts calendar days). Each employee uses the default calendar unless HR sets another one (Employees & balances).
        </p>
      </Card>
      {draft.calendars.map((c) => (
        <CalendarCard key={c.key} cal={c} isDefault={draft.defaultCalendar === c.key} onChange={update} onDefault={() => setDraft({ ...draft, defaultCalendar: c.key })} onDelete={draft.calendars.length > 1 && draft.defaultCalendar !== c.key ? () => setDraft({ ...draft, calendars: draft.calendars.filter((x) => x.key !== c.key) }) : undefined} />
      ))}
      <div>
        <PrimaryButton iconProps={{ iconName: 'Add' }} text="New calendar" onClick={add} />
      </div>
    </>
  );
}

function CalendarCard(props: { cal: IWorkCalendar; isDefault: boolean; onChange: (c: IWorkCalendar) => void; onDefault: () => void; onDelete?: () => void }): JSX.Element {
  const c = props.cal;
  const [date, setDate] = React.useState('');
  const [name, setName] = React.useState('');
  const [year, setYear] = React.useState(new Date().getFullYear());
  const fileRef = React.useRef<HTMLInputElement>(null);
  const set = (patch: Partial<IWorkCalendar>): void => props.onChange({ ...c, ...patch });
  const holidays = c.holidays.slice().sort((a, b) => (a.date < b.date ? -1 : 1));
  const years = holidays.map((h) => Number(h.date.substring(0, 4))).filter((y, i, all) => all.indexOf(y) === i);
  const shown = holidays.filter((h) => Number(h.date.substring(0, 4)) === year);
  const addHoliday = (h: IHoliday): void => set({ holidays: c.holidays.filter((x) => x.date !== h.date).concat(h) });
  const copyToNext = (): void => {
    const next = shown.map((h) => ({ ...h, date: `${year + 1}${h.date.substring(4)}` })).filter((h) => !c.holidays.some((x) => x.date === h.date));
    set({ holidays: c.holidays.concat(next) });
    setYear(year + 1);
  };
  const importCsv = async (f: File): Promise<void> => {
    const rows = parseCsv(await f.text());
    const add = rows.filter((r) => isIsoDate(r[0])).map((r) => ({ date: r[0], name: r[1] || 'Holiday', optional: /^(yes|true|1|optional)$/i.test(r[2] || '') }));
    set({ holidays: c.holidays.filter((x) => !add.some((a) => a.date === x.date)).concat(add) });
  };
  return (
    <Card
      title={
        <>
          {c.name} {props.isDefault && <span className="to-badge">Default</span>}
        </>
      }
      icon="Calendar"
      actions={
        <span className="to-row">
          {!props.isDefault && <DefaultButton text="Make default" onClick={props.onDefault} />}
          {props.onDelete && <DefaultButton text="Delete" onClick={props.onDelete} />}
        </span>
      }
    >
      <div className="to-form-grid">
        <TextField label="Name" value={c.name} onChange={(_, v) => set({ name: v || '' })} />
        <TextField label="Key (used in the employee list and CSV)" value={c.key} readOnly />
        <NumberBox label="Hours per working day" min={1} max={24} step={0.5} value={String(c.hoursPerDay)} onChange={(_, v) => set({ hoursPerDay: Number(v) || 8 })} />
        <TextField label="Day starts at" type="time" value={c.dayStart} onChange={(_, v) => set({ dayStart: v || '09:00' })} />
      </div>
      <div style={{ marginTop: 12 }}>
        <div className="to-label">Working days</div>
        <div className="to-row">
          {[1, 2, 3, 4, 5, 6, 0].map((d) => (
            <Checkbox key={d} label={WEEKDAY_SHORT[d]} checked={c.workDays.indexOf(d) >= 0} onChange={(_, ch) => set({ workDays: ch ? c.workDays.concat(d) : c.workDays.filter((x) => x !== d) })} />
          ))}
        </div>
      </div>
      <div style={{ marginTop: 16 }}>
        <div className="to-row">
          <div className="to-label" style={{ margin: 0 }}>
            Holidays
          </div>
          <span className="to-spacer" />
          {[...years, year].filter((y, i, all) => all.indexOf(y) === i).sort().map((y) => (
            <DefaultButton key={y} text={String(y)} checked={y === year} onClick={() => setYear(y)} styles={{ root: { minWidth: 56 } }} />
          ))}
          <DefaultButton text={`Copy to ${year + 1}`} disabled={!shown.length} onClick={copyToNext} />
          <DefaultButton text="Import CSV" onClick={() => fileRef.current && fileRef.current.click()} title="Columns: date (yyyy-mm-dd), name, optional (yes/no)" />
          <input ref={fileRef} type="file" accept=".csv" style={{ display: 'none' }} onChange={(e) => e.target.files && e.target.files[0] && void importCsv(e.target.files[0])} />
        </div>
        <table className="to-table" style={{ marginTop: 8 }}>
          <tbody>
            {shown.map((h) => (
              <tr key={h.date}>
                <td style={{ width: 170 }}>{fmtDate(h.date, true)}</td>
                <td>
                  <TextField borderless value={h.name} onChange={(_, v) => set({ holidays: c.holidays.map((x) => (x.date === h.date ? { ...x, name: v || '' } : x)) })} />
                </td>
                <td style={{ width: 150 }}>
                  <Toggle inlineLabel label="Optional" checked={!!h.optional} onChange={(_, ch) => set({ holidays: c.holidays.map((x) => (x.date === h.date ? { ...x, optional: !!ch } : x)) })} />
                </td>
                <td style={{ width: 40 }}>
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove holiday" onClick={() => set({ holidays: c.holidays.filter((x) => x.date !== h.date) })} />
                </td>
              </tr>
            ))}
            {!shown.length && (
              <tr>
                <td className="to-muted">No holidays in {year}.</td>
              </tr>
            )}
          </tbody>
        </table>
        <div className="to-row" style={{ marginTop: 8, alignItems: 'flex-end' }}>
          <IsoDatePicker label="Date" value={date} onChange={setDate} />
          <TextField label="Holiday" value={name} onChange={(_, v) => setName(v || '')} />
          <PrimaryButton
            text="Add"
            disabled={!date || !name.trim()}
            onClick={() => {
              addHoliday({ date, name: name.trim() });
              setYear(Number(date.substring(0, 4)));
              setDate('');
              setName('');
            }}
          />
        </div>
        <div className="to-small to-muted" style={{ marginTop: 6 }}>Optional holidays do not reduce leave; employees take them with the “Optional / floating holiday” leave type.</div>
      </div>
    </Card>
  );
}
