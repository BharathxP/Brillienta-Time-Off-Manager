import * as React from 'react';
import { DefaultButton, Dropdown, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, SearchBox, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IEmployee } from '../../models/types';
import { Card, EmployeePicker, Kpi, PageHeader, PersonLine, useBusy, useConfirm, NumberBox } from '../common/Ui';
import { BalanceCards, balanceLines } from '../requests/BalanceCards';
import { IsoDatePicker } from '../requests/RequestForm';
import { fmtDate, fmtNum, isIsoDate, leaveYearLabel } from '../../services/dates';
import { lc } from '../../services/workflow';
import { GraphService } from '../../services/GraphService';
import { download, downloadXlsx, parseCsv, toCsv } from '../../services/exporting';
import { serviceYears } from '../../services/balances';
import { SETTING_KEYS } from '../../services/SettingsStore';

const CSV_HEADERS = ['Name', 'Email', 'Manager email', 'Department', 'Job title', 'Grade', 'Start date', 'Work calendar', 'Location', 'Gender', 'Leaving date', 'Active'];
const OPENING_HEADERS = ['Email', 'Leave type', 'Available'];

/** HR: the employee list (who takes leave, their manager, grade, start date, calendar) and everyone's balances. */
export function EmployeesPage(): JSX.Element {
  const { world, svc, reload, graph, config, settings, me, licensed, engine, year } = useApp();
  const [q, setQ] = React.useState('');
  const [edit, setEdit] = React.useState<IEmployee | undefined>();
  const [balances, setBalances] = React.useState<IEmployee | undefined>();
  const [busy, run] = useBusy();
  const [log, setLog] = React.useState<string[]>([]);
  const [ask, dialog] = useConfirm();
  const fileRef = React.useRef<HTMLInputElement>(null);
  const openingRef = React.useRef<HTMLInputElement>(null);
  const [bulk, setBulk] = React.useState(false);
  const norm = q.trim().toLowerCase();
  const rows = world.employees
    .filter((e) => !norm || [e.name, e.email, e.department, e.jobTitle, e.grade, e.managerEmail].join(' ').toLowerCase().indexOf(norm) >= 0)
    .sort((a, b) => Number(b.active) - Number(a.active) || a.name.localeCompare(b.name));
  const lastImport = settings.get(SETTING_KEYS.lastImportAt);
  const noManager = world.employees.filter((e) => e.active && !e.managerEmail).length;
  const noHire = world.employees.filter((e) => e.active && !e.hireDate).length;

  const importGraph = (): Promise<boolean> =>
    run(async () => {
      if (!graph) throw new Error('Microsoft Graph is not available here. Approve the API permissions in the SharePoint admin center (API access).');
      setLog(['Reading people from Microsoft Entra ID…']);
      const users = await graph.loadUsers((n) => setLog([`Read ${n} people…`]));
      const plan = GraphService.merge(world.employees, users);
      const out: string[] = [`${users.length} people in Microsoft 365: ${plan.add.length} new, ${plan.update.length} changed.`];
      for (const e of plan.add) await svc.saveEmployee(e);
      for (const e of plan.update) await svc.saveEmployee(e);
      await settings.set(SETTING_KEYS.lastImportAt, new Date().toISOString());
      await svc.audit('Imported employees', { email: me.email, name: me.title }, 'Microsoft 365', out[0]);
      setLog(out);
      await reload();
    }, 'Employees imported from Microsoft 365.');

  const importCsv = async (file: File): Promise<void> => {
    await run(async () => {
      const rowsIn = parseCsv(await file.text());
      const head = (rowsIn.shift() || []).map((h) => h.toLowerCase());
      const col = (name: string): number => head.indexOf(name.toLowerCase());
      let added = 0;
      let updated = 0;
      for (const r of rowsIn) {
        const email = r[col('Email')];
        if (!email || email.indexOf('@') < 0) continue;
        const cur = world.employees.filter((e) => lc(e.email) === lc(email))[0];
        const val = (name: string, fallback?: string): string | undefined => (col(name) >= 0 && r[col(name)] !== undefined && r[col(name)] !== '' ? r[col(name)] : fallback);
        const hire = val('Start date', cur && cur.hireDate);
        const e: IEmployee = {
          ...(cur || { active: true }),
          email,
          name: val('Name', cur ? cur.name : email) as string,
          managerEmail: val('Manager email', cur && cur.managerEmail),
          department: val('Department', cur && cur.department),
          jobTitle: val('Job title', cur && cur.jobTitle),
          grade: val('Grade', cur && cur.grade),
          hireDate: hire && isIsoDate(hire) ? hire : cur && cur.hireDate,
          calendar: val('Work calendar', cur && cur.calendar),
          location: val('Location', cur && cur.location),
          gender: val('Gender', cur && cur.gender),
          leavingDate: (() => {
            const v = val('Leaving date', cur && cur.leavingDate);
            return v && isIsoDate(v) ? v : cur && cur.leavingDate;
          })(),
          active: col('Active') >= 0 && r[col('Active')] ? !/^(no|false|0)$/i.test(r[col('Active')]) : cur ? cur.active : true
        };
        await svc.saveEmployee(e);
        if (cur) updated++;
        else added++;
      }
      setLog([`CSV: ${added} added, ${updated} updated.`]);
      await svc.audit('Imported employees', { email: me.email, name: me.title }, 'CSV', `${added} added, ${updated} updated`);
      await reload();
    }, 'CSV imported.');
  };

  const exportCsv = (): void => {
    download('employees.csv', toCsv(CSV_HEADERS, world.employees.map((e) => [e.name, e.email, e.managerEmail, e.department, e.jobTitle, e.grade, e.hireDate, e.calendar, e.location, e.gender, e.leavingDate, e.active ? 'Yes' : 'No'])));
  };
  const exportXlsx = (): void => {
    downloadXlsx('employees.xlsx', [{ name: 'Employees', columns: CSV_HEADERS, rows: world.employees.map((e) => [e.name, e.email, e.managerEmail, e.department, e.jobTitle, e.grade, e.hireDate, e.calendar, e.location, e.gender, e.leavingDate, e.active ? 'Yes' : 'No']) }]);
  };

  const importOpening = async (file: File): Promise<void> => {
    await run(async () => {
      const rowsIn = parseCsv(await file.text());
      const head = (rowsIn.shift() || []).map((h) => h.toLowerCase());
      const ci = (n: string): number => head.indexOf(n.toLowerCase());
      if (ci('email') < 0 || ci('leave type') < 0 || ci('available') < 0) throw new Error('The file needs the columns Email, Leave type and Available.');
      const res = await engine.importOpeningBalances(world, rowsIn.map((r) => ({ email: r[ci('email')], leaveType: r[ci('leave type')], available: Number((r[ci('available')] || '').replace(',', '.')) })), year, { email: me.email, name: me.title });
      setLog([`Opening balances: ${res.done} set for ${leaveYearLabel(year, config.yearStartMonth)}.`].concat(res.skipped.length ? [`Skipped (unknown person or leave type): ${res.skipped.slice(0, 20).join(', ')}`] : []));
      await reload();
    }, 'Opening balances imported.');
  };

  const carry = async (): Promise<void> => {
    const ok = await ask(`Carry forward ${leaveYearLabel(year - 1, config.yearStartMonth)} into ${leaveYearLabel(year, config.yearStartMonth)}?`, 'What is left of each balance moves into this leave year, up to each type’s carry-forward limit. People already carried forward are skipped.', 'Carry forward');
    if (!ok) return;
    const lines: string[] = [];
    await run(async () => {
      const n = await engine.carryForward(world, year - 1, { email: me.email, name: me.title }, (m) => lines.push(m));
      setLog([`${n} balance(s) carried forward.`].concat(lines.slice(0, 30)));
      await reload();
    });
  };

  return (
    <>
      {dialog}
      <PageHeader
        eyebrow="Manage"
        title="Employees & balances"
        sub={lastImport ? `Last import from Microsoft 365: ${fmtDate(lastImport.substring(0, 10))}` : 'Import people from Microsoft 365 (Entra ID), a CSV file, or add them one by one.'}
        actions={
          <>
            <DefaultButton iconProps={{ iconName: 'CloudDownload' }} text="Import from Microsoft 365" disabled={busy} onClick={() => void importGraph()} />
            <DefaultButton
              iconProps={{ iconName: 'ExcelDocument' }}
              text="CSV"
              menuProps={{
                items: [
                  { key: 'imp', text: 'Import CSV…', iconProps: { iconName: 'Upload' }, onClick: () => {
                      if (fileRef.current) fileRef.current.click();
                    } },
                  { key: 'exp', text: 'Export CSV', iconProps: { iconName: 'Download' }, onClick: () => {
                      exportCsv();
                    } },
                  { key: 'xlsx', text: 'Export Excel', iconProps: { iconName: 'ExcelDocument' }, onClick: () => {
                      exportXlsx();
                    } },
                  { key: 'div', text: '-', itemType: 1 },
                  { key: 'open', text: 'Import opening balances…', iconProps: { iconName: 'Upload' }, onClick: () => {
                      if (openingRef.current) openingRef.current.click();
                    } },
                  { key: 'opentpl', text: 'Opening balances template', iconProps: { iconName: 'Page' }, onClick: () => {
                      download('opening-balances-template.csv', toCsv(OPENING_HEADERS, world.employees.filter((e) => e.active).slice(0, 3).map((e) => [e.email, 'Annual leave', 12])));
                    } },
                  { key: 'tpl', text: 'Download template', iconProps: { iconName: 'Page' }, onClick: () => {
                      download('employees-template.csv', toCsv(CSV_HEADERS, [['Asha Rao', 'asha@contoso.com', 'manager@contoso.com', 'Sales', 'Account manager', 'L2', '2024-04-01', 'default', 'Bengaluru', 'Female', 'Yes']]));
                    } }
                ]
              }}
            />
            <DefaultButton iconProps={{ iconName: 'AddTo' }} text="Bulk adjust" disabled={busy} onClick={() => setBulk(true)} />
            <DefaultButton iconProps={{ iconName: 'Forward' }} text="Carry forward" disabled={busy} onClick={() => void carry()} />
            <PrimaryButton iconProps={{ iconName: 'AddFriend' }} text="Add employee" onClick={() => setEdit({ email: '', name: '', active: true })} />
          </>
        }
      />
      <input ref={fileRef} type="file" accept=".csv,text/csv" style={{ display: 'none' }} onChange={(e) => e.target.files && e.target.files[0] && void importCsv(e.target.files[0])} />
      <input ref={openingRef} type="file" accept=".csv,text/csv" style={{ display: 'none' }} onChange={(e) => e.target.files && e.target.files[0] && void importOpening(e.target.files[0])} />
      {log.length > 0 && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline onDismiss={() => setLog([])} styles={{ root: { marginBottom: 12 } }}>
          {log.map((l, i) => (
            <div key={i}>{l}</div>
          ))}
        </MessageBar>
      )}
      <div className="to-kpis">
        <Kpi label="Active employees" value={world.employees.filter((e) => e.active).length} icon="People" tint={1} />
        <Kpi label="Licensed" value={licensed.size} icon="Permissions" tint={2} />
        <Kpi label="Without a manager" value={noManager} icon="Org" tint={3} tone={noManager ? 'warning' : undefined} />
        <Kpi label="Without a start date" value={noHire} icon="Calendar" tint={4} tone={noHire ? 'warning' : undefined} />
      </div>
      <Card>
        <SearchBox placeholder="Search by name, email, department, grade…" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { maxWidth: 420, marginBottom: 12 } }} />
        <div className="to-table-wrap">
          <table className="to-table">
            <thead>
              <tr>
                <th>Employee</th>
                <th>Manager</th>
                <th>Department</th>
                <th>Grade</th>
                <th>Start date</th>
                <th>Calendar</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {rows.map((e) => {
                const mgr = world.employees.filter((x) => lc(x.email) === lc(e.managerEmail))[0];
                const cal = config.calendars.filter((c) => c.key === (e.calendar || config.defaultCalendar))[0];
                return (
                  <tr key={e.email} className={e.active ? '' : 'is-inactive'}>
                    <td>
                      <PersonLine email={e.email} name={e.name} sub={e.jobTitle || e.email} size="sm" />
                      {!e.active && <span className="to-badge is-muted">Inactive</span>}
                      {e.active && !licensed.has(lc(e.email)) && <span className="to-badge is-warn">Over licence</span>}
                      {e.leavingDate && <span className="to-badge is-muted">Leaving {fmtDate(e.leavingDate)}</span>}
                    </td>
                    <td>{mgr ? mgr.name : e.managerEmail || <span className="to-muted">—</span>}</td>
                    <td>{e.department || '—'}</td>
                    <td>{e.grade || '—'}</td>
                    <td>
                      {e.hireDate ? fmtDate(e.hireDate) : '—'}
                      {e.hireDate && <div className="to-small to-muted">{serviceYears(e.hireDate, new Date().toISOString().substring(0, 10))} yr</div>}
                    </td>
                    <td>{cal ? cal.name : e.calendar}</td>
                    <td style={{ whiteSpace: 'nowrap' }}>
                      <DefaultButton text="Balances" onClick={() => setBalances(e)} />
                      <DefaultButton text="Edit" onClick={() => setEdit(e)} styles={{ root: { marginLeft: 6 } }} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
      {edit && <EmployeeEditor employee={edit} onClose={() => setEdit(undefined)} />}
      {balances && <BalancePanel employee={balances} onClose={() => setBalances(undefined)} />}
      {bulk && <BulkAdjust onClose={() => setBulk(false)} />}
    </>
  );
}

function EmployeeEditor(props: { employee: IEmployee; onClose: () => void }): JSX.Element {
  const { svc, reload, config, me } = useApp();
  const [e, setE] = React.useState<IEmployee>(props.employee);
  const [busy, run] = useBusy();
  const set = (patch: Partial<IEmployee>): void => setE((x) => ({ ...x, ...patch }));
  const save = (): Promise<boolean> =>
    run(async () => {
      if (!e.name.trim() || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.email)) throw new Error('Name and a valid email are needed.');
      await svc.saveEmployee(e);
      await svc.audit(e.id ? 'Employee updated' : 'Employee added', { email: me.email, name: me.title }, e.email, JSON.stringify({ grade: e.grade, manager: e.managerEmail, hireDate: e.hireDate, calendar: e.calendar }));
      await reload();
      props.onClose();
    }, 'Saved.');
  return (
    <Panel isOpen type={PanelType.medium} onDismiss={props.onClose} headerText={e.id ? e.name : 'Add employee'} isFooterAtBottom onRenderFooterContent={() => (
      <div className="to-row">
        <PrimaryButton text="Save" disabled={busy} onClick={() => void save()} />
        <DefaultButton text="Cancel" onClick={props.onClose} />
      </div>
    )}>
      <div className="to-stack">
        <TextField label="Name" required value={e.name} onChange={(_, v) => set({ name: v || '' })} />
        <TextField label="Email" required value={e.email} onChange={(_, v) => set({ email: (v || '').trim() })} disabled={!!e.id} />
        <EmployeePicker label="Manager (approves their leave)" value={e.managerEmail} onChange={(m) => set({ managerEmail: m ? m.email : '' })} filter={(m) => lc(m.email) !== lc(e.email)} />
        <div className="to-form-grid">
          <TextField label="Department" value={e.department || ''} onChange={(_, v) => set({ department: v || '' })} />
          <TextField label="Job title" value={e.jobTitle || ''} onChange={(_, v) => set({ jobTitle: v || '' })} />
          <Dropdown label="Level / grade" selectedKey={e.grade || ''} options={[{ key: '', text: '—' }].concat(config.grades.map((g) => ({ key: g, text: g })))} onChange={(_, o) => o && set({ grade: String(o.key) })} />
          <IsoDatePicker label="Start date (tenure, proration)" value={e.hireDate || ''} onChange={(d) => set({ hireDate: d })} />
          <Dropdown label="Work calendar" selectedKey={e.calendar || config.defaultCalendar} options={config.calendars.map((c) => ({ key: c.key, text: c.name }))} onChange={(_, o) => o && set({ calendar: String(o.key) })} />
          <TextField label="Location" value={e.location || ''} onChange={(_, v) => set({ location: v || '' })} />
          <IsoDatePicker label="Leaving date (last working day)" value={e.leavingDate || ''} onChange={(d) => set({ leavingDate: d })} />
          <Dropdown label="Gender (for gender-specific leave)" selectedKey={e.gender || ''} options={[{ key: '', text: 'Not set' }, { key: 'Female', text: 'Female' }, { key: 'Male', text: 'Male' }, { key: 'Other', text: 'Other' }]} onChange={(_, o) => o && set({ gender: String(o.key) })} />
        </div>
        {e.leavingDate && <DefaultButton text="Clear leaving date" onClick={() => set({ leavingDate: '' })} styles={{ root: { alignSelf: 'flex-start' } }} />}
        <Toggle label="Active" checked={e.active} onChange={(_, c) => set({ active: !!c })} />
      </div>
    </Panel>
  );
}

/** One person's balances with the calculation, ledger entries, adjustments and comp-off grants. */
export function BalancePanel(props: { employee: IEmployee; onClose: () => void; readOnly?: boolean }): JSX.Element {
  const { config, world, engine, reload, me, year, openRequest, svc } = useApp();
  const e = props.employee;
  const [viewYear, setViewYear] = React.useState(year);
  const [type, setType] = React.useState(config.leaveTypes.filter((t) => t.active)[0].key);
  const [days, setDays] = React.useState(1);
  const [note, setNote] = React.useState('');
  const [busy, run] = useBusy();
  const t = config.leaveTypes.filter((x) => x.key === type)[0];
  const b = engine.balancesFor(world, e, viewYear).filter((x) => x.leaveType === type)[0];
  const entries = world.ledger.filter((l) => lc(l.employeeEmail) === lc(e.email) && (l.leaveYear === viewYear || l.kind === 'compoff')).sort((a, b2) => (a.effective < b2.effective ? 1 : -1));
  const adjust = (): Promise<boolean> =>
    run(async () => {
      if (!note.trim()) throw new Error('Add a note: it is shown to HR and in the audit log.');
      await engine.adjust(world, e, type, viewYear, days, note, { email: me.email, name: me.title }, t && t.entitlement.mode === 'earned' ? 'compoff' : 'adjustment');
      await reload();
      setNote('');
    }, 'Balance updated.');
  return (
    <Panel isOpen type={PanelType.large} onDismiss={props.onClose} headerText={`${e.name} · balances`}>
      <div className="to-row" style={{ margin: '8px 0 14px' }}>
        <Dropdown ariaLabel="Leave year" selectedKey={viewYear} options={[year - 1, year, year + 1].map((y) => ({ key: y, text: `Leave year ${leaveYearLabel(y, config.yearStartMonth)}` }))} onChange={(_, o) => o && setViewYear(Number(o.key))} styles={{ root: { width: 200 } }} />
        <span className="to-spacer" />
        <DefaultButton
          iconProps={{ iconName: 'Add' }}
          text="Request on their behalf"
          onClick={() => {
            props.onClose();
            openRequest({ employeeEmail: e.email });
          }}
        />
      </div>
      <BalanceCards employee={e} year={viewYear} compact showAll onRequest={setType} />
      <div className="to-grid to-grid-2" style={{ marginTop: 16 }}>
        {props.readOnly ? (
          <Card title="This year" icon="Calendar">
            <div className="to-stack">
              {world.requests
                .filter((r) => lc(r.employeeEmail) === lc(e.email) && r.leaveYear === viewYear && r.kind === 'leave' && r.status !== 'Cancelled')
                .sort((a, b2) => (a.start < b2.start ? 1 : -1))
                .slice(0, 15)
                .map((r) => (
                  <div key={r.id} className="to-small">
                    {fmtDate(r.start)}{r.end !== r.start ? ` – ${fmtDate(r.end)}` : ''} · {(config.leaveTypes.filter((x) => x.key === r.leaveType)[0] || { name: r.leaveType }).name} · {fmtNum(r.days)} d · <span className="to-muted">{r.status}</span>
                  </div>
                ))}
            </div>
          </Card>
        ) : (
        <Card title="Adjust a balance" icon="Edit">
          <div className="to-stack">
            <Dropdown label="Leave type" selectedKey={type} options={config.leaveTypes.filter((x) => x.active && x.entitlement.mode !== 'unlimited').map((x) => ({ key: x.key, text: x.name }))} onChange={(_, o) => o && setType(String(o.key))} />
            {b && (
              <div className="to-calc">
                {balanceLines(b).map(([k, v]) => (
                  <div key={k}>
                    <span>{k}</span>
                    <b>{v}</b>
                  </div>
                ))}
              </div>
            )}
            <NumberBox label={t && t.entitlement.mode === 'earned' ? 'Grant comp-off days (use − to remove)' : 'Days to add (use − to deduct)'} value={String(days)} step={0.5} min={-365} max={365} onChange={(_, v) => setDays(Number(v) || 0)} />
            <TextField label="Note" value={note} onChange={(_, v) => setNote(v || '')} placeholder="e.g. Joining bonus, correction, special grant" />
            <PrimaryButton text="Save adjustment" disabled={busy || !days} onClick={() => void adjust()} />
          </div>
        </Card>
        )}
        <Card title="Ledger" icon="List">
          {entries.length ? (
            <table className="to-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Type</th>
                  <th>Kind</th>
                  <th className="to-number">Days</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {entries.map((l) => (
                  <tr key={l.id}>
                    <td>{fmtDate(l.effective)}</td>
                    <td>
                      {(config.leaveTypes.filter((x) => x.key === l.leaveType)[0] || { name: l.leaveType }).name}
                      <div className="to-small to-muted">{l.note}{l.expiresOn ? ` · expires ${fmtDate(l.expiresOn)}` : ''}</div>
                    </td>
                    <td>{l.kind}</td>
                    <td className="to-number">{l.days > 0 ? '+' : ''}{fmtNum(l.days)}</td>
                    <td>
                      {l.id && l.kind !== 'compoff' && !props.readOnly && (
                        <DefaultButton
                          text="Undo"
                          disabled={busy}
                          onClick={() =>
                            void run(async () => {
                              await svc.deleteLedger(l.id as number);
                              await svc.audit('Ledger entry removed', { email: me.email, name: me.title }, e.email, `${l.kind} ${l.days} ${l.leaveType}`);
                              await reload();
                            }, 'Removed.')
                          }
                        />
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <div className="to-muted to-small">No adjustments, carry-forward or comp-off credits.</div>
          )}
        </Card>
      </div>
    </Panel>
  );
}

/** The same adjustment for everyone in a department / grade (e.g. an extra company day off). */
function BulkAdjust(props: { onClose: () => void }): JSX.Element {
  const { world, config, engine, reload, me, year } = useApp();
  const [dept, setDept] = React.useState('');
  const [grade, setGrade] = React.useState('');
  const [type, setType] = React.useState(config.leaveTypes.filter((t) => t.active && t.entitlement.mode !== 'unlimited')[0].key);
  const [days, setDays] = React.useState(1);
  const [note, setNote] = React.useState('');
  const [busy, run] = useBusy();
  const departments = world.employees.map((e) => e.department || '').filter((d, i, all) => d && all.indexOf(d) === i).sort();
  const people = world.employees.filter((e) => e.active && (!dept || e.department === dept) && (!grade || e.grade === grade));
  return (
    <Panel isOpen type={PanelType.medium} onDismiss={props.onClose} headerText="Bulk adjust balances" isFooterAtBottom onRenderFooterContent={() => (
      <div className="to-row">
        <PrimaryButton text={`Apply to ${people.length} people`} disabled={busy || !days || !note.trim() || !people.length} onClick={() => void run(async () => {
          await engine.bulkAdjust(world, people, type, year, days, note, { email: me.email, name: me.title });
          await reload();
          props.onClose();
        }, `Balances of ${people.length} people updated.`)} />
        <DefaultButton text="Cancel" onClick={props.onClose} />
      </div>
    )}>
      <div className="to-stack">
        <Dropdown label="Department" selectedKey={dept} options={[{ key: '', text: 'Everyone' }].concat(departments.map((d) => ({ key: d, text: d })))} onChange={(_, o) => o && setDept(String(o.key))} />
        <Dropdown label="Grade" selectedKey={grade} options={[{ key: '', text: 'Every grade' }].concat(config.grades.map((g) => ({ key: g, text: g })))} onChange={(_, o) => o && setGrade(String(o.key))} />
        <Dropdown label="Leave type" selectedKey={type} options={config.leaveTypes.filter((t) => t.active && t.entitlement.mode !== 'unlimited').map((t) => ({ key: t.key, text: t.name }))} onChange={(_, o) => o && setType(String(o.key))} />
        <NumberBox label={`Days for leave year ${leaveYearLabel(year, config.yearStartMonth)} (use − to deduct)`} value={String(days)} step={0.5} min={-365} max={365} onChange={(_, v) => setDays(Number(v) || 0)} />
        <TextField label="Note" value={note} onChange={(_, v) => setNote(v || '')} placeholder="e.g. Company anniversary day" />
        <div className="to-small to-muted">{people.length} active people match. Each gets a ledger entry, so it can be undone person by person.</div>
      </div>
    </Panel>
  );
}
