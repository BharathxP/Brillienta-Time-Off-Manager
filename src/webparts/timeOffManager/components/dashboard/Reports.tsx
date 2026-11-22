import * as React from 'react';
import { DefaultButton, Dropdown, Icon, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Card, Kpi, PageHeader } from '../common/Ui';
import { BarList, Columns } from '../common/Charts';
import { addMonths, fmtNum, leaveYearLabel, leaveYearStart, today } from '../../services/dates';
import { download, toCsv } from '../../services/exporting';
import { lc } from '../../services/workflow';
import { listTitle } from '../../schema/schema';

/** HR reports: usage by type, department and month, balances of everyone, CSV exports and the Power BI connection. */
export function Reports(): JSX.Element {
  const { world, config, year, engine, sp, svc } = useApp();
  const [viewYear, setViewYear] = React.useState(year);
  const [dept, setDept] = React.useState('');
  const departments = world.employees.map((e) => e.department || '').filter((d, i, all) => d && all.indexOf(d) === i).sort();
  const people = world.employees.filter((e) => e.active && (!dept || e.department === dept));
  const emails = people.map((e) => lc(e.email));
  const reqs = world.requests.filter((r) => r.leaveYear === viewYear && emails.indexOf(lc(r.employeeEmail)) >= 0);
  const taken = reqs.filter((r) => r.kind === 'leave' && r.status === 'Approved');
  const t = today();
  const daysTaken = taken.reduce((s, r) => s + r.days, 0);
  const pending = reqs.filter((r) => r.status === 'Pending').length;
  const decided = reqs.filter((r) => r.status === 'Approved' || r.status === 'Rejected');
  const hoursToDecide = decided
    .map((r) => {
      const last = r.steps.filter((s) => s.actedAt).map((s) => Date.parse(s.actedAt as string)).sort().pop();
      return last ? (last - Date.parse(r.created)) / 3600000 : NaN;
    })
    .filter((h) => !isNaN(h) && h >= 0);
  const avgHours = hoursToDecide.length ? hoursToDecide.reduce((a, b) => a + b, 0) / hoursToDecide.length : 0;
  const autoApproved = reqs.filter((r) => r.history.some((h) => /Auto-approved/.test(h.note || '') || /Auto-approved/.test(h.action))).length;
  const awayToday = world.requests.filter((r) => r.kind === 'leave' && r.status === 'Approved' && r.start <= t && r.end >= t && emails.indexOf(lc(r.employeeEmail)) >= 0).length;

  const byType = config.leaveTypes
    .map((ty) => ({ label: ty.name, value: Math.round(taken.filter((r) => r.leaveType === ty.key).reduce((s, r) => s + r.days, 0) * 10) / 10, color: ty.color }))
    .filter((d) => d.value > 0)
    .sort((a, b) => b.value - a.value);
  const byDept = departments
    .map((d) => ({ label: d, value: Math.round(taken.filter((r) => (world.employees.filter((e) => lc(e.email) === lc(r.employeeEmail))[0] || { department: '' }).department === d).reduce((s, r) => s + r.days, 0) * 10) / 10 }))
    .filter((d) => d.value > 0)
    .sort((a, b) => b.value - a.value);
  const start = leaveYearStart(viewYear, config.yearStartMonth);
  const months = Array.from({ length: 12 }, (_, i) => addMonths(start, i).substring(0, 7));
  const monthly = months.map((m) => taken.filter((r) => r.start.substring(0, 7) === m).reduce((s, r) => s + r.days, 0));

  const mainTypes = config.leaveTypes.filter((ty) => ty.active && ty.entitlement.mode !== 'unlimited').slice(0, 6);
  const balanceRows = people.map((e) => ({ e, b: engine.balancesFor(world, e, viewYear) }));

  const exportRequests = (): void => {
    const rows = reqs.map((r) => [r.id, r.employeeName, r.employeeEmail, (world.employees.filter((e) => lc(e.email) === lc(r.employeeEmail))[0] || { department: '' }).department, r.kind, (config.leaveTypes.filter((x) => x.key === r.leaveType)[0] || { name: r.leaveType }).name, r.start, r.end, r.unit, r.hours, r.days, r.status, r.submittedBy ? r.submittedBy.email : '', r.created.substring(0, 10)]);
    download(`leave-requests-${viewYear}.csv`, toCsv(['Id', 'Employee', 'Email', 'Department', 'Kind', 'Leave type', 'Start', 'End', 'Unit', 'Hours', 'Days', 'Status', 'Submitted by', 'Created'], rows));
  };
  const exportBalances = (): void => {
    const types = config.leaveTypes.filter((ty) => ty.active);
    const headers = ['Employee', 'Email', 'Department', 'Grade', 'Start date'].concat(...types.map((ty) => [`${ty.name} entitled`, `${ty.name} taken`, `${ty.name} pending`, `${ty.name} available`]));
    const rows = balanceRows.map(({ e, b }) =>
      [e.name, e.email, e.department || '', e.grade || '', e.hireDate || ''].concat(
        ...types.map((ty) => {
          const x = b.filter((y) => y.leaveType === ty.key)[0];
          return x ? [x.unlimited ? 'unlimited' : fmtNum(x.entitled + x.carried + x.adjusted + x.earned), fmtNum(x.used), fmtNum(x.pending), x.unlimited ? '' : fmtNum(x.available)] : ['', '', '', ''];
        })
      )
    );
    download(`leave-balances-${viewYear}.csv`, toCsv(headers, rows));
  };

  const feed = (key: 'requests' | 'employees' | 'ledger'): string => `${sp.webUrl}/_api/web/lists/getbytitle('${listTitle(svc.prefix, key)}')/items`;

  return (
    <>
      <PageHeader
        eyebrow="Insights"
        title="Leave reports"
        sub="Who took what, when — and what everyone has left."
        actions={
          <>
            <Dropdown ariaLabel="Department" selectedKey={dept} options={[{ key: '', text: 'All departments' }].concat(departments.map((d) => ({ key: d, text: d })))} onChange={(_, o) => o && setDept(String(o.key))} styles={{ root: { width: 180 } }} />
            <Dropdown ariaLabel="Leave year" selectedKey={viewYear} options={[year - 1, year].map((y) => ({ key: y, text: leaveYearLabel(y, config.yearStartMonth) }))} onChange={(_, o) => o && setViewYear(Number(o.key))} styles={{ root: { width: 110 } }} />
          </>
        }
      />
      <div className="to-kpis">
        <Kpi label="Days taken" value={fmtNum(daysTaken)} icon="Airplane" tint={1} />
        <Kpi label="Requests" value={reqs.length} icon="BulletedList" tint={2} />
        <Kpi label="Pending" value={pending} icon="Clock" tint={3} tone={pending ? 'warning' : undefined} />
        <Kpi label="Avg. time to decide" value={avgHours ? (avgHours < 48 ? `${Math.round(avgHours)} h` : `${Math.round(avgHours / 24)} d`) : '—'} icon="Timer" tint={4} />
        <Kpi label="Auto-approved" value={autoApproved} icon="Lightning" tint={5} />
        <Kpi label="Away today" value={awayToday} icon="People" tint={6} />
      </div>
      <div className="to-grid to-grid-2">
        <Card title="Days taken by leave type" icon="BarChartHorizontal">
          <BarList items={byType} />
        </Card>
        <Card title="Days taken by department" icon="Org">
          <BarList items={byDept} />
        </Card>
      </div>
      <Card title="Days taken by month" icon="BarChartVertical" style={{ marginTop: 16 }}>
        <Columns labels={months.map((m) => ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][Number(m.substring(5)) - 1])} values={monthly} />
      </Card>
      <Card
        title="Balances"
        icon="Table"
        style={{ marginTop: 16 }}
        actions={
          <span className="to-row">
            <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export balances" onClick={exportBalances} />
            <DefaultButton iconProps={{ iconName: 'Download' }} text="Export requests" onClick={exportRequests} />
          </span>
        }
      >
        <div className="to-table-wrap">
          <table className="to-table">
            <thead>
              <tr>
                <th>Employee</th>
                {mainTypes.map((ty) => (
                  <th key={ty.key} className="to-number">
                    {ty.name}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {balanceRows.slice(0, 200).map(({ e, b }) => (
                <tr key={e.email}>
                  <td>
                    <strong>{e.name}</strong>
                    <div className="to-small to-muted">{[e.department, e.grade].filter((x) => !!x).join(' · ')}</div>
                  </td>
                  {mainTypes.map((ty) => {
                    const x = b.filter((y) => y.leaveType === ty.key)[0];
                    return (
                      <td key={ty.key} className="to-number">
                        {x ? (x.eligible ? <span className={x.available < 0 ? 'to-danger' : ''}>{fmtNum(x.available)}</span> : <span className="to-muted">n/a</span>) : ''}
                        {x && x.used > 0 && <div className="to-small to-muted">{fmtNum(x.used)} taken</div>}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {balanceRows.length > 200 && <div className="to-small to-muted" style={{ marginTop: 8 }}>Showing 200 of {balanceRows.length}; the export has everyone.</div>}
      </Card>
      <Card title="Connect Power BI" icon="PowerBILogo" style={{ marginTop: 16 }}>
        <p className="to-muted" style={{ marginTop: 0 }}>
          In Power BI Desktop choose <b>Get data › SharePoint Online list</b> (2.0), enter the site address below and pick the lists. Or use <b>OData feed</b> with the addresses. Refresh on a schedule in the Power BI service.
        </p>
        <div className="to-stack">
          <TextField label="Site address" readOnly value={sp.webUrl} />
          <TextField label="Requests (OData)" readOnly value={feed('requests')} />
          <TextField label="Employees (OData)" readOnly value={feed('employees')} />
          <TextField label="Ledger — adjustments, carry-forward, comp off (OData)" readOnly value={feed('ledger')} />
        </div>
        <p className="to-small to-muted">
          <Icon iconName="Info" /> Useful columns: Requests › Status, LeaveTypeName, Department, StartOn, EndOn, Days, Hours, LeaveYear. The balance export above is the easiest source for a balances report.
        </p>
      </Card>
    </>
  );
}
