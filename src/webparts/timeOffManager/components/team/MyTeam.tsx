import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IEmployee } from '../../models/types';
import { Card, Empty, Kpi, PageHeader, PersonLine, Segmented, StatusBadge, TypeChip } from '../common/Ui';
import { BalancePanel } from '../admin/EmployeesPage';
import { teamOf } from '../calendar/TeamCalendar';
import { addDays, fmtNum, fmtRange, today } from '../../services/dates';
import { lc, waitingFor } from '../../services/workflow';
import { downloadXlsx } from '../../services/exporting';

/** Managers: the people below them — balances, upcoming and pending leave, and requesting on their behalf. */
export function MyTeam(): JSX.Element {
  const { world, me, config, engine, year, openRequest, navigate } = useApp();
  const [scope, setScope] = React.useState<'direct' | 'all'>('direct');
  const [open, setOpen] = React.useState<IEmployee | undefined>();
  const all = teamOf(world.employees, me.email, 'reports').filter((e) => lc(e.email) !== lc(me.email));
  const people = (scope === 'direct' ? all.filter((e) => lc(e.managerEmail) === lc(me.email)) : all).sort((a, b) => a.name.localeCompare(b.name));
  const emails = people.map((e) => lc(e.email));
  const t = today();
  const horizon = addDays(t, 30);
  const reqs = world.requests.filter((r) => emails.indexOf(lc(r.employeeEmail)) >= 0);
  const awayToday = reqs.filter((r) => r.kind === 'leave' && r.status === 'Approved' && r.start <= t && r.end >= t);
  const upcoming = reqs.filter((r) => r.kind === 'leave' && (r.status === 'Approved' || r.status === 'Pending') && r.end >= t && r.start <= horizon).sort((a, b) => (a.start < b.start ? -1 : 1));
  const waiting = waitingFor(world.requests, me.email).filter((r) => emails.indexOf(lc(r.employeeEmail)) >= 0);
  const key = config.leaveTypes.filter((x) => x.active && x.entitlement.mode !== 'unlimited' && !x.eligibility.genders.length).slice(0, 4);
  const typeOf = (k: string): (typeof config.leaveTypes)[0] | undefined => config.leaveTypes.filter((x) => x.key === k)[0];
  const exportXlsx = (): void => {
    const types = config.leaveTypes.filter((x) => x.active && x.entitlement.mode !== 'unlimited');
    const rows = people.map((e) => {
      const b = engine.balancesFor(world, e, year);
      const values: Array<string | number | null> = [e.name, e.email, e.jobTitle || '', e.department || ''];
      types.forEach((ty) => {
        const x = b.filter((y) => y.leaveType === ty.key)[0];
        values.push(x && x.eligible ? x.available : null);
      });
      return values;
    });
    downloadXlsx('team-balances.xlsx', [{ name: 'Team balances', columns: ['Name', 'Email', 'Job title', 'Department'].concat(types.map((ty) => `${ty.name} available`)), rows }]);
  };
  if (!all.length) {
    return (
      <>
        <PageHeader eyebrow="My team" title="My team" />
        <Empty icon="People" title="Nobody reports to you in the employee list." />
      </>
    );
  }
  return (
    <>
      <PageHeader
        eyebrow="My team"
        title="My team"
        sub="Balances and plans of the people who report to you."
        actions={
          <>
            <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export balances" onClick={exportXlsx} />
            {config.onBehalf.managers && <PrimaryButton iconProps={{ iconName: 'Add' }} text="Request for someone" onClick={() => openRequest({ employeeEmail: people[0] && people[0].email })} />}
          </>
        }
      />
      <div className="to-kpis">
        <Kpi label="People" value={people.length} icon="People" tint={1} />
        <Kpi label="Away today" value={awayToday.length} icon="Airplane" tint={2} onClick={() => navigate('calendar')} />
        <Kpi label="Leave in the next 30 days" value={upcoming.length} icon="Calendar" tint={3} />
        <Kpi label="Waiting for you" value={waiting.length} icon="CheckList" tint={4} tone={waiting.length ? 'warning' : undefined} onClick={() => navigate('approvals')} />
      </div>
      <div className="to-grid">
        <Card>
          <div className="to-row" style={{ marginBottom: 12 }}>
            <Segmented value={scope} onChange={setScope} options={[{ key: 'direct', text: 'Direct reports' }, { key: 'all', text: `Everyone below me (${all.length})` }]} />
          </div>
          <div className="to-table-wrap">
            <table className="to-table">
              <thead>
                <tr>
                  <th>Employee</th>
                  {key.map((ty) => (
                    <th key={ty.key} className="to-number">
                      {ty.name}
                    </th>
                  ))}
                  <th>Next leave</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {people.map((e) => {
                  const b = engine.balancesFor(world, e, year);
                  const next = upcoming.filter((r) => lc(r.employeeEmail) === lc(e.email))[0];
                  return (
                    <tr key={e.email}>
                      <td>
                        <PersonLine email={e.email} name={e.name} sub={e.jobTitle} size="sm" />
                      </td>
                      {key.map((ty) => {
                        const x = b.filter((y) => y.leaveType === ty.key)[0];
                        return (
                          <td key={ty.key} className="to-number">
                            {x && x.eligible ? <span className={x.available < 0 ? 'to-danger' : ''}>{fmtNum(x.available)}</span> : <span className="to-muted">–</span>}
                          </td>
                        );
                      })}
                      <td className="to-small">{next ? `${fmtRange(next.start, next.end)}${next.status === 'Pending' ? ' (pending)' : ''}` : <span className="to-muted">—</span>}</td>
                      <td style={{ whiteSpace: 'nowrap' }}>
                        <DefaultButton text="Details" onClick={() => setOpen(e)} />
                        {config.onBehalf.managers && <DefaultButton text="Request" onClick={() => openRequest({ employeeEmail: e.email })} styles={{ root: { marginLeft: 6 } }} />}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
        <Card title="Coming up (30 days)" icon="Calendar">
          {upcoming.length ? (
            upcoming.slice(0, 12).map((r) => (
              <div key={r.id} className="to-list-row is-static">
                <span className="to-list-main">
                  <strong>{r.employeeName}</strong>
                  <div className="to-small to-muted">{fmtRange(r.start, r.end)}</div>
                </span>
                <TypeChip type={typeOf(r.leaveType)} label={(typeOf(r.leaveType) || { code: '' }).code} />
                <StatusBadge status={r.status} />
              </div>
            ))
          ) : (
            <Empty icon="Airplane" title="No leave planned." />
          )}
        </Card>
      </div>
      {open && <BalancePanel employee={open} readOnly onClose={() => setOpen(undefined)} />}
    </>
  );
}
