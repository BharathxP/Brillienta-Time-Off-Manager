import * as React from 'react';
import { DefaultButton, Dropdown, Icon, IconButton, TooltipHost } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IEmployee, ILeaveRequest } from '../../models/types';
import { Avatar, Card, Empty, PageHeader } from '../common/Ui';
import { RequestDetails } from '../requests/RequestDetails';
import { addMonths, calendarFor, eachDay, holidayOn, isWorkingDay, MONTH_NAMES, monthEnd, monthStart, today, weekday, WEEKDAY_SHORT } from '../../services/dates';
import { lc } from '../../services/workflow';
import { download, toIcs } from '../../services/exporting';

type Scope = 'team' | 'reports' | 'department' | 'everyone';

/** People in the viewer's team: their manager's team, their own reports and everyone below them. */
export function teamOf(employees: IEmployee[], email: string, scope: Scope, department?: string): IEmployee[] {
  const active = employees.filter((e) => e.active);
  const me = active.filter((e) => lc(e.email) === lc(email))[0];
  if (scope === 'everyone') return active;
  if (scope === 'department') return active.filter((e) => e.department && e.department === (department || (me && me.department)));
  const below = (root: string): IEmployee[] => {
    const out: IEmployee[] = [];
    const queue = [lc(root)];
    while (queue.length && out.length < 2000) {
      const m = queue.shift() as string;
      active.forEach((e) => {
        if (lc(e.managerEmail) === m && out.indexOf(e) < 0) {
          out.push(e);
          queue.push(lc(e.email));
        }
      });
    }
    return out;
  };
  if (scope === 'reports') return (me ? [me] : []).concat(below(email));
  const peers = me && me.managerEmail ? active.filter((e) => lc(e.managerEmail) === lc(me.managerEmail) || lc(e.email) === lc(me.managerEmail)) : me ? [me] : [];
  const mine = below(email);
  const all = peers.concat(mine);
  return all.filter((e, i) => all.indexOf(e) === i);
}

export function TeamCalendar(): JSX.Element {
  const { world, me, config, isHr, isManager, myEmployee } = useApp();
  const [month, setMonth] = React.useState(monthStart(today()));
  const [scope, setScope] = React.useState<Scope>(isHr && config.teamCalendar.scope === 'everyone' ? 'everyone' : (config.teamCalendar.scope as Scope) || 'team');
  const [dept, setDept] = React.useState((myEmployee && myEmployee.department) || '');
  const [typeFilter, setTypeFilter] = React.useState('');
  const [open, setOpen] = React.useState<ILeaveRequest | undefined>();
  const days = eachDay(month, monthEnd(month));
  const people = teamOf(world.employees, me.email, scope, dept).sort((a, b) => a.name.localeCompare(b.name));
  const emails = people.map((p) => lc(p.email));
  const shown = world.requests.filter(
    (r) => r.kind === 'leave' && (r.status === 'Approved' || r.status === 'Pending' || r.status === 'CancelRequested') && r.end >= month && r.start <= monthEnd(month) && emails.indexOf(lc(r.employeeEmail)) >= 0 && (!typeFilter || r.leaveType === typeFilter)
  );
  const byPerson: Record<string, ILeaveRequest[]> = {};
  shown.forEach((r) => (byPerson[lc(r.employeeEmail)] = (byPerson[lc(r.employeeEmail)] || []).concat(r)));
  const departments = world.employees.map((e) => e.department || '').filter((d, i, all) => d && all.indexOf(d) === i).sort();
  const cal = calendarFor(config, myEmployee);
  const t = today();
  const offToday = world.requests.filter((r) => r.kind === 'leave' && r.status === 'Approved' && r.start <= t && r.end >= t && emails.indexOf(lc(r.employeeEmail)) >= 0);
  const typeOf = (k: string): { name: string; color: string; icon: string } => config.leaveTypes.filter((x) => x.key === k)[0] || { name: k, color: '#8a959c', icon: 'Calendar' };
  const canSee = (r: ILeaveRequest): boolean => isHr || isManager || lc(r.employeeEmail) === lc(me.email) || config.teamCalendar.showLeaveType;
  const canOpen = (r: ILeaveRequest): boolean => isHr || lc(r.employeeEmail) === lc(me.email) || r.steps.some((s) => s.approvers.some((a) => lc(a.email) === lc(me.email)));
  const usedTypes = shown.map((r) => r.leaveType).filter((k, i, all) => all.indexOf(k) === i);

  return (
    <>
      <PageHeader
        eyebrow="Team calendar"
        title="Who's away"
        sub="Approved and pending time off of your team, with weekends and holidays. Approved leave is also added to each person's Outlook and Teams calendar."
        actions={
          <DefaultButton
            iconProps={{ iconName: 'CalendarReply' }}
            text="Export (.ics)"
            onClick={() => download(`team-${month.substring(0, 7)}.ics`, toIcs(config, shown.filter((r) => r.status === 'Approved'), (k) => typeOf(k).name), 'text/calendar')}
          />
        }
      />
      {offToday.length > 0 && (
        <div className="to-off-today">
          <strong>Away today</strong>
          {offToday.map((r) => (
            <span key={r.id} className="to-off-chip">
              <Avatar email={r.employeeEmail} name={r.employeeName} size="sm" />
              {r.employeeName}
              {r.unit !== 'day' && <span className="to-muted"> · {r.unit === 'half' ? (r.halfDayPart === 'pm' ? 'afternoon' : 'morning') : `${r.hours} h`}</span>}
            </span>
          ))}
        </div>
      )}
      <Card className="to-tc-card">
        <div className="to-tc-toolbar">
          <IconButton iconProps={{ iconName: 'ChevronLeft' }} ariaLabel="Previous month" onClick={() => setMonth(addMonths(month, -1))} />
          <h2 className="to-tc-title">
            {MONTH_NAMES[Number(month.substring(5, 7)) - 1]} {month.substring(0, 4)}
          </h2>
          <IconButton iconProps={{ iconName: 'ChevronRight' }} ariaLabel="Next month" onClick={() => setMonth(addMonths(month, 1))} />
          <DefaultButton text="Today" onClick={() => setMonth(monthStart(today()))} />
          <span className="to-spacer" />
          <Dropdown
            ariaLabel="Whose time off"
            selectedKey={scope}
            options={[
              { key: 'team', text: 'My team' },
              ...(isManager || isHr ? [{ key: 'reports', text: 'My reports' }] : []),
              { key: 'department', text: 'Department' },
              ...(isHr || config.teamCalendar.scope === 'everyone' ? [{ key: 'everyone', text: 'Everyone' }] : [])
            ]}
            onChange={(_, o) => o && setScope(o.key as Scope)}
            styles={{ root: { width: 150 } }}
          />
          {scope === 'department' && (
            <Dropdown ariaLabel="Department" selectedKey={dept} options={departments.map((d) => ({ key: d, text: d }))} onChange={(_, o) => o && setDept(String(o.key))} styles={{ root: { width: 170 } }} disabled={!isHr && !!myEmployee} />
          )}
          <Dropdown
            ariaLabel="Leave type"
            selectedKey={typeFilter}
            options={[{ key: '', text: 'All leave types' }].concat(config.leaveTypes.filter((x) => x.active).map((x) => ({ key: x.key, text: x.name })))}
            onChange={(_, o) => o && setTypeFilter(String(o.key))}
            styles={{ root: { width: 170 } }}
          />
        </div>
        {!people.length ? (
          <Empty icon="People" title="Nobody to show.">
            Ask HR to add your manager in the employee list.
          </Empty>
        ) : (
          <div className="to-tc-scroll">
            <table className="to-tc" style={{ '--to-days': days.length } as unknown as React.CSSProperties}>
              <thead>
                <tr>
                  <th className="to-tc-person">{people.length} people</th>
                  {days.map((d) => {
                    const h = holidayOn(cal, d);
                    return (
                      <th key={d} className={(isWorkingDay(cal, d) ? '' : 'is-off') + (d === t ? ' is-today' : '') + (h ? ' is-holiday' : '')} title={h}>
                        <span>{WEEKDAY_SHORT[weekday(d)].substring(0, 2)}</span>
                        <strong>{Number(d.substring(8))}</strong>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {people.map((p) => {
                  const pc = calendarFor(config, p);
                  const list = byPerson[lc(p.email)] || [];
                  const away = list.reduce((s, r) => s + (r.status === 'Approved' ? r.days : 0), 0);
                  return (
                    <tr key={p.email} className={lc(p.email) === lc(me.email) ? 'is-me' : ''}>
                      <th className="to-tc-person">
                        <span className="to-person">
                          <Avatar email={p.email} name={p.name} size="sm" />
                          <span style={{ minWidth: 0 }}>
                            <span className="to-person-name">{p.name}</span>
                            <span className="to-person-sub">{away ? `${Math.round(away * 10) / 10} d away` : p.jobTitle || ''}</span>
                          </span>
                        </span>
                      </th>
                      {days.map((d) => {
                        const r = list.filter((x) => x.start <= d && x.end >= d)[0];
                        const off = !isWorkingDay(pc, d);
                        const h = holidayOn(pc, d);
                        if (!r || (off && !(config.leaveTypes.filter((x) => x.key === r.leaveType)[0] || { countCalendarDays: false }).countCalendarDays)) {
                          return <td key={d} className={(off ? 'is-off' : '') + (d === t ? ' is-today' : '') + (h ? ' is-holiday' : '')} title={h}>{h ? <Icon iconName="Flag" className="to-tc-flag" /> : null}</td>;
                        }
                        const ty = typeOf(r.leaveType);
                        const visible = canSee(r);
                        const label = `${r.employeeName}: ${visible ? ty.name : 'Away'}${r.status !== 'Approved' ? ' (pending)' : ''}${r.unit === 'half' ? ` · ${r.halfDayPart === 'pm' ? 'afternoon' : 'morning'}` : r.unit === 'hours' ? ` · ${r.hours} h` : ''}`;
                        return (
                          <td key={d} className={(d === t ? 'is-today ' : '') + 'has-leave'}>
                            <TooltipHost content={label}>
                              <button
                                type="button"
                                className={'to-tc-leave' + (r.status !== 'Approved' ? ' is-pending' : '') + (r.unit !== 'day' ? ' is-partial' : '') + (r.unit === 'half' && r.halfDayPart === 'pm' ? ' is-pm' : '')}
                                style={{ '--to-type': visible ? ty.color : '#8a959c' } as unknown as React.CSSProperties}
                                aria-label={label}
                                onClick={() => canOpen(r) && setOpen(r)}
                              >
                                {d === r.start || d === month ? <Icon iconName={visible ? ty.icon : 'Airplane'} /> : null}
                              </button>
                            </TooltipHost>
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
        <div className="to-tc-legend">
          {usedTypes.map((k) => (
            <span key={k}>
              <i style={{ background: typeOf(k).color }} /> {typeOf(k).name}
            </span>
          ))}
          <span>
            <i className="is-pending" /> Pending
          </span>
          <span>
            <i className="is-off" /> Weekend / non-working
          </span>
          <span>
            <Icon iconName="Flag" /> Holiday
          </span>
        </div>
      </Card>
      {open && <RequestDetails request={open} onClose={() => setOpen(undefined)} />}
    </>
  );
}
