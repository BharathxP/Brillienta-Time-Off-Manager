import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ILeaveType } from '../../models/types';
import { Avatar, Card, Empty, Kpi, StatusBadge, TypeChip } from '../common/Ui';
import { BalanceCards } from '../requests/BalanceCards';
import { addDays, calendarFor, fmtDate, fmtRange, startOfWeek, today } from '../../services/dates';
import { lc, waitingFor } from '../../services/workflow';
import { teamOf } from '../calendar/TeamCalendar';
import { lengthText } from '../requests/RequestDetails';

export function Home(): JSX.Element {
  const { me, myEmployee, world, config, year, openRequest, navigate, isHr, isManager, licensed, license, isAdmin, engine } = useApp();
  const first = (me.title || '').split(' ')[0];
  const hour = new Date().getHours();
  const greeting = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
  const t = today();
  const mine = world.requests.filter((r) => lc(r.employeeEmail) === lc(me.email));
  const upcoming = mine.filter((r) => r.kind === 'leave' && r.end >= t && (r.status === 'Approved' || r.status === 'Pending' || r.status === 'CancelRequested')).sort((a, b) => (a.start < b.start ? -1 : 1));
  const pending = mine.filter((r) => r.status === 'Pending');
  const waiting = waitingFor(world.requests, me.email);
  const team = teamOf(world.employees, me.email, 'team').map((e) => lc(e.email));
  const weekStart = startOfWeek(t, config.weekStartsOn);
  const weekEnd = addDays(weekStart, 6);
  const teamAway = world.requests.filter((r) => r.kind === 'leave' && r.status === 'Approved' && r.start <= weekEnd && r.end >= weekStart && team.indexOf(lc(r.employeeEmail)) >= 0 && lc(r.employeeEmail) !== lc(me.email));
  const cal = calendarFor(config, myEmployee);
  const holidays = cal.holidays.filter((h) => h.date >= t).sort((a, b) => (a.date < b.date ? -1 : 1)).slice(0, 5);
  const annual = config.leaveTypes.filter((x) => x.key === 'annual' && x.active)[0];
  const annualBalance = annual && myEmployee ? engine.balancesFor(world, myEmployee, year).filter((b) => b.leaveType === 'annual')[0] : undefined;
  const typeOf = (k: string): ILeaveType | undefined => config.leaveTypes.filter((x) => x.key === k)[0];
  const over = world.employees.filter((e) => e.active).length - licensed.size;
  const words = (config.headline || 'Plan time off, together.').split(' ');
  const last = words.pop();

  return (
    <>
      <section className="to-hero to-hero-to">
        <p className="to-eyebrow">
          <span className="to-star">✳</span> {greeting}, {first}
        </p>
        <h1>
          {words.join(' ')} <span className="to-underline">{last}</span>
        </h1>
        <p className="to-hero-lead">Request leave in seconds, see who is away, and keep every balance right — inside SharePoint and Microsoft Teams.</p>
        <div className="to-row" style={{ marginTop: 18 }}>
          <PrimaryButton iconProps={{ iconName: 'Add' }} text="Request time off" onClick={() => openRequest()} />
          {config.compOff.enabled && <DefaultButton iconProps={{ iconName: 'Switch' }} text="Claim comp off" onClick={() => openRequest({ kind: 'compoff' })} />}
          <DefaultButton iconProps={{ iconName: 'CalendarWeek' }} text="Team calendar" onClick={() => navigate('calendar')} />
        </div>
        {config.templates.length > 0 && (
          <div className="to-hero-templates">
            <span className="to-small to-muted">Quick requests:</span>
            {config.templates.slice(0, 5).map((tp) => (
              <button key={tp.key} type="button" className="to-chip is-button" onClick={() => openRequest({ template: tp })}>
                <Icon iconName="LightningBolt" /> {tp.name}
              </button>
            ))}
          </div>
        )}
        <span className="to-hero-star">✳</span>
      </section>

      {!myEmployee && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 16 } }}>
          You are not in the employee list yet, so you cannot request time off. {isHr ? 'Add people under Employees & balances (or import them from Microsoft 365).' : 'Ask HR to add you.'}
        </MessageBar>
      )}
      {isAdmin && over > 0 && (
        <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { marginBottom: 16 } }} actions={<DefaultButton text="Licence" onClick={() => navigate('admin/licence')} />}>
          {over} active employee{over === 1 ? ' is' : 's are'} above the {license.users} users in the licence and cannot request time off.
        </MessageBar>
      )}

      <div className="to-kpis">
        {annualBalance && <Kpi label="Annual leave left" value={annualBalance.available} hint={`of ${annualBalance.annual} this year`} icon="Airplane" tint={1} onClick={() => navigate('requests')} />}
        <Kpi label="Upcoming time off" value={upcoming.length} icon="Calendar" tint={2} onClick={() => navigate('requests')} />
        <Kpi label="My pending requests" value={pending.length} icon="Clock" tint={3} onClick={() => navigate('requests')} />
        {(isManager || isHr) && <Kpi label="Waiting for my approval" value={waiting.length} icon="CheckList" tint={4} tone={waiting.length ? 'warning' : undefined} onClick={() => navigate('approvals')} />}
        <Kpi label="Team away this week" value={teamAway.length} icon="People" tint={5} onClick={() => navigate('calendar')} />
      </div>

      <div className="to-section-head">
        <h2>Your balances</h2>
        <button className="to-link" onClick={() => navigate('requests')}>
          All my leave
        </button>
      </div>
      <BalanceCards employee={myEmployee} year={year} compact limit={5} onRequest={(k) => openRequest({ leaveType: k })} />

      <div className="to-grid to-grid-3" style={{ marginTop: 18 }}>
        <Card title="Coming up" icon="Calendar">
          {upcoming.length ? (
            upcoming.slice(0, 6).map((r) => (
              <button key={r.id} type="button" className="to-list-row" onClick={() => navigate(`requests?id=${r.id}`)}>
                <TypeChip type={typeOf(r.leaveType)} />
                <span className="to-list-main">
                  {fmtRange(r.start, r.end)}
                  <span className="to-small to-muted"> · {lengthText(r)}</span>
                </span>
                <StatusBadge status={r.status} />
              </button>
            ))
          ) : (
            <Empty icon="Airplane" title="No time off planned." />
          )}
        </Card>
        {(isManager || isHr) && waiting.length > 0 ? (
          <Card title="Waiting for you" icon="CheckList" actions={<button className="to-link" onClick={() => navigate('approvals')}>Open approvals</button>}>
            {waiting.slice(0, 6).map((r) => (
              <button key={r.id} type="button" className="to-list-row" onClick={() => navigate(`approvals?id=${r.id}`)}>
                <Avatar email={r.employeeEmail} name={r.employeeName} size="sm" />
                <span className="to-list-main">
                  <strong>{r.employeeName}</strong>
                  <span className="to-small to-muted">
                    {' '}
                    · {r.kind === 'compoff' ? 'Comp off claim' : (typeOf(r.leaveType) || { name: r.leaveType }).name}, {fmtRange(r.start, r.end)}
                  </span>
                </span>
                <StatusBadge status={r.status} />
              </button>
            ))}
          </Card>
        ) : (
          <Card title="Team away this week" icon="People" actions={<button className="to-link" onClick={() => navigate('calendar')}>Team calendar</button>}>
            {teamAway.length ? (
              teamAway.slice(0, 6).map((r) => (
                <div key={r.id} className="to-list-row is-static">
                  <Avatar email={r.employeeEmail} name={r.employeeName} size="sm" />
                  <span className="to-list-main">
                    <strong>{r.employeeName}</strong>
                    <span className="to-small to-muted"> · {fmtRange(r.start, r.end)}</span>
                  </span>
                </div>
              ))
            ) : (
              <Empty icon="People" title="Everyone is in this week." />
            )}
          </Card>
        )}
        <Card title={`Holidays · ${cal.name}`} icon="Flag">
          {holidays.length ? (
            holidays.map((h) => (
              <div key={h.date} className="to-list-row is-static">
                <span className="to-datebadge">
                  <span className="to-datebadge-m">{fmtDate(h.date).split(' ')[1]}</span>
                  <span className="to-datebadge-d">{Number(h.date.substring(8))}</span>
                </span>
                <span className="to-list-main">
                  <strong>{h.name}</strong>
                  <span className="to-small to-muted"> · {fmtDate(h.date, true)}</span>
                </span>
              </div>
            ))
          ) : (
            <Empty icon="Flag" title="No more holidays this year." />
          )}
        </Card>
      </div>
    </>
  );
}
