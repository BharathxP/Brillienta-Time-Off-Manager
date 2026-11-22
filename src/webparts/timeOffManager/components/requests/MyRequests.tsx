import * as React from 'react';
import { DefaultButton, Dropdown, PrimaryButton } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ILeaveRequest } from '../../models/types';
import { Card, PageHeader, Segmented } from '../common/Ui';
import { RequestTable } from './RequestTable';
import { RequestDetails } from './RequestDetails';
import { BalanceCards } from './BalanceCards';
import { leaveYearLabel, today } from '../../services/dates';
import { lc } from '../../services/workflow';
import { download, toIcs } from '../../services/exporting';

type Tab = 'upcoming' | 'all' | 'compoff';

export function MyRequests(): JSX.Element {
  const { world, me, myEmployee, year, config, openRequest, route } = useApp();
  const [tab, setTab] = React.useState<Tab>('upcoming');
  const [viewYear, setViewYear] = React.useState(year);
  const [open, setOpen] = React.useState<ILeaveRequest | undefined>();
  const mine = world.requests.filter((r) => lc(r.employeeEmail) === lc(me.email));
  React.useEffect(() => {
    const id = Number(route.query.id);
    if (id) {
      const r = world.requests.filter((x) => x.id === id)[0];
      if (r) setOpen(r);
    }
  }, [route.query.id, world.requests]);
  const rows =
    tab === 'upcoming'
      ? mine.filter((r) => r.kind === 'leave' && r.end >= today() && r.status !== 'Cancelled').sort((a, b) => (a.start < b.start ? -1 : 1))
      : tab === 'compoff'
        ? mine.filter((r) => r.kind === 'compoff' || r.leaveType === config.compOff.leaveType)
        : mine.filter((r) => r.leaveYear === viewYear);
  const typeName = (k: string): string => (config.leaveTypes.filter((t) => t.key === k)[0] || { name: k }).name;
  return (
    <>
      <PageHeader
        eyebrow="My leave"
        title="My time off"
        sub={myEmployee ? `${myEmployee.jobTitle || ''}${myEmployee.department ? ' · ' + myEmployee.department : ''}` : 'You are not in the employee list yet — ask HR to add you.'}
        actions={
          <>
            {config.compOff.enabled && <DefaultButton iconProps={{ iconName: 'Switch' }} text="Claim comp off" onClick={() => openRequest({ kind: 'compoff' })} />}
            <DefaultButton
              iconProps={{ iconName: 'CalendarReply' }}
              text="Export to calendar"
              onClick={() => download('my-leave.ics', toIcs(config, mine.filter((r) => r.status === 'Approved' && r.kind === 'leave'), typeName), 'text/calendar')}
            />
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="Request time off" onClick={() => openRequest()} />
          </>
        }
      />
      <div className="to-row" style={{ marginBottom: 12 }}>
        <h2 className="to-h2">Balances</h2>
        <span className="to-spacer" />
        <Dropdown
          ariaLabel="Leave year"
          selectedKey={viewYear}
          options={[year - 1, year, year + 1].map((y) => ({ key: y, text: `Leave year ${leaveYearLabel(y, config.yearStartMonth)}` }))}
          onChange={(_, o) => o && setViewYear(Number(o.key))}
          styles={{ root: { width: 200 } }}
        />
      </div>
      <BalanceCards employee={myEmployee} year={viewYear} limit={8} onRequest={(t) => openRequest({ leaveType: t })} />
      <Card style={{ marginTop: 18 }}>
        <div className="to-row" style={{ marginBottom: 12 }}>
          <Segmented
            value={tab}
            onChange={setTab}
            options={[
              { key: 'upcoming', text: 'Upcoming' },
              { key: 'all', text: `All of ${leaveYearLabel(viewYear, config.yearStartMonth)}` },
              ...(config.compOff.enabled ? [{ key: 'compoff' as Tab, text: 'Comp off' }] : [])
            ]}
          />
        </div>
        <RequestTable rows={rows} onOpen={setOpen} empty={tab === 'upcoming' ? 'No upcoming time off. Time for a break?' : 'No requests.'} />
      </Card>
      {open && <RequestDetails request={world.requests.filter((x) => x.id === open.id)[0] || open} onClose={() => setOpen(undefined)} />}
    </>
  );
}
