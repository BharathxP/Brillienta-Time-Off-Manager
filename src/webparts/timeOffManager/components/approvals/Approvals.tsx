import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IDelegation, ILeaveRequest } from '../../models/types';
import { Card, EmployeePicker, Kpi, PageHeader, Segmented, useBusy, useConfirm } from '../common/Ui';
import { RequestTable } from '../requests/RequestTable';
import { RequestDetails } from '../requests/RequestDetails';
import { IsoDatePicker } from '../requests/RequestForm';
import { addDays, fmtDate, today } from '../../services/dates';
import { canDecide, hoursWaiting, lc, waitingFor } from '../../services/workflow';

type Tab = 'mine' | 'all' | 'decided';

/** Requests waiting for the signed-in approver (and their delegates), HR's view of every open request, bulk decisions. */
export function Approvals(): JSX.Element {
  const { world, me, isHr, isAdmin, config, engine, reload, route } = useApp();
  const [tab, setTab] = React.useState<Tab>('mine');
  const [open, setOpen] = React.useState<ILeaveRequest | undefined>();
  const [selected, setSelected] = React.useState<number[]>([]);
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const actor = { email: me.email, name: me.title, isHr, isAdmin };
  const mine = waitingFor(world.requests, me.email);
  const allOpen = world.requests.filter((r) => r.status === 'Pending' || r.status === 'CancelRequested');
  const decided = world.requests.filter((r) => r.steps.some((s) => s.actedBy && lc(s.actedBy.email) === lc(me.email)) || r.history.some((h) => h.by === me.title && /Cancellation (approved|declined)/.test(h.action)));
  const rows = tab === 'mine' ? mine : tab === 'all' ? allOpen : decided;
  React.useEffect(() => {
    const id = Number(route.query.id);
    if (id) {
      const r = world.requests.filter((x) => x.id === id)[0];
      if (r) setOpen(r);
    }
  }, [route.query.id, world.requests]);
  React.useEffect(() => setSelected([]), [tab]);
  const overdue = allOpen.filter((r) => hoursWaiting(r.steps[r.currentStep], Date.now()) > config.reminders.afterHours).length;

  const bulk = async (decision: 'approve' | 'reject'): Promise<void> => {
    const chosen = rows.filter((r) => selected.indexOf(r.id) >= 0 && canDecide(r, actor, config).ok);
    let comment = '';
    if (decision === 'reject') {
      const why = await ask(`Decline ${chosen.length} request${chosen.length === 1 ? '' : 's'}?`, 'The reason is sent to each person.', 'Decline', 'Reason');
      if (!why) return;
      comment = why === 'ok' ? '' : why;
    }
    await run(async () => {
      for (const r of chosen) await engine.decide(world, r, decision, comment, actor, canDecide(r, actor, config).onBehalf);
      await reload();
      setSelected([]);
    }, `${chosen.length} request${chosen.length === 1 ? '' : 's'} ${decision === 'approve' ? 'approved' : 'declined'}.`);
  };

  return (
    <>
      {dialog}
      <PageHeader eyebrow="Approvals" title="Waiting for a decision" sub="Approve or decline leave, comp-off claims and cancellations. HR and admins can decide on behalf of an approver." />
      <div className="to-kpis">
        <Kpi label="Waiting for you" value={mine.length} icon="CheckList" tint={1} onClick={() => setTab('mine')} />
        {isHr && <Kpi label="Open in the organisation" value={allOpen.length} icon="BulletedList" tint={2} onClick={() => setTab('all')} />}
        {isHr && <Kpi label={`Waiting over ${config.reminders.afterHours} h`} value={overdue} icon="Clock" tint={3} tone={overdue ? 'warning' : undefined} onClick={() => setTab('all')} />}
        <Kpi label="Decided by you" value={decided.length} icon="Completed" tint={4} onClick={() => setTab('decided')} />
      </div>
      <div className="to-grid to-approvals-grid">
        <Card>
          <div className="to-row" style={{ marginBottom: 12 }}>
            <Segmented
              value={tab}
              onChange={setTab}
              options={[{ key: 'mine', text: `For me (${mine.length})` }, ...(isHr ? [{ key: 'all' as Tab, text: `All open (${allOpen.length})` }] : []), { key: 'decided', text: 'Decided' }]}
            />
            <span className="to-spacer" />
            {tab !== 'decided' && selected.length > 0 && (
              <>
                <PrimaryButton iconProps={{ iconName: 'CheckMark' }} text={`Approve ${selected.length}`} disabled={busy} onClick={() => void bulk('approve')} />
                <DefaultButton iconProps={{ iconName: 'Cancel' }} text="Decline" disabled={busy} onClick={() => void bulk('reject')} />
              </>
            )}
            {isHr && tab === 'all' && (
              <DefaultButton
                iconProps={{ iconName: 'Ringer' }}
                text="Send reminders now"
                disabled={busy}
                onClick={() =>
                  void run(async () => {
                    let n = 0;
                    for (const r of allOpen.filter((x) => x.status === 'Pending')) {
                      await engine.remindNow(world, r);
                      n++;
                    }
                    await reload();
                    return n;
                  }, 'Reminders sent to the approvers.')
                }
              />
            )}
          </div>
          <RequestTable rows={rows} onOpen={setOpen} showPerson waitingColumn={tab !== 'decided'} selected={tab === 'decided' ? undefined : selected} onSelect={tab === 'decided' ? undefined : setSelected} empty={tab === 'mine' ? 'Nothing is waiting for you.' : 'No requests.'} />
        </Card>
        <DelegationCard />
      </div>
      {open && <RequestDetails request={world.requests.filter((x) => x.id === open.id)[0] || open} onClose={() => setOpen(undefined)} />}
    </>
  );
}

/** "Out of office": hand my approvals to someone for a period. HR can set it up for any approver. */
export function DelegationCard(props: { forAll?: boolean }): JSX.Element {
  const { world, me, isHr, svc, reload, dir, engine } = useApp();
  const [busy, run] = useBusy();
  const [who, setWho] = React.useState(me.email);
  const [delegate, setDelegate] = React.useState('');
  const [from, setFrom] = React.useState(today());
  const [to, setTo] = React.useState(addDays(today(), 7));
  const [note, setNote] = React.useState('');
  const mine = world.delegations.filter((d) => props.forAll || lc(d.delegator.email) === lc(me.email) || lc(d.delegate.email) === lc(me.email));
  const current = engine.activeDelegation(world, me.email);
  const save = (): Promise<boolean> =>
    run(async () => {
      const a = dir.byEmail[lc(who)];
      const b = dir.byEmail[lc(delegate)];
      if (!a || !b) throw new Error('Choose both people.');
      if (lc(a.email) === lc(b.email)) throw new Error('Choose someone else as the delegate.');
      if (to < from) throw new Error('The end date is before the start date.');
      const d: IDelegation = { delegator: { email: a.email, name: a.name }, delegate: { email: b.email, name: b.name }, from, to, note, active: true };
      await svc.saveDelegation(d);
      await svc.audit('Delegation set', { email: me.email, name: me.title }, a.email, `${b.email} from ${from} to ${to}`);
      await reload();
      // Requests already waiting move to the delegate on the next reminder run; do it now for this approver.
      const fresh = { ...world, delegations: world.delegations.concat(d) };
      await engine.runReminders(fresh, { email: me.email, name: me.title, isHr, isAdmin: false });
      await reload();
      setDelegate('');
      setNote('');
    }, 'Delegation saved. Requests waiting for them now also go to the delegate.');
  return (
    <Card title="Approval delegation" icon="Switch">
      {current && (
        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 10 } }}>
          Your approvals go to {current.delegate.name} until {fmtDate(current.to)}.
        </MessageBar>
      )}
      <div className="to-stack">
        {isHr && <EmployeePicker label="Approver" value={who} onChange={(e) => e && setWho(e.email)} />}
        <EmployeePicker label="Delegate to" value={delegate} onChange={(e) => setDelegate(e ? e.email : '')} filter={(e) => lc(e.email) !== lc(who)} />
        <div className="to-form-grid">
          <IsoDatePicker label="From" value={from} onChange={setFrom} />
          <IsoDatePicker label="To" value={to} onChange={setTo} minDate={from} />
        </div>
        <TextField label="Note" value={note} onChange={(_, v) => setNote(v || '')} placeholder="e.g. On holiday" />
        <PrimaryButton text="Delegate" disabled={busy || !delegate} onClick={() => void save()} />
      </div>
      {mine.length > 0 && (
        <div style={{ marginTop: 16 }}>
          <div className="to-label">Delegations</div>
          {mine.map((d) => (
            <div key={d.id} className="to-deleg-row">
              <div>
                <strong>{d.delegator.name}</strong> → {d.delegate.name}
                <div className="to-small to-muted">
                  {fmtDate(d.from)} – {fmtDate(d.to)}
                  {d.note ? ` · ${d.note}` : ''}
                  {d.to < today() ? ' · ended' : ''}
                </div>
              </div>
              {(isHr || lc(d.delegator.email) === lc(me.email)) && d.id && (
                <Toggle
                  checked={d.active}
                  onText="On"
                  offText="Off"
                  onChange={(_, c) => void run(async () => {
                    await svc.saveDelegation({ ...d, active: !!c });
                    await reload();
                  })}
                />
              )}
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}
