import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ILeaveRequest } from '../../models/types';
import { PersonLine, StatusBadge, TypeChip, useBusy, useConfirm } from '../common/Ui';
import { fmtDate, fmtDays, fmtRange, relativeTime, today } from '../../services/dates';
import { canDecide, lc } from '../../services/workflow';
import { teamClash } from '../../services/requests';
import { download, toIcs } from '../../services/exporting';

export function lengthText(r: ILeaveRequest): string {
  if (r.unit === 'hours') return `${r.hours} h${r.startTime ? ' from ' + r.startTime : ''}`;
  if (r.unit === 'half') return `Half day (${r.halfDayPart === 'pm' ? 'afternoon' : 'morning'})`;
  const part = r.partial === 'start' ? ' · from the afternoon' : r.partial === 'end' ? ' · back in the afternoon' : r.partial === 'both' ? ' · half first and last day' : '';
  return fmtDays(r.days) + part;
}

/** One request: what, when, the approval route and history, and what the viewer may do with it. */
export function RequestDetails(props: { request: ILeaveRequest; onClose: () => void }): JSX.Element {
  const app = useApp();
  const { config, me, isHr, isAdmin, world, engine, reload, svc } = app;
  const r = props.request;
  const type = config.leaveTypes.filter((t) => t.key === r.leaveType)[0];
  const actor = { email: me.email, name: me.title, isHr, isAdmin };
  const decision = canDecide(r, actor, config);
  const mine = lc(r.employeeEmail) === lc(me.email);
  const canCancel = (mine || isHr || isAdmin) && (r.status === 'Pending' || r.status === 'Approved') && (r.start > today() || config.cancellation.allowPast || isHr || isAdmin);
  const [comment, setComment] = React.useState('');
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const [files, setFiles] = React.useState<Array<{ name: string; url: string }>>([]);
  React.useEffect(() => {
    if (r.hasAttachment) svc.attachments(r.id).then(setFiles).catch(() => undefined);
  }, [r.id]);
  const series = r.seriesId ? world.requests.filter((x) => x.seriesId === r.seriesId) : [];
  const openSeries = series.filter((x) => (x.status === 'Pending' || x.status === 'Approved') && x.start > today());
  const canEdit = r.status === 'Pending' && (mine || isHr || isAdmin || (!!r.submittedBy && lc(r.submittedBy.email) === lc(me.email)));
  const employee = app.dir.byEmail[lc(r.employeeEmail)];
  const clash = decision.ok && r.kind === 'leave' ? teamClash(config, world.employees, world.requests, employee, r.start, r.end, r.id) : undefined;

  const act = async (d: 'approve' | 'reject'): Promise<void> => {
    if (d === 'reject' && !comment.trim()) {
      const why = await ask('Decline this request?', 'Tell them why (they see this).', 'Decline', 'Reason');
      if (!why) return;
      setComment(why === 'ok' ? '' : why);
      await run(async () => {
        await engine.decide(world, r, d, why === 'ok' ? '' : why, actor, decision.onBehalf);
        await reload();
        props.onClose();
      }, 'Request declined.');
      return;
    }
    await run(async () => {
      await engine.decide(world, r, d, comment, actor, decision.onBehalf);
      await reload();
      props.onClose();
    }, d === 'approve' ? 'Approved.' : 'Declined.');
  };

  const cancel = async (all: boolean): Promise<void> => {
    const why = await ask(all ? `Cancel ${openSeries.length} upcoming requests?` : 'Cancel this leave?', r.status === 'Approved' && config.cancellation.approvedNeedsApproval && !isHr ? 'The approvers are asked to confirm the cancellation.' : 'The days go back to the balance.', 'Cancel leave', 'Reason (optional)');
    if (!why) return;
    await run(async () => {
      for (const x of all ? openSeries : [r]) await engine.cancel(world, x, why === 'ok' ? '' : why, actor);
      await reload();
      props.onClose();
    }, 'Done.');
  };

  return (
    <Panel isOpen type={PanelType.medium} onDismiss={props.onClose} headerText={r.kind === 'compoff' ? 'Comp-off claim' : 'Leave request'} closeButtonAriaLabel="Close">
      {dialog}
      <div className="to-stack" style={{ marginTop: 8 }}>
        <div className="to-row" style={{ justifyContent: 'space-between' }}>
          <PersonLine email={r.employeeEmail} name={r.employeeName} sub={(app.dir.byEmail[lc(r.employeeEmail)] || { department: '' }).department} size="lg" />
          <StatusBadge status={r.status} />
        </div>
        <dl className="to-kv">
          <dt>Type</dt>
          <dd>{r.kind === 'compoff' ? <TypeChip type={type} label="Comp off credit" /> : <TypeChip type={type} />}</dd>
          <dt>When</dt>
          <dd>{fmtRange(r.start, r.end)}</dd>
          <dt>Length</dt>
          <dd>{lengthText(r)}{r.unit !== 'day' ? ` · ${fmtDays(r.days)}` : ''}</dd>
          {r.reason && (
            <>
              <dt>Reason</dt>
              <dd>{r.reason}</dd>
            </>
          )}
          {r.submittedBy && (
            <>
              <dt>Submitted by</dt>
              <dd>{r.submittedBy.name} (on behalf)</dd>
            </>
          )}
          {r.seriesRule && (
            <>
              <dt>Repeats</dt>
              <dd>
                {r.seriesRule} · {series.length} in the series
              </dd>
            </>
          )}
          {r.autoReply && (
            <>
              <dt>Outlook</dt>
              <dd>Automatic reply {r.history.some((h) => h.action === 'Automatic reply scheduled') ? 'scheduled' : 'set when the leave is approved'}</dd>
            </>
          )}
          {r.cancelReason && (
            <>
              <dt>Cancellation</dt>
              <dd>{r.cancelReason}</dd>
            </>
          )}
          {files.length > 0 && (
            <>
              <dt>Files</dt>
              <dd>
                {files.map((f) => (
                  <div key={f.url}>
                    <a href={f.url} target="_blank" rel="noopener noreferrer" data-interception="off">
                      <Icon iconName="Attach" /> {f.name}
                    </a>
                  </div>
                ))}
              </dd>
            </>
          )}
        </dl>
        <div>
          <div className="to-label">Approval route</div>
          {r.steps.length ? (
            <ol className="to-timeline">
              {r.steps.map((s, i) => (
                <li key={i} className={`is-${s.status}`}>
                  <span className="to-tl-dot">
                    <Icon iconName={s.status === 'approved' ? 'CheckMark' : s.status === 'rejected' ? 'Cancel' : s.status === 'pending' ? 'Clock' : s.status === 'skipped' ? 'Forward' : 'CircleRing'} />
                  </span>
                  <div>
                    <strong>{s.label}</strong> <span className="to-small to-muted">· {s.status === 'waiting' ? 'later' : s.status}</span>
                    <div className="to-small">
                      {s.actedBy ? `${s.actedBy.name}${s.onBehalf ? ' (on behalf)' : ''}, ${relativeTime(s.actedAt)}` : s.approvers.map((a) => a.name + (a.delegatedFrom ? ` (delegate of ${a.delegatedFrom})` : '')).join(', ')}
                    </div>
                    {s.comment && <div className="to-small to-muted">“{s.comment}”</div>}
                    {s.status === 'pending' && s.reminders ? <div className="to-small to-muted">Reminded {s.reminders}×</div> : null}
                  </div>
                </li>
              ))}
            </ol>
          ) : (
            <div className="to-small to-muted">No approval needed.</div>
          )}
        </div>
        {clash && (clash.away.length > 0 || clash.limit > 0) && (
          <div className={'to-clash' + (clash.limit > 0 && clash.peak.count > clash.limit ? ' is-over' : '')}>
            <div className="to-label">
              Team cover · {employee && employee.department}
            </div>
            <div className="to-small">
              Up to {clash.peak.count} of {clash.peak.size} away on {fmtDate(clash.peak.date)} with this request{clash.limit > 0 ? ` (limit ${clash.limit})` : ''}.
            </div>
            {clash.away.slice(0, 6).map((a, i) => (
              <div key={i} className="to-small to-muted">
                {a.name}: {fmtRange(a.start, a.end)}{a.status !== 'Approved' ? ' (pending)' : ''}
              </div>
            ))}
          </div>
        )}
        {decision.ok && (
          <div className="to-decide">
            {decision.onBehalf && (
              <MessageBar messageBarType={MessageBarType.warning}>You are deciding on behalf of {r.status === 'CancelRequested' ? 'the approvers' : (r.steps[r.currentStep] || { approvers: [] }).approvers.map((a) => a.name).join(', ')}. This is recorded.</MessageBar>
            )}
            <TextField label="Comment" multiline rows={2} value={comment} onChange={(_, v) => setComment(v || '')} />
            <div className="to-row" style={{ marginTop: 10 }}>
              <PrimaryButton iconProps={{ iconName: 'CheckMark' }} text={r.status === 'CancelRequested' ? 'Approve cancellation' : 'Approve'} disabled={busy} onClick={() => void act('approve')} />
              <DefaultButton iconProps={{ iconName: 'Cancel' }} text={r.status === 'CancelRequested' ? 'Keep the leave' : 'Decline'} disabled={busy} onClick={() => void act('reject')} />
            </div>
          </div>
        )}
        <div className="to-row">
          {canEdit && (
            <DefaultButton
              iconProps={{ iconName: 'Edit' }}
              text="Change"
              onClick={() => {
                props.onClose();
                app.openRequest({ edit: r });
              }}
            />
          )}
          {canCancel && <DefaultButton iconProps={{ iconName: 'EventDecline' }} text={r.status === 'Pending' ? 'Withdraw' : 'Cancel leave'} disabled={busy} onClick={() => void cancel(false)} />}
          {canCancel && openSeries.length > 1 && <DefaultButton text={`Cancel the series (${openSeries.length})`} disabled={busy} onClick={() => void cancel(true)} />}
          {r.status === 'Pending' && (mine || isHr) && (
            <DefaultButton iconProps={{ iconName: 'Ringer' }} text="Remind approver" disabled={busy} onClick={() => void run(() => engine.remindNow(world, r).then(reload), 'Reminder sent.')} />
          )}
          {r.status === 'Approved' && r.kind === 'leave' && (
            <DefaultButton iconProps={{ iconName: 'CalendarReply' }} text="Add to calendar (.ics)" onClick={() => download(`leave-${r.id}.ics`, toIcs(config, [r], (k) => (config.leaveTypes.filter((t) => t.key === k)[0] || { name: k }).name), 'text/calendar')} />
          )}
        </div>
        <div>
          <div className="to-label">History</div>
          {r.history
            .slice()
            .reverse()
            .map((h, i) => (
              <div key={i} className="to-small" style={{ padding: '3px 0' }}>
                <span className="to-muted">{fmtDate(h.at.substring(0, 10))} {h.at.substring(11, 16)}</span> · {h.action} — {h.by}
                {h.note ? <span className="to-muted"> · {h.note}</span> : null}
              </div>
            ))}
        </div>
      </div>
    </Panel>
  );
}
