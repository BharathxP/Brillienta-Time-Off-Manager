// What happens when someone submits, approves, rejects or cancels: the rules (workflow.ts, requests.ts,
// balances.ts) applied to SharePoint rows, with notifications, the events outbox and the Outlook calendar.

import { IConfig, IDelegation, IEmployee, ILeaveRequest, ILeaveType, ILedgerEntry, IPerson } from '../models/types';
import { LeaveService } from './LeaveService';
import { Notifier } from './Notifier';
import { GraphService } from './GraphService';
import { IDraft, duration, describeRecurrence, newSeriesId, occurrences } from './requests';
import { IActor, IDirectory, buildSteps, decide, delegateFor, indexEmployees, lc, refreshDelegates, reminderAction, startWorkflow } from './workflow';
import { addDays, daysBetween, fmtDate, leaveYearOf, today } from './dates';
import { allBalances, carryForwardDays, computeBalance } from './balances';

export interface IWorld {
  config: IConfig;
  employees: IEmployee[];
  delegations: IDelegation[];
  requests: ILeaveRequest[];
  ledger: ILedgerEntry[];
}

export class Engine {
  public readonly svc: LeaveService;
  public readonly notifier: Notifier;
  public graph?: GraphService;

  public constructor(svc: LeaveService, notifier: Notifier, graph?: GraphService) {
    this.svc = svc;
    this.notifier = notifier;
    this.graph = graph;
  }

  private _type(w: IWorld, key: string): ILeaveType | undefined {
    return w.config.leaveTypes.filter((t) => t.key === key)[0];
  }

  public typeName(w: IWorld, key: string): string {
    const t = this._type(w, key);
    return t ? t.name : key;
  }

  private _dir(w: IWorld): IDirectory {
    return indexEmployees(w.employees);
  }

  private _dept(w: IWorld, email: string): string {
    const e = this._dir(w).byEmail[lc(email)];
    return (e && e.department) || '';
  }

  private async _save(w: IWorld, r: ILeaveRequest): Promise<void> {
    await this.svc.updateRequest(r, this.typeName(w, r.leaveType), this._dept(w, r.employeeEmail));
  }

  private _hrRecipients(w: IWorld): IPerson[] {
    return w.config.notifications.notifyHr ? w.config.hr : [];
  }

  /**
   * Submit a draft for an employee (themselves, or on their behalf). A recurring draft becomes one request per
   * occurrence, linked by a series id. Returns the created requests.
   */
  public async submit(w: IWorld, draft: IDraft, actor: IActor, files: File[]): Promise<ILeaveRequest[]> {
    const dir = this._dir(w);
    const e = dir.byEmail[lc(draft.employeeEmail)];
    if (!e) throw new Error('This person is not in the employee list.');
    const type = draft.kind === 'compoff' ? this._type(w, w.config.compOff.leaveType) : this._type(w, draft.leaveType);
    if (!type) throw new Error('Unknown leave type.');
    const now = new Date().toISOString();
    const span = draft.unit === 'day' ? daysBetween(draft.start, draft.end) : 0;
    const starts = draft.kind === 'leave' && draft.recurrence && draft.recurrence.freq !== 'none' ? occurrences(draft.start, draft.recurrence) : [draft.start];
    const seriesId = starts.length > 1 ? newSeriesId() : undefined;
    const onBehalf = lc(actor.email) !== lc(e.email);
    const created: ILeaveRequest[] = [];
    for (const start of starts) {
      const end = draft.unit === 'day' ? addDays(start, span) : start;
      const dur = duration(w.config, type, e, { ...draft, start, end });
      if (dur.days <= 0) continue; // an occurrence on a holiday or weekend needs no leave
      const steps = buildSteps(w.config, dir, w.delegations, type, e, dur.days, today());
      const wf = draft.kind === 'compoff' ? { ...type, autoApprove: { enabled: false, maxDays: 0 } } : type;
      const started = startWorkflow(wf, steps, dur.days, now, actor.name);
      if (onBehalf) started.history[0] = { ...started.history[0], action: `Submitted on behalf of ${e.name}` };
      const r: Omit<ILeaveRequest, 'id' | 'created'> = {
        kind: draft.kind,
        employeeEmail: e.email,
        employeeName: e.name,
        leaveType: type.key,
        start,
        end,
        unit: draft.unit,
        halfDayPart: draft.unit === 'half' ? draft.halfDayPart || 'am' : undefined,
        partial: draft.unit === 'day' && draft.partial ? draft.partial : undefined,
        autoReply: draft.kind === 'leave' && !!draft.autoReply,
        startTime: draft.unit === 'hours' ? draft.startTime : undefined,
        hours: dur.hours,
        days: dur.days,
        leaveYear: leaveYearOf(start, w.config.yearStartMonth),
        status: started.status,
        currentStep: started.currentStep,
        steps: started.steps,
        reason: draft.reason,
        submittedBy: onBehalf ? { email: actor.email, name: actor.name } : undefined,
        seriesId,
        seriesRule: seriesId ? describeRecurrence(draft.recurrence) : undefined,
        history: started.history
      };
      const saved = await this.svc.createRequest(r, type.name, e.department || '');
      for (const f of files) await this.svc.addAttachment(saved.id, f);
      saved.hasAttachment = files.length > 0;
      created.push(saved);
      w.requests = [saved].concat(w.requests);
      if (saved.status === 'Approved') await this._finalised(w, saved, actor);
      else await this.notifier.send('submitted', saved, saved.steps[saved.currentStep].approvers, type.name);
      if (onBehalf) await this.notifier.send(saved.status === 'Approved' ? 'approved' : 'submitted', saved, [{ email: e.email, name: e.name }], type.name, `Submitted for you by ${actor.name}`).catch(() => undefined);
      await this.notifier.publish('submitted', saved, type.name, actor);
    }
    if (created.length) await this.svc.audit(onBehalf ? 'Submitted on behalf' : 'Submitted', actor, e.email, `${created.length} request(s), ${type.name}, from ${draft.start}`);
    return created;
  }

  /** After final approval: comp-off credits, the employee's notice, HR copy, events, calendar. */
  private async _finalised(w: IWorld, r: ILeaveRequest, actor: IPerson): Promise<void> {
    const name = this.typeName(w, r.leaveType);
    if (r.kind === 'compoff') {
      const type = this._type(w, r.leaveType);
      const expiry = type && type.entitlement.expiryDays > 0 ? addDays(r.end, type.entitlement.expiryDays) : undefined;
      const entry = await this.svc.addLedger({
        employeeEmail: r.employeeEmail,
        leaveType: r.leaveType,
        leaveYear: r.leaveYear,
        days: r.days,
        kind: 'compoff',
        note: `Comp off for ${r.start}${r.end !== r.start ? ' – ' + r.end : ''}: ${r.reason}`,
        effective: r.end,
        expiresOn: expiry,
        uniqueKey: `compoff:${r.id}`,
        requestId: r.id,
        by: actor.email
      });
      w.ledger = w.ledger.concat(entry);
      await this.notifier.send('compoffCredited', r, [{ email: r.employeeEmail, name: r.employeeName }], name, expiry ? `Use it by ${expiry}` : undefined);
    } else {
      await this.notifier.send('approved', r, [{ email: r.employeeEmail, name: r.employeeName }].concat(this._hrRecipients(w)), name);
      await this._publishToGroup(w, r);
    }
    await this.notifier.publish('approved', r, name, actor);
  }

  /** Approved leave on the Microsoft 365 group (Teams channel) calendar, posted by whoever approves it. */
  private async _publishToGroup(w: IWorld, r: ILeaveRequest): Promise<void> {
    const g = w.config.groupCalendar;
    if (!this.graph || !g || !g.enabled || !g.groupId || r.groupEventId) return;
    try {
      const e = this._dir(w).byEmail[lc(r.employeeEmail)];
      const ev = GraphService.eventFor(w.config, r, e, this.typeName(w, r.leaveType));
      const id = await this.graph.addGroupEvent(g.groupId, { ...ev, subject: `${r.employeeName} – ${ev.subject}` });
      await this.svc.setGroupEvent(r.id, id);
    } catch {
      // Group.ReadWrite.All not approved, or the approver is not a member: the app's team calendar still shows it.
    }
  }

  private async _unpublishFromGroup(w: IWorld, r: ILeaveRequest): Promise<void> {
    const g = w.config.groupCalendar;
    if (!this.graph || !g || !g.groupId || !r.groupEventId || r.groupEventId === 'removed') return;
    try {
      await this.graph.removeGroupEvent(g.groupId, r.groupEventId);
      await this.svc.setGroupEvent(r.id, 'removed');
    } catch {
      // the event may already be gone
    }
  }

  public async decide(w: IWorld, r: ILeaveRequest, decision: 'approve' | 'reject', comment: string, actor: IActor, onBehalf: boolean): Promise<ILeaveRequest> {
    const fresh = await this.svc.getRequest(r.id);
    if (fresh.status !== r.status || fresh.currentStep !== r.currentStep) throw new Error('Someone else has just decided this request. The list has been refreshed.');
    const wasCancelRequest = fresh.status === 'CancelRequested';
    const res = decide(fresh, decision, actor, comment, onBehalf, new Date().toISOString());
    const next: ILeaveRequest = { ...fresh, status: res.status, steps: res.steps, currentStep: res.currentStep, history: res.history };
    await this._save(w, next);
    const name = this.typeName(w, next.leaveType);
    const employee = [{ email: next.employeeEmail, name: next.employeeName }];
    if (wasCancelRequest) {
      const note = next.status === 'Cancelled' ? 'Your cancellation was approved' : next.status === 'Approved' ? 'The leave stands' : 'The request is still waiting for approval';
      await this.notifier.send(next.status === 'Cancelled' ? 'cancelled' : 'cancelDeclined', next, employee, name, note);
      if (next.status === 'Cancelled') {
        await this.notifier.publish('cancelled', next, name, actor);
        await this._unpublishFromGroup(w, next);
      }
    } else if (res.status === 'Approved') await this._finalised(w, next, actor);
    else if (res.status === 'Rejected') {
      await this.notifier.send('rejected', next, employee, name, comment);
      await this.notifier.publish('rejected', next, name, actor);
    } else {
      await this.notifier.send('stepApproved', next, res.next, name);
      await this.notifier.publish('stepApproved', next, name, actor);
    }
    if (onBehalf) await this.svc.audit(`${decision === 'approve' ? 'Approved' : 'Rejected'} on behalf`, actor, next.employeeEmail, `Request ${next.id} (${name}, ${next.start})`);
    return next;
  }

  /**
   * Cancel: a pending request is withdrawn at once; an approved one that has not started is cancelled at once by HR,
   * or goes to the approvers when the policy asks for it.
   */
  public async cancel(w: IWorld, r: ILeaveRequest, reason: string, actor: IActor): Promise<ILeaveRequest> {
    const c = w.config.cancellation;
    const now = new Date().toISOString();
    const started = r.start <= today();
    if (r.status !== 'Pending' && r.status !== 'Approved') throw new Error('Only pending or approved requests can be cancelled.');
    if (started && !c.allowPast && !actor.isHr && !actor.isAdmin) throw new Error('This leave has already started. Ask HR to cancel it.');
    const privileged = actor.isHr || actor.isAdmin;
    let status: ILeaveRequest['status'] = 'Cancelled';
    if (r.status === 'Approved' && c.approvedNeedsApproval && !privileged && r.steps.some((s) => s.status === 'approved')) status = 'CancelRequested';
    if (r.status === 'Pending' && !c.pendingDirect && !privileged) status = 'CancelRequested';
    const history = r.history.concat({ at: now, by: actor.name, action: status === 'Cancelled' ? 'Cancelled' : 'Cancellation requested', note: reason || undefined });
    const next: ILeaveRequest = { ...r, status, history, cancelReason: reason };
    await this._save(w, next);
    const name = this.typeName(w, r.leaveType);
    if (status === 'CancelRequested') {
      const approvers = r.steps.reduce<IPerson[]>((all, s) => all.concat(s.approvers), []);
      await this.notifier.send('cancelRequested', next, approvers, name, reason);
    } else {
      const told = r.steps.filter((s) => s.status === 'approved' || s.status === 'pending').reduce<IPerson[]>((all, s) => all.concat(s.approvers), []);
      if (lc(actor.email) !== lc(r.employeeEmail)) told.push({ email: r.employeeEmail, name: r.employeeName });
      await this.notifier.send('cancelled', next, told, name, reason);
      await this.notifier.publish('cancelled', next, name, actor);
      await this._unpublishFromGroup(w, next);
    }
    if (r.kind === 'compoff' && r.status === 'Approved' && status === 'Cancelled') {
      const credit = w.ledger.filter((l) => l.uniqueKey === `compoff:${r.id}`)[0];
      if (credit && credit.id) await this.svc.deleteLedger(credit.id);
    }
    return next;
  }

  /**
   * Change a pending request (dates, type, length, reason). The approval starts again from the first level, so
   * nobody approves something they have not seen.
   */
  public async edit(w: IWorld, r: ILeaveRequest, draft: IDraft, actor: IActor, files: File[]): Promise<ILeaveRequest> {
    const fresh = await this.svc.getRequest(r.id);
    if (fresh.status !== 'Pending') throw new Error('Only pending requests can be changed. Cancel and submit a new request instead.');
    const dir = this._dir(w);
    const e = dir.byEmail[lc(fresh.employeeEmail)];
    const type = this._type(w, draft.kind === 'compoff' ? w.config.compOff.leaveType : draft.leaveType);
    if (!type) throw new Error('Unknown leave type.');
    const end = draft.unit === 'day' ? draft.end : draft.start;
    const dur = duration(w.config, type, e, { ...draft, end });
    const now = new Date().toISOString();
    const steps = buildSteps(w.config, dir, w.delegations, type, e, dur.days, today());
    const wf = draft.kind === 'compoff' ? { ...type, autoApprove: { enabled: false, maxDays: 0 } } : type;
    const started = startWorkflow(wf, steps, dur.days, now, actor.name);
    const changes: string[] = [];
    if (fresh.leaveType !== type.key) changes.push(`type ${this.typeName(w, fresh.leaveType)} → ${type.name}`);
    if (fresh.start !== draft.start || fresh.end !== end) changes.push(`dates ${fresh.start}${fresh.end !== fresh.start ? '–' + fresh.end : ''} → ${draft.start}${end !== draft.start ? '–' + end : ''}`);
    if (fresh.days !== dur.days) changes.push(`${fresh.days} → ${dur.days} days`);
    const next: ILeaveRequest = {
      ...fresh,
      leaveType: type.key,
      start: draft.start,
      end,
      unit: draft.unit,
      halfDayPart: draft.unit === 'half' ? draft.halfDayPart || 'am' : undefined,
      partial: draft.unit === 'day' && draft.partial ? draft.partial : undefined,
      startTime: draft.unit === 'hours' ? draft.startTime : undefined,
      autoReply: !!draft.autoReply,
      hours: dur.hours,
      days: dur.days,
      leaveYear: leaveYearOf(draft.start, w.config.yearStartMonth),
      reason: draft.reason,
      status: started.status,
      steps: started.steps,
      currentStep: started.currentStep,
      history: fresh.history.concat({ at: now, by: actor.name, action: 'Changed — approval restarted', note: changes.join('; ') || undefined }).concat(started.history.slice(1))
    };
    await this._save(w, next);
    for (const f of files) await this.svc.addAttachment(next.id, f);
    if (next.status === 'Approved') await this._finalised(w, next, actor);
    else await this.notifier.send('submitted', next, next.steps[next.currentStep].approvers, type.name, 'The request was changed');
    await this.notifier.publish('changed', next, type.name, actor);
    return next;
  }

  /** The same adjustment for many people (e.g. a company-wide extra day). */
  public async bulkAdjust(w: IWorld, people: IEmployee[], leaveType: string, year: number, days: number, note: string, actor: IPerson): Promise<number> {
    const type = this._type(w, leaveType);
    const kind: ILedgerEntry['kind'] = type && type.entitlement.mode === 'earned' ? 'compoff' : 'grant';
    const expires = kind === 'compoff' && type && type.entitlement.expiryDays > 0 ? addDays(today(), type.entitlement.expiryDays) : undefined;
    const added: ILedgerEntry[] = [];
    for (const e of people) {
      added.push(await this.svc.addLedger({ employeeEmail: e.email, leaveType, leaveYear: year, days, kind, note, effective: today(), expiresOn: expires, by: actor.email }));
    }
    w.ledger = w.ledger.concat(added);
    await this.svc.audit('Bulk adjustment', actor, `${people.length} people`, `${this.typeName(w, leaveType)} ${year}: ${days > 0 ? '+' : ''}${days} (${note})`);
    return people.length;
  }

  /**
   * Opening balances (moving from another system): for each row, the adjustment that makes the available balance
   * equal the given number on the import day. Re-importing replaces the earlier opening entry (unique key).
   */
  public async importOpeningBalances(w: IWorld, rows: Array<{ email: string; leaveType: string; available: number }>, year: number, actor: IPerson): Promise<{ done: number; skipped: string[] }> {
    const dir = this._dir(w);
    const skipped: string[] = [];
    let done = 0;
    for (const row of rows) {
      const e = dir.byEmail[lc(row.email)];
      const type = w.config.leaveTypes.filter((t) => t.key === row.leaveType || lc(t.name) === lc(row.leaveType) || lc(t.code) === lc(row.leaveType))[0];
      if (!e || !type || isNaN(row.available)) {
        skipped.push(`${row.email} / ${row.leaveType}`);
        continue;
      }
      const key = `opening:${year}:${lc(e.email)}:${type.key}`;
      const old = w.ledger.filter((l) => l.uniqueKey === key)[0];
      if (old && old.id) {
        await this.svc.deleteLedger(old.id);
        w.ledger = w.ledger.filter((l) => l !== old);
      }
      const current = computeBalance(w.config, type, e, year, w.ledger, w.requests, today());
      const delta = Math.round((row.available - current.available) * 100) / 100;
      if (delta !== 0) {
        const entry = await this.svc.addLedger({ employeeEmail: e.email, leaveType: type.key, leaveYear: year, days: delta, kind: type.entitlement.mode === 'earned' ? 'compoff' : 'adjustment', note: `Opening balance ${row.available}`, effective: today(), uniqueKey: key, by: actor.email });
        w.ledger = w.ledger.concat(entry);
      }
      done++;
    }
    await this.svc.audit('Opening balances imported', actor, String(year), `${done} balance(s), ${skipped.length} skipped`);
    return { done, skipped };
  }

  /** Remove the sample organisation added by the setup wizard (people at @example.com, their requests and ledger). */
  public async removeSampleData(w: IWorld, actor: IPerson): Promise<number> {
    const sample = (email: string): boolean => /@example\.com$/i.test(email);
    let n = 0;
    for (const r of w.requests.filter((x) => sample(x.employeeEmail))) {
      await this.svc.deleteRequest(r.id);
      n++;
    }
    for (const l of w.ledger.filter((x) => sample(x.employeeEmail) && x.id)) await this.svc.deleteLedger(l.id as number);
    for (const e of w.employees.filter((x) => sample(x.email) && x.id)) {
      await this.svc.deleteEmployee(e.id as number);
      n++;
    }
    await this.svc.audit('Sample data removed', actor, 'example.com', `${n} rows`);
    return n;
  }

  // ---------------------------------------------------------------- balances

  public async adjust(w: IWorld, employee: IEmployee, leaveType: string, year: number, days: number, note: string, actor: IPerson, kind: ILedgerEntry['kind'] = 'adjustment'): Promise<ILedgerEntry> {
    const type = this._type(w, leaveType);
    const expires = kind === 'compoff' && type && type.entitlement.expiryDays > 0 ? addDays(today(), type.entitlement.expiryDays) : undefined;
    const entry = await this.svc.addLedger({ employeeEmail: employee.email, leaveType, leaveYear: year, days, kind, note, effective: today(), expiresOn: expires, by: actor.email });
    await this.svc.audit('Balance adjusted', actor, employee.email, `${this.typeName(w, leaveType)} ${year}: ${days > 0 ? '+' : ''}${days} (${note})`);
    return entry;
  }

  /** Carry what is left of `year` into `year + 1` for everyone (once per person and type: unique key). */
  public async carryForward(w: IWorld, year: number, actor: IPerson, log: (m: string) => void): Promise<number> {
    const end = addDays(`${year + 1}-${String(w.config.yearStartMonth).padStart(2, '0')}-01`, -1);
    let n = 0;
    for (const e of w.employees.filter((x) => x.active)) {
      for (const t of w.config.leaveTypes.filter((x) => x.active && x.entitlement.carryForwardMax !== 0)) {
        const key = `carry:${year}:${lc(e.email)}:${t.key}`;
        if (w.ledger.some((l) => l.uniqueKey === key)) continue;
        const closing = computeBalance(w.config, t, e, year, w.ledger, w.requests, end);
        const days = carryForwardDays(t, closing);
        if (days <= 0) continue;
        const entry = await this.svc.addLedger({ employeeEmail: e.email, leaveType: t.key, leaveYear: year + 1, days, kind: 'carry', note: `Carried forward from ${year}`, effective: addDays(end, 1), uniqueKey: key, by: actor.email });
        w.ledger = w.ledger.concat(entry);
        n++;
        log(`${e.name}: ${days} day(s) of ${t.name}`);
      }
    }
    await this.svc.audit('Carry forward', actor, String(year), `${n} balance(s) carried into ${year + 1}`);
    return n;
  }

  public balancesFor(w: IWorld, e: IEmployee | undefined, year: number): ReturnType<typeof allBalances> {
    return allBalances(w.config, e, year, w.ledger, w.requests, today());
  }

  // ---------------------------------------------------------------- reminders, escalation, auto-approval

  /** One pass over pending requests. Returns what was done (for the admin's run log). */
  public async runReminders(w: IWorld, actor: IActor): Promise<{ reminded: number; escalated: number; autoApproved: number; redelegated: number }> {
    const out = { reminded: 0, escalated: 0, autoApproved: 0, redelegated: 0 };
    const now = Date.now();
    const iso = new Date(now).toISOString();
    for (const r of w.requests.filter((x) => x.status === 'Pending')) {
      try {
        const delegated = refreshDelegates(r, w.delegations, today());
        if (delegated) {
          const next = { ...r, steps: delegated };
          await this._save(w, next);
          const fresh = delegated[r.currentStep].approvers.filter((a) => a.delegatedFrom && !r.steps[r.currentStep].approvers.some((b) => lc(b.email) === lc(a.email)));
          await this.notifier.send('submitted', next, fresh, this.typeName(w, r.leaveType), 'Delegated to you');
          Object.assign(r, next);
          out.redelegated++;
        }
        const action = reminderAction(r, w.config, now);
        const step = r.steps[r.currentStep];
        if (action === 'autoApprove') {
          const system: IActor = { email: actor.email, name: 'System (no decision in time)', isHr: true, isAdmin: true };
          const res = decide(r, 'approve', system, `Auto-approved after ${w.config.autoApproveAfterHours} hours without a decision`, true, iso);
          let next: ILeaveRequest = { ...r, status: res.status, steps: res.steps, currentStep: res.currentStep, history: res.history };
          // Auto-approval after the timeout approves every remaining level.
          while (next.status === 'Pending') {
            const more = decide(next, 'approve', system, 'Auto-approved (timeout)', true, iso);
            next = { ...next, status: more.status, steps: more.steps, currentStep: more.currentStep, history: more.history };
          }
          await this._save(w, next);
          await this._finalised(w, next, system);
          out.autoApproved++;
        } else if (action === 'escalate') {
          const steps = r.steps.map((s, i) => (i === r.currentStep ? { ...s, approvers: s.approvers.concat(w.config.hr.filter((h) => !s.approvers.some((a) => lc(a.email) === lc(h.email)))), remindedAt: iso } : s));
          const next = { ...r, steps, history: r.history.concat({ at: iso, by: 'System', action: 'Escalated to HR', note: `No decision after ${w.config.reminders.escalateAfterHours} hours` }) };
          await this._save(w, next);
          await this.notifier.send('escalated', next, w.config.hr, this.typeName(w, r.leaveType));
          out.escalated++;
        } else if (action === 'remind' && step) {
          const steps = r.steps.map((s, i) => (i === r.currentStep ? { ...s, remindedAt: iso, reminders: (s.reminders || 0) + 1 } : s));
          const next = { ...r, steps };
          await this._save(w, next);
          await this.notifier.send('reminder', next, step.approvers, this.typeName(w, r.leaveType));
          out.reminded++;
        }
      } catch {
        // one failing request never stops the run
      }
    }
    return out;
  }

  /** Remind the current approvers of one request now (button on the request). */
  public async remindNow(w: IWorld, r: ILeaveRequest): Promise<void> {
    const step = r.steps[r.currentStep];
    if (!step || r.status !== 'Pending') return;
    const iso = new Date().toISOString();
    const next = { ...r, steps: r.steps.map((s, i) => (i === r.currentStep ? { ...s, remindedAt: iso, reminders: (s.reminders || 0) + 1 } : s)) };
    await this._save(w, next);
    await this.notifier.send('reminder', next, step.approvers, this.typeName(w, r.leaveType));
  }

  // ---------------------------------------------------------------- Outlook / Teams calendar

  /** Put the signed-in person's approved future leave on their calendar, and take cancelled leave off it. */
  public async syncMyCalendar(w: IWorld, me: IEmployee | undefined): Promise<number> {
    if (!this.graph || !me || !w.config.calendarSync.outlook) return 0;
    let n = 0;
    const mine = w.requests.filter((r) => r.kind === 'leave' && lc(r.employeeEmail) === lc(me.email) && r.end >= today());
    for (const r of mine) {
      try {
        if (r.status === 'Approved' && !r.calendarEventId) {
          const id = await this.graph.addToMyCalendar(GraphService.eventFor(w.config, r, me, this.typeName(w, r.leaveType)));
          await this.svc.setCalendarEvent(r.id, id);
          r.calendarEventId = id;
          n++;
        } else if ((r.status === 'Cancelled' || r.status === 'Rejected') && r.calendarEventId && r.calendarEventId !== 'removed') {
          await this.graph.removeFromMyCalendar(r.calendarEventId).catch(() => undefined);
          await this.svc.setCalendarEvent(r.id, 'removed');
          r.calendarEventId = 'removed';
          n++;
        }
      } catch {
        // Calendars.ReadWrite not approved yet: the team calendar and .ics export still work.
        break;
      }
    }
    // Outlook automatic reply: Outlook holds one schedule, so the next approved leave that asked for one.
    const ar = w.config.autoReply;
    const next = mine.filter((r) => r.status === 'Approved' && r.autoReply && r.unit === 'day').sort((a, b) => (a.start < b.start ? -1 : 1))[0];
    if (ar && ar.enabled && next && !next.history.some((h) => h.action === 'Automatic reply scheduled')) {
      try {
        const mgr = this._dir(w).byEmail[lc(me.managerEmail)];
        const msg = ar.message.replace(/\{start\}/g, fmtDate(next.start)).replace(/\{end\}/g, fmtDate(next.end)).replace(/\{manager\}/g, mgr ? `${mgr.name} (${mgr.email})` : 'my manager').replace(/\{name\}/g, me.name);
        await this.graph.scheduleAutoReply(next.start, next.end, msg);
        const updated = { ...next, history: next.history.concat({ at: new Date().toISOString(), by: me.name, action: 'Automatic reply scheduled' }) };
        await this._save(w, updated);
        Object.assign(next, updated);
        n++;
      } catch {
        // MailboxSettings.ReadWrite not approved yet
      }
    }
    return n;
  }

  /** Delegation in force for someone today (shown on their approvals page). */
  public activeDelegation(w: IWorld, email: string): IDelegation | undefined {
    return delegateFor(w.delegations, email, today());
  }
}
