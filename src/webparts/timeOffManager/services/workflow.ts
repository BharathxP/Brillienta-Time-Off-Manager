// Approval workflow: who approves a request, in which order, and what each decision does. Pure functions
// (unit tested in tests/workflow.test.cjs).

import { IApprovalLevel, IApprovalStep, IConfig, IDelegation, IEmployee, IHistoryEntry, ILeaveRequest, ILeaveType, IPerson, RequestStatus } from '../models/types';

export const lc = (s: string | undefined): string => (s || '').trim().toLowerCase();

export interface IDirectory {
  byEmail: Record<string, IEmployee>;
}

export function indexEmployees(list: IEmployee[]): IDirectory {
  const byEmail: Record<string, IEmployee> = {};
  list.forEach((e) => (byEmail[lc(e.email)] = e));
  return { byEmail };
}

function person(e: IEmployee | undefined): IPerson | undefined {
  return e ? { email: e.email, name: e.name } : undefined;
}

/** Department head: the configured one, else the most senior manager in the department (manager outside it). */
export function departmentHead(config: IConfig, dir: IDirectory, department: string | undefined): IPerson | undefined {
  if (!department) return undefined;
  const set = config.departmentHeads[department];
  if (set && set.email) return set;
  const members = Object.keys(dir.byEmail)
    .map((k) => dir.byEmail[k])
    .filter((e) => e.active && e.department === department);
  const reports = (m: IEmployee): number => members.filter((x) => lc(x.managerEmail) === lc(m.email)).length;
  const top = members.filter((m) => {
    const mgr = dir.byEmail[lc(m.managerEmail)];
    return reports(m) > 0 && (!mgr || mgr.department !== department);
  });
  top.sort((a, b) => reports(b) - reports(a));
  return person(top[0]);
}

/** Active delegation of an approver on a day. */
export function delegateFor(delegations: IDelegation[], approverEmail: string, on: string): IDelegation | undefined {
  return delegations.filter((d) => d.active && lc(d.delegator.email) === lc(approverEmail) && d.from <= on && d.to >= on)[0];
}

/** People who approve one level for this employee (before delegation). */
export function levelApprovers(config: IConfig, dir: IDirectory, level: IApprovalLevel, employee: IEmployee | undefined): IPerson[] {
  const mgr = employee ? dir.byEmail[lc(employee.managerEmail)] : undefined;
  switch (level.kind) {
    case 'manager':
      return mgr ? [person(mgr) as IPerson] : [];
    case 'skipManager': {
      const skip = mgr ? dir.byEmail[lc(mgr.managerEmail)] : undefined;
      return skip ? [person(skip) as IPerson] : mgr ? [person(mgr) as IPerson] : [];
    }
    case 'departmentHead': {
      const h = departmentHead(config, dir, employee && employee.department);
      return h ? [h] : [];
    }
    case 'hr':
      return config.hr.slice();
    case 'person':
      return level.person && level.person.email ? [level.person] : [];
    default:
      return [];
  }
}

function withDelegates(people: IPerson[], delegations: IDelegation[], on: string): IStepApprover[] {
  const out: IStepApprover[] = [];
  people.forEach((p) => {
    if (!out.some((x) => lc(x.email) === lc(p.email))) out.push({ email: p.email, name: p.name });
    const d = delegateFor(delegations, p.email, on);
    if (d && !out.some((x) => lc(x.email) === lc(d.delegate.email))) out.push({ email: d.delegate.email, name: d.delegate.name, delegatedFrom: p.name || p.email });
  });
  return out;
}

type IStepApprover = IPerson & { delegatedFrom?: string };

/**
 * The steps of a request: each workflow level that applies to its length, with its approvers (and their delegates).
 * The requester never approves their own request (that step goes to the next level, or HR); a level that resolves to
 * the same approvers as the step before is merged; a level nobody can approve goes to HR.
 */
export function buildSteps(config: IConfig, dir: IDirectory, delegations: IDelegation[], type: ILeaveType, employee: IEmployee | undefined, days: number, on: string): IApprovalStep[] {
  const wf = config.workflows.filter((w) => w.key === type.workflowKey)[0] || config.workflows[0];
  const levels = wf ? wf.levels.filter((l) => !l.minDays || days >= l.minDays) : [];
  const self = lc(employee && employee.email);
  const steps: IApprovalStep[] = [];
  levels.forEach((level) => {
    let people = levelApprovers(config, dir, level, employee).filter((p) => lc(p.email) !== self);
    if (!people.length && level.kind !== 'hr') people = config.hr.filter((p) => lc(p.email) !== self);
    if (!people.length) people = config.admins.filter((p) => lc(p.email) !== self);
    if (!people.length) return;
    const approvers = withDelegates(people, delegations, on);
    const prev = steps[steps.length - 1];
    const key = (a: IPerson[]): string => a.map((x) => lc(x.email)).sort().join(';');
    if (prev && key(prev.approvers) === key(approvers)) return;
    steps.push({ label: level.label || level.kind, kind: level.kind, approvers, status: 'waiting' });
  });
  return steps;
}

export interface IStartResult {
  status: RequestStatus;
  steps: IApprovalStep[];
  currentStep: number;
  autoApproved: boolean;
  history: IHistoryEntry[];
}

/** Start the workflow: auto-approve when the type's rule allows it or nobody needs to approve, else open step 1. */
export function startWorkflow(type: ILeaveType, steps: IApprovalStep[], days: number, nowIso: string, by: string): IStartResult {
  const history: IHistoryEntry[] = [{ at: nowIso, by, action: 'Submitted' }];
  const auto = type.autoApprove.enabled && (type.autoApprove.maxDays <= 0 || days <= type.autoApprove.maxDays);
  if (auto || !steps.length) {
    const why = auto ? 'Auto-approved by the leave policy' : 'Auto-approved: no approver is set up';
    history.push({ at: nowIso, by: 'System', action: 'Approved', note: why });
    return { status: 'Approved', steps: steps.map((s) => ({ ...s, status: 'skipped', comment: why })), currentStep: steps.length, autoApproved: true, history };
  }
  const first = steps.map((s, i) => (i === 0 ? { ...s, status: 'pending' as const, startedAt: nowIso } : s));
  return { status: 'Pending', steps: first, currentStep: 0, autoApproved: false, history };
}

export type Decision = 'approve' | 'reject';

export interface IActor {
  email: string;
  name: string;
  isHr: boolean;
  isAdmin: boolean;
}

/** Whether someone is one of the approvers of the current step. */
export function isStepApprover(r: ILeaveRequest, email: string): boolean {
  const s = r.steps[r.currentStep];
  return !!s && s.status === 'pending' && s.approvers.some((a) => lc(a.email) === lc(email));
}

/** Whether someone may decide the current step: its approver, or HR / an admin deciding for them. */
export function canDecide(r: ILeaveRequest, actor: IActor, config: IConfig): { ok: boolean; onBehalf: boolean } {
  if (r.status !== 'Pending' && r.status !== 'CancelRequested') return { ok: false, onBehalf: false };
  if (lc(r.employeeEmail) === lc(actor.email) && !actor.isAdmin) return { ok: false, onBehalf: false };
  if (r.status === 'CancelRequested') {
    const approvers = r.steps.reduce<string[]>((all, s) => all.concat(s.approvers.map((a) => lc(a.email))), []);
    if (approvers.indexOf(lc(actor.email)) >= 0) return { ok: true, onBehalf: false };
    return { ok: actor.isAdmin || (actor.isHr && config.onBehalf.hr), onBehalf: true };
  }
  if (isStepApprover(r, actor.email)) return { ok: true, onBehalf: false };
  if (actor.isAdmin || (actor.isHr && config.onBehalf.hr)) return { ok: true, onBehalf: true };
  return { ok: false, onBehalf: false };
}

export interface IDecisionResult {
  status: RequestStatus;
  steps: IApprovalStep[];
  currentStep: number;
  history: IHistoryEntry[];
  /** The request is finished (approved, rejected, cancelled). */
  final: boolean;
  /** People to ask next (the next step's approvers). */
  next: IPerson[];
}

/** Apply a decision on the current step. */
export function decide(r: ILeaveRequest, decision: Decision, actor: IActor, comment: string, onBehalf: boolean, nowIso: string): IDecisionResult {
  const history = r.history.slice();
  const by = actor.name || actor.email;
  if (r.status === 'CancelRequested') {
    const note = comment || undefined;
    if (decision === 'approve') {
      history.push({ at: nowIso, by, action: 'Cancellation approved', note });
      return { status: 'Cancelled', steps: r.steps, currentStep: r.currentStep, history, final: true, next: [] };
    }
    history.push({ at: nowIso, by, action: 'Cancellation declined', note });
    return { status: 'Approved', steps: r.steps, currentStep: r.currentStep, history, final: true, next: [] };
  }
  const steps = r.steps.map((s) => ({ ...s }));
  const step = steps[r.currentStep];
  if (!step) return { status: r.status, steps, currentStep: r.currentStep, history, final: true, next: [] };
  step.status = decision === 'approve' ? 'approved' : 'rejected';
  step.actedBy = { email: actor.email, name: actor.name };
  step.actedAt = nowIso;
  step.comment = comment || undefined;
  step.onBehalf = onBehalf || undefined;
  const label = `${decision === 'approve' ? 'Approved' : 'Rejected'} (${step.label})${onBehalf ? ' on behalf of the approver' : ''}`;
  history.push({ at: nowIso, by, action: label, note: comment || undefined });
  if (decision === 'reject') {
    steps.forEach((s, i) => {
      if (i > r.currentStep && s.status === 'waiting') s.status = 'skipped';
    });
    return { status: 'Rejected', steps, currentStep: r.currentStep, history, final: true, next: [] };
  }
  // Later steps that the same person would approve again are approved with this decision.
  let next = r.currentStep + 1;
  const sameApprover = (i: number): boolean => !onBehalf && i < steps.length && steps[i].approvers.some((a) => lc(a.email) === lc(actor.email));
  while (sameApprover(next)) {
    steps[next] = { ...steps[next], status: 'approved', actedBy: step.actedBy, actedAt: nowIso, comment: 'Approved at the previous level' };
    next++;
  }
  if (next >= steps.length) {
    history.push({ at: nowIso, by: 'System', action: 'Fully approved' });
    return { status: 'Approved', steps, currentStep: steps.length, history, final: true, next: [] };
  }
  steps[next] = { ...steps[next], status: 'pending', startedAt: nowIso };
  return { status: 'Pending', steps, currentStep: next, history, final: false, next: steps[next].approvers };
}

/** Everyone who may act on the request now, as a ";a;b;" string (CurrentApprovers column, for flows and views). */
export function currentApproverKey(r: Pick<ILeaveRequest, 'status' | 'steps' | 'currentStep'>): string {
  if (r.status === 'CancelRequested') {
    const all = r.steps.reduce<string[]>((acc, s) => acc.concat(s.approvers.map((a) => lc(a.email))), []);
    return all.length ? ';' + all.filter((x, i) => all.indexOf(x) === i).join(';') + ';' : '';
  }
  const s = r.steps[r.currentStep];
  return r.status === 'Pending' && s ? ';' + s.approvers.map((a) => lc(a.email)).join(';') + ';' : '';
}

/** Pending requests waiting for someone (as approver or delegate). */
export function waitingFor(requests: ILeaveRequest[], email: string): ILeaveRequest[] {
  return requests.filter((r) => (r.status === 'Pending' && isStepApprover(r, email)) || (r.status === 'CancelRequested' && currentApproverKey(r).indexOf(';' + lc(email) + ';') >= 0 && lc(r.employeeEmail) !== lc(email)));
}

/** Apply new delegations to the open step of pending requests (an approver going away after a request came in). */
export function refreshDelegates(r: ILeaveRequest, delegations: IDelegation[], on: string): IApprovalStep[] | undefined {
  const s = r.steps[r.currentStep];
  if (r.status !== 'Pending' || !s) return undefined;
  const originals = s.approvers.filter((a) => !a.delegatedFrom);
  const updated = withDelegates(originals, delegations, on);
  const key = (a: IStepApprover[]): string => a.map((x) => lc(x.email) + (x.delegatedFrom ? '*' : '')).join(';');
  if (key(updated) === key(s.approvers)) return undefined;
  return r.steps.map((x, i) => (i === r.currentStep ? { ...x, approvers: updated } : x));
}

/** Hours a step has been waiting. */
export function hoursWaiting(step: IApprovalStep | undefined, now: number): number {
  if (!step || !step.startedAt) return 0;
  const t = Date.parse(step.startedAt);
  return isNaN(t) ? 0 : (now - t) / 3600000;
}

/** What the reminder run should do with one pending request. */
export function reminderAction(r: ILeaveRequest, config: IConfig, now: number): 'none' | 'remind' | 'escalate' | 'autoApprove' {
  if (r.status !== 'Pending') return 'none';
  const step = r.steps[r.currentStep];
  if (!step) return 'none';
  const waited = hoursWaiting(step, now);
  if (config.autoApproveAfterHours > 0 && waited >= config.autoApproveAfterHours) return 'autoApprove';
  const rc = config.reminders;
  if (!rc.enabled) return 'none';
  if (rc.escalateAfterHours > 0 && waited >= rc.escalateAfterHours && step.kind !== 'hr' && !step.approvers.some((a) => config.hr.some((h) => lc(h.email) === lc(a.email))) && config.hr.length) return 'escalate';
  if (waited < rc.afterHours) return 'none';
  if ((step.reminders || 0) >= rc.maxReminders) return 'none';
  const last = step.remindedAt ? Date.parse(step.remindedAt) : 0;
  if (last && (now - last) / 3600000 < rc.repeatHours) return 'none';
  return 'remind';
}
