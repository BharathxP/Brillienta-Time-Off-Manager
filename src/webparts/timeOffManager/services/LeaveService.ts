import { IApprovalStep, IDelegation, IEmployee, IHistoryEntry, ILeaveRequest, ILedgerEntry, INotification, IPerson, RequestStatus } from '../models/types';
import { ListKey, listTitle } from '../schema/schema';
import { SPClient, SPItem, odataString } from './SPClient';
import { currentApproverKey, lc } from './workflow';

function str(v: unknown): string {
  return v === undefined || v === null ? '' : String(v);
}

function num(v: unknown): number {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function json<T>(v: unknown, fallback: T): T {
  if (typeof v !== 'string' || !v) return fallback;
  try {
    return JSON.parse(v) as T;
  } catch {
    return fallback;
  }
}

/** A date-only column value for "yyyy-mm-dd" (noon UTC, so no time zone moves it to another day). */
function dateOnly(iso: string): string | null {
  return iso ? `${iso}T12:00:00Z` : null;
}

// ------------------------------------------------------------------ mappers

export function itemToEmployee(i: SPItem): IEmployee {
  return {
    id: num(i.Id),
    email: str(i.Email),
    name: str(i.Title),
    aadId: str(i.AadId) || undefined,
    jobTitle: str(i.JobTitle) || undefined,
    department: str(i.Department) || undefined,
    managerEmail: str(i.ManagerEmail) || undefined,
    grade: str(i.Grade) || undefined,
    hireDate: str(i.HireDate).substring(0, 10) || undefined,
    calendar: str(i.WorkCalendar) || undefined,
    location: str(i.Location) || undefined,
    gender: str(i.Gender) || undefined,
    leavingDate: str(i.LeavingDate) || undefined,
    active: i.IsActive !== false
  };
}

export function employeeToItem(e: IEmployee): SPItem {
  return {
    Title: e.name,
    Email: e.email,
    AadId: e.aadId || '',
    JobTitle: e.jobTitle || '',
    Department: e.department || '',
    ManagerEmail: e.managerEmail || '',
    Grade: e.grade || '',
    HireDate: e.hireDate || '',
    WorkCalendar: e.calendar || '',
    Location: e.location || '',
    Gender: e.gender || '',
    LeavingDate: e.leavingDate || '',
    IsActive: e.active
  };
}

export function itemToRequest(i: SPItem): ILeaveRequest {
  const submittedBy = str(i.SubmittedByEmail);
  return {
    id: num(i.Id),
    kind: str(i.RequestKind) === 'compoff' ? 'compoff' : 'leave',
    employeeEmail: str(i.EmployeeEmail),
    employeeName: str(i.EmployeeName),
    leaveType: str(i.LeaveType),
    start: str(i.StartDate),
    end: str(i.EndDate) || str(i.StartDate),
    unit: (str(i.Unit) as ILeaveRequest['unit']) || 'day',
    halfDayPart: (str(i.HalfDayPart) as 'am' | 'pm') || undefined,
    partial: (str(i.PartialDays) as ILeaveRequest['partial']) || undefined,
    autoReply: !!i.AutoReply,
    groupEventId: str(i.GroupEventId) || undefined,
    startTime: str(i.StartTime) || undefined,
    hours: num(i.Hours),
    days: num(i.Days),
    leaveYear: num(i.LeaveYear),
    status: (str(i.Status) as RequestStatus) || 'Pending',
    currentStep: num(i.CurrentStep),
    steps: json<IApprovalStep[]>(i.Steps, []),
    reason: str(i.Reason),
    submittedBy: submittedBy ? { email: submittedBy, name: str(i.SubmittedByName) } : undefined,
    seriesId: str(i.SeriesId) || undefined,
    seriesRule: str(i.SeriesRule) || undefined,
    history: json<IHistoryEntry[]>(i.History, []),
    calendarEventId: str(i.CalendarEventId) || undefined,
    cancelReason: str(i.CancelReason) || undefined,
    hasAttachment: !!i.Attachments,
    created: str(i.Created),
    modified: str(i.Modified) || undefined
  };
}

export function requestToItem(r: Omit<ILeaveRequest, 'id' | 'created'>, typeName: string, department: string): SPItem {
  return {
    Title: `${r.employeeName} · ${typeName} · ${r.start}${r.end && r.end !== r.start ? ' → ' + r.end : ''}`,
    RequestKind: r.kind,
    EmployeeEmail: r.employeeEmail,
    EmployeeName: r.employeeName,
    Department: department,
    LeaveType: r.leaveType,
    LeaveTypeName: typeName,
    StartDate: r.start,
    EndDate: r.end,
    StartOn: dateOnly(r.start),
    EndOn: dateOnly(r.end),
    Unit: r.unit,
    HalfDayPart: r.halfDayPart || '',
    PartialDays: r.partial || '',
    AutoReply: !!r.autoReply,
    GroupEventId: r.groupEventId || '',
    StartTime: r.startTime || '',
    Hours: r.hours,
    Days: r.days,
    LeaveYear: r.leaveYear,
    Status: r.status,
    CurrentStep: r.currentStep,
    CurrentApprovers: currentApproverKey(r),
    Steps: JSON.stringify(r.steps),
    Reason: r.reason,
    SubmittedByEmail: r.submittedBy ? r.submittedBy.email : '',
    SubmittedByName: r.submittedBy ? r.submittedBy.name : '',
    SeriesId: r.seriesId || '',
    SeriesRule: r.seriesRule || '',
    History: JSON.stringify(r.history),
    CalendarEventId: r.calendarEventId || '',
    CancelReason: r.cancelReason || ''
  };
}

export function itemToLedger(i: SPItem): ILedgerEntry {
  return {
    id: num(i.Id),
    employeeEmail: str(i.EmployeeEmail),
    leaveType: str(i.LeaveType),
    leaveYear: num(i.LeaveYear),
    days: num(i.Days),
    kind: (str(i.EntryKind) as ILedgerEntry['kind']) || 'adjustment',
    note: str(i.Note),
    effective: str(i.Effective) || str(i.Created).substring(0, 10),
    expiresOn: str(i.ExpiresOn) || undefined,
    uniqueKey: str(i.UniqueKey) || undefined,
    requestId: num(i.RequestId) || undefined,
    by: str(i.EnteredBy) || undefined
  };
}

export function itemToDelegation(i: SPItem): IDelegation {
  return {
    id: num(i.Id),
    delegator: { email: str(i.DelegatorEmail), name: str(i.DelegatorName) },
    delegate: { email: str(i.DelegateEmail), name: str(i.DelegateName) },
    from: str(i.FromDate),
    to: str(i.ToDate),
    note: str(i.Note) || undefined,
    active: i.IsActive !== false
  };
}

/**
 * Reads and writes the app's SharePoint lists. Business rules live in the pure modules (balances, workflow,
 * requests); this class only moves rows.
 */
export class LeaveService {
  public readonly sp: SPClient;
  public readonly prefix: string;

  public constructor(sp: SPClient, prefix: string) {
    this.sp = sp;
    this.prefix = prefix;
  }

  public title(key: ListKey): string {
    return listTitle(this.prefix, key);
  }

  private _items(key: ListKey): string {
    return `${this.sp.list(this.title(key))}/items`;
  }

  private async _add(key: ListKey, body: SPItem): Promise<SPItem> {
    return this.sp.post<SPItem>(this._items(key), body);
  }

  // ---------------------------------------------------------------- employees

  public async loadEmployees(): Promise<IEmployee[]> {
    const rows = await this.sp.getAll<SPItem>(`${this._items('employees')}?$top=5000&$select=Id,Title,Email,AadId,JobTitle,Department,ManagerEmail,Grade,HireDate,WorkCalendar,Location,Gender,LeavingDate,IsActive`);
    return rows.map(itemToEmployee).filter((e) => !!e.email);
  }

  public async saveEmployee(e: IEmployee): Promise<IEmployee> {
    if (e.id) {
      await this.sp.merge(`${this._items('employees')}(${e.id})`, employeeToItem(e));
      return e;
    }
    const row = await this._add('employees', employeeToItem(e));
    return { ...e, id: num(row.Id) };
  }

  public async deleteEmployee(id: number): Promise<void> {
    await this.sp.del(`${this._items('employees')}(${id})`);
  }

  // ---------------------------------------------------------------- requests

  public async loadRequests(fromYear: number): Promise<ILeaveRequest[]> {
    const rows = await this.sp.getAll<SPItem>(`${this._items('requests')}?$top=5000&$filter=LeaveYear ge ${fromYear}&$orderby=Id desc`);
    return rows.map(itemToRequest);
  }

  public async getRequest(id: number): Promise<ILeaveRequest> {
    return itemToRequest(await this.sp.get<SPItem>(`${this._items('requests')}(${id})`));
  }

  public async createRequest(r: Omit<ILeaveRequest, 'id' | 'created'>, typeName: string, department: string): Promise<ILeaveRequest> {
    const row = await this._add('requests', requestToItem(r, typeName, department));
    return { ...r, id: num(row.Id), created: str(row.Created) || new Date().toISOString() };
  }

  public async updateRequest(r: ILeaveRequest, typeName: string, department: string): Promise<void> {
    await this.sp.merge(`${this._items('requests')}(${r.id})`, requestToItem(r, typeName, department));
  }

  public async setCalendarEvent(id: number, eventId: string): Promise<void> {
    await this.sp.merge(`${this._items('requests')}(${id})`, { CalendarEventId: eventId });
  }

  public async setGroupEvent(id: number, eventId: string): Promise<void> {
    await this.sp.merge(`${this._items('requests')}(${id})`, { GroupEventId: eventId });
  }

  public async deleteRequest(id: number): Promise<void> {
    await this.sp.del(`${this._items('requests')}(${id})`);
  }

  public async addAttachment(id: number, file: File): Promise<void> {
    const buf = await file.arrayBuffer();
    const name = file.name.replace(/[^\w.\- ]/g, '_');
    await this.sp.post(`${this._items('requests')}(${id})/AttachmentFiles/add(FileName=${odataString(name)})`, buf);
  }

  public async attachments(id: number): Promise<Array<{ name: string; url: string }>> {
    const r = await this.sp.get<{ value: Array<{ FileName: string; ServerRelativeUrl: string }> }>(`${this._items('requests')}(${id})/AttachmentFiles`);
    const origin = new URL(this.sp.webUrl).origin;
    return (r.value || []).map((a) => ({ name: a.FileName, url: origin + a.ServerRelativeUrl }));
  }

  // ---------------------------------------------------------------- ledger

  public async loadLedger(fromYear: number): Promise<ILedgerEntry[]> {
    const rows = await this.sp.getAll<SPItem>(`${this._items('ledger')}?$top=5000&$filter=LeaveYear ge ${fromYear - 1} or EntryKind eq 'compoff'`);
    return rows.map(itemToLedger);
  }

  public async addLedger(l: ILedgerEntry): Promise<ILedgerEntry> {
    const row = await this._add('ledger', {
      Title: `${l.kind} · ${l.employeeEmail} · ${l.leaveType}`,
      EmployeeEmail: l.employeeEmail,
      LeaveType: l.leaveType,
      LeaveYear: l.leaveYear,
      Days: l.days,
      EntryKind: l.kind,
      Note: l.note,
      Effective: l.effective,
      ExpiresOn: l.expiresOn || '',
      UniqueKey: l.uniqueKey || '',
      RequestId: l.requestId || null,
      EnteredBy: l.by || ''
    });
    return { ...l, id: num(row.Id) };
  }

  public async deleteLedger(id: number): Promise<void> {
    await this.sp.del(`${this._items('ledger')}(${id})`);
  }

  // ---------------------------------------------------------------- delegations

  public async loadDelegations(): Promise<IDelegation[]> {
    const rows = await this.sp.getAll<SPItem>(`${this._items('delegations')}?$top=2000`);
    return rows.map(itemToDelegation);
  }

  public async saveDelegation(d: IDelegation): Promise<void> {
    const body = {
      Title: `${d.delegator.name} → ${d.delegate.name}`,
      DelegatorEmail: d.delegator.email,
      DelegatorName: d.delegator.name,
      DelegateEmail: d.delegate.email,
      DelegateName: d.delegate.name,
      FromDate: d.from,
      ToDate: d.to,
      Note: d.note || '',
      IsActive: d.active
    };
    if (d.id) await this.sp.merge(`${this._items('delegations')}(${d.id})`, body);
    else await this._add('delegations', body);
  }

  public async deleteDelegation(id: number): Promise<void> {
    await this.sp.del(`${this._items('delegations')}(${id})`);
  }

  // ---------------------------------------------------------------- notifications, email, events, audit

  public async loadNotifications(email: string): Promise<INotification[]> {
    const rows = await this.sp.get<{ value: SPItem[] }>(`${this._items('notifications')}?$top=40&$orderby=Id desc&$filter=Recipient eq ${odataString(lc(email))}`);
    return (rows.value || []).map((i) => ({
      id: num(i.Id),
      recipient: str(i.Recipient),
      title: str(i.Title),
      body: str(i.Body),
      link: str(i.Link),
      read: !!i.IsRead,
      created: str(i.Created)
    }));
  }

  public async notify(recipient: string, title: string, body: string, link: string): Promise<void> {
    await this._add('notifications', { Title: title.substring(0, 250), Recipient: lc(recipient), Body: body, Link: link, IsRead: false });
  }

  public async markRead(id: number): Promise<void> {
    await this.sp.merge(`${this._items('notifications')}(${id})`, { IsRead: true });
  }

  public async queueEmail(to: string[], cc: string[], subject: string, html: string, template: string): Promise<void> {
    await this._add('emailQueue', { Title: subject.substring(0, 250), MailTo: to.join(';'), MailCc: cc.join(';'), MailBody: html, MailStatus: 'Queued', Template: template });
  }

  public async emailQueueStats(): Promise<{ queued: number; sent: number; failed: number }> {
    const rows = await this.sp.get<{ value: SPItem[] }>(`${this._items('emailQueue')}?$top=500&$select=Id,MailStatus&$orderby=Id desc`);
    const out = { queued: 0, sent: 0, failed: 0 };
    (rows.value || []).forEach((r) => {
      const s = str(r.MailStatus);
      if (s === 'Sent') out.sent++;
      else if (s === 'Failed') out.failed++;
      else out.queued++;
    });
    return out;
  }

  public async addEvent(type: string, requestId: number, employeeEmail: string, payload: unknown): Promise<void> {
    await this._add('events', { Title: type, EventType: type, RequestId: requestId, EmployeeEmail: employeeEmail, Payload: JSON.stringify(payload), ProcessedStatus: 'New' });
  }

  public async audit(action: string, actor: IPerson, target: string, details: string): Promise<void> {
    try {
      await this._add('audit', { Title: action, Action: action, Actor: actor.email, Target: target, Details: details });
    } catch {
      // the audit log never blocks the change itself
    }
  }

  public async loadAudit(): Promise<Array<{ id: number; action: string; actor: string; target: string; details: string; created: string }>> {
    const rows = await this.sp.get<{ value: SPItem[] }>(`${this._items('audit')}?$top=200&$orderby=Id desc`);
    return (rows.value || []).map((i) => ({ id: num(i.Id), action: str(i.Action), actor: str(i.Actor), target: str(i.Target), details: str(i.Details), created: str(i.Created) }));
  }
}
