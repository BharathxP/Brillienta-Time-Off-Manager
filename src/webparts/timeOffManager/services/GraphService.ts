import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { IConfig, IEmployee, ILeaveRequest } from '../models/types';
import { addDays, addTime, calendarFor } from './dates';
import { lc } from './workflow';

interface IGraphUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  employeeHireDate?: string;
  accountEnabled?: boolean;
  userType?: string;
  manager?: { mail?: string; userPrincipalName?: string };
}

/** Microsoft Graph: import people from Entra ID, and add approved leave to the person's Outlook / Teams calendar. */
export class GraphService {
  private readonly _factory: MSGraphClientFactory;
  private _client: MSGraphClientV3 | undefined;

  public constructor(factory: MSGraphClientFactory) {
    this._factory = factory;
  }

  public async client(): Promise<MSGraphClientV3> {
    if (!this._client) this._client = await this._factory.getClient('3');
    return this._client;
  }

  /** Every enabled member account with its manager (User.Read.All). */
  public async loadUsers(onProgress?: (n: number) => void): Promise<IGraphUser[]> {
    const c = await this.client();
    const out: IGraphUser[] = [];
    let url: string | undefined =
      '/users?$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,employeeHireDate,accountEnabled,userType&$expand=manager($select=mail,userPrincipalName)&$top=999&$filter=accountEnabled eq true';
    while (url) {
      const page: { value: IGraphUser[]; '@odata.nextLink'?: string } = await c.api(url).header('ConsistencyLevel', 'eventual').get();
      out.push(...page.value.filter((u) => u.userType !== 'Guest' && (u.mail || u.userPrincipalName)));
      if (onProgress) onProgress(out.length);
      url = page['@odata.nextLink'];
    }
    return out;
  }

  /** Graph users merged into the employee list: Entra values win; grade, calendar and gender stay as HR set them. */
  public static merge(existing: IEmployee[], users: IGraphUser[]): { add: IEmployee[]; update: IEmployee[] } {
    const byEmail: Record<string, IEmployee> = {};
    existing.forEach((e) => (byEmail[lc(e.email)] = e));
    const add: IEmployee[] = [];
    const update: IEmployee[] = [];
    users.forEach((u) => {
      const email = u.mail || u.userPrincipalName;
      const mgr = u.manager ? u.manager.mail || u.manager.userPrincipalName || '' : '';
      const fromGraph = {
        email,
        name: u.displayName,
        aadId: u.id,
        jobTitle: u.jobTitle || undefined,
        department: u.department || undefined,
        managerEmail: mgr || undefined,
        location: u.officeLocation || undefined,
        hireDate: u.employeeHireDate ? u.employeeHireDate.substring(0, 10) : undefined
      };
      const cur = byEmail[lc(email)];
      if (!cur) {
        add.push({ ...fromGraph, active: true });
        return;
      }
      const next: IEmployee = { ...cur, ...fromGraph, hireDate: fromGraph.hireDate || cur.hireDate, managerEmail: fromGraph.managerEmail || cur.managerEmail };
      const changed = (['name', 'aadId', 'jobTitle', 'department', 'managerEmail', 'location', 'hireDate'] as Array<keyof IEmployee>).some((k) => (next[k] || '') !== (cur[k] || ''));
      if (changed) update.push(next);
    });
    return { add, update };
  }

  /** Graph event body for a leave request (all-day, or the hours for half-day and hourly leave). */
  public static eventFor(config: IConfig, r: ILeaveRequest, employee: IEmployee | undefined, typeName: string): Record<string, unknown> {
    const cal = calendarFor(config, employee);
    const subject = `${typeName}${r.unit === 'half' ? ' (half day)' : r.unit === 'hours' ? ` (${r.hours} h)` : ''}`;
    const base = { subject, showAs: config.calendarSync.showAs === 'oof' ? 'oof' : config.calendarSync.showAs, isReminderOn: false, categories: ['Time off'], body: { contentType: 'text', content: `${config.appName}: ${r.reason || typeName}` } };
    if (r.unit === 'day') {
      return { ...base, isAllDay: true, start: { dateTime: `${r.start}T00:00:00`, timeZone: 'UTC' }, end: { dateTime: `${addDays(r.end, 1)}T00:00:00`, timeZone: 'UTC' } };
    }
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    const startTime = r.unit === 'half' ? (r.halfDayPart === 'pm' ? addTime(cal.dayStart, cal.hoursPerDay / 2) : cal.dayStart) : r.startTime || cal.dayStart;
    const hours = r.unit === 'half' ? cal.hoursPerDay / 2 : r.hours;
    return { ...base, isAllDay: false, start: { dateTime: `${r.start}T${startTime}:00`, timeZone: tz }, end: { dateTime: `${r.start}T${addTime(startTime, hours)}:00`, timeZone: tz } };
  }

  /** Add the leave to the signed-in person's calendar (Calendars.ReadWrite). Returns the event id. */
  public async addToMyCalendar(event: Record<string, unknown>): Promise<string> {
    const c = await this.client();
    const res: { id: string } = await c.api('/me/events').post(event);
    return res.id;
  }

  public async removeFromMyCalendar(eventId: string): Promise<void> {
    const c = await this.client();
    await c.api(`/me/events/${encodeURIComponent(eventId)}`).delete();
  }
}
