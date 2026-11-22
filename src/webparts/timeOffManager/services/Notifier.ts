import { IConfig, ILeaveRequest, IPerson } from '../models/types';
import { LeaveService } from './LeaveService';
import { fmtDays, fmtRange } from './dates';
import { lc } from './workflow';

export type LeaveEvent = 'submitted' | 'approved' | 'rejected' | 'stepApproved' | 'cancelled' | 'cancelRequested' | 'reminder' | 'escalated' | 'compoffCredited';

const SUBJECTS: Record<LeaveEvent, string> = {
  submitted: 'Leave request waiting for your approval',
  stepApproved: 'Leave request waiting for your approval',
  approved: 'Your leave request is approved',
  rejected: 'Your leave request was declined',
  cancelled: 'Leave cancelled',
  cancelRequested: 'Cancellation waiting for your approval',
  reminder: 'Reminder: leave request waiting for your approval',
  escalated: 'Escalated: leave request waiting for approval',
  compoffCredited: 'Comp off credited'
};

function esc(s: string): string {
  return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Email body: a simple branded card that renders in Outlook. */
export function emailHtml(config: IConfig, title: string, lines: Array<[string, string]>, link: string, intro: string): string {
  const rows = lines.map(([k, v]) => `<tr><td style="padding:4px 12px 4px 0;color:#647079">${esc(k)}</td><td style="padding:4px 0;color:#17232e"><b>${esc(v)}</b></td></tr>`).join('');
  return (
    `<div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;border:1px solid #dfe5e6;border-radius:10px;overflow:hidden">` +
    `<div style="background:${esc(config.brandColor)};color:#fff;padding:14px 20px;font-weight:700">${esc(config.appName)}${config.orgName ? ' · ' + esc(config.orgName) : ''}</div>` +
    `<div style="padding:18px 20px"><h2 style="margin:0 0 6px;font-size:18px;color:#17232e">${esc(title)}</h2><p style="color:#4b5862;margin:0 0 12px">${esc(intro)}</p>` +
    `<table style="font-size:14px;border-collapse:collapse">${rows}</table>` +
    (link ? `<p style="margin:18px 0 0"><a href="${esc(link)}" style="background:#17232e;color:#fff;padding:9px 16px;border-radius:7px;text-decoration:none;font-weight:600">Open in ${esc(config.appName)}</a></p>` : '') +
    `</div></div>`
  );
}

/**
 * Tells people what happened to a request: the bell (Notifications list), an email (EmailQueue list, sent by a Power
 * Automate flow), the Events outbox (for flows and other systems) and the optional webhook (HTTP POST of the event).
 */
export class Notifier {
  private readonly _svc: LeaveService;
  private readonly _config: () => IConfig;
  private readonly _link: string;

  public constructor(svc: LeaveService, config: () => IConfig, pageUrl: string) {
    this._svc = svc;
    this._config = config;
    this._link = pageUrl;
  }

  public requestLink(r: ILeaveRequest, forApprover: boolean): string {
    return `${this._link}#/${forApprover ? 'approvals' : 'requests'}?id=${r.id}`;
  }

  public async send(event: LeaveEvent, r: ILeaveRequest, recipients: IPerson[], typeName: string, note?: string): Promise<void> {
    const config = this._config();
    const to = recipients.filter((p, i, all) => p && p.email && all.findIndex((x) => lc(x.email) === lc(p.email)) === i);
    const forApprover = event === 'submitted' || event === 'stepApproved' || event === 'reminder' || event === 'escalated' || event === 'cancelRequested';
    const link = this.requestLink(r, forApprover);
    const what = r.kind === 'compoff' ? `Comp-off claim for ${fmtRange(r.start, r.end)}` : `${typeName}, ${fmtRange(r.start, r.end)}`;
    const length = r.unit === 'hours' ? `${r.hours} h${r.startTime ? ' from ' + r.startTime : ''}` : r.unit === 'half' ? `Half day (${r.halfDayPart === 'pm' ? 'afternoon' : 'morning'})` : fmtDays(r.days);
    const title = `${SUBJECTS[event]}`;
    const body = forApprover ? `${r.employeeName}: ${what} (${length})` : `${what} (${length})${note ? ' — ' + note : ''}`;
    const jobs: Array<Promise<unknown>> = [];
    if (config.notifications.inApp) to.forEach((p) => jobs.push(this._svc.notify(p.email, title, body, forApprover ? `approvals?id=${r.id}` : `requests?id=${r.id}`)));
    if (config.notifications.email && to.length) {
      const cc = event === 'approved' && config.notifications.ccEmails ? config.notifications.ccEmails.split(/[;,\s]+/).filter((x) => !!x) : [];
      const lines: Array<[string, string]> = [
        ['Employee', r.employeeName],
        ['Leave', r.kind === 'compoff' ? 'Comp off (credit)' : typeName],
        ['Dates', fmtRange(r.start, r.end)],
        ['Length', length]
      ];
      if (r.reason) lines.push(['Reason', r.reason]);
      if (note) lines.push(['Comment', note]);
      const html = emailHtml(config, title, lines, link, forApprover ? `${r.employeeName} asks for time off.` : `Here is the latest on your request.`);
      jobs.push(this._svc.queueEmail(to.map((p) => p.email), cc, `${title}: ${r.employeeName}, ${fmtRange(r.start, r.end)}`, html, event));
    }
    await Promise.all(jobs.map((j) => j.catch(() => undefined)));
  }

  /** Outbox row and webhook for every change of a request (Power Automate, Power BI refresh, HR systems). */
  public async publish(type: string, r: ILeaveRequest, typeName: string, actor: IPerson): Promise<void> {
    const config = this._config();
    const payload = {
      event: type,
      at: new Date().toISOString(),
      actor: actor.email,
      request: {
        id: r.id,
        kind: r.kind,
        employeeEmail: r.employeeEmail,
        employeeName: r.employeeName,
        leaveType: r.leaveType,
        leaveTypeName: typeName,
        start: r.start,
        end: r.end,
        unit: r.unit,
        hours: r.hours,
        days: r.days,
        status: r.status,
        leaveYear: r.leaveYear,
        link: this.requestLink(r, false)
      }
    };
    if (config.integrations.outbox) await this._svc.addEvent(type, r.id, r.employeeEmail, payload).catch(() => undefined);
    const url = (config.integrations.webhookUrl || '').trim();
    if (/^https:\/\//i.test(url)) {
      // text/plain + no-cors: no CORS preflight, so it reaches a Power Automate "When an HTTP request is received"
      // trigger, an Azure Function or a Logic App without extra setup. The body is the JSON event.
      try {
        await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(payload) });
      } catch {
        // a webhook that is down never blocks the request
      }
    }
  }
}
