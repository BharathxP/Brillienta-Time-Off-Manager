# Integrations: Power Automate, API and Power BI

## Email (Power Automate)

The app never sends mail itself. Every email is a row in `BrillTM_EmailQueue` (`MailTo`, `MailCc`, `Title` = subject,
`MailBody` = HTML, `MailStatus` = Queued). One flow sends them:

1. **When an item is created** (SharePoint) on `BrillTM_EmailQueue`.
2. **Send an email (V2)** — To `MailTo`, CC `MailCc`, Subject `Title`, Body `MailBody` (HTML). Use *Send an email from a
   shared mailbox* to send from HR.
3. **Update item** — `MailStatus` = `Sent`; a parallel branch configured to *run after: has failed* sets `Failed` and
   `MailError`.

## Events outbox

Every change of a request adds a row to `BrillTM_Events`:

| Column | Value |
| --- | --- |
| `EventType` | `submitted`, `stepApproved`, `approved`, `rejected`, `cancelled` |
| `RequestId` | `BrillTM_Requests` item id |
| `Payload` | JSON: `{ event, at, actor, request: { id, kind, employeeEmail, employeeName, leaveType, leaveTypeName, start, end, unit, hours, days, status, leaveYear, link } }` |

Flow ideas: post approved leave to a Teams channel; push unpaid leave to payroll; create an *Out of office* shift in
Shifts; a **Recurrence** flow that reads `BrillTM_Requests` with `Status eq 'Pending'` and emails `CurrentApprovers`
(scheduled reminders in addition to the in-app ones).

## Webhook

Admin › Power Automate & API › *Webhook address*: each event is POSTed as JSON (sent as `text/plain`, no CORS
preflight) to an HTTPS endpoint — Power Automate *When an HTTP request is received*, Logic Apps, Azure Functions or
an HR system. Parse the body as JSON.

## REST API

Everything is a SharePoint list, so the SharePoint REST API and Microsoft Graph work with Entra ID authentication:

```
GET  {site}/_api/web/lists/getbytitle('BrillTM_Requests')/items?$filter=Status eq 'Approved' and LeaveYear eq 2026
GET  {site}/_api/web/lists/getbytitle('BrillTM_Employees')/items
GET  https://graph.microsoft.com/v1.0/sites/{site-id}/lists/BrillTM_Requests/items?expand=fields
```

Writing requests from outside should go through the app (it runs the workflow); read access is safe for any system.

## Power BI

*Get data › SharePoint Online list* (2.0) with the site address, or *OData feed* with the list addresses shown in
Reports › Connect Power BI. Useful columns in `BrillTM_Requests`: `Status`, `LeaveTypeName`, `Department`, `StartOn`,
`EndOn`, `Days`, `Hours`, `LeaveYear`, `RequestKind`. For balances, schedule the *Export balances* CSV or compute from
`BrillTM_Ledger` plus the policy.
