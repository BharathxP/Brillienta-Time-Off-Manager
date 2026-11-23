# Brillienta Time Off Manager

Leave and time-off management for **SharePoint Online and Microsoft Teams**, built with SPFx 1.23 (React 17,
Fluent UI 8). All data stays in SharePoint lists on your own site. Same design system and branding as
[Brillienta Employee Directory](../README.md): navy-and-lime workspace, Manrope / DM Sans, the Brillienta green logo tile.

## Features

| Area | What you get |
| --- | --- |
| **Leave types** | 30 ready-made types (annual, casual, sick, earned, comp off, WFH, maternity, paternity, bereavement, sabbatical, unpaid, …) — rename, recolour, add or switch off any |
| **Entitlements** | Fixed per year, **accrual** (monthly, bi-weekly, quarterly, yearly; start or end of period), **level / grade based**, **tenure based** tiers, proration for new joiners, carry-forward limit, balance cap, unlimited (tracked) |
| **Comp off / in lieu** | Employees claim extra days worked (weekends / holidays); approval credits the comp-off balance; credits lapse after N days (used first-expiring-first) |
| **Units** | Full days, **half days** (morning / afternoon) and **hours** (e.g. two-hour leave) with start time — hourly leave goes through approval like any other |
| **Non-working days** | Work calendars (working weekdays, hours per day, holidays with CSV import, optional holidays); weekends and holidays are excluded from the count unless the type counts calendar days |
| **Requests** | Live day count and balance preview, approval route preview, validation (notice, longest request, eligibility, overlap, balance, documents), attachments, **leave templates**, **recurrence** (weekly, **bi-weekly**, monthly) |
| **On behalf** | HR submit for anyone; managers for the people below them; the employee is told and the history records who submitted |
| **Workflows** | **Multilevel approval** (manager, manager's manager, department head, HR, named person), levels that apply only from N days, **auto-approval** rules per type, auto-approval after N hours without decision |
| **Approvals** | Bulk approve / decline, **approval delegation** (out of office), **approve on behalf of the manager** (HR / admins, recorded), **reminders** to approvers with repeat and escalation to HR |
| **Cancellation** | Withdraw pending requests; cancel approved leave directly (HR) or with the approvers' confirmation; cancel a whole series; days return to the balance automatically |
| **Teams calendar** | Month timeline of the team / department / everyone with weekends, holidays, pending and partial-day leave; "away today"; `.ics` export; approved leave added to each person's **Outlook / Teams calendar** (Graph) |
| **HR** | Employees (import from Microsoft Entra ID or CSV), balances with line-by-line calculation, ledger adjustments, comp-off grants, year-end carry-forward, audit log |
| **Reports** | Usage by type, department and month, approval times, balance and request **CSV exports**, **Power BI** connection details |
| **Integrations** | Emails through a **Power Automate** flow (no Mail.Send needed); Events outbox list for flows; **webhook** (JSON POST) for **API connectivity**; SharePoint REST / Microsoft Graph endpoints |
| **Branding** | App name, brand and accent colour, **logo** upload, home headline |
| **Editing** | Change a pending request (dates, type, length, reason): the approval starts again and the history shows what changed |
| **Partial & sandwich** | Multi-day leave can start in the afternoon and/or end at midday; per type **sandwich rule** (weekends / holidays between leave days count) |
| **Blackouts & staffing** | **Blackout periods** (block or warn, by department and leave type); **minimum staffing** limits per department (% or people, warn or block); approvers see the team cover on every request |
| **Managers** | **My team**: direct and indirect reports' balances, next leave, waiting approvals, Excel export, request on their behalf |
| **Balance admin** | **Opening balances** CSV import (migration), **bulk adjustments** by department / grade, **leaving date** (prorated entitlement, no leave after it), remove sample data |
| **Reports+** | **Excel (.xlsx)** exports, **leave register** for any date range (payroll), **Bradford factor**, **leave liability** |
| **Microsoft 365** | **Outlook automatic replies** during leave; approved leave on a **Microsoft 365 group / Teams channel calendar**; downloadable **Power Automate flow package** for emails |
| **Licensing** | 14-day trial; signed licence keys per tenant with a user count; pricing calculator (base users, **$1 add-on per 50 users**, **additional users at ₹100/user**), free updates and unlimited support |

## Build

Node 22.

```bash
npm install
npm run build        # production bundle + sharepoint/solution/Brillienta-Time-Off-Manager.sppkg
npm test             # build + unit tests of the leave rules (tests/*.test.cjs)
npm run preview      # local preview against an in-memory SharePoint: preview-dist/index.html
```

Preview roles: `?role=admin` (default), `?role=hr`, `?role=employee`; `?fresh=1` shows the setup wizard.

## Install

1. Upload `Brillienta-Time-Off-Manager.sppkg` to the tenant (or site) app catalog and deploy it.
2. SharePoint admin center › **Advanced › API access**: approve what you want to use (all optional) —
   `User.Read.All` (employee import), `Calendars.ReadWrite` (each person's Outlook calendar),
   `MailboxSettings.ReadWrite` (automatic replies), `Group.Read.All` + `Group.ReadWrite.All` (Teams channel calendar).
3. Add the **Brillienta Time Off Manager** web part to a page (full-width section or a single-part app page). A site
   owner opens it once: the setup wizard creates the lists and imports people (or a sample organisation).
4. Optional: add it to Teams (*Sync to Teams* in the app catalog) as a tab or personal app; the **Team Leave
   Calendar** web part shows only the calendar.
5. Admin › Notifications: **Download the flow package** and import it in Power Automate (My flows › Import ›
   Import package (legacy)), or build the 3-step flow by hand. Admin › Access: apply the recommended list permissions.

**Updates:** deploy the new package; when a release adds lists or columns, the next visit by a site owner adds them.

## Data (SharePoint lists, prefix `BrillTM_`)

| List | Holds |
| --- | --- |
| `BrillTM_Settings` | Configuration JSON (leave types, workflows, calendars, templates, roles), licence key, logo |
| `BrillTM_Employees` | Email, manager, department, grade, start date, work calendar, gender, active |
| `BrillTM_Requests` | Leave and comp-off requests: dates, unit, hours, days, status, approval steps (JSON), history (JSON), `CurrentApprovers` |
| `BrillTM_Ledger` | Adjustments, carry-forward, comp-off credits (with expiry), grants |
| `BrillTM_Delegations` | Approver → delegate, from / to |
| `BrillTM_Notifications` | The bell |
| `BrillTM_EmailQueue` | Emails for the Power Automate flow |
| `BrillTM_Events` | Outbox of request events for flows and other systems |
| `BrillTM_Audit` | Settings, balance and on-behalf changes |

Balances are never stored: they are computed from the policy, the ledger and the requests
(`services/balances.ts`), so policy changes apply at once and every number can be explained.

## Code map

```
src/webparts/timeOffManager/
  models/types.ts           domain model
  schema/                   lists (schema.ts), 30 leave types, default configuration
  services/
    dates.ts                calendar-day maths, working days, holidays, leave years
    balances.ts             entitlements (fixed, accrual, grade, tenure, proration), comp-off FIFO, carry-forward
    workflow.ts             approval steps, delegation, decisions, on-behalf, reminders/escalation
    requests.ts             duration, validation, recurrence
    Engine.ts               submit / decide / cancel / adjust / carry forward / reminders / Outlook sync
    LeaveService.ts         SharePoint list access; Provisioner.ts creates lists and permissions
    Notifier.ts             bell, email queue, events outbox, webhook
    GraphService.ts         Entra ID import, calendar events
    licensing/              licence keys, trial, pricing
  components/               Shell, Home, My leave, request form, approvals, team calendar, reports, employees, admin
tools/license/license.cjs   vendor tool: issue TO1 licence keys
```

See [docs/INTEGRATIONS.md](docs/INTEGRATIONS.md) for Power Automate, the API and Power BI.
