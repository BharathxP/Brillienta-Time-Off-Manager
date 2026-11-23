# Release notes

## 1.2.0

- All lists carry the Brillienta brand prefix: `BrillTM_Requests`, `BrillTM_Employees`, `BrillTM_Settings`, ...
- Permissions go through branded SharePoint groups: BrillTM Admins, BrillTM HR and BrillTM Users. Admin and HR
  membership follows Admin › Access; the site's Members and Visitors groups are taken off the app's lists.

## 1.1.1 — fixes from end-to-end testing

- Open pages refresh their data when the tab regains focus and every 2 minutes (approvers saw stale lists).
- Accrual leave can be booked against what accrues later in the leave year (next year's leave was refused with
  "0 available"); setting per leave type.
- A new installation no longer carries forward the previous year's full entitlement.
- Number boxes no longer reset to their defaults when they lose focus (e.g. a recurrence of 3 created 26 requests;
  admin settings could revert).
- An empty date box losing focus no longer clears related values.
- Declined future requests stay visible under "Upcoming".
- Date boxes accept typed dates (2026-12-07, 7/12/2026, 7 Dec 2026).
- New: `npm run e2e` — 30 multi-user scenarios against a strict SharePoint emulator; docs/TENANT-TEST-PLAN.md.

## 1.1.0

- Change a pending request; the approval restarts and the history records the change.
- Partial first / last day on multi-day leave; sandwich rule per leave type.
- Blackout periods and minimum staffing (warn or block); team cover shown to approvers and in the request form.
- My team page for managers.
- Opening-balance CSV import, bulk adjustments, leaving dates (prorated entitlement), remove sample data.
- Excel exports, leave register for any date range, Bradford factor and leave liability reports.
- Outlook automatic replies, Microsoft 365 group (Teams channel) calendar, downloadable Power Automate flow package.

## 1.0.0

First release of Brillienta Time Off Manager: 30 leave types; fixed, accrual, grade and tenure entitlements with
proration and carry-forward; comp off with expiry; full, half-day and hourly leave; work calendars and holidays;
templates and recurring (weekly, bi-weekly, monthly) leave; multilevel approvals with delegation, reminders,
escalation, auto-approval and on-behalf decisions; cancellation with approval; team calendar with Outlook sync;
reports, CSV export, Power Automate, webhook, REST and Power BI; custom branding; licensing with trial.
