# Test plan for a real Microsoft 365 tenant

The automated suite (`npm run e2e`) covers the app against a SharePoint emulator. This plan repeats the important
paths on a real tenant, plus what only a tenant can show (Microsoft Graph, Teams, Power Automate, permissions).
Use a test site and 4 test accounts: **Admin** (site owner), **Manager**, **Employee** (reports to Manager), **HR**.

## 1. Deploy

| # | Step | Expected |
| --- | --- | --- |
| 1.1 | Upload `Brillienta-Time-Off-Manager.sppkg` to the app catalog, *Enable app* (tick "make available to all sites") | Deployed, no errors |
| 1.2 | SharePoint admin center › Advanced › API access: approve User.Read.All, Calendars.ReadWrite, MailboxSettings.ReadWrite (and Group.Read.All / Group.ReadWrite.All to test the channel calendar) | Approved |
| 1.3 | Admin: new page on the test site, add **Brillienta Time Off Manager**, publish | Setup wizard appears |
| 1.4 | Setup › Import from Microsoft 365, tick "Make me HR" | 9 `BrillTM_` lists in Site contents (hidden ones via *Site contents › Settings*), employees imported with managers |
| 1.5 | Employees & balances › Edit the Employee: grade, start date, manager | Saved; balances change with grade / tenure |

## 2. Request and approve (two browsers, or one normal + one InPrivate window)

| # | Who | Step | Expected |
| --- | --- | --- | --- |
| 2.1 | Admin | Admin › Access: add HR. Leave types › Annual › Approval: "Manager, then HR for 5+ days". Save | Saved |
| 2.2 | Employee | Request 6 working days of annual leave across a weekend | Summary: 6 days, weekend not counted, route Manager → HR |
| 2.3 | Manager | Bell shows 1; Approvals › open › Approve with a comment | Moves to HR |
| 2.4 | HR | Approve | Approved; Employee's bell + email (after step 4.1) |
| 2.5 | Employee | Reopen the app | Event "Annual leave" (Out of office) in Outlook and the Teams calendar |
| 2.6 | Employee | Casual leave, half day afternoon | Approved at once (auto-approval) |
| 2.7 | Employee | Cancel the annual leave | "Cancellation pending"; Manager approves; Outlook event removed on next open |
| 2.8 | Employee | Claim comp off for last Saturday; Manager approves | Comp off balance +1, expiry shown in Employees › Balances › Ledger |
| 2.9 | Manager | Approvals › Delegate to someone for a week; Employee submits | Request shows both; delegate can approve |
| 2.10 | HR | Request for the Employee (on behalf); Admin approves on behalf | History and Audit log show both |
| 2.11 | Employee | Edit a pending request | Approval restarts, history shows the change |
| 2.12 | Employee | WFH every two weeks, 3 times | 3 requests in a series |
| 2.13 | Employee | Annual leave with "Set an Outlook automatic reply"; approve; Employee reopens the app | Automatic reply scheduled in Outlook |

## 3. Admin and data

| # | Step | Expected |
| --- | --- | --- |
| 3.1 | Blackouts & staffing: block a week for the Employee's department | Employee cannot submit there; HR gets a warning |
| 3.2 | Employees › CSV › opening balances template, fill, import | Balances equal the file |
| 3.3 | Reports › Export balances (Excel), Leave register | Files open in Excel |
| 3.4 | Admin › Access › Apply recommended permissions | Lists stop inheriting; Employee can still request (test 2.2 again) |
| 3.5 | As Employee, open `…#/admin` | Overview, no admin navigation |
| 3.6 | Teams & Outlook: pick a team, approve a leave | Event in the team's channel calendar |
| 3.7 | Sync to Teams (app catalog), add as a tab and personal app | App works in Teams |

## 4. Power Automate and Power BI

| # | Step | Expected |
| --- | --- | --- |
| 4.1 | Admin › Notifications › Download the flow package; Power Automate › Import package (legacy); choose connections; turn on | New requests send emails; `BrillTM_EmailQueue` rows turn Sent |
| 4.2 | Power BI Desktop › SharePoint Online list › site URL › `BrillTM_Requests` | Data loads |
| 4.3 | Admin › Power Automate & API › webhook = a "When an HTTP request is received" flow URL | Each change runs the flow |
