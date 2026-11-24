# Brillienta Time Off Manager — Privacy statement

Brillienta Time Off Manager is a SharePoint Framework app that runs entirely inside your Microsoft 365 tenant, in the
browser of the person using it (on a SharePoint page or in Microsoft Teams).

## What data is used

- **Employee data**: names, email addresses, job titles, departments, managers, grades, start and leaving dates, work
  calendars and (only when HR enters it, for gender-specific leave types) gender. It is entered by HR, imported from a
  CSV file, or read from Microsoft Entra ID through Microsoft Graph with the signed-in person's delegated permission
  (`User.Read.All`), and stored in the `BrillTM_Employees` list on your Time Off Manager site.
- **Leave data**: requests, reasons, attachments, approval steps and comments, balances adjustments, carry-forward,
  comp-off credits, delegations, notifications and an audit log, stored in the `BrillTM_` lists on the same site.
- **Calendars and mailbox settings** (optional): approved leave is added to the person's own Outlook calendar
  (`Calendars.ReadWrite`) and, when they ask for it, their Outlook automatic reply is scheduled
  (`MailboxSettings.ReadWrite`) — both from that person's own browser. When an admin switches it on, approved leave is
  posted to a Microsoft 365 group (Teams channel) calendar (`Group.Read.All`, `Group.ReadWrite.All`).
- **Email**: notifications are added to the `BrillTM_EmailQueue` list and sent by a Power Automate flow your admins
  create in your tenant, from the mailbox they choose. An optional webhook, set by your admins, receives leave events.

## Where data is stored and who can see it

All data stays in your tenant, in SharePoint lists on the site that hosts the app, protected by SharePoint permissions:
the `BrillTM Admins`, `BrillTM HR` and `BrillTM Users` groups (see the README). In Teams the app finds that site with
SharePoint search, which only returns sites the person can already open.

The publisher operates **no server** and receives **no employee data, leave data, usage data or telemetry**. The app
calls only your SharePoint sites' REST API, Microsoft Graph, Microsoft's CDN for the Fluent UI icon font, and the
webhook address your admins enter (if any).

## Retention and removal

Data is kept until your admins delete it. Removing the app and deleting the `BrillTM_` lists and `BrillTM` groups
removes all of its data.

## Contact

Questions about how your organisation uses Time Off Manager go to your Microsoft 365 administrators. Questions for
Brillienta: https://brillienta.com/contact or support@brillienta.com.
