# End-to-end tests

`npm run e2e` builds the web part and runs 30 scenarios in Chromium: the **real compiled app** against
`spEmulator.cjs`, a strict stand-in for the SharePoint REST API (lists and columns must be created first, unknown
columns and over-long text are rejected, every call carries the signed-in user). Seven people use the same site at the
same time from separate browser sessions, starting from an **empty site**: setup wizard, CSV import, configuration,
multi-level approval, auto-approval, cancellation with approval, comp off, delegation, decline, on-behalf submit and
approval, editing, bi-weekly recurrence, blackouts, opening balances, bulk adjust, Excel export, roles, team calendar,
reminders, permissions, the 1.0 → 1.1 upgrade, licence lock, sample data, concurrent decisions and partial days.

Each scenario checks the screen **and** what was written to the lists. Results and failure screenshots go to
`e2e-results/`. Microsoft Graph (Entra import, Outlook calendar, automatic replies, group calendar) needs a real
tenant — see [docs/TENANT-TEST-PLAN.md](../docs/TENANT-TEST-PLAN.md).
