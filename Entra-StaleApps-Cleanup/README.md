# Entra-StaleApps-Cleanup

## Why this tool?

Over time every tenant accumulates app registrations nobody uses any more — proofs-of-concept, retired integrations, abandoned dev apps. Each one carries credentials that can be stolen, API permissions that can be abused, and a surface area you don't monitor.

This tool finds every app registration that hasn't signed in for X days (or ever), shows you who owns it, and lets you disable or delete it safely — with a full Excel audit trail.

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph module (auto-installs if missing)
- ImportExcel module (auto-installs if missing)
- Entra ID P1 or P2 (the sign-in activity report requires a licensed tenant)
- Permissions: `Application.Read.All`, `Directory.Read.All`, `AuditLog.Read.All` for audit mode; `Application.ReadWrite.All` for disable/delete

## How it works

The script runs as a single flow — audit first, then optionally act:

1. **Setup** — prompts for admin UPN, staleness threshold (30 / 60 / 90 / 180 / custom days, default 90), and whether to include Microsoft first-party apps (default excluded)
2. **Connect** — authenticates to Microsoft Graph via browser
3. **Fetch** — pulls all app registrations, service principals, and the `servicePrincipalSignInActivities` aggregated report
4. **Analyze** — matches each app to its sign-in activity and classifies it:
   - **Active** — signed in within the threshold window
   - **Stale** — last sign-in older than the threshold
   - **Never used** — no sign-in activity on record
   - **No service principal** — orphaned app registration, can't sign in here
5. **Review & decide** — shows a summary; lets you choose:
   - **Export only** — audit trail, no changes (default)
   - **Disable SP** — set `accountEnabled=false` on the service principal (reversible; requires typing `YES`)
   - **Delete app registration** — permanent removal (requires typing `YES` then `DELETE`)
6. **Export** — writes an Excel report with one sheet per category and an Action Log

## A note on sign-in log retention

Raw Entra sign-in logs are retained for **30 days (P1/P2) or 7 days (Free)**. The built-in sign-in blade only shows that window.

This tool uses `/beta/reports/servicePrincipalSignInActivities` instead — Microsoft's **aggregated** last-sign-in report, which persists the "last time this app was ever used" timestamp well beyond the 30-day raw retention (their own docs show examples dating back years). That is why a 90-day threshold is meaningful even though raw logs only cover 30 days.

The endpoint is in the Graph `beta` channel. Widely used in production but subject to change — documented in the script banner.

## Output

An Excel report on your Desktop: `StaleApps_Audit_<timestamp>.xlsx`

Sheets:

- **Summary** — tenant, threshold, run mode, category counts
- **Stale** — apps past the threshold, with days-since-last-signin, all 4 sign-in flow timestamps (delegated/app-only × client/resource), credential counts, portal link
- **Never Used** — apps with no sign-in record (often the biggest category — cleanup starts here)
- **No SP** — orphaned app registrations without a service principal
- **Active** — for reference
- **Action Log** — only appears when destructive actions were performed, one row per attempted disable or delete

## Usage

```powershell
pwsh ./stale-apps-cleanup.ps1
```

## Safety defaults

- **Audit-only by default.** You have to pick destructive actions explicitly.
- **Disable before delete.** Set `accountEnabled=false` first; wait 30 days to confirm nothing breaks; then come back and delete.
- **Microsoft first-party apps are excluded by default.** These are Graph, Teams, SharePoint, etc. — apps Microsoft instantiates in your tenant. You can't delete them anyway.
- **Two confirmations for delete.** Typing `YES` lists the targets; typing `DELETE` performs the action.
- **No bulk surprises.** Every destructive call is logged row-by-row.
