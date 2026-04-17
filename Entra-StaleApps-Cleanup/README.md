# Entra-StaleApps-Cleanup

> **Find and clean up unused app registrations** — audit, disable, or delete apps that haven't signed in for X days.

<sub>[← Back to Griffin31 ToolKit](../)</sub>

---

## What you get

- **Classification of every app** — Active / Stale / Never used / No service principal
- **All 4 sign-in flow timestamps** — delegated+app-only × client+resource
- **Credential counts** per app (total + expired)
- **Excel with portal deep-links** per row, one sheet per category
- **Two destructive actions** — disable SP (reversible) or delete app (typed `YES` then `DELETE`)
- **Action log** recorded row-by-row

## Quick start

```powershell
pwsh ./stale-apps-cleanup.ps1
```

Prompts for: admin UPN, threshold (default 90 days), and whether to include Microsoft first-party apps (default excluded).

## Why this tool?

Over time every tenant accumulates app registrations nobody uses — proofs-of-concept, retired integrations, abandoned dev apps. Each one carries credentials that can be stolen, API permissions that can be abused, and a surface area you don't monitor. This tool finds every app that hasn't signed in for X days (or ever), shows who owns it, and lets you disable or delete it safely.

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- `ImportExcel` module — auto-installs if missing
- **Entra ID P1 or P2** (sign-in activity report requires a licensed tenant)
- Permissions: `Application.Read.All`, `Directory.Read.All`, `AuditLog.Read.All` (audit); `Application.ReadWrite.All` (disable/delete)

## How it works

Single flow — audit first, then optionally act:

1. **Setup** — admin UPN, threshold, first-party include/exclude
2. **Connect** — Microsoft Graph sign-in via browser
3. **Fetch** — all app registrations, service principals, `servicePrincipalSignInActivities` aggregated report
4. **Analyze** — classify each app by last sign-in
5. **Review & decide** — export only / disable SP / delete app registration
6. **Export** — Excel with one sheet per category and an Action Log

## Sign-in log retention — why 90 days works

Raw Entra sign-in logs are retained for **30 days (P1/P2) or 7 days (Free)**. This tool uses `/beta/reports/servicePrincipalSignInActivities` — Microsoft's **aggregated** last-sign-in report, which persists the "last time this app was ever used" timestamp well beyond raw retention (MS docs show examples dating back years). That's why a 90-day threshold is meaningful even though raw logs only cover 30 days.

The endpoint is in Graph `beta` — widely used in production but subject to change.

## Output

`StaleApps_Audit_<timestamp>.xlsx` on your Desktop:

- **Summary** — tenant, threshold, mode, category counts
- **Stale** — apps past the threshold, with all 4 sign-in flow timestamps, credential counts, portal link
- **Never Used** — apps with no sign-in record (often the biggest category)
- **No SP** — orphaned app registrations without a service principal
- **Active** — for reference
- **Action Log** — appears only when destructive actions run

## Safety

- **Audit-only by default.**
- **Disable before delete.** Set `accountEnabled=false` first; wait 30 days; then come back and delete if nothing broke.
- **Microsoft first-party apps excluded by default.**
- **Two confirmations for delete.** Type `YES` to list targets; type `DELETE` to execute.

## Related tools

- [Entra-AppCredentials-Audit](../Entra-AppCredentials-Audit/) — for apps you keep, audit their credential expiry
- [Entra-StaleDevices-Cleanup](../Entra-StaleDevices-Cleanup/) — same pattern for devices
