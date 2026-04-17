# Entra-AppCredentials-Audit

> **Catch expired and expiring app registration credentials** — Excel report with owner + remove-expired action.

<sub>[← Back to Griffin31 ToolKit](../)</sub>

---

## What you get

- **KPI dashboard** — total apps, flagged credentials, breakdown by type (cert vs secret)
- **Expired credentials** sheet — full list with dates and days overdue
- **Expiring soon** sheet — color-coded by urgency (red ≤ 7 days, amber ≤ 14)
- **Owner resolved** for each app
- **Direct Entra portal links** on every row
- **Optional remove action** — delete expired credentials (preview + typed-YES confirmation)

## Quick start

```powershell
pwsh ./app-credentials-audit.ps1
```

The script will prompt for: admin UPN, expiry threshold (30/60/90/custom days).

## Why this tool?

Expired or soon-to-expire app registration credentials (certificates and client secrets) are a common cause of service outages and a security risk. Manually checking each app in the Entra admin portal doesn't scale. This tool scans every app, flags expired and expiring credentials, shows who owns each app, and optionally removes expired ones — with a full Excel audit trail.

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- `ImportExcel` module — auto-installs if missing
- Permissions: `Application.Read.All`, `Directory.Read.All` (audit mode); plus `Application.ReadWrite.All` for removal

## How it works

Single flow — audit first, then optionally act:

1. **Setup** — prompts for admin UPN and expiry threshold (30/60/90/custom)
2. **Connect** — Microsoft Graph sign-in via browser
3. **Fetch apps** — all app registrations with pagination
4. **Analyze credentials** — every cert + secret checked; owners resolved
5. **Review & decide** — export only (default) OR remove expired (typed-YES confirmation)
6. **Export** — Excel with portal hyperlinks

## Output

`AppCredentials_Audit_<timestamp>.xlsx` on your Desktop:

- **Summary** — KPI dashboard
- **Expired** — expired credentials with owner and portal link
- **Expiring Soon** — color-coded by urgency
- **Action Log** — removal results (appears only when actions taken)

## Safety

- Audit-only by default. Remove action requires typing `YES`.
- Preview shown before any deletion.
- Action log recorded row-by-row so you can reverse mistakes via credential re-issuance.

## Related tools

- [Entra-StaleApps-Cleanup](../Entra-StaleApps-Cleanup/) — catch unused apps (credentials may be live but the app is abandoned)
- [CA-Update-AffectedApps](../CA-Update-AffectedApps/) — for apps with MFA gaps ahead of the May 2026 CA change
