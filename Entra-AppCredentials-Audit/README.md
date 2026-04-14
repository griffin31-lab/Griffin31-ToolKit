# Entra-AppCredentials-Audit

## Why this tool?

Expired or soon-to-expire app registration credentials (certificates and client secrets) are a common cause of service outages and a security risk. Manually checking each app in the Entra admin portal doesn't scale.

This tool scans all your app registrations, flags expired and expiring credentials, shows you who owns each app, and optionally removes expired ones — with a full Excel audit trail.

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph module (auto-installs if missing)
- ImportExcel module (auto-installs if missing)
- Permissions: `Application.Read.All`, `Directory.Read.All` (audit mode), plus `Application.ReadWrite.All` for removal

## How it works

The script runs as a single flow — audit first, then optionally act:

1. **Setup** — prompts for admin UPN and expiry warning threshold (30/60/90/custom days)
2. **Connect** — authenticates to Microsoft Graph via browser
3. **Fetch apps** — pulls all app registrations with pagination
4. **Analyze credentials** — checks every certificate and client secret for expiry status, resolves app owners
5. **Review & decide** — after seeing the results, choose to:
   - **Export only** — save the report, no changes (default)
   - **Remove expired** — delete expired credentials (shows preview, requires typing 'YES')
6. **Export** — generates Excel report with hyperlinks to each app's credentials page in Entra portal

## Output

An Excel report saved to your Desktop: `AppCredentials_Audit_<timestamp>.xlsx`

Sheets:
- **Summary** — KPI dashboard with total/flagged counts and credential type breakdown
- **Expired** — full list of expired credentials with app name, type, dates, days expired, owner, and portal link
- **Expiring Soon** — credentials expiring within your chosen threshold, color-coded by urgency
- **Action Log** — results of removal operations (only when actions are taken)

## Usage

```powershell
pwsh ./app-credentials-audit.ps1
```
