# Entra-StaleDevices-Cleanup

## Why this tool?

Stale devices in Entra ID are a security risk — they clutter your directory, inflate compliance reports, and can be exploited if their identities are still trusted. Microsoft recommends regularly auditing and cleaning up devices that haven't signed in.

Doing this manually through the Entra admin center is slow and doesn't give you granular control over OS type, inactivity threshold, or bulk actions.

This tool automates the full workflow: find stale devices, filter by OS and days, then audit, disable, or delete — with confirmation before any changes.

**Microsoft documentation:** [Manage stale devices in Microsoft Entra ID](https://learn.microsoft.com/en-us/entra/identity/devices/manage-stale-devices)

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph module (auto-installs if missing)
- ImportExcel module (auto-installs if missing)
- Permissions: `Device.Read.All`, `Directory.Read.All` (audit mode), plus `Device.ReadWrite.All` for disable/delete

## How it works

The script runs as a single flow — audit first, then optionally act:

1. **Setup** — prompts for admin UPN, stale threshold (30/60/90/custom days), and OS filter (All, Windows, macOS, Linux, iOS, Android, or custom)
2. **Connect** — authenticates to Microsoft Graph via browser
3. **Fetch devices** — pulls all devices with pagination
4. **Filter & audit** — identifies stale devices by last sign-in date, applies OS filter, shows summary with OS breakdown and enabled/disabled counts
5. **Review & decide** — after seeing the results, choose to:
   - **Export only** — save the report, no changes (default)
   - **Disable** — disable stale devices (shows preview, requires typing 'YES')
   - **Delete** — delete stale devices (shows preview, requires typing 'YES')
6. **Export** — generates Excel report regardless of action taken

## Output

An Excel report saved to your Desktop: `StaleDevices_Report_<timestamp>.xlsx`

Sheets:
- **Summary** — KPI dashboard with total/stale counts and OS breakdown
- **Stale Devices** — full device list with name, OS, last sign-in, days inactive, trust type, management type
- **Action Log** — results of disable/delete operations (only when actions are taken)

## Usage

```powershell
pwsh ./stale-devices-cleanup.ps1
```
