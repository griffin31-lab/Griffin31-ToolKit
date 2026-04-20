# Entra-StaleDevices-Cleanup

> **Audit, disable, or delete stale Entra ID devices** — filter by OS and inactivity threshold, act in bulk with confirmation.

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform (Windows · macOS · Linux) · PowerShell 7</sub>

---

## What you get

- **Summary dashboard** with OS breakdown + enabled/disabled counts
- **Full device list** — name, OS, last sign-in, days inactive, trust type, management type
- **OS filters** — All / Windows / macOS / Linux / iOS / Android / custom
- **Bulk actions** — disable or delete, with preview + typed-YES confirmation
- **Excel action log** recording every attempt with success/fail

## Quick start

```powershell
pwsh ./stale-devices-cleanup.ps1
```

Prompts for: admin UPN, stale threshold (30/60/90/custom), OS filter.

## Why this tool?

Stale devices in Entra ID are a security risk — they clutter the directory, inflate compliance reports, and can be exploited if their identities are still trusted. The Entra admin center is slow for this task and lacks granular control over OS type, threshold, or bulk actions. This tool automates the full workflow.

Microsoft guidance: [Manage stale devices in Microsoft Entra ID](https://learn.microsoft.com/en-us/entra/identity/devices/manage-stale-devices)

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- `ImportExcel` module — auto-installs if missing
- Permissions: `Device.Read.All`, `Directory.Read.All` (audit); `Device.ReadWrite.All` (disable/delete)

## How it works

Single flow — audit first, then optionally act:

1. **Setup** — admin UPN, threshold, OS filter
2. **Connect** — Microsoft Graph sign-in via browser
3. **Fetch devices** — paginated Graph pull
4. **Filter & audit** — identify stale by last sign-in, apply OS filter, summarize
5. **Review & decide** — export only / disable / delete
6. **Export** — Excel regardless of action

## Output

`StaleDevices_Report_<timestamp>.xlsx` on your Desktop:

- **Summary** — KPI dashboard with OS breakdown
- **Stale Devices** — full list
- **Action Log** — disable/delete results (appears only when actions taken)

## Safety

- Audit-only by default.
- Each destructive action requires a preview + typing `YES`.
- Action log preserves which devices were touched and the result.

## Related tools

- [Entra-StaleApps-Cleanup](../Entra-StaleApps-Cleanup/) — same pattern for app registrations
- [Entra-AppCredentials-Audit](../Entra-AppCredentials-Audit/) — catch expiring app credentials
