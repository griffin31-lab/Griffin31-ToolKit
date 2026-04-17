# EXO-AppPermissions-Manager

> **Scope Exchange Online app permissions to specific mailboxes** — automated management scope + service principal + role assignment + verification.

<sub>[← Back to Griffin31 ToolKit](../)</sub>

---

## What you get

- **One-shot configuration** — management scope, service principal, role assignment all in one flow
- **All 13 Exchange application roles** supported (Mail, Calendar, Contacts, MailboxSettings, EWS)
- **Works with every mailbox type** — Shared, User, Room, Equipment, Distribution Groups
- **Verification mode** (`-VerifyOnly`) — read current config without changing anything
- **List-roles mode** (`-ListAvailableRoles`) — discover available roles first

## Quick start

```powershell
# Interactive — prompts for everything
.\Configure-AppPermissions.ps1

# Default role set (Mail.ReadWrite + Mail.Send + Calendars.ReadWrite)
.\Configure-AppPermissions.ps1 -AppId "<app-id>" -TargetMailbox "mailbox@domain.com"

# Verify without making changes
.\Configure-AppPermissions.ps1 -AppId "<app-id>" -TargetMailbox "mailbox@domain.com" -VerifyOnly

# Discover available roles
.\Configure-AppPermissions.ps1 -ListAvailableRoles
```

## Why this tool?

Microsoft recommends scoping app access to **specific mailboxes** rather than granting tenant-wide Exchange access. Manually doing this via PowerShell takes 4 separate steps (management scope → service principal → role assignment → verification) and is easy to get wrong.

MS guidance: [Limiting application permissions to specific Exchange Online mailboxes](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)

## Requirements

- PowerShell 5.1+
- `ExchangeOnlineManagement` module
- **Exchange Online Administrator** or **Global Administrator** role

## Before you start

Fill in `CONFIG.ps1` with your environment values:

- `$AppId` — your Entra ID app registration ID
- `$TargetMailbox` — the mailbox to scope

## How it works

Four steps per app:

1. **Management scope** — creates or reuses a scope that filters to the target mailbox
2. **Service principal** — ensures the app's SP exists in Exchange Online
3. **Role assignment** — adds or removes Exchange application roles for the app, scoped to the target
4. **Authorization test** — verifies the app can access the mailbox with the assigned permissions

`-VerifyOnly` reads current state without changes. `-RolesToAdd` / `-RolesToRemove` control role changes explicitly.

## Files

| File | Purpose |
|------|---------|
| `Configure-AppPermissions.ps1` | Main script |
| `CONFIG.ps1` | Template for environment variables |
| `EXAMPLES.ps1` | Usage examples |

## Notes

- After configuration, you still need to add matching **Microsoft Graph API permissions** in the Entra portal and grant admin consent separately
- Permission propagation can take **30-60 minutes**
- Exchange allows only **one management scope per mailbox filter** — the script handles this by reusing existing scopes

## Related tools

- [Entra-AppCredentials-Audit](../Entra-AppCredentials-Audit/) — once your apps are scoped, audit their credential health
- [Entra-StaleApps-Cleanup](../Entra-StaleApps-Cleanup/) — if you inherit a mess of over-permissioned apps, start here
