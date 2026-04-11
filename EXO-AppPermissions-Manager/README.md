# EXO-AppPermissions-Manager

## Why this tool?

When apps need access to Exchange Online mailboxes, Microsoft recommends scoping them to **specific mailboxes** instead of granting tenant-wide access. This requires creating management scopes, service principals, and role assignments — a multi-step PowerShell process that's easy to get wrong.

This tool automates all of it in one script.

**Microsoft documentation:** [Limiting application permissions to specific Exchange Online mailboxes](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)

## Requirements

- PowerShell 5.1+
- ExchangeOnlineManagement module
- Exchange Online Administrator or Global Administrator role

## How it works

The script performs 4 steps:

1. **Management scope** — creates (or reuses) a management scope that filters to the target mailbox
2. **Service principal** — ensures the app's service principal exists in Exchange Online
3. **Role assignment** — adds or removes Exchange application roles for the app, scoped to the target mailbox
4. **Authorization test** — verifies the app can access the mailbox with the assigned permissions

In **verification mode** (`-VerifyOnly`), it only reads current state without making changes.

## Supported mailbox types

Shared, User, Room, Equipment, Distribution Groups

## Available roles

13 Exchange application roles including Mail (Read/ReadWrite/Send), Calendar (Read/ReadWrite), Contacts (Read/ReadWrite), MailboxSettings (Read/ReadWrite), and EWS access. Run `-ListAvailableRoles` to see the full list.

## Usage

```powershell
# Interactive mode — prompts for everything
.\Configure-AppPermissions.ps1

# Configure default roles (Mail.ReadWrite + Mail.Send + Calendars.ReadWrite)
.\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "mailbox@domain.com"

# Add specific roles
.\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "mailbox@domain.com" -RolesToAdd @("Application Mail.Read")

# Remove specific roles
.\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "mailbox@domain.com" -RolesToRemove @("Application Mail.Send")

# Verify only — no changes
.\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "mailbox@domain.com" -VerifyOnly

# List all available roles
.\Configure-AppPermissions.ps1 -ListAvailableRoles
```

## Files

| File | Purpose |
|------|---------|
| `Configure-AppPermissions.ps1` | Main script — configuration + verification |
| `CONFIG.ps1` | Template for environment variables |
| `EXAMPLES.ps1` | Usage examples for all scenarios |

## Notes

- After configuration, you still need to add the matching **Microsoft Graph API permissions** in Azure Portal and grant admin consent
- Permission propagation can take **30-60 minutes**
- Exchange only allows **one management scope per mailbox filter** — the script handles this by reusing existing scopes
