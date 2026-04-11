# CA-Update-AffectedApps

## Why this tool?

Starting **May 13, 2026**, Microsoft is changing how Conditional Access policies enforce rules on applications that request only basic authentication scopes (OIDC or limited directory scopes). Until now, these apps could bypass CA policies that target "All resources" when resource exclusions were present. After the change, **all CA policies will be enforced on these sign-ins**, closing a security gap.

If your tenant has apps that rely on minimal scopes and can't handle CA challenges (like MFA or device compliance), they may break after this change.

This tool scans your tenant and identifies those affected apps so you can prepare in advance.

**Microsoft announcement:** [Upcoming Conditional Access change: Improved enforcement for policies with resource exclusions](https://techcommunity.microsoft.com/blog/microsoft-entra-blog/upcoming-conditional-access-change-improved-enforcement-for-policies-with-resour/4488925)

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph module (auto-installs if missing)
- ImportExcel module (auto-installs if missing)
- Delegated Graph permissions: `DelegatedPermissionGrant.Read.All`, `Directory.Read.All`, `Reports.Read.All`, `AuditLog.Read.All`

## How it works

The script runs 6 phases:

1. **Connect** — signs in to Microsoft Graph with delegated permissions (browser prompt)
2. **Grants** — fetches all delegated permission grants (`oauth2PermissionGrants`) and identifies apps using only baseline OIDC/directory scopes (e.g. `openid`, `profile`, `User.Read`)
3. **Sign-in summary** — pulls app sign-in activity for the selected audit period (7 or 30 days)
4. **Enrich apps** — resolves service principal details, classifies apps as tenant-owned or external
5. **MFA audit** — checks sign-in logs to determine if each active app's users signed in with MFA or single-factor (SFA)
6. **Export** — generates a formatted Excel report on your Desktop

## Risk levels

| Risk | Meaning |
|------|---------|
| HIGH | SFA-only sign-ins detected — these apps will likely break when CA enforces MFA |
| MEDIUM | Active sign-ins all via MFA — should handle enforcement, but monitor |
| LOW | No recent sign-in activity — lower immediate risk |
| UNKNOWN | Audit log query failed — verify manually |

## Output

An Excel file saved to your Desktop: `ConditionalAccess_Readiness_Report_<timestamp>.xlsx`

Sheets included:
- **Executive Dashboard** — KPI summary with risk breakdown
- **Tenant-Owned Apps** — apps registered in your tenant, sorted by risk
- **External Apps** — third-party/multi-tenant apps, sorted by risk
- **Scopes Reference** — list of baseline scopes that define which apps are in scope

## Usage

```powershell
pwsh ./ca-affected-apps.ps1
```

The script will prompt for:
- Global Admin UPN
- Audit period (7 or 30 days)
