# CA-Update-AffectedApps

> **Prepare for Microsoft's May 13 2026 Conditional Access enforcement change** — Excel risk report with MFA status per app.

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform (Windows · macOS · Linux) · PowerShell 7</sub>

---

## What you get

- **Executive dashboard** — KPI summary with risk breakdown
- **Per-app risk classification** — HIGH / MEDIUM / LOW / UNKNOWN
- **MFA vs SFA breakdown** from actual sign-in logs
- **Tenant-owned vs external apps** on separate sheets
- **Scopes reference** — exactly which baseline scopes put an app in scope

## Quick start

```powershell
pwsh ./ca-affected-apps.ps1
```

The script will prompt for: Global Admin UPN, audit period (7 or 30 days).

## Why this tool?

Starting **May 13, 2026**, Microsoft is changing how Conditional Access policies enforce rules on applications that request only basic authentication scopes (OIDC or limited directory scopes). Until now, these apps could bypass CA policies that target "All resources" when resource exclusions were present. After the change, **all CA policies will be enforced on these sign-ins** — closing a security gap.

If your tenant has apps that rely on minimal scopes and can't handle CA challenges (MFA, device compliance), they may break after this change.

Microsoft announcement: [Upcoming Conditional Access change](https://techcommunity.microsoft.com/blog/microsoft-entra-blog/upcoming-conditional-access-change-improved-enforcement-for-policies-with-resour/4488925)

## Requirements

- PowerShell 7.x (Windows or macOS)
- `Microsoft.Graph` module — auto-installs if missing
- `ImportExcel` module — auto-installs if missing
- Delegated Graph permissions: `DelegatedPermissionGrant.Read.All`, `Directory.Read.All`, `Reports.Read.All`, `AuditLog.Read.All`

## How it works

Six-phase flow:

1. **Connect** — Microsoft Graph sign-in via browser
2. **Grants** — fetches delegated permission grants, identifies apps using only baseline OIDC/directory scopes
3. **Sign-in summary** — pulls app sign-in activity for your audit window
4. **Enrich apps** — resolves service principals, classifies tenant-owned vs external
5. **MFA audit** — checks sign-in logs per app for MFA vs single-factor
6. **Export** — formatted Excel report to Desktop

## Risk levels

| Risk | Meaning |
|------|---------|
| **HIGH** | SFA-only sign-ins detected — will likely break when CA enforces MFA |
| **MEDIUM** | All active sign-ins are MFA — should handle enforcement, but monitor |
| **LOW** | No recent sign-in activity — lower immediate risk |
| **UNKNOWN** | Audit log query failed — verify manually |

## Output

`ConditionalAccess_Readiness_Report_<timestamp>.xlsx` on your Desktop:

- **Executive Dashboard** — KPI summary with risk breakdown
- **Tenant-Owned Apps** — apps registered in your tenant, sorted by risk
- **External Apps** — third-party / multi-tenant apps, sorted by risk
- **Scopes Reference** — baseline scopes that define which apps are in scope

## Related tools

- [CA-Policy-Analyzer](../CA-Policy-Analyzer/) — full CA posture report with the May 2026 enforcement check + 14 other insights
- [Entra-AppCredentials-Audit](../Entra-AppCredentials-Audit/) — for apps flagged HIGH, also check their credential hygiene
