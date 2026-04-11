# CA-Update-AffectedApps

## Why this tool?

Starting **May 13, 2026**, Microsoft is changing how Conditional Access policies enforce rules on applications that request only basic authentication scopes (OIDC or limited directory scopes). Until now, these apps could bypass CA policies that target "All resources" when resource exclusions were present. After the change, **all CA policies will be enforced on these sign-ins**, closing a security gap.

If your tenant has apps that rely on minimal scopes and can't handle CA challenges (like MFA or device compliance), they may break after this change.

This tool scans your tenant and identifies those affected apps so you can prepare in advance.

**Microsoft announcement:** [Upcoming Conditional Access change: Improved enforcement for policies with resource exclusions](https://techcommunity.microsoft.com/blog/microsoft-entra-blog/upcoming-conditional-access-change-improved-enforcement-for-policies-with-resour/4488925)

## Requirements

- PowerShell 7.x (Windows or macOS)
- Microsoft.Graph PowerShell module (auto-installs if missing)
- ImportExcel module (auto-installs if missing)
- Permissions: read access to app registrations and service principals in your tenant

## Usage

```powershell
pwsh ./ca-affected-apps.ps1
```

The script will:

1. Connect to Microsoft Graph interactively
2. Scan all app registrations and service principals
3. Identify apps using only basic OIDC/directory scopes
4. Generate an Excel report of affected apps
