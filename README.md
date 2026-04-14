<p align="center">
  <img src="https://avatars.githubusercontent.com/u/230988388?s=120" alt="Griffin31 Logo" width="120"/>
</p>

<h1 align="center">Griffin31 ToolKit</h1>

<p align="center">
  <strong>Free, open-source security tools for M365, Entra ID, and cloud platforms</strong>
</p>

<p align="center">
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT License"/></a>
  <img src="https://img.shields.io/badge/tools-5-green.svg" alt="5 Tools"/>
  <img src="https://img.shields.io/badge/PowerShell-7.x-blue.svg?logo=powershell&logoColor=white" alt="PowerShell 7"/>
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey.svg" alt="Platform"/>
</p>

<p align="center">
  <a href="https://www.griffin31.com">Website</a> &bull;
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/discussions">Discussions</a> &bull;
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/issues">Report an Issue</a>
</p>

---

Built by the [Griffin31](https://www.griffin31.com) team to help M365 admins stay secure and prepared. These tools complement our SaaS platform for security posture management, alerts, audit, and compliance.

## Tools

### [CA-Update-AffectedApps](CA-Update-AffectedApps/)
> Identifies tenant apps affected by Microsoft's upcoming Conditional Access enforcement change (May 2026)

Scans your tenant for apps using only basic OIDC scopes, checks sign-in logs for MFA status, and generates an Excel report with risk levels. Helps you prepare before the change breaks your apps.

**Tags:** `Conditional Access` `Entra ID` `App Assessment` `MFA`

---

### [EXO-AppPermissions-Manager](EXO-AppPermissions-Manager/)
> Automates Exchange Online app-to-mailbox permission scoping

Creates management scopes, assigns roles, and verifies configuration — all in one script. Supports all 13 Exchange application roles and all mailbox types.

**Tags:** `Exchange Online` `App Permissions` `RBAC` `Mailbox Scoping`

---

### [SPF-Lookup-Validator](SPF-Lookup-Validator/)
> Recursively analyzes SPF records, counts DNS lookups, and validates RFC 7208 compliance

Walks your entire SPF include chain, counts the real DNS lookup total, and warns when you're over the limit of 10. Catches issues before they cause email delivery failures.

**Tags:** `SPF` `Email Security` `DNS` `RFC 7208`

---

### [Entra-StaleDevices-Cleanup](Entra-StaleDevices-Cleanup/)
> Audit, disable, or delete stale devices in Entra ID

Finds devices that haven't signed in for X days, filters by OS type, shows a full audit with owner and MDM info, then lets you decide — export only, disable, or delete.

**Tags:** `Entra ID` `Device Management` `Cleanup` `Compliance`

---

### [Entra-AppCredentials-Audit](Entra-AppCredentials-Audit/)
> Audit expired and expiring app registration certificates and client secrets

Scans all app registrations, flags expired and soon-to-expire credentials, resolves owners, and optionally removes expired ones. Generates an Excel report with direct links to each app in the Entra portal.

**Tags:** `Entra ID` `App Registrations` `Certificates` `Secrets` `Credential Hygiene`

---

## Getting Started

Each tool is self-contained in its own folder with a dedicated README. Pick a tool, follow its instructions, and run it.

Most tools require:
- **PowerShell 7.x** — [Install](https://aka.ms/install-powershell)
- **Microsoft.Graph module** — auto-installs on first run
- **Appropriate admin permissions** — see each tool's README

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines. Issues and ideas welcome in [Discussions](https://github.com/griffin31-lab/Griffin31-ToolKit/discussions).

## License

[MIT](LICENSE) — free to use, modify, and distribute.
