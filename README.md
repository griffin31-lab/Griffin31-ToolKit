<div align="center">

<img src="https://avatars.githubusercontent.com/u/230988388?s=160" alt="Griffin31" width="120" height="120" style="border-radius: 20px"/>

# Griffin31 ToolKit

### Open-source security tooling for Microsoft 365, Entra ID, and cloud identity.

<p>
  Built by security engineers, for security engineers. Production-grade PowerShell tools<br/>
  that complement the <a href="https://www.griffin31.com"><strong>Griffin31</strong></a> posture-management platform.
</p>

<p>
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/stargazers"><img alt="Stars" src="https://img.shields.io/github/stars/griffin31-lab/Griffin31-ToolKit?style=for-the-badge&color=4472C4&labelColor=1B2A4A"/></a>
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/blob/main/LICENSE"><img alt="License" src="https://img.shields.io/badge/License-MIT-4472C4?style=for-the-badge&labelColor=1B2A4A"/></a>
  <img alt="PowerShell" src="https://img.shields.io/badge/PowerShell-7.x-4472C4?style=for-the-badge&labelColor=1B2A4A&logo=powershell&logoColor=white"/>
  <img alt="Platform" src="https://img.shields.io/badge/Windows%20%7C%20macOS-lightgrey?style=for-the-badge&labelColor=1B2A4A"/>
</p>

<p>
  <a href="https://www.griffin31.com"><strong>Website</strong></a> &nbsp;·&nbsp;
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/discussions"><strong>Discussions</strong></a> &nbsp;·&nbsp;
  <a href="https://github.com/griffin31-lab/Griffin31-ToolKit/issues"><strong>Report an Issue</strong></a> &nbsp;·&nbsp;
  <a href="SECURITY.md"><strong>Security Policy</strong></a>
</p>

</div>

---

## Why Griffin31 ToolKit?

Microsoft 365 and Entra ID are the modern enterprise's front door — and the most frequent target. Built-in tooling doesn't cover every corner: stale devices, over-permissioned apps, expiring credentials, Conditional Access blind spots, SPF misconfigurations, Exchange permission sprawl.

The ToolKit is a **focused set of production-grade scripts** — each one solves a single real problem we hit in the field. No installers, no cloud back-end, no telemetry. Download, run, review the output.

Used internally at Griffin31 against real customer tenants, published here for the community.

---

## The Tools

<table>
<tr>
<td width="50%" valign="top">

### [CA-Policy-Analyzer](CA-Policy-Analyzer/)

**Conditional Access posture, gaps, and insights**

Exports your full CA configuration, scores every policy 0-100, flags tenant-wide gaps against Microsoft's 2026 best practices — including the May 2026 enforcement change — and produces a self-contained HTML report with posture score, priority-sorted insights, and per-policy drill-down.

`Conditional Access` · `Entra ID` · `Posture` · `Zero Trust`

</td>
<td width="50%" valign="top">

### [Entra-AppCredentials-Audit](Entra-AppCredentials-Audit/)

**Catch expired and expiring app registration credentials**

Scans every app registration, flags expired and soon-to-expire certificates and client secrets, resolves owners, and optionally removes expired credentials. Generates an Excel report with direct links into the Entra portal.

`Entra ID` · `App Registrations` · `Credential Hygiene`

</td>
</tr>
<tr>
<td width="50%" valign="top">

### [CA-Update-AffectedApps](CA-Update-AffectedApps/)

**Prepare for Microsoft's May 2026 CA enforcement change**

Identifies tenant apps using basic OIDC scopes, cross-references sign-in logs for MFA status, and generates an Excel risk report so you can remediate before Microsoft's enforcement change breaks authentication.

`Conditional Access` · `App Assessment` · `MFA`

</td>
<td width="50%" valign="top">

### [Entra-StaleDevices-Cleanup](Entra-StaleDevices-Cleanup/)

**Audit, disable, or delete stale devices**

Finds devices that haven't signed in for X days, filters by OS and ownership, shows a full audit with MDM info, then gives you the decision — export only, disable, or delete.

`Entra ID` · `Device Management` · `Compliance`

</td>
</tr>
<tr>
<td width="50%" valign="top">

### [EXO-AppPermissions-Manager](EXO-AppPermissions-Manager/)

**Exchange Online app-to-mailbox scoping, automated**

Creates management scopes, assigns roles, and verifies configuration in one flow. Supports all 13 Exchange application roles and every mailbox type.

`Exchange Online` · `RBAC` · `Mailbox Scoping`

</td>
<td width="50%" valign="top">

### [SPF-Lookup-Validator](SPF-Lookup-Validator/)

**RFC 7208-compliant SPF chain analysis**

Recursively walks your entire SPF include chain, counts the real DNS lookup total against the 10-lookup limit, and catches misconfigurations before they break email delivery.

`SPF` · `Email Security` · `DNS`

</td>
</tr>
</table>

---

## Getting Started

```powershell
# Clone the repository
git clone https://github.com/griffin31-lab/Griffin31-ToolKit.git
cd Griffin31-ToolKit

# Pick a tool, open its folder, follow its README
cd CA-Policy-Analyzer
pwsh ./CA-Manager.ps1
```

Every tool is self-contained. Most tools require:

- **PowerShell 7.x** — [install guide](https://aka.ms/install-powershell)
- **Microsoft.Graph module** — auto-installs on first run
- **Delegated admin permissions** — each tool's README lists exact scopes

No Griffin31 account needed. No telemetry. No network calls outside Microsoft Graph.

---

## Design Principles

<table>
<tr>
<td width="33%" valign="top">

#### **One problem, one tool**

Every script solves a single, real, high-frequency problem. No frameworks, no abstractions you don't need.

</td>
<td width="33%" valign="top">

#### **Safe by default**

Read-only by default. Any destructive action requires explicit confirmation. Break-glass logic baked in.

</td>
<td width="33%" valign="top">

#### **Honest output**

Reports show what the tool can and cannot determine. No inflated scores, no marketing numbers — the output is what you'd want in an audit.

</td>
</tr>
</table>

---

## Security

Found a vulnerability? Please follow our [security policy](SECURITY.md) — responsible disclosure to **security@griffin31.ai**. We respond within 48 hours.

All tools are scanned for supply-chain risks and reviewed against OWASP and Microsoft Graph least-privilege guidance before release.

---

## Contributing

We welcome contributions — bug fixes, new tools, documentation improvements.

- **Issues** — [open a ticket](https://github.com/griffin31-lab/Griffin31-ToolKit/issues) for bugs and feature requests
- **Discussions** — [join the conversation](https://github.com/griffin31-lab/Griffin31-ToolKit/discussions) for questions, ideas, and showcases
- **Contribution guide** — see [CONTRIBUTING.md](CONTRIBUTING.md)
- **Code of conduct** — see [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)

---

## About Griffin31

[Griffin31](https://www.griffin31.com) is building the security posture management platform for Microsoft 365 — continuous monitoring, prioritized recommendations, and automated remediation for Entra ID, Exchange Online, SharePoint, Teams, and Intune.

The ToolKit is the free, open-source foundation of the same analysis engines that power our commercial product.

<div align="center">
  <br/>
  <a href="https://www.griffin31.com">
    <img src="https://avatars.githubusercontent.com/u/230988388?s=80" alt="Griffin31" width="60" height="60" style="border-radius: 12px"/>
  </a>
  <br/>
  <sub><strong>Made with care by the Griffin31 team</strong></sub>
  <br/>
  <sub><a href="https://www.griffin31.com">griffin31.com</a> · <a href="https://github.com/griffin31-lab">@griffin31-lab</a></sub>
</div>

---

<div align="center">
<sub>
Released under the <a href="LICENSE">MIT License</a>. Free to use, modify, and distribute.
</sub>
</div>
