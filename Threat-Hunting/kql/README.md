# KQL Hunt Queries

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

KQL queries for Microsoft Sentinel, Defender XDR Advanced Hunting, and Azure Monitor Log Analytics. Each `.kql` file is self-contained — metadata header, full incident/background context, tuning notes, hardening guidance, source attribution, and the query body are all inline as comments.

## Index

### OAuth / Application Consent
- [App-Credential-Added.kql](./App-Credential-Added.kql) — Password/certificate added to an app registration (credential persistence)
- [First-Time-App-Consenter.kql](./First-Time-App-Consenter.kql) — User consenting for the first time to any OAuth app
- [First-Time-App-Credential-Addition.kql](./First-Time-App-Credential-Addition.kql) — First credential ever added to an app (pivot for backdoor detection)
- [Graph-Mail-Permissions-Added.kql](./Graph-Mail-Permissions-Added.kql) — New Graph Mail.* permissions granted to any app
- [High-Privilege-Graph-Permissions-Added.kql](./High-Privilege-Graph-Permissions-Added.kql) — High-impact Graph scopes (Mail.*, Files.*, Sites.FullControl.*) granted
- [Service-Principal-Added-To-Role.kql](./Service-Principal-Added-To-Role.kql) — SP added to a directory role (non-interactive privilege escalation)
- [Service-Principal-Signin-New-IP.kql](./Service-Principal-Signin-New-IP.kql) — SP signing in from a previously-unseen IP
- [Suspicious-OAuth-Consent.kql](./Suspicious-OAuth-Consent.kql) — OAuth consent grants and delegated permission assignments (Vercel/Context.ai breach pattern)

### Identity / Authentication
- [Break-Glass-Account-Signin.kql](./Break-Glass-Account-Signin.kql) — Any sign-in by a break-glass / emergency account (should be vanishingly rare)
- [Device-Code-Signin-Unmanaged-Device.kql](./Device-Code-Signin-Unmanaged-Device.kql) — Device code flow from unmanaged device (Storm-2372 pattern)
- [Entra-Role-Additions.kql](./Entra-Role-Additions.kql) — Any Entra directory role assignment
- [First-Time-Legacy-Auth-Success.kql](./First-Time-Legacy-Auth-Success.kql) — Legacy auth success from a user/IP not seen before
- [Impossible-Travel-Multi-Country.kql](./Impossible-Travel-Multi-Country.kql) — Multi-country sign-ins for one user within an implausibly short window
- [MFA-Only-SMS-Method.kql](./MFA-Only-SMS-Method.kql) — Users whose only MFA method is SMS (SIM-swap exposure)
- [Privileged-User-Single-Factor-Signin.kql](./Privileged-User-Single-Factor-Signin.kql) — Global Admin / Priv-Role user signing in with single factor
- [SSPR-Followed-By-Risky-Signin.kql](./SSPR-Followed-By-Risky-Signin.kql) — Self-service password reset followed by risky sign-in (Scattered Spider pattern)
- [Signin-From-New-Country.kql](./Signin-From-New-Country.kql) — First sign-in from a country not seen in user's 90-day baseline

### Data Access (SharePoint / OneDrive / Exchange)
- [Anomalous-Access-Other-Mailboxes.kql](./Anomalous-Access-Other-Mailboxes.kql) — User accessing other users' mailboxes (NOBELIUM pattern)
- [Exchange-Audit-Log-Disabled.kql](./Exchange-Audit-Log-Disabled.kql) — Mailbox audit logging disabled
- [External-User-SharePoint-Access.kql](./External-User-SharePoint-Access.kql) — External user activity summary on SharePoint
- [Inbox-Forwarding-Rule-Created.kql](./Inbox-Forwarding-Rule-Created.kql) — New inbox rule with external forwarding (classic BEC indicator)
- [Mailbox-ForwardingSmtpAddress-Set.kql](./Mailbox-ForwardingSmtpAddress-Set.kql) — `Set-Mailbox -ForwardingSmtpAddress` invoked
- [Multiple-Users-Forwarded-Same-Destination.kql](./Multiple-Users-Forwarded-Same-Destination.kql) — Multiple compromised mailboxes forwarding to the same external address
- [SharePoint-Bulk-Download.kql](./SharePoint-Bulk-Download.kql) — User downloading abnormally high volume from SharePoint/OneDrive
- [Teams-Files-Uploaded-Access-Summary.kql](./Teams-Files-Uploaded-Access-Summary.kql) — Teams file upload activity summary

### Persistence (Apps, Roles, CA, Guests)
- [App-Ownership-Changed.kql](./App-Ownership-Changed.kql) — Owner added/removed on an app registration
- [CA-Policy-Deleted-Or-Disabled.kql](./CA-Policy-Deleted-Or-Disabled.kql) — Conditional Access policy deleted or disabled
- [CA-Policy-State-Report-Only-Downgrade.kql](./CA-Policy-State-Report-Only-Downgrade.kql) — CA policy silently downgraded to report-only
- [Credential-Added-To-App-Registration.kql](./Credential-Added-To-App-Registration.kql) — Secret or cert added to an existing app
- [Custom-Security-Attribute-Set.kql](./Custom-Security-Attribute-Set.kql) — Custom security attribute modified (abused for CA-policy bypass)
- [Federation-Trust-Modified.kql](./Federation-Trust-Modified.kql) — Federation trust changed (Golden SAML precursor)
- [Guest-Invited-By-New-Inviter.kql](./Guest-Invited-By-New-Inviter.kql) — First-time inviter creates a guest account
- [New-App-Registration-First-Time-Actor.kql](./New-App-Registration-First-Time-Actor.kql) — App registered by an actor who has never registered one before
- [Role-Assigned-Outside-PIM.kql](./Role-Assigned-Outside-PIM.kql) — Permanent role assignment bypassing PIM

### Endpoint (Defender for Endpoint)
- [BloodHound-SharpHound-Collector.kql](./BloodHound-SharpHound-Collector.kql) — SharpHound / AzureHound execution (AD / Entra ID attack-path enumeration)
- [LOLBAS-From-Office-App.kql](./LOLBAS-From-Office-App.kql) — `rundll32` / `regsvr32` / `mshta` spawned from Office apps
- [LSASS-Memory-Access.kql](./LSASS-Memory-Access.kql) — LSASS memory read attempts (Mimikatz / ProcDump core step)
- [Lumma-Stealer-Indicators.kql](./Lumma-Stealer-Indicators.kql) — Lumma Stealer HTTP exfil via `TeslaBrowser/5.5` user-agent
- [Persistence-Scheduled-Task-Run-Key.kql](./Persistence-Scheduled-Task-Run-Key.kql) — Scheduled tasks, registry Run keys, Startup drops
- [PowerShell-Encoded-Command.kql](./PowerShell-Encoded-Command.kql) — Decodes Base64 PowerShell inline

### Exfiltration
- [Cloud-Storage-Exfil.kql](./Cloud-Storage-Exfil.kql) — Outbound to mega, wetransfer, transfer.sh, rclone process
- [DNS-Tunneling-Indicators.kql](./DNS-Tunneling-Indicators.kql) — Long FQDNs / high subdomain fan-out (DNS tunneling)
- [OneDrive-SharePoint-Bulk-Download.kql](./OneDrive-SharePoint-Bulk-Download.kql) — Per-user time-series anomaly on cloud-file downloads
- [Unusual-GraphAPI-UserAgent.kql](./Unusual-GraphAPI-UserAgent.kql) — Graph API called from non-Microsoft clients (ROADtools, GraphRunner)

## File format

Every `.kql` file has this structure inline:

```kql
// Name:          <short name>
// Description:   <one line>
// MITRE ATT&CK:  T<id> (<technique name>), ...
// Data Source:   <tables required>
// Platform:      Sentinel | Defender XDR | Advanced Hunting | Log Analytics
// Time window:   <default lookback>
// Context:       <what incident / TTP motivated this>
//
// Recommended actions on hits:
// - <action>
//
// ---------------------------------------------------------------------------
// Extended context, background, tuning, and attribution
// ---------------------------------------------------------------------------
//
// <full background, expected output table, tuning, hardening, references,
//  and MIT/BSD-3 attribution — all as // comments>

<query body>
```

## Sources

All queries sourced from permissively-licensed public repos (full attribution in each file):
- [Azure/Azure-Sentinel](https://github.com/Azure/Azure-Sentinel) · MIT
- [Bert-JanP/Hunting-Queries-Detection-Rules](https://github.com/Bert-JanP/Hunting-Queries-Detection-Rules) · BSD-3-Clause
- [reprise99/Sentinel-Queries](https://github.com/reprise99/Sentinel-Queries) · MIT
- [cyb3rmik3/KQL-threat-hunting-queries](https://github.com/cyb3rmik3/KQL-threat-hunting-queries) · MIT
- [microsoft/Microsoft-365-Defender-Hunting-Queries](https://github.com/microsoft/Microsoft-365-Defender-Hunting-Queries) · MIT (archived)
