# Detection Rules

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

Scheduled analytic rules (Sentinel YAML) and Defender XDR custom detections (markdown per rule) — productionized forms of the KQL hunts in [../kql/](../kql/). Sentinel YAML follows the format used in the [Azure/Azure-Sentinel](https://github.com/Azure/Azure-Sentinel/tree/master/Detections) public repo.

File prefix tells you the platform: `sentinel-*.yaml` for Sentinel analytics, `defender-xdr-*.md` for Defender XDR custom detections.

## Sentinel analytic rules

| # | Rule | File | Severity | Tactic | MITRE |
|---|---|---|---|---|---|
| 1 | Suspicious OAuth Consent | [sentinel-SuspiciousOAuthConsent.yaml](./sentinel-SuspiciousOAuthConsent.yaml) | High | Initial Access | T1528 |
| 2 | Impossible Travel | [sentinel-ImpossibleTravel.yaml](./sentinel-ImpossibleTravel.yaml) | High | Initial Access | T1078.004 |
| 3 | High-Privilege Role Assignment | [sentinel-HighPrivilegeRoleAssignment.yaml](./sentinel-HighPrivilegeRoleAssignment.yaml) | High | Privilege Escalation | T1098.003 |
| 4 | New App Registration | [sentinel-NewAppRegistration.yaml](./sentinel-NewAppRegistration.yaml) | Medium | Persistence | T1098.001 |
| 5 | Mailbox Forwarding Rule Created | [sentinel-MailboxForwardingRuleCreated.yaml](./sentinel-MailboxForwardingRuleCreated.yaml) | Medium | Collection | T1114.003 |
| 6 | Legacy Authentication Success | [sentinel-LegacyAuthSuccess.yaml](./sentinel-LegacyAuthSuccess.yaml) | Medium | Credential Access | T1110 |
| 7 | MFA via Weak / Phishable Method | [sentinel-MFAWeakMethodBypass.yaml](./sentinel-MFAWeakMethodBypass.yaml) | Medium | Credential Access | T1621 |
| 8 | Mass SharePoint / OneDrive Download | [sentinel-MassSharePointDownload.yaml](./sentinel-MassSharePointDownload.yaml) | High | Exfiltration | T1530 |
| 9 | Anonymous Proxy / Tor Sign-in | [sentinel-AnonymousProxyTorSignin.yaml](./sentinel-AnonymousProxyTorSignin.yaml) | High | Initial Access | T1090.003 |
| 10 | Conditional Access Policy Disabled | [sentinel-ConditionalAccessDisabled.yaml](./sentinel-ConditionalAccessDisabled.yaml) | High | Defense Evasion | T1562.007 |
| 11 | Break-Glass Account Sign-in | [sentinel-BreakGlassAccountSignin.yaml](./sentinel-BreakGlassAccountSignin.yaml) | High | Initial Access | T1078.004 |
| 12 | Guest Invite by Non-Admin | [sentinel-GuestInviteByNonAdmin.yaml](./sentinel-GuestInviteByNonAdmin.yaml) | Medium | Persistence | T1136.003 |
| 13 | App Credential (Secret/Cert) Added | [sentinel-AppCredentialAdded.yaml](./sentinel-AppCredentialAdded.yaml) | High | Persistence | T1098.001 |

## Defender XDR custom detections

| # | Rule | File | Severity | Tactic | MITRE |
|---|---|---|---|---|---|
| 14 | LSASS Memory Access | [defender-xdr-LSASSMemoryAccess.md](./defender-xdr-LSASSMemoryAccess.md) | High | Credential Access | T1003.001 |
| 15 | Infostealer Indicator Match | [defender-xdr-InfostealerIndicator.md](./defender-xdr-InfostealerIndicator.md) | High | Credential Access / Exfiltration | T1555 / T1567 |

## Deployment

### Sentinel (`sentinel-*.yaml`)

1. Review the rule in-place; tune thresholds and allow-lists for your tenant.
2. Deploy via one of:
   - **Azure portal** — Sentinel → Analytics → Create → Import from template → paste YAML
   - **ARM / Bicep** — wrap the YAML body in an ARM template and deploy via `New-AzResourceGroupDeployment`
   - **GitHub Actions** — the Azure-Sentinel repo's sync action works unchanged on this folder
3. Validate that all required `connectorId`s are enabled in your workspace.
4. Some rules reference Sentinel watchlists (`BreakGlassAccounts`, `AuthorizedGuestInviters`) — create those before enabling.

### Defender XDR (`defender-xdr-*.md`)

1. Open [security.microsoft.com](https://security.microsoft.com) → Hunting → Custom detection rules → Create
2. Paste the KQL from the markdown file
3. Set severity, frequency, entity mappings to match the doc
4. Configure automated response (isolation, investigation package) per the file's "Suggested response actions" section

## Tuning & expected FP rate

| Rule | Expected FP rate | Primary noise source | Recommended mitigation |
|---|---|---|---|
| Suspicious OAuth Consent | Low if user consent is restricted | First-party Microsoft apps during onboarding | Allow-list by publisher / app ID |
| Impossible Travel | Medium | VPN hops, Entra Connect Sync, service accounts | Exclude service accounts; raise threshold for VPN-heavy users |
| Mass SharePoint Download | Medium | Bulk export by legitimate users, OneDrive sync events | Baseline per user (included in query) |
| MFA Weak Method | Medium | Users who genuinely prefer SMS | Compare to user's 14-day baseline (included) |
| Legacy Auth Success | Low on modern tenants | Legacy POP/IMAP mail clients | Block legacy auth via CA — then any hit is true positive |

## Disclaimer

Always validate in a non-production workspace before production use. Thresholds and allow-lists are starting points — tune to your tenant's baseline noise profile.
