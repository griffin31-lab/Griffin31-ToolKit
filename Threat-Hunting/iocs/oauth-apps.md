# Malicious OAuth Applications

<sub>[← Back to IoCs](../) · [← Back to Threat-Hunting](../../)</sub>

OAuth application identifiers (Entra ID client IDs and Google Workspace OAuth client IDs) tied to documented consent-phishing and business-email-compromise (BEC) incidents. These are the sanctioned, attributable app IDs that have been confirmed malicious in published threat-intel reports.

## Usage

| Platform | How to apply |
|---|---|
| **Microsoft Entra ID** | Enterprise applications → search by `AppId` → if present, revoke consent + remove service principal; block the `AppId` via Consent and Permissions policy |
| **Microsoft Sentinel** | Import as a Watchlist (`MaliciousOAuthApps`) and reference from analytic rules against `AuditLogs` → `TargetResources[0].id` |
| **Defender XDR Advanced Hunting** | `CloudAppEvents` / `AADSpnSignInEventsBeta` — filter on `ApplicationId in (watchlist)` |
| **Google Workspace** | Admin Console → Security → API controls → Domain wide delegation / Third-party app access → block by Client ID |

## Indicators

| App ID | Display Name (as seen) | Platform | Campaign / Incident | First Published | Confidence | Source | Notes |
|---|---|---|---|---|---|---|---|
| `110671459871-30f1spbu0hptbs60cb4vsmv79i7bbvqj.apps.googleusercontent.com` | Context.ai OAuth client | Google Workspace | Vercel / Context.ai breach | 2026-04-19 | High | [Vercel incident disclosure](https://vercel.com/changelog) / [Griffin31 OAuth hunt](../../kql/oauth/Suspicious-OAuth-Consent.md) | SaaS supply-chain pivot; employee granted "Allow All" scope from corporate Google account |
| `2ef68ccc-8a4d-42ff-ae88-2d7bb89ad139` | PerfectData Software (rebranded to Mail_Backup) | Entra ID | BEC / mailbox exfil, Midnight Blizzard-adjacent | 2023-09 (earliest public) | High | [Darktrace analysis](https://www.darktrace.com/blog/how-abuse-of-perfectdata-software-may-create-a-perfect-storm-an-emerging-trend-in-account-takeovers) · [Cyber Corner writeup](https://cybercorner.tech/malicious-azure-application-perfectdata-software-and-office365-business-email-compromise/) | Legitimate app abused post-takeover to clone full mailbox — requests `EWS.AccessAsUser.All` + `offline_access` |
| `04b07795-8ddb-461a-bbee-02f9e1bf7b46` | Microsoft Azure CLI (first-party — abused) | Entra ID | ConsentFix campaigns | 2025-12-31 | High | [Push Security — ConsentFix](https://pushsecurity.com/blog/consentfix) · [Mitiga ConsentFix](https://www.mitiga.io/blog/consentfix-oauth-phishing-explained-how-token-based-attacks-bypass-mfa-in-microsoft-entra-id) | A legitimate Microsoft first-party app ID abused as the target client in OAuth consent phishing. Detection must focus on the *who granted it and why* rather than the app ID alone |
| `854189f9-4c71-44bb-9880-dd0c2f75922a` | Proofpoint-tracked phishing redirector | Entra ID | Adobe / RingCentral / SharePoint / DocuSign impersonation cluster | 2025-06 | High | [Proofpoint — OAuth app impersonation](https://www.proofpoint.com/us/blog/threat-insight/microsoft-oauth-app-impersonation-campaign-leads-mfa-phishing) | Used as the `client_id` on a redirector app chaining to Tycoon AitM phishing pages |
| `aebc6443-996d-45c2-90f0-388ff96faa56` | Proofpoint-tracked phishing redirector (sibling) | Entra ID | Same cluster as above | 2025-06 | High | [Proofpoint — OAuth app impersonation](https://www.proofpoint.com/us/blog/threat-insight/microsoft-oauth-app-impersonation-campaign-leads-mfa-phishing) | Parallel `client_id` observed in the same 2025 Microsoft OAuth impersonation cluster |

## Source & attribution

- [Microsoft Threat Intelligence — Storm-2477 / Lumma Stealer](https://www.microsoft.com/en-us/security/blog/2025/05/21/lumma-stealer-breaking-down-the-delivery-techniques-and-capabilities-of-a-prolific-infostealer/)
- [Darktrace — PerfectData abuse and account takeover risks](https://www.darktrace.com/blog/how-abuse-of-perfectdata-software-may-create-a-perfect-storm-an-emerging-trend-in-account-takeovers)
- [Proofpoint Threat Research — Microsoft OAuth App Impersonation Campaign](https://www.proofpoint.com/us/blog/threat-insight/microsoft-oauth-app-impersonation-campaign-leads-mfa-phishing)
- [Push Security — ConsentFix browser-native OAuth grant hijack](https://pushsecurity.com/blog/consentfix)
- [Mitiga — ConsentFix OAuth Phishing](https://www.mitiga.io/blog/consentfix-oauth-phishing-explained-how-token-based-attacks-bypass-mfa-in-microsoft-entra-id)

## Notes

- Malicious OAuth apps rotate aggressively. The `AppId` is the only durable artifact; display names are trivially changed (as PerfectData Software → Mail_Backup demonstrates).
- The Context.ai Google ID is a Google Workspace identifier and will **not** appear in Entra `AuditLogs`. Use it inside Google Workspace Admin audit logs only.
- First-party Microsoft app IDs appearing as the target of consent phishing (e.g., Azure CLI) mean the rule must evaluate *behavior* (where the consent was granted, from what IP, followed by what token use), not just the App ID.
