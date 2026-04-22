# Suspicious OAuth Consent Activity

<sub>[← Back to KQL index](../README.md)</sub>

## Source incident: Vercel / Context.ai Breach (April 19, 2026)

Vercel disclosed a security incident traced to a third-party AI tool called Context.ai. The chain of events is a textbook SaaS-to-SaaS supply chain attack:

1. A Context.ai employee was infected with **Lumma Stealer** infostealer malware in February 2026 (Microsoft tracks Lumma as **Storm-2477**). Corporate credentials and OAuth tokens were exfiltrated.
2. The attacker used the stolen OAuth token to access Context.ai's Google Workspace.
3. From there they discovered that a Vercel employee had signed up for Context.ai using their **Vercel enterprise Google account** and granted "Allow All" permissions.
4. The attacker pivoted into Vercel's Google Workspace, then into internal Vercel environments.
5. They harvested customer **environment variables that were not marked as "Sensitive"** — API keys, DB credentials, signing keys, webhook secrets.
6. The dataset was offered for sale on BreachForums for $2M.

**The critical point:** Vercel was not breached through their product, code, or infrastructure. The attack rode in through a third-party SaaS app that one employee connected to a corporate identity. **This pattern is not specific to Vercel or Google Workspace** — it works identically against Microsoft 365 / Entra ID environments.

## Why this hunt

In an M365 / Entra ID context, the equivalent attack surface is **OAuth apps in Entra ID with delegated permissions to Microsoft Graph**. The `AuditLogs` table records every consent grant, delegated-permission addition, and app-role assignment. Running this query against the last 90 days surfaces every consent event — giving the analyst a complete picture of which apps have been trusted, by whom, and when.

## Expected output

Each row represents one consent / permission operation:

| Column | What it tells you |
|---|---|
| `TimeGenerated` | When the consent happened |
| `AppName` | Display name of the consented app |
| `AppId` | Entra object ID of the service principal |
| `OperationName` | Which of the 4 consent operations fired |
| `InitiatedByUser` | UPN of the person who performed the consent |
| `ModifiedProps` | Raw JSON of the before/after state — includes the exact permissions granted |

A **normal tenant** will see entries from:
- Microsoft first-party apps (e.g., Office, Teams, Outlook) during onboarding
- Known admin-consented apps your org has approved
- Regular "end users consenting to low-impact apps" if legacy consent is still allowed

A **suspicious hit** has one or more of:
- `AppName` you don't recognize, especially with generic names (e.g., "My App", "Test App")
- High-privilege permissions in `ModifiedProps` (`Mail.Read`, `Files.Read.All`, `Sites.FullControl.All`, `User.ReadWrite.All`)
- `InitiatedByUser` who shouldn't be consenting (standard users rather than designated reviewers)
- Consents happening outside business hours or from unusual geographies (cross-reference with `SigninLogs`)

## Tuning notes

- **Legacy tenants** — if you haven't restricted user consent, this query can return hundreds of low-impact app consents. Filter out known Microsoft first-party app IDs first.
- **Audit log retention** — adjust `ago(90d)` to your workspace's retention window.
- **Shorten for alerting** — for a scheduled analytic rule, use `ago(1h)` and alert on any non-allow-listed `AppId`.

## Recommended hardening (prevent, not just detect)

Detection without prevention is fragile. Pair this hunt with these Entra ID controls:

1. **Restrict User Consent** — Entra admin > Enterprise apps > Consent and permissions > set to *"Allow user consent for apps from verified publishers, for selected permissions"* (`microsoft-user-default-low` policy)
2. **Enable the Admin Consent Workflow** — so users have a sanctioned path to request high-impact permissions
3. **Classify permissions** — define which Graph permissions are "low impact" and which are not, then wire the classification into the consent policy
4. **App Governance in Defender for Cloud Apps** — the strongest detective control for this class: flags OAuth apps with rare community use, suspicious reply URLs, anomalous Graph activity, unusual user agents
5. **Token Protection** (Conditional Access) — binds tokens to the originating device, neutralizing the token-replay portion of this attack chain
6. **Block Device Code Flow** in Conditional Access wherever it isn't explicitly needed (common infostealer exfiltration path)
7. **Periodic review** — restricting future consent does not remove existing grants. Review Enterprise Applications quarterly; revoke anything from unverified publishers, anything with high-privilege scope, anything authorized by a small user count

## Referenced IOC (Vercel incident)

Vercel published this OAuth App ID as an IOC — Google Workspace admins should check for it, but since it is a Google-side identifier it does not appear in Entra ID logs:

```
110671459871-30f1spbu0hptbs60cb4vsmv79i7bbvqj.apps.googleusercontent.com
```

Our mirror in Entra ID terms is any app with broad delegated Graph scope granted by an end user — the format of this hunt.

## References

- [Microsoft Learn: Configure user consent settings](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-user-consent)
- [MITRE ATT&CK T1528 — Steal Application Access Token](https://attack.mitre.org/techniques/T1528/)
- [MITRE ATT&CK T1550.001 — Use Alternate Authentication Material: Application Access Token](https://attack.mitre.org/techniques/T1550/001/)
- [Microsoft Threat Intelligence: Storm-2477 (Lumma Stealer)](https://learn.microsoft.com/en-us/unified-secops-platform/threat-actor-naming)
