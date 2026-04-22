# Threat-Hunting

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform · Microsoft Sentinel · Defender XDR · Advanced Hunting</sub>

> **Curated KQL hunt queries, IoCs, and detection rules** — built for Microsoft 365, Entra ID, Microsoft Defender XDR, and Microsoft Sentinel.

Every item in this library is tied to a documented real-world incident or a concrete attacker technique, mapped to MITRE ATT&CK, and shipped with:

- **Why it matters** — the incident or TTP that motivated the hunt
- **Data source** — the log tables / products required to run it
- **MITRE mapping** — ATT&CK tactic and technique IDs
- **Expected output** — what a legitimate vs. suspicious hit looks like
- **Recommended actions** — what to do when it fires
- **Source attribution** — every query / rule / IOC cites its upstream public source

Each artifact is a **single self-contained file** — all context lives inline as comments. No companion docs to chase.

---

## Structure

```
Threat-Hunting/
├── kql/                   44 KQL hunt queries — each a standalone .kql with inline
│                          metadata, background, tuning, hardening, and query body
├── iocs/                  5 IOC categories — one standalone .md per category listing
│                          atomic indicators with source + first-publish date
└── detection-rules/       15 productionized detections — Sentinel YAML (13) +
                           Defender XDR markdown (2); file prefix tells you the platform
```

---

## Counts at a glance

| Category | Count |
|---|---|
| KQL hunt queries | 44 |
| IOC categories (atomic + pattern indicators) | 5 |
| Sentinel analytic rules | 13 |
| Defender XDR custom detections | 2 |
| **Total artifacts** | **64** |

---

## KQL highlights

| Theme | File | Source incident / TTP |
|---|---|---|
| OAuth consent abuse | [Suspicious-OAuth-Consent.kql](./kql/Suspicious-OAuth-Consent.kql) | Vercel / Context.ai breach (Apr 2026) |
| Device code phishing | [Device-Code-Signin-Unmanaged-Device.kql](./kql/Device-Code-Signin-Unmanaged-Device.kql) | Storm-2372 (Feb 2025) |
| Infostealer C2 | [Lumma-Stealer-Indicators.kql](./kql/Lumma-Stealer-Indicators.kql) | Storm-2477 / Lumma |
| Credential persistence | [Credential-Added-To-App-Registration.kql](./kql/Credential-Added-To-App-Registration.kql) | Storm-0558 / Midnight Blizzard |
| Helpdesk phishing chain | [SSPR-Followed-By-Risky-Signin.kql](./kql/SSPR-Followed-By-Risky-Signin.kql) | Scattered Spider |
| Mailbox BEC | [Inbox-Forwarding-Rule-Created.kql](./kql/Inbox-Forwarding-Rule-Created.kql) | Classic BEC / Business Email Compromise |
| Golden SAML precursor | [Federation-Trust-Modified.kql](./kql/Federation-Trust-Modified.kql) | NOBELIUM |

See [kql/README.md](./kql/) for the full index grouped by OAuth / Identity / Data Access / Persistence / Endpoint / Exfiltration.

---

## IoCs

| File | Coverage | Lead source |
|---|---|---|
| [iocs/oauth-apps.md](./iocs/oauth-apps.md) | 5 malicious OAuth app IDs (Entra + Google) | Vercel/Context.ai; PerfectData / Mail_Backup |
| [iocs/infostealer-c2.md](./iocs/infostealer-c2.md) | 7 Lumma / AuraStealer / Vidar / StealC signatures | CISA AA25-141B, Microsoft, abuse.ch ThreatFox |
| [iocs/phishing-infra.md](./iocs/phishing-infra.md) | 4 Evilginx2/3 + Tycoon OAuth redirector patterns | LevelBlue, Deepwatch, Abnormal AI, Proofpoint |
| [iocs/threat-actor-ips.md](./iocs/threat-actor-ips.md) | 5 CISA advisories (Salt Typhoon, Interlock, Scattered Spider, LummaC2, Volt Typhoon) | CISA |
| [iocs/file-hashes.md](./iocs/file-hashes.md) | 4 SHA-256 hashes (StealC V2, Shuyal Stealer, Temp Stealer) | Picus, Point Wild, Cyble |

See [iocs/README.md](./iocs/) for usage notes (Sentinel watchlists, Defender XDR indicators, firewall/DNS lists, MISP/STIX feeds).

---

## Detection rules

Productionized forms of the highest-value KQL hunts:

- **Sentinel YAML** (13 rules) — deploy via portal, ARM, or GitHub Actions. File prefix: `sentinel-`
- **Defender XDR** (2 custom detections) — paste via Hunting → Custom detection rules. File prefix: `defender-xdr-`

See [detection-rules/README.md](./detection-rules/) for the full index + tuning table + deployment steps.

---

## How to use

### Microsoft Sentinel
1. Log Analytics in the Sentinel workspace
2. Copy the KQL body from any `.kql` file (skip the `//` header comments — they're reference)
3. Paste into a new query window, tune `ago(Nd)` to your retention, run

### Microsoft Defender XDR Advanced Hunting
1. [security.microsoft.com](https://security.microsoft.com) → Hunting → Advanced hunting
2. Most queries target `AuditLogs` or `SigninLogs` (Entra) and need Sentinel; Defender-native queries run directly
3. Paste, run, refine

### As a scheduled analytic rule
Most queries can be promoted to a scheduled analytic in Sentinel. The ones with tested, tuned forms already appear in [detection-rules/](./detection-rules/) as exportable YAML.

---

## Contributing

New queries land with this structure — all inline, no companion docs:

```kql
// Name:          <short name>
// Description:   <one-line purpose>
// MITRE ATT&CK:  T<id> (<technique name>)
// Data Source:   <tables / products required>
// Platform:      Sentinel | Defender XDR | Advanced Hunting | Log Analytics
// Time window:   <default lookback>
// Context:       <incident / TTP that motivated this>
//
// Recommended actions on hits:
// - ...
//
// ---------------------------------------------------------------------------
// Extended context, background, tuning, and attribution
// ---------------------------------------------------------------------------
//
// <full background, expected output, tuning, hardening, references,
//  and upstream source attribution — all as // comments>

<query body>
```

See [Suspicious-OAuth-Consent.kql](./kql/Suspicious-OAuth-Consent.kql) for the reference example.

---

## Disclaimer

Queries are provided as-is. Always validate in a non-production environment before use. False-positive rates vary by tenant baseline — tune thresholds to your environment. Griffin31 takes no responsibility for misinterpretation or misuse.

## License

MIT — see [LICENSE](../LICENSE). Upstream sources retain their own licenses (MIT, BSD-3-Clause) — attribution is in each file.
