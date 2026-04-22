# Threat-Hunting

<sub>[← Back to Griffin31 ToolKit](../) · Cross-platform · Microsoft Sentinel · Defender XDR · Advanced Hunting</sub>

> **Curated KQL hunt queries, IoCs, and detection rules** — built for Microsoft 365, Entra ID, Microsoft Defender XDR, and Microsoft Sentinel.

Every item in this library is tied to a documented real-world incident or a concrete attacker technique, mapped to MITRE ATT&CK, and shipped with:

- **Why it matters** — the incident or TTP that motivated the hunt
- **Data source** — the log tables / products required to run it
- **MITRE mapping** — ATT&CK tactic and technique IDs
- **Expected output** — what a legitimate vs. suspicious hit looks like
- **Recommended actions** — what to do when it fires

---

## Structure

```
Threat-Hunting/
├── kql/                  KQL hunt queries (Sentinel, Defender XDR, Advanced Hunting, Log Analytics)
│   ├── oauth/            OAuth consent abuse, app registration abuse, delegated permission misuse
│   ├── (more categories as library grows)
│   └── README.md
├── iocs/                 Indicators of Compromise tied to documented incidents
│   └── README.md
└── detection-rules/      Defender XDR / Sentinel analytic rules in native export format
    └── README.md
```

---

## Quick index

| Category | Query / Rule | Source incident / TTP |
|---|---|---|
| OAuth | [Suspicious-OAuth-Consent.kql](./kql/oauth/Suspicious-OAuth-Consent.kql) | Vercel / Context.ai breach (Apr 2026) |

New entries land here with each release. Check back regularly or watch the repo.

---

## How to use

### In Microsoft Sentinel
1. Open **Log Analytics** in the Sentinel workspace
2. Copy the KQL from the `.kql` file
3. Paste into a new query window, adjust `TimeGenerated > ago(Nd)` as needed, run

### In Microsoft Defender XDR Advanced Hunting
1. Open [security.microsoft.com](https://security.microsoft.com) → **Hunting** → **Advanced hunting**
2. Some queries target `AuditLogs` (Entra) and will need to be run from the Sentinel side; Defender-native queries will be in the `defender/` subfolder when added
3. Paste, run, refine

### As a scheduled analytic rule
Most queries can be promoted to a scheduled analytic in Sentinel. When a query has a tested, tuned form ready for rule deployment, it will also appear in `detection-rules/` as an exported ARM / YAML template.

---

## Contributing

Each new query should include:
- A `.kql` file with a metadata header block (Name, Description, MITRE, Data Source, Reference)
- A companion `.md` file in the same folder documenting background, expected output, and remediation
- Tested against a real Sentinel / Defender tenant before submission
- Tuning notes if the query is prone to false positives

See the first entry ([OAuth consent hunt](./kql/oauth/Suspicious-OAuth-Consent.md)) for the expected format.

---

## Disclaimer

Queries are provided as-is. Always validate in a non-production environment before use. False-positive rates vary by tenant baseline — tune thresholds to match your environment. Griffin31 takes no responsibility for misinterpretation or misuse of results.

## License

MIT — see [LICENSE](../LICENSE).
