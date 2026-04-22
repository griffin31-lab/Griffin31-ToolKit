# KQL Hunt Queries

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

KQL (Kusto Query Language) queries for Microsoft Sentinel, Defender XDR Advanced Hunting, and Azure Monitor Log Analytics.

## Index

### OAuth / Application Consent
- [Suspicious-OAuth-Consent.kql](./oauth/Suspicious-OAuth-Consent.kql) — Detects suspicious OAuth consent grants and delegated permission assignments. **Context:** Vercel / Context.ai breach (April 2026). [Background](./oauth/Suspicious-OAuth-Consent.md)

_More categories will be added as the library grows._

## File conventions

Every query file follows this pattern:

```kql
// Name: <short name>
// Description: <one-line purpose>
// MITRE ATT&CK: T<id> (<technique name>), ...
// Data Source: <table(s) required>
// Platform: <Sentinel | Defender XDR | Advanced Hunting | Log Analytics>
// Time window: <default lookback>
// Reference: <link to companion .md>
// Context: <what incident / TTP motivated this>
//
// Recommended actions on hits:
// - <action 1>
// - <action 2>

<the query>
```

Each `.kql` has a sibling `.md` documenting background, expected output, tuning notes, and remediation steps.
