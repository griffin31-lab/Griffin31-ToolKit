# Indicators of Compromise (IoCs)

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

Atomic and pattern-based indicators tied to documented incidents, published advisories, or tracked campaigns. Each file below is a standalone category listing with full source attribution inline.

## Index

| File | Coverage | Lead source |
|---|---|---|
| [oauth-apps.md](./oauth-apps.md) | Malicious OAuth application IDs (Entra + Google) — 5 entries | Microsoft Security Blog, Push Security, Mitiga, Proofpoint, Darktrace |
| [infostealer-c2.md](./infostealer-c2.md) | Infostealer C2 domains, URL paths, protocol signatures — 7 entries | CISA AA25-141B, Microsoft, Trend Micro, abuse.ch ThreatFox |
| [phishing-infra.md](./phishing-infra.md) | Phishing kit infrastructure patterns — 4 entries | LevelBlue, Deepwatch, Abnormal AI, Proofpoint |
| [threat-actor-ips.md](./threat-actor-ips.md) | State-sponsored / criminal IP infrastructure — 5 CISA advisories | CISA AA25-239A / AA25-203A / AA25-141B / AA23-320A / AA23-144A |
| [file-hashes.md](./file-hashes.md) | SHA-256 hashes of documented malware samples — 4 entries | Picus, Point Wild, Cyble |

## Cross-cutting usage

- **Sentinel** — most categories import cleanly as Watchlists. Sample KQL joins are in each file.
- **Defender XDR** — paste atomic values into Settings → Endpoints → Indicators (by type).
- **Firewall / DNS filter** — prefer `threat-actor-ips.md` and `infostealer-c2.md`. Expire aggressively (30-90 days max) — infostealer C2 rotates weekly.
- **MISP / OpenCTI / STIX feeds** — CISA and abuse.ch both publish machine-readable feeds; each file links to the canonical feed.

## Notes

- **Every indicator has a source URL and first-publish date.** Griffin31 does not fabricate, guess, or speculate.
- **Expiry matters.** Infostealer C2s rotate weekly. State-actor IPs are revoked by CISA on a rolling basis. Re-verify at the source before acting.
- **Prefer TTP-based hunts** when the indicator lifespan is short. Atomic IoCs in this library are best for *retrospective* detection — "was I ever touched by this?" — not for long-term block lists.
