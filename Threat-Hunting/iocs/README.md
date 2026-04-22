# Indicators of Compromise (IoCs)

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

Atomic indicators — OAuth app IDs, file hashes, domains, IPs, URLs — tied to documented incidents or active campaigns.

## Structure (planned)

```
iocs/
├── oauth-apps/       Malicious OAuth App IDs (Entra, Google Workspace)
├── hashes/           SHA-256 file hashes (Defender XDR searchable)
├── domains/          C2 / phishing / exfil domains
├── ips/              C2 / scanner / exfil IPs
└── urls/             Full malicious URLs
```

Each indicator file carries:
- `source` — where it came from (vendor disclosure, public report, internal detection)
- `published` — date the indicator was first published
- `threat_actor` — attribution if known (e.g., Storm-2477 / Lumma Stealer)
- `confidence` — High / Medium / Low
- `expires` — when to stop acting on the indicator (many IoCs are ephemeral)

## Index

_Empty — indicators will be added as new incidents are documented._

## Notes

- IoCs here are published as plain text for copy/paste into SIEM watchlists, Defender XDR indicators, and firewall deny-lists.
- Always confirm an indicator is still active before acting on it.
- Many infostealer C2 domains rotate weekly — prefer TTP-based hunts over IoC blocks.
