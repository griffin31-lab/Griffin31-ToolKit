# Infostealer C2 Infrastructure

<sub>[← Back to IoCs](../) · [← Back to Threat-Hunting](../../)</sub>

Command-and-control domains, URL path patterns, and IP endpoints tied to publicly-documented infostealer campaigns (Lumma / Storm-2477, RedLine, Vidar, StealC, AuraStealer). All entries reference the FBI/CISA, Microsoft, or abuse.ch disclosure they were drawn from.

## Usage

| Platform | How to apply |
|---|---|
| **Microsoft Sentinel** | Import as `InfostealerC2` watchlist → reference from DNS / proxy / firewall tables (`CommonSecurityLog`, `DeviceNetworkEvents`) |
| **Defender XDR** | Indicators → URLs/Domains → import list with action `Block and remediate` |
| **Firewall / DNS filter** | Add to deny list; expect rapid rotation on infostealer families (7-day TTL recommended) |
| **Threat intel platform (MISP, OpenCTI)** | Tag with the family (`lumma`, `redline`, `vidar`, `stealc`, `aurastealer`) and source advisory |

## Indicators

| Indicator | Type | Family / Actor | First Published | Confidence | Source | Notes |
|---|---|---|---|---|---|---|
| `/c2sock` | URI path | Lumma / Storm-2477 | 2025-05-21 | High | [Microsoft Security Blog — Lumma](https://www.microsoft.com/en-us/security/blog/2025/05/21/lumma-stealer-breaking-down-the-delivery-techniques-and-capabilities-of-a-prolific-infostealer/) · [CISA AA25-141B](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b) | Stolen data is posted via HTTP POST `multipart/form-data` to this path — hunt on outbound HTTP logs |
| `*.shop` (unusually low-rep) | TLD pattern | AuraStealer | 2025-07 | Medium | [abuse.ch URLhaus](https://urlhaus.abuse.ch/browse/tag/infostealer/) · [Cybersecurity News — AuraStealer](https://cybersecuritynews.com/threat-actors-deploy-aurastealer-infostealer/) | 48 C2 domains observed since mid-2025 on `.shop` and `.cfd` |
| `*.cfd` (unusually low-rep) | TLD pattern | AuraStealer | 2025-07 | Medium | [abuse.ch URLhaus](https://urlhaus.abuse.ch/browse/tag/infostealer/) · [Cybersecurity News — AuraStealer](https://cybersecuritynews.com/threat-actors-deploy-aurastealer-infostealer/) | Co-occurring TLD for AuraStealer C2 registration |
| `95.213.224.25:80` | IP:port | RedLine Stealer | 2023-01 (first documented); still active 2025 | Medium | [ThreatFox — RedLine](https://threatfox.abuse.ch/browse/malware/win.redline_stealer/) · [CircleID RedLine analysis](https://circleid.com/posts/20230104-redline-stealer-ioc-analysis-and-expansion) | Historically-documented C2; validate on abuse.ch before blocking |
| `*.xyz`, `*.top`, `*.duckdns.org`, `*.ddns.net` (dynamic DNS + low-rep gTLDs) | Domain pattern | RedLine / commodity infostealers | 2023-01 onward | Medium | [CircleID RedLine analysis](https://circleid.com/posts/20230104-redline-stealer-ioc-analysis-and-expansion) · [abuse.ch URLhaus](https://urlhaus.abuse.ch/) | Meta-IOC: statistically over-represented in commodity infostealer C2; combine with new-domain + low-reputation signals |
| Telegram + Mastodon "resolver" accounts | Protocol / infra | Vidar Stealer 2.0 | 2025-10-06 | High | [Trend Micro — Vidar 2.0](https://www.trendmicro.com/en_us/research/25/j/how-vidar-stealer-2-upgrades-infostealer-capabilities.html) · [SC Media — Vidar 2.0](https://www.scworld.com/news/vidar-stealer-2-0-what-to-know-about-new-infostealer-features) | Vidar binaries fetch C2 IP from attacker-registered social-media profiles (dead drop resolver). Hunt DNS to t.me/mastodon.social followed by traffic to newly-seen IP |
| HTTP POST with RC4-encrypted JSON body | Protocol signature | StealC v2 | 2025-03 | High | [Medium — StealC v2 network analysis](https://medium.com/@pavol.kluka/network-traffic-analysis-how-to-analyze-stealc-version-2-infostealer-which-uses-rc4-e9f23d89aa06) · [Picus Security — StealC v2](https://www.picussecurity.com/resource/blog/stealc-v2-malware-enhances-stealth-and-expands-data-theft-features) | StealC v2 switched to JSON-based C2 protocol; hunt for RC4-looking payloads to uncommon hosts |

## Source & attribution

- [CISA / FBI AA25-141B — LummaC2 advisory](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b)
- [Microsoft Security Blog — Lumma Stealer delivery techniques](https://www.microsoft.com/en-us/security/blog/2025/05/21/lumma-stealer-breaking-down-the-delivery-techniques-and-capabilities-of-a-prolific-infostealer/)
- [Trend Micro — Vidar Stealer 2.0 analysis](https://www.trendmicro.com/en_us/research/25/j/how-vidar-stealer-2-upgrades-infostealer-capabilities.html)
- [abuse.ch ThreatFox — RedLine Stealer IOCs](https://threatfox.abuse.ch/browse/malware/win.redline_stealer/)
- [abuse.ch URLhaus — infostealer-tagged URLs](https://urlhaus.abuse.ch/browse/tag/infostealer/)
- [Picus Security — StealC v2 behavior](https://www.picussecurity.com/resource/blog/stealc-v2-malware-enhances-stealth-and-expands-data-theft-features)
- [Cybersecurity News — AuraStealer 48 C2 domains](https://cybersecuritynews.com/threat-actors-deploy-aurastealer-infostealer/)

## Notes

- Infostealer C2 domains **rotate every 4-7 days**. Prefer TTP-based hunts (URI path, User-Agent, payload shape) over atomic domain blocks.
- The specific domain lists in CISA AA25-141B are historical (Nov 2023 – May 2025) — Microsoft seized 2,300 in May 2025. Treat any domain list older than 30 days as retrospective hunt material, not preventive block.
- For the live, up-to-date block list, pull from abuse.ch ThreatFox API daily.
