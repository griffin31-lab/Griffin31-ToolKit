# File Hashes (SHA-256)

<sub>[← Back to IoCs](../) · [← Back to Threat-Hunting](../../)</sub>

SHA-256 hashes of publicly-documented malware samples — infostealers, loaders, droppers. Every hash below traces back to a published vendor analysis or public sandbox verdict. Hashes are durable (unlike domains/IPs), but polymorphism means any single sample is only a slice of a family — treat these as known-good seeds for VirusTotal pivots and Defender XDR retrospective hunts.

## Usage

| Platform | How to apply |
|---|---|
| **Microsoft Sentinel** | Watchlist `MaliciousFileHashes` → join on `DeviceFileEvents.SHA256`, `DeviceImageLoadEvents.SHA256`, `DeviceProcessEvents.SHA256` |
| **Defender XDR** | Indicators → File hashes → action `Block and remediate` |
| **VirusTotal / pivot tools** | Query each hash; pivot on Imphash, ssdeep, parent samples |
| **EDR (any)** | Upload to the "custom hash block" list on your EDR; include child-process detection for loader hashes |

## Indicators

| SHA-256 | Family / Variant | First Published | Source | Confidence | Notes |
|---|---|---|---|---|---|
| `c62e094cf89f9a2d3b5018fdd5ce30e664d40023b2ace19acc1fd7c6b2347143` | StealC V2 | 2025-03 | [Picus — StealC V2 analysis](https://www.picussecurity.com/resource/blog/stealc-v2-malware-enhances-stealth-and-expands-data-theft-features) | High | V2 release with JSON-based C2 protocol and RC4 transport |
| `B95F39B3C110D5FC7E89E50209C560FE7077B9B66A5FC31065F0C17C7F06EE83` | PowerShell loader (StealC V2 chain) | 2025 | [Medium — StealC V2 network analysis](https://medium.com/@pavol.kluka/network-traffic-analysis-how-to-analyze-stealc-version-2-infostealer-which-uses-rc4-e9f23d89aa06) | Medium | First-stage PowerShell loader observed pulling StealC V2 |
| `d5889aac10527ddc7d4b03407a8933a84a1ea0550f61d442493d4f3237203e3c` | Temp Stealer (Bundled installer) | 2025 | [Cyble — Bundled installer infostealer](https://cyble.com/blog/infostealer-distributed-using-bundled-installer/) | Medium | Dropped by a trojanized installer; credential and cookie theft |
| `8bbeafcc91a43936ae8a91de31795842cd93d2d8be3f72ce5c6ed27a08cdc092` | Shuyal Stealer | 2025 | [Point Wild — Shuyal Stealer analysis](https://www.pointwild.com/threat-intelligence/shuyal-stealer-advanced-infostealer-targeting-19-browsers/) | High | Targets 19 browsers; documented in Point Wild threat report |

## Source & attribution

- [Picus Security — StealC v2 Malware Enhances Stealth and Expands Data Theft Features](https://www.picussecurity.com/resource/blog/stealc-v2-malware-enhances-stealth-and-expands-data-theft-features)
- [Medium / Pavol Kluka — Network Traffic Analysis: StealC v2 with RC4](https://medium.com/@pavol.kluka/network-traffic-analysis-how-to-analyze-stealc-version-2-infostealer-which-uses-rc4-e9f23d89aa06)
- [Cyble — Infostealer distributed using bundled installer](https://cyble.com/blog/infostealer-distributed-using-bundled-installer/)
- [Point Wild — Shuyal Stealer: Advanced Infostealer Targeting 19 Browsers](https://www.pointwild.com/threat-intelligence/shuyal-stealer-advanced-infostealer-targeting-19-browsers/)
- [CISA AA25-141B — LummaC2 advisory (additional LummaC2 hashes in Tables 6-7)](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b)

## Notes

- Hashes are **single samples**, not family signatures. A retrospective hunt with these 4 hashes will find only the specific binaries; use the family names to pivot into YARA rules.
- For Lumma-specific hashes, prefer pulling directly from CISA AA25-141B Tables 6-7 (the list is too long to reproduce and is updated by the advisory itself).
- Defender XDR automatically ingests many of these via Microsoft Threat Intelligence; verify the indicator is present in your tenant's TI feed before duplicating.
