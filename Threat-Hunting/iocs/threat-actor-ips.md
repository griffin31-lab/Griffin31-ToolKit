# Threat Actor IP Infrastructure

<sub>[← Back to IoCs](../) · [← Back to Threat-Hunting](../../)</sub>

IP ranges and individual addresses published in CISA, NSA, FBI, and partner joint advisories, attributable to named threat actors (Salt Typhoon, Volt Typhoon, Interlock, Scattered Spider, LummaC2 operators). Treat historical indicators as **retrospective hunt material** — most of these are rotated out by the time they're published.

## Usage

| Platform | How to apply |
|---|---|
| **Microsoft Sentinel** | Watchlist `ThreatActorIPs` → join with `SigninLogs`, `DeviceNetworkEvents`, `CommonSecurityLog` on IP; retrospective hunt across 90-365 days |
| **Defender XDR** | Indicators → IPs → action `Alert only` for historical hunting (not block, due to FP risk on aged IOCs) |
| **Firewall / NDR** | Add to alerting list. CISA specifically warns against blind blocking of listed IPs — some may be legitimate infrastructure |
| **Purview / Sentinel TI** | Ingest the CISA STIX XML/JSON feeds directly (auto-expire on publisher pull) |

## Indicators

Specific IPs from these advisories deliberately kept as *references* rather than re-published — the CISA STIX feeds are the canonical source of truth and are updated/revoked as actors rotate.

| Advisory | Actor(s) | Published | IOC Source | Context | Confidence |
|---|---|---|---|---|---|
| [CISA AA25-239A](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-239a) | Salt Typhoon / OPERATOR PANDA / RedMike / UNC5807 / GhostEmperor | 2025-08-27 | [STIX XML](https://www.cisa.gov/sites/default/files/2025-09/AA25-239A_Countering_Chinese_State-Sponsored_Actors_Compromise_of_Networks_Worldwide_to_Feed_Global_Espionage_System.stix_.xml) | PRC state-sponsored APT; backbone router compromise; telecom / government / transportation / military targets; activity from 2021 – June 2025 | High (for activity); Low-Medium (for IP currency — updated Sep 2025 with removals) |
| [CISA AA25-203A](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-203a) | Interlock ransomware | 2025-07-22 | STIX XML + JSON on advisory page | Double-extortion ransomware; critical infrastructure in North America + Europe; ClickFix social engineering delivery; activity since Sept 2024 | High |
| [CISA AA25-141B](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b) | LummaC2 / Storm-2477 | 2025-05-21 | Tables 6 + 7 in PDF advisory | MaaS infostealer; infections from Nov 2023 – May 2025; 2,300 C2 domains seized by Microsoft DCU May 13, 2025 | Medium (historical — infra seized or rotated) |
| [CISA AA23-320A](https://www.cisa.gov/news-events/cybersecurity-advisories/aa23-320a) (updated 2025-07-29) | Scattered Spider / UNC3944 / Octo Tempest / Storm-0875 | 2023-11; updated 2025-07-29 | Advisory body | Social-engineering-first criminal group; 600+ impersonation domains tracked by Silent Push Q1 2022 – Q1 2025; 81% impersonate tech vendors; registrar of choice NiceNIC; DragonForce ransomware variant used 2025 | High (TTPs); Medium (atomic IOCs) |
| [CISA AA23-144A / AA24-038A](https://www.cisa.gov/news-events/cybersecurity-advisories/aa23-144a) | Volt Typhoon / Insidious Taurus | 2023-05; updated 2024-02 | STIX JSON (`AA23-144A.stix.json`) | PRC state-sponsored; living-off-the-land; 5-year persistence observed; critical infrastructure | Medium (actor uses minimal atomic IOCs by design; behavioral hunts preferred) |

## Source & attribution

- [CISA Cybersecurity Advisories index](https://www.cisa.gov/news-events/cybersecurity-advisories)
- [CISA AA25-239A — Countering Chinese State-Sponsored Actors (Salt Typhoon et al.)](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-239a)
- [CISA AA25-203A — #StopRansomware: Interlock](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-203a)
- [CISA AA25-141B — LummaC2](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b)
- [CISA AA23-320A — Scattered Spider](https://www.cisa.gov/news-events/cybersecurity-advisories/aa23-320a)
- [CISA AA23-144A — PRC actor living off the land (Volt Typhoon)](https://www.cisa.gov/news-events/cybersecurity-advisories/aa23-144a)

## Notes

- **CISA's own guidance** is that "some IP addresses in the advisory may be associated with legitimate activity" — always investigate before blocking, especially for IPs older than 90 days.
- IPs here are **pointers to the STIX feeds**, not re-published IP strings. The authoritative, version-controlled list lives on CISA's site — and pulling from STIX keeps you in sync with additions/removals.
- For Volt Typhoon specifically, atomic IOCs are sparse by design. The value is in the behavioral hunt (wmic, ntdsutil, netsh, PowerShell patterns), not the IP list.
