# Infostealer Indicator Match on Endpoint

<sub>[← Back to Detection Rules](../) · [← Back to Threat-Hunting](../../)</sub>

| Field | Value |
|---|---|
| **Rule ID (internal)** | `d82a5f46-3e71-4c89-bf05-7a6c9d1b4e38` |
| **Platform** | Microsoft Defender XDR (custom detection) |
| **Severity** | High |
| **Frequency** | Every 30 minutes |
| **MITRE ATT&CK Tactic** | Credential Access, Exfiltration |
| **MITRE ATT&CK Technique** | [T1555 — Credentials from Password Stores](https://attack.mitre.org/techniques/T1555/), [T1567 — Exfiltration Over Web Service](https://attack.mitre.org/techniques/T1567/) |
| **Data sources** | `DeviceNetworkEvents`, `DeviceFileEvents`, `DeviceProcessEvents` |

## Description

Matches endpoint telemetry against documented infostealer IOCs from the Griffin31 library:
- File hashes (StealC V2, Shuyal Stealer, Temp Stealer)
- Lumma / Storm-2477 C2 URI pattern (`/c2sock` POST)
- StealC V2 RC4-encrypted JSON exfil pattern
- AuraStealer `.shop` / `.cfd` C2 TLDs (pattern match)

Cross-references: [iocs/infostealer-c2](../../iocs/infostealer-c2/) and [iocs/file-hashes](../../iocs/file-hashes/).

## KQL

```kql
let LummaURIPattern = "/c2sock";
let KnownInfostealerHashes = dynamic([
  "c62e094cf89f9a2d3b5018fdd5ce30e664d40023b2ace19acc1fd7c6b2347143",
  "B95F39B3C110D5FC7E89E50209C560FE7077B9B66A5FC31065F0C17C7F06EE83",
  "d5889aac10527ddc7d4b03407a8933a84a1ea0550f61d442493d4f3237203e3c",
  "8bbeafcc91a43936ae8a91de31795842cd93d2d8be3f72ce5c6ed27a08cdc092"
]);
let SuspiciousTLDs = dynamic([".shop", ".cfd", ".xyz", ".top", ".duckdns.org", ".ddns.net"]);
let FileHits =
    DeviceFileEvents
    | where Timestamp > ago(30m)
    | where SHA256 in~ (KnownInfostealerHashes) or InitiatingProcessSHA256 in~ (KnownInfostealerHashes)
    | extend MatchReason = "known_infostealer_sha256", MatchValue = SHA256
    | project Timestamp, DeviceId, DeviceName, InitiatingProcessAccountName, FileName, SHA256, MatchReason, MatchValue;
let NetworkHits =
    DeviceNetworkEvents
    | where Timestamp > ago(30m)
    | where RemoteUrl has LummaURIPattern
       or RemoteUrl has_any (SuspiciousTLDs)
    | extend MatchReason = iff(RemoteUrl has LummaURIPattern, "lumma_c2sock_uri", "suspicious_infostealer_tld")
    | extend MatchValue = RemoteUrl
    | project Timestamp, DeviceId, DeviceName, InitiatingProcessAccountName, InitiatingProcessFileName, RemoteUrl, RemoteIP, MatchReason, MatchValue;
union FileHits, NetworkHits
| order by Timestamp desc
```

## Entity mapping

| Defender Entity | Column |
|---|---|
| Device | `DeviceId`, `DeviceName` |
| User | `InitiatingProcessAccountName` |
| File | `SHA256`, `FileName` |
| URL | `RemoteUrl` |
| IP | `RemoteIP` |

## Suggested response actions

1. **Automated** — Isolate device; collect investigation package; submit file to Microsoft for deep analysis
2. **Immediate** — Rotate browser-saved credentials, cookies, MFA tokens for the affected user; revoke Entra sessions (`Revoke-AzureADUserAllRefreshToken`)
3. **Wider hunt** — Take the matched SHA256 and pivot across the whole estate (`DeviceFileEvents | where SHA256 == "<hash>"`)
4. **Block** — Add the matched URL / domain to Defender XDR Indicators with action `Block and remediate`
5. **User awareness** — Notify user; guide them through password reset, device reimage, and token revocation

## Source & attribution

- [Microsoft Security Blog — Lumma Stealer](https://www.microsoft.com/en-us/security/blog/2025/05/21/lumma-stealer-breaking-down-the-delivery-techniques-and-capabilities-of-a-prolific-infostealer/)
- [CISA AA25-141B — LummaC2 advisory](https://www.cisa.gov/news-events/cybersecurity-advisories/aa25-141b)
- [Picus Security — StealC v2 analysis](https://www.picussecurity.com/resource/blog/stealc-v2-malware-enhances-stealth-and-expands-data-theft-features)
- [Point Wild — Shuyal Stealer analysis](https://www.pointwild.com/threat-intelligence/shuyal-stealer-advanced-infostealer-targeting-19-browsers/)
- [abuse.ch URLhaus infostealer tag](https://urlhaus.abuse.ch/browse/tag/infostealer/)

## Tuning notes

- Hash-based matching is high-fidelity. TLD-pattern matching has a meaningful false-positive rate on legitimate `.shop` / `.xyz` domains — combine with a newly-observed-domain or low-reputation signal before auto-response.
- Treat the hash list as a *seed*; pull the full CISA AA25-141B IOC set for the richest coverage, and refresh weekly.
- If you are running MDE with MDI TI, many of these hashes will already block at the agent; this rule acts as a safety net.
