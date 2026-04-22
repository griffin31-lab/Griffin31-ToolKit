# LSASS Memory Access from Non-Microsoft Binary

<sub>[← Back to Detection Rules](../) · [← Back to Threat-Hunting](../../)</sub>

| Field | Value |
|---|---|
| **Rule ID (internal)** | `f7b29c40-5d14-4a82-9e36-8c1d3b0e7a21` |
| **Platform** | Microsoft Defender XDR (custom detection) |
| **Severity** | High |
| **Frequency** | Every 1 hour |
| **MITRE ATT&CK Tactic** | Credential Access |
| **MITRE ATT&CK Technique** | [T1003.001 — OS Credential Dumping: LSASS Memory](https://attack.mitre.org/techniques/T1003/001/) |
| **Data sources** | `DeviceEvents` (action: `OpenProcessApiCall`), `DeviceProcessEvents` |

## Description

Detects processes that open a handle to `lsass.exe` with memory-read rights from a non-Microsoft-signed initiating binary. LSASS holds plaintext credentials and Kerberos tickets — opening a handle for read access is the first half of every credential-dumping tool (Mimikatz, ProcDump misuse, SharpKatz, NanoDump).

## KQL

```kql
DeviceEvents
| where Timestamp > ago(1h)
| where ActionType == "OpenProcessApiCall"
| where FileName =~ "lsass.exe"
| extend InitiatingProcess = tostring(InitiatingProcessFileName)
| where InitiatingProcessFileName !in~ ("MsMpEng.exe", "svchost.exe", "wininit.exe", "csrss.exe", "SenseIR.exe", "MsSense.exe", "System", "taskmgr.exe")
| where InitiatingProcessSignatureStatus != "Valid" or InitiatingProcessSignerType != "Microsoft"
| project Timestamp, DeviceName, DeviceId, InitiatingProcessAccountName, InitiatingProcessFileName, InitiatingProcessCommandLine, InitiatingProcessSignerType, InitiatingProcessSignatureStatus, ReportId
```

## Entity mapping

| Defender Entity | Column |
|---|---|
| Device | `DeviceId`, `DeviceName` |
| User | `InitiatingProcessAccountName` |
| File | `InitiatingProcessFileName` |
| Process | `InitiatingProcessCommandLine` |

## Suggested response actions

1. **Automated** — Isolate the device (network quarantine) via Defender Automated Investigation & Remediation
2. **Manual** — Collect investigation package; rotate any passwords of accounts that logged into the device in the previous 30 days
3. **Prevent** — Enable ASR rule `Block credential stealing from the Windows local security authority subsystem (lsass.exe)` (GUID `9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2`)
4. **Hunt** — Pivot on `InitiatingProcessSHA256` and check for the same hash on other endpoints

## Source & attribution

- [MITRE ATT&CK T1003.001 — LSASS Memory dumping](https://attack.mitre.org/techniques/T1003/001/)
- [Azure/Azure-Sentinel: LSASS Memory Access Detections](https://github.com/Azure/Azure-Sentinel/tree/master/Detections)
- [Microsoft Learn — ASR rules reference](https://learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction-rules-reference)

## Tuning notes

- Legitimate noise sources: `MsMpEng.exe` (Defender), `SenseIR.exe` / `MsSense.exe` (MDE agent), vendor EDR agents, Process Explorer when run by admins. The allow-list above is a starting point — extend for your security tool estate.
- False-positive rate on a cleanly-configured estate is typically 1-3 hits/day from admin tooling; in an EDR-heavy environment this can climb. Tune before alerting.
