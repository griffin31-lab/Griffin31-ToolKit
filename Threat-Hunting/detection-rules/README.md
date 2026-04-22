# Detection Rules

<sub>[← Back to Threat-Hunting](../) · [← Back to Griffin31 ToolKit](../../)</sub>

Scheduled analytic rules in native export format for Microsoft Sentinel, Defender XDR (Custom Detections), and related platforms.

## Structure (planned)

```
detection-rules/
├── sentinel/         Sentinel analytic rules (ARM template or YAML)
├── defender-xdr/     Defender XDR custom detections
└── purview/          Microsoft Purview alert policies
```

Each rule carries:
- **Severity** and **Tactic** (MITRE ATT&CK)
- **Data sources** required
- **Entity mappings** (user, host, IP, app)
- **Tuning guidance** — thresholds, exclusions, expected false-positive rate
- **Companion KQL** — many rules here are productionized versions of hunt queries in [../kql/](../kql/)

## Index

_Empty — rules will be added once the corresponding hunt queries are tested and tuned._

## Deployment notes

- **Sentinel ARM templates** are deployed via `New-AzResourceGroupDeployment` or the Azure portal's "Import from template" flow.
- **Defender XDR custom detections** are created manually via [security.microsoft.com](https://security.microsoft.com) → Hunting → Custom detection rules. The companion KQL from `../kql/` can be copy-pasted.
- Always test in a non-production workspace first. Tune thresholds to match your tenant's baseline.
