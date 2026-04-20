<sub>[← Back to Griffin31 ToolKit](../)</sub>

# CA Emergency Account Exclusion — Logic App (Consumption)

## Overview

Azure **Consumption Logic App** that continuously enforces the exclusion of two emergency/break-glass accounts from **every enabled** Conditional Access policy in the tenant. Runs every 30 minutes and idempotently patches any enabled policy missing the exclusion. Authenticates to Microsoft Graph via the Logic App's system-assigned Managed Identity.

## Why this tool

A single misconfigured Conditional Access policy can lock every admin out of the tenant. Break-glass accounts are meant to be excluded from all CA policies — but exclusions get removed during policy edits, policy cloning, drift, and accidental overwrites. This tool re-applies the exclusion every 30 minutes so the break-glass path is never silently closed.

## Requirements

- **Entra ID Premium P1 or P2** (Conditional Access)
- **Azure Subscription**: Contributor (on the target resource group)
- **Entra ID**: Privileged Role Administrator or Global Administrator (to grant the app role to the MSI)
- PowerShell 7+ with: `Az.Accounts`, `Az.Resources`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (auto-installed by the deploy script)

## Quick start

```powershell
pwsh ./scripts/Deploy-BreakGlassEnforcer.ps1
```

Interactive. Prompts for resource group, Logic App name, region, two emergency account Object IDs, and recurrence interval.

The deploy script:
1. Validates both Object IDs resolve to **enabled** users via Graph
2. Rejects duplicate Object IDs
3. Creates the resource group if missing
4. Validates and deploys the ARM template
5. Grants `Policy.ReadWrite.ConditionalAccess` to the new Managed Identity

## Architecture

```
Recurrence (every N min)
   |
   v
GET /v1.0/identity/conditionalAccess/policies   (MSI)
   |
   v
For each policy where state == 'enabled':
   - Both accounts already excluded? -> skip
   - Otherwise:
       GET /v1.0/identity/conditionalAccess/policies/{id}   (capture @odata.etag)
       Build excludeUsers = coalesce(current, []) + [acct1, acct2]
       PATCH /v1.0/identity/conditionalAccess/policies/{id}
           Headers: If-Match: <etag>
       Log result (ok / failed with status + body) to RunSummary
```

```
CA-BreakGlass-Enforcer/
|-- README.md
|-- .gitignore
|-- infrastructure/
|   `-- CA-BreakGlass-Enforcer.json   # ARM template (workflow embedded)
|-- scripts/
|   |-- Deploy-BreakGlassEnforcer.ps1
|   |-- Grant-GraphPermissions.ps1
|   |-- Remove-BreakGlassEnforcer.ps1
|   `-- Test-BreakGlassEnforcer.ps1
`-- docs/
    |-- SETUP-GUIDE.md
    `-- TROUBLESHOOTING.md
```

## Required Graph permission

| Permission | Type | Reason |
|---|---|---|
| `Policy.ReadWrite.ConditionalAccess` | Application | Read and patch CA policies (read is covered by ReadWrite) |

## Safety

This Logic App **modifies Conditional Access policies tenant-wide, every 30 minutes**. Treat it as a production control, not a toy:

- Put it in a **dedicated resource group** with RBAC locked down
- Grant **Contributor** on that resource group only via **PIM** (time-bound, approval-gated)
- Restrict who can delete the Logic App — an attacker who can delete it can then freely remove break-glass exclusions
- Emit Logic App run history to **Log Analytics** so every patch is auditable

### Enable diagnostic export to Log Analytics

```bash
az monitor diagnostic-settings create \
  --name 'ca-emergency-diag' \
  --resource '<logic-app-resource-id>' \
  --workspace '<log-analytics-workspace-resource-id>' \
  --logs '[{"category":"WorkflowRuntime","enabled":true}]'
```

The workflow writes per-policy result lines (`OK policyId=... displayName=... patchResult=ok` or `FAIL ... statusCode=... patchResult=failed`) to a run-summary string that is visible in run history and flows to Log Analytics when diagnostic settings are enabled.

## Security notes

- **Disabled / report-only policies are skipped** — the workflow only patches policies with `state == 'enabled'`
- **If-Match header** on every PATCH prevents last-writer-wins during concurrent admin edits
- **Null-safe** — the `excludeUsers` array is coalesced to `[]` before contains/append, so policies without any current exclusions are handled correctly
- **MSI-only authentication** — no stored credentials, no client secret, no certificate
- **v1.0 Graph endpoints** only (no preview surface)
- **Least-privilege** — only `Policy.ReadWrite.ConditionalAccess`

## Removal

```powershell
# Remove the Logic App only (policies still contain the exclusions)
pwsh ./scripts/Remove-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation"

# Remove the Logic App AND strip the exclusions from every policy
pwsh ./scripts/Remove-BreakGlassEnforcer.ps1 `
    -ResourceGroupName "rg-ca-automation" `
    -RestorePolicies `
    -EmergencyAccount1ObjectId "<id1>" `
    -EmergencyAccount2ObjectId "<id2>"

# Remove the entire resource group
pwsh ./scripts/Remove-BreakGlassEnforcer.ps1 `
    -ResourceGroupName "rg-ca-automation" -RemoveResourceGroup
```

Without `-RestorePolicies` the script prints a warning that emergency account Object IDs remain in every CA policy's excludeUsers. That's often the desired behavior (you want the break-glass exclusions to persist even after the Logic App is retired) — but if the emergency accounts have been deleted, re-run with `-RestorePolicies` to clean them out.

## Troubleshooting

See [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

## Setup details

See [docs/SETUP-GUIDE.md](docs/SETUP-GUIDE.md).

## License

See [../LICENSE](../LICENSE).

## Related resources

- [Emergency access accounts in Microsoft Entra ID](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access)
- [Conditional Access API](https://learn.microsoft.com/en-us/graph/api/resources/conditionalaccesspolicy)
- [Logic Apps Consumption pricing](https://azure.microsoft.com/en-us/pricing/details/logic-apps/)
