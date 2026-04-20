# Setup Guide — CA Emergency Account Exclusion Logic App

This tool ships as one ARM template and four PowerShell scripts. There is no Bicep, no Functions Core Tools, no `config/local.settings.json`, and no separate workflow file — the workflow definition is embedded inside the ARM template.

## Table of contents

1. [Prerequisites](#prerequisites)
2. [Find the emergency account Object IDs](#find-the-emergency-account-object-ids)
3. [Deploy (interactive, recommended)](#deploy-interactive-recommended)
4. [Deploy (manual, ARM only)](#deploy-manual-arm-only)
5. [Post-deployment validation](#post-deployment-validation)
6. [Maintenance](#maintenance)
7. [Diagnostic export](#diagnostic-export)

---

## Prerequisites

| Scope | Required role |
|---|---|
| Azure subscription | Contributor on the target resource group |
| Entra ID | Privileged Role Administrator or Global Administrator (to grant the Graph app role to the MSI) |

PowerShell modules (auto-installed by the deploy script if missing):

```powershell
Install-Module Az.Accounts, Az.Resources, Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Force
```

---

## Find the emergency account Object IDs

```powershell
Connect-MgGraph -Scopes "User.Read.All"

# Replace with your emergency account UPNs
Get-MgUser -Filter "userPrincipalName eq 'emergency1@yourdomain.tld'" | Select-Object Id, DisplayName, AccountEnabled
Get-MgUser -Filter "userPrincipalName eq 'emergency2@yourdomain.tld'" | Select-Object Id, DisplayName, AccountEnabled
```

Both accounts must be **enabled**. The deploy script re-validates this via Graph before provisioning anything.

---

## Deploy (interactive, recommended)

```powershell
pwsh ./scripts/Deploy-BreakGlassEnforcer.ps1
```

You will be prompted for:

- Resource group name (default: `rg-ca-automation`)
- Logic App name (default: `CA-BreakGlass-Enforcer`)
- Azure region (default: `eastus`)
- Emergency account 1 Object ID
- Emergency account 2 Object ID (must differ from account 1)
- Recurrence interval in minutes (default: `30`, minimum: `5`)

The script:
1. Installs / imports required modules
2. Runs `Connect-AzAccount` if no context exists
3. Calls Graph `GET /v1.0/users/{id}` for each Object ID and aborts if not a user or if disabled
4. Creates the resource group if needed
5. Runs `Test-AzResourceGroupDeployment` (errors surface immediately)
6. Deploys the ARM template
7. Runs `Grant-GraphPermissions.ps1` to assign `Policy.ReadWrite.ConditionalAccess` to the new MSI

---

## Deploy (manual, ARM only)

```powershell
# 1. Deploy the ARM template
New-AzResourceGroupDeployment `
    -ResourceGroupName "rg-ca-automation" `
    -TemplateFile "./infrastructure/CA-BreakGlass-Enforcer.json" `
    -logicAppName "CA-BreakGlass-Enforcer" `
    -emergencyAccount1ObjectId "<object-id-1>" `
    -emergencyAccount2ObjectId "<object-id-2>" `
    -recurrenceIntervalMinutes 30

# 2. Grant Graph permission to the MSI
#    Use the principalId from the deployment Outputs
./scripts/Grant-GraphPermissions.ps1 -PrincipalId "<principal-id>"
```

---

## Post-deployment validation

```powershell
pwsh ./scripts/Test-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation"
```

This triggers the recurrence manually, waits 15 seconds, then reports the latest run status and a portal link.

Additional manual checks:

1. Open a CA policy in the Entra portal. Confirm both emergency account Object IDs appear under **Users -> Exclude -> Users**.
2. In **Entra ID -> Monitoring -> Audit logs**, filter `Activity = Update conditional access policy` and confirm the initiator is the Logic App's Managed Identity display name.
3. In the Logic App run history, expand the most recent run and inspect the `RunSummary` output — each patched policy emits an `OK` or `FAIL` line.

---

## Maintenance

### Change the emergency accounts

Redeploy with new Object IDs:

```powershell
pwsh ./scripts/Deploy-BreakGlassEnforcer.ps1
```

ARM is idempotent — redeploying the same template with updated parameters simply updates the Logic App in place.

Note: redeploying does **not** remove the old emergency account IDs from existing CA policies. If the old accounts have been deleted, run:

```powershell
pwsh ./scripts/Remove-BreakGlassEnforcer.ps1 `
    -ResourceGroupName "rg-ca-automation" `
    -RestorePolicies `
    -EmergencyAccount1ObjectId "<old-id-1>" `
    -EmergencyAccount2ObjectId "<old-id-2>"
```

(then redeploy with the new IDs).

### Change the recurrence interval

Re-run the deploy script and enter a new interval. Minimum 5 minutes, maximum 1440 (1 day).

### Add a third emergency account

The template parameterizes exactly two accounts. To support three, add a third parameter and a third `AppendToArrayVariable` in the workflow. Keep in mind that Microsoft recommends exactly two break-glass accounts — a third typically indicates a process problem, not a technical one.

---

## Diagnostic export

Logic App run history is retained for only 90 days in the resource. To keep an auditable trail, enable diagnostic settings to a Log Analytics workspace:

```bash
az monitor diagnostic-settings create \
  --name 'ca-emergency-diag' \
  --resource '<logic-app-resource-id>' \
  --workspace '<log-analytics-workspace-resource-id>' \
  --logs '[{"category":"WorkflowRuntime","enabled":true}]'
```

Example KQL query — failures in the last 24 hours:

```kusto
AzureDiagnostics
| where ResourceType == "WORKFLOWS"
| where status_s == "Failed"
| where TimeGenerated > ago(24h)
| project TimeGenerated, resource_workflowName_s, resource_actionName_s, status_s, code_s, error_message_s
```
