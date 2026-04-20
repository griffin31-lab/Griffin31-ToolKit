# Troubleshooting — CA Emergency Account Exclusion Logic App

## Common issues

---

### Workflow fails at "GET all conditional access policies"

**Symptom:** The first HTTP action returns 401 or 403.

**Cause:** The Managed Identity has not been granted `Policy.ReadWrite.ConditionalAccess`, or the grant has not propagated yet.

**Fix:**
```powershell
./scripts/Grant-GraphPermissions.ps1 -PrincipalId "<managed-identity-principal-id>"
```

Verify the assignment:
```powershell
Connect-MgGraph -Scopes "Application.Read.All"
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId "<principal-id>"
```

If you just granted the role, wait 1-2 minutes for Azure AD replication and re-run.

---

### Managed Identity not found right after deployment

**Symptom:** `Grant-GraphPermissions.ps1` throws "service principal not found."

**Cause:** The Logic App was just deployed and the Managed Identity has not fully propagated to the directory.

**Fix:** Wait 60-120 seconds and re-run the grant script.

---

### PATCH returns 412 Precondition Failed

**Symptom:** The PATCH action on a policy fails with HTTP 412.

**Cause:** Another admin (or another automation) edited the policy between the GET and PATCH calls — the `If-Match` header protected you from silently overwriting their change.

**Fix:** Expected behavior. The policy will be re-processed on the next recurrence (default: 30 minutes). If the pattern is frequent, increase the recurrence interval or coordinate with the other writer.

---

### PATCH returns 409 Conflict

**Symptom:** Occasional 409 on the PATCH call.

**Cause:** Transient conflict with another concurrent edit.

**Fix:** The workflow uses `concurrency: repetitions: 1` (sequential). The next run will retry automatically.

---

### Emergency account was deleted

**Symptom:** Graph returns an error when patching, or the account ID simply no longer resolves.

**Cause:** The break-glass user object was deleted from Entra ID.

**Fix:** Stop using that Object ID.

1. Deploy a replacement break-glass user
2. Remove the stale Object ID from existing policies:
   ```powershell
   ./scripts/Remove-BreakGlassEnforcer.ps1 `
       -ResourceGroupName "rg-ca-automation" `
       -RestorePolicies `
       -EmergencyAccount1ObjectId "<stale-id>" `
       -EmergencyAccount2ObjectId "<other-current-id>"
   ```
3. Redeploy the Logic App with the new Object ID.

---

### Workflow runs but policies are not being updated

**Symptom:** Run history shows success but a CA policy still does not list the exclusions.

**Cause 1:** The policy is disabled or in report-only mode. The workflow deliberately skips policies where `state != 'enabled'`. Patching disabled or staged policies would silently mutate in-progress work.

**Cause 2:** Both emergency accounts are already in the policy's `excludeUsers` — the workflow short-circuits.

**Verification:** Open the policy in the Entra portal and confirm:
- `state == enabled`
- Both emergency accounts under **Users -> Exclude -> Users**

---

### Contributor permission denied on the resource group

**Symptom:** The deploy script fails at "Validating ARM template" or at `New-AzResourceGroupDeployment`.

**Cause:** The signed-in account lacks Contributor on the target resource group.

**Fix:**
- If the resource group is locked down under PIM, activate the Contributor role assignment first.
- If the group does not exist, grant Contributor at the subscription level temporarily so the group can be created.

---

### Diagnostic logs not appearing in Log Analytics

**Symptom:** Run history is visible in the Azure portal but nothing shows up in Log Analytics.

**Cause:** Diagnostic settings are not configured on the Logic App.

**Fix:**
```bash
az monitor diagnostic-settings create \
  --name 'ca-emergency-diag' \
  --resource '<logic-app-resource-id>' \
  --workspace '<log-analytics-workspace-resource-id>' \
  --logs '[{"category":"WorkflowRuntime","enabled":true}]'
```

Allow up to 15 minutes before querying.

---

## Useful diagnostic queries

### Current exclusions on all CA policies

```powershell
Connect-MgGraph -Scopes "Policy.Read.All"

$policies = Get-MgIdentityConditionalAccessPolicy -All
$policies | ForEach-Object {
    [PSCustomObject]@{
        DisplayName   = $_.DisplayName
        State         = $_.State
        ExcludedUsers = ($_.Conditions.Users.ExcludeUsers -join ", ")
    }
} | Format-Table -AutoSize
```

### Audit log: policy changes made by the Logic App

In the Azure portal -> **Entra ID -> Monitoring -> Audit logs**:
- Activity: `Update conditional access policy`
- Initiated by (actor): the Logic App's Managed Identity display name (equals the Logic App name)
