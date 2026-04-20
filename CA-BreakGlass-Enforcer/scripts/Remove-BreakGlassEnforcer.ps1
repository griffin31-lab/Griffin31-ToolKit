#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$LogicAppName = 'CA-BreakGlass-Enforcer',

    [switch]$RemoveResourceGroup,

    [switch]$RestorePolicies,

    [Parameter(Mandatory = $false)]
    [string]$EmergencyAccount1ObjectId,

    [Parameter(Mandatory = $false)]
    [string]$EmergencyAccount2ObjectId
)

<#
.SYNOPSIS
    Removes the CA Emergency Account Exclusion Logic App, with optional policy cleanup.

.DESCRIPTION
    Deletes the Logic App (or entire resource group). Optionally, when -RestorePolicies
    is specified, iterates all Conditional Access policies and removes the two
    emergency account Object IDs from each policy's excludeUsers before removal.

    Without -RestorePolicies the Logic App is removed but the emergency account Object
    IDs remain in every CA policy's excludeUsers. A warning is printed.

.PARAMETER ResourceGroupName
    The resource group where the Logic App is deployed.

.PARAMETER LogicAppName
    Name of the Logic App to remove.

.PARAMETER RemoveResourceGroup
    If specified, deletes the entire resource group (and all resources in it).

.PARAMETER RestorePolicies
    If specified, removes the two emergency account Object IDs from every CA policy
    before removing the Logic App. Requires -EmergencyAccount1ObjectId and
    -EmergencyAccount2ObjectId.

.PARAMETER EmergencyAccount1ObjectId
    Object ID of emergency account 1 (required with -RestorePolicies).

.PARAMETER EmergencyAccount2ObjectId
    Object ID of emergency account 2 (required with -RestorePolicies).

.EXAMPLE
    ./Remove-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation"

.EXAMPLE
    ./Remove-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation" `
        -RestorePolicies -EmergencyAccount1ObjectId "<id1>" -EmergencyAccount2ObjectId "<id2>"

.EXAMPLE
    ./Remove-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation" -RemoveResourceGroup

.NOTES
    Version : 1.1
    Author  : Griffin31 Security Team
#>

$ErrorActionPreference = 'Stop'

if (-not (Get-AzContext)) { Connect-AzAccount }

if ($RestorePolicies) {
    if (-not $EmergencyAccount1ObjectId -or -not $EmergencyAccount2ObjectId) {
        Write-Host "[!] -RestorePolicies requires both -EmergencyAccount1ObjectId and -EmergencyAccount2ObjectId." -ForegroundColor Red
        exit 1
    }

    Write-Host ""
    Write-Host "Restoring CA policies (removing emergency account exclusions)..." -ForegroundColor Cyan

    $token = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
    $headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }

    $list = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" `
        -Headers $headers -Method Get
    $targets = @($EmergencyAccount1ObjectId, $EmergencyAccount2ObjectId) | ForEach-Object { $_.ToLower() }
    $patchedCount = 0

    foreach ($p in $list.value) {
        $excluded = @($p.conditions.users.excludeUsers)
        if (-not $excluded) { continue }
        $remaining = @($excluded | Where-Object { $targets -notcontains $_.ToLower() })
        if ($remaining.Count -eq $excluded.Count) { continue }

        Write-Host "   Restoring: $($p.displayName)" -ForegroundColor White
        if ($PSCmdlet.ShouldProcess($p.displayName, "Remove emergency account exclusions")) {
            $body = @{
                conditions = @{
                    users = @{
                        excludeUsers = $remaining
                    }
                }
            } | ConvertTo-Json -Depth 10

            $patchHeaders = $headers.Clone()
            if ($p.'@odata.etag') { $patchHeaders['If-Match'] = $p.'@odata.etag' }

            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($p.id)" `
                -Headers $patchHeaders -Method Patch -Body $body | Out-Null
            $patchedCount++
        }
    }
    Write-Host "   [OK] Restored $patchedCount policies." -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host "WARNING: Emergency account Object IDs remain in every CA policy's excludeUsers." -ForegroundColor Yellow
    Write-Host "         Re-run with -RestorePolicies -EmergencyAccount1ObjectId ... -EmergencyAccount2ObjectId ..." -ForegroundColor Yellow
    Write-Host "         to remove the exclusions from all policies." -ForegroundColor Yellow
    Write-Host ""
}

if ($RemoveResourceGroup) {
    Write-Host ""
    Write-Host "WARNING: This will delete the entire resource group '$ResourceGroupName' and ALL resources in it." -ForegroundColor Yellow
    if ($PSCmdlet.ShouldProcess($ResourceGroupName, "Delete resource group")) {
        Remove-AzResourceGroup -Name $ResourceGroupName -Force
        Write-Host "[OK] Resource group '$ResourceGroupName' deleted." -ForegroundColor Green
    }
}
else {
    Write-Host "Removing Logic App '$LogicAppName' from '$ResourceGroupName'..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($LogicAppName, "Delete Logic App")) {
        Remove-AzResource -ResourceGroupName $ResourceGroupName `
            -ResourceName $LogicAppName -ResourceType "Microsoft.Logic/workflows" -Force
        Write-Host "[OK] Logic App '$LogicAppName' removed." -ForegroundColor Green
    }
}
