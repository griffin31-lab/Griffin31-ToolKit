#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$PrincipalId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId
)

<#
.SYNOPSIS
    Grants Microsoft Graph API permissions to the Logic App's Managed Identity.

.DESCRIPTION
    Assigns Policy.ReadWrite.ConditionalAccess to the system-assigned managed identity
    of the CA Emergency Account Exclusion Logic App. ReadWrite covers reads on CA
    policies, so a separate Policy.Read.All role is not required.

.PARAMETER PrincipalId
    The Object ID of the Logic App's system-assigned managed identity.

.PARAMETER TenantId
    Azure AD Tenant ID. Optional - uses current context if omitted.

.EXAMPLE
    ./Grant-GraphPermissions.ps1 -PrincipalId "12345678-1234-1234-1234-123456789012"

.NOTES
    Version  : 1.1
    Author   : Griffin31 Security Team
    Requires : Microsoft.Graph.Authentication, Microsoft.Graph.Applications
    Requires : Global Administrator or Privileged Role Administrator
#>

$ErrorActionPreference = 'Stop'

# Stable Microsoft Graph app-role ID - no runtime lookup required
$Permissions = @(
    @{ Name = 'Policy.ReadWrite.ConditionalAccess'; Id = '01c0a623-fc9b-48e9-b794-0756f8e8f067' }
)

Write-Host ""
Write-Host "Checking required modules..." -ForegroundColor Cyan
foreach ($m in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "   Installing $m..." -ForegroundColor White
        Install-Module -Name $m -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module $m -ErrorAction Stop
}
Write-Host "   [OK] Modules loaded" -ForegroundColor Green

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$connectParams = @{ Scopes = @('Application.Read.All', 'AppRoleAssignment.ReadWrite.All'); NoWelcome = $true }
if ($TenantId) { $connectParams['TenantId'] = $TenantId }
Connect-MgGraph @connectParams
Write-Host "   [OK] Connected" -ForegroundColor Green

# Get Microsoft Graph service principal
$graphSP = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
if (-not $graphSP) { throw "Microsoft Graph service principal not found in tenant." }

Write-Host ""
Write-Host "Granting permissions to principal: $PrincipalId" -ForegroundColor Cyan
foreach ($perm in $Permissions) {
    Write-Host "   - $($perm.Name)..." -ForegroundColor White

    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId |
        Where-Object { $_.AppRoleId -eq $perm.Id -and $_.ResourceId -eq $graphSP.Id }

    if ($existing) {
        Write-Host "     Already assigned - skipping." -ForegroundColor Gray
        continue
    }

    if ($PSCmdlet.ShouldProcess($perm.Name, "Grant AppRole")) {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId `
            -PrincipalId $PrincipalId -ResourceId $graphSP.Id -AppRoleId $perm.Id | Out-Null
        Write-Host "     [OK] Granted" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "[OK] All permissions granted successfully." -ForegroundColor Green
Write-Host ""
