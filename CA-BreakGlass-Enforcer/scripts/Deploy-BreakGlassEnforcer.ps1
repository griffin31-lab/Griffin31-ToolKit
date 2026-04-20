#Requires -Version 7.0
[CmdletBinding()]
param()

<#
.SYNOPSIS
    Interactively deploys the CA Emergency Account Exclusion Logic App (Consumption tier).

.DESCRIPTION
    Walks through deployment step by step, prompting for values with sensible defaults.
    Connects to Azure, creates the resource group if needed, deploys the ARM template,
    validates that the emergency account Object IDs resolve to enabled users, and
    grants the required Microsoft Graph API permissions to the Managed Identity.

.EXAMPLE
    pwsh ./Deploy-BreakGlassEnforcer.ps1

.NOTES
    Version : 1.1
    Author  : Griffin31 Security Team

    Required Graph API Permissions (granted automatically to the Managed Identity):
    - Policy.ReadWrite.ConditionalAccess
#>

$ErrorActionPreference = 'Stop'

function Read-WithDefault {
    param([string]$Question, [string]$Default)
    $answer = Read-Host "$Question$(if ($Default) { " [$Default]" })"
    if ([string]::IsNullOrWhiteSpace($answer)) { $Default } else { $answer }
}

function Write-Step {
    param([string]$Text)
    Write-Host "`n[$(([array]$script:step++)[0])] $Text" -ForegroundColor Cyan
}

$script:step = 1

# == Banner ====================================================================
Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "   CA Emergency Account Exclusion - Logic App Deployment" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Enter to accept the default shown in [brackets]." -ForegroundColor White

# == Collect inputs ============================================================
$ResourceGroupName = Read-WithDefault "Resource group name" "rg-ca-automation"
$LogicAppName      = Read-WithDefault "Logic App name"      "CA-BreakGlass-Enforcer"
$Location          = Read-WithDefault "Azure region"        "eastus"

Write-Host ""
Write-Host "Emergency account Object IDs (Entra ID > Users > select user > Object ID):" -ForegroundColor Cyan

$guidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

do {
    $Account1 = Read-Host "Emergency account 1 Object ID"
    if ($Account1 -notmatch $guidPattern) { Write-Host "   Invalid GUID - please try again." -ForegroundColor Red }
} until ($Account1 -match $guidPattern)

do {
    $Account2 = Read-Host "Emergency account 2 Object ID"
    if ($Account2 -notmatch $guidPattern) { Write-Host "   Invalid GUID - please try again." -ForegroundColor Red }
    elseif ($Account2.Trim().ToLower() -eq $Account1.Trim().ToLower()) {
        Write-Host "   Account 2 must be different from Account 1." -ForegroundColor Red
        $Account2 = $null
    }
} until ($Account2 -match $guidPattern)

$intervalInput            = Read-WithDefault "Recurrence interval (minutes)" "30"
$RecurrenceIntervalMinutes = [int]$intervalInput

# == Confirm ===================================================================
Write-Host ""
Write-Host "Deployment plan:" -ForegroundColor Cyan
Write-Host "  Resource Group : $ResourceGroupName" -ForegroundColor White
Write-Host "  Logic App      : $LogicAppName"      -ForegroundColor White
Write-Host "  Location       : $Location"          -ForegroundColor White
Write-Host "  Account 1      : $Account1"          -ForegroundColor White
Write-Host "  Account 2      : $Account2"          -ForegroundColor White
Write-Host "  Recurrence     : Every $RecurrenceIntervalMinutes minutes" -ForegroundColor White
Write-Host "  Tier           : Consumption (pay-per-execution)" -ForegroundColor White

Write-Host ""
Write-Host "WARNING: This tool modifies Conditional Access policies tenant-wide every $RecurrenceIntervalMinutes minutes." -ForegroundColor Yellow

$confirm = Read-Host "`nProceed? [Y/n]"
if ($confirm -match '^[Nn]') {
    Write-Host "Deployment cancelled." -ForegroundColor White
    exit 0
}

try {
    # == Check modules =========================================================
    Write-Step "Checking required modules"
    foreach ($module in @('Az.Accounts', 'Az.Resources')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "   Installing $module..." -ForegroundColor White
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -ErrorAction Stop
    }
    Write-Host "   [OK] Modules ready" -ForegroundColor Green

    # == Connect to Azure ======================================================
    Write-Step "Connecting to Azure"
    if (-not (Get-AzContext)) { Connect-AzAccount }
    $ctx = Get-AzContext
    Write-Host "   [OK] Signed in as : $($ctx.Account.Id)" -ForegroundColor Green
    Write-Host "        Subscription : $($ctx.Subscription.Name)" -ForegroundColor Gray

    # == Validate emergency accounts via Graph =================================
    Write-Step "Validating emergency account Object IDs resolve to enabled users"
    $graphToken = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
    $graphHeaders = @{ Authorization = "Bearer $graphToken" }

    function Test-EmergencyAccount {
        param([string]$ObjectId, [string]$Label)
        $uri = "https://graph.microsoft.com/v1.0/users/$ObjectId`?`$select=id,userPrincipalName,accountEnabled"
        try {
            $user = Invoke-RestMethod -Uri $uri -Headers $graphHeaders -Method Get
        }
        catch {
            throw "$Label ($ObjectId) did not resolve to a user object: $($_.Exception.Message)"
        }
        if (-not $user.id) {
            throw "$Label ($ObjectId) did not resolve to a user object."
        }
        if (-not $user.accountEnabled) {
            throw "$Label ($ObjectId / $($user.userPrincipalName)) is disabled. Cannot deploy with a disabled break-glass account."
        }
        Write-Host "   [OK] $Label : $($user.userPrincipalName)" -ForegroundColor Green
    }

    Test-EmergencyAccount -ObjectId $Account1 -Label "Account 1"
    Test-EmergencyAccount -ObjectId $Account2 -Label "Account 2"

    # == Resource group ========================================================
    Write-Step "Resource group"
    if (-not (Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue)) {
        Write-Host "   Creating '$ResourceGroupName' in $Location..." -ForegroundColor White
        New-AzResourceGroup -Name $ResourceGroupName -Location $Location | Out-Null
    }
    Write-Host "   [OK] Resource group ready" -ForegroundColor Green

    # == Validate template =====================================================
    Write-Step "Validating ARM template"
    $templatePath = Join-Path $PSScriptRoot "..\infrastructure\CA-BreakGlass-Enforcer.json"
    $deployParams = @{
        logicAppName              = $LogicAppName
        location                  = $Location
        emergencyAccount1ObjectId = $Account1
        emergencyAccount2ObjectId = $Account2
        recurrenceIntervalMinutes = $RecurrenceIntervalMinutes
    }

    $validation = Test-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName `
        -TemplateFile $templatePath @deployParams
    if ($validation) {
        throw "Template validation failed: $($validation.Message)"
    }
    Write-Host "   [OK] Template valid" -ForegroundColor Green

    # == Deploy ================================================================
    Write-Step "Deploying Logic App (Consumption)"
    $deployName = "ca-emergency-$(Get-Date -Format 'yyyyMMddHHmmss')"
    $deployment = New-AzResourceGroupDeployment -Name $deployName `
        -ResourceGroupName $ResourceGroupName -TemplateFile $templatePath @deployParams

    if ($deployment.ProvisioningState -ne 'Succeeded') {
        throw "Deployment failed: $($deployment.ProvisioningState)"
    }

    $principalId = $deployment.Outputs['principalId'].Value
    Write-Host "   [OK] Deployed successfully" -ForegroundColor Green
    Write-Host "        Managed Identity : $principalId" -ForegroundColor Gray

    # == Grant Graph permissions ===============================================
    Write-Step "Granting Microsoft Graph API permissions"
    & (Join-Path $PSScriptRoot "Grant-GraphPermissions.ps1") -PrincipalId $principalId

    # == Done ==================================================================
    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Green
    Write-Host "   [OK] Deployment complete" -ForegroundColor Green
    Write-Host "==============================================================" -ForegroundColor Green
    Write-Host "   Logic App '$LogicAppName' is now running in '$ResourceGroupName'." -ForegroundColor White
    Write-Host "   It will enforce emergency account exclusions every $RecurrenceIntervalMinutes minutes." -ForegroundColor White
    Write-Host ""
    Write-Host "   To export run history to Log Analytics, enable diagnostic settings:" -ForegroundColor White
    Write-Host "     az monitor diagnostic-settings create \\" -ForegroundColor Gray
    Write-Host "       --name 'ca-emergency-diag' \\" -ForegroundColor Gray
    Write-Host "       --resource $($deployment.Outputs['logicAppId'].Value) \\" -ForegroundColor Gray
    Write-Host "       --workspace <log-analytics-workspace-resource-id> \\" -ForegroundColor Gray
    Write-Host "       --logs '[{\"category\":\"WorkflowRuntime\",\"enabled\":true}]'" -ForegroundColor Gray
}
catch {
    Write-Host ""
    Write-Host "[!] Deployment failed: $_" -ForegroundColor Red
    throw
}
