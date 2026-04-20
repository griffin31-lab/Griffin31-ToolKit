#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$LogicAppName = 'CA-BreakGlass-Enforcer'
)

<#
.SYNOPSIS
    Triggers a manual run of the CA Emergency Account Exclusion Logic App and checks the result.

.DESCRIPTION
    Finds the Logic App, triggers its recurrence trigger via REST, waits briefly, then
    reports the latest run status. Prints a portal link for detailed run history.

.PARAMETER ResourceGroupName
    The resource group where the Logic App is deployed.

.PARAMETER LogicAppName
    Name of the Logic App to test.

.EXAMPLE
    ./Test-BreakGlassEnforcer.ps1 -ResourceGroupName "rg-ca-automation"

.NOTES
    Version : 1.1
    Author  : Griffin31 Security Team
#>

$ErrorActionPreference = 'Stop'

if (-not (Get-AzContext)) { Connect-AzAccount }

Write-Host ""
Write-Host "Testing Logic App: $LogicAppName" -ForegroundColor Cyan

# Get the Logic App
$la = Get-AzResource -ResourceGroupName $ResourceGroupName -ResourceName $LogicAppName `
    -ResourceType "Microsoft.Logic/workflows" -ErrorAction SilentlyContinue
if (-not $la) {
    Write-Host "[!] Logic App '$LogicAppName' not found in '$ResourceGroupName'." -ForegroundColor Red
    exit 1
}
Write-Host "   [OK] Logic App found: $($la.ResourceId)" -ForegroundColor Green

# Trigger a manual run via REST
Write-Host ""
Write-Host "Triggering manual run..." -ForegroundColor Cyan
$token = (Get-AzAccessToken).Token
$triggerName = "Schedule_automation_of_emergency_account_management_within_conditional_access"
$runUri = "https://management.azure.com$($la.ResourceId)/triggers/$triggerName/run?api-version=2016-06-01"
Invoke-RestMethod -Uri $runUri -Method Post -Headers @{ Authorization = "Bearer $token" } | Out-Null

Write-Host "   Run triggered. Waiting 15 seconds for the run to complete..." -ForegroundColor Gray
Start-Sleep -Seconds 15

# Check the latest run
$runsUri = "https://management.azure.com$($la.ResourceId)/runs?api-version=2016-06-01&`$top=1"
$runs = Invoke-RestMethod -Uri $runsUri -Method Get -Headers @{ Authorization = "Bearer $token" }
$latest = $runs.value | Select-Object -First 1

if ($latest) {
    $status = $latest.properties.status
    $color = if ($status -eq 'Succeeded') { 'Green' } elseif ($status -eq 'Running') { 'Cyan' } else { 'Red' }
    Write-Host ""
    Write-Host "   Latest run status : $status" -ForegroundColor $color
    Write-Host "   Start time        : $($latest.properties.startTime)" -ForegroundColor White
    Write-Host "   End time          : $($latest.properties.endTime)"   -ForegroundColor White
}
else {
    Write-Host ""
    Write-Host "   No runs found yet. Check the Azure Portal run history." -ForegroundColor White
}

Write-Host ""
Write-Host "View run history in the portal:" -ForegroundColor Cyan
Write-Host "   https://portal.azure.com/#resource$($la.ResourceId)/runs" -ForegroundColor White
