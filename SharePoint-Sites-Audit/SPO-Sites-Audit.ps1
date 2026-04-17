[CmdletBinding()]
param(
    [string]$DataDir    = "data",
    [string]$ReportsDir = "reports"
)

# Orchestrator: runs all analysis modules against a data export, then generates the HTML report.
Write-Host "Starting SharePoint Sites Audit..." -ForegroundColor Green

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([System.IO.Path]::IsPathRooted($DataDir))    { $dataPath = $DataDir }    else { $dataPath = Join-Path $scriptDir $DataDir }
if ([System.IO.Path]::IsPathRooted($ReportsDir)) { $reportsPath = $ReportsDir } else { $reportsPath = Join-Path $scriptDir $ReportsDir }

if (-not (Test-Path $dataPath))    { New-Item -Path $dataPath    -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $reportsPath)) { New-Item -Path $reportsPath -ItemType Directory -Force | Out-Null }

try {
    Write-Host "Running per-site analysis..." -ForegroundColor Yellow
    & "$scriptDir/modules/Analyze-Sites.ps1"      -DataDir $dataPath -OutputPath "$dataPath/site-findings.json"

    Write-Host "Running OneDrive analysis..." -ForegroundColor Yellow
    & "$scriptDir/modules/Analyze-OneDrive.ps1"   -DataDir $dataPath -OutputPath "$dataPath/onedrive-findings.json"

    Write-Host "Running Groups & Teams analysis..." -ForegroundColor Yellow
    & "$scriptDir/modules/Analyze-GroupsTeams.ps1" -DataDir $dataPath -OutputPath "$dataPath/group-findings.json"

    Write-Host "Aggregating key insights..." -ForegroundColor Yellow
    & "$scriptDir/modules/Analyze-KeyInsights.ps1" -DataDir $dataPath -OutputPath "$dataPath/key-insights.json"

    Write-Host "Generating HTML report..." -ForegroundColor Yellow
    $reportPath = "$reportsPath/SP_Sites_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    & "$scriptDir/modules/Generate-Report.ps1" -AnalysisDir $dataPath -OutputPath $reportPath

    Write-Host "Analysis complete. Report: $reportPath" -ForegroundColor Green
} catch {
    Write-Error "Analysis failed: $($_.Exception.Message)"
    exit 1
}
