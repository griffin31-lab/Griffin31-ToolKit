[CmdletBinding()]
param(
    [string]$DataDir = "data",
    [string]$ReportsDir = "reports"
)

# Main orchestrator script for Conditional Access Gap Analysis
Write-Host "Starting Conditional Access Gap Analysis..." -ForegroundColor Green

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Support both absolute and relative paths
if ([System.IO.Path]::IsPathRooted($DataDir)) {
    $dataPath = $DataDir
} else {
    $dataPath = Join-Path $scriptDir $DataDir
}

if ([System.IO.Path]::IsPathRooted($ReportsDir)) {
    $reportsPath = $ReportsDir
} else {
    $reportsPath = Join-Path $scriptDir $ReportsDir
}

# Ensure directories exist
New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
New-Item -Path $reportsPath -ItemType Directory -Force | Out-Null

try {
    # 1. Break Glass Detection
    Write-Host "Running break-glass detection..." -ForegroundColor Yellow
    & "$scriptDir\modules\Detect-BreakGlass.ps1" -ExportDir $dataPath -OutputPath "$dataPath\breakglass.json"
    
    # 2. Nested Group Detection
    Write-Host "Running nested group analysis..." -ForegroundColor Yellow
    & "$scriptDir\modules\Analyze-NestedGroups.ps1" -ExportDir $dataPath -OutputPath "$dataPath\nested-groups.json"
    
    # 3. Policy Evaluation
    Write-Host "Running policy gap analysis..." -ForegroundColor Yellow
    & "$scriptDir\modules\Analyze-PolicyGaps.ps1" -ExportDir $dataPath -OutputPath "$dataPath\policy-gaps.json"
    
    # 4. Missing Controls Analysis
    Write-Host "Running missing controls analysis..." -ForegroundColor Yellow
    & "$scriptDir\modules\Analyze-MissingControls.ps1" -ExportDir $dataPath -OutputPath "$dataPath\missing-controls.json"

    # 5. Key Insights (posture score + 15 checks)
    Write-Host "Running key insights analysis..." -ForegroundColor Yellow
    & "$scriptDir\modules\Analyze-KeyInsights.ps1" -ExportDir $dataPath -AnalysisDir $dataPath -OutputPath "$dataPath\key-insights.json"

    # 6. Generate HTML Report
    Write-Host "Generating HTML report..." -ForegroundColor Yellow
    $reportPath = "$reportsPath\CA_Gap_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    & "$scriptDir\modules\Generate-Report.ps1" -AnalysisDir $dataPath -OutputPath $reportPath
    
    Write-Host "Analysis complete! Report available at: $reportPath" -ForegroundColor Green
    
} catch {
    Write-Error "Analysis failed: $($_.Exception.Message)"
    exit 1
}
