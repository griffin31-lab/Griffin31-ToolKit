[CmdletBinding()]
param()

<#
.SYNOPSIS
    SharePoint Sites Audit Manager - Interactive menu system

.DESCRIPTION
    Iterates SharePoint sites, OneDrive accounts, M365 groups, and Teams.
    Runs 14 per-entity security checks and produces a self-contained HTML report.

.PREREQUISITES
    - PowerShell 7.x
    - Microsoft.Online.SharePoint.PowerShell module
    - Microsoft.Graph module

.AUTHOR
    Griffin31 Security Team
#>

# ── Prerequisite checks ──
$ErrorActionPreference = "Stop"

# 1. PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Host ""
  Write-Host "  [!] This script requires PowerShell 7 or later." -ForegroundColor Red
  Write-Host ""
  if ($IsWindows -or $env:OS -match "Windows") {
    Write-Host "  Install it from: https://aka.ms/install-powershell" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
  }
  Write-Host ""
  return
}

# 2. Microsoft.Online.SharePoint.PowerShell module
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
  Write-Host ""
  Write-Host "  [!] Microsoft.Online.SharePoint.PowerShell module is not installed." -ForegroundColor Red
  Write-Host ""
  $installChoice = Read-Host "  Would you like to install it now? (Y/n)"
  if ($installChoice -eq 'n' -or $installChoice -eq 'N') {
    Write-Host "  Cannot continue without the SPO module. Exiting." -ForegroundColor Red
    return
  }
  Write-Host "  Installing Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Yellow
  Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
  Write-Host "  Installed." -ForegroundColor Green
}

# 3. Microsoft.Graph module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
  Write-Host ""
  Write-Host "  [!] Microsoft.Graph module is not installed." -ForegroundColor Red
  Write-Host ""
  $installChoice = Read-Host "  Would you like to install it now? (Y/n)"
  if ($installChoice -eq 'n' -or $installChoice -eq 'N') {
    Write-Host "  Cannot continue without Microsoft.Graph. Exiting." -ForegroundColor Red
    return
  }
  Write-Host "  Installing Microsoft.Graph (this may take a few minutes)..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  Write-Host "  Installed." -ForegroundColor Green
} else {
  Remove-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
  Import-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
}

# ── UI helpers ──
function Write-ColorText {
  param([string]$Text, [ConsoleColor]$Color = 'White')
  Write-Host $Text -ForegroundColor $Color
}

function Write-Header {
  param([string]$Title)
  Clear-Host
  Write-ColorText "===================================================================" -Color Cyan
  Write-ColorText "           SHAREPOINT SITES AUDIT MANAGER v1.0" -Color Yellow
  Write-ColorText "===================================================================" -Color Cyan
  Write-ColorText ""
  if ($Title) {
    Write-ColorText "  $Title" -Color Green
    Write-ColorText ""
  }
}

function Get-TenantInfo {
  Write-Header "TENANT CONFIGURATION"

  do {
    Write-ColorText "Enter the tenant domain or UPN:" -Color White
    Write-ColorText "Examples: contoso.onmicrosoft.com OR admin@contoso.onmicrosoft.com" -Color Gray
    Write-Host -NoNewline "Tenant: " -ForegroundColor Yellow
    $inputVal = Read-Host

    $domainPattern = '^[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$'
    $upnPattern    = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$'
    if ($inputVal -match $upnPattern) {
      $upn = $inputVal
      $domain = $inputVal.Split('@')[1]
    } elseif ($inputVal -match $domainPattern) {
      $domain = $inputVal
      $upn = $null
    } else {
      Write-ColorText "`nInvalid format." -Color Red
      Write-ColorText "Press any key to try again..." -Color Gray
      $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
      continue
    }
    if ($domain -match '\.\.' -or $domain.Contains('/') -or $domain.Contains('\') -or $domain.Contains(':')) {
      Write-ColorText "`nRejected: domain contains invalid characters." -Color Red
      $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
      continue
    }

    # SPO admin URL — extract tenant part (the part before .onmicrosoft.com)
    $tenantPart = $domain -replace '\.onmicrosoft\.com$', '' -replace '\..*$', ''
    $spoAdminUrl = "https://$tenantPart-admin.sharepoint.com"

    Write-ColorText "`nTenant domain: $domain" -Color Green
    Write-ColorText "SPO admin URL: $spoAdminUrl" -Color Gray
    if ($upn) { Write-ColorText "Admin user: $upn" -Color Gray }

    return @{
      UPN         = $upn
      Domain      = $domain
      SpoAdminUrl = $spoAdminUrl
      TenantDir   = Join-Path $PSScriptRoot "tenants/$domain"
    }
  } while ($true)
}

function Initialize-TenantDirectory {
  param($TenantInfo)
  $dataDir    = Join-Path $TenantInfo.TenantDir "data"
  $reportsDir = Join-Path $TenantInfo.TenantDir "reports"
  foreach ($d in @($TenantInfo.TenantDir, $dataDir, $reportsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
  }
  return @{ DataDir = $dataDir; ReportsDir = $reportsDir }
}

function Show-MainMenu {
  param($TenantInfo)
  Write-Header "MAIN MENU - $($TenantInfo.Domain)"
  Write-ColorText "Tenant directory: $($TenantInfo.TenantDir)" -Color Gray
  Write-ColorText ""
  Write-ColorText "1. Export data only (SharePoint + Graph)" -Color White
  Write-ColorText "2. Analyze existing data" -Color White
  Write-ColorText "3. Full pipeline — Sample mode (top 100 sites by storage)" -Color Green
  Write-ColorText "4. Full pipeline — Full scan (all sites; slow on large tenants)" -Color Yellow
  Write-ColorText "5. View previous reports" -Color White
  Write-ColorText "6. Exit" -Color White
  Write-ColorText ""
  Write-Host -NoNewline "Select option (1-6): " -ForegroundColor Yellow
}

function Invoke-DataExport {
  param($TenantInfo, $Directories, [bool]$FullScan = $false)
  Write-Header "DATA EXPORT - $($TenantInfo.Domain)"
  Write-ColorText "Mode: $(if ($FullScan) { 'Full scan' } else { 'Sample (top 100 sites)' })" -Color Gray

  $exportScript = Join-Path $PSScriptRoot "Export-Data.ps1"
  if (-not (Test-Path $exportScript)) {
    Write-ColorText "ERROR: Export-Data.ps1 not found at: $exportScript" -Color Red
    Write-ColorText "Press any key to return..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    return $false
  }

  try {
    # Run in fresh PS process to avoid Graph assembly conflicts
    $pwshCmd = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if (-not $pwshCmd) { $pwshCmd = (Get-Command powershell -ErrorAction SilentlyContinue).Source }
    if (-not $pwshCmd) { throw "PowerShell not found." }

    $pwshArgs = @(
      '-NoProfile', '-ExecutionPolicy', 'Bypass',
      '-File', $exportScript,
      '-OutputFolder', $Directories.DataDir,
      '-TenantDomain', $TenantInfo.Domain,
      '-SpoAdminUrl', $TenantInfo.SpoAdminUrl
    )
    if ($FullScan) { $pwshArgs += '-FullScan' }

    $process = Start-Process -FilePath $pwshCmd -ArgumentList $pwshArgs -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) { throw "Export exited with code $($process.ExitCode)" }

    Write-ColorText "`nData export completed." -Color Green
    $files = Get-ChildItem -Path $Directories.DataDir -Filter "*.json"
    if ($files.Count -gt 0) {
      Write-ColorText "`nExported files:" -Color White
      foreach ($f in $files) { Write-ColorText "  - $($f.Name)" -Color Gray }
    }
    return $true
  } catch {
    Write-ColorText "`nERROR during export: $($_.Exception.Message)" -Color Red
    return $false
  } finally {
    Write-ColorText "`nPress any key to continue..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
  }
}

function Invoke-Analysis {
  param($TenantInfo, $Directories)
  Write-Header "ANALYSIS - $($TenantInfo.Domain)"

  $requiredFiles = @("sites.json", "groups.json", "tenant-baseline.json")
  foreach ($f in $requiredFiles) {
    $p = Join-Path $Directories.DataDir $f
    if (-not (Test-Path $p)) {
      Write-ColorText "ERROR: Missing data file: $f" -Color Red
      Write-ColorText "Run Export first (option 1 or 3)." -Color Yellow
      Write-ColorText "Press any key..." -Color Gray
      $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
      return $null
    }
  }

  $analysisScript = Join-Path $PSScriptRoot "SPO-Sites-Audit.ps1"
  if (-not (Test-Path $analysisScript)) {
    Write-ColorText "ERROR: SPO-Sites-Audit.ps1 not found." -Color Red
    Write-ColorText "Press any key..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    return $null
  }

  try {
    & $analysisScript -DataDir $Directories.DataDir -ReportsDir $Directories.ReportsDir

    $reports = Get-ChildItem -Path $Directories.ReportsDir -Filter "SP_Sites_Audit_*.html" |
               Sort-Object LastWriteTime -Descending
    if ($reports.Count -gt 0) {
      $latest = $reports[0]
      Write-ColorText "`nAnalysis complete. Report: $($latest.Name)" -Color Green
      Write-Host -NoNewline "`nOpen report in browser? (Y/N): " -ForegroundColor Yellow
      $answer = Read-Host
      if ($answer -match '^[Yy]') {
        try {
          if ($IsWindows)    { Start-Process $latest.FullName }
          elseif ($IsMacOS)  { & open $latest.FullName }
          elseif ($IsLinux)  { & xdg-open $latest.FullName }
          else { Start-Process $latest.FullName }
        } catch { Write-ColorText "Could not open: $($_.Exception.Message)" -Color Red }
      }
      return $latest.FullName
    }
    Write-ColorText "`nAnalysis completed but no report file found." -Color Yellow
    return $null
  } catch {
    Write-ColorText "`nERROR during analysis: $($_.Exception.Message)" -Color Red
    return $null
  } finally {
    Write-ColorText "`nPress any key..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
  }
}

function Show-PreviousReports {
  param($TenantInfo, $Directories)
  Write-Header "PREVIOUS REPORTS - $($TenantInfo.Domain)"

  $reports = Get-ChildItem -Path $Directories.ReportsDir -Filter "SP_Sites_Audit_*.html" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending
  if ($reports.Count -eq 0) {
    Write-ColorText "No reports found in $($Directories.ReportsDir)" -Color Yellow
    Write-ColorText "`nPress any key..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    return
  }
  Write-ColorText "Found $($reports.Count) report(s):" -Color White
  for ($i = 0; $i -lt $reports.Count; $i++) {
    $r = $reports[$i]
    Write-ColorText "$($i+1). $($r.Name)" -Color White
    Write-ColorText "   Created: $($r.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Color Gray
    Write-ColorText "   Size:    $([math]::Round($r.Length / 1KB, 2)) KB" -Color Gray
    Write-ColorText ""
  }
  Write-Host -NoNewline "Select report to open (1-$($reports.Count)) or Enter to return: " -ForegroundColor Yellow
  $sel = Read-Host
  if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $reports.Count) {
    $chosen = $reports[[int]$sel - 1]
    try {
      if ($IsWindows)   { Start-Process $chosen.FullName }
      elseif ($IsMacOS) { & open $chosen.FullName }
      elseif ($IsLinux) { & xdg-open $chosen.FullName }
      else { Start-Process $chosen.FullName }
    } catch { Write-ColorText "Could not open: $($_.Exception.Message)" -Color Red }
  }
}

# ── Main loop ──
$tenant = Get-TenantInfo
$dirs   = Initialize-TenantDirectory -TenantInfo $tenant

while ($true) {
  Show-MainMenu -TenantInfo $tenant
  $choice = Read-Host
  switch ($choice) {
    '1' { Invoke-DataExport -TenantInfo $tenant -Directories $dirs -FullScan $false | Out-Null }
    '2' { Invoke-Analysis   -TenantInfo $tenant -Directories $dirs | Out-Null }
    '3' {
      if (Invoke-DataExport -TenantInfo $tenant -Directories $dirs -FullScan $false) {
        Invoke-Analysis -TenantInfo $tenant -Directories $dirs | Out-Null
      }
    }
    '4' {
      Write-ColorText "`nFull scan iterates every site. 10k-site tenants take 20-40 min. Continue? (YES to confirm)" -Color Yellow
      $confirm = Read-Host
      if ($confirm -eq 'YES') {
        if (Invoke-DataExport -TenantInfo $tenant -Directories $dirs -FullScan $true) {
          Invoke-Analysis -TenantInfo $tenant -Directories $dirs | Out-Null
        }
      } else {
        Write-ColorText "Cancelled." -Color Gray
        Start-Sleep -Seconds 1
      }
    }
    '5' { Show-PreviousReports -TenantInfo $tenant -Directories $dirs }
    '6' { Write-ColorText "`nGoodbye." -Color Cyan; return }
    default {
      Write-ColorText "Invalid choice." -Color Red
      Start-Sleep -Seconds 1
    }
  }
}
