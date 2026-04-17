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
    Write-Host "  Install it from: https://aka.ms/install-powershell" -ForegroundColor DarkGray
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor DarkGray
  }
  Write-Host ""
  return
}

# 2. PnP.PowerShell module (cross-platform SharePoint admin — works on Windows, macOS, Linux)
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  Write-Host ""
  Write-Host "  [!] PnP.PowerShell module is not installed." -ForegroundColor Red
  Write-Host ""
  $installChoice = Read-Host "  Would you like to install it now? (Y/n)"
  if ($installChoice -eq 'n' -or $installChoice -eq 'N') {
    Write-Host "  Cannot continue without PnP.PowerShell. Exiting." -ForegroundColor Red
    return
  }
  Write-Host "  Installing PnP.PowerShell (this may take a few minutes)..." -ForegroundColor Cyan
  Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
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
  Write-Host "  Installing Microsoft.Graph (this may take a few minutes)..." -ForegroundColor Cyan
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
  Write-ColorText "===================================================================" -Color DarkBlue
  Write-ColorText "           SHAREPOINT SITES AUDIT MANAGER v1.0" -Color White
  Write-ColorText "===================================================================" -Color DarkBlue
  Write-ColorText ""
  if ($Title) {
    Write-ColorText "  $Title" -Color Cyan
    Write-ColorText ""
  }
}

function Get-TenantInfo {
  Write-Header "TENANT CONFIGURATION"

  do {
    Write-ColorText "Enter the tenant domain or UPN:" -Color White
    Write-ColorText "Examples: contoso.onmicrosoft.com OR admin@contoso.onmicrosoft.com" -Color Gray
    Write-Host -NoNewline "Tenant: " -ForegroundColor Cyan
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
  return @{
    DataDir    = $dataDir
    ReportsDir = $reportsDir
    ConfigPath = Join-Path $TenantInfo.TenantDir "config.json"
  }
}

function Get-PnPClientId {
  param($Directories)
  if (Test-Path $Directories.ConfigPath) {
    try {
      $cfg = Get-Content $Directories.ConfigPath -Raw | ConvertFrom-Json
      if ($cfg.ClientId) { return $cfg.ClientId }
    } catch {}
  }
  return $null
}

function Invoke-FirstTimeSetup {
  param($TenantInfo, $Directories)
  Write-Header "FIRST-TIME SETUP - $($TenantInfo.Domain)"

  $regScript = Join-Path $PSScriptRoot "Register-PnPApp.ps1"
  if (-not (Test-Path $regScript)) {
    Write-ColorText "ERROR: Register-PnPApp.ps1 not found." -Color Red
    Write-ColorText "`nPress any key..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    return
  }

  try {
    & $regScript -TenantDomain $TenantInfo.Domain -ConfigPath $Directories.ConfigPath
    Write-ColorText "`nSetup finished. Check above for success or errors." -Color Green
  } catch {
    Write-ColorText "`nSetup failed: $($_.Exception.Message)" -Color Red
  }
  Write-ColorText "`nPress any key..." -Color Gray
  $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Show-MainMenu {
  param($TenantInfo)
  Write-Header "MAIN MENU - $($TenantInfo.Domain)"
  Write-ColorText "Tenant directory: $($TenantInfo.TenantDir)" -Color Gray
  Write-ColorText ""
  Write-ColorText "1. Export data only (SharePoint + Graph)" -Color White
  Write-ColorText "2. Analyze existing data" -Color White
  Write-ColorText "3. Full pipeline — Sample mode (top 100 sites by storage)" -Color Green
  Write-ColorText "4. Full pipeline — Full scan (all sites; slow on large tenants)" -Color White
  Write-ColorText "5. View previous reports" -Color White
  Write-ColorText "6. Exit" -Color White
  Write-ColorText ""
  Write-Host -NoNewline "Select option (1-6): " -ForegroundColor Cyan
}

function Confirm-Setup {
  param($TenantInfo, $Directories)
  # If config exists, nothing to do.
  if (Get-PnPClientId -Directories $Directories) { return $true }

  Write-Host ""
  Write-ColorText "No PnP app is registered for this tenant yet." -Color White
  Write-ColorText "First-time setup will register an Entra ID app + generate a cert." -Color DarkGray
  Write-ColorText "You will need Global Administrator access ONCE to approve consent." -Color DarkGray
  Write-Host ""
  Write-Host -NoNewline "Run first-time setup now? (Y/n): " -ForegroundColor Cyan
  $answer = Read-Host
  if ($answer -eq 'n' -or $answer -eq 'N') {
    Write-ColorText "Cancelled. Cannot run without a registered app." -Color Red
    Write-ColorText "Press any key..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    return $false
  }

  Invoke-FirstTimeSetup -TenantInfo $TenantInfo -Directories $Directories
  # Verify setup succeeded
  if (-not (Get-PnPClientId -Directories $Directories)) {
    return $false
  }
  return $true
}

function Invoke-DataExport {
  param($TenantInfo, $Directories, [bool]$FullScan = $false)
  if (-not (Confirm-Setup -TenantInfo $TenantInfo -Directories $Directories)) { return $false }

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
    # Require PS7 for the spawned subprocess — sub-scripts use cross-platform APIs that don't exist in Windows PS 5.1
    $pwshCmd = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if (-not $pwshCmd) {
      Write-ColorText "`n[!] 'pwsh' (PowerShell 7) not found on PATH. This tool requires PS7." -Color Red
      Write-ColorText "    Install from https://aka.ms/install-powershell" -Color DarkGray
      Write-ColorText "`nPress any key..." -Color Gray
      $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
      return $false
    }

    $pwshArgs = @(
      '-NoProfile', '-ExecutionPolicy', 'Bypass',
      '-File', $exportScript,
      '-OutputFolder', $Directories.DataDir,
      '-TenantDomain', $TenantInfo.Domain,
      '-SpoAdminUrl', $TenantInfo.SpoAdminUrl,
      '-ConfigPath', $Directories.ConfigPath
    )
    if ($FullScan) { $pwshArgs += '-FullScan' }

    $process = Start-Process -FilePath $pwshCmd -ArgumentList $pwshArgs -Wait -PassThru -NoNewWindow

    # Exit code 2 = app registration was deleted; Export-Data already moved config.json aside.
    # Auto-recover ONCE: run setup, then retry the export. Never loop (prevents runaway setup
    # loops if Azure propagation is unusually slow).
    if ($process.ExitCode -eq 2) {
      Write-ColorText "`n[!] Stale app detected — running setup to register a new app..." -Color Yellow
      Invoke-FirstTimeSetup -TenantInfo $TenantInfo -Directories $Directories
      if (-not (Get-PnPClientId -Directories $Directories)) {
        Write-ColorText "Setup did not complete. Cancelling." -Color Red
        return $false
      }
      Write-ColorText "`nRetrying export with the new app..." -Color Cyan
      $process = Start-Process -FilePath $pwshCmd -ArgumentList $pwshArgs -Wait -PassThru -NoNewWindow

      # Second attempt — if still exit 2, Azure propagation is unusually slow for the new app.
      # Don't loop setup. Tell the user to wait and retry manually.
      if ($process.ExitCode -eq 2) {
        Write-ColorText "`n[!] The new app was registered but Azure has not finished propagating it." -Color Yellow
        Write-ColorText "    This can occasionally take 2-5 minutes. Wait a moment, then pick the" -Color Yellow
        Write-ColorText "    same menu option again. No need to re-run setup." -Color Yellow
        return $false
      }
    }

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
      Write-ColorText "Run Export first (option 1 or 3)." -Color White
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
      Write-Host -NoNewline "`nOpen report in browser? (Y/N): " -ForegroundColor Cyan
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
    Write-ColorText "`nAnalysis completed but no report file found." -Color White
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
    Write-ColorText "No reports found in $($Directories.ReportsDir)" -Color DarkGray
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
  Write-Host -NoNewline "Select report to open (1-$($reports.Count)) or Enter to return: " -ForegroundColor Cyan
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
