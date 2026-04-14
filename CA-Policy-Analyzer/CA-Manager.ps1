[CmdletBinding()]
param()

<#
.SYNOPSIS
    Conditional Access Analysis Manager - Interactive menu system

.DESCRIPTION
    Interactive menu system for managing Conditional Access policy analysis.
    Handles data export, analysis, and report generation with tenant-specific organization.

.PREREQUISITES
    - PowerShell 7.x
    - Microsoft Graph PowerShell SDK
    - Appropriate permissions for Graph API access

.AUTHOR
    Griffin31 Security Team

.VERSION
    1.0
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
    Write-Host "  Then run:  pwsh .\CA-Manager.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./CA-Manager.ps1" -ForegroundColor Yellow
  }
  Write-Host ""
  return
}

# 2. Microsoft.Graph module
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
  Write-Host "  Microsoft.Graph installed successfully." -ForegroundColor Green
} else {
  # Force reimport to avoid version conflicts with stale loaded assemblies
  Remove-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
  Import-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
}

# Color functions for better UX
function Write-ColorText {
    param(
        [string]$Text,
        [ConsoleColor]$Color = 'White'
    )
    Write-Host $Text -ForegroundColor $Color
}

function Write-Header {
    param([string]$Title)
    
    Clear-Host
    Write-ColorText "═══════════════════════════════════════════════════════════════════" -Color Cyan
    Write-ColorText "           CONDITIONAL ACCESS ANALYSIS MANAGER v1.0" -Color Yellow
    Write-ColorText "═══════════════════════════════════════════════════════════════════" -Color Cyan
    Write-ColorText ""
    if ($Title) {
        Write-ColorText "  $Title" -Color Green
        Write-ColorText ""
    }
}

function Get-TenantInfo {
    Write-Header "TENANT CONFIGURATION"
    
    do {
        Write-ColorText "Enter the tenant domain or UPN for the target tenant:" -Color White
        Write-ColorText "Examples: contoso.com OR admin@contoso.com" -Color Gray
        Write-Host -NoNewline "Tenant: " -ForegroundColor Yellow
        $input = Read-Host
        
        # Strict validation — reject path separators, '..', and anything outside a safe domain charset
        $domainPattern = '^[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$'
        $upnPattern    = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$'
        if ($input -match $upnPattern) {
            $upn = $input
            $domain = $input.Split('@')[1]
        } elseif ($input -match $domainPattern) {
            $domain = $input
            $upn = $null
        } else {
            Write-ColorText "`nInvalid format. Please enter a valid domain (contoso.com) or UPN (admin@contoso.com)." -Color Red
            Write-ColorText "Press any key to try again..." -Color Gray
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            continue
        }
        # Defense in depth: refuse path traversal artefacts even if the regex above permitted them
        if ($domain -match '\.\.' -or $domain.Contains('/') -or $domain.Contains('\') -or $domain.Contains(':')) {
            Write-ColorText "`nRejected: domain contains invalid characters." -Color Red
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            continue
        }
        
        Write-ColorText "`nTarget tenant domain: $domain" -Color Green
        if ($upn) {
            Write-ColorText "User provided: $upn" -Color Gray
        }
        
        return @{
            UPN = $upn
            Domain = $domain
            TenantDir = Join-Path $PSScriptRoot "tenants\$domain"
        }
    } while ($true)
}

function Show-MainMenu {
    param($TenantInfo)
    
    Write-Header "MAIN MENU - $($TenantInfo.Domain)"
    
    Write-ColorText "Tenant Directory: $($TenantInfo.TenantDir)" -Color Gray
    Write-ColorText ""
    Write-ColorText "1. Export Data from Microsoft Graph" -Color White
    Write-ColorText "2. Analyze Existing Data" -Color White
    Write-ColorText "3. Full Pipeline (Export + Analyze)" -Color White
    Write-ColorText "4. View Previous Reports" -Color White
    Write-ColorText "5. Exit" -Color White
    Write-ColorText ""
    Write-Host -NoNewline "Select option (1-5): " -ForegroundColor Yellow
}

function Initialize-TenantDirectory {
    param($TenantInfo)
    
    # Create tenant-specific directories
    $dataDir = Join-Path $TenantInfo.TenantDir "data"
    $reportsDir = Join-Path $TenantInfo.TenantDir "reports"
    
    if (!(Test-Path $TenantInfo.TenantDir)) {
        New-Item -ItemType Directory -Path $TenantInfo.TenantDir -Force | Out-Null
        Write-ColorText "Created tenant directory: $($TenantInfo.TenantDir)" -Color Green
    }
    
    if (!(Test-Path $dataDir)) {
        New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
        Write-ColorText "Created data directory: $dataDir" -Color Green
    }
    
    if (!(Test-Path $reportsDir)) {
        New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
        Write-ColorText "Created reports directory: $reportsDir" -Color Green
    }
    
    return @{
        DataDir = $dataDir
        ReportsDir = $reportsDir
    }
}

function Invoke-DataExport {
    param($TenantInfo, $Directories)
    
    Write-Header "DATA EXPORT - $($TenantInfo.Domain)"
    
    Write-ColorText "Starting data export for tenant: $($TenantInfo.Domain)" -Color Green
    if ($TenantInfo.UPN) {
        Write-ColorText "User: $($TenantInfo.UPN)" -Color Gray
    }
    Write-ColorText "Target directory: $($Directories.DataDir)" -Color Gray
    Write-ColorText ""
    
    # Check if Export-Data.ps1 exists
    $exportScript = Join-Path $PSScriptRoot "Export-Data.ps1"
    if (!(Test-Path $exportScript)) {
        Write-ColorText "ERROR: Export-Data.ps1 not found at: $exportScript" -Color Red
        Write-ColorText "Press any key to return to main menu..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return $false
    }
    
    try {
        Write-ColorText "Executing export script..." -Color Yellow
        # Run the export script in a fresh PowerShell process to avoid module/assembly conflicts
        $pwshCmd = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
        if (-not $pwshCmd) {
            # Fallback to Windows PowerShell if pwsh is not available
            $pwshCmd = (Get-Command powershell -ErrorAction SilentlyContinue).Source
        }

        if (-not $pwshCmd) {
            throw "No PowerShell executable (pwsh or powershell) was found in PATH to spawn a clean process."
        }

        # Prepare safe argument array and use -File to avoid complex quoting
        $pwshArgs = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $exportScript,
            '-OutputFolder', $Directories.DataDir,
            '-TenantDomain', $TenantInfo.Domain
        )

        Write-ColorText "Launching external PowerShell: $pwshCmd -File $exportScript" -Color Gray

        # Invoke pwsh synchronously with redirected streams to avoid input/output interference
        $process = Start-Process -FilePath $pwshCmd -ArgumentList $pwshArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -ne 0) {
            throw "Export script exited with code $($process.ExitCode)"
        }

        Write-ColorText "`nData export completed successfully!" -Color Green
        Write-ColorText "Exported files are located in: $($Directories.DataDir)" -Color Gray
        
        # List exported files
        $exportedFiles = Get-ChildItem -Path $Directories.DataDir -Filter "*.json"
        if ($exportedFiles.Count -gt 0) {
            Write-ColorText "`nExported files:" -Color White
            foreach ($file in $exportedFiles) {
                Write-ColorText "  - $($file.Name)" -Color Gray
            }
        }
        
        return $true
    }
    catch {
        Write-ColorText "`nERROR during export: $($_.Exception.Message)" -Color Red
        return $false
    }
    finally {
        Write-ColorText "`nPress any key to continue..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

function Invoke-Analysis {
    param($TenantInfo, $Directories)
    
    Write-Header "ANALYSIS - $($TenantInfo.Domain)"
    
    # Check for required data files
    $requiredFiles = @(
        "ConditionalAccessPolicies.json",
        "DirectoryRoleAssignments.json",
        "Members.json",
        "SecurityGroups.json",
        "SecurityGroupMemberships.json"
    )
    
    $missingFiles = @()
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $Directories.DataDir $file
        if (!(Test-Path $filePath)) {
            $missingFiles += $file
        }
    }
    
    if ($missingFiles.Count -gt 0) {
        Write-ColorText "ERROR: Missing required data files:" -Color Red
        foreach ($file in $missingFiles) {
            Write-ColorText "  - $file" -Color Red
        }
        Write-ColorText "`nPlease run data export first or ensure all required files exist." -Color Yellow
        Write-ColorText "Press any key to return to main menu..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return $null
    }
    
    Write-ColorText "Starting analysis for tenant: $($TenantInfo.Domain)" -Color Green
    Write-ColorText "Data directory: $($Directories.DataDir)" -Color Gray
    Write-ColorText "Reports directory: $($Directories.ReportsDir)" -Color Gray
    Write-ColorText ""
    
    # Check if CA-Gap-Analysis.ps1 exists
    $analysisScript = Join-Path $PSScriptRoot "CA-Gap-Analysis.ps1"
    if (!(Test-Path $analysisScript)) {
        Write-ColorText "ERROR: CA-Gap-Analysis.ps1 not found at: $analysisScript" -Color Red
        Write-ColorText "Press any key to return to main menu..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return $null
    }
    
    try {
        Write-ColorText "Executing analysis script..." -Color Yellow
        & $analysisScript -DataDir $Directories.DataDir -ReportsDir $Directories.ReportsDir
        
        # Find the latest report
        $reports = Get-ChildItem -Path $Directories.ReportsDir -Filter "CA_Gap_Analysis_*.html" | Sort-Object LastWriteTime -Descending
        if ($reports.Count -gt 0) {
            $latestReport = $reports[0]
            Write-ColorText "`nAnalysis completed successfully!" -Color Green
            Write-ColorText "Report generated: $($latestReport.Name)" -Color Gray
            
            # Ask if user wants to open the report
            Write-ColorText "`nWould you like to open the report in your default browser? (Y/N): " -Color Yellow -NoNewline
            $response = Read-Host
            
            if ($response -match '^[Yy]') {
                Write-ColorText "Opening report in browser..." -Color Green
                try {
                    if ($IsWindows) {
                        Start-Process $latestReport.FullName
                    } elseif ($IsMacOS) {
                        & open $latestReport.FullName
                    } elseif ($IsLinux) {
                        & xdg-open $latestReport.FullName
                    } else {
                        # Fallback for older PowerShell versions
                        if ($env:OS -eq "Windows_NT") {
                            Start-Process $latestReport.FullName
                        } else {
                            & open $latestReport.FullName
                        }
                    }
                } catch {
                    Write-ColorText "ERROR during analysis: $($_.Exception.Message)" -Color Red
                }
            }
            
            return $latestReport.FullName
        } else {
            Write-ColorText "`nWARNING: Analysis completed but no report file found." -Color Yellow
            return $null
        }
    }
    catch {
        Write-ColorText "`nERROR during analysis: $($_.Exception.Message)" -Color Red
        return $null
    }
    finally {
        Write-ColorText "`nPress any key to continue..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

function Show-PreviousReports {
    param($TenantInfo, $Directories)
    
    Write-Header "PREVIOUS REPORTS - $($TenantInfo.Domain)"
    
    $reports = Get-ChildItem -Path $Directories.ReportsDir -Filter "CA_Gap_Analysis_*.html" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    
    if ($reports.Count -eq 0) {
        Write-ColorText "No reports found in: $($Directories.ReportsDir)" -Color Yellow
        Write-ColorText "`nPress any key to return to main menu..." -Color Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return
    }
    
    Write-ColorText "Found $($reports.Count) report(s):" -Color White
    Write-ColorText ""
    
    for ($i = 0; $i -lt $reports.Count; $i++) {
        $report = $reports[$i]
        Write-ColorText "$($i + 1). $($report.Name)" -Color White
        Write-ColorText "   Created: $($report.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Color Gray
        Write-ColorText "   Size: $([math]::Round($report.Length / 1KB, 2)) KB" -Color Gray
        Write-ColorText ""
    }
    
    Write-Host -NoNewline "Select report to open (1-$($reports.Count)) or press Enter to return: " -ForegroundColor Yellow
    $selection = Read-Host
    
    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $reports.Count) {
        $selectedReport = $reports[[int]$selection - 1]
        Write-ColorText "Opening: $($selectedReport.Name)" -Color Green
        try {
            if ($IsWindows) {
                Start-Process $selectedReport.FullName
            } elseif ($IsMacOS) {
                & open $selectedReport.FullName
            } elseif ($IsLinux) {
                & xdg-open $selectedReport.FullName
            } else {
                # Fallback for older PowerShell versions
                if ($env:OS -eq "Windows_NT") {
                    Start-Process $selectedReport.FullName
                } else {
                    & open $selectedReport.FullName
                }
            }
        } catch {
            Write-ColorText "CRITICAL ERROR: $($_.Exception.Message)" -Color Red
            Write-Host "Press any key to exit..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}

# Main execution
try {
    # Get tenant information
    $tenantInfo = Get-TenantInfo
    
    # Ensure directories exist
    $directories = Initialize-TenantDirectory -TenantInfo $tenantInfo
    
    # Main menu loop
    do {
        Show-MainMenu -TenantInfo $tenantInfo
        
        # Clear any pending input to avoid interference from spawned processes
        while ($Host.UI.RawUI.KeyAvailable) {
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
        
        $choice = Read-Host
        
        # Trim whitespace and handle empty input
        $choice = $choice.Trim()
        
        if ([string]::IsNullOrEmpty($choice)) {
            Write-ColorText "`nNo selection made. Please choose 1-5." -Color Red
            Write-ColorText "Press any key to continue..." -Color Gray
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            continue
        }
        
        switch ($choice) {
            "1" {
                Invoke-DataExport -TenantInfo $tenantInfo -Directories $directories
            }
            "2" {
                Invoke-Analysis -TenantInfo $tenantInfo -Directories $directories
            }
            "3" {
                Write-Header "FULL PIPELINE - $($tenantInfo.Domain)"
                Write-ColorText "Running complete pipeline: Export + Analysis" -Color Green
                Write-ColorText ""
                
                $exportSuccess = Invoke-DataExport -TenantInfo $tenantInfo -Directories $directories
                if ($exportSuccess) {
                    Write-ColorText "`nProceeding to analysis..." -Color Yellow
                    Start-Sleep -Seconds 2
                    Invoke-Analysis -TenantInfo $tenantInfo -Directories $directories
                } else {
                    Write-ColorText "`nSkipping analysis due to export failure." -Color Red
                    Write-ColorText "Press any key to continue..." -Color Gray
                    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                }
            }
            "4" {
                Show-PreviousReports -TenantInfo $tenantInfo -Directories $directories
            }
            "5" {
                Write-Header "GOODBYE"
                Write-ColorText "Thank you for using CA Analysis Manager!" -Color Green
                Write-ColorText "Session completed for tenant: $($tenantInfo.Domain)" -Color Gray
                break
            }
            default {
                Write-ColorText "`nInvalid selection. Please choose 1-5." -Color Red
                Write-ColorText "Press any key to continue..." -Color Gray
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            }
        }
    } while ($choice -ne "5")
}
catch {
    Write-ColorText "`nCRITICAL ERROR: $($_.Exception.Message)" -Color Red
    Write-ColorText "Press any key to exit..." -Color Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}
