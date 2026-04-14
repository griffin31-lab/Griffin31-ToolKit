# ==============================
# Entra ID Stale Devices Cleanup
# Identifies, audits, disables, and deletes stale devices
# based on last sign-in activity.
#
# Single-file script — no Python dependency.
# Supports PowerShell 7.x on Windows and macOS.
# ==============================

# ── Prerequisite checks ──
$ErrorActionPreference = "Stop"

# 1. PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Host ""
  Write-Host "  [!] This script requires PowerShell 7 or later." -ForegroundColor Red
  Write-Host ""
  if ($IsWindows -or $env:OS -match "Windows") {
    Write-Host "  Install it from: https://aka.ms/install-powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh .\stale-devices-cleanup.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./stale-devices-cleanup.ps1" -ForegroundColor Yellow
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
}

# 3. ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
  Write-Host ""
  Write-Host "  [!] ImportExcel module is not installed." -ForegroundColor Red
  Write-Host ""
  $installChoice = Read-Host "  Would you like to install it now? (Y/n)"
  if ($installChoice -eq 'n' -or $installChoice -eq 'N') {
    Write-Host "  Cannot continue without ImportExcel. Exiting." -ForegroundColor Red
    return
  }
  Write-Host "  Installing ImportExcel..." -ForegroundColor Yellow
  Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
  Write-Host "  ImportExcel installed successfully." -ForegroundColor Green
}

Import-Module ImportExcel -ErrorAction Stop

# ── Helper: Write-ProgressBar ──
function Write-ProgressBar {
  param(
    [int]$Current,
    [int]$Total,
    [string]$Activity = "Processing",
    [string]$Status   = ""
  )
  if ($Total -le 0) { return }
  $pct = [math]::Min(100, [math]::Round(($Current / $Total) * 100))
  $elapsed = if ($script:swPhase) { $script:swPhase.Elapsed.ToString("mm\:ss") } else { "--:--" }
  Write-Progress -Activity $Activity -Status "$pct%  ($Current/$Total)  $Status  [$elapsed]" -PercentComplete $pct
}

# ── Interactive banner ──
$banner = @"

  ============================================================
     Entra ID Stale Devices Cleanup
     Audit, Disable & Delete inactive devices
  ============================================================

"@
Write-Host $banner -ForegroundColor Cyan

# ── Step 1: Collect UPN ──
Write-Host "  This script connects via delegated Graph permissions." -ForegroundColor Gray
Write-Host "  You will be prompted to sign in with a browser." -ForegroundColor Gray
Write-Host ""

$AdminUPN = Read-Host "  Enter admin UPN (e.g. admin@contoso.onmicrosoft.com)"
if (-not $AdminUPN -or $AdminUPN.Trim().Length -eq 0) {
  Write-Host "  [!] No UPN provided. Exiting." -ForegroundColor Red
  return
}
$AdminUPN = $AdminUPN.Trim()

if ($AdminUPN -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
  Write-Host "  [!] '$AdminUPN' doesn't look like a valid UPN." -ForegroundColor Red
  return
}

$upnDomain = ($AdminUPN -split '@')[-1]
Write-Host ""
Write-Host "  Tenant domain detected: " -NoNewline; Write-Host $upnDomain -ForegroundColor Yellow

# ── Step 2: Stale threshold ──
Write-Host ""
Write-Host "  Select stale device threshold:" -ForegroundColor White
Write-Host "    [1]  30 days   (default)" -ForegroundColor Green
Write-Host "    [2]  60 days" -ForegroundColor Yellow
Write-Host "    [3]  90 days" -ForegroundColor Yellow
Write-Host "    [4]  Custom" -ForegroundColor Yellow
Write-Host ""
$thresholdChoice = Read-Host "  Enter choice (1-4) [default: 1]"

switch ($thresholdChoice) {
  "2" { $StaleDays = 60 }
  "3" { $StaleDays = 90 }
  "4" {
    $customDays = Read-Host "  Enter number of days"
    if ($customDays -match '^\d+$' -and [int]$customDays -gt 0) {
      $StaleDays = [int]$customDays
    } else {
      Write-Host "  [!] Invalid number. Using 30 days." -ForegroundColor Red
      $StaleDays = 30
    }
  }
  default { $StaleDays = 30 }
}
Write-Host "  -> Threshold: $StaleDays days" -ForegroundColor Green

# ── Step 3: OS filter ──
Write-Host ""
Write-Host "  Filter by operating system:" -ForegroundColor White
Write-Host "    [1]  All OS types (default)" -ForegroundColor Green
Write-Host "    [2]  Windows only" -ForegroundColor Yellow
Write-Host "    [3]  macOS only" -ForegroundColor Yellow
Write-Host "    [4]  Linux only" -ForegroundColor Yellow
Write-Host "    [5]  iOS only" -ForegroundColor Yellow
Write-Host "    [6]  Android only" -ForegroundColor Yellow
Write-Host "    [7]  Custom (enter OS name)" -ForegroundColor Yellow
Write-Host ""
$osChoice = Read-Host "  Enter choice (1-7) [default: 1]"

switch ($osChoice) {
  "2" { $OSFilter = "Windows" }
  "3" { $OSFilter = "macOS" }
  "4" { $OSFilter = "Linux" }
  "5" { $OSFilter = "iOS" }
  "6" { $OSFilter = "Android" }
  "7" {
    $OSFilter = Read-Host "  Enter OS name (e.g. ChromeOS)"
    if (-not $OSFilter -or $OSFilter.Trim().Length -eq 0) {
      Write-Host "  [!] No OS entered. Using All." -ForegroundColor Red
      $OSFilter = $null
    } else {
      $OSFilter = $OSFilter.Trim()
    }
  }
  default { $OSFilter = $null }
}

if ($OSFilter) {
  Write-Host "  -> OS filter: $OSFilter" -ForegroundColor Green
} else {
  Write-Host "  -> OS filter: All" -ForegroundColor Green
}

# ── Summary & confirm ──
$ActionMode = "Audit"
Write-Host ""
Write-Host "  ──────────────────────────────────────" -ForegroundColor Gray
Write-Host "  Threshold:  $StaleDays days" -ForegroundColor White
Write-Host "  OS filter:  $(if ($OSFilter) { $OSFilter } else { 'All' })" -ForegroundColor White
Write-Host "  Flow:       Audit -> Review -> Optionally disable/delete" -ForegroundColor Gray
Write-Host "  ──────────────────────────────────────" -ForegroundColor Gray
Write-Host ""

$confirm = Read-Host "  Press ENTER to start or type 'q' to quit"
if ($confirm -and $confirm.Trim().ToLower() -ne '') {
  Write-Host "  Cancelled." -ForegroundColor Red
  return
}

# ── Overall timer ──
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$cutoffDate = (Get-Date).AddDays(-$StaleDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# ── Phase 1: Connect ──
# Request ReadWrite upfront so we don't need to reconnect if user chooses to act
$totalPhases = 4
$phase = 1

Write-Host ""
Write-Host "  [$phase/$totalPhases] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$requiredScopes = @(
  "Device.Read.All",
  "Device.ReadWrite.All",
  "Directory.Read.All"
)

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$needsConnect = $true

if ($ctx -and $ctx.TenantId) {
  $currentScopes = @($ctx.Scopes)
  $missing = $requiredScopes | Where-Object { $_ -notin $currentScopes }
  if ($missing.Count -eq 0) {
    Write-Host "  Reusing existing Graph session ($($ctx.Account))" -ForegroundColor Green
    $needsConnect = $false
  } else {
    Write-Host "  Existing session missing scopes: $($missing -join ', ')" -ForegroundColor Yellow
    Write-Host "  Reconnecting..." -ForegroundColor Yellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
  }
}

if ($needsConnect) {
  Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
  $ctx = Get-MgContext
}

Write-Host "  Connected to tenant: $($ctx.TenantId)  ($($ctx.Account))" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 2: Fetch all devices ──
$phase++
Write-Host "  [$phase/$totalPhases] Fetching devices from Entra ID..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$allDevices = @()
$deviceUri = "https://graph.microsoft.com/v1.0/devices?`$select=id,displayName,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime,accountEnabled,deviceId,registrationDateTime,trustType,managementType,mdmAppId,profileType&`$top=999"

while ($deviceUri) {
  $resp = Invoke-MgGraphRequest -Method GET -Uri $deviceUri -ErrorAction Stop
  if ($resp.value) {
    $allDevices += $resp.value
  }
  $deviceUri = $resp.'@odata.nextLink'
  if ($deviceUri) {
    Write-ProgressBar -Current $allDevices.Count -Total ($allDevices.Count + 999) -Activity "Fetching devices" -Status "$($allDevices.Count) fetched"
  }
}

Write-Progress -Activity "Fetching devices" -Completed
Write-Host "  Total devices fetched: $($allDevices.Count)" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 3: Filter stale devices ──
$phase++
Write-Host "  [$phase/$totalPhases] Filtering stale devices (inactive $StaleDays+ days)..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$now = Get-Date
$staleDevices = @()

foreach ($device in $allDevices) {
  $lastSignIn = $device.approximateLastSignInDateTime
  $isStale = $false

  if (-not $lastSignIn) {
    # Never signed in — check registration date
    $regDate = $device.registrationDateTime
    if ($regDate) {
      $regParsed = [datetime]::Parse($regDate, [System.Globalization.CultureInfo]::InvariantCulture)
      if (($now - $regParsed).TotalDays -ge $StaleDays) {
        $isStale = $true
        $daysSince = [math]::Round(($now - $regParsed).TotalDays)
        $lastActivity = "Never (registered $($regParsed.ToString('yyyy-MM-dd')))"
      }
    } else {
      $isStale = $true
      $daysSince = 9999
      $lastActivity = "Never (no date available)"
    }
  } else {
    $signInParsed = [datetime]::Parse($lastSignIn, [System.Globalization.CultureInfo]::InvariantCulture)
    $daysSince = [math]::Round(($now - $signInParsed).TotalDays)
    if ($daysSince -ge $StaleDays) {
      $isStale = $true
      $lastActivity = $signInParsed.ToString('yyyy-MM-dd')
    }
  }

  if (-not $isStale) { continue }

  # OS filter
  $deviceOS = if ($device.operatingSystem) { $device.operatingSystem } else { "Unknown" }
  if ($OSFilter -and $deviceOS -notlike "*$OSFilter*") { continue }

  $regDateFormatted = "N/A"
  if ($device.registrationDateTime) {
    try {
      $regDateFormatted = ([datetime]::Parse($device.registrationDateTime, [System.Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
    } catch { $regDateFormatted = "N/A" }
  }

  # Resolve MDM name from well-known app IDs
  $mdmName = "N/A"
  if ($device.mdmAppId) {
    $mdmName = switch ($device.mdmAppId) {
      "0000000a-0000-0000-c000-000000000000" { "Intune" }
      "54b943f8-d761-4f8d-951e-9cea1846db5a" { "Jamf Pro" }
      "in6877d1-4354-47e5-b381-8c76f9a1a903" { "VMware Workspace ONE" }
      default { $device.mdmAppId }
    }
  }

  # Join type (trustType mapping)
  $joinType = switch ($device.trustType) {
    "AzureAd"       { "Azure AD Joined" }
    "Hybrid"        { "Hybrid Azure AD Joined" }
    "ServerAd"      { "Azure AD Registered" }
    "Workplace"     { "Workplace Joined" }
    default         { if ($device.trustType) { $device.trustType } else { "N/A" } }
  }

  $staleDevices += [PSCustomObject]@{
    DeviceId       = $device.id
    DisplayName    = $device.displayName
    OS             = $deviceOS
    OSVersion      = if ($device.operatingSystemVersion) { $device.operatingSystemVersion } else { "N/A" }
    RegisteredDate = $regDateFormatted
    LastSignIn     = $lastActivity
    DaysInactive   = $daysSince
    Enabled        = $device.accountEnabled
    JoinType       = $joinType
    MDM            = $mdmName
    Owner          = ""
    ManagementType = if ($device.managementType) { $device.managementType } else { "N/A" }
  }
}

# Sort by days inactive descending
$staleDevices = $staleDevices | Sort-Object -Property DaysInactive -Descending

Write-Host "  Stale devices found: $($staleDevices.Count)" -ForegroundColor $(if ($staleDevices.Count -gt 0) { "Yellow" } else { "Green" })

# OS breakdown
$osGroups = $staleDevices | Group-Object -Property OS | Sort-Object Count -Descending
if ($osGroups.Count -gt 0) {
  Write-Host ""
  Write-Host "  OS breakdown:" -ForegroundColor White
  foreach ($g in $osGroups) {
    Write-Host "    $($g.Name): $($g.Count)" -ForegroundColor Gray
  }
}

# Enabled/disabled breakdown
$enabledCount  = @($staleDevices | Where-Object { $_.Enabled -eq $true }).Count
$disabledCount = @($staleDevices | Where-Object { $_.Enabled -eq $false }).Count
Write-Host ""
Write-Host "  Status: $enabledCount enabled, $disabledCount already disabled" -ForegroundColor Gray

$script:swPhase.Stop()

# ── Phase 4: Enrich with device owners ──
if ($staleDevices.Count -gt 0) {
  $phase++
  Write-Host ""
  Write-Host "  [$phase/$totalPhases] Fetching device owners..." -ForegroundColor Cyan
  $script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

  for ($i = 0; $i -lt $staleDevices.Count; $i++) {
    $d = $staleDevices[$i]
    Write-ProgressBar -Current ($i + 1) -Total $staleDevices.Count -Activity "Fetching owners" -Status $d.DisplayName

    try {
      $owners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$($d.DeviceId)/registeredOwners?`$select=displayName,userPrincipalName" -ErrorAction Stop
      if ($owners.value -and $owners.value.Count -gt 0) {
        $ownerNames = ($owners.value | ForEach-Object {
          if ($_.displayName) { $_.displayName } elseif ($_.userPrincipalName) { $_.userPrincipalName } else { "Unknown" }
        }) -join ", "
        $d.Owner = $ownerNames
      } else {
        $d.Owner = "No owner"
      }
    } catch {
      $d.Owner = "N/A"
    }
  }

  Write-Progress -Activity "Fetching owners" -Completed
  Write-Host "  Owners resolved for $($staleDevices.Count) devices" -ForegroundColor Green
  $script:swPhase.Stop()
}

if ($staleDevices.Count -eq 0) {
  Write-Host ""
  Write-Host "  No stale devices found matching your criteria. Nothing to do." -ForegroundColor Green
  $swTotal.Stop()
  Write-Host ""
  Write-Host "  Total time: $($swTotal.Elapsed.ToString('mm\:ss'))" -ForegroundColor Gray
  return
}

# ── Post-audit: Ask what to do ──
$actionResults = @()

Write-Host ""
Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "  Audit complete. What would you like to do?" -ForegroundColor White
Write-Host ""
Write-Host "    [1]  Export report only — no changes (default)" -ForegroundColor Green
Write-Host "    [2]  Disable stale devices, then export" -ForegroundColor Yellow
Write-Host "    [3]  Delete stale devices, then export" -ForegroundColor Red
Write-Host ""
$postChoice = Read-Host "  Enter choice (1-3) [default: 1]"

switch ($postChoice) {
  "2" { $ActionMode = "Disable" }
  "3" { $ActionMode = "Delete" }
  default { $ActionMode = "Audit" }
}

if ($ActionMode -ne "Audit") {
  # Show preview of affected devices
  $targetDevices = if ($ActionMode -eq "Disable") {
    @($staleDevices | Where-Object { $_.Enabled -eq $true })
  } else {
    $staleDevices
  }

  if ($targetDevices.Count -eq 0) {
    if ($ActionMode -eq "Disable") {
      Write-Host ""
      Write-Host "  All stale devices are already disabled. Exporting report only." -ForegroundColor Green
    } else {
      Write-Host ""
      Write-Host "  No devices to delete. Exporting report only." -ForegroundColor Green
    }
    $ActionMode = "Audit"
  } else {
    Write-Host ""
    Write-Host "  The following $($targetDevices.Count) device(s) will be $($ActionMode.ToLower())d:" -ForegroundColor White
    Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Gray

    $preview = $targetDevices | Select-Object -First 20
    foreach ($d in $preview) {
      $statusTag = if ($d.Enabled) { "[Enabled]" } else { "[Disabled]" }
      Write-Host "    $($d.DisplayName)  |  $($d.OS)  |  $($d.DaysInactive) days  |  $statusTag" -ForegroundColor Gray
    }
    if ($targetDevices.Count -gt 20) {
      Write-Host "    ... and $($targetDevices.Count - 20) more" -ForegroundColor Gray
    }

    Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host ""

    $confirmAction = Read-Host "  Type 'YES' to confirm $($ActionMode.ToLower()) of $($targetDevices.Count) device(s)"
    if ($confirmAction -ne 'YES') {
      Write-Host "  Action cancelled. Exporting report only." -ForegroundColor Yellow
      $ActionMode = "Audit"
    } else {
      Write-Host ""
      Write-Host "  $($ActionMode) in progress..." -ForegroundColor $(if ($ActionMode -eq "Delete") { "Red" } else { "Yellow" })
      $script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()
      $successCount = 0
      $failCount = 0

      for ($i = 0; $i -lt $targetDevices.Count; $i++) {
        $d = $targetDevices[$i]
        Write-ProgressBar -Current ($i + 1) -Total $targetDevices.Count -Activity "$ActionMode devices" -Status $d.DisplayName

        try {
          if ($ActionMode -eq "Disable") {
            $body = @{ accountEnabled = $false } | ConvertTo-Json
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/devices/$($d.DeviceId)" -Body $body -ContentType "application/json" -ErrorAction Stop
          } elseif ($ActionMode -eq "Delete") {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/devices/$($d.DeviceId)" -ErrorAction Stop
          }
          $successCount++
          $actionResults += [PSCustomObject]@{
            DeviceName = $d.DisplayName
            DeviceId   = $d.DeviceId
            Action     = $ActionMode
            Result     = "Success"
          }
        } catch {
          $failCount++
          $actionResults += [PSCustomObject]@{
            DeviceName = $d.DisplayName
            DeviceId   = $d.DeviceId
            Action     = $ActionMode
            Result     = "Failed: $($_.Exception.Message)"
          }
        }
      }

      Write-Progress -Activity "$ActionMode devices" -Completed
      Write-Host ""
      Write-Host "  $ActionMode complete: $successCount succeeded, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
      $script:swPhase.Stop()
    }
  }
}

# ── Export to Excel ──
Write-Host ""
Write-Host "  Generating Excel report..." -ForegroundColor Cyan

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$desktopPath = [Environment]::GetFolderPath("Desktop")
$exportPath = Join-Path $desktopPath "StaleDevices_Report_$timestamp.xlsx"

# ── Excel color helpers ──
function HexColor([string]$hex) {
  $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
  $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
  $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
  return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

$NAVY           = HexColor "1B2A4A"
$DARK_BLUE      = HexColor "2C3E6B"
$ACCENT_BLUE    = HexColor "4472C4"
$LIGHT_BLUE     = HexColor "D6E4F0"
$WHITE_C        = [System.Drawing.Color]::White
$MED_GRAY       = HexColor "D9D9D9"
$DARK_GRAY      = HexColor "404040"
$RED_BG         = HexColor "FDE8E8"
$RED_TEXT        = HexColor "B91C1C"
$AMBER_BG       = HexColor "FEF3C7"
$AMBER_TEXT      = HexColor "92400E"
$GREEN_BG       = HexColor "D1FAE5"
$GREEN_TEXT      = HexColor "065F46"

function Set-Fill($cells, $color) {
  $cells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
  $cells.Style.Fill.BackgroundColor.SetColor($color)
}

function Set-Font($cells, [int]$size = 10, [bool]$bold = $false, $color = $null) {
  $cells.Style.Font.Size = $size
  $cells.Style.Font.Bold = $bold
  if ($color) { $cells.Style.Font.Color.SetColor($color) }
}

function Set-ThinBorder($cells, $color = $null) {
  if ($null -eq $color) { $color = [System.Drawing.Color]::FromArgb(217, 217, 217) }
  $b = $cells.Style.Border
  $b.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Top.Color.SetColor($color)
  $b.Bottom.Color.SetColor($color)
  $b.Left.Color.SetColor($color)
  $b.Right.Color.SetColor($color)
}

# ── Build workbook ──
$excel = New-Object OfficeOpenXml.ExcelPackage

# ── Sheet 1: Summary Dashboard ──
$ws = $excel.Workbook.Worksheets.Add("Summary")
$ws.TabColor = HexColor "2E75B6"
$ws.View.ShowGridLines = $false

# Header
$ws.Column(1).Width = 2
$ws.Column(2).Width = 25
$ws.Column(3).Width = 20
$ws.Column(4).Width = 20
$ws.Column(5).Width = 20
$ws.Column(6).Width = 20

# Title bar
$ws.Row(1).Height = 8
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[1, $c] $NAVY }
$ws.Row(2).Height = 36
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[2, $c] $NAVY }
$ws.Cells["B2:F2"].Merge = $true
$ws.Cells["B2"].Value = "STALE DEVICES REPORT"
Set-Font $ws.Cells["B2"] -size 18 -bold $true -color $WHITE_C
$ws.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Subtitle
$ws.Row(3).Height = 24
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[3, $c] $DARK_BLUE }
$ws.Cells["B3:F3"].Merge = $true
$ws.Cells["B3"].Value = "Tenant: $($ctx.TenantId)  |  Threshold: $StaleDays days  |  OS: $(if ($OSFilter) { $OSFilter } else { 'All' })  |  Action: $ActionMode  |  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Set-Font $ws.Cells["B3"] -size 10 -color (HexColor "B0C4DE")

# KPI row
$ws.Row(4).Height = 8
$ws.Row(5).Height = 60
$kpis = @(
  @{ Col = 2; Number = "$($allDevices.Count)"; Label = "TOTAL DEVICES" },
  @{ Col = 3; Number = "$($staleDevices.Count)"; Label = "STALE DEVICES" },
  @{ Col = 4; Number = "$enabledCount"; Label = "ENABLED (STALE)" },
  @{ Col = 5; Number = "$disabledCount"; Label = "ALREADY DISABLED" }
)

foreach ($kpi in $kpis) {
  $cell = $ws.Cells[5, $kpi.Col]
  $cell.IsRichText = $true
  $rt = $cell.RichText
  $num = $rt.Add($kpi.Number + "`n")
  $num.Size = 22
  $num.Bold = $true
  $num.Color = $NAVY
  $lbl = $rt.Add($kpi.Label)
  $lbl.Size = 9
  $lbl.Bold = $false
  $lbl.Color = $DARK_GRAY
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $cell.Style.WrapText = $true
  Set-Fill $cell $LIGHT_BLUE
  Set-ThinBorder $cell
}

# OS breakdown section
$row = 7
$ws.Cells["B$row"].Value = "OS BREAKDOWN"
Set-Font $ws.Cells["B$row"] -size 12 -bold $true -color $NAVY
$row++

$ws.Cells["B$row"].Value = "Operating System"
$ws.Cells["C$row"].Value = "Stale Devices"
$ws.Cells["D$row"].Value = "% of Total Stale"
Set-Font $ws.Cells["B$row"] -size 10 -bold $true -color $WHITE_C
Set-Font $ws.Cells["C$row"] -size 10 -bold $true -color $WHITE_C
Set-Font $ws.Cells["D$row"] -size 10 -bold $true -color $WHITE_C
Set-Fill $ws.Cells["B$row"] $ACCENT_BLUE
Set-Fill $ws.Cells["C$row"] $ACCENT_BLUE
Set-Fill $ws.Cells["D$row"] $ACCENT_BLUE
$row++

foreach ($g in $osGroups) {
  $ws.Cells["B$row"].Value = $g.Name
  $ws.Cells["C$row"].Value = $g.Count
  $pctVal = if ($staleDevices.Count -gt 0) { [math]::Round(($g.Count / $staleDevices.Count) * 100, 1) } else { 0 }
  $ws.Cells["D$row"].Value = "$pctVal%"
  Set-ThinBorder $ws.Cells["B$row"]
  Set-ThinBorder $ws.Cells["C$row"]
  Set-ThinBorder $ws.Cells["D$row"]
  $ws.Cells["C$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells["D$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $row++
}

# ── Sheet 2: Device Details ──
$ws2 = $excel.Workbook.Worksheets.Add("Stale Devices")
$ws2.TabColor = HexColor "E67E22"
$ws2.View.ShowGridLines = $false

# Title
$ws2.Column(1).Width = 2
$ws2.Column(2).Width = 28
$ws2.Column(3).Width = 13
$ws2.Column(4).Width = 16
$ws2.Column(5).Width = 14
$ws2.Column(6).Width = 16
$ws2.Column(7).Width = 14
$ws2.Column(8).Width = 10
$ws2.Column(9).Width = 25
$ws2.Column(10).Width = 22
$ws2.Column(11).Width = 14
$ws2.Column(12).Width = 16

$ws2.Row(1).Height = 8
for ($c = 1; $c -le 12; $c++) { Set-Fill $ws2.Cells[1, $c] $NAVY }
$ws2.Row(2).Height = 32
for ($c = 1; $c -le 12; $c++) { Set-Fill $ws2.Cells[2, $c] $NAVY }
$ws2.Cells["B2:L2"].Merge = $true
$ws2.Cells["B2"].Value = "STALE DEVICES — DETAILED LIST"
Set-Font $ws2.Cells["B2"] -size 16 -bold $true -color $WHITE_C
$ws2.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Headers
$hdrRow = 4
$headers = @(
  @{ Col = 2; Name = "Device Name" },
  @{ Col = 3; Name = "OS" },
  @{ Col = 4; Name = "OS Version" },
  @{ Col = 5; Name = "Registered" },
  @{ Col = 6; Name = "Last Sign-In" },
  @{ Col = 7; Name = "Days Inactive" },
  @{ Col = 8; Name = "Enabled" },
  @{ Col = 9; Name = "Owner" },
  @{ Col = 10; Name = "Join Type" },
  @{ Col = 11; Name = "MDM" },
  @{ Col = 12; Name = "Management" }
)

$ws2.Row($hdrRow).Height = 28
foreach ($h in $headers) {
  $cell = $ws2.Cells[$hdrRow, $h.Col]
  $cell.Value = $h.Name
  Set-Fill $cell $ACCENT_BLUE
  Set-Font $cell -size 10 -bold $true -color $WHITE_C
  $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
}

# Data rows
$dataRow = $hdrRow + 1
foreach ($d in $staleDevices) {
  $ws2.Row($dataRow).Height = 22

  $ws2.Cells[$dataRow, 2].Value = $d.DisplayName
  $ws2.Cells[$dataRow, 3].Value = $d.OS
  $ws2.Cells[$dataRow, 4].Value = $d.OSVersion
  $ws2.Cells[$dataRow, 5].Value = $d.RegisteredDate
  $ws2.Cells[$dataRow, 6].Value = $d.LastSignIn
  $ws2.Cells[$dataRow, 7].Value = $d.DaysInactive
  $ws2.Cells[$dataRow, 8].Value = if ($d.Enabled) { "Yes" } else { "No" }
  $ws2.Cells[$dataRow, 9].Value = $d.Owner
  $ws2.Cells[$dataRow, 10].Value = $d.JoinType
  $ws2.Cells[$dataRow, 11].Value = $d.MDM
  $ws2.Cells[$dataRow, 12].Value = $d.ManagementType

  # Color enabled/disabled
  $enabledCell = $ws2.Cells[$dataRow, 8]
  if ($d.Enabled) {
    Set-Fill $enabledCell $AMBER_BG
    Set-Font $enabledCell -color $AMBER_TEXT
  } else {
    Set-Fill $enabledCell $GREEN_BG
    Set-Font $enabledCell -color $GREEN_TEXT
  }

  # Color high inactivity
  $daysCell = $ws2.Cells[$dataRow, 7]
  $daysCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  if ($d.DaysInactive -ge 180) {
    Set-Fill $daysCell $RED_BG
    Set-Font $daysCell -color $RED_TEXT
  } elseif ($d.DaysInactive -ge 90) {
    Set-Fill $daysCell $AMBER_BG
    Set-Font $daysCell -color $AMBER_TEXT
  }

  # Borders
  for ($c = 2; $c -le 12; $c++) {
    Set-ThinBorder $ws2.Cells[$dataRow, $c]
  }

  # Alternating row color
  if ($dataRow % 2 -eq 0) {
    for ($c = 2; $c -le 12; $c++) {
      if ($c -ne 7 -and $c -ne 8) {
        Set-Fill $ws2.Cells[$dataRow, $c] (HexColor "F8F9FA")
      }
    }
  }

  $dataRow++
}

# ── Sheet 3: Action Log (if actions were taken) ──
if ($actionResults.Count -gt 0) {
  $ws3 = $excel.Workbook.Worksheets.Add("Action Log")
  $ws3.TabColor = HexColor "E74C3C"
  $ws3.View.ShowGridLines = $false

  $ws3.Column(1).Width = 2
  $ws3.Column(2).Width = 30
  $ws3.Column(3).Width = 40
  $ws3.Column(4).Width = 14
  $ws3.Column(5).Width = 40

  $ws3.Row(1).Height = 8
  for ($c = 1; $c -le 5; $c++) { Set-Fill $ws3.Cells[1, $c] $NAVY }
  $ws3.Row(2).Height = 32
  for ($c = 1; $c -le 5; $c++) { Set-Fill $ws3.Cells[2, $c] $NAVY }
  $ws3.Cells["B2:E2"].Merge = $true
  $ws3.Cells["B2"].Value = "ACTION LOG — $($ActionMode.ToUpper()) RESULTS"
  Set-Font $ws3.Cells["B2"] -size 16 -bold $true -color $WHITE_C

  # Headers
  $logHdr = 4
  $ws3.Row($logHdr).Height = 28
  $logHeaders = @(
    @{ Col = 2; Name = "Device Name" },
    @{ Col = 3; Name = "Device ID" },
    @{ Col = 4; Name = "Action" },
    @{ Col = 5; Name = "Result" }
  )
  foreach ($h in $logHeaders) {
    $cell = $ws3.Cells[$logHdr, $h.Col]
    $cell.Value = $h.Name
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
  }

  $logRow = $logHdr + 1
  foreach ($r in $actionResults) {
    $ws3.Cells[$logRow, 2].Value = $r.DeviceName
    $ws3.Cells[$logRow, 3].Value = $r.DeviceId
    $ws3.Cells[$logRow, 4].Value = $r.Action
    $ws3.Cells[$logRow, 5].Value = $r.Result

    $resultCell = $ws3.Cells[$logRow, 5]
    if ($r.Result -eq "Success") {
      Set-Fill $resultCell $GREEN_BG
      Set-Font $resultCell -color $GREEN_TEXT
    } else {
      Set-Fill $resultCell $RED_BG
      Set-Font $resultCell -color $RED_TEXT
    }

    for ($c = 2; $c -le 5; $c++) {
      Set-ThinBorder $ws3.Cells[$logRow, $c]
    }
    $logRow++
  }
}

# ── Save ──
$excel.SaveAs($exportPath)
$excel.Dispose()

$swTotal.Stop()

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  Report saved: $exportPath" -ForegroundColor Green
Write-Host "  Total devices scanned: $($allDevices.Count)" -ForegroundColor White
Write-Host "  Stale devices found:   $($staleDevices.Count)" -ForegroundColor White
if ($actionResults.Count -gt 0) {
  $successTotal = @($actionResults | Where-Object { $_.Result -eq "Success" }).Count
  $failTotal    = @($actionResults | Where-Object { $_.Result -ne "Success" }).Count
  Write-Host "  Action ($ActionMode):       $successTotal succeeded, $failTotal failed" -ForegroundColor White
}
Write-Host "  Total time: $($swTotal.Elapsed.ToString('mm\:ss'))" -ForegroundColor Gray
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""
