# ==============================
# Entra ID Stale App Registrations Cleanup
# Identifies, audits, disables, and deletes app registrations
# based on service principal sign-in activity.
#
# Uses /beta/reports/servicePrincipalSignInActivities — the
# aggregated last-sign-in report, which persists far beyond the
# 30-day raw sign-in log retention.
#
# Single-file script. Supports PowerShell 7.x on Windows and macOS.
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
    Write-Host "  Then run:  pwsh .\stale-apps-cleanup.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./stale-apps-cleanup.ps1" -ForegroundColor Yellow
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
     Entra ID Stale App Registrations Cleanup
     Audit, Disable & Delete inactive app registrations
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

if ($AdminUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9][a-zA-Z0-9.-]{0,252}\.[a-zA-Z]{2,63}$') {
  Write-Host "  [!] '$AdminUPN' doesn't look like a valid UPN." -ForegroundColor Red
  return
}

$upnDomain = ($AdminUPN -split '@')[-1]
Write-Host ""
Write-Host "  Tenant domain detected: " -NoNewline; Write-Host $upnDomain -ForegroundColor Yellow

# ── Step 2: Stale threshold ──
Write-Host ""
Write-Host "  Select stale app threshold:" -ForegroundColor White
Write-Host "    [1]  90 days   (default, recommended)" -ForegroundColor Green
Write-Host "    [2]  30 days" -ForegroundColor Yellow
Write-Host "    [3]  60 days" -ForegroundColor Yellow
Write-Host "    [4]  180 days" -ForegroundColor Yellow
Write-Host "    [5]  Custom" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Note: the aggregated sign-in report retains last-activity" -ForegroundColor DarkGray
Write-Host "        timestamps beyond the 30-day raw log retention window." -ForegroundColor DarkGray
Write-Host ""
$thresholdChoice = Read-Host "  Enter choice (1-5) [default: 1]"

switch ($thresholdChoice) {
  "2" { $StaleDays = 30 }
  "3" { $StaleDays = 60 }
  "4" { $StaleDays = 180 }
  "5" {
    $customDays = Read-Host "  Enter number of days"
    if ($customDays -match '^\d+$' -and [int]$customDays -gt 0) {
      $StaleDays = [int]$customDays
    } else {
      Write-Host "  [!] Invalid number. Using 90 days." -ForegroundColor Red
      $StaleDays = 90
    }
  }
  default { $StaleDays = 90 }
}
Write-Host "  -> Threshold: $StaleDays days" -ForegroundColor Green

# ── Step 3: Include / exclude Microsoft first-party apps ──
Write-Host ""
Write-Host "  Include Microsoft-published (first-party) apps?" -ForegroundColor White
Write-Host "    [1]  Exclude first-party apps (default)" -ForegroundColor Green
Write-Host "    [2]  Include all" -ForegroundColor Yellow
Write-Host ""
$msChoice = Read-Host "  Enter choice (1-2) [default: 1]"
$IncludeFirstParty = ($msChoice -eq "2")
if ($IncludeFirstParty) {
  Write-Host "  -> Including Microsoft first-party apps" -ForegroundColor Green
} else {
  Write-Host "  -> Excluding Microsoft first-party apps" -ForegroundColor Green
}

# ── Summary & confirm ──
Write-Host ""
Write-Host "  --------------------------------------" -ForegroundColor Gray
Write-Host "  Threshold:        $StaleDays days" -ForegroundColor White
Write-Host "  First-party apps: $(if ($IncludeFirstParty) { 'Included' } else { 'Excluded' })" -ForegroundColor White
Write-Host "  Flow:             Audit -> Review -> Optionally disable/delete" -ForegroundColor Gray
Write-Host "  --------------------------------------" -ForegroundColor Gray
Write-Host ""

$confirm = Read-Host "  Press ENTER to start or type 'q' to quit"
if ($confirm -and $confirm.Trim().ToLower() -ne '') {
  Write-Host "  Cancelled." -ForegroundColor Red
  return
}

# ── Overall timer ──
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$cutoffDate = (Get-Date).AddDays(-$StaleDays).ToUniversalTime()

# ── Phase 1: Connect ──
# Request ReadWrite upfront so we don't reconnect if the user chooses to act
$totalPhases = 4
$phase = 1

Write-Host ""
Write-Host "  [$phase/$totalPhases] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$requiredScopes = @(
  "Application.ReadWrite.All",
  "Directory.Read.All",
  "AuditLog.Read.All"
)

Write-Host "  A browser window will open for sign-in." -ForegroundColor Gray
Write-Host "  If it doesn't, look for a device-code URL below." -ForegroundColor DarkGray
Write-Host ""

try {
  Connect-MgGraph -Scopes $requiredScopes -NoWelcome
  $ctx = Get-MgContext
  if (-not $ctx) {
    Write-Host "  [!] Graph context not established." -ForegroundColor Red
    return
  }
  Write-Host ""
  Write-Host "  Connected as: " -NoNewline; Write-Host $ctx.Account -ForegroundColor Green
  Write-Host "  Tenant:       " -NoNewline; Write-Host $ctx.TenantId -ForegroundColor Green
} catch {
  Write-Host "  [!] Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
  return
}

# ── Phase 2: Fetch app registrations, service principals, sign-in activity ──
$phase = 2
Write-Host ""
Write-Host "  [$phase/$totalPhases] Fetching app registrations and sign-in activity..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Graph pagination helper
function Invoke-GraphPaged {
  param([string]$Uri)
  $all = @()
  $next = $Uri
  while ($next) {
    try {
      $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
    } catch {
      Write-Host "  [!] Graph request failed: $($_.Exception.Message)" -ForegroundColor Red
      throw
    }
    if ($resp.value) { $all += $resp.value }
    $next = $resp.'@odata.nextLink'
  }
  return $all
}

# (a) App registrations (applications in this tenant)
Write-Host "    - App registrations..." -ForegroundColor Gray
$appRegs = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/applications?`$top=999&`$select=id,appId,displayName,createdDateTime,publisherDomain,signInAudience,keyCredentials,passwordCredentials,web,tags"
Write-Host "    -> $($appRegs.Count) app registrations" -ForegroundColor Green

# (b) Service principals (for appId -> sp mapping + accountEnabled + owner lookup)
Write-Host "    - Service principals..." -ForegroundColor Gray
$sps = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/servicePrincipals?`$top=999&`$select=id,appId,displayName,accountEnabled,servicePrincipalType,appOwnerOrganizationId,tags"
Write-Host "    -> $($sps.Count) service principals" -ForegroundColor Green

# Build appId -> SP lookup
$spByAppId = @{}
foreach ($sp in $sps) { if ($sp.appId) { $spByAppId[$sp.appId] = $sp } }

# Microsoft's first-party tenant (Microsoft Services)
$msFirstPartyTenantId = "f8cdef31-a31e-4b4a-93e4-5f571e91255a"

# (c) Sign-in activity per SP (beta endpoint)
Write-Host "    - Service principal sign-in activity (beta)..." -ForegroundColor Gray
try {
  $signInActivity = Invoke-GraphPaged "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?`$top=999"
  Write-Host "    -> $($signInActivity.Count) sign-in activity records" -ForegroundColor Green
} catch {
  Write-Host "  [!] Could not fetch sign-in activity report. Tenant may lack Entra ID P1/P2." -ForegroundColor Red
  Write-Host "      Continuing without sign-in data (all apps will be flagged as 'unknown')." -ForegroundColor Yellow
  $signInActivity = @()
}

# Build appId -> activity lookup
$activityByAppId = @{}
foreach ($a in $signInActivity) { if ($a.appId) { $activityByAppId[$a.appId] = $a } }

# ── Phase 3: Analyze ──
$phase = 3
Write-Host ""
Write-Host "  [$phase/$totalPhases] Analyzing..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

function Get-MaxDate {
  param($dates)
  $valid = @($dates | Where-Object { $_ -and $_ -ne '' })
  if ($valid.Count -eq 0) { return $null }
  return ($valid | ForEach-Object { [datetime]$_ } | Sort-Object -Descending | Select-Object -First 1)
}

function Get-CredentialStatus {
  param($app)
  $total = 0; $expired = 0; $expiringSoon = 0
  $now = Get-Date
  $in30 = $now.AddDays(30)
  foreach ($set in @($app.keyCredentials, $app.passwordCredentials)) {
    foreach ($c in @($set)) {
      if (-not $c) { continue }
      $total++
      if ($c.endDateTime) {
        $exp = [datetime]$c.endDateTime
        if     ($exp -lt $now) { $expired++ }
        elseif ($exp -lt $in30) { $expiringSoon++ }
      }
    }
  }
  return @{ Total = $total; Expired = $expired; ExpiringSoon = $expiringSoon }
}

$analysis = @()
$i = 0
$total = $appRegs.Count

foreach ($app in $appRegs) {
  $i++
  if (($i % 25) -eq 0 -or $i -eq $total) {
    Write-ProgressBar -Current $i -Total $total -Activity "Analyzing apps" -Status $app.displayName
  }

  # First-party check: app reg owned by MS tenant (rare — most first-party apps only have an SP)
  $sp = if ($app.appId -and $spByAppId.ContainsKey($app.appId)) { $spByAppId[$app.appId] } else { $null }
  $isFirstParty = $false
  if ($sp -and $sp.appOwnerOrganizationId -eq $msFirstPartyTenantId) { $isFirstParty = $true }
  if ($app.publisherDomain -match '(microsoft\.com|microsoft\.onmicrosoft\.com)$') { $isFirstParty = $true }

  if ($isFirstParty -and -not $IncludeFirstParty) { continue }

  # Sign-in activity
  $activity = if ($app.appId -and $activityByAppId.ContainsKey($app.appId)) { $activityByAppId[$app.appId] } else { $null }
  $lastSignIn    = if ($activity -and $activity.lastSignInActivity) { $activity.lastSignInActivity.lastSignInDateTime } else { $null }
  $lastDelegated = if ($activity -and $activity.delegatedClientSignInActivity) { $activity.delegatedClientSignInActivity.lastSignInDateTime } else { $null }
  $lastAppOnly   = if ($activity -and $activity.applicationAuthenticationClientSignInActivity) { $activity.applicationAuthenticationClientSignInActivity.lastSignInDateTime } else { $null }
  $lastAsResDel  = if ($activity -and $activity.delegatedResourceSignInActivity) { $activity.delegatedResourceSignInActivity.lastSignInDateTime } else { $null }
  $lastAsResApp  = if ($activity -and $activity.applicationAuthenticationResourceSignInActivity) { $activity.applicationAuthenticationResourceSignInActivity.lastSignInDateTime } else { $null }

  # Status
  if (-not $sp) {
    $status = "No service principal"
    $reason = "App registration exists but has no service principal in this tenant — cannot sign in here."
    $daysSince = $null
  } elseif (-not $lastSignIn) {
    $status = "Never used"
    $reason = "No sign-in activity recorded. App has never been used since the report was populated."
    $daysSince = $null
  } else {
    $lastDt = [datetime]$lastSignIn
    $daysSince = [int]((Get-Date) - $lastDt).TotalDays
    if ($lastDt -lt $cutoffDate) {
      $status = "Stale"
      $reason = "Last sign-in was $daysSince days ago (> $StaleDays)."
    } else {
      $status = "Active"
      $reason = "Last sign-in was $daysSince days ago."
    }
  }

  $creds = Get-CredentialStatus $app

  $analysis += [PSCustomObject]@{
    DisplayName          = $app.displayName
    AppId                = $app.appId
    ObjectId             = $app.id
    SPObjectId           = if ($sp) { $sp.id } else { $null }
    SPEnabled            = if ($sp) { [bool]$sp.accountEnabled } else { $null }
    SPType               = if ($sp) { $sp.servicePrincipalType } else { $null }
    FirstParty           = $isFirstParty
    Status               = $status
    Reason               = $reason
    DaysSinceLastSignIn  = $daysSince
    LastSignIn           = $lastSignIn
    LastDelegatedClient  = $lastDelegated
    LastAppOnlyClient    = $lastAppOnly
    LastDelegatedResource = $lastAsResDel
    LastAppOnlyResource  = $lastAsResApp
    CreatedDateTime      = $app.createdDateTime
    PublisherDomain      = $app.publisherDomain
    SignInAudience       = $app.signInAudience
    TotalCredentials     = $creds.Total
    ExpiredCredentials   = $creds.Expired
    ExpiringSoon         = $creds.ExpiringSoon
    PortalLink           = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($app.appId)"
  }
}

Write-Progress -Activity "Analyzing apps" -Completed

# Categorize
$active      = @($analysis | Where-Object { $_.Status -eq 'Active' })
$stale       = @($analysis | Where-Object { $_.Status -eq 'Stale' })
$neverUsed   = @($analysis | Where-Object { $_.Status -eq 'Never used' })
$orphaned    = @($analysis | Where-Object { $_.Status -eq 'No service principal' })

Write-Host ""
Write-Host "  Summary:" -ForegroundColor White
Write-Host "    Active:              $($active.Count)" -ForegroundColor Green
Write-Host "    Stale (> $StaleDays d):   $($stale.Count)" -ForegroundColor Yellow
Write-Host "    Never used:          $($neverUsed.Count)" -ForegroundColor Yellow
Write-Host "    No SP (orphaned):    $($orphaned.Count)" -ForegroundColor DarkGray
Write-Host "    TOTAL analyzed:      $($analysis.Count)" -ForegroundColor White
Write-Host ""

# ── Phase 4: Decide ──
$phase = 4
Write-Host "  [$phase/$totalPhases] Review & decide" -ForegroundColor Cyan
Write-Host ""

$cleanupTargets = @($stale + $neverUsed)
if ($cleanupTargets.Count -eq 0) {
  Write-Host "  No stale or never-used apps found. Exporting audit only." -ForegroundColor Green
  $ActionChoice = "1"
} else {
  Write-Host "  $($cleanupTargets.Count) app(s) are candidates for cleanup." -ForegroundColor Yellow
  Write-Host ""
  Write-Host "    [1]  Export only (audit, no changes)       (default, safe)" -ForegroundColor Green
  Write-Host "    [2]  Disable service principal (reversible)" -ForegroundColor Yellow
  Write-Host "    [3]  Delete app registration (PERMANENT)" -ForegroundColor Red
  Write-Host ""
  $ActionChoice = Read-Host "  Enter choice (1-3) [default: 1]"
  if (-not $ActionChoice) { $ActionChoice = "1" }
}

$actionLog = @()
$actionMode = "Audit"

switch ($ActionChoice) {
  "2" {
    $actionMode = "Disable SP"
    Write-Host ""
    Write-Host "  Preview — SPs that will be disabled (accountEnabled=false):" -ForegroundColor Yellow
    $preview = @($cleanupTargets | Where-Object { $_.SPObjectId -and $_.SPEnabled } | Select-Object -First 10)
    foreach ($p in $preview) {
      Write-Host "    - $($p.DisplayName)  ($($p.AppId))  [$($p.Status)]" -ForegroundColor Gray
    }
    $remaining = @($cleanupTargets | Where-Object { $_.SPObjectId -and $_.SPEnabled }).Count
    if ($remaining -gt 10) { Write-Host "    ... and $($remaining - 10) more" -ForegroundColor Gray }
    Write-Host ""
    $confirm = Read-Host "  Type YES to disable $remaining service principal(s), anything else to cancel"
    if ($confirm -ne "YES") {
      Write-Host "  Cancelled — no changes made." -ForegroundColor Yellow
      $actionMode = "Audit (cancelled disable)"
    } else {
      $j = 0
      foreach ($t in $cleanupTargets) {
        if (-not $t.SPObjectId -or -not $t.SPEnabled) { continue }
        $j++
        Write-ProgressBar -Current $j -Total $remaining -Activity "Disabling SPs" -Status $t.DisplayName
        try {
          Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($t.SPObjectId)" `
            -Body (@{ accountEnabled = $false } | ConvertTo-Json) `
            -ContentType "application/json" | Out-Null
          $actionLog += [PSCustomObject]@{
            Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action       = "Disable SP"
            DisplayName  = $t.DisplayName
            AppId        = $t.AppId
            Result       = "Success"
            ErrorMessage = $null
          }
        } catch {
          $actionLog += [PSCustomObject]@{
            Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action       = "Disable SP"
            DisplayName  = $t.DisplayName
            AppId        = $t.AppId
            Result       = "Failed"
            ErrorMessage = $_.Exception.Message
          }
        }
      }
      Write-Progress -Activity "Disabling SPs" -Completed
      Write-Host "  Done. Review the Action Log sheet in the report." -ForegroundColor Green
    }
  }
  "3" {
    $actionMode = "Delete app registration"
    Write-Host ""
    Write-Host "  [!] DELETE is PERMANENT. The app registration cannot be recovered." -ForegroundColor Red
    Write-Host "  [!] Consumers of the app will immediately fail authentication." -ForegroundColor Red
    Write-Host ""
    $targets = @($cleanupTargets | Where-Object { $_.ObjectId })
    Write-Host "  Preview — app registrations that will be DELETED:" -ForegroundColor Red
    $preview = @($targets | Select-Object -First 10)
    foreach ($p in $preview) {
      Write-Host "    - $($p.DisplayName)  ($($p.AppId))  [$($p.Status)]" -ForegroundColor Gray
    }
    if ($targets.Count -gt 10) { Write-Host "    ... and $($targets.Count - 10) more" -ForegroundColor Gray }
    Write-Host ""
    $confirm1 = Read-Host "  Type YES to continue to final confirmation"
    if ($confirm1 -ne "YES") {
      Write-Host "  Cancelled — no changes made." -ForegroundColor Yellow
      $actionMode = "Audit (cancelled delete)"
    } else {
      Write-Host ""
      $confirm2 = Read-Host "  FINAL CONFIRMATION — type DELETE to permanently remove $($targets.Count) app registration(s)"
      if ($confirm2 -ne "DELETE") {
        Write-Host "  Cancelled — no changes made." -ForegroundColor Yellow
        $actionMode = "Audit (cancelled delete)"
      } else {
        $j = 0
        foreach ($t in $targets) {
          $j++
          Write-ProgressBar -Current $j -Total $targets.Count -Activity "Deleting app registrations" -Status $t.DisplayName
          try {
            Invoke-MgGraphRequest -Method DELETE `
              -Uri "https://graph.microsoft.com/v1.0/applications/$($t.ObjectId)" | Out-Null
            $actionLog += [PSCustomObject]@{
              Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
              Action       = "Delete app reg"
              DisplayName  = $t.DisplayName
              AppId        = $t.AppId
              Result       = "Success"
              ErrorMessage = $null
            }
          } catch {
            $actionLog += [PSCustomObject]@{
              Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
              Action       = "Delete app reg"
              DisplayName  = $t.DisplayName
              AppId        = $t.AppId
              Result       = "Failed"
              ErrorMessage = $_.Exception.Message
            }
          }
        }
        Write-Progress -Activity "Deleting app registrations" -Completed
        Write-Host "  Done. Review the Action Log sheet in the report." -ForegroundColor Green
      }
    }
  }
  default {
    $actionMode = "Audit"
    Write-Host "  Audit-only run. No changes made." -ForegroundColor Green
  }
}

# ── Export Excel report (Griffin31 styling) ──
Write-Host ""
Write-Host "  Writing Excel report..." -ForegroundColor Cyan

$reportDir = if ($IsMacOS -or $IsLinux) { "$HOME/Desktop" } else { [Environment]::GetFolderPath("Desktop") }
if (-not (Test-Path $reportDir)) { $reportDir = $HOME }
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$reportFile = Join-Path $reportDir "StaleApps_Audit_$timestamp.xlsx"

# ── Excel color helpers ──
function HexColor([string]$hex) {
  $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
  $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
  $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
  return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

$NAVY        = HexColor "1B2A4A"
$DARK_BLUE   = HexColor "2C3E6B"
$ACCENT_BLUE = HexColor "4472C4"
$LIGHT_BLUE  = HexColor "D6E4F0"
$WHITE_C     = [System.Drawing.Color]::White
$DARK_GRAY   = HexColor "404040"
$ROW_ALT     = HexColor "F8F9FA"
$RED_BG      = HexColor "FDE8E8"
$RED_TEXT    = HexColor "B91C1C"
$AMBER_BG    = HexColor "FEF3C7"
$AMBER_TEXT  = HexColor "92400E"
$GREEN_BG    = HexColor "D1FAE5"
$GREEN_TEXT  = HexColor "065F46"
$GRAY_BG     = HexColor "E5E7EB"
$GRAY_TEXT   = HexColor "374151"

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
  $b.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $b.Top.Color.SetColor($color)
  $b.Bottom.Color.SetColor($color)
  $b.Left.Color.SetColor($color)
  $b.Right.Color.SetColor($color)
}

# ── Build workbook ──
$excel = New-Object OfficeOpenXml.ExcelPackage

# ── Sheet: Summary Dashboard ──
$ws = $excel.Workbook.Worksheets.Add("Summary")
$ws.TabColor = HexColor "2E75B6"
$ws.View.ShowGridLines = $false

$ws.Column(1).Width = 2
2..6 | ForEach-Object { $ws.Column($_).Width = 25 }

# Title bar (navy band)
$ws.Row(1).Height = 8
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[1, $c] $NAVY }
$ws.Row(2).Height = 36
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[2, $c] $NAVY }
$ws.Cells["B2:F2"].Merge = $true
$ws.Cells["B2"].Value = "STALE APP REGISTRATIONS AUDIT"
Set-Font $ws.Cells["B2"] -size 18 -bold $true -color $WHITE_C
$ws.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Subtitle
$ws.Row(3).Height = 24
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[3, $c] $DARK_BLUE }
$ws.Cells["B3:F3"].Merge = $true
$ws.Cells["B3"].Value = "Tenant: $($ctx.TenantId)  |  Threshold: $StaleDays days  |  Mode: $actionMode  |  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Set-Font $ws.Cells["B3"] -size 10 -color (HexColor "B0C4DE")

# KPI cards
$ws.Row(4).Height = 8
$ws.Row(5).Height = 60
$kpis = @(
  @{ Col = 2; Number = "$($analysis.Count)"; Label = "TOTAL APPS" }
  @{ Col = 3; Number = "$($stale.Count)";    Label = "STALE" }
  @{ Col = 4; Number = "$($neverUsed.Count)"; Label = "NEVER USED" }
  @{ Col = 5; Number = "$($active.Count)";   Label = "ACTIVE" }
  @{ Col = 6; Number = "$($orphaned.Count)"; Label = "NO SERVICE PRINCIPAL" }
)
foreach ($kpi in $kpis) {
  $cell = $ws.Cells[5, $kpi.Col]
  $cell.IsRichText = $true
  $rt = $cell.RichText
  $num = $rt.Add($kpi.Number + "`n")
  $num.Size = 22; $num.Bold = $true; $num.Color = $NAVY
  $lbl = $rt.Add($kpi.Label)
  $lbl.Size = 9; $lbl.Bold = $false; $lbl.Color = $DARK_GRAY
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $cell.Style.WrapText = $true
  Set-Fill $cell $LIGHT_BLUE
  Set-ThinBorder $cell
}

# Breakdown by status
$row = 7
$ws.Cells["B$row"].Value = "BREAKDOWN BY STATUS"
Set-Font $ws.Cells["B$row"] -size 12 -bold $true -color $NAVY
$row++

$ws.Cells["B$row"].Value = "Status"
$ws.Cells["C$row"].Value = "Count"
$ws.Cells["D$row"].Value = "% of Total"
for ($c = 2; $c -le 4; $c++) {
  Set-Fill $ws.Cells[$row, $c] $ACCENT_BLUE
  Set-Font $ws.Cells[$row, $c] -size 10 -bold $true -color $WHITE_C
}
$row++

$totalCount = [math]::Max(1, $analysis.Count)
$breakdown = @(
  @{ Status = "Active";                Count = $active.Count;     BgC = $GREEN_BG; FgC = $GREEN_TEXT },
  @{ Status = "Stale (> $StaleDays d)"; Count = $stale.Count;     BgC = $AMBER_BG; FgC = $AMBER_TEXT },
  @{ Status = "Never used";            Count = $neverUsed.Count;  BgC = $RED_BG;   FgC = $RED_TEXT },
  @{ Status = "No service principal";  Count = $orphaned.Count;   BgC = $GRAY_BG;  FgC = $GRAY_TEXT }
)
foreach ($b in $breakdown) {
  $pct = [math]::Round(($b.Count / $totalCount) * 100, 1)
  $ws.Cells["B$row"].Value = $b.Status
  $ws.Cells["C$row"].Value = $b.Count
  $ws.Cells["D$row"].Value = "$pct%"
  Set-Fill $ws.Cells["B$row"] $b.BgC
  Set-Font $ws.Cells["B$row"] -color $b.FgC -bold $true
  for ($c = 2; $c -le 4; $c++) {
    Set-ThinBorder $ws.Cells[$row, $c]
    $ws.Cells[$row, $c].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  }
  $ws.Cells["C$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells["D$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $row++
}

# ── Detail sheet builder ──
function Add-AppSheet {
  param(
    [string]$Name,
    [string]$TabHex,
    [string]$TitleText,
    [array]$Apps,
    [string]$DaysColor  # "red" | "amber" | "none"
  )
  if (-not $Apps -or $Apps.Count -eq 0) { return }
  $ws = $excel.Workbook.Worksheets.Add($Name)
  $ws.TabColor = HexColor $TabHex
  $ws.View.ShowGridLines = $false

  # Column widths
  $ws.Column(1).Width = 2
  $ws.Column(2).Width = 32  # App Name
  $ws.Column(3).Width = 38  # App ID
  $ws.Column(4).Width = 14  # Status
  $ws.Column(5).Width = 15  # Days since
  $ws.Column(6).Width = 20  # Last Sign-In
  $ws.Column(7).Width = 12  # SP Enabled
  $ws.Column(8).Width = 14  # First-Party
  $ws.Column(9).Width = 12  # Creds total
  $ws.Column(10).Width = 12 # Creds expired
  $ws.Column(11).Width = 14 # Created
  $ws.Column(12).Width = 22 # Portal Link

  $lastCol = 12

  # Banner
  $ws.Row(1).Height = 8
  for ($c = 1; $c -le $lastCol; $c++) { Set-Fill $ws.Cells[1, $c] $NAVY }
  $ws.Row(2).Height = 32
  for ($c = 1; $c -le $lastCol; $c++) { Set-Fill $ws.Cells[2, $c] $NAVY }
  $ws.Cells[2, 2, 2, $lastCol].Merge = $true
  $ws.Cells[2, 2].Value = $TitleText
  Set-Font $ws.Cells[2, 2] -size 16 -bold $true -color $WHITE_C
  $ws.Cells[2, 2].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  # Header row
  $hdrRow = 4
  $headers = @("App Name","App ID","Status","Days Since","Last Sign-In","SP Enabled","First-Party","Credentials","Expired Creds","Created","Portal Link")
  $ws.Row($hdrRow).Height = 28
  for ($i = 0; $i -lt $headers.Count; $i++) {
    $col = 2 + $i
    $cell = $ws.Cells[$hdrRow, $col]
    $cell.Value = $headers[$i]
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
    $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  }

  # Data rows
  $dataRow = $hdrRow + 1
  foreach ($app in $Apps) {
    $ws.Row($dataRow).Height = 20

    $ws.Cells[$dataRow, 2].Value  = $app.DisplayName
    $ws.Cells[$dataRow, 3].Value  = $app.AppId
    $ws.Cells[$dataRow, 4].Value  = $app.Status
    $ws.Cells[$dataRow, 5].Value  = if ($null -ne $app.DaysSinceLastSignIn) { $app.DaysSinceLastSignIn } else { "—" }
    $ws.Cells[$dataRow, 6].Value  = if ($app.LastSignIn) { ([datetime]$app.LastSignIn).ToString("yyyy-MM-dd HH:mm") } else { "—" }
    $ws.Cells[$dataRow, 7].Value  = if ($null -ne $app.SPEnabled) { if ($app.SPEnabled) { "Yes" } else { "No" } } else { "—" }
    $ws.Cells[$dataRow, 8].Value  = if ($app.FirstParty) { "Yes" } else { "No" }
    $ws.Cells[$dataRow, 9].Value  = $app.TotalCredentials
    $ws.Cells[$dataRow,10].Value  = $app.ExpiredCredentials
    $ws.Cells[$dataRow,11].Value  = if ($app.CreatedDateTime) { ([datetime]$app.CreatedDateTime).ToString("yyyy-MM-dd") } else { "—" }

    $ws.Cells[$dataRow,12].Value = "Open in Entra"
    if ($app.AppId) {
      $ws.Cells[$dataRow,12].Hyperlink = [System.Uri]::new($app.PortalLink)
      $ws.Cells[$dataRow,12].Style.Font.UnderLine = $true
      Set-Font $ws.Cells[$dataRow,12] -color $ACCENT_BLUE
    }

    # App ID in mono-ish alignment
    $ws.Cells[$dataRow, 3].Style.Font.Name = "Consolas"
    $ws.Cells[$dataRow, 3].Style.Font.Size = 9

    # Center status and small columns
    foreach ($col in 4,5,7,8,9,10,11,12) {
      $ws.Cells[$dataRow, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }

    # Color status cell
    $statusCell = $ws.Cells[$dataRow, 4]
    switch ($app.Status) {
      "Active"               { Set-Fill $statusCell $GREEN_BG; Set-Font $statusCell -color $GREEN_TEXT -bold $true }
      "Stale"                { Set-Fill $statusCell $AMBER_BG; Set-Font $statusCell -color $AMBER_TEXT -bold $true }
      "Never used"           { Set-Fill $statusCell $RED_BG;   Set-Font $statusCell -color $RED_TEXT   -bold $true }
      "No service principal" { Set-Fill $statusCell $GRAY_BG;  Set-Font $statusCell -color $GRAY_TEXT  -bold $true }
    }

    # Color Days Since column based on value
    if ($DaysColor -ne "none" -and $null -ne $app.DaysSinceLastSignIn) {
      $daysCell = $ws.Cells[$dataRow, 5]
      $d = [int]$app.DaysSinceLastSignIn
      if ($d -ge 180) { Set-Fill $daysCell $RED_BG;   Set-Font $daysCell -color $RED_TEXT -bold $true }
      elseif ($d -ge 90) { Set-Fill $daysCell $AMBER_BG; Set-Font $daysCell -color $AMBER_TEXT -bold $true }
      else { Set-Fill $daysCell $GREEN_BG; Set-Font $daysCell -color $GREEN_TEXT -bold $true }
    }

    # Expired credential warning
    if ($app.ExpiredCredentials -gt 0) {
      $ec = $ws.Cells[$dataRow, 10]
      Set-Fill $ec $RED_BG; Set-Font $ec -color $RED_TEXT -bold $true
    }

    # Borders + zebra striping
    for ($col = 2; $col -le $lastCol; $col++) {
      Set-ThinBorder $ws.Cells[$dataRow, $col]
    }
    if ($dataRow % 2 -eq 0) {
      for ($col = 2; $col -le $lastCol; $col++) {
        $isColored = ($col -eq 4) -or ($col -eq 5 -and $DaysColor -ne "none" -and $null -ne $app.DaysSinceLastSignIn) -or ($col -eq 10 -and $app.ExpiredCredentials -gt 0)
        if (-not $isColored) { Set-Fill $ws.Cells[$dataRow, $col] $ROW_ALT }
      }
    }

    $dataRow++
  }

  # Freeze below header, enable filter
  $ws.View.FreezePanes($hdrRow + 1, 1)
  $ws.Cells[$hdrRow, 2, $hdrRow, $lastCol].AutoFilter = $true
}

Add-AppSheet -Name "Stale"      -TabHex "E67E22" -TitleText "STALE — Last sign-in older than $StaleDays days" -Apps $stale    -DaysColor "amber"
Add-AppSheet -Name "Never Used" -TabHex "E74C3C" -TitleText "NEVER USED — No sign-in activity on record"       -Apps $neverUsed -DaysColor "red"
Add-AppSheet -Name "No SP"      -TabHex "7F8C8D" -TitleText "NO SERVICE PRINCIPAL — Orphaned app registrations" -Apps $orphaned  -DaysColor "none"
Add-AppSheet -Name "Active"     -TabHex "27AE60" -TitleText "ACTIVE — Sign-in within the last $StaleDays days"  -Apps $active    -DaysColor "none"

# ── Action Log (only if destructive actions ran) ──
if ($actionLog.Count -gt 0) {
  $ws4 = $excel.Workbook.Worksheets.Add("Action Log")
  $ws4.TabColor = HexColor "8E44AD"
  $ws4.View.ShowGridLines = $false
  $ws4.Column(1).Width = 2
  $ws4.Column(2).Width = 20; $ws4.Column(3).Width = 20; $ws4.Column(4).Width = 34
  $ws4.Column(5).Width = 38; $ws4.Column(6).Width = 14; $ws4.Column(7).Width = 50

  $ws4.Row(1).Height = 8
  for ($c = 1; $c -le 7; $c++) { Set-Fill $ws4.Cells[1, $c] $NAVY }
  $ws4.Row(2).Height = 32
  for ($c = 1; $c -le 7; $c++) { Set-Fill $ws4.Cells[2, $c] $NAVY }
  $ws4.Cells["B2:G2"].Merge = $true
  $ws4.Cells["B2"].Value = "ACTION LOG"
  Set-Font $ws4.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws4.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  $hdrRow = 4
  $ws4.Row($hdrRow).Height = 28
  $logHeaders = @("Timestamp","Action","Display Name","App ID","Result","Error Message")
  for ($i = 0; $i -lt $logHeaders.Count; $i++) {
    $col = 2 + $i
    $cell = $ws4.Cells[$hdrRow, $col]
    $cell.Value = $logHeaders[$i]
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  }

  $r = $hdrRow + 1
  foreach ($entry in $actionLog) {
    $ws4.Cells[$r, 2].Value = $entry.Timestamp
    $ws4.Cells[$r, 3].Value = $entry.Action
    $ws4.Cells[$r, 4].Value = $entry.DisplayName
    $ws4.Cells[$r, 5].Value = $entry.AppId
    $ws4.Cells[$r, 6].Value = $entry.Result
    $ws4.Cells[$r, 7].Value = $entry.ErrorMessage

    $resCell = $ws4.Cells[$r, 6]
    if ($entry.Result -eq "Success") { Set-Fill $resCell $GREEN_BG; Set-Font $resCell -color $GREEN_TEXT -bold $true }
    else { Set-Fill $resCell $RED_BG; Set-Font $resCell -color $RED_TEXT -bold $true }

    $ws4.Cells[$r, 5].Style.Font.Name = "Consolas"; $ws4.Cells[$r, 5].Style.Font.Size = 9
    for ($c = 2; $c -le 7; $c++) { Set-ThinBorder $ws4.Cells[$r, $c] }
    if ($r % 2 -eq 0) {
      for ($c = 2; $c -le 7; $c++) {
        if ($c -ne 6) { Set-Fill $ws4.Cells[$r, $c] $ROW_ALT }
      }
    }
    $r++
  }
  $ws4.View.FreezePanes($hdrRow + 1, 1)
  $ws4.Cells[$hdrRow, 2, $hdrRow, 7].AutoFilter = $true
}

# ── Save ──
try {
  $excel.SaveAs($reportFile)
  $excel.Dispose()
  Write-Host "  Report saved: $reportFile" -ForegroundColor Green
} catch {
  Write-Host "  [!] Export failed: $($_.Exception.Message)" -ForegroundColor Red
}

# ── Done ──
$swTotal.Stop()
Write-Host ""
Write-Host "  Completed in $($swTotal.Elapsed.ToString('mm\:ss'))." -ForegroundColor Cyan
Write-Host ""

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
