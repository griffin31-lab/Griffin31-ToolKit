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

try {
  Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null
  $ctx = Get-MgContext
  if (-not $ctx) {
    Write-Host "  [!] Graph context not established." -ForegroundColor Red
    return
  }
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

# ── Export Excel report ──
Write-Host ""
Write-Host "  Writing Excel report..." -ForegroundColor Cyan

$reportDir = if ($IsMacOS -or $IsLinux) { "$HOME/Desktop" } else { [Environment]::GetFolderPath("Desktop") }
if (-not (Test-Path $reportDir)) { $reportDir = $HOME }
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$reportFile = Join-Path $reportDir "StaleApps_Audit_$timestamp.xlsx"

# Summary sheet data
$summary = @(
  [PSCustomObject]@{ Metric = "Tenant";                      Value = $ctx.TenantId }
  [PSCustomObject]@{ Metric = "Run by";                      Value = $ctx.Account }
  [PSCustomObject]@{ Metric = "Generated";                   Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
  [PSCustomObject]@{ Metric = "Staleness threshold (days)";  Value = $StaleDays }
  [PSCustomObject]@{ Metric = "First-party apps";            Value = if ($IncludeFirstParty) { "Included" } else { "Excluded" } }
  [PSCustomObject]@{ Metric = "Action mode";                 Value = $actionMode }
  [PSCustomObject]@{ Metric = "";                            Value = "" }
  [PSCustomObject]@{ Metric = "Total app registrations";     Value = $analysis.Count }
  [PSCustomObject]@{ Metric = "Active";                      Value = $active.Count }
  [PSCustomObject]@{ Metric = "Stale (> $StaleDays d)";      Value = $stale.Count }
  [PSCustomObject]@{ Metric = "Never used";                  Value = $neverUsed.Count }
  [PSCustomObject]@{ Metric = "No service principal";        Value = $orphaned.Count }
)

try {
  $summary | Export-Excel -Path $reportFile -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow

  if ($stale.Count -gt 0) {
    $stale | Export-Excel -Path $reportFile -WorksheetName "Stale" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium2
  }
  if ($neverUsed.Count -gt 0) {
    $neverUsed | Export-Excel -Path $reportFile -WorksheetName "Never Used" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium4
  }
  if ($orphaned.Count -gt 0) {
    $orphaned | Export-Excel -Path $reportFile -WorksheetName "No SP" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium6
  }
  if ($active.Count -gt 0) {
    $active | Export-Excel -Path $reportFile -WorksheetName "Active" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium1
  }
  if ($actionLog.Count -gt 0) {
    $actionLog | Export-Excel -Path $reportFile -WorksheetName "Action Log" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium3
  }

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
