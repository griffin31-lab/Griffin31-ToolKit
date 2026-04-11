# ==============================
# Conditional Access Readiness Assessment
# Identifies apps using only OIDC/directory scopes affected by the
# May 13 2026 CA enforcement change.
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
    Write-Host "  Then run:  pwsh .\graph-all-apps.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./graph-all-apps.ps1" -ForegroundColor Yellow
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
     Conditional Access Readiness Assessment
     Microsoft Entra ID  -  May 13 2026 CA Enforcement Change
  ============================================================

"@
Write-Host $banner -ForegroundColor Cyan

# ── Step 1: Collect UPN and resolve tenant ──
Write-Host "  This script connects via delegated Graph permissions." -ForegroundColor Gray
Write-Host "  You will be prompted to sign in with a browser." -ForegroundColor Gray
Write-Host ""

$AdminUPN = Read-Host "  Enter Global Admin UPN (e.g. admin@contoso.onmicrosoft.com)"
if (-not $AdminUPN -or $AdminUPN.Trim().Length -eq 0) {
  Write-Host "  [!] No UPN provided. Exiting." -ForegroundColor Red
  return
}
$AdminUPN = $AdminUPN.Trim()

# Basic UPN format check
if ($AdminUPN -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
  Write-Host "  [!] '$AdminUPN' doesn't look like a valid UPN (expected: user@domain.com)." -ForegroundColor Red
  Write-Host "  Please re-run the script with a valid UPN." -ForegroundColor Yellow
  return
}

# Extract tenant hint from UPN domain
$upnDomain = ($AdminUPN -split '@')[-1]
Write-Host ""
Write-Host "  Tenant domain detected: " -NoNewline; Write-Host $upnDomain -ForegroundColor Yellow

# ── Step 2: Audit period selection ──
Write-Host ""
Write-Host "  Select audit period for sign-in analysis:" -ForegroundColor White
Write-Host "    [1]  7 days   (faster  ~ 2-5 min)" -ForegroundColor Green
Write-Host "    [2]  30 days  (deeper  ~ 5-15 min)" -ForegroundColor Yellow
Write-Host ""
$periodChoice = Read-Host "  Enter choice (1 or 2) [default: 1]"
if ($periodChoice -eq "2") {
  $AuditDays   = 30
  $GraphPeriod = "D30"
  Write-Host "  -> 30-day audit selected" -ForegroundColor Yellow
} else {
  $AuditDays   = 7
  $GraphPeriod  = "D7"
  Write-Host "  -> 7-day audit selected" -ForegroundColor Green
}

# ── Time estimates ──
$estMin = if ($AuditDays -eq 7) { "2-5" } else { "5-15" }
Write-Host ""
Write-Host "  Estimated run time: $estMin minutes (depends on tenant size)" -ForegroundColor Gray
Write-Host "  Phases: Connect -> Grants -> Sign-in summary -> Enrich apps -> MFA audit -> Export" -ForegroundColor Gray
Write-Host ""

# Confirm
$confirm = Read-Host "  Press ENTER to start or type 'q' to quit"
if ($confirm -and $confirm.Trim().ToLower() -ne '') {
  Write-Host "  Cancelled." -ForegroundColor Red
  return
}

# ── Overall timer ──
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()

$BaselineScopes = @(
  "openid", "profile", "email", "offline_access",
  "User.Read", "User.Read.All", "User.ReadBasic.All",
  "People.Read", "People.Read.All",
  "GroupMember.Read.All", "Member.Read.Hidden"
)

# ── Phase 1: Connect ──
Write-Host ""
Write-Host "  [1/6] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$requiredScopes = @(
  "DelegatedPermissionGrant.Read.All",
  "Directory.Read.All",
  "Reports.Read.All",
  "AuditLog.Read.All"
)

# Check if already connected with adequate scopes
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
  try {
    Connect-MgGraph -Scopes $requiredScopes -TenantId $upnDomain -ErrorAction Stop
    $ctx = Get-MgContext
  }
  catch {
    Write-Host ""
    Write-Host "  [!] Failed to connect to Microsoft Graph." -ForegroundColor Red
    Write-Host "  This usually means the browser sign-in was cancelled or timed out." -ForegroundColor Yellow
    Write-Host "  Please re-run the script and complete the sign-in prompt." -ForegroundColor Yellow
    Write-Host ""
    return
  }
}

$TenantId = $ctx.TenantId
if (-not $TenantId) {
  Write-Host "  [!] Failed to connect - no tenant ID resolved." -ForegroundColor Red
  Write-Host "  Please re-run the script and complete the browser sign-in." -ForegroundColor Yellow
  return
}
Write-Host "  Connected to tenant: $TenantId  ($($ctx.Account))" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 2: Pull oauth2PermissionGrants ──
Write-Host "  [2/6] Fetching delegated permission grants..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$uri = "https://graph.microsoft.com/beta/oauth2PermissionGrants?`$select=clientId,scope,consentType"
$grants = @()
$pageNum = 0
while ($uri) {
  $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
  $grants += $resp.value
  $pageNum++
  Write-Host "`r  Grants fetched: $($grants.Count) (page $pageNum)..." -NoNewline -ForegroundColor Gray
  $uri = $resp.'@odata.nextLink'
}
Write-Host "`r  Grants fetched: $($grants.Count) total                    " -ForegroundColor Green
$script:swPhase.Stop()

# ------------------------------
# Build baseline-only candidate set
# ------------------------------
$scopesByClient = @{}

foreach ($g in $grants) {
  if (-not $g -or -not $g.clientId) { continue }
  $cid = $g.clientId.ToString().Trim()
  if (-not $cid) { continue }

  if (-not $scopesByClient.ContainsKey($cid)) {
    $scopesByClient[$cid] = [System.Collections.Generic.HashSet[string]]::new(
      [System.StringComparer]::OrdinalIgnoreCase
    )
  }

  foreach ($s in ($g.scope -split '\s+')) {
    if ($s -and $s.Trim().Length -gt 0) {
      [void]$scopesByClient[$cid].Add($s.Trim())
    }
  }
}

$candidates = foreach ($cid in $scopesByClient.Keys) {
  $scopes = $scopesByClient[$cid]
  if ($scopes.Count -le 0) { continue }

  $outside = $scopes | Where-Object { $_ -notin $BaselineScopes }
  if ($outside.Count -eq 0) {
    [PSCustomObject]@{
      ServicePrincipalObjectId = $cid
      Scopes = ($scopes -join ' ')
    }
  }
}

# ── Phase 3: Pull sign-in summary ──
Write-Host "  [3/6] Fetching app sign-in summary ($AuditDays days)..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$signInSummary = @()
$signInUri = "https://graph.microsoft.com/beta/reports/getAzureADApplicationSignInSummary(period='$GraphPeriod')"

while ($signInUri) {
  $resp = Invoke-MgGraphRequest -Method GET -Uri $signInUri

  if ($resp -and $resp.value) {
    $signInSummary += $resp.value
  }

  $signInUri = $resp.'@odata.nextLink'
}
Write-Host "  Sign-in summaries: $($signInSummary.Count) apps" -ForegroundColor Green
$script:swPhase.Stop()

# appId -> total sign-ins
$signInCountByAppId = @{}
foreach ($s in $signInSummary) {
  $appId = $s.id
  if (-not $appId) { continue }

  $success = 0
  $failed  = 0
  if ($null -ne $s.successfulSignInCount) { $success = [int]$s.successfulSignInCount }
  if ($null -ne $s.failedSignInCount)     { $failed  = [int]$s.failedSignInCount }

  $signInCountByAppId[$appId] = $success + $failed
}

$resultsTenantOwned = @()
$resultsNotTenantOwned = @()

# ── Phase 4: Enrich candidates with SP details ──
Write-Host "  [4/6] Enriching $($candidates.Count) candidate apps with details..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$enrichIdx = 0
$enrichTotal = @($candidates).Count

foreach ($c in $candidates) {
  $enrichIdx++
  $shortId = if ($c.ServicePrincipalObjectId.Length -ge 8) { $c.ServicePrincipalObjectId.Substring(0,8) } else { $c.ServicePrincipalObjectId }
  Write-ProgressBar -Current $enrichIdx -Total $enrichTotal -Activity "Enriching service principals" -Status $shortId

  try {
    $spUri = "https://graph.microsoft.com/beta/servicePrincipals/$($c.ServicePrincipalObjectId)?`$select=id,appId,displayName,appOwnerOrganizationId"
    $sp = Invoke-MgGraphRequest -Method GET -Uri $spUri

    $signinCount = 0
    if ($sp.appId -and $signInCountByAppId.ContainsKey($sp.appId)) {
      $signinCount = $signInCountByAppId[$sp.appId]
    }

    $row = [PSCustomObject]@{
      ServicePrincipalObjectId = $c.ServicePrincipalObjectId
      AppId                   = $sp.appId
      AppDisplayName          = $sp.displayName
      AppOwnerOrganizationId  = $sp.appOwnerOrganizationId
      Scopes                  = $c.Scopes
      SigninCount             = $signinCount
    }

    if ($sp.appOwnerOrganizationId -eq $TenantId) {
      $resultsTenantOwned += $row
    }
    else {
      $resultsNotTenantOwned += $row
    }
  }
  catch {
    # Skip non-enumerable / missing service principals
  }
}
Write-Progress -Activity "Enriching service principals" -Completed
Write-Host "  Found: $($resultsTenantOwned.Count) tenant-owned, $($resultsNotTenantOwned.Count) external" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 5: MFA audit on active apps ──
Write-Host "  [5/6] Auditing MFA status on sign-in logs ($AuditDays days)..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

function Get-MfaAudit {
  param([string]$AppId)
  if (-not $AppId) { return @{ SFA = 0; MFA = 0; Unknown = 0 } }

  $sfaCount = 0
  $mfaCount = 0
  $unknownCount = 0
  $cutoff = (Get-Date).AddDays(-$AuditDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
  $signInLogsUri = "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=appId eq '$AppId' and createdDateTime ge $cutoff&`$select=id,authenticationRequirement,status&`$top=100"

  try {
    while ($signInLogsUri) {
      $resp = Invoke-MgGraphRequest -Method GET -Uri $signInLogsUri

      foreach ($signIn in $resp.value) {
        if ($signIn.status -and $signIn.status.errorCode -ne 0) { continue }

        $req = $signIn.authenticationRequirement
        if ($req -eq 'multiFactorAuthentication') {
          $mfaCount++
        }
        elseif ($req -eq 'singleFactorAuthentication') {
          $sfaCount++
        }
        else {
          $unknownCount++
        }
      }

      $signInLogsUri = $resp.'@odata.nextLink'
    }
  }
  catch {
    return @{ SFA = -1; MFA = -1; Unknown = -1 }
  }

  return @{ SFA = $sfaCount; MFA = $mfaCount; Unknown = $unknownCount }
}

$allApps = @($resultsTenantOwned) + @($resultsNotTenantOwned)
$activeApps = @($allApps | Where-Object { $_.SigninCount -gt 0 })
$mfaIdx = 0
$mfaTotal = $activeApps.Count

Write-Host "  Active apps to audit: $mfaTotal" -ForegroundColor Gray

foreach ($app in $allApps) {
  $app | Add-Member -NotePropertyName MfaSignins     -NotePropertyValue 0 -Force
  $app | Add-Member -NotePropertyName SfaSignins     -NotePropertyValue 0 -Force
  $app | Add-Member -NotePropertyName UnknownSignins -NotePropertyValue 0 -Force

  if ($app.SigninCount -gt 0) {
    $mfaIdx++
    Write-ProgressBar -Current $mfaIdx -Total $mfaTotal -Activity "Auditing MFA enforcement" -Status $app.AppDisplayName
    $audit = Get-MfaAudit -AppId $app.AppId
    $app.MfaSignins     = $audit.MFA
    $app.SfaSignins     = $audit.SFA
    $app.UnknownSignins = $audit.Unknown
  }
}
Write-Progress -Activity "Auditing MFA enforcement" -Completed
Write-Host "  MFA audit complete" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 6: Export to Excel (via ImportExcel / EPPlus) ──
Write-Host "  [6/6] Generating Excel report..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Build export data
$exportData = @()

foreach ($app in $resultsTenantOwned) {
  $isFailed = ($app.SfaSignins -eq -1)
  $risk = if ($isFailed) { "UNKNOWN" }
          elseif ($app.SfaSignins -gt 0) { "HIGH" }
          elseif ($app.SigninCount -gt 0) { "MEDIUM" }
          else { "LOW" }

  $exportData += [PSCustomObject]@{
    AppType  = "Tenant-Owned"
    AppName  = $app.AppDisplayName
    AppId    = $app.AppId
    SPID     = $app.ServicePrincipalObjectId
    Scopes   = $app.Scopes
    Signins  = [int]$app.SigninCount
    MFA      = [int]$app.MfaSignins
    SFA      = [int]$app.SfaSignins
    Unknown  = [int]$app.UnknownSignins
    Risk     = $risk
  }
}

foreach ($app in $resultsNotTenantOwned) {
  $isFailed = ($app.SfaSignins -eq -1)
  $risk = if ($isFailed) { "UNKNOWN" }
          elseif ($app.SfaSignins -gt 0) { "HIGH" }
          elseif ($app.SigninCount -gt 0) { "MEDIUM" }
          else { "LOW" }

  $exportData += [PSCustomObject]@{
    AppType  = "External"
    AppName  = $app.AppDisplayName
    AppId    = $app.AppId
    SPID     = $app.ServicePrincipalObjectId
    Scopes   = $app.Scopes
    Signins  = [int]$app.SigninCount
    MFA      = [int]$app.MfaSignins
    SFA      = [int]$app.SfaSignins
    Unknown  = [int]$app.UnknownSignins
    Risk     = $risk
  }
}

# ── Determine output path (Desktop preferred) ──
$desktopPath = [System.Environment]::GetFolderPath("Desktop")
if (-not $desktopPath -or -not (Test-Path $desktopPath)) {
  $desktopPath = $PWD
}
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportPath = Join-Path $desktopPath "ConditionalAccess_Readiness_Report_$timestamp.xlsx"

# ── Stats for report ──
$tenantOwnedData = @($exportData | Where-Object { $_.AppType -eq "Tenant-Owned" })
$externalData    = @($exportData | Where-Object { $_.AppType -eq "External" })
$highRisk        = @($exportData | Where-Object { $_.Risk -eq "HIGH" })
$medRisk         = @($exportData | Where-Object { $_.Risk -eq "MEDIUM" })
$lowRisk         = @($exportData | Where-Object { $_.Risk -eq "LOW" })
$unknownRisk     = @($exportData | Where-Object { $_.Risk -eq "UNKNOWN" })
$activeTotal     = @($exportData | Where-Object { $_.Signins -gt 0 })
$generatedDate   = Get-Date -Format "MMMM dd, yyyy HH:mm"

# ── Colors (System.Drawing.Color) ──
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
$VERY_LIGHT_BLUE= HexColor "EDF2F9"
$WHITE_C        = [System.Drawing.Color]::White
$MED_GRAY       = HexColor "D9D9D9"
$DARK_GRAY      = HexColor "404040"
$CHARCOAL       = HexColor "333333"
$RED_BG_C       = HexColor "FDE8E8"
$RED_TEXT_C      = HexColor "B91C1C"
$AMBER_BG_C     = HexColor "FEF3C7"
$AMBER_TEXT_C    = HexColor "92400E"
$GREEN_BG_C     = HexColor "D1FAE5"
$GREEN_TEXT_C    = HexColor "065F46"
$GRAY_BG_C      = HexColor "E5E7EB"
$GRAY_TEXT_C     = HexColor "6B7280"
$KPI_BG_1       = HexColor "EBF5FB"
$KPI_BG_2       = HexColor "FEF9E7"
$KPI_BG_3       = HexColor "F9EBEA"
$KPI_BG_4       = HexColor "EAFAF1"

# ── Helper: apply solid fill ──
function Set-Fill($cells, [System.Drawing.Color]$color) {
  $cells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
  $cells.Style.Fill.BackgroundColor.SetColor($color)
}

# ── Helper: apply font ──
function Set-Font($cells, [string]$name = "Arial", [int]$size = 10, [bool]$bold = $false, [System.Drawing.Color]$color = $null, [bool]$underline = $false) {
  $cells.Style.Font.Name = $name
  $cells.Style.Font.Size = $size
  $cells.Style.Font.Bold = $bold
  if ($color) { $cells.Style.Font.Color.SetColor($color) }
  if ($underline) { $cells.Style.Font.UnderLine = $true }
}

# ── Helper: thin border on all sides ──
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

# ── Create workbook ──
$excel = New-Object OfficeOpenXml.ExcelPackage

# ═══════════════════════════════════════════
# SHEET 1: EXECUTIVE DASHBOARD
# ═══════════════════════════════════════════
$ws = $excel.Workbook.Worksheets.Add("Executive Dashboard")
$ws.TabColor = $NAVY
$ws.View.ShowGridLines = $false

# Column widths
$ws.Column(1).Width = 3
$ws.Column(2).Width = 22
$ws.Column(3).Width = 22
$ws.Column(4).Width = 22
$ws.Column(5).Width = 22
$ws.Column(6).Width = 3

# Title banner (rows 1-3)
for ($r = 1; $r -le 3; $r++) {
  for ($c = 1; $c -le 6; $c++) {
    Set-Fill $ws.Cells[$r, $c] $NAVY
  }
}
$ws.Cells["B1:E1"].Merge = $true
$ws.Cells["B2:E2"].Merge = $true
$ws.Cells["B3:E3"].Merge = $true
$ws.Row(1).Height = 12
$ws.Row(2).Height = 38
$ws.Row(3).Height = 22

$ws.Cells["B2"].Value = "CONDITIONAL ACCESS READINESS ASSESSMENT"
Set-Font $ws.Cells["B2"] -size 20 -bold $true -color $WHITE_C
$ws.Cells["B2"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$ws.Cells["B2"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

$ws.Cells["B3"].Value = "Tenant: $TenantId  |  Generated: $generatedDate  |  Assessment Period: $AuditDays days"
Set-Font $ws.Cells["B3"] -size 11 -color (HexColor "B0C4DE")
$ws.Cells["B3"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$ws.Cells["B3"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

$ws.Row(4).Height = 15

# KPI cards (rows 5-8)
$kpiData = @(
  @{ Col = 2; BG = $KPI_BG_1; Number = "$($exportData.Count)"; Label = "TOTAL APPS" },
  @{ Col = 3; BG = $KPI_BG_2; Number = "$($activeTotal.Count)"; Label = "ACTIVE ($($AuditDays)d)" },
  @{ Col = 4; BG = $KPI_BG_3; Number = "$($highRisk.Count)"; Label = "SFA ONLY (HIGH)" },
  @{ Col = 5; BG = $KPI_BG_4; Number = "$($medRisk.Count)"; Label = "MFA OK (MEDIUM)" }
)
$ws.Row(5).Height = 8
$ws.Row(6).Height = 42
$ws.Row(7).Height = 22
$ws.Row(8).Height = 8

foreach ($kpi in $kpiData) {
  $col = $kpi.Col
  for ($r = 5; $r -le 8; $r++) {
    Set-Fill $ws.Cells[$r, $col] $kpi.BG
    Set-ThinBorder $ws.Cells[$r, $col]
  }
  $ws.Cells[6, $col].Value = $kpi.Number
  Set-Font $ws.Cells[6, $col] -size 28 -bold $true -color $NAVY
  $ws.Cells[6, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells[6, $col].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  $ws.Cells[7, $col].Value = $kpi.Label
  Set-Font $ws.Cells[7, $col] -size 9 -color $DARK_GRAY
  $ws.Cells[7, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells[7, $col].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
}

# ── Section header helper ──
function Add-SectionHeader($ws, [int]$row, [string]$text) {
  $ws.Row($row).Height = 24
  $ws.Cells["B${row}:E${row}"].Merge = $true
  $ws.Cells[$row, 2].Value = $text
  Set-Font $ws.Cells[$row, 2] -size 13 -bold $true -color $NAVY
  for ($c = 2; $c -le 5; $c++) {
    $ws.Cells[$row, $c].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
    $ws.Cells[$row, $c].Style.Border.Bottom.Color.SetColor($ACCENT_BLUE)
  }
}

# The Change
$ws.Row(9).Height = 8
Add-SectionHeader $ws 10 "THE MICROSOFT CHANGE"
$ws.Row(11).Height = 6

$infoRows = @(
  @{ Label = "Effective Date"; Value = "May 13, 2026" },
  @{ Label = "Impact"; Value = "CA policies targeting 'All resources' with exclusions will now enforce on apps requesting only OIDC/directory scopes (openid, profile, email, offline_access, User.Read, etc.)" },
  @{ Label = "Before"; Value = "These apps bypassed CA policies when exclusions were present" },
  @{ Label = "After"; Value = "CA policies will be consistently enforced for ALL authentication flows regardless of requested scopes" }
)
$row = 12
foreach ($info in $infoRows) {
  $rowHeight = [math]::Max(20, 16 + [math]::Floor($info.Value.Length / 80) * 14)
  $ws.Row($row).Height = $rowHeight
  $ws.Cells["C${row}:E${row}"].Merge = $true

  $ws.Cells[$row, 2].Value = $info.Label
  Set-Font $ws.Cells[$row, 2] -size 10 -bold $true -color $CHARCOAL
  $ws.Cells[$row, 2].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 2].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
  $ws.Cells[$row, 2].Style.WrapText = $true

  $ws.Cells[$row, 3].Value = $info.Value
  Set-Font $ws.Cells[$row, 3] -size 10 -color $CHARCOAL
  $ws.Cells[$row, 3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 3].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
  $ws.Cells[$row, 3].Style.WrapText = $true
  $row++
}

# Risk Breakdown
$row++
Add-SectionHeader $ws $row "RISK BREAKDOWN"
$row++
$ws.Row($row).Height = 6
$row++

$riskRows = @(
  @{ Label = "HIGH"; Count = $highRisk.Count; BG = $RED_BG_C; FG = $RED_TEXT_C; Desc = "Single-factor (SFA) sign-ins detected. These apps will likely break when CA enforces MFA." },
  @{ Label = "MEDIUM"; Count = $medRisk.Count; BG = $AMBER_BG_C; FG = $AMBER_TEXT_C; Desc = "Active sign-ins all via MFA. Should handle enforcement, but monitor closely." },
  @{ Label = "LOW"; Count = $lowRisk.Count; BG = $GREEN_BG_C; FG = $GREEN_TEXT_C; Desc = "No recent sign-in activity. Lower immediate risk." },
  @{ Label = "UNKNOWN"; Count = $unknownRisk.Count; BG = $GRAY_BG_C; FG = $GRAY_TEXT_C; Desc = "Audit log query failed (insufficient permissions or data unavailable). Verify manually." }
)

foreach ($rr in $riskRows) {
  $ws.Row($row).Height = 36
  $ws.Cells["C${row}:E${row}"].Merge = $true

  $ws.Cells[$row, 2].Value = "  $($rr.Label)  ($($rr.Count))"
  Set-Font $ws.Cells[$row, 2] -size 11 -bold $true -color $rr.FG
  Set-Fill $ws.Cells[$row, 2] $rr.BG
  $ws.Cells[$row, 2].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells[$row, 2].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  Set-ThinBorder $ws.Cells[$row, 2]

  $ws.Cells[$row, 3].Value = $rr.Desc
  Set-Font $ws.Cells[$row, 3] -size 10 -color $CHARCOAL
  $ws.Cells[$row, 3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 3].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $ws.Cells[$row, 3].Style.WrapText = $true
  $row++
}

# Recommended Actions
$row++
Add-SectionHeader $ws $row "RECOMMENDED ACTIONS"
$row++
$ws.Row($row).Height = 6
$row++

$actions = @(
  @{ Num = "1"; Title = "Fix HIGH-risk apps immediately"; Desc = "These apps have SFA-only sign-ins and will break when CA enforces MFA. Update them to support MFA/claims challenges." },
  @{ Num = "2"; Title = "Test before May 13"; Desc = "Run authentication tests on tenant-owned and critical external apps to confirm they handle MFA challenges." },
  @{ Num = "3"; Title = "Update custom apps"; Desc = "Ensure apps use MSAL and support Conditional Access claims challenges. See aka.ms/CAforLowValueScopes" },
  @{ Num = "4"; Title = "Enable monitoring"; Desc = "Configure sign-in log alerts to catch authentication failures post-enforcement." }
)

foreach ($act in $actions) {
  $ws.Row($row).Height = 40
  $ws.Cells["C${row}:E${row}"].Merge = $true

  $ws.Cells[$row, 2].Value = "  $($act.Num).  $($act.Title)"
  Set-Font $ws.Cells[$row, 2] -size 10 -bold $true -color $CHARCOAL
  $ws.Cells[$row, 2].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 2].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  Set-Fill $ws.Cells[$row, 2] $VERY_LIGHT_BLUE
  Set-ThinBorder $ws.Cells[$row, 2]

  $ws.Cells[$row, 3].Value = $act.Desc
  Set-Font $ws.Cells[$row, 3] -size 10 -color $CHARCOAL
  $ws.Cells[$row, 3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 3].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $ws.Cells[$row, 3].Style.WrapText = $true
  $row++
}

# References
$row++
Add-SectionHeader $ws $row "REFERENCES"
$row++
$ws.Row($row).Height = 6
$row++

$references = @(
  @{ Label = "Microsoft Announcement"; URL = "https://techcommunity.microsoft.com/blog/microsoft-entra-blog/upcoming-conditional-access-change-improved-enforcement-for-policies-with-resour/4488925" },
  @{ Label = "Detailed Documentation"; URL = "https://aka.ms/CAforLowValueScopes" },
  @{ Label = "Developer Guide"; URL = "https://learn.microsoft.com/en-us/entra/identity-platform/v2-conditional-access-dev-guide" }
)

foreach ($ref in $references) {
  $ws.Cells["C${row}:E${row}"].Merge = $true
  $ws.Row($row).Height = 20

  $ws.Cells[$row, 2].Value = $ref.Label
  Set-Font $ws.Cells[$row, 2] -size 10 -bold $true -color $CHARCOAL
  $ws.Cells[$row, 2].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 2].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  $ws.Cells[$row, 3].Value = $ref.URL
  Set-Font $ws.Cells[$row, 3] -size 10 -color $ACCENT_BLUE -underline $true
  $ws.Cells[$row, 3].Hyperlink = [System.Uri]::new($ref.URL)
  $ws.Cells[$row, 3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells[$row, 3].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $row++
}

# ═══════════════════════════════════════════
# APP SHEET BUILDER
# ═══════════════════════════════════════════
function Build-AppSheet {
  param(
    [OfficeOpenXml.ExcelPackage]$Package,
    [string]$SheetName,
    [string]$Title,
    [array]$Data,
    [System.Drawing.Color]$TabColor
  )

  $ws = $Package.Workbook.Worksheets.Add($SheetName)
  $ws.TabColor = $TabColor
  $ws.View.ShowGridLines = $false

  $colDefs = @(
    @{ Header = "Application Name";       Width = 32 },
    @{ Header = "App ID";                 Width = 38 },
    @{ Header = "Delegated Scopes";       Width = 44 },
    @{ Header = "Sign-ins ($($AuditDays)d)"; Width = 14 },
    @{ Header = "MFA ($($AuditDays)d)";   Width = 12 },
    @{ Header = "SFA ($($AuditDays)d)";   Width = 12 },
    @{ Header = "Risk";                   Width = 12 },
    @{ Header = "Action Required";        Width = 38 }
  )

  $lastCol = $colDefs.Count + 1  # +1 because data starts at col B

  # Column A spacer
  $ws.Column(1).Width = 2
  for ($i = 0; $i -lt $colDefs.Count; $i++) {
    $ws.Column($i + 2).Width = $colDefs[$i].Width
  }

  # Title banner (rows 1-3)
  for ($r = 1; $r -le 3; $r++) {
    for ($c = 1; $c -le $lastCol; $c++) {
      Set-Fill $ws.Cells[$r, $c] $NAVY
    }
  }
  $ws.Row(1).Height = 8
  $ws.Row(2).Height = 32
  $ws.Row(3).Height = 8

  $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $lastCol).Address -replace '[0-9]',''
  $ws.Cells["B2:${lastColLetter}2"].Merge = $true
  $ws.Cells["B2"].Value = $Title
  Set-Font $ws.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws.Cells["B2"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $ws.Cells["B2"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  # Summary bar (row 4)
  $ws.Row(4).Height = 28
  for ($c = 1; $c -le $lastCol; $c++) {
    Set-Fill $ws.Cells[4, $c] $LIGHT_BLUE
  }
  $ws.Cells["B4:${lastColLetter}4"].Merge = $true

  $high   = @($Data | Where-Object { $_.Risk -eq "HIGH" }).Count
  $med    = @($Data | Where-Object { $_.Risk -eq "MEDIUM" }).Count
  $low    = @($Data | Where-Object { $_.Risk -eq "LOW" }).Count
  $active = @($Data | Where-Object { $_.Signins -gt 0 }).Count

  $ws.Cells["B4"].Value = "Total: $($Data.Count)    |    Active: $active    |    High: $high    |    Medium: $med    |    Low: $low"
  Set-Font $ws.Cells["B4"] -size 10 -bold $true -color $NAVY
  $ws.Cells["B4"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $ws.Cells["B4"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  # Header row
  $ws.Row(5).Height = 6
  $hdrRow = 6
  $ws.Row($hdrRow).Height = 32

  for ($i = 0; $i -lt $colDefs.Count; $i++) {
    $col = $i + 2
    $cell = $ws.Cells[$hdrRow, $col]
    $cell.Value = $colDefs[$i].Header
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
    Set-Fill $cell $DARK_BLUE
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
    $cell.Style.Border.Bottom.Color.SetColor($NAVY)
  }

  # Data rows
  $riskOrder = @{ "HIGH" = 0; "UNKNOWN" = 1; "MEDIUM" = 2; "LOW" = 3 }
  $sortedData = @($Data | Sort-Object { $riskOrder[$_.Risk] }, { -$_.Signins })

  $r = $hdrRow + 1
  $idx = 0
  foreach ($app in $sortedData) {
    $ws.Row($r).Height = 26
    $stripe = if ($idx % 2 -eq 0) { $WHITE_C } else { $VERY_LIGHT_BLUE }

    $action = switch ($app.Risk) {
      "HIGH"    { "SFA sign-ins detected - will break when MFA enforced" }
      "UNKNOWN" { "Audit log unavailable - verify MFA status manually" }
      "MEDIUM"  { "All sign-ins via MFA - monitor after enforcement" }
      default   { "No recent sign-ins - lower priority" }
    }

    $mfaVal = if ($app.MFA -eq -1) { "N/A" } else { $app.MFA }
    $sfaVal = if ($app.MFA -eq -1) { "N/A" } else { $app.SFA }

    $values = @($app.AppName, $app.AppId, $app.Scopes, $app.Signins, $mfaVal, $sfaVal, $app.Risk, $action)

    for ($i = 0; $i -lt $values.Count; $i++) {
      $col = $i + 2
      $cell = $ws.Cells[$r, $col]
      $cell.Value = $values[$i]
      Set-Font $cell -size 10 -color $CHARCOAL
      Set-Fill $cell $stripe

      if ($i -lt 3 -or $i -eq 7) {
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $cell.Style.WrapText = $true
      } else {
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
      }
      $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
      $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
      $cell.Style.Border.Bottom.Color.SetColor($MED_GRAY)

      # Risk badge
      if ($i -eq 6) {
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        switch ($values[$i]) {
          "HIGH"    { Set-Fill $cell $RED_BG_C;   Set-Font $cell -size 10 -bold $true -color $RED_TEXT_C }
          "UNKNOWN" { Set-Fill $cell $GRAY_BG_C;  Set-Font $cell -size 10 -bold $true -color $GRAY_TEXT_C }
          "MEDIUM"  { Set-Fill $cell $AMBER_BG_C; Set-Font $cell -size 10 -bold $true -color $AMBER_TEXT_C }
          "LOW"     { Set-Fill $cell $GREEN_BG_C; Set-Font $cell -size 10 -bold $true -color $GREEN_TEXT_C }
        }
      }

      # Bold non-zero sign-in numbers
      if ($i -in @(3, 4, 5) -and $values[$i] -is [int] -and $values[$i] -gt 0) {
        Set-Font $cell -size 10 -bold $true -color $CHARCOAL
      }
      # Red SFA count
      if ($i -eq 5 -and $values[$i] -is [int] -and $values[$i] -gt 0) {
        Set-Font $cell -size 10 -bold $true -color $RED_TEXT_C
      }
    }
    $r++
    $idx++
  }

  # Freeze panes and auto-filter
  $ws.View.FreezePanes($hdrRow + 1, 2)
  if ($sortedData.Count -gt 0) {
    $ws.Cells["B${hdrRow}:${lastColLetter}$($r - 1)"].AutoFilter = $true
  }
}

# ═══════════════════════════════════════════
# Build app sheets
# ═══════════════════════════════════════════
Build-AppSheet -Package $excel -SheetName "Tenant-Owned Apps" -Title "TENANT-OWNED APPLICATIONS" -Data $tenantOwnedData -TabColor (HexColor "2E75B6")
Build-AppSheet -Package $excel -SheetName "External Apps" -Title "EXTERNAL APPLICATIONS" -Data $externalData -TabColor (HexColor "E67E22")

# ═══════════════════════════════════════════
# SHEET 4: SCOPES REFERENCE
# ═══════════════════════════════════════════
$ws4 = $excel.Workbook.Worksheets.Add("Scopes Reference")
$ws4.TabColor = HexColor "7F8C8D"
$ws4.View.ShowGridLines = $false
$ws4.Column(1).Width = 2
$ws4.Column(2).Width = 28
$ws4.Column(3).Width = 70

# Title banner
for ($r = 1; $r -le 3; $r++) {
  for ($c = 1; $c -le 3; $c++) {
    Set-Fill $ws4.Cells[$r, $c] $NAVY
  }
}
$ws4.Row(1).Height = 8
$ws4.Row(2).Height = 32
$ws4.Row(3).Height = 8
$ws4.Cells["B2:C2"].Merge = $true
$ws4.Cells["B2"].Value = "IN-SCOPE DELEGATED PERMISSIONS"
Set-Font $ws4.Cells["B2"] -size 16 -bold $true -color $WHITE_C
$ws4.Cells["B2"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$ws4.Cells["B2"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Subtitle
$ws4.Row(4).Height = 28
for ($c = 1; $c -le 3; $c++) {
  Set-Fill $ws4.Cells[4, $c] $LIGHT_BLUE
}
$ws4.Cells["B4:C4"].Merge = $true
$ws4.Cells["B4"].Value = "These scopes define which apps are in scope of the May 13 CA enforcement change"
Set-Font $ws4.Cells["B4"] -size 10 -bold $true -color $NAVY
$ws4.Cells["B4"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$ws4.Cells["B4"].Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Headers
$ws4.Row(5).Height = 6
$hdrRow = 6
$ws4.Row($hdrRow).Height = 28

foreach ($hdr in @(@{Col=2;Text="Scope"}, @{Col=3;Text="Description"})) {
  $cell = $ws4.Cells[$hdrRow, $hdr.Col]
  $cell.Value = $hdr.Text
  Set-Font $cell -size 10 -bold $true -color $WHITE_C
  Set-Fill $cell $DARK_BLUE
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}

# Scope data
$scopesInfo = @(
  @{ Scope = "openid";              Desc = "Core OIDC scope - requests an ID token" },
  @{ Scope = "profile";             Desc = "Access to user's basic profile (name, picture, etc.)" },
  @{ Scope = "email";               Desc = "Access to user's email address" },
  @{ Scope = "offline_access";      Desc = "Requests a refresh token for long-lived access" },
  @{ Scope = "User.Read";           Desc = "Read the signed-in user's basic profile" },
  @{ Scope = "User.Read.All";       Desc = "Read all users' basic profiles (delegated)" },
  @{ Scope = "User.ReadBasic.All";  Desc = "Read all users' basic properties" },
  @{ Scope = "People.Read";         Desc = "Read the signed-in user's relevant people list" },
  @{ Scope = "People.Read.All";     Desc = "Read all users' relevant people lists" },
  @{ Scope = "GroupMember.Read.All"; Desc = "Read all group memberships" },
  @{ Scope = "Member.Read.Hidden";  Desc = "Read hidden group memberships" }
)

$r = 7
$idx = 0
foreach ($s in $scopesInfo) {
  $stripe = if ($idx % 2 -eq 0) { $WHITE_C } else { $VERY_LIGHT_BLUE }
  $ws4.Row($r).Height = 24

  $cellScope = $ws4.Cells[$r, 2]
  $cellScope.Value = $s.Scope
  Set-Font $cellScope -size 10 -bold $true -color $CHARCOAL
  Set-Fill $cellScope $stripe
  $cellScope.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $cellScope.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $cellScope.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
  $cellScope.Style.Border.Bottom.Color.SetColor($MED_GRAY)

  $cellDesc = $ws4.Cells[$r, 3]
  $cellDesc.Value = $s.Desc
  Set-Font $cellDesc -size 10 -color $CHARCOAL
  Set-Fill $cellDesc $stripe
  $cellDesc.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $cellDesc.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $cellDesc.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
  $cellDesc.Style.Border.Bottom.Color.SetColor($MED_GRAY)

  $r++
  $idx++
}

# ── Save workbook ──
$excel.SaveAs($exportPath)
$excel.Dispose()

$script:swPhase.Stop()
$swTotal.Stop()
$totalTime = $swTotal.Elapsed.ToString("mm\:ss")

$highCount = $highRisk.Count
$medCount  = $medRisk.Count
$lowCount  = $lowRisk.Count

# ── Disconnect from Graph ──
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "     ASSESSMENT COMPLETE" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Report:  $exportPath" -ForegroundColor White
Write-Host "  Tenant:  $TenantId" -ForegroundColor Gray
Write-Host "  Period:  $AuditDays days" -ForegroundColor Gray
Write-Host "  Time:    $totalTime" -ForegroundColor Gray
Write-Host ""
Write-Host "  Results:" -ForegroundColor White
Write-Host "    Total apps:     $($exportData.Count)" -ForegroundColor White
Write-Host "    Tenant-owned:   $($tenantOwnedData.Count)" -ForegroundColor White
Write-Host "    External:       $($externalData.Count)" -ForegroundColor White
if ($highCount -gt 0) {
  Write-Host "    HIGH risk:      $highCount" -ForegroundColor Red
} else {
  Write-Host "    HIGH risk:      $highCount" -ForegroundColor Green
}
Write-Host "    MEDIUM risk:    $medCount" -ForegroundColor Yellow
Write-Host "    LOW risk:       $lowCount" -ForegroundColor Gray
Write-Host ""
Write-Host "  Open the .xlsx file in Excel for full details." -ForegroundColor Cyan
Write-Host ""
