# ==============================
# Entra ID App Credentials Audit
# Identifies expired and soon-to-expire app registration
# certificates and client secrets.
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
    Write-Host "  Then run:  pwsh .\app-credentials-audit.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./app-credentials-audit.ps1" -ForegroundColor Yellow
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
     Entra ID App Credentials Audit
     Expired & expiring certificates and client secrets
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

# ── Step 2: Expiry warning threshold ──
Write-Host ""
Write-Host "  Select expiry warning threshold:" -ForegroundColor White
Write-Host "    [1]  30 days  (default)" -ForegroundColor Green
Write-Host "    [2]  60 days" -ForegroundColor Yellow
Write-Host "    [3]  90 days" -ForegroundColor Yellow
Write-Host "    [4]  Custom" -ForegroundColor Yellow
Write-Host ""
$thresholdChoice = Read-Host "  Enter choice (1-4) [default: 1]"

switch ($thresholdChoice) {
  "2" { $WarningDays = 60 }
  "3" { $WarningDays = 90 }
  "4" {
    $customDays = Read-Host "  Enter number of days"
    if ($customDays -match '^\d+$' -and [int]$customDays -gt 0) {
      $WarningDays = [int]$customDays
    } else {
      Write-Host "  [!] Invalid number. Using 30 days." -ForegroundColor Red
      $WarningDays = 30
    }
  }
  default { $WarningDays = 30 }
}
Write-Host "  -> Warning threshold: $WarningDays days" -ForegroundColor Green

# ── Summary & confirm ──
Write-Host ""
Write-Host "  ──────────────────────────────────────" -ForegroundColor Gray
Write-Host "  Warning threshold:  $WarningDays days" -ForegroundColor White
Write-Host "  Flow:               Audit -> Review -> Optionally remove expired credentials" -ForegroundColor Gray
Write-Host "  ──────────────────────────────────────" -ForegroundColor Gray
Write-Host ""

$confirm = Read-Host "  Press ENTER to start or type 'q' to quit"
if ($confirm -and $confirm.Trim().ToLower() -ne '') {
  Write-Host "  Cancelled." -ForegroundColor Red
  return
}

# ── Overall timer ──
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$now = Get-Date

# ── Phase 1: Connect ──
$totalPhases = 4
$phase = 1

Write-Host ""
Write-Host "  [$phase/$totalPhases] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$requiredScopes = @(
  "Application.Read.All",
  "Application.ReadWrite.All",
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

$tenantId = $ctx.TenantId
Write-Host "  Connected to tenant: $tenantId  ($($ctx.Account))" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 2: Fetch all app registrations ──
$phase++
Write-Host "  [$phase/$totalPhases] Fetching app registrations..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$allApps = @()
$appUri = "https://graph.microsoft.com/v1.0/applications?`$select=id,appId,displayName,keyCredentials,passwordCredentials&`$top=999"

while ($appUri) {
  $resp = Invoke-MgGraphRequest -Method GET -Uri $appUri -ErrorAction Stop
  if ($resp.value) {
    $allApps += $resp.value
  }
  $appUri = $resp.'@odata.nextLink'
  if ($appUri) {
    Write-ProgressBar -Current $allApps.Count -Total ($allApps.Count + 999) -Activity "Fetching apps" -Status "$($allApps.Count) fetched"
  }
}

Write-Progress -Activity "Fetching apps" -Completed
Write-Host "  Total app registrations: $($allApps.Count)" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 3: Analyze credentials ──
$phase++
Write-Host "  [$phase/$totalPhases] Analyzing certificates and secrets..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$warningDate = $now.AddDays($WarningDays)
$credentials = @()

for ($i = 0; $i -lt $allApps.Count; $i++) {
  $app = $allApps[$i]
  Write-ProgressBar -Current ($i + 1) -Total $allApps.Count -Activity "Analyzing credentials" -Status $app.displayName

  $portalUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/$($app.appId)/isMSAApp~/false"

  # Check certificates (keyCredentials)
  if ($app.keyCredentials) {
    foreach ($cert in $app.keyCredentials) {
      $endDate = $null
      $startDate = $null
      try {
        $endDate = [datetime]::Parse($cert.endDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
        $startDate = [datetime]::Parse($cert.startDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
      } catch { continue }

      $daysUntilExpiry = [math]::Round(($endDate - $now).TotalDays)

      $status = if ($daysUntilExpiry -lt 0) {
        "Expired"
      } elseif ($daysUntilExpiry -le $WarningDays) {
        "Expiring Soon"
      } else {
        "Valid"
      }

      if ($status -eq "Valid") { continue }

      $credentials += [PSCustomObject]@{
        AppName       = $app.displayName
        AppId         = $app.appId
        ObjectId      = $app.id
        CredentialType = "Certificate"
        Description   = if ($cert.displayName) { $cert.displayName } else { "N/A" }
        KeyId         = $cert.keyId
        StartDate     = $startDate.ToString('yyyy-MM-dd')
        EndDate       = $endDate.ToString('yyyy-MM-dd')
        DaysToExpiry  = $daysUntilExpiry
        Status        = $status
        PortalUrl     = $portalUrl
      }
    }
  }

  # Check client secrets (passwordCredentials)
  if ($app.passwordCredentials) {
    foreach ($secret in $app.passwordCredentials) {
      $endDate = $null
      $startDate = $null
      try {
        $endDate = [datetime]::Parse($secret.endDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
        $startDate = [datetime]::Parse($secret.startDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
      } catch { continue }

      $daysUntilExpiry = [math]::Round(($endDate - $now).TotalDays)

      $status = if ($daysUntilExpiry -lt 0) {
        "Expired"
      } elseif ($daysUntilExpiry -le $WarningDays) {
        "Expiring Soon"
      } else {
        "Valid"
      }

      if ($status -eq "Valid") { continue }

      $credentials += [PSCustomObject]@{
        AppName       = $app.displayName
        AppId         = $app.appId
        ObjectId      = $app.id
        CredentialType = "Client Secret"
        Description   = if ($secret.displayName) { $secret.displayName } else { "N/A" }
        KeyId         = $secret.keyId
        StartDate     = $startDate.ToString('yyyy-MM-dd')
        EndDate       = $endDate.ToString('yyyy-MM-dd')
        DaysToExpiry  = $daysUntilExpiry
        Status        = $status
        PortalUrl     = $portalUrl
      }
    }
  }
}

Write-Progress -Activity "Analyzing credentials" -Completed

# Sort: expired first (most negative days), then expiring soon
$credentials = $credentials | Sort-Object -Property DaysToExpiry

$expiredCreds  = @($credentials | Where-Object { $_.Status -eq "Expired" })
$expiringCreds = @($credentials | Where-Object { $_.Status -eq "Expiring Soon" })

# Unique apps affected
$expiredApps  = @($expiredCreds | Select-Object -Property AppId -Unique).Count
$expiringApps = @($expiringCreds | Select-Object -Property AppId -Unique).Count

Write-Host ""
Write-Host "  Results:" -ForegroundColor White
Write-Host "    Expired credentials:       $($expiredCreds.Count) (across $expiredApps apps)" -ForegroundColor $(if ($expiredCreds.Count -gt 0) { "Red" } else { "Green" })
Write-Host "    Expiring within $WarningDays days:  $($expiringCreds.Count) (across $expiringApps apps)" -ForegroundColor $(if ($expiringCreds.Count -gt 0) { "Yellow" } else { "Green" })

# Credential type breakdown
$certCount   = @($credentials | Where-Object { $_.CredentialType -eq "Certificate" }).Count
$secretCount = @($credentials | Where-Object { $_.CredentialType -eq "Client Secret" }).Count
Write-Host ""
Write-Host "  Breakdown: $certCount certificates, $secretCount client secrets" -ForegroundColor Gray

$script:swPhase.Stop()

if ($credentials.Count -eq 0) {
  Write-Host ""
  Write-Host "  No expired or expiring credentials found. Your apps are healthy." -ForegroundColor Green
  $swTotal.Stop()
  Write-Host ""
  Write-Host "  Total time: $($swTotal.Elapsed.ToString('mm\:ss'))" -ForegroundColor Gray
  return
}

# ── Phase 4: Enrich with app owners ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Fetching app owners..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Get unique app object IDs to avoid duplicate lookups
$uniqueAppIds = $credentials | Select-Object -Property ObjectId -Unique
$ownerCache = @{}

for ($i = 0; $i -lt $uniqueAppIds.Count; $i++) {
  $objId = $uniqueAppIds[$i].ObjectId
  Write-ProgressBar -Current ($i + 1) -Total $uniqueAppIds.Count -Activity "Fetching owners" -Status "App $($i + 1) of $($uniqueAppIds.Count)"

  try {
    $owners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$objId/owners?`$select=displayName,userPrincipalName" -ErrorAction Stop
    if ($owners.value -and $owners.value.Count -gt 0) {
      $ownerNames = ($owners.value | ForEach-Object {
        if ($_.displayName) { $_.displayName } elseif ($_.userPrincipalName) { $_.userPrincipalName } else { "Unknown" }
      }) -join ", "
      $ownerCache[$objId] = $ownerNames
    } else {
      $ownerCache[$objId] = "No owner"
    }
  } catch {
    $ownerCache[$objId] = "N/A"
  }
}

# Apply owners to credentials
foreach ($cred in $credentials) {
  $cred | Add-Member -NotePropertyName "Owner" -NotePropertyValue ($ownerCache[$cred.ObjectId]) -Force
}

Write-Progress -Activity "Fetching owners" -Completed
Write-Host "  Owners resolved for $($uniqueAppIds.Count) apps" -ForegroundColor Green
$script:swPhase.Stop()

# ── Post-audit: Ask what to do with expired credentials ──
$actionResults = @()

if ($expiredCreds.Count -gt 0) {
  Write-Host ""
  Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
  Write-Host "  Found $($expiredCreds.Count) expired credential(s). What would you like to do?" -ForegroundColor White
  Write-Host ""
  Write-Host "    [1]  Export report only — no changes (default)" -ForegroundColor Green
  Write-Host "    [2]  Remove expired credentials, then export" -ForegroundColor Red
  Write-Host ""
  $postChoice = Read-Host "  Enter choice (1-2) [default: 1]"

  $ActionMode = if ($postChoice -eq "2") { "Remove" } else { "Audit" }

  if ($ActionMode -eq "Remove") {
    Write-Host ""
    Write-Host "  The following $($expiredCreds.Count) expired credential(s) will be removed:" -ForegroundColor White
    Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Gray

    $preview = $expiredCreds | Select-Object -First 20
    foreach ($c in $preview) {
      $daysAgo = [math]::Abs($c.DaysToExpiry)
      Write-Host "    $($c.AppName)  |  $($c.CredentialType)  |  $($c.Description)  |  expired $daysAgo days ago" -ForegroundColor Gray
    }
    if ($expiredCreds.Count -gt 20) {
      Write-Host "    ... and $($expiredCreds.Count - 20) more" -ForegroundColor Gray
    }

    Write-Host "  ──────────────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host ""

    $confirmAction = Read-Host "  Type 'YES' to confirm removal of $($expiredCreds.Count) expired credential(s)"
    if ($confirmAction -ne 'YES') {
      Write-Host "  Action cancelled. Exporting report only." -ForegroundColor Yellow
      $ActionMode = "Audit"
    } else {
      Write-Host ""
      Write-Host "  Removing expired credentials..." -ForegroundColor Red
      $script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()
      $successCount = 0
      $failCount = 0

      for ($i = 0; $i -lt $expiredCreds.Count; $i++) {
        $c = $expiredCreds[$i]
        Write-ProgressBar -Current ($i + 1) -Total $expiredCreds.Count -Activity "Removing credentials" -Status "$($c.AppName) — $($c.CredentialType)"

        try {
          if ($c.CredentialType -eq "Client Secret") {
            $body = @{ passwordCredentialId = $c.KeyId } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($c.ObjectId)/removePassword" -Body $body -ContentType "application/json" -ErrorAction Stop
          } elseif ($c.CredentialType -eq "Certificate") {
            $body = @{ keyCredentialId = $c.KeyId } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($c.ObjectId)/removeKey" -Body $body -ContentType "application/json" -ErrorAction Stop
          }
          $successCount++
          $actionResults += [PSCustomObject]@{
            AppName    = $c.AppName
            AppId      = $c.AppId
            Type       = $c.CredentialType
            KeyId      = $c.KeyId
            Action     = "Remove"
            Result     = "Success"
          }
        } catch {
          $failCount++
          $actionResults += [PSCustomObject]@{
            AppName    = $c.AppName
            AppId      = $c.AppId
            Type       = $c.CredentialType
            KeyId      = $c.KeyId
            Action     = "Remove"
            Result     = "Failed: $($_.Exception.Message)"
          }
        }
      }

      Write-Progress -Activity "Removing credentials" -Completed
      Write-Host ""
      Write-Host "  Removal complete: $successCount succeeded, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
      $script:swPhase.Stop()
    }
  }
} else {
  $ActionMode = "Audit"
}

# ── Export to Excel ──
Write-Host ""
Write-Host "  Generating Excel report..." -ForegroundColor Cyan

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$desktopPath = [Environment]::GetFolderPath("Desktop")
$exportPath = Join-Path $desktopPath "AppCredentials_Audit_$timestamp.xlsx"

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

$ws.Column(1).Width = 2
$ws.Column(2).Width = 25
$ws.Column(3).Width = 25
$ws.Column(4).Width = 25
$ws.Column(5).Width = 25
$ws.Column(6).Width = 25

# Title bar
$ws.Row(1).Height = 8
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[1, $c] $NAVY }
$ws.Row(2).Height = 36
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[2, $c] $NAVY }
$ws.Cells["B2:F2"].Merge = $true
$ws.Cells["B2"].Value = "APP CREDENTIALS AUDIT"
Set-Font $ws.Cells["B2"] -size 18 -bold $true -color $WHITE_C
$ws.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Subtitle
$ws.Row(3).Height = 24
for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[3, $c] $DARK_BLUE }
$ws.Cells["B3:F3"].Merge = $true
$ws.Cells["B3"].Value = "Tenant: $tenantId  |  Warning: $WarningDays days  |  Action: $ActionMode  |  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Set-Font $ws.Cells["B3"] -size 10 -color (HexColor "B0C4DE")

# KPI row
$ws.Row(4).Height = 8
$ws.Row(5).Height = 60
$kpis = @(
  @{ Col = 2; Number = "$($allApps.Count)"; Label = "TOTAL APPS" },
  @{ Col = 3; Number = "$($credentials.Count)"; Label = "FLAGGED CREDENTIALS" },
  @{ Col = 4; Number = "$($expiredCreds.Count)"; Label = "EXPIRED" },
  @{ Col = 5; Number = "$($expiringCreds.Count)"; Label = "EXPIRING SOON" }
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

# Type breakdown
$row = 7
$ws.Cells["B$row"].Value = "CREDENTIAL TYPE BREAKDOWN"
Set-Font $ws.Cells["B$row"] -size 12 -bold $true -color $NAVY
$row++

$ws.Cells["B$row"].Value = "Type"
$ws.Cells["C$row"].Value = "Expired"
$ws.Cells["D$row"].Value = "Expiring Soon"
$ws.Cells["E$row"].Value = "Total Flagged"
Set-Font $ws.Cells["B$row"] -size 10 -bold $true -color $WHITE_C
Set-Font $ws.Cells["C$row"] -size 10 -bold $true -color $WHITE_C
Set-Font $ws.Cells["D$row"] -size 10 -bold $true -color $WHITE_C
Set-Font $ws.Cells["E$row"] -size 10 -bold $true -color $WHITE_C
for ($c = 2; $c -le 5; $c++) { Set-Fill $ws.Cells[$row, $c] $ACCENT_BLUE }
$row++

foreach ($type in @("Certificate", "Client Secret")) {
  $typeExpired  = @($expiredCreds | Where-Object { $_.CredentialType -eq $type }).Count
  $typeExpiring = @($expiringCreds | Where-Object { $_.CredentialType -eq $type }).Count
  $ws.Cells["B$row"].Value = $type
  $ws.Cells["C$row"].Value = $typeExpired
  $ws.Cells["D$row"].Value = $typeExpiring
  $ws.Cells["E$row"].Value = ($typeExpired + $typeExpiring)
  for ($c = 2; $c -le 5; $c++) {
    Set-ThinBorder $ws.Cells[$row, $c]
    $ws.Cells[$row, $c].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  }
  $ws.Cells["B$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
  $row++
}

# ── Sheet 2: Expired Credentials ──
if ($expiredCreds.Count -gt 0) {
  $ws2 = $excel.Workbook.Worksheets.Add("Expired")
  $ws2.TabColor = HexColor "E74C3C"
  $ws2.View.ShowGridLines = $false

  $ws2.Column(1).Width = 2
  $ws2.Column(2).Width = 28
  $ws2.Column(3).Width = 15
  $ws2.Column(4).Width = 22
  $ws2.Column(5).Width = 14
  $ws2.Column(6).Width = 14
  $ws2.Column(7).Width = 16
  $ws2.Column(8).Width = 22
  $ws2.Column(9).Width = 40

  $ws2.Row(1).Height = 8
  for ($c = 1; $c -le 9; $c++) { Set-Fill $ws2.Cells[1, $c] $NAVY }
  $ws2.Row(2).Height = 32
  for ($c = 1; $c -le 9; $c++) { Set-Fill $ws2.Cells[2, $c] $NAVY }
  $ws2.Cells["B2:I2"].Merge = $true
  $ws2.Cells["B2"].Value = "EXPIRED CREDENTIALS"
  Set-Font $ws2.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws2.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  $hdrRow = 4
  $headers = @(
    @{ Col = 2; Name = "App Name" },
    @{ Col = 3; Name = "Type" },
    @{ Col = 4; Name = "Description" },
    @{ Col = 5; Name = "Start Date" },
    @{ Col = 6; Name = "End Date" },
    @{ Col = 7; Name = "Days Expired" },
    @{ Col = 8; Name = "Owner" },
    @{ Col = 9; Name = "Portal Link" }
  )

  $ws2.Row($hdrRow).Height = 28
  foreach ($h in $headers) {
    $cell = $ws2.Cells[$hdrRow, $h.Col]
    $cell.Value = $h.Name
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
    $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  }

  $dataRow = $hdrRow + 1
  foreach ($c in $expiredCreds) {
    $ws2.Row($dataRow).Height = 22

    $ws2.Cells[$dataRow, 2].Value = $c.AppName
    $ws2.Cells[$dataRow, 3].Value = $c.CredentialType
    $ws2.Cells[$dataRow, 4].Value = $c.Description
    $ws2.Cells[$dataRow, 5].Value = $c.StartDate
    $ws2.Cells[$dataRow, 6].Value = $c.EndDate
    $ws2.Cells[$dataRow, 7].Value = [math]::Abs($c.DaysToExpiry)
    $ws2.Cells[$dataRow, 8].Value = $c.Owner

    # Hyperlink to Entra portal
    $ws2.Cells[$dataRow, 9].Value = "Open in Entra"
    $ws2.Cells[$dataRow, 9].Hyperlink = [System.Uri]::new($c.PortalUrl)
    $ws2.Cells[$dataRow, 9].Style.Font.UnderLine = $true
    Set-Font $ws2.Cells[$dataRow, 9] -color $ACCENT_BLUE

    # Color days expired
    $daysCell = $ws2.Cells[$dataRow, 7]
    $daysCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    Set-Fill $daysCell $RED_BG
    Set-Font $daysCell -color $RED_TEXT

    for ($col = 2; $col -le 9; $col++) {
      Set-ThinBorder $ws2.Cells[$dataRow, $col]
    }

    if ($dataRow % 2 -eq 0) {
      for ($col = 2; $col -le 9; $col++) {
        if ($col -ne 7) {
          Set-Fill $ws2.Cells[$dataRow, $col] (HexColor "F8F9FA")
        }
      }
    }

    $dataRow++
  }
}

# ── Sheet 3: Expiring Soon ──
if ($expiringCreds.Count -gt 0) {
  $ws3 = $excel.Workbook.Worksheets.Add("Expiring Soon")
  $ws3.TabColor = HexColor "E67E22"
  $ws3.View.ShowGridLines = $false

  $ws3.Column(1).Width = 2
  $ws3.Column(2).Width = 28
  $ws3.Column(3).Width = 15
  $ws3.Column(4).Width = 22
  $ws3.Column(5).Width = 14
  $ws3.Column(6).Width = 14
  $ws3.Column(7).Width = 16
  $ws3.Column(8).Width = 22
  $ws3.Column(9).Width = 40

  $ws3.Row(1).Height = 8
  for ($c = 1; $c -le 9; $c++) { Set-Fill $ws3.Cells[1, $c] $NAVY }
  $ws3.Row(2).Height = 32
  for ($c = 1; $c -le 9; $c++) { Set-Fill $ws3.Cells[2, $c] $NAVY }
  $ws3.Cells["B2:I2"].Merge = $true
  $ws3.Cells["B2"].Value = "EXPIRING SOON (within $WarningDays days)"
  Set-Font $ws3.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws3.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

  $hdrRow = 4
  $headers = @(
    @{ Col = 2; Name = "App Name" },
    @{ Col = 3; Name = "Type" },
    @{ Col = 4; Name = "Description" },
    @{ Col = 5; Name = "Start Date" },
    @{ Col = 6; Name = "End Date" },
    @{ Col = 7; Name = "Days Left" },
    @{ Col = 8; Name = "Owner" },
    @{ Col = 9; Name = "Portal Link" }
  )

  $ws3.Row($hdrRow).Height = 28
  foreach ($h in $headers) {
    $cell = $ws3.Cells[$hdrRow, $h.Col]
    $cell.Value = $h.Name
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
    $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  }

  $dataRow = $hdrRow + 1
  foreach ($c in $expiringCreds) {
    $ws3.Row($dataRow).Height = 22

    $ws3.Cells[$dataRow, 2].Value = $c.AppName
    $ws3.Cells[$dataRow, 3].Value = $c.CredentialType
    $ws3.Cells[$dataRow, 4].Value = $c.Description
    $ws3.Cells[$dataRow, 5].Value = $c.StartDate
    $ws3.Cells[$dataRow, 6].Value = $c.EndDate
    $ws3.Cells[$dataRow, 7].Value = $c.DaysToExpiry
    $ws3.Cells[$dataRow, 8].Value = $c.Owner

    # Hyperlink to Entra portal
    $ws3.Cells[$dataRow, 9].Value = "Open in Entra"
    $ws3.Cells[$dataRow, 9].Hyperlink = [System.Uri]::new($c.PortalUrl)
    $ws3.Cells[$dataRow, 9].Style.Font.UnderLine = $true
    Set-Font $ws3.Cells[$dataRow, 9] -color $ACCENT_BLUE

    # Color days left
    $daysCell = $ws3.Cells[$dataRow, 7]
    $daysCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    if ($c.DaysToExpiry -le 7) {
      Set-Fill $daysCell $RED_BG
      Set-Font $daysCell -color $RED_TEXT
    } elseif ($c.DaysToExpiry -le 14) {
      Set-Fill $daysCell $AMBER_BG
      Set-Font $daysCell -color $AMBER_TEXT
    } else {
      Set-Fill $daysCell $GREEN_BG
      Set-Font $daysCell -color $GREEN_TEXT
    }

    for ($col = 2; $col -le 9; $col++) {
      Set-ThinBorder $ws3.Cells[$dataRow, $col]
    }

    if ($dataRow % 2 -eq 0) {
      for ($col = 2; $col -le 9; $col++) {
        if ($col -ne 7) {
          Set-Fill $ws3.Cells[$dataRow, $col] (HexColor "F8F9FA")
        }
      }
    }

    $dataRow++
  }
}

# ── Sheet 4: Action Log (if actions were taken) ──
if ($actionResults.Count -gt 0) {
  $ws4 = $excel.Workbook.Worksheets.Add("Action Log")
  $ws4.TabColor = HexColor "E74C3C"
  $ws4.View.ShowGridLines = $false

  $ws4.Column(1).Width = 2
  $ws4.Column(2).Width = 28
  $ws4.Column(3).Width = 15
  $ws4.Column(4).Width = 38
  $ws4.Column(5).Width = 12
  $ws4.Column(6).Width = 40

  $ws4.Row(1).Height = 8
  for ($c = 1; $c -le 6; $c++) { Set-Fill $ws4.Cells[1, $c] $NAVY }
  $ws4.Row(2).Height = 32
  for ($c = 1; $c -le 6; $c++) { Set-Fill $ws4.Cells[2, $c] $NAVY }
  $ws4.Cells["B2:F2"].Merge = $true
  $ws4.Cells["B2"].Value = "ACTION LOG — REMOVAL RESULTS"
  Set-Font $ws4.Cells["B2"] -size 16 -bold $true -color $WHITE_C

  $logHdr = 4
  $ws4.Row($logHdr).Height = 28
  $logHeaders = @(
    @{ Col = 2; Name = "App Name" },
    @{ Col = 3; Name = "Type" },
    @{ Col = 4; Name = "Key ID" },
    @{ Col = 5; Name = "Action" },
    @{ Col = 6; Name = "Result" }
  )
  foreach ($h in $logHeaders) {
    $cell = $ws4.Cells[$logHdr, $h.Col]
    $cell.Value = $h.Name
    Set-Fill $cell $ACCENT_BLUE
    Set-Font $cell -size 10 -bold $true -color $WHITE_C
  }

  $logRow = $logHdr + 1
  foreach ($r in $actionResults) {
    $ws4.Cells[$logRow, 2].Value = $r.AppName
    $ws4.Cells[$logRow, 3].Value = $r.Type
    $ws4.Cells[$logRow, 4].Value = $r.KeyId
    $ws4.Cells[$logRow, 5].Value = $r.Action
    $ws4.Cells[$logRow, 6].Value = $r.Result

    $resultCell = $ws4.Cells[$logRow, 6]
    if ($r.Result -eq "Success") {
      Set-Fill $resultCell $GREEN_BG
      Set-Font $resultCell -color $GREEN_TEXT
    } else {
      Set-Fill $resultCell $RED_BG
      Set-Font $resultCell -color $RED_TEXT
    }

    for ($c = 2; $c -le 6; $c++) {
      Set-ThinBorder $ws4.Cells[$logRow, $c]
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
Write-Host "  Total apps scanned:       $($allApps.Count)" -ForegroundColor White
Write-Host "  Expired credentials:      $($expiredCreds.Count)" -ForegroundColor White
Write-Host "  Expiring soon:            $($expiringCreds.Count)" -ForegroundColor White
if ($actionResults.Count -gt 0) {
  $successTotal = @($actionResults | Where-Object { $_.Result -eq "Success" }).Count
  $failTotal    = @($actionResults | Where-Object { $_.Result -ne "Success" }).Count
  Write-Host "  Removed:                  $successTotal succeeded, $failTotal failed" -ForegroundColor White
}
Write-Host "  Total time: $($swTotal.Elapsed.ToString('mm\:ss'))" -ForegroundColor Gray
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""
