# ==============================
# Entra ID App Registration Audit
# One scan that combines, per app registration:
#   - API permission risk        (High / Medium / Low, by what the permission can do)
#   - Credential health           (expired & expiring certs/secrets)
#   - Staleness                   (sign-in activity: Active / Stale / Never used / No SP)
#   - Owners, tenant vs external
#
# Audit-first; optional gated actions (remove expired creds / disable SP / delete app).
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
    Write-Host "  Then run:  pwsh .\entra-appregistration-audit.ps1" -ForegroundColor Yellow
  } else {
    Write-Host "  Install it with:  brew install powershell/tap/powershell" -ForegroundColor Yellow
    Write-Host "  Then run:  pwsh ./entra-appregistration-audit.ps1" -ForegroundColor Yellow
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

# ── Helper: Invoke-GraphPaged ──
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

# ── Helper: Invoke-GraphBatch ──
# Runs many GET requests in parallel via Graph JSON batching (20 per round trip)
# instead of one HTTP call each. Returns a hashtable: request id -> sub-response body
# (a hashtable with .value for collections, or the object itself for single-item GETs;
# $null on failure). Chunks by 20 and retries throttled / 5xx sub-requests with backoff.
function Invoke-GraphBatch {
  param([object[]]$Requests, [string]$Activity = "Batch")   # each item: @{ id=<string>; url=<relative GET url> }
  $results = @{}
  $items = @($Requests)
  if ($items.Count -eq 0) { return $results }
  $done = 0
  for ($i = 0; $i -lt $items.Count; $i += 20) {
    $chunk = @($items[$i..([math]::Min($i + 19, $items.Count - 1))])
    $attempt = 0
    while ($chunk.Count -gt 0) {
      $attempt++
      $payload = @{ requests = @($chunk | ForEach-Object { @{ id = "$($_.id)"; method = 'GET'; url = $_.url } }) }
      try {
        $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body ($payload | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop
      } catch {
        if ($attempt -ge 4) { foreach ($c in $chunk) { $results["$($c.id)"] = $null }; break }
        Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $attempt))); continue
      }
      $retry = @()
      foreach ($r in @($resp.responses)) {
        $rid = "$($r.id)"; $status = [int]$r.status
        if (($status -eq 429 -or $status -ge 500) -and $attempt -lt 4) {
          $retry += ($chunk | Where-Object { "$($_.id)" -eq $rid })
        } else {
          $results[$rid] = $r.body
        }
      }
      $chunk = @($retry)
      if ($chunk.Count -gt 0) { Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $attempt))) }
    }
    $done += 20
    Write-ProgressBar -Current ([math]::Min($done, $items.Count)) -Total $items.Count -Activity $Activity -Status "batched"
  }
  Write-Progress -Activity $Activity -Completed
  return $results
}

# ── Helper: Get-LastSignIn ──
# The servicePrincipalSignInActivities report exposes 5 timestamps per app:
# delegated (user) / app-only (daemon), each as client or resource, plus an overall
# lastSignInActivity. We take the most recent across ALL of them (so daemon/app-only
# and user-delegated sign-ins are both covered) and report which flow it was.
function Get-LastSignIn {
  param($activity)
  if (-not $activity) { return @{ DateTime = $null; Flow = $null } }
  $cands = @(
    @{ F = 'Delegated (user)';     V = $activity.delegatedClientSignInActivity },
    @{ F = 'Delegated (resource)'; V = $activity.delegatedResourceSignInActivity },
    @{ F = 'App-only (daemon)';    V = $activity.applicationAuthenticationClientSignInActivity },
    @{ F = 'App-only (resource)';  V = $activity.applicationAuthenticationResourceSignInActivity },
    @{ F = 'Unknown';              V = $activity.lastSignInActivity }
  )
  $bestRaw = $null; $bestDt = $null; $bestFlow = $null
  foreach ($c in $cands) {
    $raw = if ($c.V -and $c.V.lastSignInDateTime) { $c.V.lastSignInDateTime } else { $null }
    if (-not $raw) { continue }
    try { $dt = [datetime]$raw } catch { continue }
    if (-not $bestDt -or $dt -gt $bestDt) { $bestDt = $dt; $bestRaw = $raw; $bestFlow = $c.F }
  }
  return @{ DateTime = $bestRaw; Flow = $bestFlow }
}

# ── Helper: Confirm-WriteAccess ──
# Called only when the user picks an action. Elevates the Graph session to
# Application.ReadWrite.All on demand. Returns $false (no error) if the user
# declines or consent isn't granted, so the run cleanly falls back to audit.
function Confirm-WriteAccess {
  $needed = "Application.ReadWrite.All"
  $ctx = Get-MgContext -ErrorAction SilentlyContinue
  if ($ctx -and ($needed -in @($ctx.Scopes))) { return $true }

  Write-Host ""
  Write-Host "  This action needs write permission ($needed)." -ForegroundColor Yellow
  Write-Host "  You'll be asked to sign in / consent once. Decline to keep audit-only." -ForegroundColor Gray
  $ok = Read-Host "  Grant write access and continue? (y/N)"
  if ($ok -notmatch '^(y|yes)$') {
    Write-Host "  Keeping read-only — no changes will be made." -ForegroundColor Yellow
    return $false
  }
  try {
    Connect-MgGraph -Scopes @("Application.Read.All","Directory.Read.All","AuditLog.Read.All","Application.ReadWrite.All") -NoWelcome
    $ctx = Get-MgContext
    if ($ctx -and ($needed -in @($ctx.Scopes))) { return $true }
    Write-Host "  [!] Write permission was not granted. Staying audit-only." -ForegroundColor Yellow
    return $false
  } catch {
    Write-Host "  [!] Could not get write permission: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "      Staying audit-only — no changes made." -ForegroundColor Yellow
    return $false
  }
}

# ============================================================
#  Permission risk dictionary
#  Classifies a Graph/API permission by what it can actually
#  do — NOT by how many permissions an app has.
#  Overall app risk = its single highest-risk permission.
# ============================================================

# Explicit HIGH — tenant takeover, write to directory/roles, or read/send
# all mail/files. (Many also match the *.ReadWrite.All heuristic below; listed
# here for clarity and to catch the ones that don't.)
$script:HighPerms = @{}
@(
  'RoleManagement.ReadWrite.Directory','Directory.ReadWrite.All','Application.ReadWrite.All',
  'Application.ReadWrite.OwnedBy','AppRoleAssignment.ReadWrite.All','User.ReadWrite.All',
  'Group.ReadWrite.All','GroupMember.ReadWrite.All','Device.ReadWrite.All',
  'Mail.ReadWrite','Mail.Send','MailboxSettings.ReadWrite','Files.ReadWrite.All',
  'Sites.ReadWrite.All','Sites.FullControl.All','Sites.Manage.All','full_access_as_app',
  'Exchange.ManageAsApp','Domain.ReadWrite.All','Organization.ReadWrite.All',
  'Policy.ReadWrite.ConditionalAccess','Policy.ReadWrite.AuthenticationMethod',
  'UserAuthenticationMethod.ReadWrite.All','PrivilegedAccess.ReadWrite.AzureAD',
  'PrivilegedAccess.ReadWrite.AzureADGroup','PrivilegedAccess.ReadWrite.AzureResources',
  'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup',
  'DeviceManagementConfiguration.ReadWrite.All','DeviceManagementManagedDevices.ReadWrite.All',
  'DeviceManagementApps.ReadWrite.All','DeviceManagementRBAC.ReadWrite.All',
  'IdentityRiskyUser.ReadWrite.All','RoleManagement.ReadWrite.Exchange'
) | ForEach-Object { $script:HighPerms[$_.ToLower()] = $true }

# Explicit MEDIUM — broad tenant-wide read of directory, mail, files, audit.
$script:MedPerms = @{}
@(
  'Directory.Read.All','User.Read.All','Group.Read.All','GroupMember.Read.All',
  'Application.Read.All','Device.Read.All','Mail.Read','Mail.ReadBasic','Mail.ReadBasic.All',
  'Files.Read.All','Sites.Read.All','AuditLog.Read.All','Reports.Read.All',
  'Policy.Read.All','People.Read.All','Calendars.ReadWrite','Calendars.Read',
  'Contacts.ReadWrite','Contacts.Read','Chat.Read.All','ChannelMessage.Read.All',
  'TeamMember.Read.All','Organization.Read.All','RoleManagement.Read.Directory',
  'IdentityRiskyUser.Read.All','UserAuthenticationMethod.Read.All'
) | ForEach-Object { $script:MedPerms[$_.ToLower()] = $true }

# When granted as an APPLICATION (app-only) permission, these tenant-wide reads
# run with no user context and effectively expose the whole tenant — escalate to High.
$script:AppReadEscalate = @{}
@(
  'Directory.Read.All','User.Read.All','Group.Read.All','GroupMember.Read.All',
  'Mail.Read','Mail.ReadBasic','Mail.ReadBasic.All','Files.Read.All','Sites.Read.All',
  'Calendars.Read','Contacts.Read','People.Read.All','Chat.Read.All','ChannelMessage.Read.All'
) | ForEach-Object { $script:AppReadEscalate[$_.ToLower()] = $true }

# Heuristics for permissions not in the explicit lists.
$script:HighLike = @('*.ReadWrite.All','*FullControl*','*ReadWrite.Directory','*.ReadWrite.OwnedBy','*Manage.All')

function Get-PermissionRisk {
  param([string]$Name, [string]$Type)   # Type = 'Application' or 'Delegated'
  if (-not $Name) { return 'Low' }
  $n = $Name.ToLower()
  $rank = 1   # Low by default

  if     ($script:HighPerms.ContainsKey($n)) { $rank = 3 }
  elseif ($script:MedPerms.ContainsKey($n))  { $rank = 2 }
  else {
    foreach ($pat in $script:HighLike) { if ($n -like $pat.ToLower()) { $rank = 3; break } }
    if ($rank -lt 3 -and $n -like '*.read.all') { $rank = 2 }
  }

  # App-only escalation for tenant-wide reads
  if ($Type -eq 'Application' -and $script:AppReadEscalate.ContainsKey($n) -and $rank -lt 3) {
    $rank = 3
  }

  switch ($rank) { 3 { 'High' } 2 { 'Medium' } default { 'Low' } }
}

function Get-RiskRank { param([string]$Risk)
  switch ($Risk) { 'High' { 3 } 'Medium' { 2 } 'Low' { 1 } default { 0 } }
}

# Friendly short names for the most common resource APIs
$script:ResourceShortNames = @{
  '00000003-0000-0000-c000-000000000000' = 'Graph'
  '00000002-0000-0000-c000-000000000000' = 'AAD Graph'
  '00000002-0000-0ff1-ce00-000000000000' = 'Exchange'
  '00000003-0000-0ff1-ce00-000000000000' = 'SharePoint'
  'fc780465-2017-40d4-a0c5-307022471b92' = 'WindowsVPN'
}

# ── Interactive banner ──
$banner = @"

  ============================================================
     Entra ID App Registration Audit
     Permission risk · Credentials · Staleness · Owners
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

# ── Step 2: Credential expiry warning threshold ──
Write-Host ""
Write-Host "  Credential expiry warning threshold:" -ForegroundColor White
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
    if ($customDays -match '^\d+$' -and [int]$customDays -gt 0) { $WarningDays = [int]$customDays }
    else { Write-Host "  [!] Invalid number. Using 30 days." -ForegroundColor Red; $WarningDays = 30 }
  }
  default { $WarningDays = 30 }
}
Write-Host "  -> Expiry warning: $WarningDays days" -ForegroundColor Green

# ── Step 3: Stale threshold ──
Write-Host ""
Write-Host "  Stale app threshold (no sign-in within):" -ForegroundColor White
Write-Host "    [1]  90 days   (default, recommended)" -ForegroundColor Green
Write-Host "    [2]  30 days" -ForegroundColor Yellow
Write-Host "    [3]  60 days" -ForegroundColor Yellow
Write-Host "    [4]  180 days" -ForegroundColor Yellow
Write-Host "    [5]  Custom" -ForegroundColor Yellow
Write-Host ""
$staleChoice = Read-Host "  Enter choice (1-5) [default: 1]"
switch ($staleChoice) {
  "2" { $StaleDays = 30 }
  "3" { $StaleDays = 60 }
  "4" { $StaleDays = 180 }
  "5" {
    $customDays = Read-Host "  Enter number of days"
    if ($customDays -match '^\d+$' -and [int]$customDays -gt 0) { $StaleDays = [int]$customDays }
    else { Write-Host "  [!] Invalid number. Using 90 days." -ForegroundColor Red; $StaleDays = 90 }
  }
  default { $StaleDays = 90 }
}
Write-Host "  -> Stale threshold: $StaleDays days" -ForegroundColor Green

# ── (Scope is chosen after connecting — see the object-type prompt below, ──
#     which shows live counts per type so you can pick what to audit.) ──

# ── Summary & confirm ──
Write-Host ""
Write-Host "  --------------------------------------" -ForegroundColor Gray
Write-Host "  Expiry warning:   $WarningDays days" -ForegroundColor White
Write-Host "  Stale threshold:  $StaleDays days" -ForegroundColor White
Write-Host "  Scope:            chosen after connect (with live counts)" -ForegroundColor White
Write-Host "  Flow:             Audit -> Review -> Optionally act (gated)" -ForegroundColor Gray
Write-Host "  --------------------------------------" -ForegroundColor Gray
Write-Host ""

$confirm = Read-Host "  Press ENTER to start or type 'q' to quit"
if ($confirm -and $confirm.Trim().ToLower() -ne '') {
  Write-Host "  Cancelled." -ForegroundColor Red
  return
}

# ── Overall timer ──
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$now = Get-Date
$warningDate = $now.AddDays($WarningDays)
$cutoffDate  = $now.AddDays(-$StaleDays).ToUniversalTime()

# ── Phase 1: Connect ──
$totalPhases = 7
$phase = 1
Write-Host ""
Write-Host "  [$phase/$totalPhases] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Read-only by default — write permission is requested later, and only if the
# user actually chooses an action (remove / disable / delete). An audit run
# never asks the admin to consent to write access.
$requiredScopes = @(
  "Application.Read.All",
  "Directory.Read.All",
  "AuditLog.Read.All"
)

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$needsConnect = $true
if ($ctx -and $ctx.TenantId) {
  $missing = $requiredScopes | Where-Object { $_ -notin @($ctx.Scopes) }
  if ($missing.Count -eq 0) {
    Write-Host "  Reusing existing Graph session ($($ctx.Account))" -ForegroundColor Green
    $needsConnect = $false
  } else {
    Write-Host "  Existing session missing scopes: $($missing -join ', ')" -ForegroundColor Yellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
  }
}
if ($needsConnect) {
  try {
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $ctx = Get-MgContext
  } catch {
    Write-Host "  [!] Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    return
  }
}
if (-not $ctx) { Write-Host "  [!] Graph context not established." -ForegroundColor Red; return }
$tenantId = $ctx.TenantId
Write-Host "  Connected to tenant: $tenantId  ($($ctx.Account))" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 2: Fetch apps, service principals, sign-in activity ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Fetching app registrations and sign-in activity..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host "    - App registrations..." -ForegroundColor Gray
$appRegs = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/applications?`$top=999&`$select=id,appId,displayName,createdDateTime,publisherDomain,signInAudience,keyCredentials,passwordCredentials,requiredResourceAccess"
Write-Host "    -> $($appRegs.Count) app registrations" -ForegroundColor Green

Write-Host "    - Service principals..." -ForegroundColor Gray
$sps = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/servicePrincipals?`$top=999&`$select=id,appId,displayName,accountEnabled,servicePrincipalType,appOwnerOrganizationId,keyCredentials,passwordCredentials"
Write-Host "    -> $($sps.Count) service principals" -ForegroundColor Green
$spByAppId = @{}
foreach ($sp in $sps) { if ($sp.appId) { $spByAppId[$sp.appId] = $sp } }

$msFirstPartyTenantId = "f8cdef31-a31e-4b4a-93e4-5f571e91255a"

Write-Host "    - Service principal sign-in activity (beta)..." -ForegroundColor Gray
try {
  $signInActivity = Invoke-GraphPaged "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?`$top=999"
  Write-Host "    -> $($signInActivity.Count) sign-in activity records" -ForegroundColor Green
} catch {
  Write-Host "  [!] Could not fetch sign-in activity. Tenant may lack Entra ID P1/P2." -ForegroundColor Red
  Write-Host "      Staleness will show as 'Unknown' for all apps." -ForegroundColor Yellow
  $signInActivity = @()
}
$activityByAppId = @{}
foreach ($a in $signInActivity) { if ($a.appId) { $activityByAppId[$a.appId] = $a } }

# Delegated permission grants (consented OAuth2 scopes) — fetched once, tenant-wide,
# then grouped by the client SP so each enterprise app knows what it was granted.
Write-Host "    - Enterprise app delegated grants (OAuth2)..." -ForegroundColor Gray
try {
  $oauthGrants = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$top=999"
  Write-Host "    -> $($oauthGrants.Count) delegated grant record(s)" -ForegroundColor Green
} catch {
  Write-Host "  [!] Could not fetch OAuth2 grants: $($_.Exception.Message)" -ForegroundColor Yellow
  $oauthGrants = @()
}
$grantsByClient = @{}
foreach ($g in $oauthGrants) {
  if (-not $g.clientId) { continue }
  if (-not $grantsByClient.ContainsKey($g.clientId)) { $grantsByClient[$g.clientId] = @() }
  $grantsByClient[$g.clientId] += $g
}
# SP object lookup by objectId (resource resolution) and the set of appIds that have a local app registration.
$spById = @{}
foreach ($sp in $sps) { if ($sp.id) { $spById[$sp.id] = $sp } }
$appRegAppIds = @{}
foreach ($app in $appRegs) { if ($app.appId) { $appRegAppIds[$app.appId] = $true } }
$script:swPhase.Stop()

# ── Scope: pick which object types to audit (live counts from this tenant) ──
# Enterprise apps (service principals) are split by who owns them so you can, e.g.,
# audit only the external apps users consented to and skip the Microsoft noise.
function Get-SpCategory($sp) {
  if ("$($sp.servicePrincipalType)" -eq 'ManagedIdentity') { return 'ManagedId' }
  if ($sp.appOwnerOrganizationId -eq $msFirstPartyTenantId)  { return 'MS' }
  if ($sp.appOwnerOrganizationId -eq $tenantId)              { return 'Org' }
  return 'ThirdParty'
}
$catCount = @{ MS = 0; Org = 0; ThirdParty = 0; ManagedId = 0 }
foreach ($sp in $sps) { $catCount[(Get-SpCategory $sp)]++ }

Write-Host ""
Write-Host "  What should this scan audit?  (live counts from your tenant)" -ForegroundColor White
Write-Host "    [1]  App registrations             $("{0,5}" -f $appRegs.Count)   Entra > App registrations (apps built in your tenant)" -ForegroundColor Green
Write-Host "    [2]  Enterprise apps - your org     $("{0,5}" -f $catCount.Org)   Entra > Enterprise applications (single-tenant apps your org created)" -ForegroundColor Green
Write-Host "    [3]  Enterprise apps - third-party  $("{0,5}" -f $catCount.ThirdParty)   external apps users consented to (top consent-phishing risk)" -ForegroundColor Green
Write-Host "    [4]  Enterprise apps - Microsoft    $("{0,5}" -f $catCount.MS)   Microsoft first-party apps (usually noise; slower)" -ForegroundColor Yellow
Write-Host "    [5]  Managed identities            $("{0,5}" -f $catCount.ManagedId)   Azure managed identities" -ForegroundColor Yellow
Write-Host "    [6]  All of the above            $("{0,5}" -f ($appRegs.Count + $catCount.Org + $catCount.ThirdParty + $catCount.MS + $catCount.ManagedId))   everything (slower)" -ForegroundColor Yellow
$scopeInput = Read-Host "  Enter choices, comma-separated (e.g. 1,3), 6 or 'all' for everything [default: 1,2,3]"
if (-not $scopeInput -or $scopeInput.Trim() -eq '') { $scopeInput = '1,2,3' }
if ($scopeInput -match '(?i)\ball\b' -or (@($scopeInput -split '[,\s]+') -contains '6')) {
  $scopeSel = @(1,2,3,4,5)
} else {
  $scopeSel = @($scopeInput -split '[,\s]+' | Where-Object { $_ -match '^[1-5]$' } | ForEach-Object { [int]$_ } | Sort-Object -Unique)
  if ($scopeSel.Count -eq 0) { Write-Host "  [!] No valid choice — using default 1,2,3." -ForegroundColor Yellow; $scopeSel = @(1,2,3) }
}
$doAppRegs       = $scopeSel -contains 1
$doEntOrg        = $scopeSel -contains 2
$doEntThirdParty = $scopeSel -contains 3
$doEntMS         = $scopeSel -contains 4
$doManagedId     = $scopeSel -contains 5
$doEnt           = $doEntOrg -or $doEntThirdParty -or $doEntMS -or $doManagedId

$scopeLabels = @()
if ($doAppRegs)       { $scopeLabels += "App regs ($($appRegs.Count))" }
if ($doEntOrg)        { $scopeLabels += "Ent:org ($($catCount.Org))" }
if ($doEntThirdParty) { $scopeLabels += "Ent:third-party ($($catCount.ThirdParty))" }
if ($doEntMS)         { $scopeLabels += "Ent:Microsoft ($($catCount.MS))" }
if ($doManagedId)     { $scopeLabels += "Managed IDs ($($catCount.ManagedId))" }
$scopeSummary = if ($scopeLabels.Count) { $scopeLabels -join ', ' } else { 'None' }
Write-Host "  -> Scope: $scopeSummary" -ForegroundColor Green
if ($doEntMS -and $catCount.MS -gt 200) {
  Write-Host "  [i] Auditing $($catCount.MS) Microsoft apps means a Graph call each — this can take several minutes." -ForegroundColor Yellow
}

# ── Phase 3: Resolve permission names from resource service principals ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Resolving API permission names..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Collect every distinct resource API referenced by any app
$resourceAppIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($app in $appRegs) {
  foreach ($r in @($app.requiredResourceAccess)) {
    if ($r -and $r.resourceAppId) { [void]$resourceAppIds.Add($r.resourceAppId) }
  }
}

# resourceAppId -> @{ Name; Roles(id->value); Scopes(id->value) }
$resourceMap = @{}
$ri = 0
foreach ($resId in $resourceAppIds) {
  $ri++
  Write-ProgressBar -Current $ri -Total $resourceAppIds.Count -Activity "Resolving resource APIs" -Status $resId
  $roles  = @{}
  $scopes = @{}
  $rName  = if ($script:ResourceShortNames.ContainsKey($resId)) { $script:ResourceShortNames[$resId] } else { $resId }
  try {
    $rsp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$resId')?`$select=displayName,appId,appRoles,oauth2PermissionScopes" -ErrorAction Stop
    if ($rsp.displayName) { $rName = $rsp.displayName }
    foreach ($ar in @($rsp.appRoles))             { if ($ar.id) { $roles[$ar.id.ToString()]  = $ar.value } }
    foreach ($sc in @($rsp.oauth2PermissionScopes)) { if ($sc.id) { $scopes[$sc.id.ToString()] = $sc.value } }
  } catch {
    # resource SP not present in tenant — keep GUIDs as fallback names
  }
  $resourceMap[$resId] = @{ Name = $rName; Roles = $roles; Scopes = $scopes }
}
Write-Progress -Activity "Resolving resource APIs" -Completed
Write-Host "  Resolved $($resourceAppIds.Count) resource API(s)" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 4: Analyze every app ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Analyzing apps (permissions, credentials, staleness)..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

function Resolve-Permissions {
  param($app)
  $perms = @()
  foreach ($r in @($app.requiredResourceAccess)) {
    if (-not $r) { continue }
    $resId = $r.resourceAppId
    $rmap  = if ($resId -and $resourceMap.ContainsKey($resId)) { $resourceMap[$resId] } else { $null }
    $resName = if ($rmap) { $rmap.Name } else { $resId }
    foreach ($acc in @($r.resourceAccess)) {
      if (-not $acc) { continue }
      $isRole = ($acc.type -eq 'Role')
      $type   = if ($isRole) { 'Application' } else { 'Delegated' }
      $name   = $null
      if ($rmap) {
        $key = if ($acc.id) { $acc.id.ToString() } else { $null }
        if ($isRole -and $key -and $rmap.Roles.ContainsKey($key))  { $name = $rmap.Roles[$key] }
        elseif (-not $isRole -and $key -and $rmap.Scopes.ContainsKey($key)) { $name = $rmap.Scopes[$key] }
      }
      if (-not $name) { $name = "($($acc.id))" }
      $risk = Get-PermissionRisk -Name $name -Type $type
      $perms += [PSCustomObject]@{
        Resource = $resName
        Name     = $name
        Type     = $type
        Risk     = $risk
      }
    }
  }
  return ,$perms
}

function Get-CredentialBuckets {
  param($app, $portalUrl, $source = 'App reg')
  $out = @()
  foreach ($pair in @(
      @{ Set = $app.keyCredentials;      Kind = 'Certificate'   },
      @{ Set = $app.passwordCredentials; Kind = 'Client Secret' })) {
    foreach ($c in @($pair.Set)) {
      if (-not $c) { continue }
      $endDate = $null; $startDate = $null
      try {
        $endDate   = [datetime]::Parse($c.endDateTime,   [System.Globalization.CultureInfo]::InvariantCulture)
        $startDate = [datetime]::Parse($c.startDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
      } catch { continue }
      $days = [math]::Round(($endDate - $now).TotalDays)
      $status = if ($days -lt 0) { "Expired" } elseif ($days -le $WarningDays) { "Expiring Soon" } else { "Valid" }
      $out += [PSCustomObject]@{
        AppName = $app.displayName; AppId = $app.appId; ObjectId = $app.id
        CredentialType = $pair.Kind
        Description = if ($c.displayName) { $c.displayName } else { "N/A" }
        KeyId = $c.keyId
        StartDate = $startDate.ToString('yyyy-MM-dd'); EndDate = $endDate.ToString('yyyy-MM-dd')
        DaysToExpiry = $days; Status = $status; PortalUrl = $portalUrl; Owner = "N/A"; Source = $source
      }
    }
  }
  return ,$out
}

$analysis = @()
$allCreds = @()
$i = 0; $total = $appRegs.Count
foreach ($app in $appRegs) {
  $i++
  if (($i % 25) -eq 0 -or $i -eq $total) {
    Write-ProgressBar -Current $i -Total $total -Activity "Analyzing apps" -Status $app.displayName
  }

  if (-not $doAppRegs) { continue }
  $sp = if ($app.appId -and $spByAppId.ContainsKey($app.appId)) { $spByAppId[$app.appId] } else { $null }
  $isFirstParty = $false
  if ($sp -and $sp.appOwnerOrganizationId -eq $msFirstPartyTenantId) { $isFirstParty = $true }
  if ($app.publisherDomain -match '(microsoft\.com|microsoft\.onmicrosoft\.com)$') { $isFirstParty = $true }

  $portalUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($app.appId)"

  # Permissions + risk
  $perms = Resolve-Permissions $app
  $highP = @($perms | Where-Object { $_.Risk -eq 'High' })
  $medP  = @($perms | Where-Object { $_.Risk -eq 'Medium' })
  $lowP  = @($perms | Where-Object { $_.Risk -eq 'Low' })
  $overallRisk = if ($highP.Count) { 'High' } elseif ($medP.Count) { 'Medium' } elseif ($lowP.Count) { 'Low' } else { 'None' }
  $sensitiveList = (@($highP + $medP) | ForEach-Object { "$($_.Resource): $($_.Name) ($($_.Type))" }) -join "`n"
  # Distinct API products this app touches (all perms) + a flat string of every
  # permission name, so the HTML report can filter by product and search details.
  $permProducts  = @($perms | ForEach-Object { $_.Resource } | Where-Object { $_ } | Sort-Object -Unique)
  $permNames     = (@($perms | ForEach-Object { $_.Name }) -join ' ')

  # Credentials
  $creds = Get-CredentialBuckets $app $portalUrl
  $allCreds += $creds
  $expiredN  = @($creds | Where-Object { $_.Status -eq 'Expired' }).Count
  $expiringN = @($creds | Where-Object { $_.Status -eq 'Expiring Soon' }).Count

  # Staleness — most recent sign-in across all flows (delegated + app-only)
  $activity   = if ($app.appId -and $activityByAppId.ContainsKey($app.appId)) { $activityByAppId[$app.appId] } else { $null }
  $si = Get-LastSignIn $activity
  $lastSignIn = $si.DateTime; $signInFlow = $si.Flow
  if (-not $sp) {
    $staleStatus = "No service principal"; $daysSince = $null
  } elseif ($signInActivity.Count -eq 0) {
    $staleStatus = "Unknown"; $daysSince = $null
  } elseif (-not $lastSignIn) {
    $staleStatus = "Never used"; $daysSince = $null
  } else {
    $lastDt = [datetime]$lastSignIn
    $daysSince = [int]((Get-Date) - $lastDt).TotalDays
    $staleStatus = if ($lastDt -lt $cutoffDate) { "Stale" } else { "Active" }
  }

  $analysis += [PSCustomObject]@{
    DisplayName    = $app.displayName
    AppId          = $app.appId
    ObjectId       = $app.id
    SPObjectId     = if ($sp) { $sp.id } else { $null }
    SPEnabled      = if ($sp) { [bool]$sp.accountEnabled } else { $null }
    FirstParty     = $isFirstParty
    Tenancy        = if ($isFirstParty) { 'Microsoft' } elseif ($sp -and $sp.appOwnerOrganizationId -eq $tenantId) { 'Tenant' } else { 'Tenant/External' }
    WorkloadIdCand = if (-not $sp) { 'No SP — N/A' }
                     elseif ($isFirstParty) { 'No (Microsoft app)' }
                     elseif ($sp.servicePrincipalType -eq 'ManagedIdentity') { 'No (managed identity)' }
                     elseif ($sp.appOwnerOrganizationId -eq $tenantId) { 'CA + ID Protection' }
                     else { 'ID Protection only' }
    OverallRisk    = $overallRisk
    HighCount      = $highP.Count
    MedCount       = $medP.Count
    LowCount       = $lowP.Count
    TotalPerms     = $perms.Count
    SensitivePerms = $sensitiveList
    Products       = $permProducts
    PermSearch     = $permNames
    StaleStatus    = $staleStatus
    DaysSinceLastSignIn = $daysSince
    LastSignIn     = $lastSignIn
    SignInFlow     = $signInFlow
    TotalCreds     = $creds.Count
    ExpiredCreds   = $expiredN
    ExpiringCreds  = $expiringN
    CreatedDateTime = $app.createdDateTime
    Owner          = "N/A"
    PortalLink     = $portalUrl
  }
}
Write-Progress -Activity "Analyzing apps" -Completed

$expiredCreds  = @($allCreds | Where-Object { $_.Status -eq 'Expired' } | Sort-Object DaysToExpiry)
$expiringCreds = @($allCreds | Where-Object { $_.Status -eq 'Expiring Soon' } | Sort-Object DaysToExpiry)

$highRiskApps = @($analysis | Where-Object { $_.OverallRisk -eq 'High' })
$medRiskApps  = @($analysis | Where-Object { $_.OverallRisk -eq 'Medium' })
$staleApps    = @($analysis | Where-Object { $_.StaleStatus -eq 'Stale' })
$neverUsed    = @($analysis | Where-Object { $_.StaleStatus -eq 'Never used' })
$orphaned     = @($analysis | Where-Object { $_.StaleStatus -eq 'No service principal' })
$caCandidates = @($analysis | Where-Object { $_.WorkloadIdCand -eq 'CA + ID Protection' })
$idpOnly      = @($analysis | Where-Object { $_.WorkloadIdCand -eq 'ID Protection only' })

Write-Host ""
Write-Host "  Results:" -ForegroundColor White
Write-Host "    Apps analyzed:        $($analysis.Count)" -ForegroundColor White
Write-Host "    HIGH-risk perms:      $($highRiskApps.Count)" -ForegroundColor $(if ($highRiskApps.Count) { 'Red' } else { 'Green' })
Write-Host "    MEDIUM-risk perms:    $($medRiskApps.Count)" -ForegroundColor $(if ($medRiskApps.Count) { 'Yellow' } else { 'Green' })
Write-Host "    Expired credentials:  $($expiredCreds.Count)" -ForegroundColor $(if ($expiredCreds.Count) { 'Red' } else { 'Green' })
Write-Host "    Expiring soon:        $($expiringCreds.Count)" -ForegroundColor $(if ($expiringCreds.Count) { 'Yellow' } else { 'Green' })
Write-Host "    Stale apps:           $($staleApps.Count)" -ForegroundColor $(if ($staleApps.Count) { 'Yellow' } else { 'Green' })
Write-Host "    Never used:           $($neverUsed.Count)" -ForegroundColor Gray
Write-Host "    Workload ID (CA):     $($caCandidates.Count)  [+ $($idpOnly.Count) ID Protection only]" -ForegroundColor Gray
$script:swPhase.Stop()

# ── Phase 5: Analyze enterprise apps (actual GRANTED permissions) ──
# App registrations above describe what an app *requests*. Enterprise apps
# (service principals) are the identities that actually hold consented access —
# including third-party / SaaS / gallery apps that have no local app registration.
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Analyzing enterprise apps (granted permissions, credentials, staleness)..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

# Resolve a resource SP's app-role ids -> permission values (cached by SP objectId).
$resRolesByObjId = @{}
function Get-ResourceRoleInfo {
  param([string]$ObjId)
  if (-not $ObjId) { return @{ Name = $null; Roles = @{} } }
  if ($resRolesByObjId.ContainsKey($ObjId)) { return $resRolesByObjId[$ObjId] }
  $name = if ($spById.ContainsKey($ObjId)) { $spById[$ObjId].displayName } else { $ObjId }
  $roles = @{}
  try {
    $rsp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ObjId`?`$select=displayName,appRoles" -ErrorAction Stop
    if ($rsp.displayName) { $name = $rsp.displayName }
    foreach ($ar in @($rsp.appRoles)) { if ($ar.id) { $roles[$ar.id.ToString()] = $ar.value } }
  } catch { }
  $info = @{ Name = $name; Roles = $roles }
  $resRolesByObjId[$ObjId] = $info
  return $info
}

# Pass 1 — which SPs are in scope (by the categories chosen at the scope prompt).
$entTargets = @()
foreach ($sp in $sps) {
  $cat = Get-SpCategory $sp
  $include = switch ($cat) { 'Org' { $doEntOrg } 'ThirdParty' { $doEntThirdParty } 'MS' { $doEntMS } 'ManagedId' { $doManagedId } default { $false } }
  if ($include) { $entTargets += [PSCustomObject]@{ Sp = $sp; Cat = $cat } }
}
Write-Host "    - $($entTargets.Count) enterprise app(s) in scope; fetching grants in parallel batches..." -ForegroundColor Gray

# Batch-fetch each in-scope SP's app-role assignments (application permissions), 20 per round trip.
$asgReqs = @($entTargets | ForEach-Object { @{ id = "$($_.Sp.id)"; url = "/servicePrincipals/$($_.Sp.id)/appRoleAssignments?`$select=appRoleId,resourceId,resourceDisplayName&`$top=999" } })
$asgBody = Invoke-GraphBatch -Requests $asgReqs -Activity "Fetching app-role assignments"

# Pre-warm the resource app-role dictionaries for every distinct resource referenced (one batch).
$resIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($k in $asgBody.Keys) { foreach ($a in @($asgBody[$k].value)) { if ($a.resourceId) { [void]$resIds.Add("$($a.resourceId)") } } }
$resReqs = @($resIds | ForEach-Object { @{ id = "$_"; url = "/servicePrincipals/$_`?`$select=id,displayName,appRoles" } })
$resBody = Invoke-GraphBatch -Requests $resReqs -Activity "Resolving app-role names"
foreach ($rid in $resBody.Keys) {
  $b = $resBody[$rid]; $roles = @{}
  $name = if ($spById.ContainsKey($rid)) { $spById[$rid].displayName } else { $rid }
  if ($b) { if ($b.displayName) { $name = $b.displayName }; foreach ($ar in @($b.appRoles)) { if ($ar.id) { $roles[$ar.id.ToString()] = $ar.value } } }
  $resRolesByObjId[$rid] = @{ Name = $name; Roles = $roles }
}

$entAnalysis = @()
$entCreds = @()
$ei = 0; $etotal = $entTargets.Count
foreach ($t in $entTargets) {
  $ei++
  $sp = $t.Sp; $cat = $t.Cat
  if (($ei % 25) -eq 0 -or $ei -eq $etotal) { Write-ProgressBar -Current $ei -Total $etotal -Activity "Analyzing enterprise apps" -Status $sp.displayName }
  $isFirstParty = ($cat -eq 'MS')

  $portalUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_Managed/ManagedAppMenuBlade/~/Permissions/objectId/$($sp.id)/appId/$($sp.appId)"

  # --- Granted permissions ---
  $perms = @()
  # Delegated (consented OAuth2 scopes) — scope values are already human-readable.
  foreach ($g in @($grantsByClient[$sp.id])) {
    if (-not $g) { continue }
    $resName = if ($g.resourceId -and $spById.ContainsKey($g.resourceId)) { $spById[$g.resourceId].displayName } else { $g.resourceId }
    foreach ($scope in @(("$($g.scope)").Trim() -split '\s+')) {
      if (-not $scope) { continue }
      $perms += [PSCustomObject]@{ Resource = $resName; Name = $scope; Type = 'Delegated'; Risk = (Get-PermissionRisk -Name $scope -Type 'Delegated') }
    }
  }
  # Application (app roles the SP has been granted) — from the pre-fetched batch.
  $assignments = @($asgBody["$($sp.id)"].value)
  foreach ($asg in @($assignments)) {
    if (-not $asg) { continue }
    $roleKey = "$($asg.appRoleId)"
    if (-not $roleKey -or $roleKey -eq '00000000-0000-0000-0000-000000000000') { continue }   # default access, no permission
    $rinfo = Get-ResourceRoleInfo -ObjId "$($asg.resourceId)"
    $resName = if ($asg.resourceDisplayName) { $asg.resourceDisplayName } elseif ($rinfo.Name) { $rinfo.Name } else { $asg.resourceId }
    $pname = if ($rinfo.Roles.ContainsKey($roleKey)) { $rinfo.Roles[$roleKey] } else { "($roleKey)" }
    $perms += [PSCustomObject]@{ Resource = $resName; Name = $pname; Type = 'Application'; Risk = (Get-PermissionRisk -Name $pname -Type 'Application') }
  }

  $highP = @($perms | Where-Object { $_.Risk -eq 'High' })
  $medP  = @($perms | Where-Object { $_.Risk -eq 'Medium' })
  $lowP  = @($perms | Where-Object { $_.Risk -eq 'Low' })
  $overallRisk = if ($highP.Count) { 'High' } elseif ($medP.Count) { 'Medium' } elseif ($lowP.Count) { 'Low' } else { 'None' }
  $sensitiveList = (@($highP + $medP) | ForEach-Object { "$($_.Resource): $($_.Name) ($($_.Type))" }) -join "`n"
  $permProducts  = @($perms | ForEach-Object { $_.Resource } | Where-Object { $_ } | Sort-Object -Unique)
  $permNames     = (@($perms | ForEach-Object { $_.Name }) -join ' ')

  # --- Credentials on the SP (SAML signing certs, secrets) ---
  $creds = Get-CredentialBuckets $sp $portalUrl 'Enterprise'
  $entCreds += $creds
  $expiredN  = @($creds | Where-Object { $_.Status -eq 'Expired' }).Count
  $expiringN = @($creds | Where-Object { $_.Status -eq 'Expiring Soon' }).Count

  # --- Staleness --- most recent sign-in across all flows (delegated + app-only)
  $activity   = if ($sp.appId -and $activityByAppId.ContainsKey($sp.appId)) { $activityByAppId[$sp.appId] } else { $null }
  $si = Get-LastSignIn $activity
  $lastSignIn = $si.DateTime; $signInFlow = $si.Flow
  if ($cat -eq 'ManagedId' -and -not $lastSignIn) {
    # Managed identity token use largely isn't captured by this report — don't call it "Never used".
    $staleStatus = "N/A (not tracked)"; $daysSince = $null
  } elseif ($signInActivity.Count -eq 0) { $staleStatus = "Unknown"; $daysSince = $null }
  elseif (-not $lastSignIn)        { $staleStatus = "Never used"; $daysSince = $null }
  else {
    $lastDt = [datetime]$lastSignIn
    $daysSince = [int]((Get-Date) - $lastDt).TotalDays
    $staleStatus = if ($lastDt -lt $cutoffDate) { "Stale" } else { "Active" }
  }

  $entAnalysis += [PSCustomObject]@{
    DisplayName    = $sp.displayName
    AppId          = $sp.appId
    SPObjectId     = $sp.id
    SPEnabled      = [bool]$sp.accountEnabled
    SPType         = "$($sp.servicePrincipalType)"
    FirstParty     = $isFirstParty
    Category       = switch ($cat) { 'Org' { 'Your org' } 'ThirdParty' { 'Third-party' } 'MS' { 'Microsoft' } 'ManagedId' { 'Managed identity' } default { $cat } }
    HasAppReg      = if ($sp.appId -and $appRegAppIds.ContainsKey($sp.appId)) { $true } else { $false }
    Tenancy        = if ($isFirstParty) { 'Microsoft' } elseif ($sp.appOwnerOrganizationId -eq $tenantId) { 'Tenant' } else { 'Tenant/External' }
    OverallRisk    = $overallRisk
    HighCount      = $highP.Count
    MedCount       = $medP.Count
    LowCount       = $lowP.Count
    TotalPerms     = $perms.Count
    SensitivePerms = $sensitiveList
    Products       = $permProducts
    PermSearch     = $permNames
    StaleStatus    = $staleStatus
    DaysSinceLastSignIn = $daysSince
    LastSignIn     = $lastSignIn
    SignInFlow     = $signInFlow
    TotalCreds     = $creds.Count
    ExpiredCreds   = $expiredN
    ExpiringCreds  = $expiringN
    Owner          = "N/A"
    PortalLink     = $portalUrl
  }
}
Write-Progress -Activity "Analyzing enterprise apps" -Completed

# Owners for the enterprise apps that surface in the report — batched.
$entReport = @($entAnalysis | Where-Object {
  $_.OverallRisk -in @('High','Medium') -or $_.ExpiredCreds -gt 0 -or $_.ExpiringCreds -gt 0 -or $_.StaleStatus -in @('Stale','Never used')
})
$entOwnReqs = @($entReport | ForEach-Object { @{ id = "$($_.SPObjectId)"; url = "/servicePrincipals/$($_.SPObjectId)/owners?`$select=displayName,userPrincipalName" } })
$entOwnBody = Invoke-GraphBatch -Requests $entOwnReqs -Activity "Fetching enterprise app owners"
foreach ($e in $entReport) {
  $ob = $entOwnBody["$($e.SPObjectId)"]
  if ($ob -and $ob.value -and @($ob.value).Count -gt 0) {
    $e.Owner = (@($ob.value) | ForEach-Object { if ($_.displayName) { $_.displayName } elseif ($_.userPrincipalName) { $_.userPrincipalName } else { "Unknown" } }) -join ", "
  } elseif ($null -eq $ob) { $e.Owner = "N/A" } else { $e.Owner = "No owner" }
}
$entOwnerByObjId = @{}
foreach ($e in $entAnalysis) { $entOwnerByObjId[$e.SPObjectId] = $e.Owner }
foreach ($c in $entCreds) { if ($c.ObjectId -and $entOwnerByObjId.ContainsKey($c.ObjectId)) { $c.Owner = $entOwnerByObjId[$c.ObjectId] } }

$entHigh     = @($entAnalysis | Where-Object { $_.OverallRisk -eq 'High' })
$entMed      = @($entAnalysis | Where-Object { $_.OverallRisk -eq 'Medium' })
$entStale    = @($entAnalysis | Where-Object { $_.StaleStatus -eq 'Stale' })
$entNoAppReg = @($entAnalysis | Where-Object { -not $_.HasAppReg })
$entExpired  = @($entCreds | Where-Object { $_.Status -eq 'Expired' } | Sort-Object DaysToExpiry)
$entExpiring = @($entCreds | Where-Object { $_.Status -eq 'Expiring Soon' } | Sort-Object DaysToExpiry)

Write-Host ""
Write-Host "  Enterprise apps:" -ForegroundColor White
Write-Host "    Analyzed:             $($entAnalysis.Count)  (no app registration: $($entNoAppReg.Count))" -ForegroundColor White
Write-Host "    HIGH-risk grants:     $($entHigh.Count)" -ForegroundColor $(if ($entHigh.Count) { 'Red' } else { 'Green' })
Write-Host "    MEDIUM-risk grants:   $($entMed.Count)" -ForegroundColor $(if ($entMed.Count) { 'Yellow' } else { 'Green' })
Write-Host "    Stale:                $($entStale.Count)" -ForegroundColor Gray
$script:swPhase.Stop()

# ── Phase 6: Resolve owners for apps that appear in the report ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Resolving owners..." -ForegroundColor Cyan
$script:swPhase = [System.Diagnostics.Stopwatch]::StartNew()

$reportApps = @($analysis | Where-Object {
  $_.OverallRisk -in @('High','Medium') -or $_.ExpiredCreds -gt 0 -or $_.ExpiringCreds -gt 0 -or
  $_.StaleStatus -in @('Stale','Never used','No service principal')
})
$ownerCache = @{}
$ownReqs = @($reportApps | Where-Object { $_.ObjectId } | ForEach-Object { @{ id = "$($_.ObjectId)"; url = "/applications/$($_.ObjectId)/owners?`$select=displayName,userPrincipalName" } })
$ownBody = Invoke-GraphBatch -Requests $ownReqs -Activity "Fetching owners"
foreach ($a in $reportApps) {
  if ($ownerCache.ContainsKey($a.ObjectId)) { continue }
  $ob = $ownBody["$($a.ObjectId)"]
  if ($ob -and $ob.value -and @($ob.value).Count -gt 0) {
    $ownerCache[$a.ObjectId] = (@($ob.value) | ForEach-Object {
      if ($_.displayName) { $_.displayName } elseif ($_.userPrincipalName) { $_.userPrincipalName } else { "Unknown" }
    }) -join ", "
  } elseif ($null -eq $ob) { $ownerCache[$a.ObjectId] = "N/A" } else { $ownerCache[$a.ObjectId] = "No owner" }
}
foreach ($a in $analysis)  { if ($ownerCache.ContainsKey($a.ObjectId)) { $a.Owner = $ownerCache[$a.ObjectId] } }
foreach ($c in $allCreds)  { if ($ownerCache.ContainsKey($c.ObjectId)) { $c.Owner = $ownerCache[$c.ObjectId] } }
Write-Host "  Owners resolved for $($reportApps.Count) app(s)" -ForegroundColor Green
$script:swPhase.Stop()

# ── Phase 7: Review & act (gated) ──
$phase++
Write-Host ""
Write-Host "  [$phase/$totalPhases] Review & decide" -ForegroundColor Cyan
Write-Host ""
Write-Host "    [1]  Export report only — no changes        (default, safe)" -ForegroundColor Green
Write-Host "    [2]  Remove EXPIRED credentials, then export" -ForegroundColor Red
Write-Host "    [3]  Disable service principals of stale/never-used apps (reversible)" -ForegroundColor Yellow
Write-Host "    [4]  DELETE stale/never-used app registrations (PERMANENT)" -ForegroundColor Red
Write-Host ""
$actionChoice = Read-Host "  Enter choice (1-4) [default: 1]"
if (-not $actionChoice) { $actionChoice = "1" }

$actionResults = @()
$actionMode = "Audit"
$cleanupTargets = @($staleApps + $neverUsed)

switch ($actionChoice) {
  "2" {
    $actionMode = "Remove expired credentials"
    if ($expiredCreds.Count -eq 0) {
      Write-Host "  No expired credentials to remove. Audit only." -ForegroundColor Green
      $actionMode = "Audit"
    } elseif (-not (Confirm-WriteAccess)) {
      $actionMode = "Audit (write permission declined)"
    } else {
      Write-Host ""
      Write-Host "  $($expiredCreds.Count) expired credential(s) will be removed:" -ForegroundColor White
      foreach ($c in @($expiredCreds | Select-Object -First 15)) {
        Write-Host "    - $($c.AppName) | $($c.CredentialType) | $($c.Description) | expired $([math]::Abs($c.DaysToExpiry))d ago" -ForegroundColor Gray
      }
      if ($expiredCreds.Count -gt 15) { Write-Host "    ... and $($expiredCreds.Count - 15) more" -ForegroundColor Gray }
      Write-Host ""
      $cf = Read-Host "  Type YES to confirm removal of $($expiredCreds.Count) expired credential(s)"
      if ($cf -ne 'YES') { Write-Host "  Cancelled — exporting audit only." -ForegroundColor Yellow; $actionMode = "Audit (cancelled removal)" }
      else {
        $j = 0
        foreach ($c in $expiredCreds) {
          $j++; Write-ProgressBar -Current $j -Total $expiredCreds.Count -Activity "Removing credentials" -Status "$($c.AppName) — $($c.CredentialType)"
          try {
            if ($c.CredentialType -eq "Client Secret") {
              Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($c.ObjectId)/removePassword" -Body (@{ passwordCredentialId = $c.KeyId } | ConvertTo-Json) -ContentType "application/json" | Out-Null
            } else {
              Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($c.ObjectId)/removeKey" -Body (@{ keyCredentialId = $c.KeyId } | ConvertTo-Json) -ContentType "application/json" | Out-Null
            }
            $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Remove cred"; DisplayName = $c.AppName; AppId = $c.AppId; Result = "Success"; ErrorMessage = $null }
          } catch {
            $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Remove cred"; DisplayName = $c.AppName; AppId = $c.AppId; Result = "Failed"; ErrorMessage = $_.Exception.Message }
          }
        }
        Write-Progress -Activity "Removing credentials" -Completed
        Write-Host "  Done. See the Action Log sheet." -ForegroundColor Green
      }
    }
  }
  "3" {
    $actionMode = "Disable SP"
    $targets = @($cleanupTargets | Where-Object { $_.SPObjectId -and $_.SPEnabled })
    if ($targets.Count -eq 0) { Write-Host "  No enabled SPs among stale/never-used apps. Audit only." -ForegroundColor Green; $actionMode = "Audit" }
    elseif (-not (Confirm-WriteAccess)) { $actionMode = "Audit (write permission declined)" }
    else {
      Write-Host ""
      Write-Host "  Preview — SPs to disable (accountEnabled=false):" -ForegroundColor Yellow
      foreach ($p in @($targets | Select-Object -First 10)) { Write-Host "    - $($p.DisplayName)  ($($p.AppId))  [$($p.StaleStatus)]" -ForegroundColor Gray }
      if ($targets.Count -gt 10) { Write-Host "    ... and $($targets.Count - 10) more" -ForegroundColor Gray }
      Write-Host ""
      $cf = Read-Host "  Type YES to disable $($targets.Count) service principal(s)"
      if ($cf -ne 'YES') { Write-Host "  Cancelled — no changes." -ForegroundColor Yellow; $actionMode = "Audit (cancelled disable)" }
      else {
        $j = 0
        foreach ($t in $targets) {
          $j++; Write-ProgressBar -Current $j -Total $targets.Count -Activity "Disabling SPs" -Status $t.DisplayName
          try {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($t.SPObjectId)" -Body (@{ accountEnabled = $false } | ConvertTo-Json) -ContentType "application/json" | Out-Null
            $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Disable SP"; DisplayName = $t.DisplayName; AppId = $t.AppId; Result = "Success"; ErrorMessage = $null }
          } catch {
            $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Disable SP"; DisplayName = $t.DisplayName; AppId = $t.AppId; Result = "Failed"; ErrorMessage = $_.Exception.Message }
          }
        }
        Write-Progress -Activity "Disabling SPs" -Completed
        Write-Host "  Done. See the Action Log sheet." -ForegroundColor Green
      }
    }
  }
  "4" {
    $actionMode = "Delete app registration"
    $targets = @($cleanupTargets | Where-Object { $_.ObjectId })
    if ($targets.Count -eq 0) { Write-Host "  No stale/never-used apps to delete. Audit only." -ForegroundColor Green; $actionMode = "Audit" }
    elseif (-not (Confirm-WriteAccess)) { $actionMode = "Audit (write permission declined)" }
    else {
      Write-Host ""
      Write-Host "  [!] DELETE is PERMANENT. Consumers will immediately fail authentication." -ForegroundColor Red
      Write-Host "  Preview — app registrations to DELETE:" -ForegroundColor Red
      foreach ($p in @($targets | Select-Object -First 10)) { Write-Host "    - $($p.DisplayName)  ($($p.AppId))  [$($p.StaleStatus)]" -ForegroundColor Gray }
      if ($targets.Count -gt 10) { Write-Host "    ... and $($targets.Count - 10) more" -ForegroundColor Gray }
      Write-Host ""
      $cf1 = Read-Host "  Type YES to continue to final confirmation"
      if ($cf1 -ne 'YES') { Write-Host "  Cancelled — no changes." -ForegroundColor Yellow; $actionMode = "Audit (cancelled delete)" }
      else {
        $cf2 = Read-Host "  FINAL CONFIRMATION — type DELETE to permanently remove $($targets.Count) app registration(s)"
        if ($cf2 -ne 'DELETE') { Write-Host "  Cancelled — no changes." -ForegroundColor Yellow; $actionMode = "Audit (cancelled delete)" }
        else {
          $j = 0
          foreach ($t in $targets) {
            $j++; Write-ProgressBar -Current $j -Total $targets.Count -Activity "Deleting app registrations" -Status $t.DisplayName
            try {
              Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$($t.ObjectId)" | Out-Null
              $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Delete app reg"; DisplayName = $t.DisplayName; AppId = $t.AppId; Result = "Success"; ErrorMessage = $null }
            } catch {
              $actionResults += [PSCustomObject]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Action = "Delete app reg"; DisplayName = $t.DisplayName; AppId = $t.AppId; Result = "Failed"; ErrorMessage = $_.Exception.Message }
            }
          }
          Write-Progress -Activity "Deleting app registrations" -Completed
          Write-Host "  Done. See the Action Log sheet." -ForegroundColor Green
        }
      }
    }
  }
  default { $actionMode = "Audit"; Write-Host "  Audit-only run. No changes made." -ForegroundColor Green }
}

# ============================================================
#  Export to Excel (Griffin31 styling)
# ============================================================
Write-Host ""
Write-Host "  Generating Excel report..." -ForegroundColor Cyan

$reportDir = if ($IsMacOS -or $IsLinux) { "$HOME/Desktop" } else { [Environment]::GetFolderPath("Desktop") }
if (-not (Test-Path $reportDir)) { $reportDir = $HOME }
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$exportPath = Join-Path $reportDir "AppRegistration_Audit_$timestamp.xlsx"

function HexColor([string]$hex) {
  $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
  $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
  $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
  return [System.Drawing.Color]::FromArgb($r, $g, $b)
}
$NAVY = HexColor "1B2A4A"; $DARK_BLUE = HexColor "2C3E6B"; $ACCENT_BLUE = HexColor "4472C4"
$LIGHT_BLUE = HexColor "D6E4F0"; $WHITE_C = [System.Drawing.Color]::White; $DARK_GRAY = HexColor "404040"
$ROW_ALT = HexColor "F8F9FA"; $RED_BG = HexColor "FDE8E8"; $RED_TEXT = HexColor "B91C1C"
$AMBER_BG = HexColor "FEF3C7"; $AMBER_TEXT = HexColor "92400E"; $GREEN_BG = HexColor "D1FAE5"
$GREEN_TEXT = HexColor "065F46"; $GRAY_BG = HexColor "E5E7EB"; $GRAY_TEXT = HexColor "374151"

function Set-Fill($cells, $color) {
  $cells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
  $cells.Style.Fill.BackgroundColor.SetColor($color)
}
function Set-Font($cells, [int]$size = 10, [bool]$bold = $false, $color = $null) {
  $cells.Style.Font.Size = $size; $cells.Style.Font.Bold = $bold
  if ($color) { $cells.Style.Font.Color.SetColor($color) }
}
function Set-ThinBorder($cells, $color = $null) {
  if ($null -eq $color) { $color = [System.Drawing.Color]::FromArgb(217, 217, 217) }
  $b = $cells.Style.Border
  foreach ($e in @($b.Top, $b.Bottom, $b.Left, $b.Right)) { $e.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin; $e.Color.SetColor($color) }
}
function Set-RiskCell($cell, [string]$risk) {
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  switch ($risk) {
    'High'   { Set-Fill $cell $RED_BG;   Set-Font $cell -bold $true -color $RED_TEXT }
    'Medium' { Set-Fill $cell $AMBER_BG; Set-Font $cell -bold $true -color $AMBER_TEXT }
    'Low'    { Set-Fill $cell $GREEN_BG; Set-Font $cell -color $GREEN_TEXT }
    default  { Set-Fill $cell $GRAY_BG;  Set-Font $cell -color $GRAY_TEXT }
  }
}
function Set-WlidCell($cell, [string]$val) {
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  switch ($val) {
    'CA + ID Protection'  { Set-Fill $cell $GREEN_BG; Set-Font $cell -bold $true -color $GREEN_TEXT }
    'ID Protection only'  { Set-Fill $cell $AMBER_BG; Set-Font $cell -color $AMBER_TEXT }
    default               { Set-Fill $cell $GRAY_BG;  Set-Font $cell -color $GRAY_TEXT }
  }
}
# Enabled/disabled = enterprise-app (service principal) accountEnabled.
function Get-EnabledLabel($v) { if ($v -eq $true) { 'Enabled' } elseif ($v -eq $false) { 'Disabled' } else { 'No SP' } }
function Set-EnabledCell($cell, [string]$val) {
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  switch ($val) {
    'Enabled'  { Set-Fill $cell $GREEN_BG; Set-Font $cell -color $GREEN_TEXT }
    'Disabled' { Set-Fill $cell $AMBER_BG; Set-Font $cell -bold $true -color $AMBER_TEXT }
    default    { Set-Fill $cell $GRAY_BG;  Set-Font $cell -color $GRAY_TEXT }
  }
}
# App enabled/disabled state keyed by appId — used by Excel (below) and the HTML report.
$appStateByAppId = @{}
foreach ($a in $analysis) { $appStateByAppId[$a.AppId] = Get-EnabledLabel $a.SPEnabled }
foreach ($e in $entAnalysis) { if ($e.AppId -and -not $appStateByAppId.ContainsKey($e.AppId)) { $appStateByAppId[$e.AppId] = Get-EnabledLabel $e.SPEnabled } }

$excel = New-Object OfficeOpenXml.ExcelPackage

# ── Sheet 1: Summary ──
$ws = $excel.Workbook.Worksheets.Add("Summary")
$ws.TabColor = HexColor "2E75B6"; $ws.View.ShowGridLines = $false
$ws.Column(1).Width = 2
for ($c = 2; $c -le 6; $c++) { $ws.Column($c).Width = 25 }
$ws.Row(1).Height = 8; for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[1, $c] $NAVY }
$ws.Row(2).Height = 36; for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[2, $c] $NAVY }
$ws.Cells["B2:F2"].Merge = $true; $ws.Cells["B2"].Value = "APP REGISTRATION AUDIT"
Set-Font $ws.Cells["B2"] -size 18 -bold $true -color $WHITE_C
$ws.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$ws.Row(3).Height = 24; for ($c = 1; $c -le 6; $c++) { Set-Fill $ws.Cells[3, $c] $DARK_BLUE }
$ws.Cells["B3:F3"].Merge = $true
$ws.Cells["B3"].Value = "Tenant: $tenantId  |  Scope: $scopeSummary  |  Expiry: $WarningDays d  |  Stale: $StaleDays d  |  Action: $actionMode  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Set-Font $ws.Cells["B3"] -size 10 -color (HexColor "B0C4DE")

$ws.Row(4).Height = 8; $ws.Row(5).Height = 60
$kpis = @(
  @{ Col = 2; Number = "$($analysis.Count)";      Label = "APPS ANALYZED" },
  @{ Col = 3; Number = "$($highRiskApps.Count)";   Label = "HIGH-RISK PERMS" },
  @{ Col = 4; Number = "$($expiredCreds.Count)";   Label = "EXPIRED CREDS" },
  @{ Col = 5; Number = "$($staleApps.Count)";      Label = "STALE APPS" },
  @{ Col = 6; Number = "$($neverUsed.Count)";      Label = "NEVER USED" }
)
foreach ($kpi in $kpis) {
  $cell = $ws.Cells[5, $kpi.Col]; $cell.IsRichText = $true; $rt = $cell.RichText
  $num = $rt.Add($kpi.Number + "`n"); $num.Size = 22; $num.Bold = $true; $num.Color = $NAVY
  $lbl = $rt.Add($kpi.Label); $lbl.Size = 9; $lbl.Color = $DARK_GRAY
  $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $cell.Style.WrapText = $true; Set-Fill $cell $LIGHT_BLUE; Set-ThinBorder $cell
}

$row = 7
$ws.Cells["B$row"].Value = "PERMISSION RISK BREAKDOWN"; Set-Font $ws.Cells["B$row"] -size 12 -bold $true -color $NAVY; $row++
$ws.Cells["B$row"].Value = "Overall Risk"; $ws.Cells["C$row"].Value = "Apps"
Set-Font $ws.Cells["B$row"] -bold $true -color $WHITE_C; Set-Font $ws.Cells["C$row"] -bold $true -color $WHITE_C
Set-Fill $ws.Cells[$row,2] $ACCENT_BLUE; Set-Fill $ws.Cells[$row,3] $ACCENT_BLUE; $row++
foreach ($lvl in @(
    @{ N='High';   C=$highRiskApps.Count },
    @{ N='Medium'; C=$medRiskApps.Count },
    @{ N='Low';    C=@($analysis | Where-Object { $_.OverallRisk -eq 'Low' }).Count },
    @{ N='None';   C=@($analysis | Where-Object { $_.OverallRisk -eq 'None' }).Count })) {
  $ws.Cells["B$row"].Value = $lvl.N; $ws.Cells["C$row"].Value = $lvl.C
  Set-RiskCell $ws.Cells[$row,2] $lvl.N
  $ws.Cells[$row,3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  Set-ThinBorder $ws.Cells[$row,2]; Set-ThinBorder $ws.Cells[$row,3]; $row++
}

$row++
$ws.Cells["B$row"].Value = "WORKLOAD IDENTITY PROTECTION CANDIDATES"; Set-Font $ws.Cells["B$row"] -size 12 -bold $true -color $NAVY; $row++
$ws.Cells["B$row"].Value = "Eligibility"; $ws.Cells["C$row"].Value = "Apps"
Set-Font $ws.Cells["B$row"] -bold $true -color $WHITE_C; Set-Font $ws.Cells["C$row"] -bold $true -color $WHITE_C
Set-Fill $ws.Cells[$row,2] $ACCENT_BLUE; Set-Fill $ws.Cells[$row,3] $ACCENT_BLUE; $row++
foreach ($wl in @(
    @{ N='CA + ID Protection';  C=$caCandidates.Count },
    @{ N='ID Protection only';  C=$idpOnly.Count })) {
  $ws.Cells["B$row"].Value = $wl.N; $ws.Cells["C$row"].Value = $wl.C
  Set-WlidCell $ws.Cells[$row,2] $wl.N
  $ws.Cells[$row,3].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  Set-ThinBorder $ws.Cells[$row,2]; Set-ThinBorder $ws.Cells[$row,3]; $row++
}
$ws.Cells["B$row"].Value = "CA = single-tenant SPs you can target with Conditional Access (needs Workload ID Premium). ID Protection also covers multi-tenant. Microsoft apps & managed identities excluded."
Set-Font $ws.Cells["B$row"] -size 9 -color $DARK_GRAY; $ws.Cells["B$row:F$row"].Merge = $true; $ws.Cells["B$row"].Style.WrapText = $true; $ws.Row($row).Height = 28

# ── Sheet 2: Permission Risk ──
$ws2 = $excel.Workbook.Worksheets.Add("Permission Risk")
$ws2.TabColor = HexColor "B91C1C"; $ws2.View.ShowGridLines = $false
$pCols = @(2,32),@(3,24),@(4,11),@(5,6),@(6,6),@(7,6),@(8,14),@(9,12),@(10,19),@(11,50),@(12,34)
$ws2.Column(1).Width = 2; foreach ($pc in $pCols) { $ws2.Column($pc[0]).Width = $pc[1] }
$ws2.Row(1).Height = 8; for ($c = 1; $c -le 12; $c++) { Set-Fill $ws2.Cells[1,$c] $NAVY }
$ws2.Row(2).Height = 32; for ($c = 1; $c -le 12; $c++) { Set-Fill $ws2.Cells[2,$c] $NAVY }
$ws2.Cells["B2:L2"].Merge = $true; $ws2.Cells["B2"].Value = "API PERMISSION RISK (overall = highest-risk permission)"
Set-Font $ws2.Cells["B2"] -size 16 -bold $true -color $WHITE_C
$ws2.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$hdrRow = 4; $ws2.Row($hdrRow).Height = 28
$headers2 = @("App Name","Owner","Risk","High","Med","Low","Tenancy","Enabled","Workload ID","Sensitive Permissions","Portal Link")
for ($k = 0; $k -lt $headers2.Count; $k++) {
  $cell = $ws2.Cells[$hdrRow, ($k + 2)]; $cell.Value = $headers2[$k]
  Set-Fill $cell $ACCENT_BLUE; Set-Font $cell -bold $true -color $WHITE_C
  $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
}
$dataRow = $hdrRow + 1
foreach ($a in @($analysis | Where-Object { $_.OverallRisk -ne 'None' } | Sort-Object @{E={Get-RiskRank $_.OverallRisk};Descending=$true}, @{E='HighCount';Descending=$true})) {
  $ws2.Cells[$dataRow,2].Value = $a.DisplayName
  $ws2.Cells[$dataRow,3].Value = $a.Owner
  $ws2.Cells[$dataRow,4].Value = $a.OverallRisk; Set-RiskCell $ws2.Cells[$dataRow,4] $a.OverallRisk
  $ws2.Cells[$dataRow,5].Value = $a.HighCount; $ws2.Cells[$dataRow,6].Value = $a.MedCount; $ws2.Cells[$dataRow,7].Value = $a.LowCount
  $ws2.Cells[$dataRow,8].Value = $a.Tenancy
  $enLbl = Get-EnabledLabel $a.SPEnabled
  $ws2.Cells[$dataRow,9].Value = $enLbl; Set-EnabledCell $ws2.Cells[$dataRow,9] $enLbl
  $ws2.Cells[$dataRow,10].Value = $a.WorkloadIdCand; Set-WlidCell $ws2.Cells[$dataRow,10] $a.WorkloadIdCand
  $ws2.Cells[$dataRow,11].Value = $a.SensitivePerms; $ws2.Cells[$dataRow,11].Style.WrapText = $true
  $ws2.Cells[$dataRow,12].Value = "Open in Entra"; $ws2.Cells[$dataRow,12].Hyperlink = [System.Uri]::new($a.PortalLink)
  Set-Font $ws2.Cells[$dataRow,12] -color $ACCENT_BLUE; $ws2.Cells[$dataRow,12].Style.Font.UnderLine = $true
  foreach ($col in 5,6,7) { $ws2.Cells[$dataRow,$col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
  for ($col = 2; $col -le 12; $col++) { Set-ThinBorder $ws2.Cells[$dataRow,$col] }
  if ($dataRow % 2 -eq 0) { foreach ($col in 2,3,8,11) { Set-Fill $ws2.Cells[$dataRow,$col] $ROW_ALT } }
  $ws2.Cells[$dataRow,11].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
  $dataRow++
}
if ($dataRow -gt ($hdrRow + 1)) { $ws2.Cells[$hdrRow, 2, ($dataRow - 1), 12].AutoFilter = $true }

# ── Helper to build a credentials sheet block ──
function Add-CredSheet($name, $tab, $title, $rows) {
  $w = $excel.Workbook.Worksheets.Add($name); $w.TabColor = HexColor $tab; $w.View.ShowGridLines = $false
  $cw = @(2,30),@(3,15),@(4,22),@(5,13),@(6,13),@(7,13),@(8,24),@(9,12),@(10,14),@(11,30)
  $w.Column(1).Width = 2; foreach ($x in $cw) { $w.Column($x[0]).Width = $x[1] }
  $w.Row(1).Height = 8; for ($c=1;$c-le11;$c++){ Set-Fill $w.Cells[1,$c] $NAVY }
  $w.Row(2).Height = 32; for ($c=1;$c-le11;$c++){ Set-Fill $w.Cells[2,$c] $NAVY }
  $w.Cells["B2:K2"].Merge = $true; $w.Cells["B2"].Value = $title
  Set-Font $w.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $w.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $hr = 4; $w.Row($hr).Height = 28
  $hs = @("App Name","Type","Description","Start","End","Days","Owner","Enabled","Source","Portal Link")
  for ($k=0;$k-lt$hs.Count;$k++){ $cell=$w.Cells[$hr,($k+2)]; $cell.Value=$hs[$k]; Set-Fill $cell $ACCENT_BLUE; Set-Font $cell -bold $true -color $WHITE_C }
  $dr = $hr + 1
  foreach ($c in $rows) {
    $w.Cells[$dr,2].Value=$c.AppName; $w.Cells[$dr,3].Value=$c.CredentialType; $w.Cells[$dr,4].Value=$c.Description
    $w.Cells[$dr,5].Value=$c.StartDate; $w.Cells[$dr,6].Value=$c.EndDate
    $w.Cells[$dr,7].Value= if ($c.Status -eq 'Expired') { [math]::Abs($c.DaysToExpiry) } else { $c.DaysToExpiry }
    $w.Cells[$dr,8].Value=$c.Owner
    $enLbl = if ($c.AppId -and $appStateByAppId.ContainsKey($c.AppId)) { $appStateByAppId[$c.AppId] } else { 'No SP' }
    $w.Cells[$dr,9].Value=$enLbl; Set-EnabledCell $w.Cells[$dr,9] $enLbl
    $w.Cells[$dr,10].Value=$c.Source; $w.Cells[$dr,10].Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $w.Cells[$dr,11].Value="Open in Entra"; $w.Cells[$dr,11].Hyperlink=[System.Uri]::new($c.PortalUrl)
    Set-Font $w.Cells[$dr,11] -color $ACCENT_BLUE; $w.Cells[$dr,11].Style.Font.UnderLine=$true
    $dc=$w.Cells[$dr,7]; $dc.Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    if ($c.Status -eq 'Expired' -or $c.DaysToExpiry -le 7) { Set-Fill $dc $RED_BG; Set-Font $dc -color $RED_TEXT }
    elseif ($c.DaysToExpiry -le 14) { Set-Fill $dc $AMBER_BG; Set-Font $dc -color $AMBER_TEXT }
    else { Set-Fill $dc $GREEN_BG; Set-Font $dc -color $GREEN_TEXT }
    for ($col=2;$col-le11;$col++){ Set-ThinBorder $w.Cells[$dr,$col] }
    if ($dr % 2 -eq 0) { foreach ($col in 2,3,4,5,6,8,10) { Set-Fill $w.Cells[$dr,$col] $ROW_ALT } }
    $dr++
  }
  if ($dr -gt ($hr + 1)) { $w.Cells[$hr, 2, ($dr - 1), 11].AutoFilter = $true }
}
# Credential sheets combine app-registration and enterprise-app credentials (Source column distinguishes them).
$allExpiredCreds  = @($expiredCreds  + $entExpired  | Sort-Object DaysToExpiry)
$allExpiringCreds = @($expiringCreds + $entExpiring | Sort-Object DaysToExpiry)
if ($allExpiredCreds.Count -gt 0)  { Add-CredSheet "Expired Creds"  "E74C3C" "EXPIRED CREDENTIALS" $allExpiredCreds }
if ($allExpiringCreds.Count -gt 0) { Add-CredSheet "Expiring Creds" "E67E22" "EXPIRING SOON (within $WarningDays days)" $allExpiringCreds }

# ── Sheet: Stale & Unused ──
$cleanupAll = @($staleApps + $neverUsed + $orphaned)
if ($cleanupAll.Count -gt 0) {
  $ws5 = $excel.Workbook.Worksheets.Add("Stale & Unused")
  $ws5.TabColor = HexColor "7F8C8D"; $ws5.View.ShowGridLines = $false
  $sc = @(2,32),@(3,18),@(4,14),@(5,16),@(6,20),@(7,10),@(8,24),@(9,12),@(10,30)
  $ws5.Column(1).Width = 2; foreach ($x in $sc) { $ws5.Column($x[0]).Width = $x[1] }
  $ws5.Row(1).Height = 8; for ($c=1;$c-le10;$c++){ Set-Fill $ws5.Cells[1,$c] $NAVY }
  $ws5.Row(2).Height = 32; for ($c=1;$c-le10;$c++){ Set-Fill $ws5.Cells[2,$c] $NAVY }
  $ws5.Cells["B2:J2"].Merge = $true; $ws5.Cells["B2"].Value = "STALE & UNUSED APPS"
  Set-Font $ws5.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws5.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $hr = 4; $ws5.Row($hr).Height = 28
  $hs = @("App Name","Activity","Days Since","Last Sign-In","Last Flow","Creds","Owner","Enabled","Portal Link")
  for ($k=0;$k-lt$hs.Count;$k++){ $cell=$ws5.Cells[$hr,($k+2)]; $cell.Value=$hs[$k]; Set-Fill $cell $ACCENT_BLUE; Set-Font $cell -bold $true -color $WHITE_C }
  $dr = $hr + 1
  foreach ($a in @($cleanupAll | Sort-Object @{E='DaysSinceLastSignIn';Descending=$true})) {
    $ws5.Cells[$dr,2].Value=$a.DisplayName
    $ws5.Cells[$dr,3].Value=$a.StaleStatus
    $ws5.Cells[$dr,4].Value= if ($null -ne $a.DaysSinceLastSignIn) { $a.DaysSinceLastSignIn } else { "—" }
    $ws5.Cells[$dr,5].Value= if ($a.LastSignIn) { ([datetime]$a.LastSignIn).ToString('yyyy-MM-dd') } else { "—" }
    $ws5.Cells[$dr,6].Value= if ($a.SignInFlow) { $a.SignInFlow } else { "—" }
    $ws5.Cells[$dr,7].Value=$a.TotalCreds
    $ws5.Cells[$dr,8].Value=$a.Owner
    $enLbl = Get-EnabledLabel $a.SPEnabled
    $ws5.Cells[$dr,9].Value=$enLbl; Set-EnabledCell $ws5.Cells[$dr,9] $enLbl
    $ws5.Cells[$dr,10].Value="Open in Entra"; $ws5.Cells[$dr,10].Hyperlink=[System.Uri]::new($a.PortalLink)
    Set-Font $ws5.Cells[$dr,10] -color $ACCENT_BLUE; $ws5.Cells[$dr,10].Style.Font.UnderLine=$true
    $sCell=$ws5.Cells[$dr,3]; $sCell.Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    switch ($a.StaleStatus) {
      'Stale'      { Set-Fill $sCell $AMBER_BG; Set-Font $sCell -color $AMBER_TEXT }
      'Never used' { Set-Fill $sCell $RED_BG;   Set-Font $sCell -color $RED_TEXT }
      default      { Set-Fill $sCell $GRAY_BG;  Set-Font $sCell -color $GRAY_TEXT }
    }
    for ($col=2;$col-le10;$col++){ Set-ThinBorder $ws5.Cells[$dr,$col] }
    if ($dr % 2 -eq 0) { foreach ($col in 2,4,5,6,7,8) { Set-Fill $ws5.Cells[$dr,$col] $ROW_ALT } }
    $dr++
  }
  if ($dr -gt ($hr + 1)) { $ws5.Cells[$hr, 2, ($dr - 1), 10].AutoFilter = $true }
}

# ── Sheet: Enterprise Apps (actual granted permissions) ──
$entWithPerms = @($entAnalysis | Where-Object { $_.OverallRisk -ne 'None' })
if ($entWithPerms.Count -gt 0) {
  $ws7 = $excel.Workbook.Worksheets.Add("Enterprise Apps")
  $ws7.TabColor = HexColor "8E44AD"; $ws7.View.ShowGridLines = $false
  $ec = @(2,32),@(3,24),@(4,11),@(5,6),@(6,6),@(7,6),@(8,10),@(9,18),@(10,12),@(11,16),@(12,16),@(13,50),@(14,30)
  $ws7.Column(1).Width = 2; foreach ($x in $ec) { $ws7.Column($x[0]).Width = $x[1] }
  $ws7.Row(1).Height = 8; for ($c=1;$c-le14;$c++){ Set-Fill $ws7.Cells[1,$c] $NAVY }
  $ws7.Row(2).Height = 32; for ($c=1;$c-le14;$c++){ Set-Fill $ws7.Cells[2,$c] $NAVY }
  $ws7.Cells["B2:N2"].Merge = $true; $ws7.Cells["B2"].Value = "ENTERPRISE APPS — GRANTED PERMISSIONS (what apps actually hold consent to do)"
  Set-Font $ws7.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $ws7.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
  $hr = 4; $ws7.Row($hr).Height = 28
  $hs = @("App Name","Owner","Risk","High","Med","Low","App Reg?","Type","Enabled","Activity","Category","Granted Permissions","Portal Link")
  for ($k=0;$k-lt$hs.Count;$k++){ $cell=$ws7.Cells[$hr,($k+2)]; $cell.Value=$hs[$k]; Set-Fill $cell $ACCENT_BLUE; Set-Font $cell -bold $true -color $WHITE_C; $cell.Style.VerticalAlignment=[OfficeOpenXml.Style.ExcelVerticalAlignment]::Center }
  $dr = $hr + 1
  foreach ($e in @($entWithPerms | Sort-Object @{E={Get-RiskRank $_.OverallRisk};Descending=$true}, @{E='HighCount';Descending=$true})) {
    $ws7.Cells[$dr,2].Value=$e.DisplayName
    $ws7.Cells[$dr,3].Value=$e.Owner
    $ws7.Cells[$dr,4].Value=$e.OverallRisk; Set-RiskCell $ws7.Cells[$dr,4] $e.OverallRisk
    $ws7.Cells[$dr,5].Value=$e.HighCount; $ws7.Cells[$dr,6].Value=$e.MedCount; $ws7.Cells[$dr,7].Value=$e.LowCount
    $ws7.Cells[$dr,8].Value= if ($e.HasAppReg) { 'Yes' } else { 'No' }
    $ws7.Cells[$dr,9].Value=$e.SPType
    $enLbl = Get-EnabledLabel $e.SPEnabled
    $ws7.Cells[$dr,10].Value=$enLbl; Set-EnabledCell $ws7.Cells[$dr,10] $enLbl
    $ws7.Cells[$dr,11].Value=$e.StaleStatus
    $aCell=$ws7.Cells[$dr,11]; $aCell.Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    switch ($e.StaleStatus) {
      'Active'     { Set-Fill $aCell $GREEN_BG; Set-Font $aCell -color $GREEN_TEXT }
      'Stale'      { Set-Fill $aCell $AMBER_BG; Set-Font $aCell -color $AMBER_TEXT }
      'Never used' { Set-Fill $aCell $RED_BG;   Set-Font $aCell -color $RED_TEXT }
      default      { Set-Fill $aCell $GRAY_BG;  Set-Font $aCell -color $GRAY_TEXT }
    }
    $ws7.Cells[$dr,12].Value=$e.Category
    $ws7.Cells[$dr,13].Value=$e.SensitivePerms; $ws7.Cells[$dr,13].Style.WrapText=$true; $ws7.Cells[$dr,13].Style.VerticalAlignment=[OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    $ws7.Cells[$dr,14].Value="Open in Entra"; $ws7.Cells[$dr,14].Hyperlink=[System.Uri]::new($e.PortalLink)
    Set-Font $ws7.Cells[$dr,14] -color $ACCENT_BLUE; $ws7.Cells[$dr,14].Style.Font.UnderLine=$true
    foreach ($col in 5,6,7,8) { $ws7.Cells[$dr,$col].Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    for ($col=2;$col-le14;$col++){ Set-ThinBorder $ws7.Cells[$dr,$col] }
    if ($dr % 2 -eq 0) { foreach ($col in 2,3,9,12,13) { Set-Fill $ws7.Cells[$dr,$col] $ROW_ALT } }
    $dr++
  }
  if ($dr -gt ($hr + 1)) { $ws7.Cells[$hr, 2, ($dr - 1), 14].AutoFilter = $true }
}

# ── Sheet: Action Log ──
if ($actionResults.Count -gt 0) {
  $ws6 = $excel.Workbook.Worksheets.Add("Action Log")
  $ws6.TabColor = HexColor "E74C3C"; $ws6.View.ShowGridLines = $false
  $ac = @(2,20),@(3,16),@(4,32),@(5,12),@(6,45)
  $ws6.Column(1).Width = 2; foreach ($x in $ac) { $ws6.Column($x[0]).Width = $x[1] }
  $ws6.Row(1).Height = 8; for ($c=1;$c-le6;$c++){ Set-Fill $ws6.Cells[1,$c] $NAVY }
  $ws6.Row(2).Height = 32; for ($c=1;$c-le6;$c++){ Set-Fill $ws6.Cells[2,$c] $NAVY }
  $ws6.Cells["B2:F2"].Merge = $true; $ws6.Cells["B2"].Value = "ACTION LOG"
  Set-Font $ws6.Cells["B2"] -size 16 -bold $true -color $WHITE_C
  $hr = 4; $ws6.Row($hr).Height = 28
  $hs = @("Timestamp","Action","App Name","Result","Error")
  for ($k=0;$k-lt$hs.Count;$k++){ $cell=$ws6.Cells[$hr,($k+2)]; $cell.Value=$hs[$k]; Set-Fill $cell $ACCENT_BLUE; Set-Font $cell -bold $true -color $WHITE_C }
  $dr = $hr + 1
  foreach ($r in $actionResults) {
    $ws6.Cells[$dr,2].Value=$r.Timestamp; $ws6.Cells[$dr,3].Value=$r.Action; $ws6.Cells[$dr,4].Value=$r.DisplayName
    $ws6.Cells[$dr,5].Value=$r.Result; $ws6.Cells[$dr,6].Value=$r.ErrorMessage
    $rc=$ws6.Cells[$dr,5]
    if ($r.Result -eq 'Success') { Set-Fill $rc $GREEN_BG; Set-Font $rc -color $GREEN_TEXT } else { Set-Fill $rc $RED_BG; Set-Font $rc -color $RED_TEXT }
    for ($col=2;$col-le6;$col++){ Set-ThinBorder $ws6.Cells[$dr,$col] }
    $dr++
  }
}

$excel.SaveAs($exportPath)
$excel.Dispose()

# ============================================================
#  Export interactive HTML report
# ============================================================
Write-Host "  Generating interactive HTML report..." -ForegroundColor Cyan
$htmlPath = Join-Path $reportDir "AppRegistration_Audit_$timestamp.html"

$reportApps = $analysis | Select-Object DisplayName, Owner, OverallRisk, HighCount, MedCount, LowCount,
  TotalPerms, @{n='Sensitive';e={ $_.SensitivePerms }},
  @{n='Products';e={ @($_.Products) }}, @{n='PermSearch';e={ $_.PermSearch }},
  Tenancy, WorkloadIdCand, StaleStatus,
  @{n='Status';e={ if ($_.SPEnabled -eq $true) { 'Enabled' } elseif ($_.SPEnabled -eq $false) { 'Disabled' } else { 'No SP' } }},
  @{n='DaysSince';e={ $_.DaysSinceLastSignIn }},
  @{n='LastSignIn';e={ if ($_.LastSignIn) { ([datetime]$_.LastSignIn).ToString('yyyy-MM-dd') } else { '' } }},
  @{n='SignInFlow';e={ $_.SignInFlow }},
  TotalCreds, ExpiredCreds, ExpiringCreds, @{n='Portal';e={ $_.PortalLink }}

# Credentials table = app-registration creds + enterprise-app creds, tagged by Source.
$reportCreds = @($allCreds + $entCreds) | Select-Object AppName, CredentialType, Description, StartDate, EndDate,
  DaysToExpiry, Status, Owner,
  @{n='AppState';e={ if ($_.AppId -and $appStateByAppId.ContainsKey($_.AppId)) { $appStateByAppId[$_.AppId] } else { 'No SP' } }},
  @{n='Source';e={ $_.Source }},
  @{n='Portal';e={ $_.PortalUrl }}

# Enterprise apps = identities holding actual granted consent (incl. apps with no app registration).
$reportEntApps = $entAnalysis | Where-Object { $_.OverallRisk -ne 'None' } | Select-Object DisplayName, Owner, OverallRisk, HighCount, MedCount, LowCount,
  TotalPerms, @{n='Sensitive';e={ $_.SensitivePerms }},
  @{n='Products';e={ @($_.Products) }}, @{n='PermSearch';e={ $_.PermSearch }},
  Tenancy, @{n='Category';e={ $_.Category }}, @{n='SPType';e={ $_.SPType }}, @{n='HasAppReg';e={ [bool]$_.HasAppReg }},
  @{n='Status';e={ if ($_.SPEnabled -eq $true) { 'Enabled' } elseif ($_.SPEnabled -eq $false) { 'Disabled' } else { 'No SP' } }},
  StaleStatus,
  @{n='LastSignIn';e={ if ($_.LastSignIn) { ([datetime]$_.LastSignIn).ToString('yyyy-MM-dd') } else { '' } }},
  @{n='SignInFlow';e={ $_.SignInFlow }},
  TotalCreds, @{n='Portal';e={ $_.PortalLink }}

$payload = [ordered]@{
  meta = [ordered]@{ tenant = "$tenantId"; account = "$($ctx.Account)"; generated = (Get-Date -Format 'yyyy-MM-dd HH:mm'); expiry = $WarningDays; stale = $StaleDays; action = "$actionMode"; scope = "$scopeSummary" }
  kpis = [ordered]@{ apps = $analysis.Count; high = $highRiskApps.Count; medium = $medRiskApps.Count; expired = $expiredCreds.Count; expiring = $expiringCreds.Count; stale = $staleApps.Count; never = $neverUsed.Count; caCand = $caCandidates.Count; idpOnly = $idpOnly.Count; entApps = $entAnalysis.Count; entHigh = $entHigh.Count; entNoReg = $entNoAppReg.Count }
  apps = @($reportApps)
  creds = @($reportCreds)
  entApps = @($reportEntApps)
}
$json = $payload | ConvertTo-Json -Depth 6 -Compress
# Harden the JSON embedded in <script>: a maliciously named app registration could
# otherwise inject markup. '<' and '>' appear only inside JSON string values, so
# escaping them (plus the JS line separators U+2028/U+2029) is lossless and blocks
# every <script>/<!-- breakout while the data still renders identically.
$bs = [char]92   # backslash built via char code so the escape sequence can't be mangled
$json = $json.Replace('<', "${bs}u003c").Replace('>', "${bs}u003e")
$json = $json.Replace([string][char]0x2028, "${bs}u2028").Replace([string][char]0x2029, "${bs}u2029")

$htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>App Registration Audit</title>
<!-- Self-contained report: no external fonts, scripts, or network calls. Uses local fonts if installed, else system fallbacks. -->
<style>
:root{
  --bg:#070b14; --panel:#0d1424; --panel-2:#111c30; --line:#1b2842; --line-2:#26375a;
  --ink:#e8edf7; --muted:#8a98b5; --dim:#5d6c8c;
  --hi:#fb5570; --hi-bg:rgba(251,85,112,.13); --hi-gl:rgba(251,85,112,.35);
  --med:#f6a93b; --med-bg:rgba(246,169,59,.13);
  --low:#3dd7a6; --low-bg:rgba(61,215,166,.12);
  --accent:#46e0d0; --accent-2:#5b8cff;
  --mono:'JetBrains Mono','SF Mono',ui-monospace,Menlo,Consolas,monospace; --sans:'Sora',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{
  font-family:var(--sans); color:var(--ink); background:var(--bg); line-height:1.5;
  min-height:100vh; padding:0 clamp(16px,4vw,56px) 80px;
  background-image:
    radial-gradient(1100px 560px at 6% -8%, rgba(70,224,208,.10), transparent 60%),
    radial-gradient(900px 520px at 102% 112%, rgba(251,85,112,.09), transparent 60%);
  background-attachment:fixed;
}
body::before{
  content:""; position:fixed; inset:0; pointer-events:none; z-index:0; opacity:.4;
  background-image:linear-gradient(var(--line) 1px,transparent 1px),linear-gradient(90deg,var(--line) 1px,transparent 1px);
  background-size:54px 54px; mask-image:radial-gradient(circle at 50% 0%,#000,transparent 75%);
}
.wrap{position:relative; z-index:1; max-width:1400px; margin:0 auto}

/* Header */
header{display:flex; flex-wrap:wrap; align-items:flex-end; justify-content:space-between; gap:20px; padding:38px 0 26px}
.brand{display:flex; flex-direction:column; gap:6px}
.logo{font-family:var(--mono); font-weight:800; font-size:13px; letter-spacing:.32em; color:var(--accent); text-transform:uppercase}
.logo .cur{display:inline-block; width:9px; height:15px; background:var(--accent); margin-left:4px; transform:translateY(2px); animation:blink 1.1s steps(1) infinite}
@keyframes blink{50%{opacity:0}}
h1{font-size:clamp(26px,3.4vw,40px); font-weight:700; letter-spacing:-.02em; line-height:1.05}
h1 b{color:var(--accent)}
.meta{font-family:var(--mono); font-size:11.5px; color:var(--muted); text-align:right; line-height:1.9}
.meta span{color:var(--ink)}
.meta .tag{display:inline-block; padding:2px 9px; border:1px solid var(--line-2); border-radius:20px; color:var(--accent); margin-left:6px}

/* KPIs */
.kpis{display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:14px; margin:6px 0 30px}
.kpi{position:relative; background:linear-gradient(160deg,var(--panel-2),var(--panel)); border:1px solid var(--line); border-radius:14px; padding:20px 20px 18px; overflow:hidden; opacity:0; transform:translateY(14px); animation:rise .6s cubic-bezier(.2,.7,.2,1) forwards}
.kpi::after{content:""; position:absolute; left:0; top:0; bottom:0; width:3px; background:var(--accent)}
.kpi.hi::after{background:var(--hi); box-shadow:0 0 18px var(--hi-gl)}
.kpi.med::after{background:var(--med)} .kpi.low::after{background:var(--low)} .kpi.acc::after{background:var(--accent-2)}
.kpi .n{font-family:var(--mono); font-size:38px; font-weight:800; letter-spacing:-.03em; line-height:1}
.kpi .l{font-family:var(--mono); font-size:10.5px; letter-spacing:.16em; text-transform:uppercase; color:var(--muted); margin-top:10px}
.kpi .s{font-size:11px; color:var(--dim); margin-top:3px}
@keyframes rise{to{opacity:1; transform:none}}

/* Tabs + controls */
.bar{display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:16px; margin-bottom:14px; border-bottom:1px solid var(--line); padding-bottom:0}
.tabs{display:flex; gap:4px; flex-wrap:wrap}
.tab{font-family:var(--mono); font-size:12px; letter-spacing:.05em; color:var(--muted); background:none; border:none; padding:11px 15px; cursor:pointer; position:relative; transition:color .2s}
.tab:hover{color:var(--ink)}
.tab.on{color:var(--accent)}
.tab.on::after{content:""; position:absolute; left:10px; right:10px; bottom:-1px; height:2px; background:var(--accent); box-shadow:0 0 10px var(--accent); border-radius:2px}
.tab .c{color:var(--dim); margin-left:6px; font-size:11px}
.ctrl{display:flex; gap:10px; align-items:center; padding-bottom:10px}
.search{font-family:var(--mono); font-size:12px; background:var(--panel); border:1px solid var(--line-2); color:var(--ink); padding:9px 13px; border-radius:9px; width:210px; outline:none; transition:border .2s,box-shadow .2s}
.search:focus{border-color:var(--accent); box-shadow:0 0 0 3px rgba(70,224,208,.12)}
.search::placeholder{color:var(--dim)}
.sel{width:auto; max-width:190px; cursor:pointer; appearance:none; -webkit-appearance:none; padding-right:26px; background-image:linear-gradient(45deg,transparent 50%,var(--muted) 50%),linear-gradient(135deg,var(--muted) 50%,transparent 50%); background-position:calc(100% - 15px) center,calc(100% - 10px) center; background-size:5px 5px,5px 5px; background-repeat:no-repeat}
.sel option{background:var(--panel); color:var(--ink)}
.chips{display:flex; gap:5px}
.chip{font-family:var(--mono); font-size:10.5px; letter-spacing:.06em; text-transform:uppercase; padding:7px 11px; border-radius:20px; border:1px solid var(--line-2); background:none; color:var(--muted); cursor:pointer; transition:.18s}
.chip:hover{color:var(--ink); border-color:var(--dim)}
.chip.on{color:var(--bg); font-weight:700}
.chip.on[data-r="all"]{background:var(--accent); border-color:var(--accent)}
.chip.on[data-r="High"]{background:var(--hi); border-color:var(--hi)}
.chip.on[data-r="Medium"]{background:var(--med); border-color:var(--med)}
.chip.on[data-r="Low"]{background:var(--low); border-color:var(--low)}

/* Table */
.tablewrap{background:var(--panel); border:1px solid var(--line); border-radius:14px; overflow:hidden}
table{width:100%; border-collapse:collapse; font-size:13px}
thead th{font-family:var(--mono); font-size:10.5px; letter-spacing:.13em; text-transform:uppercase; color:var(--muted); text-align:left; padding:14px 16px; background:var(--panel-2); border-bottom:1px solid var(--line-2); cursor:pointer; user-select:none; white-space:nowrap; position:sticky; top:0}
thead th:hover{color:var(--ink)}
thead th .ar{opacity:.4; font-size:9px; margin-left:4px}
thead th.sorted .ar{opacity:1; color:var(--accent)}
th.num,td.num{text-align:center}
tbody tr{border-bottom:1px solid var(--line); transition:background .15s; opacity:0; animation:fadein .4s ease forwards}
@keyframes fadein{to{opacity:1}}
tbody tr.row:hover{background:var(--panel-2)}
tbody tr.row.exp{cursor:pointer}
td{padding:12px 16px; vertical-align:middle}
td.app{font-weight:600; max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
td.mono{font-family:var(--mono); font-size:12px; color:var(--muted)}
.badge{display:inline-flex; align-items:center; gap:6px; font-family:var(--mono); font-size:11px; font-weight:700; letter-spacing:.03em; padding:4px 10px; border-radius:6px; white-space:nowrap}
.badge::before{content:""; width:6px; height:6px; border-radius:50%}
.b-High{background:var(--hi-bg); color:var(--hi)} .b-High::before{background:var(--hi); box-shadow:0 0 8px var(--hi)}
.b-Medium{background:var(--med-bg); color:var(--med)} .b-Medium::before{background:var(--med)}
.b-Low{background:var(--low-bg); color:var(--low)} .b-Low::before{background:var(--low)}
.b-None{background:rgba(138,152,181,.1); color:var(--dim)} .b-None::before{background:var(--dim)}
.wl{font-family:var(--mono); font-size:11px; font-weight:600; padding:3px 9px; border-radius:20px; white-space:nowrap; border:1px solid transparent}
.wl-ca{color:var(--low); background:var(--low-bg); border-color:rgba(61,215,166,.3)}
.wl-idp{color:var(--med); background:var(--med-bg); border-color:rgba(246,169,59,.25)}
.wl-no{color:var(--dim); background:rgba(138,152,181,.08)}
.st{font-family:var(--mono); font-size:11px; padding:3px 9px; border-radius:6px}
.st-Stale{color:var(--med); background:var(--med-bg)} .st-Never{color:var(--hi); background:var(--hi-bg)}
.st-Active{color:var(--low); background:var(--low-bg)} .st-other{color:var(--dim); background:rgba(138,152,181,.08)}
.st-Expired{color:var(--hi); background:var(--hi-bg)} .st-Expiring{color:var(--med); background:var(--med-bg)}
.dim{color:var(--dim)}
a.link{color:var(--accent-2); text-decoration:none; font-family:var(--mono); font-size:11px; border-bottom:1px dashed rgba(91,140,255,.4)}
a.link:hover{border-bottom-style:solid}
.detail{background:#0a1120}
.detail td{padding:0}
.detail .inner{padding:6px 20px 18px 20px; display:grid; grid-template-columns:1fr; gap:6px}
.detail h4{font-family:var(--mono); font-size:10.5px; letter-spacing:.14em; text-transform:uppercase; color:var(--muted); margin:8px 0 4px}
.perm{font-family:var(--mono); font-size:12px; padding:6px 12px; border-radius:7px; border-left:3px solid var(--line-2); background:var(--panel)}
.perm.p-high{border-left-color:var(--hi)} .perm.p-med{border-left-color:var(--med)}
.perm .t{color:var(--dim); font-size:10px; margin-left:8px}
.empty{padding:60px 20px; text-align:center; color:var(--dim); font-family:var(--mono); font-size:13px}
.foot{margin-top:26px; font-family:var(--mono); font-size:11px; color:var(--dim); display:flex; justify-content:space-between; flex-wrap:wrap; gap:10px}
.count{font-family:var(--mono); font-size:11px; color:var(--muted); padding:10px 16px; border-top:1px solid var(--line)}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="brand">
      <div class="logo">G31 // APP-REG AUDIT<span class="cur"></span></div>
      <h1>App Registration <b>Risk</b> Report</h1>
    </div>
    <div class="meta" id="meta"></div>
  </header>

  <section class="kpis" id="kpis"></section>

  <div class="bar">
    <div class="tabs" id="tabs"></div>
    <div class="ctrl">
      <div class="chips" id="chips"></div>
      <select class="search sel" id="state"></select>
      <select class="search sel" id="product"></select>
      <input class="search" id="search" placeholder="search apps, perms, products…" autocomplete="off">
    </div>
  </div>

  <div class="tablewrap">
    <div id="table"></div>
    <div class="count" id="count"></div>
  </div>

  <div class="foot">
    <span>Generated by Entra-AppRegistration-Audit · Griffin31 ToolKit</span>
    <span id="footmeta"></span>
  </div>
</div>

<script>
const D = __DATA_JSON__;
const E = s => (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const riskRank = r => ({High:3,Medium:2,Low:1,None:0}[r]??0);
const badge = r => `<span class="badge b-${E(r)}">${E(r)}</span>`;
function wlBadge(v){
  if(v && v.indexOf('CA')===0) return `<span class="wl wl-ca">CA + ID Protection</span>`;
  if(v && v.indexOf('ID Protection')===0) return `<span class="wl wl-idp">ID Protection only</span>`;
  return `<span class="wl wl-no">${E(v||'—')}</span>`;
}
function stBadge(v){
  const k = v==='Stale'?'Stale':v==='Never used'?'Never':v==='Active'?'Active':v==='Expired'?'Expired':v==='Expiring Soon'?'Expiring':'other';
  return `<span class="st st-${k}">${E(v)}</span>`;
}
function enBadge(v){
  const k = v==='Enabled'?'Active':v==='Disabled'?'Stale':'other';
  return `<span class="st st-${k}">${E(v)}</span>`;
}

let state = {tab:'risk', q:'', risk:'all', product:'all', enabled:'all', sort:{key:null,dir:-1}};
const prods = a => Array.isArray(a.Products) ? a.Products : (a.Products ? [a.Products] : []);

const TABS = [
  {id:'risk',       label:'App Reg Permissions', count:()=>D.apps.filter(a=>a.OverallRisk!=='None').length},
  {id:'enterprise', label:'Enterprise Apps',     count:()=>(D.entApps||[]).length},
  {id:'workload',   label:'Workload ID',         count:()=>D.apps.filter(a=>a.WorkloadIdCand&&(a.WorkloadIdCand.indexOf('CA')===0||a.WorkloadIdCand.indexOf('ID Protection')===0)).length},
  {id:'creds',      label:'Credentials',         count:()=>D.creds.length},
  {id:'stale',      label:'Stale & Unused',      count:()=>D.apps.filter(a=>['Stale','Never used','No service principal'].includes(a.StaleStatus)).length},
];

// ---- meta + kpis ----
const m = D.meta;
document.getElementById('meta').innerHTML =
  `tenant <span>${E(m.tenant)}</span><br>${E(m.account)}<br>${E(m.generated)} · expiry <span>${m.expiry}d</span> · stale <span>${m.stale}d</span><span class="tag">${E(m.action)}</span>${m.scope?`<br>scope: <span>${E(m.scope)}</span>`:''}`;
document.getElementById('footmeta').textContent = `${m.tenant} · ${m.generated}`;

const K = D.kpis;
const kpiDefs = [
  {n:K.apps,   l:'App Regs Analyzed', cls:'acc'},
  {n:K.high,   l:'High-Risk Perms', cls:'hi', s:`${K.medium} medium`},
  {n:K.entApps,l:'Enterprise Apps', cls:K.entHigh?'hi':'acc', s:`${K.entHigh} high-risk · ${K.entNoReg} no app reg`},
  {n:K.expired,l:'Expired Creds', cls:K.expired?'hi':'low', s:`${K.expiring} expiring`},
  {n:K.stale,  l:'Stale Apps', cls:'med', s:`${K.never} never used`},
  {n:K.caCand, l:'CA-Protectable', cls:'low', s:`${K.idpOnly} ID-Protection only`},
];
document.getElementById('kpis').innerHTML = kpiDefs.map((k,i)=>
  `<div class="kpi ${k.cls}" style="animation-delay:${i*70}ms"><div class="n" data-to="${k.n}">0</div><div class="l">${k.l}</div>${k.s?`<div class="s">${k.s}</div>`:''}</div>`).join('');
// count-up
document.querySelectorAll('.kpi .n').forEach(el=>{
  const to=+el.dataset.to, dur=850, t0=performance.now();
  (function step(t){const p=Math.min(1,(t-t0)/dur); el.textContent=Math.round(to*(1-Math.pow(1-p,3))); if(p<1)requestAnimationFrame(step);})(t0);
});

// ---- tabs + chips ----
document.getElementById('tabs').innerHTML = TABS.map(t=>
  `<button class="tab ${t.id===state.tab?'on':''}" data-t="${t.id}">${t.label}<span class="c">${t.count()}</span></button>`).join('');
document.getElementById('chips').innerHTML = ['all','High','Medium','Low'].map(r=>
  `<button class="chip ${r===state.risk?'on':''}" data-r="${r}">${r==='all'?'all':r}</button>`).join('');
const PRODUCTS = [...new Set(D.apps.flatMap(prods))].filter(Boolean).sort((a,b)=>a.toLowerCase()<b.toLowerCase()?-1:1);
document.getElementById('product').innerHTML =
  `<option value="all">all products</option>` + PRODUCTS.map(p=>`<option value="${E(p)}">${E(p)}</option>`).join('');
document.getElementById('state').innerHTML =
  `<option value="all">all states</option>` + ['Enabled','Disabled','No SP'].map(s=>`<option value="${s}">${s}</option>`).join('');

document.getElementById('tabs').addEventListener('click',e=>{
  const b=e.target.closest('.tab'); if(!b)return;
  state.tab=b.dataset.t; state.sort={key:null,dir:-1};
  document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('on',x.dataset.t===state.tab));
  document.getElementById('chips').style.visibility = (state.tab==='creds'||state.tab==='stale')?'hidden':'visible';
  document.getElementById('product').style.display = (state.tab==='creds')?'none':'';
  render();
});
document.getElementById('chips').addEventListener('click',e=>{
  const b=e.target.closest('.chip'); if(!b)return;
  state.risk=b.dataset.r;
  document.querySelectorAll('.chip').forEach(x=>x.classList.toggle('on',x.dataset.r===state.risk));
  render();
});
document.getElementById('product').addEventListener('change',e=>{state.product=e.target.value; render();});
document.getElementById('state').addEventListener('change',e=>{state.enabled=e.target.value; render();});
document.getElementById('search').addEventListener('input',e=>{state.q=e.target.value.toLowerCase(); render();});

// ---- column defs per tab ----
const COLS = {
  risk:[
    {k:'DisplayName',t:'Application',c:'app'},
    {k:'OverallRisk',t:'Risk',r:a=>badge(a.OverallRisk),sk:a=>riskRank(a.OverallRisk)},
    {k:'HighCount',t:'High',num:1},{k:'MedCount',t:'Med',num:1},{k:'LowCount',t:'Low',num:1},
    {k:'WorkloadIdCand',t:'Workload ID',r:a=>wlBadge(a.WorkloadIdCand)},
    {k:'Tenancy',t:'Tenancy',c:'mono'},
    {k:'Status',t:'Enabled',r:a=>enBadge(a.Status)},
    {k:'Owner',t:'Owner',c:'mono'},
  ],
  workload:[
    {k:'DisplayName',t:'Application',c:'app'},
    {k:'WorkloadIdCand',t:'Eligibility',r:a=>wlBadge(a.WorkloadIdCand),sk:a=>(a.WorkloadIdCand||'').indexOf('CA')===0?2:(a.WorkloadIdCand||'').indexOf('ID')===0?1:0},
    {k:'OverallRisk',t:'Risk',r:a=>badge(a.OverallRisk),sk:a=>riskRank(a.OverallRisk)},
    {k:'Tenancy',t:'Tenancy',c:'mono'},
    {k:'Status',t:'Enabled',r:a=>enBadge(a.Status)},
    {k:'StaleStatus',t:'Activity',r:a=>stBadge(a.StaleStatus)},
    {k:'LastSignIn',t:'Last Sign-In',c:'mono',r:a=>a.LastSignIn||'<span class="dim">—</span>'},
    {k:'SignInFlow',t:'Last Flow',c:'mono',r:a=>a.SignInFlow||'<span class="dim">—</span>'},
    {k:'Owner',t:'Owner',c:'mono'},
  ],
  enterprise:[
    {k:'DisplayName',t:'Application',c:'app'},
    {k:'OverallRisk',t:'Risk',r:a=>badge(a.OverallRisk),sk:a=>riskRank(a.OverallRisk)},
    {k:'HighCount',t:'High',num:1},{k:'MedCount',t:'Med',num:1},{k:'LowCount',t:'Low',num:1},
    {k:'Category',t:'Category',c:'mono'},
    {k:'HasAppReg',t:'App Reg?',r:a=>a.HasAppReg?'<span class="st st-Active">Yes</span>':'<span class="st st-Never">No</span>',sk:a=>a.HasAppReg?1:0},
    {k:'SPType',t:'Type',c:'mono'},
    {k:'Status',t:'Enabled',r:a=>enBadge(a.Status)},
    {k:'StaleStatus',t:'Activity',r:a=>stBadge(a.StaleStatus)},
    {k:'SignInFlow',t:'Last Flow',c:'mono',r:a=>a.SignInFlow||'<span class="dim">—</span>'},
    {k:'Owner',t:'Owner',c:'mono'},
  ],
  creds:[
    {k:'AppName',t:'Application',c:'app'},
    {k:'Source',t:'Source',c:'mono'},
    {k:'CredentialType',t:'Type',c:'mono'},
    {k:'Description',t:'Description',c:'mono'},
    {k:'Status',t:'Status',r:c=>stBadge(c.Status)},
    {k:'AppState',t:'Enabled',r:c=>enBadge(c.AppState)},
    {k:'EndDate',t:'Expires',c:'mono'},
    {k:'DaysToExpiry',t:'Days',num:1,r:c=>{const d=c.DaysToExpiry; const cl=d<0?'var(--hi)':d<=7?'var(--hi)':d<=14?'var(--med)':'var(--low)'; return `<span style="font-family:var(--mono);color:${cl}">${d<0?Math.abs(d)+'↓':d}</span>`;}},
    {k:'Owner',t:'Owner',c:'mono'},
  ],
  stale:[
    {k:'DisplayName',t:'Application',c:'app'},
    {k:'StaleStatus',t:'Activity',r:a=>stBadge(a.StaleStatus)},
    {k:'Status',t:'Enabled',r:a=>enBadge(a.Status)},
    {k:'DaysSince',t:'Days Since',num:1,r:a=>a.DaysSince==null?'<span class="dim">—</span>':a.DaysSince},
    {k:'LastSignIn',t:'Last Sign-In',c:'mono',r:a=>a.LastSignIn||'<span class="dim">—</span>'},
    {k:'SignInFlow',t:'Last Flow',c:'mono',r:a=>a.SignInFlow||'<span class="dim">—</span>'},
    {k:'TotalCreds',t:'Creds',num:1},
    {k:'Owner',t:'Owner',c:'mono'},
  ],
};

function rows(){
  let r;
  if(state.tab==='creds') r = D.creds.slice();
  else if(state.tab==='enterprise') r = (D.entApps||[]).slice();
  else if(state.tab==='stale') r = D.apps.filter(a=>['Stale','Never used','No service principal'].includes(a.StaleStatus));
  else if(state.tab==='workload') r = D.apps.filter(a=>a.WorkloadIdCand&&(a.WorkloadIdCand.indexOf('CA')===0||a.WorkloadIdCand.indexOf('ID Protection')===0));
  else r = D.apps.filter(a=>a.OverallRisk!=='None');
  // product filter (app tabs only — creds carry no product)
  if(state.tab!=='creds' && state.product!=='all') r=r.filter(x=>prods(x).includes(state.product));
  // enabled/disabled filter (all tabs — creds carry the owning app's state)
  if(state.enabled!=='all') r=r.filter(x=>(state.tab==='creds'?x.AppState:x.Status)===state.enabled);
  // search — matches app name plus its details (owner, tenancy, products, permissions)
  if(state.q){
    r=r.filter(x=>{
      const hay = state.tab==='creds'
        ? [x.AppName,x.Source,x.CredentialType,x.Description,x.Status,x.AppState,x.Owner]
        : [x.DisplayName,x.Owner,x.Tenancy,x.Category,x.WorkloadIdCand,x.SPType,x.StaleStatus,x.SignInFlow,x.Status,x.Sensitive,x.PermSearch,prods(x).join(' ')];
      return hay.join(' ').toLowerCase().includes(state.q);
    });
  }
  // risk filter (risk + workload + enterprise tabs)
  if((state.tab==='risk'||state.tab==='workload'||state.tab==='enterprise') && state.risk!=='all') r=r.filter(x=>x.OverallRisk===state.risk);
  // sort
  const cols=COLS[state.tab];
  if(state.sort.key){
    const cd=cols.find(c=>c.k===state.sort.key);
    r.sort((a,b)=>{
      let va = cd&&cd.sk?cd.sk(a):a[state.sort.key], vb = cd&&cd.sk?cd.sk(b):b[state.sort.key];
      if(typeof va==='string'){va=va.toLowerCase();vb=(vb||'').toLowerCase(); return va<vb?-state.sort.dir:va>vb?state.sort.dir:0;}
      return ((va||0)-(vb||0))*state.sort.dir;
    });
  } else if(state.tab==='risk'||state.tab==='workload'||state.tab==='enterprise'){
    r.sort((a,b)=>riskRank(b.OverallRisk)-riskRank(a.OverallRisk)||b.HighCount-a.HighCount);
  } else if(state.tab==='creds'){ r.sort((a,b)=>a.DaysToExpiry-b.DaysToExpiry); }
  else { r.sort((a,b)=>(b.DaysSince||0)-(a.DaysSince||0)); }
  return r;
}

function render(){
  const cols=COLS[state.tab], data=rows();
  const expandable = state.tab==='risk'||state.tab==='workload'||state.tab==='enterprise';
  let h=`<table><thead><tr>`;
  cols.forEach(c=>{ const on=state.sort.key===c.k; h+=`<th class="${c.num?'num':''} ${on?'sorted':''}" data-k="${c.k}">${c.t}<span class="ar">${on?(state.sort.dir>0?'▲':'▼'):'⇅'}</span></th>`; });
  h+=`<th></th></tr></thead><tbody>`;
  if(!data.length){ h+=`</tbody></table><div class="empty">no matching records</div>`; document.getElementById('table').innerHTML=h.replace('<tbody></tbody>',''); document.getElementById('count').textContent='0 records'; return; }
  data.forEach((x,i)=>{
    h+=`<tr class="row ${expandable?'exp':''}" data-i="${i}" style="animation-delay:${Math.min(i*14,400)}ms">`;
    cols.forEach(c=>{ const val=c.r?c.r(x):(x[c.k]==null||x[c.k]===''?'<span class="dim">—</span>':E(x[c.k])); h+=`<td class="${c.c||''} ${c.num?'num':''}">${val}</td>`; });
    const portal = x.Portal?`<a class="link" href="${E(x.Portal)}" target="_blank" rel="noopener">entra ↗</a>`:'';
    h+=`<td style="text-align:right">${portal}</td></tr>`;
    if(expandable){
      const perms=(x.Sensitive||'').split('\n').filter(Boolean);
      const pl = perms.length? perms.map(p=>{ const cls=/\(Application\)/.test(p)&&/ReadWrite|FullControl|Send|Directory/.test(p)?'p-high':'p-med'; return `<div class="perm ${cls}">${E(p)}</div>`;}).join(''):'<div class="dim" style="font-family:var(--mono);font-size:12px">no high/medium permissions</div>';
      h+=`<tr class="detail" data-d="${i}" style="display:none"><td colspan="${cols.length+1}"><div class="inner"><h4>Sensitive permissions (${perms.length})</h4>${pl}</div></td></tr>`;
    }
  });
  h+=`</tbody></table>`;
  document.getElementById('table').innerHTML=h;
  document.getElementById('count').textContent=`${data.length} record${data.length===1?'':'s'}`;
  // header sort
  document.querySelectorAll('thead th[data-k]').forEach(th=>th.onclick=()=>{
    const k=th.dataset.k;
    if(state.sort.key===k) state.sort.dir*=-1; else state.sort={key:k,dir:1};
    render();
  });
  // row expand
  if(expandable) document.querySelectorAll('tr.row.exp').forEach(tr=>tr.onclick=e=>{
    if(e.target.closest('a'))return;
    const d=document.querySelector(`tr.detail[data-d="${tr.dataset.i}"]`);
    if(d) d.style.display = d.style.display==='none'?'table-row':'none';
  });
}
render();
</script>
</body>
</html>
'@

$htmlOut = $htmlTemplate.Replace('__DATA_JSON__', $json)
[System.IO.File]::WriteAllText($htmlPath, $htmlOut, [System.Text.UTF8Encoding]::new($false))
Write-Host "  HTML report saved." -ForegroundColor Green

$swTotal.Stop()

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  Excel report: $exportPath" -ForegroundColor Green
Write-Host "  HTML report:  $htmlPath" -ForegroundColor Green
Write-Host "  App registrations:    $($analysis.Count)  (HIGH-risk perms: $($highRiskApps.Count))" -ForegroundColor White
Write-Host "  Enterprise apps:      $($entAnalysis.Count)  (HIGH-risk grants: $($entHigh.Count), no app reg: $($entNoAppReg.Count))" -ForegroundColor White
Write-Host "  Expired credentials:  $($expiredCreds.Count) app reg + $($entExpired.Count) enterprise" -ForegroundColor White
Write-Host "  Stale / never used:   $($staleApps.Count) / $($neverUsed.Count)" -ForegroundColor White
if ($actionResults.Count -gt 0) {
  $ok = @($actionResults | Where-Object { $_.Result -eq 'Success' }).Count
  $bad = @($actionResults | Where-Object { $_.Result -ne 'Success' }).Count
  Write-Host "  Actions ($actionMode): $ok succeeded, $bad failed" -ForegroundColor White
}
Write-Host "  Total time: $($swTotal.Elapsed.ToString('mm\:ss'))" -ForegroundColor Gray
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""
