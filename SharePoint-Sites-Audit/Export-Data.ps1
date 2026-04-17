[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputFolder,
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [Parameter(Mandatory)]
    [string]$SpoAdminUrl,
    [Parameter(Mandatory)]
    [string]$ConfigPath,
    [switch]$FullScan,
    [int]$SampleSize = 100
)

$ErrorActionPreference = "Stop"

# PS7 guard — this script is launched as a subprocess, so the parent's version check doesn't apply here.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "  [!] This script requires PowerShell 7 or later. Current: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "      Install from https://aka.ms/install-powershell and relaunch via 'pwsh'." -ForegroundColor DarkGray
    exit 1
}

Write-Host ""
Write-Host "  SharePoint Sites Audit — Data Export (PnP + Graph, app-only cert auth)" -ForegroundColor Cyan
Write-Host "  Tenant:     $TenantDomain" -ForegroundColor Gray
Write-Host "  SPO admin:  $SpoAdminUrl" -ForegroundColor Gray
Write-Host "  Mode:       $(if ($FullScan) { 'Full scan' } else { "Sample ($SampleSize sites)" })" -ForegroundColor Gray
Write-Host ""

if (-not (Test-Path $ConfigPath)) {
    Write-Host "  [!] Config file not found: $ConfigPath" -ForegroundColor Red
    Write-Host "      Run first-time setup via SPO-Manager.ps1" -ForegroundColor DarkGray
    exit 1
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

# Load config: ClientId, CertificatePath, EncryptedCertPassword
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$clientId     = $config.ClientId
$certPath     = $config.CertificatePath
$encryptedPw  = $config.EncryptedCertPassword
$thumbprint   = $config.CertificateThumbprint

# Detect "fresh setup" — if the app was registered within the last 2 minutes, 700016 is
# almost certainly propagation delay. Otherwise it's a deleted/stale app → fast-fail.
$isFreshSetup = $false
if ($config.RegisteredAt) {
    try {
        $registeredAt = [datetime]::Parse($config.RegisteredAt, [System.Globalization.CultureInfo]::InvariantCulture)
        $ageSeconds = ((Get-Date).ToUniversalTime() - $registeredAt.ToUniversalTime()).TotalSeconds
        if ($ageSeconds -lt 120) { $isFreshSetup = $true }
    } catch {}
}

if (-not (Test-Path $certPath)) {
    Write-Host "  [!] Certificate not found: $certPath" -ForegroundColor Red
    Write-Host "      Config may be stale. Delete config.json and re-run to redo setup." -ForegroundColor Yellow
    exit 1
}

$securePw = $encryptedPw | ConvertTo-SecureString

# ── Phase 1: Connect to SharePoint via PnP (cert-based, silent, with propagation retry) ──
Write-Host "  [1/4] Connecting to SharePoint Online (PnP, cert auth — silent)..." -ForegroundColor Cyan
Import-Module PnP.PowerShell -ErrorAction Stop

function Connect-PnPWithRetry {
    # 700016 ("app not found") handling:
    #   - If the app was registered in the last 2 minutes → treat as propagation, retry briefly
    #   - Otherwise → fast-fail so the outer handler wipes config and triggers re-setup
    $maxAttempts = 5
    $delays      = @(0, 20, 30, 45, 60)

    for ($i = 0; $i -lt $maxAttempts; $i++) {
        if ($delays[$i] -gt 0) {
            Write-Host "        Retrying in $($delays[$i])s (attempt $($i+1)/$maxAttempts)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds $delays[$i]
        }
        try {
            Connect-PnPOnline `
                -Url $SpoAdminUrl `
                -ClientId $clientId `
                -Tenant $TenantDomain `
                -CertificatePath $certPath `
                -CertificatePassword $securePw `
                -ErrorAction Stop
            return $true
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'AADSTS700016|was not found in the directory') {
                if (-not $isFreshSetup) { throw }  # app is genuinely gone — recover immediately
                if ($i -ge 2) { throw }            # fresh setup but still failing after ~50s → give up
                Write-Host "        (new app propagation pending — retrying)" -ForegroundColor DarkGray
                continue
            }
            if ($msg -match 'AADSTS700027|AADSTS7000215|AADSTS50034|AADSTS500011|AADSTS50105|unauthorized|not registered') {
                if ($i -eq $maxAttempts - 1) { throw }
                Write-Host "        (propagation pending — retrying)" -ForegroundColor DarkGray
                continue
            }
            throw
        }
    }
}

try {
    Connect-PnPWithRetry | Out-Null
    Write-Host "        Connected to $SpoAdminUrl" -ForegroundColor Green
} catch {
    $errMsg = $_.Exception.Message

    # AADSTS700016 = app not found. Handled by the outer Manager via exit code 2.
    # Skip the raw Azure error and show a clean recovery message.
    if ($errMsg -match 'AADSTS700016|was not found in the directory') {
        Write-Host "        Configured Entra app no longer exists — preparing to re-register." -ForegroundColor Yellow
        try {
            $configDir = Split-Path $ConfigPath -Parent
            # Clean up older stale backups (keep only the one we're about to create)
            Get-ChildItem -Path $configDir -Filter "config.json.bak-*" -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            $backupPath = "$ConfigPath.bak-$(Get-Date -Format 'yyyyMMddHHmmss')"
            if (Test-Path $ConfigPath) { Move-Item -Path $ConfigPath -Destination $backupPath -Force | Out-Null }
            $certFolder = Join-Path $configDir "cert"
            if (Test-Path $certFolder) { Remove-Item -Path $certFolder -Recurse -Force -ErrorAction SilentlyContinue }
        } catch {
            Write-Host "        [!] Could not clean up stale config: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "            Manually delete $ConfigPath and re-run." -ForegroundColor Yellow
        }
        exit 2
    }

    # Any other connection failure — show the raw error so the admin can diagnose
    Write-Host "        [!] PnP connection failed: $errMsg" -ForegroundColor Red
    if ($errMsg -match 'AADSTS7000215|invalid_client|AADSTS50034|unauthorized|consent|propagat') {
        Write-Host ""
        Write-Host "  [!] Admin consent may not have propagated yet. Wait 2-3 minutes and retry." -ForegroundColor Yellow
    }
    exit 1
}

# ── Phase 2: Obtain a Graph token via PnP (avoids MSAL assembly conflict) ──
# Microsoft.Graph module and PnP.PowerShell ship incompatible versions of
# Microsoft.Identity.Client.dll; loading both in the same process throws
# "Could not load type 'IMsalSFHttpClientFactory'". We skip Connect-MgGraph
# and call Graph via Invoke-RestMethod using a token minted by PnP.
Write-Host ""
Write-Host "  [2/4] Obtaining Graph token via PnP (cert auth — silent)..." -ForegroundColor Cyan

function Get-GraphTokenWithRetry {
    $maxAttempts = 5
    $delays = @(0, 20, 30, 45, 60)
    for ($i = 0; $i -lt $maxAttempts; $i++) {
        if ($delays[$i] -gt 0) {
            Write-Host "        Retrying in $($delays[$i])s (attempt $($i+1)/$maxAttempts)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds $delays[$i]
        }
        try {
            $t = [string](Get-PnPAccessToken -ResourceTypeName Graph -ErrorAction Stop)
            if (-not $t -or $t -notmatch '^[A-Za-z0-9._-]{100,}$') {
                throw "Invalid Graph token format."
            }
            return $t
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'AADSTS700027|AADSTS7000215|AADSTS50034|AADSTS500011|AADSTS50105|not registered|invalid_client') {
                if ($i -eq $maxAttempts - 1) { throw }
                Write-Host "        (cert/consent propagation pending — retrying)" -ForegroundColor DarkGray
                continue
            }
            throw
        }
    }
}

try {
    $script:graphToken = Get-GraphTokenWithRetry
    Write-Host "        Graph token acquired (length $($script:graphToken.Length))" -ForegroundColor Green
} catch {
    Write-Host "        [!] Failed to obtain Graph token after retries: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "            Cert propagation to Graph can occasionally exceed 3 minutes. Wait and re-run the same menu option." -ForegroundColor Yellow
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# ── Helpers ──
function Invoke-GraphPaged {
    param(
        [string]$Uri,
        [hashtable]$ExtraHeaders = @{},
        [switch]$SilentFail       # if set, return @() on error instead of throwing
    )
    $all = @()
    $next = $Uri
    $headers = @{ Authorization = "Bearer $script:graphToken" }
    foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] }

    while ($next) {
        $attempt = 0
        $maxAttempts = 4
        $backoff = @(0, 30, 45, 60)
        $succeeded = $false
        while ($attempt -lt $maxAttempts -and -not $succeeded) {
            if ($backoff[$attempt] -gt 0) {
                Write-Host "        (permission propagation — retrying in $($backoff[$attempt])s)" -ForegroundColor DarkGray
                Start-Sleep -Seconds $backoff[$attempt]
            }
            try {
                $resp = Invoke-RestMethod -Method GET -Uri $next -Headers $headers -ErrorAction Stop
                $succeeded = $true
            } catch {
                $statusCode = 0
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                $isRetryable = ($statusCode -in 401, 403, 429, 500, 502, 503, 504)
                if (-not $isRetryable -or $attempt -ge $maxAttempts - 1) {
                    if ($SilentFail) { return @() }
                    Write-Host "        [!] Graph request failed for $next : $($_.Exception.Message)" -ForegroundColor Red
                    throw
                }
                $attempt++
            }
        }
        if ($resp.value) { $all += $resp.value }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}

function Write-ProgressBar {
    param([int]$Current, [int]$Total, [string]$Activity = "Processing", [string]$Status = "")
    if ($Total -le 0) { return }
    $pct = [math]::Min(100, [math]::Round(($Current / $Total) * 100))
    Write-Progress -Activity $Activity -Status "$pct%  ($Current/$Total)  $Status" -PercentComplete $pct
}

# ── Phase 3: Fetch sites ──
Write-Host ""
Write-Host "  [3/4] Fetching SharePoint sites (Get-PnPTenantSite)..." -ForegroundColor Cyan
try {
    # PnP: Get-PnPTenantSite -Detailed gives sharing capability, storage, dates. IncludeOneDriveSites pulls personal.
    $allSites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites -ErrorAction Stop
    Write-Host "        Found $($allSites.Count) total sites (including OneDrive)" -ForegroundColor Green
} catch {
    Write-Host "        [!] Get-PnPTenantSite failed: $($_.Exception.Message)" -ForegroundColor Red
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    
    exit 1
}

# Split sites vs OneDrive — PnP Template 'SPSPERS#10' or URL contains -my.sharepoint.com
$nonPersonal = @($allSites | Where-Object { $_.Template -notmatch 'SPSPERS' -and $_.Url -notmatch '-my\.sharepoint\.com' })
$oneDrives   = @($allSites | Where-Object { $_.Template -match 'SPSPERS' -or $_.Url -match '-my\.sharepoint\.com' })

# Sample vs full
if (-not $FullScan) {
    $nonPersonal = @($nonPersonal | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First $SampleSize)
    Write-Host "        Sampled top $($nonPersonal.Count) sites by storage" -ForegroundColor Cyan
} else {
    Write-Host "        Full scan: $($nonPersonal.Count) sites" -ForegroundColor Cyan
}

# ── Aggregate external user counts via Get-PnPExternalUser (tenant-wide, fast) ──
Write-Host "        Fetching external user list (tenant-wide)..." -ForegroundColor Gray
$externalUserMap = @{}
try {
    $extUsers = @()
    $pageIdx = 0
    $pageSize = 50
    do {
        $batch = @(Get-PnPExternalUser -PageSize $pageSize -Position ($pageIdx * $pageSize) -ErrorAction Stop)
        $extUsers += $batch
        $pageIdx++
    } while ($batch.Count -eq $pageSize)
    Write-Host "        Found $($extUsers.Count) external users tenant-wide" -ForegroundColor Green

    # Bucket by AcceptedAs site URL where available (ExternalUser object has an InvitedAs site URL)
    # Note: Get-PnPExternalUser does not reliably map each user to their host site. Count per known site via sharing URL match.
    # For v1 we record tenant-wide total and per-site = null (marked as 'not measured').
    foreach ($s in ($nonPersonal + $oneDrives)) {
        $externalUserMap[$s.Url] = $null
    }
    # Tenant total exposed separately in tenant-baseline.json below.
    $tenantExternalTotal = $extUsers.Count
} catch {
    Write-Host "        (skip) Get-PnPExternalUser not available or returned an error: $($_.Exception.Message)" -ForegroundColor DarkGray
    $tenantExternalTotal = $null
}

# Build site records
$siteRecords = @()
foreach ($s in $nonPersonal) {
    $siteRecords += [PSCustomObject]@{
        Url                      = $s.Url
        Title                    = $s.Title
        Template                 = $s.Template
        Owner                    = $s.Owner
        SharingCapability        = [string]$s.SharingCapability
        StorageUsageCurrent      = [int64]$s.StorageUsageCurrent
        LastContentModifiedDate  = if ($s.LastContentModifiedDate) { $s.LastContentModifiedDate.ToString("o") } else { $null }
        LockState                = [string]$s.LockState
        SensitivityLabel         = [string]$s.SensitivityLabel
        GroupId                  = [string]$s.GroupId
        IsHubSite                = [bool]$s.IsHubSite
        ExternalUserCount        = $externalUserMap[$s.Url]  # may be $null if not measured
    }
}

$onedriveRecords = @()
foreach ($od in $oneDrives) {
    $onedriveRecords += [PSCustomObject]@{
        Url                     = $od.Url
        Owner                   = $od.Owner
        Title                   = $od.Title
        SharingCapability       = [string]$od.SharingCapability
        StorageUsageCurrent     = [int64]$od.StorageUsageCurrent
        LastContentModifiedDate = if ($od.LastContentModifiedDate) { $od.LastContentModifiedDate.ToString("o") } else { $null }
        LockState               = [string]$od.LockState
        ExternalUserCount       = $externalUserMap[$od.Url]
    }
}

# Tenant baseline
try {
    $tenantCfg = Get-PnPTenant
    $tenantBaseline = [PSCustomObject]@{
        SharingCapability       = [string]$tenantCfg.SharingCapability
        DefaultSharingLinkType  = [string]$tenantCfg.DefaultSharingLinkType
        DefaultLinkPermission   = [string]$tenantCfg.DefaultLinkPermission
        TenantExternalUserTotal = $tenantExternalTotal
    }
} catch {
    $tenantBaseline = [PSCustomObject]@{
        SharingCapability = "Unknown"; DefaultSharingLinkType = "Unknown"; DefaultLinkPermission = "Unknown"; TenantExternalUserTotal = $null
    }
}

# ── Phase 4: Fetch groups + teams via Graph ──
Write-Host ""
Write-Host "  [4/4] Fetching M365 groups and Teams (via Graph)..." -ForegroundColor Cyan

# `assignedLabels` and `resourceProvisioningOptions` are "advanced query" properties —
# they require ConsistencyLevel: eventual header + $count=true. One call does both.
$groupsUri = "https://graph.microsoft.com/v1.0/groups?`$top=999&`$count=true&`$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes,visibility,resourceProvisioningOptions,assignedLabels,createdDateTime"
try {
    $groups = Invoke-GraphPaged -Uri $groupsUri -ExtraHeaders @{ 'ConsistencyLevel' = 'eventual' }
    Write-Host "        Found $($groups.Count) groups" -ForegroundColor Green
} catch {
    Write-Host "        [!] Groups fetch failed: $($_.Exception.Message)" -ForegroundColor Red
    $groups = @()
}

$groupRecords = @()
$teamRecords  = @()
foreach ($g in $groups) {
    $isUnified = @($g.groupTypes) -contains 'Unified'
    $isTeam    = @($g.resourceProvisioningOptions) -contains 'Team'
    $labelIds  = @()
    if ($g.assignedLabels) { $labelIds = @($g.assignedLabels | ForEach-Object { $_.labelId }) }
    $rec = [PSCustomObject]@{
        Id                  = [string]$g.id
        DisplayName         = [string]$g.displayName
        Mail                = [string]$g.mail
        Visibility          = [string]$g.visibility
        Created             = [string]$g.createdDateTime
        AssignedLabelIds    = $labelIds
        HasSensitivityLabel = ($labelIds.Count -gt 0)
        GuestCount          = 0
        MemberCount         = 0
        IsTeam              = $isTeam
    }
    if ($isTeam)        { $teamRecords += $rec }
    elseif ($isUnified) { $groupRecords += $rec }
}

$probeSet = @($groupRecords + $teamRecords)
Write-Host "        Fetching member + guest counts for $($probeSet.Count) group(s)/team(s)..." -ForegroundColor Gray
$idx = 0
foreach ($g in $probeSet) {
    $idx++
    if (($idx % 10) -eq 0 -or $idx -eq $probeSet.Count) {
        Write-ProgressBar -Current $idx -Total $probeSet.Count -Activity "Group members" -Status $g.DisplayName
    }
    try {
        $members = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/groups/$($g.Id)/members?`$top=999&`$select=id,userType"
        $g.MemberCount = $members.Count
        $g.GuestCount  = @($members | Where-Object { $_.userType -eq 'Guest' }).Count
    } catch {
        # Ignore missing-member errors per group
    }
}
Write-Progress -Activity "Group members" -Completed

# Sensitivity label catalog (optional — beta endpoint, some tenants restrict app-only access)
$labelCatalog = @(Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?`$top=999" -SilentFail)
if ($labelCatalog.Count -eq 0) {
    Write-Host "        (sensitivity label catalog unavailable — continuing without label names)" -ForegroundColor DarkGray
}

# ── Save JSON artifacts ──
$ctxExport = @{
    TenantDomain = $TenantDomain
    ExportedAt   = (Get-Date).ToString("o")
    RunBy        = "AppOnly:$clientId"
    Mode         = if ($FullScan) { "FullScan" } else { "Sample-$SampleSize" }
    Module       = "PnP.PowerShell"
}

$siteRecords    | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "sites.json")             -Encoding UTF8
$onedriveRecords| ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "onedrives.json")         -Encoding UTF8
$groupRecords   | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "groups.json")            -Encoding UTF8
$teamRecords    | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "teams.json")             -Encoding UTF8
$tenantBaseline | ConvertTo-Json -Depth 3 | Out-File -FilePath (Join-Path $OutputFolder "tenant-baseline.json")    -Encoding UTF8
$labelCatalog   | ConvertTo-Json -Depth 4 | Out-File -FilePath (Join-Path $OutputFolder "sensitivity-labels.json") -Encoding UTF8
$ctxExport      | ConvertTo-Json -Depth 2 | Out-File -FilePath (Join-Path $OutputFolder "export-context.json")     -Encoding UTF8

Write-Host ""
Write-Host "  Export complete. Data saved in: $OutputFolder" -ForegroundColor Green
Write-Host "    sites.json           : $($siteRecords.Count) sites" -ForegroundColor Gray
Write-Host "    onedrives.json       : $($onedriveRecords.Count) OneDrives" -ForegroundColor Gray
Write-Host "    groups.json          : $($groupRecords.Count) M365 groups" -ForegroundColor Gray
Write-Host "    teams.json           : $($teamRecords.Count) teams" -ForegroundColor Gray
Write-Host "    tenant-baseline.json : tenant default sharing baseline" -ForegroundColor Gray
Write-Host "    sensitivity-labels.json : $($labelCatalog.Count) label(s)" -ForegroundColor Gray

try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}

