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

Write-Host ""
Write-Host "  SharePoint Sites Audit — Data Export (PnP + Graph, app-only cert auth)" -ForegroundColor Cyan
Write-Host "  Tenant:     $TenantDomain" -ForegroundColor Gray
Write-Host "  SPO admin:  $SpoAdminUrl" -ForegroundColor Gray
Write-Host "  Mode:       $(if ($FullScan) { 'Full scan' } else { "Sample ($SampleSize sites)" })" -ForegroundColor Gray
Write-Host ""

if (-not (Test-Path $ConfigPath)) {
    Write-Host "  [!] Config file not found: $ConfigPath" -ForegroundColor Red
    Write-Host "      Run first-time setup via SPO-Manager.ps1" -ForegroundColor Yellow
    exit 1
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

# Load config: ClientId, CertificatePath, EncryptedCertPassword
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$clientId     = $config.ClientId
$certPath     = $config.CertificatePath
$encryptedPw  = $config.EncryptedCertPassword
$thumbprint   = $config.CertificateThumbprint

if (-not (Test-Path $certPath)) {
    Write-Host "  [!] Certificate not found: $certPath" -ForegroundColor Red
    Write-Host "      Config may be stale. Delete config.json and re-run to redo setup." -ForegroundColor Yellow
    exit 1
}

$securePw = $encryptedPw | ConvertTo-SecureString

# ── Phase 1: Connect to SharePoint via PnP (cert-based, silent) ──
Write-Host "  [1/4] Connecting to SharePoint Online (PnP, cert auth — silent)..." -ForegroundColor Cyan
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline `
        -Url $SpoAdminUrl `
        -ClientId $clientId `
        -Tenant $TenantDomain `
        -CertificatePath $certPath `
        -CertificatePassword $securePw `
        -ErrorAction Stop
    Write-Host "        Connected to $SpoAdminUrl" -ForegroundColor Green
} catch {
    Write-Host "        [!] PnP connection failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -match 'AADSTS7000215|invalid_client|AADSTS50034|unauthorized|AADSTS700016') {
        Write-Host ""
        Write-Host "  [!] Admin consent may not have propagated yet. Wait 2-3 minutes and retry." -ForegroundColor Yellow
        Write-Host "      If the problem persists, delete config.json and re-run setup." -ForegroundColor Yellow
    }
    exit 1
}

# ── Phase 2: Connect to Microsoft Graph (cert-based, same app, silent) ──
Write-Host ""
Write-Host "  [2/4] Connecting to Microsoft Graph (cert auth — silent)..." -ForegroundColor Cyan

# Resolve TenantId (needed for Graph cert auth)
$tenantId = $null
try {
    # Derive from PnP connection: the tenant GUID is available via Get-PnPTenantId or similar
    $tenantId = (Get-PnPTenantId -ErrorAction Stop)
} catch {
    # Fallback: look up via OpenID discovery
    try {
        $disc = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomain/v2.0/.well-known/openid-configuration"
        if ($disc.issuer -match 'sts\.windows\.net/([0-9a-f-]{36})') { $tenantId = $Matches[1] }
    } catch {}
}
if (-not $tenantId) {
    Write-Host "        [!] Could not determine Tenant ID." -ForegroundColor Red
    exit 1
}

try {
    # Use the same PFX cert. Graph module accepts -CertificateThumbprint or -Certificate (X509Certificate2 object).
    Add-Type -AssemblyName System.Security
    $plainPw = [System.Net.NetworkCredential]::new("", $securePw).Password
    $certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $certPath, $plainPw
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $certObj -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Graph context not established." }
    Write-Host "        Connected as app $clientId (app-only)" -ForegroundColor Green
} catch {
    Write-Host "        [!] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# ── Helpers ──
function Invoke-GraphPaged {
    param([string]$Uri)
    $all = @()
    $next = $Uri
    while ($next) {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        } catch {
            Write-Host "        [!] Graph request failed for $next : $($_.Exception.Message)" -ForegroundColor Red
            throw
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
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# Split sites vs OneDrive — PnP Template 'SPSPERS#10' or URL contains -my.sharepoint.com
$nonPersonal = @($allSites | Where-Object { $_.Template -notmatch 'SPSPERS' -and $_.Url -notmatch '-my\.sharepoint\.com' })
$oneDrives   = @($allSites | Where-Object { $_.Template -match 'SPSPERS' -or $_.Url -match '-my\.sharepoint\.com' })

# Sample vs full
if (-not $FullScan) {
    $nonPersonal = @($nonPersonal | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First $SampleSize)
    Write-Host "        Sampled top $($nonPersonal.Count) sites by storage" -ForegroundColor Yellow
} else {
    Write-Host "        Full scan: $($nonPersonal.Count) sites" -ForegroundColor Yellow
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
try {
    $groups = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/groups?`$top=999&`$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes,visibility,resourceProvisioningOptions,assignedLabels,createdDateTime"
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

# Sensitivity label catalog (optional)
$labelCatalog = @()
try {
    $labelCatalog = Invoke-GraphPaged "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?`$top=999"
} catch {}

# ── Save JSON artifacts ──
$ctxExport = @{
    TenantDomain = $TenantDomain
    ExportedAt   = (Get-Date).ToString("o")
    RunBy        = if ($ctx.Account) { $ctx.Account } else { "AppOnly:$clientId" }
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
try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
