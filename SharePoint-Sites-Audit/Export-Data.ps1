[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputFolder,
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [Parameter(Mandatory)]
    [string]$SpoAdminUrl,
    [switch]$FullScan,
    [int]$SampleSize = 100,
    [int]$ExternalUserSampleSize = 25   # cap per-site Get-SPOUser enumeration to this many sites
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  SharePoint Sites Audit — Data Export" -ForegroundColor Cyan
Write-Host "  Tenant:     $TenantDomain" -ForegroundColor Gray
Write-Host "  SPO admin:  $SpoAdminUrl" -ForegroundColor Gray
Write-Host "  Mode:       $(if ($FullScan) { 'Full scan' } else { "Sample ($SampleSize sites)" })" -ForegroundColor Gray
Write-Host ""

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

# ── Phase 1: Connect to SharePoint Online ──
Write-Host "  [1/4] Connecting to SharePoint Online..." -ForegroundColor Cyan
Write-Host "        A browser window will open for sign-in." -ForegroundColor DarkGray
try {
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
    Connect-SPOService -Url $SpoAdminUrl -ErrorAction Stop
    Write-Host "        Connected to $SpoAdminUrl" -ForegroundColor Green
} catch {
    Write-Host "        [!] SPO connection failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "        Check that the admin URL is correct and you have SharePoint Administrator role." -ForegroundColor Yellow
    exit 1
}

# ── Phase 2: Connect to Microsoft Graph ──
Write-Host ""
Write-Host "  [2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$graphScopes = @(
    "Sites.Read.All",
    "Group.Read.All",
    "Directory.Read.All",
    "InformationProtectionPolicy.Read"
)
try {
    Connect-MgGraph -Scopes $graphScopes -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Graph context not established." }
    Write-Host "        Connected as $($ctx.Account)" -ForegroundColor Green
} catch {
    Write-Host "        [!] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# ── Helper: Graph pagination ──
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
Write-Host "  [3/4] Fetching SharePoint sites..." -ForegroundColor Cyan
try {
    # Include personal (OneDrive) sites for separate analysis
    $allSites = Get-SPOSite -Limit All -IncludePersonalSite:$true -Detailed
    Write-Host "        Found $($allSites.Count) total sites (including OneDrive)" -ForegroundColor Green
} catch {
    Write-Host "        [!] Get-SPOSite failed: $($_.Exception.Message)" -ForegroundColor Red
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# Split sites vs OneDrive (personal)
$nonPersonal = @($allSites | Where-Object { $_.Template -notmatch 'SPSPERS' -and $_.Url -notmatch '-my\.sharepoint\.com' })
$oneDrives   = @($allSites | Where-Object { $_.Template -match 'SPSPERS'  -or  $_.Url -match '-my\.sharepoint\.com' })

# Sample vs full
if (-not $FullScan) {
    $nonPersonal = @($nonPersonal | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First $SampleSize)
    Write-Host "        Sampled top $($nonPersonal.Count) sites by storage" -ForegroundColor Yellow
} else {
    Write-Host "        Full scan: $($nonPersonal.Count) sites" -ForegroundColor Yellow
}

# Build site records
$siteRecords = @()
$siteCount = $nonPersonal.Count
$idx = 0
foreach ($s in $nonPersonal) {
    $idx++
    if (($idx % 10) -eq 0 -or $idx -eq $siteCount) {
        Write-ProgressBar -Current $idx -Total $siteCount -Activity "Enumerating sites" -Status $s.Title
    }
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
    }
}
Write-Progress -Activity "Enumerating sites" -Completed

# ── Per-site external user counts (slow — sample only unless FullScan) ──
$externalUserMap = @{}
$sitesToProbe = if ($FullScan) { $siteRecords } else {
    @($siteRecords | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First $ExternalUserSampleSize)
}
Write-Host ""
Write-Host "        Fetching external user counts for $($sitesToProbe.Count) site(s)..." -ForegroundColor Gray
$idx = 0
foreach ($s in $sitesToProbe) {
    $idx++
    if (($idx % 5) -eq 0 -or $idx -eq $sitesToProbe.Count) {
        Write-ProgressBar -Current $idx -Total $sitesToProbe.Count -Activity "External users per site" -Status $s.Title
    }
    try {
        $ext = @(Get-SPOUser -Site $s.Url -Limit All -ErrorAction Stop | Where-Object { $_.IsExternalUser })
        $externalUserMap[$s.Url] = $ext.Count
    } catch {
        # Access denied or site unreachable — skip
        $externalUserMap[$s.Url] = $null
    }
}
Write-Progress -Activity "External users per site" -Completed

foreach ($s in $siteRecords) {
    if ($externalUserMap.ContainsKey($s.Url)) {
        $s | Add-Member -NotePropertyName ExternalUserCount -NotePropertyValue $externalUserMap[$s.Url] -Force
    } else {
        $s | Add-Member -NotePropertyName ExternalUserCount -NotePropertyValue $null -Force
    }
}

# Build OneDrive records
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
    }
}

# Fetch external user counts for OneDrive too (smaller sample)
$odProbe = if ($FullScan) { $onedriveRecords } else {
    @($onedriveRecords | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First $ExternalUserSampleSize)
}
Write-Host ""
Write-Host "        Fetching OneDrive external user counts for $($odProbe.Count) account(s)..." -ForegroundColor Gray
$idx = 0
foreach ($od in $odProbe) {
    $idx++
    if (($idx % 5) -eq 0 -or $idx -eq $odProbe.Count) {
        Write-ProgressBar -Current $idx -Total $odProbe.Count -Activity "OneDrive external users" -Status $od.Owner
    }
    try {
        $ext = @(Get-SPOUser -Site $od.Url -Limit All -ErrorAction Stop | Where-Object { $_.IsExternalUser })
        $od | Add-Member -NotePropertyName ExternalUserCount -NotePropertyValue $ext.Count -Force
    } catch {
        $od | Add-Member -NotePropertyName ExternalUserCount -NotePropertyValue $null -Force
    }
}
Write-Progress -Activity "OneDrive external users" -Completed

foreach ($od in $onedriveRecords) {
    if (-not ($od.PSObject.Properties['ExternalUserCount'])) {
        $od | Add-Member -NotePropertyName ExternalUserCount -NotePropertyValue $null -Force
    }
}

# Tenant baseline (for comparing per-site sharing to tenant default)
try {
    $tenantCfg = Get-SPOTenant
    $tenantBaseline = [PSCustomObject]@{
        SharingCapability       = [string]$tenantCfg.SharingCapability
        DefaultSharingLinkType  = [string]$tenantCfg.DefaultSharingLinkType
        DefaultLinkPermission   = [string]$tenantCfg.DefaultLinkPermission
    }
} catch {
    $tenantBaseline = [PSCustomObject]@{
        SharingCapability = "Unknown"; DefaultSharingLinkType = "Unknown"; DefaultLinkPermission = "Unknown"
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

# Classify: M365 group if groupTypes contains 'Unified'; Team if resourceProvisioningOptions contains 'Team'
$groupRecords = @()
$teamRecords  = @()
foreach ($g in $groups) {
    $isUnified = @($g.groupTypes) -contains 'Unified'
    $isTeam    = @($g.resourceProvisioningOptions) -contains 'Team'
    $labelIds  = @()
    if ($g.assignedLabels) {
        $labelIds = @($g.assignedLabels | ForEach-Object { $_.labelId })
    }
    $rec = [PSCustomObject]@{
        Id                = [string]$g.id
        DisplayName       = [string]$g.displayName
        Mail              = [string]$g.mail
        Visibility        = [string]$g.visibility
        Created           = [string]$g.createdDateTime
        AssignedLabelIds  = $labelIds
        HasSensitivityLabel = ($labelIds.Count -gt 0)
        GuestCount        = 0   # populated below
        MemberCount       = 0
        IsTeam            = $isTeam
    }
    if ($isTeam) {
        $teamRecords += $rec
    } elseif ($isUnified) {
        $groupRecords += $rec
    }
}

# For each M365 group + team, fetch guest member count (only if group has a sensitivity label gap — probe everyone in sample mode, or all in full)
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
        # Skip errors silently
    }
}
Write-Progress -Activity "Group members" -Completed

# ── Fetch sensitivity label catalog (for name lookup in the report) ──
$labelCatalog = @()
try {
    $labelCatalog = Invoke-GraphPaged "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?`$top=999"
} catch {
    # Optional — tenant may not have labels configured
}

# ── Save JSON artifacts ──
$ctxExport = @{
    TenantDomain = $TenantDomain
    ExportedAt   = (Get-Date).ToString("o")
    RunBy        = $ctx.Account
    Mode         = if ($FullScan) { "FullScan" } else { "Sample-$SampleSize" }
}

$siteRecords    | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "sites.json")          -Encoding UTF8
$onedriveRecords| ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "onedrives.json")      -Encoding UTF8
$groupRecords   | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "groups.json")         -Encoding UTF8
$teamRecords    | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $OutputFolder "teams.json")          -Encoding UTF8
$tenantBaseline | ConvertTo-Json -Depth 3 | Out-File -FilePath (Join-Path $OutputFolder "tenant-baseline.json") -Encoding UTF8
$labelCatalog   | ConvertTo-Json -Depth 4 | Out-File -FilePath (Join-Path $OutputFolder "sensitivity-labels.json") -Encoding UTF8
$ctxExport      | ConvertTo-Json -Depth 2 | Out-File -FilePath (Join-Path $OutputFolder "export-context.json")  -Encoding UTF8

Write-Host ""
Write-Host "  Export complete. Data saved in: $OutputFolder" -ForegroundColor Green
Write-Host "    sites.json           : $($siteRecords.Count) sites" -ForegroundColor Gray
Write-Host "    onedrives.json       : $($onedriveRecords.Count) OneDrives" -ForegroundColor Gray
Write-Host "    groups.json          : $($groupRecords.Count) M365 groups" -ForegroundColor Gray
Write-Host "    teams.json           : $($teamRecords.Count) teams" -ForegroundColor Gray
Write-Host "    tenant-baseline.json : tenant default sharing baseline" -ForegroundColor Gray
Write-Host "    sensitivity-labels.json : $($labelCatalog.Count) label(s)" -ForegroundColor Gray

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
