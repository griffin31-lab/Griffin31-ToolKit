[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DataDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function Read-Json($path) {
    if (Test-Path $path) { return Get-Content $path -Raw | ConvertFrom-Json }
    return $null
}

$siteFindings     = Read-Json (Join-Path $DataDir "site-findings.json")
$onedriveFindings = Read-Json (Join-Path $DataDir "onedrive-findings.json")
$groupFindings    = Read-Json (Join-Path $DataDir "group-findings.json")
$sites            = Read-Json (Join-Path $DataDir "sites.json")
$onedrives        = Read-Json (Join-Path $DataDir "onedrives.json")
$groups           = Read-Json (Join-Path $DataDir "groups.json")
$teams            = Read-Json (Join-Path $DataDir "teams.json")

# Severity point model
$severityPoints = @{ 'High' = 25; 'Medium' = 12; 'Low' = 5 }

function Get-EntityScore($entity) {
    $score = 100
    foreach ($f in @($entity.Findings)) {
        $deduct = $severityPoints[$f.Severity]
        if ($deduct) { $score -= $deduct }
    }
    return [math]::Max(0, $score)
}

function Get-ScoreBand($score) {
    if     ($score -ge 81) { "Strong" }
    elseif ($score -ge 61) { "Good" }
    elseif ($score -ge 41) { "Fair" }
    else                   { "Critical" }
}

# Flatten all entities that have findings + score them
$entityRecords = @()

foreach ($e in @($siteFindings.Findings)) {
    $entityRecords += [PSCustomObject]@{
        EntityType  = $e.EntityType
        EntityId    = $e.EntityId
        EntityName  = $e.EntityName
        EntityUrl   = $e.EntityUrl
        StorageBytes = $e.StorageBytes
        Score       = (Get-EntityScore $e)
        Findings    = $e.Findings
        Meta        = @{ SharingCap = $e.SharingCap; ExternalCount = $e.ExternalCount }
    }
}

foreach ($e in @($onedriveFindings.Findings)) {
    $entityRecords += [PSCustomObject]@{
        EntityType  = $e.EntityType
        EntityId    = $e.EntityId
        EntityName  = $e.EntityName
        EntityUrl   = $e.EntityUrl
        StorageBytes = $e.StorageBytes
        Score       = (Get-EntityScore $e)
        Findings    = $e.Findings
        Meta        = @{ SharingCap = $e.SharingCap; ExternalCount = $e.ExternalCount }
    }
}

foreach ($e in @($groupFindings.Findings)) {
    $entityRecords += [PSCustomObject]@{
        EntityType  = $e.EntityType
        EntityId    = $e.EntityId
        EntityName  = $e.EntityName
        EntityUrl   = $e.EntityUrl
        StorageBytes = 0
        Score       = (Get-EntityScore $e)
        Findings    = $e.Findings
        Meta        = @{ GuestCount = $e.GuestCount; MemberCount = $e.MemberCount; Visibility = $e.Visibility }
    }
}

# Sort by score ascending (worst first)
$entityRecords = @($entityRecords | Sort-Object -Property @{Expression='Score'; Descending=$false}, @{Expression='EntityName'; Descending=$false})

# ── Tenant-level posture score ──
# Weighted by storage for sites/OneDrive; groups/teams contribute equal weight.
# Step 1: For each scanned entity (not just those with findings), assume 100 if no finding.
# Step 2: Weight sites/OneDrive by storage, groups/teams by count.

$allSiteScores = @{}
foreach ($s in @($sites))     { $allSiteScores[$s.Url] = 100 }
foreach ($o in @($onedrives)) { $allSiteScores[$o.Url] = 100 }
$allGroupScores = @{}
foreach ($g in @($groups))    { $allGroupScores[$g.Id] = 100 }
foreach ($t in @($teams))     { $allGroupScores[$t.Id] = 100 }

foreach ($e in $entityRecords) {
    if ($e.EntityType -in @('Site','OneDrive')) { $allSiteScores[$e.EntityId] = $e.Score }
    else { $allGroupScores[$e.EntityId] = $e.Score }
}

# Storage-weighted site score
$totalStorage = 0
$weightedSum  = 0
foreach ($s in @($sites + $onedrives)) {
    $w = [double]$s.StorageUsageCurrent
    if ($w -le 0) { $w = 1 }   # avoid zero-weight
    $score = $allSiteScores[$s.Url]
    if ($null -eq $score) { $score = 100 }
    $totalStorage += $w
    $weightedSum  += ($score * $w)
}
$siteAvg = if ($totalStorage -gt 0) { [math]::Round($weightedSum / $totalStorage, 0) } else { 100 }

# Group/Team average (equal weight)
$grpAvg = if ($allGroupScores.Count -gt 0) {
    [math]::Round(($allGroupScores.Values | Measure-Object -Average).Average, 0)
} else { 100 }

# Tenant score: 70% sites+OneDrive (the bulk of data), 30% groups+teams
$tenantScore = [math]::Round(($siteAvg * 0.7) + ($grpAvg * 0.3), 0)
$tenantBand  = Get-ScoreBand $tenantScore

# Aggregate severity counts
$allFindings = @($entityRecords | ForEach-Object { $_.Findings } | Where-Object { $_ })
$critCount = 0  # Critical severity not used; reserved
$highCount = @($allFindings | Where-Object { $_.Severity -eq 'High' }).Count
$medCount  = @($allFindings | Where-Object { $_.Severity -eq 'Medium' }).Count
$lowCount  = @($allFindings | Where-Object { $_.Severity -eq 'Low' }).Count

$totalEntitiesScanned = (@($sites).Count + @($onedrives).Count + @($groups).Count + @($teams).Count)

$result = @{
    TenantScore        = $tenantScore
    TenantBand         = $tenantBand
    SiteAverageScore   = $siteAvg
    GroupAverageScore  = $grpAvg
    Summary = @{
        TotalEntitiesScanned  = $totalEntitiesScanned
        SitesScanned          = @($sites).Count
        OneDrivesScanned      = @($onedrives).Count
        GroupsScanned         = @($groups).Count
        TeamsScanned          = @($teams).Count
        EntitiesWithFindings  = $entityRecords.Count
        HighFindings          = $highCount
        MediumFindings        = $medCount
        LowFindings           = $lowCount
    }
    Entities = $entityRecords
    Timestamp = (Get-Date).ToString("o")
}

$result | ConvertTo-Json -Depth 8 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "  Key insights complete. Tenant posture: $tenantScore ($tenantBand). Entities with findings: $($entityRecords.Count)" -ForegroundColor Green
