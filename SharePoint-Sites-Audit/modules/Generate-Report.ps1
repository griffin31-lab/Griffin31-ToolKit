[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$AnalysisDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function Read-Json($path) {
    if (Test-Path $path) { return Get-Content $path -Raw | ConvertFrom-Json }
    return $null
}

$insights  = Read-Json (Join-Path $AnalysisDir "key-insights.json")
$context   = Read-Json (Join-Path $AnalysisDir "export-context.json")

$tenantName = if ($context -and $context.TenantDomain) { $context.TenantDomain } else { "Unknown" }
$exportedAt = if ($context -and $context.ExportedAt)   { ([datetime]$context.ExportedAt).ToString("yyyy-MM-dd HH:mm") } else { (Get-Date).ToString("yyyy-MM-dd HH:mm") }
$mode       = if ($context -and $context.Mode)        { $context.Mode } else { "Unknown" }

$tenantScore = if ($insights) { $insights.TenantScore } else { 0 }
$tenantBand  = if ($insights) { $insights.TenantBand }  else { "Unknown" }
$siteAvg     = if ($insights) { $insights.SiteAverageScore } else { 0 }
$groupAvg    = if ($insights) { $insights.GroupAverageScore } else { 0 }

$totalEntities    = if ($insights) { $insights.Summary.TotalEntitiesScanned } else { 0 }
$sitesScanned     = if ($insights) { $insights.Summary.SitesScanned } else { 0 }
$onedrivesScanned = if ($insights) { $insights.Summary.OneDrivesScanned } else { 0 }
$groupsScanned    = if ($insights) { $insights.Summary.GroupsScanned } else { 0 }
$teamsScanned     = if ($insights) { $insights.Summary.TeamsScanned } else { 0 }
$entWithFindings  = if ($insights) { $insights.Summary.EntitiesWithFindings } else { 0 }
$highCount        = if ($insights) { $insights.Summary.HighFindings } else { 0 }
$medCount         = if ($insights) { $insights.Summary.MediumFindings } else { 0 }
$lowCount         = if ($insights) { $insights.Summary.LowFindings } else { 0 }

function HtmlEncode($s) {
    if ($null -eq $s) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$s)
}

function BandColor($band) {
    switch ($band) {
        "Strong"   { "#065F46" }
        "Good"     { "#16A34A" }
        "Fair"     { "#D97706" }
        "Critical" { "#B91C1C" }
        default    { "#4472C4" }
    }
}

function ScoreBadge($score) {
    $color = if     ($score -ge 81) { "#065F46" }
             elseif ($score -ge 61) { "#16A34A" }
             elseif ($score -ge 41) { "#D97706" }
             else                   { "#B91C1C" }
    return "<span class='score-badge' style='background:$color'>$score</span>"
}

function SeverityPill($sev) {
    $color = switch ($sev) {
        "High"   { "#B91C1C" }
        "Medium" { "#D97706" }
        "Low"    { "#2563EB" }
        default  { "#6B7280" }
    }
    return "<span class='sev-pill' style='background:$color'>$(HtmlEncode $sev)</span>"
}

function EntityTypeBadge($type) {
    $color = switch ($type) {
        "Site"     { "#4472C4" }
        "OneDrive" { "#0EA5E9" }
        "Group"    { "#7C3AED" }
        "Team"     { "#059669" }
        default    { "#6B7280" }
    }
    return "<span class='type-badge' style='background:$color'>$(HtmlEncode $type)</span>"
}

function FormatBytes($b) {
    if ($null -eq $b -or $b -le 0) { return "—" }
    $kb = [double]$b
    $units = @("B","KB","MB","GB","TB")
    $i = 0
    while ($kb -ge 1024 -and $i -lt $units.Count - 1) { $kb = $kb / 1024; $i++ }
    return ("{0:N1} {1}" -f $kb, $units[$i])
}

$bandColor = BandColor $tenantBand

# ── Full catalog of checks the tool runs (used to compute "healthy checks") ──
$checkCatalog = @(
    [PSCustomObject]@{ Id="SP-001"; Severity="High";   Title="Publicly accessible site (Anyone links enabled)" }
    [PSCustomObject]@{ Id="SP-003"; Severity="High";   Title="Excessive external users on site (> 10)" }
    [PSCustomObject]@{ Id="SP-004"; Severity="Medium"; Title="Site sharing more permissive than tenant baseline" }
    [PSCustomObject]@{ Id="SP-005"; Severity="Medium"; Title="Inactive site (no content changes > 365 days)" }
    [PSCustomObject]@{ Id="SP-006"; Severity="Medium"; Title="Non-group site with no primary admin" }
    [PSCustomObject]@{ Id="SP-007"; Severity="Medium"; Title="Site missing sensitivity label" }
    [PSCustomObject]@{ Id="SP-008"; Severity="Medium"; Title="External users on non-group site (likely direct grants)" }
    [PSCustomObject]@{ Id="OD-009"; Severity="High";   Title="OneDrive with excessive external users (> 5)" }
    [PSCustomObject]@{ Id="OD-010"; Severity="Medium"; Title="OneDrive sharing more permissive than tenant baseline" }
    [PSCustomObject]@{ Id="GT-011"; Severity="Medium"; Title="M365 Group missing sensitivity label" }
    [PSCustomObject]@{ Id="GT-012"; Severity="Medium"; Title="Team missing sensitivity label" }
    [PSCustomObject]@{ Id="GT-013"; Severity="Medium"; Title="Group/Team has guests and no sensitivity label" }
)

# ── Build entity rows + per-type counts + grouped-by-finding view ──
$entityRowsHtml = ""
$detailsMapJson = ""
$idx = 0
$detailsMap = @{}

# Per-filter counts
$typeCounts = @{ 'Site' = 0; 'OneDrive' = 0; 'Group' = 0; 'Team' = 0 }

# Flat list of (finding + its entity) for the grouped-by-finding view
$flatFindings = @()

if ($insights -and $insights.Entities) {
    foreach ($e in $insights.Entities) {
        $idx++
        if ($typeCounts.ContainsKey($e.EntityType)) { $typeCounts[$e.EntityType]++ }

        $findingsList = ""
        foreach ($f in @($e.Findings)) {
            $findingsList += @"
<div class='finding-row' style='border-left-color:$(switch ($f.Severity) { 'High' { '#B91C1C' } 'Medium' { '#D97706' } default { '#2563EB' } })'>
  <div class='finding-head'>$(SeverityPill $f.Severity) <span class='finding-id'>$(HtmlEncode $f.Id)</span> <span class='finding-title'>$(HtmlEncode $f.Title)</span></div>
  <div class='finding-details'><strong>Finding:</strong> $(HtmlEncode $f.Details)</div>
  <div class='finding-rem'><strong>Remediation:</strong> $(HtmlEncode $f.Remediation)</div>
</div>
"@
            $flatFindings += [PSCustomObject]@{
                FindingId   = $f.Id
                Title       = $f.Title
                Severity    = $f.Severity
                Remediation = $f.Remediation
                EntityName  = $e.EntityName
                EntityType  = $e.EntityType
                EntityUrl   = $e.EntityUrl
                AdminUrl    = $e.AdminUrl
                Details     = $f.Details
            }
        }
        $detailsMap["row-$idx"] = $findingsList

        $metaChips = ""
        if ($e.EntityType -in @('Site','OneDrive')) {
            if ($e.StorageBytes) { $metaChips += "<span class='meta-chip'>Storage: $(FormatBytes $e.StorageBytes)</span>" }
            if ($e.Meta -and $e.Meta.SharingCap)     { $metaChips += "<span class='meta-chip'>Sharing: $(HtmlEncode $e.Meta.SharingCap)</span>" }
            if ($e.Meta -and $null -ne $e.Meta.ExternalCount) { $metaChips += "<span class='meta-chip'>External users: $($e.Meta.ExternalCount)</span>" }
        } else {
            if ($e.Meta -and $e.Meta.MemberCount) { $metaChips += "<span class='meta-chip'>Members: $($e.Meta.MemberCount)</span>" }
            if ($e.Meta -and $e.Meta.GuestCount)  { $metaChips += "<span class='meta-chip'>Guests: $($e.Meta.GuestCount)</span>" }
            if ($e.Meta -and $e.Meta.Visibility)  { $metaChips += "<span class='meta-chip'>$(HtmlEncode $e.Meta.Visibility)</span>" }
        }

        $issueCount = @($e.Findings).Count

        # Primary link: admin portal (where admin can actually act). Secondary chip: the site URL.
        $primary = if ($e.AdminUrl) { "<a href='$(HtmlEncode $e.AdminUrl)' target='_blank' rel='noopener'>$(HtmlEncode $e.EntityName)</a>" } else { HtmlEncode $e.EntityName }
        $siteChip = if ($e.EntityUrl -and $e.EntityUrl -ne $e.AdminUrl) {
            "<a class='site-link-chip' href='$(HtmlEncode $e.EntityUrl)' target='_blank' rel='noopener' title='Open site'>open site</a>"
        } else { "" }

        $entityRowsHtml += @"
<tr data-entity-type='$(HtmlEncode $e.EntityType)' data-details-key='row-$idx'>
  <td class='col-type'>$(EntityTypeBadge $e.EntityType)</td>
  <td class='col-name'>$primary $siteChip<div class='meta-row'>$metaChips</div></td>
  <td class='col-score'>$(ScoreBadge $e.Score)</td>
  <td class='col-issues'>$issueCount</td>
  <td class='col-expand'><button class='expand-btn' aria-label='Expand details' type='button'>+</button></td>
</tr>
"@
    }
    $detailsMapJson = ($detailsMap | ConvertTo-Json -Compress)
} else {
    $entityRowsHtml = "<tr><td colspan='5' class='empty'>No entities with findings.</td></tr>"
    $detailsMapJson = "{}"
}

# ── Build "Group by finding type" view ──
$groupedHtml = ""
if ($flatFindings.Count -gt 0) {
    $byFinding = $flatFindings | Group-Object -Property FindingId | Sort-Object -Property @{Expression={
        switch (($_.Group[0]).Severity) { 'High' {0} 'Medium' {1} 'Low' {2} default {3} }
    }}, Name
    foreach ($grp in $byFinding) {
        $first = $grp.Group[0]
        $sevColor = switch ($first.Severity) { 'High' { '#B91C1C' } 'Medium' { '#D97706' } default { '#2563EB' } }
        $entitiesList = ""
        foreach ($item in $grp.Group) {
            $itemLink = if ($item.AdminUrl) { "<a href='$(HtmlEncode $item.AdminUrl)' target='_blank' rel='noopener'>$(HtmlEncode $item.EntityName)</a>" } else { HtmlEncode $item.EntityName }
            $typePill = EntityTypeBadge $item.EntityType
            $siteSecondary = if ($item.EntityUrl -and $item.EntityUrl -ne $item.AdminUrl) {
                "<a class='site-link-chip' href='$(HtmlEncode $item.EntityUrl)' target='_blank' rel='noopener'>open site</a>"
            } else { "" }
            $entitiesList += "<li>$typePill $itemLink $siteSecondary<div class='group-details'>$(HtmlEncode $item.Details)</div></li>"
        }
        $groupedHtml += @"
<div class='grouped-finding' style='border-left-color:$sevColor'>
  <button class='grouped-head' type='button' aria-expanded='false'>
    $(SeverityPill $first.Severity)
    <span class='finding-id'>$(HtmlEncode $first.FindingId)</span>
    <span class='grouped-title'>$(HtmlEncode $first.Title)</span>
    <span class='grouped-count'>$($grp.Count) affected</span>
    <span class='chev' aria-hidden='true'>+</span>
  </button>
  <div class='grouped-body' hidden>
    <div class='grouped-rem'><strong>Remediation:</strong> $(HtmlEncode $first.Remediation)</div>
    <ul class='grouped-entities'>$entitiesList</ul>
  </div>
</div>
"@
    }
} else {
    $groupedHtml = "<p class='empty'>No findings.</p>"
}

$siteCount = $typeCounts['Site']
$onedriveCount = $typeCounts['OneDrive']
$groupCount = $typeCounts['Group']
$teamCount = $typeCounts['Team']
$totalCount = $siteCount + $onedriveCount + $groupCount + $teamCount

# ── Compute "Healthy checks" (catalog items that fired no findings) ──
$firedIds = @($flatFindings | ForEach-Object { $_.FindingId } | Select-Object -Unique)
$healthyChecks = @($checkCatalog | Where-Object { $_.Id -notin $firedIds })
$healthyHtml = ""
if ($healthyChecks.Count -gt 0) {
    foreach ($c in $healthyChecks) {
        $healthyHtml += "<li><span class='hc-id'>$(HtmlEncode $c.Id)</span> <span class='hc-title'>$(HtmlEncode $c.Title)</span></li>"
    }
}
$healthyCount = $healthyChecks.Count
$firedCount   = $firedIds.Count
$totalChecks  = $checkCatalog.Count

# ── Assemble full HTML ──
$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>SharePoint Sites Audit &mdash; $(HtmlEncode $tenantName)</title>
<link rel='preconnect' href='https://fonts.googleapis.com'>
<link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>
<link href='https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap' rel='stylesheet'>
<link href='https://cdn.datatables.net/2.1.8/css/dataTables.dataTables.min.css' rel='stylesheet' integrity='sha384-bK2L0Q3Sn2fB8YZTIjKejpTUJXd8xNtKn4RcIm7KM3EX3orMlOI31hwtNH0iyky/' crossorigin='anonymous' referrerpolicy='no-referrer'>
<script src='https://code.jquery.com/jquery-3.7.1.min.js' integrity='sha384-1H217gwSVyLSIfaLxHbE7dRb3v4mYCKbpQvzx0cegeju1MVsGrX5xXxAvs/HgeFs' crossorigin='anonymous' referrerpolicy='no-referrer'></script>
<script src='https://cdn.datatables.net/2.1.8/js/dataTables.min.js' integrity='sha384-MgwUq0TVErv5Lkj/jIAgQpC+iUIqwhwXxJMfrZQVAOhr++1MR02yXH8aXdPc3fk0' crossorigin='anonymous' referrerpolicy='no-referrer'></script>
<style>
  :root {
    --navy: #1B2A4A;
    --accent: #4472C4;
    --bg: #F5F7FB;
    --card: #FFFFFF;
    --text: #0F172A;
    --muted: #475569;
    --border: #E2E8F0;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: 'IBM Plex Sans', system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    font-size: 14px;
    background-image:
      linear-gradient(to right, rgba(27, 42, 74, 0.035) 1px, transparent 1px),
      linear-gradient(to bottom, rgba(27, 42, 74, 0.035) 1px, transparent 1px);
    background-size: 32px 32px;
    background-attachment: fixed;
  }
  .mono, code, .num { font-family: 'IBM Plex Mono', monospace; font-variant-numeric: tabular-nums; }
  h1, h2, h3 { line-height: 1.15; letter-spacing: -0.01em; margin: 0; }
  a { color: var(--accent); text-decoration: none; }
  a:hover { text-decoration: underline; }

  .skip-link { position: absolute; top: -40px; left: 0; background: var(--navy); color: white; padding: 8px 12px; z-index: 100; }
  .skip-link:focus { top: 0; }

  header.topbar {
    background: var(--navy); color: white;
    padding: 14px 24px;
    display: flex; justify-content: space-between; align-items: center;
    border-bottom: 4px solid var(--accent);
  }
  .brand { display: flex; align-items: center; gap: 12px; color: white; text-decoration: none; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.1; }
  .brand-name { font-size: 16px; font-weight: 700; letter-spacing: 0.3px; }
  .brand-sub { font-size: 12px; opacity: 0.8; font-weight: 400; }
  .topbar .meta { font-size: 12px; opacity: 0.85; }
  .topbar .meta span { margin-left: 16px; }

  main { padding: 24px 32px; max-width: 1400px; margin: 0 auto; }
  section { margin-bottom: 36px; scroll-margin-top: 20px; }
  section h2 { font-size: 20px; font-weight: 700; color: var(--navy); margin-bottom: 16px; }

  /* Hero posture */
  .posture-hero {
    background: var(--navy);
    background-image:
      linear-gradient(to right, rgba(255,255,255,0.04) 1px, transparent 1px),
      linear-gradient(to bottom, rgba(255,255,255,0.04) 1px, transparent 1px);
    background-size: 24px 24px;
    color: white;
    border-radius: 4px;
    border: 1px solid rgba(68, 114, 196, 0.4);
    padding: 32px 36px;
    display: grid;
    grid-template-columns: auto 1fr auto;
    gap: 40px;
    align-items: center;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
  }
  .posture-hero::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent 0%, var(--accent) 30%, var(--accent) 70%, transparent 100%);
    opacity: 0.7;
  }
  .posture-hero::after {
    content: 'SHAREPOINT // POSTURE'; position: absolute; top: 10px; right: 16px;
    font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: 0.15em;
    color: var(--accent); opacity: 0.75;
  }
  .posture-score { font-size: 88px; font-weight: 700; line-height: 1; font-family: 'IBM Plex Mono', monospace; }
  .posture-meta { display: flex; flex-direction: column; gap: 8px; }
  .posture-band {
    display: inline-block; padding: 4px 14px; border-radius: 20px;
    font-size: 12px; font-weight: 600; letter-spacing: 0.5px; text-transform: uppercase;
    width: fit-content; color: white; background: $bandColor;
  }
  .posture-desc { font-size: 14px; opacity: 0.85; max-width: 560px; }
  .sub-scores { display: flex; flex-direction: column; gap: 8px; }
  .sub-score-item { display: flex; justify-content: space-between; gap: 16px; font-family: 'IBM Plex Mono', monospace; font-size: 12px; opacity: 0.9; }
  .sub-score-item .val { font-weight: 700; font-size: 22px; font-variant-numeric: tabular-nums; }

  /* KPI row */
  .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }
  .kpi {
    background: white; border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px;
  }
  .kpi-label { font-size: 10px; text-transform: uppercase; color: var(--muted); font-weight: 600; letter-spacing: 0.5px; }
  .kpi-value { font-size: 24px; font-weight: 700; color: var(--navy); font-family: 'IBM Plex Mono', monospace; margin-top: 4px; }
  .kpi.high .kpi-value { color: #B91C1C; }
  .kpi.med  .kpi-value { color: #D97706; }

  /* View toggle (grouped-by-finding vs grouped-by-entity) */
  .view-toggle { display: flex; gap: 2px; border-bottom: 2px solid var(--border); margin-bottom: 20px; }
  .view-btn {
    background: transparent; border: none; padding: 10px 20px;
    font-family: inherit; font-size: 14px; font-weight: 600;
    color: var(--muted); cursor: pointer; border-bottom: 2px solid transparent;
    margin-bottom: -2px; transition: color 150ms, border-color 150ms;
  }
  .view-btn:hover { color: var(--text); }
  .view-btn.active { color: var(--navy); border-bottom-color: var(--accent); }
  .view-panel[hidden] { display: none; }

  /* Filter buttons */
  .filter-bar { display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }
  .filter-btn {
    background: white; border: 1px solid var(--border); border-radius: 6px;
    padding: 8px 14px; font-family: inherit; font-size: 13px; font-weight: 600;
    color: var(--muted); cursor: pointer; transition: all 150ms;
    display: inline-flex; align-items: center; gap: 6px;
  }
  .filter-btn:hover { border-color: var(--accent); color: var(--text); }
  .filter-btn.active { background: var(--navy); color: white; border-color: var(--navy); }
  .chip-count {
    display: inline-block; background: var(--border); color: var(--muted);
    font-size: 10px; padding: 1px 8px; border-radius: 10px; font-weight: 700;
    font-family: 'IBM Plex Mono', monospace;
  }
  .filter-btn.active .chip-count { background: var(--accent); color: white; }

  /* Group-by-finding cards (collapsible) */
  .grouped-finding {
    background: white; border-left: 4px solid #2563EB;
    border: 1px solid var(--border); border-radius: 6px;
    margin-bottom: 10px; overflow: hidden;
  }
  .grouped-head {
    width: 100%; display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
    padding: 14px 18px; background: transparent; border: none;
    font-family: inherit; font-size: 14px; text-align: left;
    cursor: pointer; transition: background 150ms;
  }
  .grouped-head:hover { background: #F8FAFC; }
  .grouped-head[aria-expanded='true'] { background: #F1F5F9; }
  .grouped-title { font-size: 15px; font-weight: 600; color: var(--navy); }
  .grouped-count {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px; color: var(--muted);
    background: var(--bg); padding: 2px 10px; border-radius: 4px;
  }
  .grouped-head .chev {
    margin-left: auto;
    width: 24px; height: 24px; line-height: 20px; text-align: center;
    border: 1px solid var(--border); border-radius: 4px;
    font-size: 16px; font-weight: 700; color: var(--muted);
    transition: all 150ms;
  }
  .grouped-head:hover .chev { border-color: var(--accent); color: var(--accent); }
  .grouped-head[aria-expanded='true'] .chev { background: var(--navy); color: white; border-color: var(--navy); }
  .grouped-body { padding: 0 18px 14px 18px; border-top: 1px solid var(--border); }
  .grouped-body[hidden] { display: none; }
  .grouped-rem { font-size: 13px; margin: 12px 0 8px 0; color: var(--text); }
  .grouped-entities { list-style: none; padding: 0; margin: 0; }
  .grouped-entities li {
    padding: 8px 10px; border-top: 1px solid var(--border);
    display: flex; flex-wrap: wrap; gap: 8px; align-items: baseline;
  }
  .group-details { flex-basis: 100%; font-size: 12px; color: var(--muted); margin-top: 2px; margin-left: 8px; }

  .site-link-chip {
    display: inline-block;
    background: var(--bg); border: 1px solid var(--border);
    color: var(--muted) !important;
    font-size: 10px; padding: 1px 8px; border-radius: 3px;
    font-family: 'IBM Plex Mono', monospace;
    text-decoration: none !important;
  }
  .site-link-chip:hover { border-color: var(--accent); color: var(--accent) !important; }

  .expand-controls { display: flex; gap: 6px; margin-bottom: 10px; }
  .expand-controls button {
    background: white; border: 1px solid var(--border); border-radius: 4px;
    padding: 6px 12px; font-family: inherit; font-size: 12px; font-weight: 600;
    color: var(--muted); cursor: pointer;
  }
  .expand-controls button:hover { border-color: var(--accent); color: var(--accent); }

  /* Healthy checks block — subtle, green, collapsed by default */
  .healthy-block {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    border-radius: 6px;
    margin-top: 24px;
    padding: 0 16px;
  }
  .healthy-block summary {
    display: flex; align-items: center; gap: 10px;
    padding: 12px 0;
    font-size: 14px; font-weight: 600; color: #065F46;
    cursor: pointer;
    list-style: none;
  }
  .healthy-block summary::-webkit-details-marker { display: none; }
  .healthy-block summary::marker { content: ''; }
  .hc-check-icon {
    display: inline-flex; align-items: center; justify-content: center;
    width: 22px; height: 22px; border-radius: 50%;
    background: #065F46; color: white;
    font-size: 13px; font-weight: 700;
  }
  .healthy-block .hc-count {
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px; color: #065F46; opacity: 0.75;
  }
  .hc-intro { font-size: 12px; color: #065F46; margin: 0 0 10px 0; opacity: 0.85; }
  .healthy-list { list-style: none; padding: 0; margin: 0 0 14px 0; }
  .healthy-list li {
    padding: 6px 0;
    border-top: 1px dashed #BBF7D0;
    font-size: 13px;
    display: flex; gap: 10px; align-items: baseline;
  }
  .hc-id {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px; color: #065F46; background: #DCFCE7;
    padding: 1px 8px; border-radius: 3px; flex-shrink: 0;
  }
  .hc-title { color: var(--text); }

  /* Entity table */
  table.entities {
    width: 100%; background: white; border-collapse: collapse;
    border-radius: 8px; overflow: hidden; border: 1px solid var(--border);
  }
  table.entities th {
    background: var(--navy); color: white; text-align: left;
    padding: 10px 12px; font-size: 12px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.3px;
  }
  table.entities td {
    padding: 10px 12px; border-top: 1px solid var(--border);
    font-size: 13px; vertical-align: top;
  }
  table.entities tbody tr:hover { background: #F8FAFC; }

  .col-type { width: 90px; }
  .col-name { min-width: 260px; font-weight: 500; }
  .col-score { width: 70px; text-align: center; }
  .col-issues { width: 80px; text-align: center; font-family: 'IBM Plex Mono', monospace; font-weight: 600; }
  .col-expand { width: 48px; text-align: center; }

  .meta-row { margin-top: 4px; display: flex; gap: 4px; flex-wrap: wrap; }
  .meta-chip {
    background: #F1F5F9; color: var(--muted);
    font-size: 10px; padding: 2px 8px; border-radius: 3px;
    font-family: 'IBM Plex Mono', monospace;
  }

  .expand-btn {
    background: transparent; border: 1px solid var(--border); border-radius: 4px;
    width: 28px; height: 28px; font-size: 16px; font-weight: 700;
    cursor: pointer; color: var(--muted); transition: all 150ms;
    font-family: inherit;
  }
  .expand-btn:hover { border-color: var(--accent); color: var(--accent); }
  .expand-btn.open { background: var(--navy); color: white; border-color: var(--navy); }

  .details-row { background: #FAFBFC; }
  .details-inner { padding: 12px 16px; }

  .finding-row {
    background: white; border-left: 4px solid #2563EB;
    border-radius: 4px; padding: 12px 16px; margin-bottom: 8px;
    border: 1px solid var(--border);
  }
  .finding-head { display: flex; gap: 10px; align-items: center; margin-bottom: 6px; flex-wrap: wrap; }
  .finding-id { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--muted); }
  .finding-title { font-size: 14px; font-weight: 600; color: var(--navy); }
  .finding-details, .finding-rem { font-size: 13px; margin: 3px 0; }

  .sev-pill, .type-badge, .score-badge {
    display: inline-block; color: white; padding: 2px 8px;
    border-radius: 4px; font-size: 11px; font-weight: 600; white-space: nowrap;
  }
  .sev-pill { text-transform: uppercase; letter-spacing: 0.5px; }
  .type-badge { font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; }
  .score-badge { font-family: 'IBM Plex Mono', monospace; min-width: 36px; text-align: center; }

  .empty { color: var(--muted); font-style: italic; padding: 20px; text-align: center; }

  /* DataTables polish */
  .dt-search, .dt-length { margin-bottom: 10px; font-size: 13px; }
  .dt-paging button.dt-paging-button { font-size: 12px !important; }

  @media (max-width: 900px) {
    .posture-hero { grid-template-columns: 1fr; gap: 16px; padding: 20px; }
    .posture-score { font-size: 60px; }
    main { padding: 16px; }
  }
  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }
  }
  @media print { .filter-bar, .skip-link { display: none; } }
</style>
</head>
<body>
<a href='#main' class='skip-link'>Skip to content</a>

<header class='topbar' role='banner'>
  <a href='https://www.griffin31.com' target='_blank' rel='noopener' class='brand' aria-label='Griffin31 website'>
    <span class='brand-text'>
      <span class='brand-name'>Griffin31</span>
      <span class='brand-sub'>SharePoint Sites Audit</span>
    </span>
  </a>
  <div class='meta'>
    <span><strong>Tenant:</strong> $(HtmlEncode $tenantName)</span>
    <span><strong>Mode:</strong> $(HtmlEncode $mode)</span>
    <span><strong>Generated:</strong> $exportedAt</span>
    <span><a href='https://www.griffin31.com' target='_blank' rel='noopener' style='color:#9BB3E0; text-decoration:underline;'>griffin31.com</a></span>
  </div>
</header>

<main id='main' role='main'>

  <section id='overview'>
    <div class='posture-hero'>
      <div class='posture-score'>$tenantScore</div>
      <div class='posture-meta'>
        <span class='posture-band'>$(HtmlEncode $tenantBand)</span>
        <div style='font-size:14px;font-weight:600;'>Tenant SharePoint Posture Score</div>
        <div class='posture-desc'>Storage-weighted score across every scanned site, OneDrive, group, and team. Each entity starts at 100 and loses points for every finding (High -25, Medium -12).</div>
      </div>
      <div class='sub-scores'>
        <div class='sub-score-item'><span>Sites &amp; OneDrive</span><span class='val'>$siteAvg</span></div>
        <div class='sub-score-item'><span>Groups &amp; Teams</span><span class='val'>$groupAvg</span></div>
      </div>
    </div>

    <div class='kpi-row'>
      <div class='kpi'><div class='kpi-label'>Sites Scanned</div><div class='kpi-value'>$sitesScanned</div></div>
      <div class='kpi'><div class='kpi-label'>OneDrives</div><div class='kpi-value'>$onedrivesScanned</div></div>
      <div class='kpi'><div class='kpi-label'>Groups</div><div class='kpi-value'>$groupsScanned</div></div>
      <div class='kpi'><div class='kpi-label'>Teams</div><div class='kpi-value'>$teamsScanned</div></div>
      <div class='kpi'><div class='kpi-label'>With Findings</div><div class='kpi-value'>$entWithFindings</div></div>
      <div class='kpi high'><div class='kpi-label'>High Findings</div><div class='kpi-value'>$highCount</div></div>
      <div class='kpi med'><div class='kpi-label'>Medium Findings</div><div class='kpi-value'>$medCount</div></div>
    </div>
  </section>

  <section id='findings'>
    <h2>Findings &amp; affected entities</h2>

    <div class='view-toggle' role='tablist' aria-label='Report view'>
      <button class='view-btn active' data-view='grouped' aria-pressed='true'>Group by finding</button>
      <button class='view-btn' data-view='entity' aria-pressed='false'>Group by entity</button>
    </div>

    <!-- GROUP BY FINDING VIEW -->
    <div id='view-grouped' class='view-panel'>
      <div class='expand-controls'>
        <button class='expand-all-btn' type='button' data-target='grouped'>Expand all</button>
        <button class='collapse-all-btn' type='button' data-target='grouped'>Collapse all</button>
      </div>
      $groupedHtml

      <details class='healthy-block'>
        <summary>
          <span class='hc-check-icon' aria-hidden='true'>&#10003;</span>
          Healthy checks
          <span class='hc-count'>$healthyCount of $totalChecks</span>
        </summary>
        <p class='hc-intro'>The tool ran these checks and your tenant passed. No findings means no matching entities were detected.</p>
        <ul class='healthy-list'>$healthyHtml</ul>
      </details>
    </div>

    <!-- GROUP BY ENTITY VIEW -->
    <div id='view-entity' class='view-panel' hidden>
      <div class='filter-bar' role='tablist' aria-label='Filter by entity type'>
        <button class='filter-btn active' data-filter='all' aria-pressed='true'>All <span class='chip-count'>$totalCount</span></button>
        <button class='filter-btn' data-filter='Site'>Sites <span class='chip-count'>$siteCount</span></button>
        <button class='filter-btn' data-filter='OneDrive'>OneDrive <span class='chip-count'>$onedriveCount</span></button>
        <button class='filter-btn' data-filter='Group'>Groups <span class='chip-count'>$groupCount</span></button>
        <button class='filter-btn' data-filter='Team'>Teams <span class='chip-count'>$teamCount</span></button>
      </div>
      <table id='entityTable' class='entities'>
        <thead>
          <tr>
            <th>Type</th>
            <th>Name</th>
            <th>Score</th>
            <th>Issues</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          $entityRowsHtml
        </tbody>
      </table>
    </div>
  </section>

</main>

<script>
  var detailsMap = $detailsMapJson;

  `$(document).ready(function() {
    var table = `$('#entityTable').DataTable({
      pageLength: 25,
      order: [[2, 'asc']],
      columnDefs: [
        { targets: [4], orderable: false, searchable: false },
        { targets: [0,2,3], className: 'dt-center' }
      ]
    });

    // Custom filter — match on data-entity-type attribute on the <tr>
    var currentTypeFilter = 'all';
    `$.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
      if (settings.nTable.id !== 'entityTable') return true;
      if (currentTypeFilter === 'all') return true;
      var tr = settings.aoData[dataIndex].nTr;
      return tr && tr.getAttribute('data-entity-type') === currentTypeFilter;
    });

    // Expand / collapse (DataTables child-row API)
    `$('#entityTable tbody').on('click', '.expand-btn', function() {
      var btn = `$(this);
      var tr = btn.closest('tr');
      var row = table.row(tr);
      var key = tr.attr('data-details-key');
      if (row.child.isShown()) {
        row.child.hide();
        btn.removeClass('open').text('+');
      } else {
        var html = '<div class="details-inner">' + (detailsMap[key] || '<em>No details.</em>') + '</div>';
        row.child(html).show();
        btn.addClass('open').text('−');
      }
    });

    // Entity type filter
    `$('.filter-btn').on('click', function() {
      `$('.filter-btn').removeClass('active').attr('aria-pressed','false');
      `$(this).addClass('active').attr('aria-pressed','true');
      currentTypeFilter = `$(this).data('filter');
      table.rows().every(function() { if (this.child.isShown()) { this.child.hide(); } });
      `$('.expand-btn.open').removeClass('open').text('+');
      table.draw();
    });

    // View toggle
    `$('.view-btn').on('click', function() {
      `$('.view-btn').removeClass('active').attr('aria-pressed','false');
      `$(this).addClass('active').attr('aria-pressed','true');
      var view = `$(this).data('view');
      `$('#view-grouped').prop('hidden', view !== 'grouped');
      `$('#view-entity').prop('hidden', view !== 'entity');
      if (view === 'entity') { table.columns.adjust(); }
    });

    // Collapse / expand per-finding card (group-by-finding view)
    `$('.grouped-head').on('click', function() {
      var head = `$(this);
      var body = head.next('.grouped-body');
      var isOpen = head.attr('aria-expanded') === 'true';
      if (isOpen) {
        head.attr('aria-expanded','false');
        body.prop('hidden', true);
        head.find('.chev').text('+');
      } else {
        head.attr('aria-expanded','true');
        body.prop('hidden', false);
        head.find('.chev').text('−');
      }
    });

    `$('.expand-all-btn').on('click', function() {
      `$('.grouped-head').attr('aria-expanded','true');
      `$('.grouped-body').prop('hidden', false);
      `$('.grouped-head .chev').text('−');
    });
    `$('.collapse-all-btn').on('click', function() {
      `$('.grouped-head').attr('aria-expanded','false');
      `$('.grouped-body').prop('hidden', true);
      `$('.grouped-head .chev').text('+');
    });
  });
</script>

</body>
</html>
"@

$html | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "  Report generated: $OutputPath" -ForegroundColor Green
