[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$AnalysisDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

# ── Load analysis outputs ──
function Read-Json($path) {
    if (Test-Path $path) { return Get-Content $path -Raw | ConvertFrom-Json }
    return $null
}

$policies      = Read-Json "$AnalysisDir/ConditionalAccessPolicies.json"
$gaps          = Read-Json "$AnalysisDir/policy-gaps.json"
$missing       = Read-Json "$AnalysisDir/missing-controls.json"
$insights      = Read-Json "$AnalysisDir/key-insights.json"
$breakglass    = Read-Json "$AnalysisDir/breakglass.json"
$nestedGroups  = Read-Json "$AnalysisDir/nested-groups.json"

$tenantName = ""
try {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($ctx) { $tenantName = $ctx.Account -replace '^.*@','' }
} catch {}
if (-not $tenantName) {
    $parentPath = Split-Path (Split-Path $AnalysisDir -Parent) -Leaf
    $tenantName = $parentPath
}

$generated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
$postureScore = if ($insights) { $insights.PostureScore } else { 0 }
$postureBand  = if ($insights) { $insights.PostureBand } else { "Unknown" }

$phaseF = if ($insights -and $insights.PhaseScores) { $insights.PhaseScores.Foundation } else { 0 }
$phaseC = if ($insights -and $insights.PhaseScores) { $insights.PhaseScores.Core } else { 0 }
$phaseA = if ($insights -and $insights.PhaseScores) { $insights.PhaseScores.Advanced } else { 0 }
$phaseAdmin = if ($insights -and $insights.PhaseScores) { $insights.PhaseScores.Admin } else { 0 }

$totalPolicies  = if ($gaps) { $gaps.Summary.TotalPolicies } else { 0 }
$enabledCount   = if ($gaps) { $gaps.Summary.EnabledPolicies } else { 0 }
$reportOnly     = if ($gaps) { $gaps.Summary.ReportOnly } else { 0 }
$disabled       = if ($gaps) { $gaps.Summary.Disabled } else { 0 }
$avgScore       = if ($gaps) { $gaps.Summary.AvgScore } else { 0 }

$critCount   = if ($insights) { $insights.Summary.Critical } else { 0 }
$highCount   = if ($insights) { $insights.Summary.High } else { 0 }
$medCount    = if ($insights) { $insights.Summary.Medium } else { 0 }
$infoCount   = if ($insights) { $insights.Summary.Info } else { 0 }

# ── HTML helpers ──
function HtmlEncode($s) {
    if ($null -eq $s) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$s)
}

function BandColor($band) {
    switch ($band) {
        "Strong"   { return "#065F46" }
        "Good"     { return "#16A34A" }
        "Fair"     { return "#92400E" }
        "Critical" { return "#B91C1C" }
        default    { return "#4472C4" }
    }
}

function SeverityColor($sev) {
    switch ($sev) {
        "Critical" { return "#B91C1C" }
        "High"     { return "#D97706" }
        "Medium"   { return "#CA8A04" }
        "Info"     { return "#2563EB" }
        default    { return "#6B7280" }
    }
}

function SeverityBg($sev) {
    switch ($sev) {
        "Critical" { return "#FEE2E2" }
        "High"     { return "#FFEDD5" }
        "Medium"   { return "#FEF3C7" }
        "Info"     { return "#DBEAFE" }
        default    { return "#F3F4F6" }
    }
}

function LicenseBadge($lic) {
    $color = switch ($lic) {
        "P1"      { "#4472C4" }
        "P2"      { "#7C3AED" }
        "Purview" { "#0891B2" }
        "WID"     { "#059669" }
        default   { "#6B7280" }
    }
    $label = switch ($lic) {
        "P1"      { "Entra ID P1" }
        "P2"      { "Entra ID P2" }
        "Purview" { "Microsoft Purview" }
        "WID"     { "Workload Identities" }
        default   { $lic }
    }
    $title = switch ($lic) {
        "P1"      { "Requires Entra ID P1 license" }
        "P2"      { "Requires Entra ID P2 license" }
        "Purview" { "Requires Microsoft Purview (Insider Risk)" }
        "WID"     { "Requires Workload Identities Premium add-on" }
        default   { "" }
    }
    return "<span class='lic-badge' style='background:$color' title='$(HtmlEncode $title)'>$(HtmlEncode $label)</span>"
}

function PriorityBadge($pri) {
    $color = switch ($pri) {
        "Critical" { "#B91C1C" }
        "High"     { "#D97706" }
        "Medium"   { "#CA8A04" }
        default    { "#6B7280" }
    }
    return "<span class='pri-badge' style='background:$color'>$(HtmlEncode $pri)</span>"
}

function StateBadge($state) {
    $color = switch ($state) {
        "enabled" { "#065F46" }
        "enabledForReportingButNotEnforced" { "#92400E" }
        "disabled" { "#6B7280" }
        default   { "#6B7280" }
    }
    $label = switch ($state) {
        "enabled" { "Enabled" }
        "enabledForReportingButNotEnforced" { "Report-only" }
        "disabled" { "Disabled" }
        default { $state }
    }
    return "<span class='state-badge' style='background:$color'>$(HtmlEncode $label)</span>"
}

function ScoreBadge($score) {
    $color = if     ($score -ge 81) { "#065F46" }
             elseif ($score -ge 61) { "#16A34A" }
             elseif ($score -ge 41) { "#CA8A04" }
             elseif ($score -ge 21) { "#D97706" }
             else                   { "#B91C1C" }
    return "<span class='score-badge' style='background:$color'>$score</span>"
}

$bandColor = BandColor $postureBand

# ── Insights HTML ──
$insightsHtml = ""
if ($insights -and $insights.Insights) {
    foreach ($i in $insights.Insights) {
        $sevColor = SeverityColor $i.Severity
        $sevBg    = SeverityBg    $i.Severity
        $affected = ""
        if ($i.AffectedPolicies -and $i.AffectedPolicies.Count -gt 0) {
            $list = ($i.AffectedPolicies | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ""
            $affected = "<details class='affected'><summary>Affected policies ($($i.AffectedPolicies.Count))</summary><ul>$list</ul></details>"
        }
        $insightsHtml += @"
<div class='insight-card' style='border-left-color:$sevColor;background:$sevBg'>
  <div class='insight-head'>
    <span class='sev-pill' style='background:$sevColor'>$(HtmlEncode $i.Severity)</span>
    <span class='phase-tag'>$(HtmlEncode $i.Phase)</span>
    <span class='insight-id'>$(HtmlEncode $i.Id)</span>
  </div>
  <h3 class='insight-title'>$(HtmlEncode $i.Title)</h3>
  <p class='insight-finding'><strong>Finding:</strong> $(HtmlEncode $i.Finding)</p>
  <p class='insight-rec'><strong>Recommendation:</strong> $(HtmlEncode $i.Recommendation)</p>
  $affected
</div>
"@
    }
} else {
    $insightsHtml = "<p class='empty'>No insights generated.</p>"
}

# ── Missing Controls HTML (grouped by priority, with license badges) ──
$missingHtml = ""
if ($missing -and $missing.MissingControls -and $missing.MissingControls.Count -gt 0) {
    $grouped = $missing.MissingControls | Group-Object -Property Priority
    $order = @("Critical","High","Medium")
    foreach ($pri in $order) {
        $grp = $grouped | Where-Object { $_.Name -eq $pri }
        if (-not $grp) { continue }
        $missingHtml += "<h3 class='mc-group-head'>$(PriorityBadge $pri) $pri priority <span class='mc-count'>($($grp.Count))</span></h3>"
        $missingHtml += "<div class='mc-grid'>"
        foreach ($c in $grp.Group) {
            $missingHtml += @"
<div class='mc-card'>
  <div class='mc-head'>
    <span class='mc-type'>$(HtmlEncode $c.ControlType)</span>
    $(LicenseBadge $c.License)
  </div>
  <div class='mc-name'>$(HtmlEncode $c.ControlName)</div>
</div>
"@
        }
        $missingHtml += "</div>"
    }
} else {
    $missingHtml = "<p class='empty'>All recommended controls are in use.</p>"
}

# ── Policies tables (split by state, same scoring columns) ──
function Build-PolicyRows($list) {
    $rows = ""
    foreach ($p in $list) {
        $flagsHtml = ""
        if ($p.Flags -and $p.Flags.Count -gt 0) {
            foreach ($f in $p.Flags) {
                $flagsHtml += "<span class='flag-pill'>$(HtmlEncode $f)</span>"
            }
        }
        # Validate PolicyId is a GUID before building the portal URL
        $pid = [string]$p.PolicyId
        if ($pid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            $portalUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/policyId/$pid"
            $nameCell = "<a href='$portalUrl' target='_blank' rel='noopener'>$(HtmlEncode $p.PolicyName)</a>"
        } else {
            $nameCell = "$(HtmlEncode $p.PolicyName)"
        }
        $assignCell = Render-GapList $p.AssignmentGaps
        $condCell   = Render-GapList $p.ConditionGaps
        $rows += @"
<tr>
  <td class='col-name'>$nameCell$flagsHtml</td>
  <td class='col-score'>$(ScoreBadge $p.Score)</td>
  <td class='col-assign'>$assignCell</td>
  <td class='col-cond'>$condCell</td>
  <td>$(HtmlEncode $p.Modified)</td>
</tr>
"@
    }
    return $rows
}

function Render-GapList($gaps) {
    if (-not $gaps -or @($gaps).Count -eq 0) { return "<span class='muted'>&mdash;</span>" }
    $arr = @($gaps)
    if ($arr.Count -eq 1) { return HtmlEncode $arr[0] }
    $items = ($arr | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ""
    return "<ul class='gap-list'>$items</ul>"
}

$enabledRows = ""; $reportRows = ""; $disabledRows = ""
$enabledTotal = 0; $reportTotal = 0; $disabledTotal = 0
if ($gaps -and $gaps.PolicyGaps) {
    $enabledList  = @($gaps.PolicyGaps | Where-Object { $_.State -eq 'enabled' })
    $reportList   = @($gaps.PolicyGaps | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' })
    $disabledList = @($gaps.PolicyGaps | Where-Object { $_.State -eq 'disabled' })
    $enabledRows  = Build-PolicyRows $enabledList
    $reportRows   = Build-PolicyRows $reportList
    $disabledRows = Build-PolicyRows $disabledList
    $enabledTotal = $enabledList.Count
    $reportTotal  = $reportList.Count
    $disabledTotal = $disabledList.Count
}

function Build-PolicyTable($id, $rows) {
    if (-not $rows) { return "<p class='empty'>No policies in this state.</p>" }
    return @"
<table id='$id' class='table policies-table'>
  <thead>
    <tr>
      <th>Policy Name</th>
      <th>Score</th>
      <th>Assignment Gaps</th>
      <th>Condition Gaps</th>
      <th>Modified</th>
    </tr>
  </thead>
  <tbody>
$rows
  </tbody>
</table>
"@
}
$tblEnabled  = Build-PolicyTable 'policiesEnabled'  $enabledRows
$tblReport   = Build-PolicyTable 'policiesReport'   $reportRows
$tblDisabled = Build-PolicyTable 'policiesDisabled' $disabledRows

# ── Break-Glass HTML ──
$bgHtml = ""
if ($breakglass -and $breakglass.BreakGlassUsers -and $breakglass.BreakGlassUsers.Count -gt 0) {
    $bgHtml = "<table class='table'><thead><tr><th>Display Name</th><th>UPN</th><th>Detection Method</th></tr></thead><tbody>"
    foreach ($u in $breakglass.BreakGlassUsers) {
        $bgHtml += "<tr><td>$(HtmlEncode $u.DisplayName)</td><td>$(HtmlEncode $u.UserPrincipalName)</td><td>$(HtmlEncode $breakglass.DetectionMethod)</td></tr>"
    }
    $bgHtml += "</tbody></table>"
} else {
    $bgHtml = "<div class='warning-box'><strong>No break-glass accounts detected.</strong> Create at least two dedicated emergency access accounts and exclude them from all CA policies.</div>"
}

# ── Nested Groups — CA blind spots (policies assigned to nested groups) ──
$ngHtml = ""
if ($nestedGroups -and $nestedGroups.PoliciesUsingNestedGroups -and $nestedGroups.PoliciesUsingNestedGroups.Count -gt 0) {
    $ngHtml += "<p class='section-lead'>CA policy assignments that reference a security group containing nested groups. The effective member set can change without editing the policy &mdash; a management blind spot.</p>"
    $ngHtml += "<table class='table'><thead><tr><th>Policy</th><th>State</th><th>Assignment</th><th>Group (contains nested)</th></tr></thead><tbody>"
    foreach ($row in $nestedGroups.PoliciesUsingNestedGroups) {
        $assignBadge = if ($row.Assignment -eq 'Exclude') {
            "<span class='assign-badge exclude'>Exclude</span>"
        } else {
            "<span class='assign-badge include'>Include</span>"
        }
        $pid = [string]$row.PolicyId
        $nameCell = if ($pid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            "<a href='https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/policyId/$pid' target='_blank' rel='noopener'>$(HtmlEncode $row.PolicyName)</a>"
        } else { HtmlEncode $row.PolicyName }
        $ngHtml += "<tr><td>$nameCell</td><td>$(StateBadge $row.PolicyState)</td><td>$assignBadge</td><td>$(HtmlEncode $row.GroupName)</td></tr>"
    }
    $ngHtml += "</tbody></table>"
} else {
    $ngHtml = "<p class='empty'>No CA policies are assigned to groups that contain nested groups.</p>"
}

# Supplementary: list of groups that contain nested groups (reference)
if ($nestedGroups -and $nestedGroups.NestedGroups -and $nestedGroups.NestedGroups.Count -gt 0) {
    $ngHtml += "<h3 class='subhead'>Reference: groups with nested membership</h3>"
    $ngHtml += "<table class='table'><thead><tr><th>Group</th><th>Direct Members</th><th>Total (recursive)</th><th>Nested Children</th></tr></thead><tbody>"
    foreach ($g in $nestedGroups.NestedGroups) {
        $children = @($g.NestedChildren | ForEach-Object { HtmlEncode $_.Name }) -join ", "
        $ngHtml += "<tr><td>$(HtmlEncode $g.GroupName)</td><td>$(HtmlEncode $g.DirectMembers)</td><td>$(HtmlEncode $g.TotalMembers)</td><td>$children</td></tr>"
    }
    $ngHtml += "</tbody></table>"
}

# ── Build full HTML ──
$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Conditional Access Analysis &mdash; $(HtmlEncode $tenantName)</title>
<link rel='preconnect' href='https://fonts.googleapis.com'>
<link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>
<link href='https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap' rel='stylesheet'>
<link href='https://cdn.datatables.net/2.1.8/css/dataTables.dataTables.min.css' rel='stylesheet' integrity='sha384-bK2L0Q3Sn2fB8YZTIjKejpTUJXd8xNtKn4RcIm7KM3EX3orMlOI31hwtNH0iyky/' crossorigin='anonymous' referrerpolicy='no-referrer'>
<script src='https://cdn.jsdelivr.net/npm/lucide@0.453.0/dist/umd/lucide.min.js' defer integrity='sha384-0v3Wh0ZtiQBYCYywQoeMDA5mAw5Z01p0lhorqy0BlwFrcN6u3lmN7K0U/e1uezmr' crossorigin='anonymous' referrerpolicy='no-referrer'></script>
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
    --sidebar-w: 240px;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: 'IBM Plex Sans', system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    font-size: 14px;
    font-feature-settings: "ss01", "ss02";
    /* Blueprint grid overlay — subtle ops-console atmosphere */
    background-image:
      linear-gradient(to right, rgba(27, 42, 74, 0.035) 1px, transparent 1px),
      linear-gradient(to bottom, rgba(27, 42, 74, 0.035) 1px, transparent 1px);
    background-size: 32px 32px;
    background-attachment: fixed;
  }
  code, .mono, .num { font-family: 'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-variant-numeric: tabular-nums; }
  h1, h2, h3 { font-family: 'IBM Plex Sans', sans-serif; line-height: 1.15; letter-spacing: -0.01em; }
  a { color: var(--accent); text-decoration: none; }
  a:hover { text-decoration: underline; }

  .skip-link {
    position: absolute; top: -40px; left: 0;
    background: var(--navy); color: white; padding: 8px 12px;
    z-index: 100;
  }
  .skip-link:focus { top: 0; }

  header.topbar {
    background: var(--navy);
    color: white;
    padding: 14px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 4px solid var(--accent);
  }
  .topbar .brand { display: flex; align-items: center; gap: 12px; color: white; text-decoration: none; }
  .topbar .brand:hover { text-decoration: none; opacity: 0.9; }
  .brand-logo { border-radius: 6px; background: white; padding: 3px; display: block; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.1; }
  .brand-name { font-size: 16px; font-weight: 700; letter-spacing: 0.3px; }
  .brand-sub { font-size: 12px; opacity: 0.8; font-weight: 400; }
  .topbar .meta { font-size: 12px; opacity: 0.85; }
  .topbar .meta span { margin-left: 16px; }
  .topbar-link { color: #9BB3E0 !important; text-decoration: underline; }
  .topbar-link:hover { color: white !important; }

  .layout {
    display: grid;
    grid-template-columns: var(--sidebar-w) 1fr;
    min-height: calc(100vh - 60px);
  }

  nav.sidebar {
    background: white;
    border-right: 1px solid var(--border);
    padding: 24px 12px;
    position: sticky;
    top: 0;
    align-self: start;
    height: 100vh;
    overflow-y: auto;
  }
  nav.sidebar a {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 12px;
    color: var(--muted);
    border-radius: 6px;
    margin-bottom: 2px;
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    transition: background 150ms, color 150ms;
  }
  nav.sidebar a:hover { background: #F1F5F9; color: var(--text); text-decoration: none; }
  nav.sidebar a.active { background: var(--navy); color: white; }
  nav.sidebar a [data-lucide] { width: 16px; height: 16px; }

  main { padding: 24px 32px; max-width: 1400px; }
  section { margin-bottom: 40px; scroll-margin-top: 20px; }
  section h2 {
    font-size: 20px;
    font-weight: 700;
    color: var(--navy);
    margin: 0 0 16px 0;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  section h2 [data-lucide] { width: 20px; height: 20px; color: var(--accent); }

  /* Hero posture — ops-console blueprint */
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
    grid-template-columns: auto 1fr;
    gap: 40px;
    align-items: center;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
  }
  .posture-hero::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent 0%, var(--accent) 30%, var(--accent) 70%, transparent 100%);
    opacity: 0.7;
  }
  .posture-hero::after {
    content: 'CA // POSTURE';
    position: absolute;
    top: 10px; right: 16px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.15em;
    color: var(--accent);
    opacity: 0.75;
  }
  .posture-score {
    font-size: 88px;
    font-weight: 700;
    line-height: 1;
    font-family: 'IBM Plex Mono', ui-monospace, monospace;
  }
  .posture-meta { display: flex; flex-direction: column; gap: 8px; }
  .posture-band {
    display: inline-block;
    padding: 4px 14px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    width: fit-content;
    color: white;
    background: $bandColor;
  }
  .posture-desc { font-size: 14px; opacity: 0.85; max-width: 560px; }

  /* Phase bars */
  .phase-bars { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }
  .phase-bar {
    background: white;
    border-radius: 8px;
    padding: 14px 16px;
    border: 1px solid var(--border);
  }
  .phase-name { font-size: 11px; text-transform: uppercase; color: var(--muted); font-weight: 600; letter-spacing: 0.5px; margin-bottom: 6px; }
  .phase-value { font-size: 22px; font-weight: 700; color: var(--navy); font-family: 'IBM Plex Mono', ui-monospace, monospace; margin-bottom: 6px; }
  .phase-bar-track { height: 6px; background: var(--border); border-radius: 3px; overflow: hidden; }
  .phase-bar-fill { height: 100%; background: var(--accent); border-radius: 3px; transition: width 400ms ease; }

  /* KPI row */
  .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 24px; }
  .kpi {
    background: white;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
  }
  .kpi-label { font-size: 11px; text-transform: uppercase; color: var(--muted); font-weight: 600; letter-spacing: 0.5px; }
  .kpi-value { font-size: 24px; font-weight: 700; color: var(--navy); font-family: 'IBM Plex Mono', ui-monospace, monospace; margin-top: 4px; }
  .kpi.crit .kpi-value { color: #B91C1C; }
  .kpi.warn .kpi-value { color: #D97706; }

  /* Insights */
  .insight-card {
    background: white;
    border-left: 4px solid var(--accent);
    border-radius: 6px;
    padding: 16px 20px;
    margin-bottom: 12px;
  }
  .insight-head { display: flex; gap: 8px; align-items: center; margin-bottom: 8px; flex-wrap: wrap; }
  .sev-pill {
    color: white;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .phase-tag {
    background: var(--navy);
    color: white;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 10px;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  .insight-id { font-family: 'IBM Plex Mono', ui-monospace, monospace; font-size: 11px; color: var(--muted); }
  .insight-title { font-size: 15px; font-weight: 600; margin: 4px 0 8px 0; color: var(--navy); }
  .insight-finding, .insight-rec { margin: 4px 0; font-size: 13px; }
  details.affected { margin-top: 8px; font-size: 13px; cursor: pointer; }
  details.affected summary { color: var(--accent); font-weight: 500; }
  details.affected ul { margin: 6px 0 0 20px; color: var(--muted); }

  /* Missing Controls */
  .mc-group-head { display: flex; align-items: center; gap: 10px; margin: 20px 0 12px 0; font-size: 15px; color: var(--navy); }
  .mc-count { color: var(--muted); font-size: 13px; font-weight: 400; }
  .mc-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 10px; margin-bottom: 20px; }
  .mc-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 14px;
    transition: border-color 150ms;
  }
  .mc-card:hover { border-color: var(--accent); }
  .mc-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
  .mc-type { font-size: 10px; text-transform: uppercase; color: var(--muted); font-weight: 600; letter-spacing: 0.5px; }
  .mc-name { font-size: 13px; color: var(--text); font-weight: 500; }

  /* Badges */
  .lic-badge, .pri-badge, .state-badge, .score-badge, .flag-pill {
    display: inline-block;
    color: white;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 600;
    white-space: nowrap;
  }
  .score-badge { font-family: 'IBM Plex Mono', ui-monospace, monospace; min-width: 36px; text-align: center; }
  .flag-pill { background: #475569; margin-left: 6px; font-size: 10px; }

  /* Tables */
  table.table {
    width: 100%;
    background: white;
    border-collapse: collapse;
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid var(--border);
  }
  table.table th {
    background: var(--navy);
    color: white;
    text-align: left;
    padding: 10px 12px;
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  table.table td {
    padding: 10px 12px;
    border-top: 1px solid var(--border);
    font-size: 13px;
    vertical-align: top;
  }
  table.table tbody tr:hover { background: #F8FAFC; }
  .col-name { min-width: 200px; font-weight: 500; }
  .col-score { width: 70px; text-align: center; }
  .col-assign, .col-cond { font-size: 12px; color: var(--muted); max-width: 320px; }
  ul.gap-list { margin: 0; padding-left: 16px; }
  ul.gap-list li { margin: 2px 0; }
  .muted { color: var(--muted); }

  /* Tabs */
  .tabs { display: flex; gap: 4px; border-bottom: 2px solid var(--border); margin-bottom: 16px; }
  .tab {
    background: transparent;
    border: none;
    padding: 10px 16px;
    font-family: inherit;
    font-size: 13px;
    font-weight: 600;
    color: var(--muted);
    cursor: pointer;
    border-bottom: 2px solid transparent;
    margin-bottom: -2px;
    transition: color 150ms, border-color 150ms;
  }
  .tab:hover { color: var(--text); }
  .tab.active { color: var(--navy); border-bottom-color: var(--accent); }
  .tab-count {
    display: inline-block;
    background: var(--border);
    color: var(--muted);
    font-size: 11px;
    padding: 1px 8px;
    border-radius: 10px;
    margin-left: 6px;
    font-weight: 600;
  }
  .tab.active .tab-count { background: var(--accent); color: white; }
  .tab-panel[hidden] { display: none; }

  /* Assignment badge in nested group table */
  .assign-badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 600;
    color: white;
  }
  .assign-badge.include { background: #2563EB; }
  .assign-badge.exclude { background: #B91C1C; }

  .section-lead { color: var(--muted); font-size: 13px; margin: 0 0 12px 0; max-width: 780px; }
  .subhead { font-size: 14px; color: var(--navy); margin: 24px 0 10px 0; font-weight: 600; }

  /* DataTables overrides */
  .dt-search, .dt-length { margin-bottom: 10px; font-size: 13px; }
  .dt-paging button.dt-paging-button { font-size: 12px !important; }

  .empty { color: var(--muted); font-style: italic; padding: 20px; text-align: center; }
  .warning-box {
    background: #FEF3C7;
    border-left: 4px solid #D97706;
    color: #92400E;
    padding: 14px 18px;
    border-radius: 6px;
  }

  /* Responsive */
  @media (max-width: 900px) {
    .layout { grid-template-columns: 1fr; }
    nav.sidebar {
      position: static;
      height: auto;
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
      padding: 10px;
    }
    nav.sidebar a { flex: 1 1 auto; padding: 8px 10px; font-size: 12px; }
    main { padding: 16px; }
    .posture-hero { grid-template-columns: 1fr; gap: 16px; padding: 20px; }
    .posture-score { font-size: 60px; }
    .phase-bars { grid-template-columns: repeat(2, 1fr); }
  }

  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
      animation-duration: 0.01ms !important;
      transition-duration: 0.01ms !important;
    }
  }

  @media print {
    nav.sidebar, .skip-link { display: none; }
    .layout { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>
<a href='#main' class='skip-link'>Skip to content</a>

<header class='topbar' role='banner'>
  <a href='https://www.griffin31.com' target='_blank' rel='noopener' class='brand' aria-label='Griffin31 website'>
    <img src='https://avatars.githubusercontent.com/u/230988388?s=120' alt='Griffin31' class='brand-logo' width='36' height='36'>
    <span class='brand-text'>
      <span class='brand-name'>Griffin31</span>
      <span class='brand-sub'>Conditional Access Analysis</span>
    </span>
  </a>
  <div class='meta'>
    <span><strong>Tenant:</strong> $(HtmlEncode $tenantName)</span>
    <span><strong>Generated:</strong> $generated</span>
    <span><a href='https://www.griffin31.com' target='_blank' rel='noopener' class='topbar-link'>griffin31.com</a></span>
  </div>
</header>

<div class='layout'>
  <nav class='sidebar' role='navigation' aria-label='Report sections'>
    <a href='#overview' class='nav-link active'><i data-lucide='layout-dashboard'></i> Overview</a>
    <a href='#insights' class='nav-link'><i data-lucide='alert-triangle'></i> Key Insights</a>
    <a href='#controls' class='nav-link'><i data-lucide='shield-check'></i> Missing Controls</a>
    <a href='#policies' class='nav-link'><i data-lucide='list-checks'></i> Policies</a>
    <a href='#breakglass' class='nav-link'><i data-lucide='key-round'></i> Break-Glass</a>
    <a href='#nested' class='nav-link'><i data-lucide='users'></i> Nested Groups</a>
  </nav>

  <main id='main' role='main'>

    <section id='overview'>
      <h2><i data-lucide='layout-dashboard'></i> Overview</h2>

      <div class='posture-hero'>
        <div class='posture-score'>$postureScore</div>
        <div class='posture-meta'>
          <span class='posture-band'>$postureBand</span>
          <div style='font-size:14px;font-weight:600;'>Tenant CA Posture Score</div>
          <div class='posture-desc'>Scored 0-100 based on Microsoft's deployment phase model (Foundation, Core, Advanced, Admin). Deductions for critical and high-severity insights.</div>
        </div>
      </div>

      <div class='phase-bars'>
        <div class='phase-bar'>
          <div class='phase-name'>Foundation</div>
          <div class='phase-value'>$phaseF</div>
          <div class='phase-bar-track'><div class='phase-bar-fill' style='width:$phaseF%'></div></div>
        </div>
        <div class='phase-bar'>
          <div class='phase-name'>Core Auth</div>
          <div class='phase-value'>$phaseC</div>
          <div class='phase-bar-track'><div class='phase-bar-fill' style='width:$phaseC%'></div></div>
        </div>
        <div class='phase-bar'>
          <div class='phase-name'>Advanced</div>
          <div class='phase-value'>$phaseA</div>
          <div class='phase-bar-track'><div class='phase-bar-fill' style='width:$phaseA%'></div></div>
        </div>
        <div class='phase-bar'>
          <div class='phase-name'>Admin Protection</div>
          <div class='phase-value'>$phaseAdmin</div>
          <div class='phase-bar-track'><div class='phase-bar-fill' style='width:$phaseAdmin%'></div></div>
        </div>
      </div>

      <div class='kpi-row'>
        <div class='kpi'><div class='kpi-label'>Total Policies</div><div class='kpi-value'>$totalPolicies</div></div>
        <div class='kpi'><div class='kpi-label'>Enabled</div><div class='kpi-value'>$enabledCount</div></div>
        <div class='kpi'><div class='kpi-label'>Report-Only</div><div class='kpi-value'>$reportOnly</div></div>
        <div class='kpi'><div class='kpi-label'>Disabled</div><div class='kpi-value'>$disabled</div></div>
        <div class='kpi'><div class='kpi-label'>Avg Policy Score</div><div class='kpi-value'>$avgScore</div></div>
        <div class='kpi crit'><div class='kpi-label'>Critical Insights</div><div class='kpi-value'>$critCount</div></div>
        <div class='kpi warn'><div class='kpi-label'>High Insights</div><div class='kpi-value'>$highCount</div></div>
      </div>
    </section>

    <section id='insights'>
      <h2><i data-lucide='alert-triangle'></i> Key Insights</h2>
      $insightsHtml
    </section>

    <section id='controls'>
      <h2><i data-lucide='shield-check'></i> Missing Controls</h2>
      $missingHtml
    </section>

    <section id='policies'>
      <h2><i data-lucide='list-checks'></i> Policies</h2>
      <div class='tabs' role='tablist'>
        <button class='tab active' role='tab' data-target='tab-enabled' aria-selected='true'>Enabled <span class='tab-count'>$enabledTotal</span></button>
        <button class='tab' role='tab' data-target='tab-report' aria-selected='false'>Report-only <span class='tab-count'>$reportTotal</span></button>
        <button class='tab' role='tab' data-target='tab-disabled' aria-selected='false'>Disabled <span class='tab-count'>$disabledTotal</span></button>
      </div>
      <div id='tab-enabled' class='tab-panel active' role='tabpanel'>$tblEnabled</div>
      <div id='tab-report' class='tab-panel' role='tabpanel' hidden>$tblReport</div>
      <div id='tab-disabled' class='tab-panel' role='tabpanel' hidden>$tblDisabled</div>
    </section>

    <section id='breakglass'>
      <h2><i data-lucide='key-round'></i> Break-Glass Accounts</h2>
      $bgHtml
    </section>

    <section id='nested'>
      <h2><i data-lucide='users'></i> Nested Groups</h2>
      $ngHtml
    </section>

  </main>
</div>

<script>
  // Lucide icons
  document.addEventListener('DOMContentLoaded', function() {
    if (typeof lucide !== 'undefined') lucide.createIcons();
  });

  // DataTables on all three policy tables
  `$(document).ready(function() {
    ['#policiesEnabled', '#policiesReport', '#policiesDisabled'].forEach(function(sel) {
      if (`$(sel).length) {
        `$(sel).DataTable({
          pageLength: 25,
          order: [[1, 'asc']],
          columnDefs: [{ targets: [2, 3], orderable: false }]
        });
      }
    });
  });

  // Policy state tabs
  document.querySelectorAll('.tab').forEach(function(btn) {
    btn.addEventListener('click', function() {
      document.querySelectorAll('.tab').forEach(function(b) {
        b.classList.remove('active');
        b.setAttribute('aria-selected','false');
      });
      document.querySelectorAll('.tab-panel').forEach(function(p) { p.hidden = true; });
      btn.classList.add('active');
      btn.setAttribute('aria-selected','true');
      var target = document.getElementById(btn.dataset.target);
      if (target) {
        target.hidden = false;
        // DataTables needs a redraw after being shown from hidden
        var tbl = target.querySelector('table.dataTable');
        if (tbl && window.jQuery) { window.jQuery(tbl).DataTable().columns.adjust(); }
      }
    });
  });

  // Sidebar active link on scroll
  const links = document.querySelectorAll('.nav-link');
  const sections = Array.from(links).map(l => document.querySelector(l.getAttribute('href')));
  function setActive() {
    const y = window.scrollY + 80;
    let idx = 0;
    sections.forEach((s, i) => { if (s && s.offsetTop <= y) idx = i; });
    links.forEach((l, i) => l.classList.toggle('active', i === idx));
  }
  window.addEventListener('scroll', setActive, { passive: true });
</script>

</body>
</html>
"@

$html | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Report generated: $OutputPath" -ForegroundColor Green
