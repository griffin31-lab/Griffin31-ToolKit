[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExportDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

# ── Load data ──
$policies = Get-Content "$ExportDir/ConditionalAccessPolicies.json" | ConvertFrom-Json
$groups   = Get-Content "$ExportDir/SecurityGroups.json" | ConvertFrom-Json
$members  = Get-Content "$ExportDir/Members.json" | ConvertFrom-Json
$guests   = Get-Content "$ExportDir/Guests.json" | ConvertFrom-Json
$apps     = Get-Content "$ExportDir/ServicePrincipals.json" | ConvertFrom-Json

# Named locations (optional)
$locMap = @{}
if (Test-Path "$ExportDir/NamedLocations.json") {
    $namedLocations = Get-Content "$ExportDir/NamedLocations.json" | ConvertFrom-Json
    foreach ($l in $namedLocations) {
        if ($l.Id -and $l.DisplayName) { $locMap[$l.Id] = $l.DisplayName }
    }
}

# Break-glass (optional, from prior stage)
$breakGlassUsers = @()
$breakGlassPath = Join-Path (Split-Path $ExportDir -Parent) "data/breakglass.json"
if (Test-Path $breakGlassPath) {
    $breakGlassData = Get-Content $breakGlassPath | ConvertFrom-Json
    $breakGlassUsers = $breakGlassData.BreakGlassUsers
}
$breakGlassIds = @($breakGlassUsers | ForEach-Object { $_.Id })

# Group memberships
$groupMems = @{}
if (Test-Path "$ExportDir/SecurityGroupMemberships.json") {
    $gm = Get-Content "$ExportDir/SecurityGroupMemberships.json" | ConvertFrom-Json
    foreach ($g in $gm) {
        if ($g.GroupId -and $g.Members) { $groupMems[$g.GroupId] = $g.Members.Id }
    }
}

$allUsers = @($members + $guests)
$userMap  = @{}; foreach ($u in $allUsers) { $userMap[$u.Id] = $u.DisplayName }
$groupMap = @{}; foreach ($g in $groups)  { $groupMap[$g.Id] = $g.DisplayName }
$appMap   = @{}; foreach ($a in $apps)    { $appMap[$a.Id]   = $a.DisplayName }

# ── Effective coverage helper ──
function Get-PolicyEffectiveCoverage {
    param($policy)

    $includedUsers = @()
    if ($policy.Conditions.Users.IncludeUsers -contains "All") {
        $includedUsers = $allUsers.Id
    } else {
        $includedUsers += $policy.Conditions.Users.IncludeUsers | Where-Object { $_ }
    }
    foreach ($groupId in $policy.Conditions.Users.IncludeGroups) {
        if ($groupMems.ContainsKey($groupId)) { $includedUsers += $groupMems[$groupId] }
    }

    $excludedUsers = @()
    $excludedUsers += $policy.Conditions.Users.ExcludeUsers | Where-Object { $_ }
    foreach ($groupId in $policy.Conditions.Users.ExcludeGroups) {
        if ($groupMems.ContainsKey($groupId)) { $excludedUsers += $groupMems[$groupId] }
    }

    $includedApps = @()
    if ($policy.Conditions.Applications.IncludeApplications -contains "All") {
        $includedApps = $apps.Id
    } else {
        $includedApps = $policy.Conditions.Applications.IncludeApplications | Where-Object { $_ }
    }
    $excludedApps = $policy.Conditions.Applications.ExcludeApplications | Where-Object { $_ }

    $includedUsers = $includedUsers | Select-Object -Unique
    $excludedUsers = $excludedUsers | Select-Object -Unique
    $includedApps  = $includedApps  | Select-Object -Unique
    $excludedApps  = $excludedApps  | Select-Object -Unique

    return @{
        IncludedUsers = $includedUsers | Where-Object { $_ -notin $excludedUsers }
        IncludedApps  = $includedApps  | Where-Object { $_ -notin $excludedApps }
        ExcludedUsers = $excludedUsers
        ExcludedApps  = $excludedApps
    }
}

# ── Global coverage (what's protected by at least one enabled policy) ──
$userCovered = @{}
$appCovered  = @{}
foreach ($policy in $policies | Where-Object { $_.state -eq 'enabled' }) {
    $eff = Get-PolicyEffectiveCoverage $policy
    foreach ($u in $eff.IncludedUsers) { $userCovered[$u] = $true }
    foreach ($a in $eff.IncludedApps)  { $appCovered[$a]  = $true }
}

# ── Per-policy scoring ──
function Get-PolicyAnalysis {
    param($policy)

    $score = 100
    $assignmentGaps = @()
    $conditionGaps  = @()
    $flags = @()

    # No-op policy check (enabled but no controls)
    $hasGrantControls = $policy.GrantControls -and (
        ($policy.GrantControls.BuiltInControls -and $policy.GrantControls.BuiltInControls.Count -gt 0) -or
        $policy.GrantControls.AuthenticationStrength -or
        $policy.GrantControls.TermsOfUse -or
        $policy.GrantControls.CustomAuthenticationFactors
    )
    $hasSessionControls = $policy.SessionControls -and (
        $policy.SessionControls.ApplicationEnforcedRestrictions -or
        $policy.SessionControls.CloudAppSecurity -or
        $policy.SessionControls.SignInFrequency -or
        $policy.SessionControls.PersistentBrowser -or
        $policy.SessionControls.ContinuousAccessEvaluation -or
        $policy.SessionControls.DisableResilienceDefaults -or
        $policy.SessionControls.SecureSignInSession
    )
    if ($policy.state -eq 'enabled' -and -not $hasGrantControls -and -not $hasSessionControls) {
        $flags += "NoOp"
        $assignmentGaps += "Policy is enabled but has neither grant nor session controls (no-op)"
        return @{
            Score = 0
            AssignmentGaps = $assignmentGaps
            ConditionGaps = $conditionGaps
            Flags = $flags
        }
    }

    # Weak MFA: MFA required but no authentication strength
    $grantBuiltIns = @()
    if ($policy.GrantControls -and $policy.GrantControls.BuiltInControls) {
        $grantBuiltIns = @($policy.GrantControls.BuiltInControls)
    }
    if ($grantBuiltIns -contains "mfa" -and -not $policy.GrantControls.AuthenticationStrength) {
        $score -= 10
        $flags += "WeakMFA"
        $conditionGaps += "MFA required but no authentication strength set (SMS/voice still allowed)"
    }

    # Report-only lingering
    if ($policy.state -eq 'enabledForReportingButNotEnforced' -and $policy.modifiedDateTime) {
        try {
            $modified = [datetime]::Parse($policy.modifiedDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            if (((Get-Date) - $modified).TotalDays -gt 30) {
                $flags += "ReportOnlyStuck"
            }
        } catch {}
    }

    # Stale policy
    if ($policy.modifiedDateTime) {
        try {
            $modified = [datetime]::Parse($policy.modifiedDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            if (((Get-Date) - $modified).TotalDays -gt 365) {
                $flags += "Stale"
            }
        } catch {}
    }

    # Assignment: group exclusions
    $relevantGroupNames = @()
    $groupMemberIds = @()
    foreach ($groupId in $policy.Conditions.Users.ExcludeGroups) {
        if ($groupMems.ContainsKey($groupId)) {
            $groupMembers = $groupMems[$groupId] | Where-Object { $_ -and $_ -notin $breakGlassIds }
            $count = @($groupMembers).Count
            $score -= ($count * 5)
            $displayName = if ($groupMap[$groupId]) { $groupMap[$groupId] } else { $groupId }
            $relevantGroupNames += "$displayName ($count members)"
            $groupMemberIds += $groupMembers
        }
    }
    if ($relevantGroupNames.Count -gt 0) {
        $assignmentGaps += "Groups excluded: $($relevantGroupNames -join ', ')"
    }

    # Assignment: direct user exclusions
    $relevantUserNames = @()
    foreach ($userId in $policy.Conditions.Users.ExcludeUsers) {
        if ($userId -and $userId -notin $breakGlassIds -and $userId -notin $groupMemberIds -and $userCovered[$userId]) {
            $score -= 5
            $relevantUserNames += if ($userMap[$userId]) { $userMap[$userId] } else { $userId }
        }
    }
    if ($relevantUserNames.Count -gt 0) {
        $summary = if ($relevantUserNames.Count -le 2) { $relevantUserNames -join ', ' }
                   else { "$($relevantUserNames[0..1] -join ', ') + $($relevantUserNames.Count - 2) more" }
        $assignmentGaps += "Users bypassed: $summary"
    }

    # Assignment: app exclusions
    $relevantAppNames = @()
    foreach ($appId in $policy.Conditions.Applications.ExcludeApplications) {
        if ($appId -and $appCovered[$appId]) {
            $score -= 10
            $relevantAppNames += if ($appMap[$appId]) { $appMap[$appId] } else { $appId }
        }
    }
    if ($relevantAppNames.Count -gt 0) {
        $summary = if ($relevantAppNames.Count -le 2) { $relevantAppNames -join ', ' }
                   else { "$($relevantAppNames[0..1] -join ', ') + $($relevantAppNames.Count - 2) more" }
        $assignmentGaps += "Apps unprotected: $summary"
    }

    # Single app targeting
    if ($policy.Conditions.Applications.IncludeApplications -and
        $policy.Conditions.Applications.IncludeApplications.Count -eq 1 -and
        $policy.Conditions.Applications.IncludeApplications -notcontains "All") {
        $score -= 15
        $assignmentGaps += "Scope: only one app targeted"
    }

    # Single user targeting
    if ($policy.Conditions.Users.IncludeUsers -and
        $policy.Conditions.Users.IncludeUsers.Count -eq 1 -and
        $policy.Conditions.Users.IncludeUsers -notcontains "All" -and
        -not $policy.Conditions.Users.IncludeGroups -and
        -not $policy.Conditions.Users.IncludeRoles) {
        $score -= 20
        $assignmentGaps += "Scope: only one user targeted"
    }

    # Condition: platform exclusions
    if ($policy.Conditions.Platforms.ExcludePlatforms) {
        $excluded = $policy.Conditions.Platforms.ExcludePlatforms | Where-Object { $_ }
        $desktop = @('windows','macOS','linux')
        $mobile  = @('android','iOS')
        $excDesktop = $excluded | Where-Object { $desktop -contains $_ }
        $excMobile  = $excluded | Where-Object { $mobile  -contains $_ }
        if ($excDesktop) {
            $score -= ($excDesktop.Count * 10)
            $conditionGaps += "Platforms excluded: $($excDesktop -join ', ')"
        }
        if ($excMobile) {
            $score -= ($excMobile.Count * 3)
            $conditionGaps += "Platforms excluded: $($excMobile -join ', ')"
        }
    }

    # Condition: platform include limitation
    if ($policy.Conditions.Platforms.IncludePlatforms -and
        $policy.Conditions.Platforms.IncludePlatforms -notcontains "all") {
        $included = $policy.Conditions.Platforms.IncludePlatforms | Where-Object { $_ -and $_ -ne 'all' }
        $conditionGaps += "Platforms limited to: $($included -join ', ')"
    }

    # Condition: location exclusions
    if ($policy.Conditions.Locations.ExcludeLocations) {
        $locs = $policy.Conditions.Locations.ExcludeLocations | Where-Object { $_ } |
            ForEach-Object { if ($locMap[$_]) { $locMap[$_] } else { $_ } }
        if ($locs) {
            $score -= ($locs.Count * 10)
            $conditionGaps += "Trusted locations bypassed: $($locs -join ', ')"
        }
    }
    if ($policy.Conditions.Locations.IncludeLocations -and
        $policy.Conditions.Locations.IncludeLocations -notcontains 'All' -and
        $policy.Conditions.Locations.IncludeLocations -notcontains 'AllTrusted') {
        $locs = $policy.Conditions.Locations.IncludeLocations | Where-Object { $_ -and $_ -ne 'All' -and $_ -ne 'AllTrusted' } |
            ForEach-Object { if ($locMap[$_]) { $locMap[$_] } else { $_ } }
        if ($locs) {
            $score -= 5
            $conditionGaps += "Only applies from: $($locs -join ', ')"
        }
    }

    # Condition: client app types (don't penalize — MS templates target specific types)
    if ($policy.Conditions.ClientAppTypes -and $policy.Conditions.ClientAppTypes -notcontains 'all') {
        $types = $policy.Conditions.ClientAppTypes | Where-Object { $_ -and $_ -ne 'all' }
        $conditionGaps += "Client app types: $($types -join ', ')"
    }

    $score = [Math]::Max(0, $score)

    return @{
        Score = $score
        AssignmentGaps = $assignmentGaps
        ConditionGaps  = $conditionGaps
        Flags = $flags
    }
}

# ── Run analysis ──
$policyGaps = @()
foreach ($policy in $policies) {
    $r = Get-PolicyAnalysis $policy

    # Emit gaps as arrays — the HTML generator escapes each item before rendering.
    $policyGaps += @{
        PolicyId       = $policy.id
        PolicyName     = $policy.displayName
        State          = $policy.state
        Created        = if ($policy.createdDateTime)  { (Get-Date $policy.createdDateTime).ToString("yyyy-MM-dd") } else { $null }
        Modified       = if ($policy.modifiedDateTime) { (Get-Date $policy.modifiedDateTime).ToString("yyyy-MM-dd") } else { $null }
        AssignmentGaps = @($r.AssignmentGaps)
        ConditionGaps  = @($r.ConditionGaps)
        Score          = $r.Score
        Flags          = $r.Flags
    }
}

$result = @{
    PolicyGaps = $policyGaps
    Summary = @{
        TotalPolicies    = $policies.Count
        EnabledPolicies  = ($policies | Where-Object { $_.state -eq 'enabled' }).Count
        ReportOnly       = ($policies | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' }).Count
        Disabled         = ($policies | Where-Object { $_.state -eq 'disabled' }).Count
        AvgScore         = if ($policyGaps.Count -gt 0) { [math]::Round(($policyGaps | Measure-Object -Property Score -Average).Average, 1) } else { 0 }
    }
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$result | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Policy gap analysis complete: Analyzed $($policies.Count) policies" -ForegroundColor Green
