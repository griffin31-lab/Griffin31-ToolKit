[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExportDir,
    [Parameter(Mandatory)]
    [string]$AnalysisDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

# ── Load data ──
$policies = Get-Content "$ExportDir/ConditionalAccessPolicies.json" | ConvertFrom-Json
$members  = Get-Content "$ExportDir/Members.json" | ConvertFrom-Json
$guests   = Get-Content "$ExportDir/Guests.json" | ConvertFrom-Json

$breakGlassPath = "$AnalysisDir/breakglass.json"
$breakGlassData = if (Test-Path $breakGlassPath) { Get-Content $breakGlassPath | ConvertFrom-Json } else { $null }
$breakGlassIds  = @()
if ($breakGlassData -and $breakGlassData.BreakGlassUsers) { $breakGlassIds = @($breakGlassData.BreakGlassUsers.Id) }

$missingCtrlPath = "$AnalysisDir/missing-controls.json"
$missingCtrlData = if (Test-Path $missingCtrlPath) { Get-Content $missingCtrlPath | ConvertFrom-Json } else { $null }

$policyGapsPath = "$AnalysisDir/policy-gaps.json"
$policyGapsData = if (Test-Path $policyGapsPath) { Get-Content $policyGapsPath | ConvertFrom-Json } else { $null }

$enabled = $policies | Where-Object { $_.state -eq 'enabled' }
$reportOnly = $policies | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' }

$insights = @()

function Add-Insight {
    param([string]$Id, [string]$Title, [string]$Severity, [string]$Phase, [string]$Finding, [string]$Recommendation, [array]$AffectedPolicies = @())
    $script:insights += [PSCustomObject]@{
        Id = $Id; Title = $Title; Severity = $Severity; Phase = $Phase
        Finding = $Finding; Recommendation = $Recommendation
        AffectedPolicies = @($AffectedPolicies)
    }
}

# ── CRITICAL CHECKS ──

# 1. Lockout risk — All-users policies without break-glass exclusion
$lockoutPolicies = @()
foreach ($p in $enabled) {
    if ($p.Conditions.Users.IncludeUsers -contains "All") {
        $excludedUsers = @($p.Conditions.Users.ExcludeUsers)
        $hasBreakGlass = $false
        if ($breakGlassIds.Count -gt 0) {
            $hasBreakGlass = $excludedUsers | Where-Object { $_ -in $breakGlassIds }
        }
        if (-not $hasBreakGlass) { $lockoutPolicies += $p.displayName }
    }
}
if ($lockoutPolicies.Count -gt 0) {
    Add-Insight -Id "INS-001" -Title "Lockout risk: All-users policies without break-glass exclusion" `
        -Severity "Critical" -Phase "Foundation" `
        -Finding "$($lockoutPolicies.Count) enabled policy(s) target All users but do not exclude a break-glass account." `
        -Recommendation "Exclude at least one dedicated emergency access account from every All-users policy. See: aka.ms/breakglass" `
        -AffectedPolicies $lockoutPolicies
}

# 2. Break-glass not configured
if ($breakGlassIds.Count -eq 0) {
    Add-Insight -Id "INS-002" -Title "No break-glass account detected" `
        -Severity "Critical" -Phase "Foundation" `
        -Finding "No account is excluded from all enabled CA policies and holds Global Administrator." `
        -Recommendation "Create two dedicated cloud-only emergency access accounts with Global Administrator, exclude them from all CA policies, and protect with FIDO2 keys."
}

# 3. No policy blocks legacy authentication
$legacyBlocked = $false
foreach ($p in $enabled) {
    if ($p.GrantControls.BuiltInControls -contains "block" -and
        $p.Conditions.ClientAppTypes -and
        ($p.Conditions.ClientAppTypes -contains "exchangeActiveSync" -or $p.Conditions.ClientAppTypes -contains "other")) {
        $legacyBlocked = $true; break
    }
}
if (-not $legacyBlocked) {
    Add-Insight -Id "INS-003" -Title "Legacy authentication not blocked" `
        -Severity "Critical" -Phase "Foundation" `
        -Finding "No enabled policy blocks legacy authentication (Exchange ActiveSync / Other clients)." `
        -Recommendation "Create a policy that blocks legacy authentication for All users targeting All cloud apps."
}

# 4. No-op policies (enabled, no grant or session controls)
$noOps = @()
foreach ($p in $enabled) {
    $hasGrant = $p.GrantControls -and (
        ($p.GrantControls.BuiltInControls -and $p.GrantControls.BuiltInControls.Count -gt 0) -or
        $p.GrantControls.AuthenticationStrength -or $p.GrantControls.TermsOfUse
    )
    $hasSession = $p.SessionControls -and (
        $p.SessionControls.ApplicationEnforcedRestrictions -or $p.SessionControls.CloudAppSecurity -or
        $p.SessionControls.SignInFrequency -or $p.SessionControls.PersistentBrowser -or
        $p.SessionControls.ContinuousAccessEvaluation -or $p.SessionControls.SecureSignInSession
    )
    if (-not $hasGrant -and -not $hasSession) { $noOps += $p.displayName }
}
if ($noOps.Count -gt 0) {
    Add-Insight -Id "INS-004" -Title "Enabled policies with no controls (no-op)" `
        -Severity "Critical" -Phase "Foundation" `
        -Finding "$($noOps.Count) enabled policy(s) have neither grant nor session controls — they block nothing." `
        -Recommendation "Add a grant or session control, or disable the policy." `
        -AffectedPolicies $noOps
}

# 5. May 2026 enforcement risk — resource exclusions on All-apps policies
$enforcementRisk = @()
foreach ($p in $enabled) {
    if ($p.Conditions.Applications.IncludeApplications -contains "All" -and
        $p.Conditions.Applications.ExcludeApplications -and
        $p.Conditions.Applications.ExcludeApplications.Count -gt 0) {
        $enforcementRisk += $p.displayName
    }
}
if ($enforcementRisk.Count -gt 0) {
    Add-Insight -Id "INS-005" -Title "May 2026 enforcement change — resource exclusions" `
        -Severity "Critical" -Phase "Foundation" `
        -Finding "$($enforcementRisk.Count) policy(s) target All apps with resource exclusions. Microsoft enforcement change on May 13, 2026 will apply these policies on all sign-ins regardless of resource-exclusion scope." `
        -Recommendation "Review policies — confirm excluded apps genuinely should bypass, or move to Filter-for-applications." `
        -AffectedPolicies $enforcementRisk
}

# ── HIGH CHECKS ──

# 6. Weak MFA — MFA without auth strength
$weakMFA = @()
foreach ($p in $enabled) {
    if ($p.GrantControls.BuiltInControls -contains "mfa" -and -not $p.GrantControls.AuthenticationStrength) {
        $weakMFA += $p.displayName
    }
}
if ($weakMFA.Count -gt 0) {
    Add-Insight -Id "INS-006" -Title "Weak MFA — no authentication strength set" `
        -Severity "High" -Phase "Core" `
        -Finding "$($weakMFA.Count) MFA policy(s) allow any MFA method including SMS and voice." `
        -Recommendation "Attach an Authentication Strength (e.g., 'Multifactor authentication' or 'Phishing-resistant MFA') to each MFA policy." `
        -AffectedPolicies $weakMFA
}

# 7. Admins not separately protected
$adminPolicies = $enabled | Where-Object {
    $_.Conditions.Users.IncludeRoles -and $_.Conditions.Users.IncludeRoles.Count -gt 0
}
if (@($adminPolicies).Count -eq 0) {
    Add-Insight -Id "INS-007" -Title "Admin accounts not separately protected" `
        -Severity "High" -Phase "Admin" `
        -Finding "No enabled policy targets privileged built-in roles specifically." `
        -Recommendation "Create a stricter policy for admin roles (Global Admin, Security Admin, etc.) requiring phishing-resistant MFA and compliant devices."
}

# 8. No phishing-resistant MFA for admins
$phishResistantAdmin = $false
foreach ($p in $enabled) {
    if ($p.GrantControls.AuthenticationStrength -and
        $p.Conditions.Users.IncludeRoles -and $p.Conditions.Users.IncludeRoles.Count -gt 0) {
        $phishResistantAdmin = $true; break
    }
}
if (-not $phishResistantAdmin) {
    Add-Insight -Id "INS-008" -Title "No phishing-resistant MFA for admins" `
        -Severity "High" -Phase "Admin" `
        -Finding "No policy enforces an Authentication Strength (phishing-resistant) on privileged roles." `
        -Recommendation "Create a policy targeting privileged roles with the 'Phishing-resistant MFA' authentication strength (FIDO2, Windows Hello, cert-based)."
}

# 9. Guest coverage gap
$guestProtected = $false
foreach ($p in $enabled) {
    $inc = $p.Conditions.Users.IncludeUsers
    if ($inc -contains "GuestsOrExternalUsers" -or $inc -contains "All") { $guestProtected = $true; break }
    if ($p.Conditions.Users.IncludeGuestsOrExternalUsers) { $guestProtected = $true; break }
}
if (-not $guestProtected -and $guests.Count -gt 0) {
    Add-Insight -Id "INS-009" -Title "Guest / external user coverage gap" `
        -Severity "High" -Phase "Core" `
        -Finding "$($guests.Count) guest user(s) present but no enabled policy targets GuestsOrExternalUsers." `
        -Recommendation "Create a policy requiring MFA for guest access to all resources."
}

# 10. Device code flow not blocked
$deviceCodeBlocked = $false
foreach ($p in $enabled) {
    if ($p.Conditions.AuthenticationFlows -and $p.Conditions.AuthenticationFlows.TransferMethods -match "deviceCodeFlow") {
        $deviceCodeBlocked = $true; break
    }
}
if (-not $deviceCodeBlocked) {
    Add-Insight -Id "INS-010" -Title "Device code flow not restricted" `
        -Severity "High" -Phase "Advanced" `
        -Finding "No policy restricts device code flow — a known phishing vector."  `
        -Recommendation "Create a policy blocking device code flow for All users (exclude only device-registration accounts that genuinely need it)."
}

# ── MEDIUM CHECKS ──

# 11. CA policies assigned to groups that contain nested groups (management blind spot)
$nestedGroupsPath = "$AnalysisDir/nested-groups.json"
$nestedData = if (Test-Path $nestedGroupsPath) { Get-Content $nestedGroupsPath | ConvertFrom-Json } else { $null }
if ($nestedData -and $nestedData.PoliciesUsingNestedGroups -and $nestedData.PoliciesUsingNestedGroups.Count -gt 0) {
    $affected = $nestedData.PoliciesUsingNestedGroups | ForEach-Object { "$($_.PolicyName) -> $($_.GroupName)" }
    Add-Insight -Id "INS-011" -Title "CA policy assigned to nested group (blind spot)" `
        -Severity "Medium" -Phase "Core" `
        -Finding "$($nestedData.PoliciesUsingNestedGroups.Count) policy assignment(s) reference a group that contains nested groups. The effective member set can change without modifying the policy, creating a management blind spot." `
        -Recommendation "Avoid nested groups in CA assignments. Use flat security groups, dynamic groups, or directory roles for predictable scope." `
        -AffectedPolicies $affected
}

# 12. Report-only stuck > 30 days
$reportStuck = @()
foreach ($p in $reportOnly) {
    if ($p.modifiedDateTime) {
        try {
            $m = [datetime]::Parse($p.modifiedDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            if (((Get-Date) - $m).TotalDays -gt 30) { $reportStuck += $p.displayName }
        } catch {}
    }
}
if ($reportStuck.Count -gt 0) {
    Add-Insight -Id "INS-012" -Title "Report-only policies stuck > 30 days" `
        -Severity "Medium" -Phase "Core" `
        -Finding "$($reportStuck.Count) policy(s) have been in report-only mode for more than 30 days." `
        -Recommendation "Review sign-in logs, then promote to Enabled or retire." `
        -AffectedPolicies $reportStuck
}

# 13. Stale policies > 365 days
$stale = @()
foreach ($p in $enabled) {
    if ($p.modifiedDateTime) {
        try {
            $m = [datetime]::Parse($p.modifiedDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            if (((Get-Date) - $m).TotalDays -gt 365) { $stale += $p.displayName }
        } catch {}
    }
}
if ($stale.Count -gt 0) {
    Add-Insight -Id "INS-013" -Title "Stale policies (not modified > 365 days)" `
        -Severity "Medium" -Phase "Core" `
        -Finding "$($stale.Count) enabled policy(s) have not been reviewed in over a year." `
        -Recommendation "Review annually to ensure scope and controls still align with current security posture." `
        -AffectedPolicies $stale
}

# ── INFO CHECKS ──

# 14. Approaching 240-policy limit
if ($policies.Count -gt 200) {
    Add-Insight -Id "INS-014" -Title "Approaching 240-policy tenant limit" `
        -Severity "Info" -Phase "Core" `
        -Finding "Tenant has $($policies.Count) CA policies (limit is 240 including disabled)." `
        -Recommendation "Consolidate policies using groups, roles, and Filter for applications to stay under the limit."
}

# ── TENANT POSTURE SCORE ──
# Phase weights: Foundation 40 / Core 30 / Advanced 20 / Admin 10
$phaseWeights = @{ Foundation = 40; Core = 30; Advanced = 20; Admin = 10 }
$phaseIssueWeights = @{ Critical = 4; High = 2; Medium = 1; Info = 0 }

# Total possible issues per phase — we normalize via max expected issues per phase
# For simplicity: each phase starts at 100, subtract weighted points per found issue, max 25 per phase
$phaseScores = @{ Foundation = 100; Core = 100; Advanced = 100; Admin = 100 }
foreach ($i in $insights) {
    $deduct = switch ($i.Severity) {
        "Critical" { 25 }
        "High"     { 15 }
        "Medium"   { 8 }
        "Info"     { 2 }
        default    { 0 }
    }
    if ($phaseScores.ContainsKey($i.Phase)) {
        $phaseScores[$i.Phase] = [Math]::Max(0, $phaseScores[$i.Phase] - $deduct)
    }
}

$postureScore = [Math]::Round(
    ($phaseScores.Foundation * $phaseWeights.Foundation / 100) +
    ($phaseScores.Core       * $phaseWeights.Core       / 100) +
    ($phaseScores.Advanced   * $phaseWeights.Advanced   / 100) +
    ($phaseScores.Admin      * $phaseWeights.Admin      / 100), 0
)

$postureBand = if     ($postureScore -ge 81) { "Strong" }
               elseif ($postureScore -ge 61) { "Good" }
               elseif ($postureScore -ge 41) { "Fair" }
               else                          { "Critical" }

$result = @{
    Insights = $insights | Sort-Object @{Expression={
        switch ($_.Severity) { "Critical" {0}; "High" {1}; "Medium" {2}; "Info" {3}; default {4} }
    }}
    PostureScore = $postureScore
    PostureBand  = $postureBand
    PhaseScores  = $phaseScores
    Summary = @{
        Total    = $insights.Count
        Critical = ($insights | Where-Object { $_.Severity -eq 'Critical' }).Count
        High     = ($insights | Where-Object { $_.Severity -eq 'High' }).Count
        Medium   = ($insights | Where-Object { $_.Severity -eq 'Medium' }).Count
        Info     = ($insights | Where-Object { $_.Severity -eq 'Info' }).Count
    }
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$result | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Key insights analysis complete: $($insights.Count) insights, posture score $postureScore ($postureBand)" -ForegroundColor Green
