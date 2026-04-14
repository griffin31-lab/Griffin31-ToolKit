[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExportDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

$policies = Get-Content "$ExportDir/ConditionalAccessPolicies.json" | ConvertFrom-Json

# ── Modern control definitions (2026, validated against MS Learn) ──
# License tags: P1 = Entra ID P1, P2 = Entra ID P2, Purview = Microsoft Purview, WID = Workload Identities license
$controlDefs = @(
    # Grant controls
    @{ Name = "Require multi-factor authentication";                  Type = "Grant";   Priority = "Critical"; License = "P1"; Check = "mfa" }
    @{ Name = "Require authentication strength";                      Type = "Grant";   Priority = "Critical"; License = "P1"; Check = "authStrength" }
    @{ Name = "Require phishing-resistant MFA for admins";            Type = "Grant";   Priority = "Critical"; License = "P1"; Check = "phishResistantAdmin" }
    @{ Name = "Require device to be marked as compliant";             Type = "Grant";   Priority = "High";     License = "P1"; Check = "compliantDevice" }
    @{ Name = "Require app protection policy";                        Type = "Grant";   Priority = "High";     License = "P1"; Check = "compliantApplication" }
    @{ Name = "Require Terms of Use acceptance";                      Type = "Grant";   Priority = "Medium";   License = "P1"; Check = "termsOfUse" }

    # Session controls
    @{ Name = "Sign-in frequency";                                    Type = "Session"; Priority = "High";     License = "P1"; Check = "signInFrequency" }
    @{ Name = "Persistent browser session control";                   Type = "Session"; Priority = "High";     License = "P1"; Check = "persistentBrowser" }
    @{ Name = "Use app enforced restrictions";                        Type = "Session"; Priority = "Medium";   License = "P1"; Check = "appEnforced" }
    @{ Name = "Use Conditional Access App Control (MDCA)";            Type = "Session"; Priority = "Medium";   License = "P1"; Check = "cloudAppSec" }
    @{ Name = "Customize Continuous Access Evaluation";               Type = "Session"; Priority = "High";     License = "P1"; Check = "cae" }
    @{ Name = "Require token protection for sign-in sessions";        Type = "Session"; Priority = "High";     License = "P1"; Check = "tokenProtection" }

    # Policy scenarios (entire policy patterns)
    @{ Name = "Block legacy authentication";                          Type = "Scenario"; Priority = "Critical"; License = "P1"; Check = "blockLegacyAuth" }
    @{ Name = "Block device code flow";                               Type = "Scenario"; Priority = "High";     License = "P1"; Check = "blockDeviceCode" }
    @{ Name = "Block authentication transfer";                        Type = "Scenario"; Priority = "High";     License = "P1"; Check = "blockAuthTransfer" }
    @{ Name = "Restrict high-risk sign-ins";                          Type = "Scenario"; Priority = "High";     License = "P2"; Check = "signInRisk" }
    @{ Name = "Restrict high-risk users";                             Type = "Scenario"; Priority = "High";     License = "P2"; Check = "userRisk" }
    @{ Name = "Insider risk policy";                                  Type = "Scenario"; Priority = "Medium";   License = "Purview"; Check = "insiderRisk" }
    @{ Name = "Conditional Access for workload identities";           Type = "Scenario"; Priority = "High";     License = "WID"; Check = "workloadIdentity" }
    @{ Name = "Block high-risk AI agent identities";                  Type = "Scenario"; Priority = "Medium";   License = "P2"; Check = "aiAgent" }
    @{ Name = "Authentication context in use";                        Type = "Scenario"; Priority = "Medium";   License = "P1"; Check = "authContext" }
    @{ Name = "Filter for devices in use";                            Type = "Scenario"; Priority = "Medium";   License = "P1"; Check = "filterDevices" }
    @{ Name = "Filter for applications in use";                       Type = "Scenario"; Priority = "Medium";   License = "P1"; Check = "filterApps" }
)

# ── Probe policies for each control ──
function Test-ControlUsed {
    param([string]$check, $policies)

    foreach ($p in $policies) {
        if ($p.state -ne 'enabled' -and $p.state -ne 'enabledForReportingButNotEnforced') { continue }

        switch ($check) {
            "mfa"                 { if ($p.GrantControls.BuiltInControls -contains "mfa") { return $true } }
            "authStrength"        { if ($p.GrantControls.AuthenticationStrength) { return $true } }
            "phishResistantAdmin" {
                if ($p.GrantControls.AuthenticationStrength -and $p.Conditions.Users.IncludeRoles -and $p.Conditions.Users.IncludeRoles.Count -gt 0) {
                    return $true
                }
            }
            "compliantDevice"     { if ($p.GrantControls.BuiltInControls -contains "compliantDevice") { return $true } }
            "compliantApplication"{ if ($p.GrantControls.BuiltInControls -contains "compliantApplication") { return $true } }
            "termsOfUse"          { if ($p.GrantControls.TermsOfUse -and $p.GrantControls.TermsOfUse.Count -gt 0) { return $true } }

            "signInFrequency"     { if ($p.SessionControls.SignInFrequency) { return $true } }
            "persistentBrowser"   { if ($p.SessionControls.PersistentBrowser) { return $true } }
            "appEnforced"         { if ($p.SessionControls.ApplicationEnforcedRestrictions) { return $true } }
            "cloudAppSec"         { if ($p.SessionControls.CloudAppSecurity) { return $true } }
            "cae"                 { if ($p.SessionControls.ContinuousAccessEvaluation) { return $true } }
            "tokenProtection"     { if ($p.SessionControls.SecureSignInSession) { return $true } }

            "blockLegacyAuth" {
                if ($p.GrantControls.BuiltInControls -contains "block" -and
                    $p.Conditions.ClientAppTypes -and
                    ($p.Conditions.ClientAppTypes -contains "exchangeActiveSync" -or $p.Conditions.ClientAppTypes -contains "other")) {
                    return $true
                }
            }
            "blockDeviceCode" {
                if ($p.Conditions.AuthenticationFlows -and $p.Conditions.AuthenticationFlows.TransferMethods -match "deviceCodeFlow") {
                    return $true
                }
            }
            "blockAuthTransfer" {
                if ($p.Conditions.AuthenticationFlows -and $p.Conditions.AuthenticationFlows.TransferMethods -match "authenticationTransfer") {
                    return $true
                }
            }
            "signInRisk"       { if ($p.Conditions.SignInRiskLevels -and $p.Conditions.SignInRiskLevels.Count -gt 0) { return $true } }
            "userRisk"         { if ($p.Conditions.UserRiskLevels -and $p.Conditions.UserRiskLevels.Count -gt 0) { return $true } }
            "insiderRisk"      { if ($p.Conditions.InsiderRiskLevels -and $p.Conditions.InsiderRiskLevels.Count -gt 0) { return $true } }
            "workloadIdentity" { if ($p.Conditions.ClientApplications -and $p.Conditions.ClientApplications.IncludeServicePrincipals) { return $true } }
            "aiAgent"          { if ($p.displayName -match "agent" -or ($p.Conditions.ClientApplications -and $p.Conditions.ClientApplications.IncludeServicePrincipals)) { return $true } }
            "authContext"      { if ($p.Conditions.Applications.IncludeAuthenticationContextClassReferences -and $p.Conditions.Applications.IncludeAuthenticationContextClassReferences.Count -gt 0) { return $true } }
            "filterDevices"    { if ($p.Conditions.Devices -and $p.Conditions.Devices.DeviceFilter) { return $true } }
            "filterApps"       { if ($p.Conditions.Applications.ApplicationFilter) { return $true } }
        }
    }
    return $false
}

$missingControls = @()
foreach ($def in $controlDefs) {
    $used = Test-ControlUsed -check $def.Check -policies $policies
    if (-not $used) {
        $missingControls += [PSCustomObject]@{
            ControlName = $def.Name
            ControlType = $def.Type
            Priority    = $def.Priority
            License     = $def.License
        }
    }
}

$result = @{
    MissingControls = $missingControls
    Summary = @{
        TotalEnabledPolicies = ($policies | Where-Object { $_.state -eq 'enabled' }).Count
        TotalMissing         = $missingControls.Count
        CriticalMissing      = ($missingControls | Where-Object { $_.Priority -eq 'Critical' }).Count
        HighMissing          = ($missingControls | Where-Object { $_.Priority -eq 'High' }).Count
        MediumMissing        = ($missingControls | Where-Object { $_.Priority -eq 'Medium' }).Count
    }
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$result | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Missing controls analysis complete: $($missingControls.Count) controls not implemented" -ForegroundColor Green
