[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DataDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$groups = Get-Content (Join-Path $DataDir "groups.json") | ConvertFrom-Json
$teams  = Get-Content (Join-Path $DataDir "teams.json")  | ConvertFrom-Json

$findings = @()

# Check 11: Sensitivity label missing on M365 Group
# Check 13: Guest members in Group without label
foreach ($g in @($groups)) {
    $groupFindings = @()
    if (-not $g.HasSensitivityLabel) {
        $groupFindings += [PSCustomObject]@{
            Id = "GT-011"
            Title = "M365 Group missing sensitivity label"
            Severity = "Medium"
            Details = "Group has no assigned sensitivity label."
            Remediation = "Apply an appropriate sensitivity label in the Microsoft 365 admin center or Purview."
        }
        if ($g.GuestCount -gt 0) {
            $groupFindings += [PSCustomObject]@{
                Id = "GT-013"
                Title = "Group has guests and no sensitivity label"
                Severity = "Medium"
                Details = "$($g.GuestCount) guest(s) in a group without any sensitivity label — content may be oversharing."
                Remediation = "Apply a sensitivity label and review guest membership."
            }
        }
    }

    # Check GT-014 (Data Management): external senders allowed
    if ($g.PSObject.Properties.Name -contains 'AllowExternalSenders' -and $g.AllowExternalSenders -eq $true) {
        $groupFindings += [PSCustomObject]@{
            Id = "GT-014"
            Title = "Group accepts email from external senders"
            Severity = "Low"
            Category = "Data Management"
            Details = "Any internet sender can email this group — phishing payload can land directly in the group inbox / Teams conversation."
            Remediation = "Admin center > Teams &amp; Groups > Group > Settings > uncheck 'Let people outside the organization email this team'."
        }
    }

    # Check GT-015 (Data Management): auto-subscribe new members to conversations
    if ($g.PSObject.Properties.Name -contains 'AutoSubscribeNewMembers' -and $g.AutoSubscribeNewMembers -eq $true) {
        $groupFindings += [PSCustomObject]@{
            Id = "GT-015"
            Title = "Group auto-subscribes new members to all conversations"
            Severity = "Low"
            Category = "Data Management"
            Details = "Every new member automatically receives all inbound group email in their personal inbox — widens blast radius of malicious messages and complicates offboarding."
            Remediation = "Admin center > Teams &amp; Groups > Group > Settings > uncheck 'Send copies of team emails and events to team members' inboxes'."
        }
    }

    if ($groupFindings.Count -gt 0) {
        $findings += [PSCustomObject]@{
            EntityType    = "Group"
            EntityId      = $g.Id
            EntityName    = $g.DisplayName
            EntityUrl     = "https://admin.microsoft.com/adminportal/home#/groups/:/GroupDetailsV3/$($g.Id)/General"
            AdminUrl      = "https://admin.microsoft.com/adminportal/home#/groups/:/GroupDetailsV3/$($g.Id)/General"
            GuestCount    = $g.GuestCount
            MemberCount   = $g.MemberCount
            Visibility    = $g.Visibility
            Findings      = $groupFindings
        }
    }
}

# Check 12: Sensitivity label missing on Team
foreach ($t in @($teams)) {
    $teamFindings = @()
    if (-not $t.HasSensitivityLabel) {
        $teamFindings += [PSCustomObject]@{
            Id = "GT-012"
            Title = "Team missing sensitivity label"
            Severity = "Medium"
            Details = "Team has no assigned sensitivity label."
            Remediation = "Apply an appropriate sensitivity label. This also governs Teams guest access and external sharing defaults."
        }
        if ($t.GuestCount -gt 0) {
            $teamFindings += [PSCustomObject]@{
                Id = "GT-013"
                Title = "Team has guests and no sensitivity label"
                Severity = "Medium"
                Details = "$($t.GuestCount) guest(s) in a team without any sensitivity label."
                Remediation = "Apply a sensitivity label; review guest membership."
            }
        }
    }

    # Check GT-016 (Data Management): members can edit their own messages
    if ($t.PSObject.Properties.Name -contains 'AllowUserEditMessages' -and $t.AllowUserEditMessages -eq $true) {
        $teamFindings += [PSCustomObject]@{
            Id = "GT-016"
            Title = "Team allows members to edit their sent messages"
            Severity = "Low"
            Category = "Data Management"
            Details = "A compromised user (or a malicious insider) can silently rewrite past messages to remove phishing links or alter instructions — breaks the forensic evidence trail."
            Remediation = "Teams admin center > Messaging policies > set 'Owners can delete sent messages' / 'Users can edit sent messages' to Off for sensitive teams."
        }
    }

    # Check GT-017 (Data Management): members can delete their own messages
    if ($t.PSObject.Properties.Name -contains 'AllowUserDeleteMessages' -and $t.AllowUserDeleteMessages -eq $true) {
        $teamFindings += [PSCustomObject]@{
            Id = "GT-017"
            Title = "Team allows members to delete their sent messages"
            Severity = "Low"
            Category = "Data Management"
            Details = "A compromised user can remove evidence of phishing delivery or lateral movement inside Teams channels — hinders incident response investigation."
            Remediation = "Teams admin center > Messaging policies > set 'Users can delete sent messages' to Off for sensitive teams."
        }
    }

    if ($teamFindings.Count -gt 0) {
        $findings += [PSCustomObject]@{
            EntityType    = "Team"
            EntityId      = $t.Id
            EntityName    = $t.DisplayName
            EntityUrl     = "https://admin.microsoft.com/adminportal/home#/groups/:/GroupDetailsV3/$($t.Id)/General"
            AdminUrl      = "https://admin.microsoft.com/adminportal/home#/groups/:/GroupDetailsV3/$($t.Id)/General"
            GuestCount    = $t.GuestCount
            MemberCount   = $t.MemberCount
            Visibility    = $t.Visibility
            Findings      = $teamFindings
        }
    }
}

# Check 14 (default library sensitivity label) — placeholder
# Proper detection requires iterating each site's /drives/{id}/list columns in Graph which is expensive.
# Deferred to v2. Documented in README.

$result = @{
    Findings = $findings
    Summary  = @{
        GroupsScanned     = $groups.Count
        TeamsScanned      = $teams.Count
        EntitiesWithFindings = $findings.Count
        HighFindings      = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'High' }).Count
        MediumFindings    = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'Medium' }).Count
    }
    Timestamp = (Get-Date).ToString("o")
}

$result | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "  Groups/Teams analysis complete: $($findings.Count) entities with findings" -ForegroundColor Green
