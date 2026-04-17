[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DataDir,
    [Parameter(Mandatory)]
    [string]$OutputPath,
    [int]$ExcessiveExternalThreshold = 10,
    [int]$InactiveDaysThreshold = 365
)

$ErrorActionPreference = "Stop"

$sites          = Get-Content (Join-Path $DataDir "sites.json") | ConvertFrom-Json
$tenantBaseline = Get-Content (Join-Path $DataDir "tenant-baseline.json") | ConvertFrom-Json
$exportCtx      = if (Test-Path (Join-Path $DataDir "export-context.json")) {
                    Get-Content (Join-Path $DataDir "export-context.json") | ConvertFrom-Json
                  } else { $null }

# SharePoint admin center URL (sites management view). Specific-site deep link isn't
# supported by the admin center URL schema, so we link to the Active Sites list where
# the admin can filter by URL or name.
$adminBaseUrl = $null
if ($exportCtx -and $exportCtx.TenantDomain) {
    # Infer admin host from tenant: contoso.onmicrosoft.com -> contoso-admin.sharepoint.com
    $tenantPart = ($exportCtx.TenantDomain -replace '\.onmicrosoft\.com$','' -replace '\..*$','')
    $adminBaseUrl = "https://$tenantPart-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/siteManagement"
}

# Define sharing capability severity order — higher = more permissive
$capabilityRank = @{
    'Disabled'                    = 0
    'ExistingExternalUserSharingOnly' = 1
    'ExternalUserSharingOnly'     = 2
    'ExternalUserAndGuestSharing' = 3   # Anyone-link capable
}

function Get-CapRank($cap) {
    if ($capabilityRank.ContainsKey("$cap")) { return $capabilityRank["$cap"] }
    return -1
}

$baseRank = Get-CapRank $tenantBaseline.SharingCapability
$findings = @()
$now = Get-Date

foreach ($s in $sites) {
    $siteFindings = @()
    $capRank = Get-CapRank $s.SharingCapability

    # Check 1: Publicly accessible (Anyone-link)
    if ($s.SharingCapability -eq 'ExternalUserAndGuestSharing') {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-001"
            Title = "Publicly accessible site (Anyone links enabled)"
            Severity = "High"
            Details = "Sharing capability is 'ExternalUserAndGuestSharing' — anyone with the link can access."
            Remediation = "Reduce sharing on this site to 'ExternalUserSharingOnly' or tighter unless public access is intentional."
        }
    }

    # Check 2: Sites allow Anonymous access (same as Anyone links — covered by Check 1)
    # Distinct check: site has guest links without authentication — can't distinguish cleanly from Check 1
    # so we fold this into Check 1. (Not creating a duplicate finding.)

    # Check 3: Excessive external sharing (count-based)
    if ($null -ne $s.ExternalUserCount -and $s.ExternalUserCount -gt $ExcessiveExternalThreshold) {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-003"
            Title = "Excessive external users on site"
            Severity = "High"
            Details = "$($s.ExternalUserCount) external users have access (threshold: $ExcessiveExternalThreshold)."
            Remediation = "Review external users and remove those who no longer need access. Consider access reviews."
        }
    }

    # Check 4: Site-level sharing more permissive than tenant baseline
    if ($baseRank -ge 0 -and $capRank -gt $baseRank) {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-004"
            Title = "Site sharing more permissive than tenant baseline"
            Severity = "Medium"
            Details = "Site sharing is '$($s.SharingCapability)' but tenant baseline is '$($tenantBaseline.SharingCapability)'."
            Remediation = "Align site sharing to tenant baseline unless a documented exception exists."
        }
    }

    # Check 5: Inactive site (no content changes > threshold)
    if ($s.LastContentModifiedDate) {
        try {
            $mod = [datetime]::Parse($s.LastContentModifiedDate, [System.Globalization.CultureInfo]::InvariantCulture)
            $daysSince = [int]($now - $mod).TotalDays
            if ($daysSince -gt $InactiveDaysThreshold -and $s.LockState -ne 'Archived' -and $s.LockState -ne 'NoAccess') {
                $siteFindings += [PSCustomObject]@{
                    Id = "SP-005"
                    Title = "Inactive site (no content changes > $InactiveDaysThreshold days)"
                    Severity = "Medium"
                    Details = "Last content modified $daysSince days ago. Site is not archived."
                    Remediation = "Consider archiving or removing the site via SharePoint site lifecycle management."
                }
            }
        } catch {}
    }

    # Check 6: Site ownership — only flag non-group-connected sites with no primary admin.
    # Group-connected sites have owners managed on the M365 group, so PnP's Owner field
    # is typically empty for them; skipping those avoids false positives.
    $isGroupConnected = ($s.GroupId -and $s.GroupId -ne '' -and $s.GroupId -ne '00000000-0000-0000-0000-000000000000')
    $ownerIsSystem    = ($s.Owner -match 'SHAREPOINT\\system' -or $s.Owner -match 'c:0\(\.s\|true')
    if (-not $isGroupConnected -and ($ownerIsSystem -or -not $s.Owner -or $s.Owner -eq '')) {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-006"
            Title = "Site has no primary admin assigned"
            Severity = "Medium"
            Details = "Non-group site has no primary admin or only a system account. Owner: '$($s.Owner)'."
            Remediation = "Assign at least one explicit site collection admin for operational continuity."
        }
    }

    # Check 7: Sensitivity label missing on site
    if (-not $s.SensitivityLabel -or $s.SensitivityLabel -eq '') {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-007"
            Title = "Site missing sensitivity label"
            Severity = "Medium"
            Details = "No sensitivity label assigned to this site."
            Remediation = "Apply an appropriate sensitivity label (Public / Internal / Confidential / Highly Confidential)."
        }
    }

    # Check 8: Direct access vs group-based — heuristic not reliable from Get-SPOSite alone.
    # Proper detection requires Get-SPOUser per site and categorizing by principal type, which is costly.
    # For v1, flag only if ExternalUserCount > 0 AND site has no group association (GroupId empty).
    if ($null -ne $s.ExternalUserCount -and $s.ExternalUserCount -gt 0 -and (-not $s.GroupId -or $s.GroupId -eq '')) {
        $siteFindings += [PSCustomObject]@{
            Id = "SP-008"
            Title = "External users on non-group site (likely direct grants)"
            Severity = "Medium"
            Details = "Site has $($s.ExternalUserCount) external users and is not backed by an M365 group — access likely granted to individuals rather than groups."
            Remediation = "Prefer group-based access. Connect the site to an M365 group where possible."
        }
    }

    if ($siteFindings.Count -gt 0) {
        $findings += [PSCustomObject]@{
            EntityType    = "Site"
            EntityId      = $s.Url
            EntityName    = $s.Title
            EntityUrl     = $s.Url                          # direct link to the site
            AdminUrl      = if ($adminBaseUrl) { $adminBaseUrl } else { $null }  # SP admin center
            StorageBytes  = $s.StorageUsageCurrent
            SharingCap    = $s.SharingCapability
            ExternalCount = $s.ExternalUserCount
            Findings      = $siteFindings
        }
    }
}

$result = @{
    Findings = $findings
    Summary  = @{
        SitesScanned        = $sites.Count
        SitesWithFindings   = $findings.Count
        HighFindings        = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'High' }).Count
        MediumFindings      = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'Medium' }).Count
    }
    Timestamp = (Get-Date).ToString("o")
}

$result | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "  Site analysis complete: $($findings.Count) sites with findings" -ForegroundColor Green
