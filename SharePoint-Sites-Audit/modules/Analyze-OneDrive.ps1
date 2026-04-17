[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DataDir,
    [Parameter(Mandatory)]
    [string]$OutputPath,
    [int]$ExcessiveExternalThreshold = 5
)

$ErrorActionPreference = "Stop"

$oneDrives      = Get-Content (Join-Path $DataDir "onedrives.json") | ConvertFrom-Json
$tenantBaseline = Get-Content (Join-Path $DataDir "tenant-baseline.json") | ConvertFrom-Json

$capabilityRank = @{
    'Disabled' = 0
    'ExistingExternalUserSharingOnly' = 1
    'ExternalUserSharingOnly' = 2
    'ExternalUserAndGuestSharing' = 3
}
function Get-CapRank($cap) {
    if ($capabilityRank.ContainsKey("$cap")) { return $capabilityRank["$cap"] }
    return -1
}
$baseRank = Get-CapRank $tenantBaseline.SharingCapability

$findings = @()
foreach ($od in $oneDrives) {
    $odFindings = @()

    # Check 9: Excessive external sharing (OneDrive)
    if ($null -ne $od.ExternalUserCount -and $od.ExternalUserCount -gt $ExcessiveExternalThreshold) {
        $odFindings += [PSCustomObject]@{
            Id = "OD-009"
            Title = "Excessive external users on OneDrive"
            Severity = "High"
            Details = "$($od.ExternalUserCount) external users (threshold: $ExcessiveExternalThreshold) have access to this OneDrive."
            Remediation = "Review and revoke external access that is no longer needed. Consider a user access review."
        }
    }

    # Check 10: OneDrive sharing more permissive than tenant baseline
    $capRank = Get-CapRank $od.SharingCapability
    if ($baseRank -ge 0 -and $capRank -gt $baseRank) {
        $odFindings += [PSCustomObject]@{
            Id = "OD-010"
            Title = "OneDrive sharing more permissive than tenant baseline"
            Severity = "Medium"
            Details = "OneDrive sharing is '$($od.SharingCapability)' but tenant baseline is '$($tenantBaseline.SharingCapability)'."
            Remediation = "Tighten OneDrive sharing to match the tenant baseline unless an exception is documented."
        }
    }

    if ($odFindings.Count -gt 0) {
        $findings += [PSCustomObject]@{
            EntityType    = "OneDrive"
            EntityId      = $od.Url
            EntityName    = if ($od.Owner) { $od.Owner } else { $od.Title }
            EntityUrl     = $od.Url
            StorageBytes  = $od.StorageUsageCurrent
            SharingCap    = $od.SharingCapability
            ExternalCount = $od.ExternalUserCount
            Findings      = $odFindings
        }
    }
}

$result = @{
    Findings = $findings
    Summary  = @{
        OneDrivesScanned      = $oneDrives.Count
        OneDrivesWithFindings = $findings.Count
        HighFindings          = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'High' }).Count
        MediumFindings        = ($findings | ForEach-Object { $_.Findings } | Where-Object { $_.Severity -eq 'Medium' }).Count
    }
    Timestamp = (Get-Date).ToString("o")
}

$result | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "  OneDrive analysis complete: $($findings.Count) OneDrives with findings" -ForegroundColor Green
