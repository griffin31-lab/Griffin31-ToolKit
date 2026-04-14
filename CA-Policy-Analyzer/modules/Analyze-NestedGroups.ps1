[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExportDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

# Load data
$groups = Get-Content "$ExportDir/SecurityGroups.json" | ConvertFrom-Json
$policies = Get-Content "$ExportDir/ConditionalAccessPolicies.json" | ConvertFrom-Json

$groupMems = @{}
if (Test-Path "$ExportDir/SecurityGroupMemberships.json") {
    $gm = Get-Content "$ExportDir/SecurityGroupMemberships.json" | ConvertFrom-Json
    foreach ($g in $gm) {
        if ($g.GroupId -and $g.Members) {
            $groupMems[$g.GroupId] = $g.Members.Id
        }
    }
}

$groupMap = @{}
foreach ($g in $groups) { $groupMap[$g.Id] = $g.DisplayName }

function Get-GroupMembersRecursive($groupId, $seen=@{}) {
    $memberIds = @()
    if ($seen.ContainsKey($groupId)) { return $memberIds }
    $seen[$groupId] = $true
    if ($groupMems[$groupId]) {
        foreach ($memberId in $groupMems[$groupId]) {
            $memberIds += $memberId
            if ($groupMap.ContainsKey($memberId)) {
                $memberIds += Get-GroupMembersRecursive $memberId $seen
            }
        }
    }
    return $memberIds | Select-Object -Unique
}

# Build a set of group IDs that contain nested groups (direct children that are themselves groups)
$nestedGroupIds = @{}
$nestedGroups = @()
foreach ($group in $groups) {
    $directMembers = @($groupMems[$group.Id] | Where-Object { $_ })
    $nestedChildren = $directMembers | Where-Object { $groupMap.ContainsKey($_) }
    if ($nestedChildren.Count -gt 0) {
        $nestedGroupIds[$group.Id] = $true
        $allMembers = Get-GroupMembersRecursive $group.Id
        $nestedGroups += [PSCustomObject]@{
            GroupId        = $group.Id
            GroupName      = $group.DisplayName
            DirectMembers  = $directMembers.Count
            TotalMembers   = @($allMembers).Count
            NestedChildren = @($nestedChildren | ForEach-Object { [PSCustomObject]@{ Id = $_; Name = $groupMap[$_] } })
        }
    }
}

# ── Blind-spot detection: CA policies that include/exclude a group containing nested groups ──
$policiesUsing = @()
foreach ($p in $policies) {
    $refs = @()
    foreach ($gid in @($p.Conditions.Users.IncludeGroups)) {
        if ($gid -and $nestedGroupIds.ContainsKey($gid)) {
            $refs += [PSCustomObject]@{ Assignment = "Include"; GroupId = $gid; GroupName = $groupMap[$gid] }
        }
    }
    foreach ($gid in @($p.Conditions.Users.ExcludeGroups)) {
        if ($gid -and $nestedGroupIds.ContainsKey($gid)) {
            $refs += [PSCustomObject]@{ Assignment = "Exclude"; GroupId = $gid; GroupName = $groupMap[$gid] }
        }
    }
    foreach ($r in $refs) {
        $policiesUsing += [PSCustomObject]@{
            PolicyId    = $p.id
            PolicyName  = $p.displayName
            PolicyState = $p.state
            Assignment  = $r.Assignment
            GroupId     = $r.GroupId
            GroupName   = $r.GroupName
        }
    }
}

$result = @{
    NestedGroups              = $nestedGroups
    PoliciesUsingNestedGroups = $policiesUsing
    TotalGroupsWithNesting    = $nestedGroups.Count
    TotalPolicyAssignments    = $policiesUsing.Count
    TotalGroups               = $groups.Count
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$result | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "Nested group analysis complete: $($nestedGroups.Count) nested groups, $($policiesUsing.Count) CA assignment(s) affected" -ForegroundColor Green
