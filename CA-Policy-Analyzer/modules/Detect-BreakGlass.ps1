[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExportDir,
    [Parameter(Mandatory)]
    [string]$OutputPath
)

# Load data
$policies = Get-Content "$ExportDir\ConditionalAccessPolicies.json" | ConvertFrom-Json
$members = Get-Content "$ExportDir\Members.json" | ConvertFrom-Json
$guests = Get-Content "$ExportDir\Guests.json" | ConvertFrom-Json
$allUsers = @($members + $guests)

$groupMems = @{}
if (Test-Path "$ExportDir\SecurityGroupMemberships.json") {
    $gm = Get-Content "$ExportDir\SecurityGroupMemberships.json" | ConvertFrom-Json
    foreach ($g in $gm) { 
        if ($g.GroupId -and $g.Members) {
            $groupMems[$g.GroupId] = $g.Members.Id 
        }
    }
}

$breakGlassUsers = @()

if (Test-Path "$ExportDir\DirectoryRoleAssignments.json") {
    $roleAssignments = Get-Content "$ExportDir\DirectoryRoleAssignments.json" | ConvertFrom-Json
    $globalAdminUserIds = ($roleAssignments | Where-Object { $_.RoleDisplayName -eq "Global Administrator" }).UserId
    
    # Find users excluded from ALL enabled policies
    $enabledPolicies = $policies | Where-Object { $_.state -eq 'enabled' }
    
    foreach ($user in $allUsers) {
        $isGlobalAdmin = $globalAdminUserIds -contains $user.Id
        if (-not $isGlobalAdmin) { continue }
        
        $excludedFromAllPolicies = $true
        
        foreach ($policy in $enabledPolicies) {
            $includedUsers = @()
            $excludedUsers = @()
            
            # Handle includes
            if ($policy.Conditions.Users.IncludeUsers -contains "All") {
                $includedUsers = $allUsers.Id
            } else {
                $includedUsers += $policy.Conditions.Users.IncludeUsers
            }
            
            # Add group members
            foreach ($groupId in $policy.Conditions.Users.IncludeGroups) {
                if ($groupMems.ContainsKey($groupId)) {
                    $includedUsers += $groupMems[$groupId]
                }
            }
            
            # Handle excludes
            $excludedUsers += $policy.Conditions.Users.ExcludeUsers
            foreach ($groupId in $policy.Conditions.Users.ExcludeGroups) {
                if ($groupMems.ContainsKey($groupId)) {
                    $excludedUsers += $groupMems[$groupId]
                }
            }
            
            # Check if user is included and not excluded
            if (($includedUsers -contains $user.Id) -and ($excludedUsers -notcontains $user.Id)) {
                $excludedFromAllPolicies = $false
                break
            }
        }
        
        if ($excludedFromAllPolicies) {
            $breakGlassUsers += $user
        }
    }
    
    $detectionMethod = "role-based"
} else {
    # Fallback to name-based detection
    $breakGlassUsers = $allUsers | Where-Object { $_.DisplayName -match "Break Glass|Break2|Break 2|breakglass|break glass" }
    $detectionMethod = "name-based"
}

$result = @{
    BreakGlassUsers = $breakGlassUsers | Select-Object Id, DisplayName, UserPrincipalName
    Count = $breakGlassUsers.Count
    DetectionMethod = $detectionMethod
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$result | ConvertTo-Json -Depth 3 | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Break-glass detection complete: Found $($breakGlassUsers.Count) accounts using $detectionMethod detection" -ForegroundColor Green
