param(
    [string]$OutputFolder = "data",
    [string]$UserPrincipalName,
    [string]$TenantDomain,
    [string]$SelectionChoice  # No default value - only set when explicitly passed
)

Write-Host "`n=== Microsoft Graph Data Export Tool ===" -ForegroundColor Cyan
Write-Host "This will export data to the '$OutputFolder' directory" -ForegroundColor Yellow

# Show the target tenant if provided
if ($TenantDomain) {
    Write-Host "Target Tenant: $TenantDomain" -ForegroundColor Cyan
} elseif ($UserPrincipalName) {
    Write-Host "Target User: $UserPrincipalName" -ForegroundColor Cyan
}

# Only show menu in interactive mode
if (-not $SelectionChoice) {
    Write-Host ""
    Write-Host "Select components to export:" -ForegroundColor Yellow
    Write-Host "1. Conditional Access Policies" -ForegroundColor Green
    Write-Host "2. Named Locations" -ForegroundColor Green
    Write-Host "3. Security Groups & Memberships" -ForegroundColor Green
    Write-Host "4. Users (Members & Guests)" -ForegroundColor Green
    Write-Host "5. Directory Roles" -ForegroundColor Green
    Write-Host "6. Role Assignments" -ForegroundColor Green
    Write-Host "7. Service Principals" -ForegroundColor Green
    Write-Host "8. Devices" -ForegroundColor Green
    Write-Host "9. All Components" -ForegroundColor Cyan
    Write-Host "0. Exit" -ForegroundColor Red
}

# Use SelectionChoice parameter if provided, otherwise prompt user
if ($SelectionChoice) {
    $choices = $SelectionChoice
    Write-Host "`nUsing selection: $choices (non-interactive mode)" -ForegroundColor Green
} else {
    $choices = Read-Host "`nEnter choices (comma-separated, e.g., 1,3,5)"
}
if ($choices -eq "0") { Write-Host "Exiting..."; exit }

$selectedComponents = if ($choices -eq "9") { @("1","2","3","4","5","6","7","8") } else { $choices -split "," | ForEach-Object { $_.Trim() } }

# Create output directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Support both absolute and relative paths
if ([System.IO.Path]::IsPathRooted($OutputFolder)) {
    $outputPath = $OutputFolder
} else {
    $outputPath = Join-Path $scriptDir $OutputFolder
}

New-Item -Path $outputPath -ItemType Directory -Force | Out-Null

try {
    Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow

    # Find the highest Microsoft.Graph version that is installed for ALL required
    # submodules. Pinning every submodule to the same version prevents assembly
    # load conflicts when multiple versions coexist.
    $graphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.DeviceManagement'
    )

    $versionSets = foreach ($m in $graphModules) {
        $versions = Get-Module -ListAvailable -Name $m | ForEach-Object { $_.Version }
        if (-not $versions) {
            throw "Module '$m' is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
        }
        ,@($versions)
    }
    $commonVersions = $versionSets[0]
    for ($i = 1; $i -lt $versionSets.Count; $i++) {
        $commonVersions = $commonVersions | Where-Object { $versionSets[$i] -contains $_ }
    }
    if (-not $commonVersions) {
        throw "No common Microsoft.Graph version found across required submodules. Reinstall with: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
    }
    $graphVersion = ($commonVersions | Sort-Object -Descending | Select-Object -First 1).ToString()
    Write-Host "Using Microsoft.Graph version: $graphVersion" -ForegroundColor Gray

    foreach ($m in $graphModules) {
        Import-Module $m -RequiredVersion $graphVersion -Force -ErrorAction Stop
    }
    
    # Use tenant domain if provided, otherwise extract from UPN
    $targetTenant = $null
    if ($TenantDomain) {
        $targetTenant = $TenantDomain
        Write-Host "Connecting to tenant: $TenantDomain" -ForegroundColor Gray
    } elseif ($UserPrincipalName) {
        $targetTenant = $UserPrincipalName.Split('@')[1]
        Write-Host "Connecting to tenant: $targetTenant (from UPN: $UserPrincipalName)" -ForegroundColor Gray
    }
    
    if ($targetTenant) {
        Connect-MgGraph -Scopes "Policy.Read.All","Group.Read.All","User.Read.All","Directory.Read.All","Application.Read.All","Device.Read.All","RoleManagement.Read.Directory" -TenantId $targetTenant -NoWelcome
    } else {
        Connect-MgGraph -Scopes "Policy.Read.All","Group.Read.All","User.Read.All","Directory.Read.All","Application.Read.All","Device.Read.All","RoleManagement.Read.Directory" -NoWelcome
    }
    
    Write-Host "Connected successfully!" -ForegroundColor Green

    if ($selectedComponents -contains "1") {
        Write-Host "`nExporting Conditional Access Policies..." -ForegroundColor Yellow
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
        $caPolicies | ConvertTo-Json -Depth 20 | Out-File "$outputPath\ConditionalAccessPolicies.json" -Encoding UTF8
        Write-Host "Exported $($caPolicies.Count) CA policies" -ForegroundColor Green
    }

    if ($selectedComponents -contains "2") {
        Write-Host "Exporting Named Locations..." -ForegroundColor Yellow
        $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All
        $namedLocations | ConvertTo-Json -Depth 10 | Out-File "$outputPath\NamedLocations.json" -Encoding UTF8
        Write-Host "Exported $($namedLocations.Count) named locations" -ForegroundColor Green
    }

    if ($selectedComponents -contains "3") {
        Write-Host "Exporting Security Groups..." -ForegroundColor Yellow
        $secGroups = Get-MgGroup -Filter "securityEnabled eq true and mailEnabled eq false" -All
        $simpleGroups = $secGroups | Select-Object Id, DisplayName, Description
        $simpleGroups | ConvertTo-Json -Depth 5 | Out-File "$outputPath\SecurityGroups.json" -Encoding UTF8
        Write-Host "Exported $($simpleGroups.Count) security groups" -ForegroundColor Green
        
        Write-Host "Exporting Group Memberships (parallel processing)..." -ForegroundColor Yellow
        
        $groupMemberships = $simpleGroups | ForEach-Object -Parallel {
            try {
                $members = Get-MgGroupMember -GroupId $_.Id -All | Select-Object Id, @{Name="ODataType";Expression={$_."@odata.type"}}
                [PSCustomObject]@{
                    GroupId = $_.Id
                    GroupDisplayName = $_.DisplayName
                    Members = $members
                }
            } catch {
                Write-Warning "Failed to get members for group $($_.DisplayName): $($_.Exception.Message)"
            }
        } -ThrottleLimit 10
        
        $groupMemberships | ConvertTo-Json -Depth 5 | Out-File "$outputPath\SecurityGroupMemberships.json" -Encoding UTF8
        Write-Host "Exported group memberships for $($groupMemberships.Count) groups"
    }

    if ($selectedComponents -contains "4") {
        Write-Host "Exporting Members..." -ForegroundColor Yellow
        $members = Get-MgUser -Filter "userType eq 'Member'" -All
        $simpleMembers = $members | Select-Object Id, DisplayName, UserPrincipalName, AccountEnabled, JobTitle, Department, Mail, UserType, CreationDateTime
        $simpleMembers | ConvertTo-Json -Depth 5 | Out-File "$outputPath\Members.json" -Encoding UTF8
        Write-Host "Exported $($simpleMembers.Count) member users" -ForegroundColor Green
        
        Write-Host "Exporting Guests..." -ForegroundColor Yellow
        $guests = Get-MgUser -Filter "userType eq 'Guest'" -All
        $simpleGuests = $guests | Select-Object Id, DisplayName, UserPrincipalName, AccountEnabled, JobTitle, Department, Mail, UserType, CreationDateTime
        $simpleGuests | ConvertTo-Json -Depth 5 | Out-File "$outputPath\Guests.json" -Encoding UTF8
        Write-Host "Exported $($simpleGuests.Count) guest users" -ForegroundColor Green
    }

    if ($selectedComponents -contains "5") {
        Write-Host "Exporting Directory Roles..." -ForegroundColor Yellow
        $roles = Get-MgDirectoryRole -All
        $simpleRoles = $roles | Select-Object Id, DisplayName, Description
        $simpleRoles | ConvertTo-Json -Depth 5 | Out-File "$outputPath\DirectoryRoles.json" -Encoding UTF8
        Write-Host "Exported $($roles.Count) directory roles" -ForegroundColor Green
    }

    if ($selectedComponents -contains "6") {
        Write-Host "Exporting Role Assignments (Focus: Global Administrator)..." -ForegroundColor Yellow
        $roleAssignments = @()
        
        # Find Global Administrator role
        $globalAdminRole = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'"
        if ($globalAdminRole) {
            Write-Host "Found Global Administrator role with ID: $($globalAdminRole.Id)"
            
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id -All
            Write-Host "Found $($members.Count) Global Admin members"
            
            foreach ($member in $members) {
                try {
                    $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                    if ($user) {
                        Write-Host "- $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Cyan
                        
                        $roleAssignments += [PSCustomObject]@{
                            RoleDisplayName = "Global Administrator"
                            UserId = $user.Id
                            UserDisplayName = $user.DisplayName
                            UserPrincipalName = $user.UserPrincipalName
                            UserType = $user.UserType
                            AccountEnabled = $user.AccountEnabled
                        }
                    }
                } catch {
                    Write-Warning "Failed to get user details for member $($member.Id): $($_.Exception.Message)"
                }
            }
        } else {
            Write-Warning "Global Administrator role not found"
        }
        
        $roleAssignments | ConvertTo-Json -Depth 5 | Out-File "$outputPath\DirectoryRoleAssignments.json" -Encoding UTF8
        Write-Host "Exported $($roleAssignments.Count) Global Administrator assignments" -ForegroundColor Green
    }

    if ($selectedComponents -contains "7") {
        Write-Host "Exporting Service Principals..." -ForegroundColor Yellow
        $apps = Get-MgServicePrincipal -All
        $simpleApps = $apps | Select-Object Id, DisplayName, AppId, PublisherName, CreatedDateTime
        $simpleApps | ConvertTo-Json -Depth 5 | Out-File "$outputPath\ServicePrincipals.json" -Encoding UTF8
        Write-Host "Exported $($simpleApps.Count) service principals" -ForegroundColor Green
    }

    if ($selectedComponents -contains "8") {
        Write-Host "Exporting Devices..." -ForegroundColor Yellow
        $devices = Get-MgDevice -All
        $simpleDevices = $devices | Select-Object Id, DisplayName, DeviceId, OperatingSystem, DeviceOwnership, AccountEnabled, IsCompliant, IsManaged, CreatedDateTime
        $simpleDevices | ConvertTo-Json -Depth 5 | Out-File "$outputPath\Devices.json" -Encoding UTF8
        Write-Host "Exported $($simpleDevices.Count) devices" -ForegroundColor Green
    }

    Write-Host "`nExport complete. Data saved in: $outputPath" -ForegroundColor Green
    Write-Host "You can now run: .\CA-Gap-Analysis.ps1" -ForegroundColor Cyan

} catch {
    Write-Error "Export failed: $($_.Exception.Message)"
    exit 1
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
