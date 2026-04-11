#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Configure and Verify Application Permissions for Exchange Online Mailboxes

.DESCRIPTION
    Universal PowerShell script for managing and verifying Exchange Online application permissions.
    Supports all 13 Exchange application roles with flexible add/remove capabilities.
    Works with any mailbox type: shared mailboxes, user mailboxes, room/equipment resources.
    
    Features:
    - Complete role support (Mail, Calendar, Contacts, MailboxSettings, EWS)
    - Add/remove specific permissions
    - Smart connection management (reuses existing connections)
    - Role validation and duplicate prevention
    - Management scope naming improvements
    - Automatic Graph permission mapping
    - Works with all mailbox types (shared, user, resource, etc.)
    - Interactive mode when run without parameters
    - Verification mode (replaces Verify-AppPermissions.ps1)

.PARAMETER AppId
    Application (client) ID from Azure AD app registration

.PARAMETER TargetMailbox
    Email address of the target mailbox (works with shared mailboxes, user mailboxes, room/equipment resources, etc.)

.PARAMETER AppDisplayName
    Display name for the application (default: "Application")

.PARAMETER RolesToAdd
    Array of Exchange roles to assign (default: Mail.ReadWrite, Mail.Send, Calendars.ReadWrite)

.PARAMETER RolesToRemove
    Array of Exchange roles to remove

.PARAMETER RemoveAllRoles
    Remove all assigned roles for this application

.PARAMETER ListAvailableRoles
    Display all available Exchange application roles with descriptions

.PARAMETER ForceNewScope
    Attempt to create app-specific scope (will show Exchange limitation)

.PARAMETER RenameExistingScope
    Rename existing scope to generic name (fixes app-specific naming)

.PARAMETER VerifyOnly
    Verification mode - only displays current permissions without making changes (like Verify-AppPermissions.ps1)

.EXAMPLE
    .\Configure-AppPermissions.ps1
    # Interactive mode - prompts for all required information

.EXAMPLE
    .\Configure-AppPermissions.ps1 -VerifyOnly
    # Interactive verification mode - prompts for AppId and TargetMailbox, then shows current permissions

.EXAMPLE
    .\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "shared@domain.com" -VerifyOnly
    # Verification mode with parameters

.EXAMPLE
    .\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "shared@domain.com"
    
.EXAMPLE
    .\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "john.doe@domain.com" -RolesToAdd @("Application Mail.Read")
    
.EXAMPLE 
    .\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "conference.room1@domain.com" -RolesToAdd @("Application Calendars.ReadWrite")
    
.EXAMPLE
    .\Configure-AppPermissions.ps1 -ListAvailableRoles

.EXAMPLE
    .\Configure-AppPermissions.ps1 -AppId "your-app-id" -TargetMailbox "shared@domain.com" -RolesToAdd @("Application Mail.Read", "Application Calendars.ReadWrite")

.NOTES
    Author: Exchange Admin Script
    Version: 3.0 (Unified Script)
    Last Updated: September 2025
    Requires: ExchangeOnlineManagement module
    Requires: Exchange Online Administrator or Global Administrator permissions
#>

# Configure and Verify Application Permissions for Exchange Online Mailboxes
# Universal script - works with any mailbox type (shared, user, resource, etc.)
# Supports all Exchange application roles with add/remove capabilities
# Includes verification mode (replaces Verify-AppPermissions.ps1)

param(
    [Parameter(Mandatory=$false)]
    [string]$AppId,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetMailbox,
    
    [Parameter(Mandatory=$false)]
    [string]$AppDisplayName = "Application",
    
    [Parameter(Mandatory=$false)]
    [string[]]$RolesToAdd = @(),
    
    [Parameter(Mandatory=$false)]
    [string[]]$RolesToRemove = @(),
    
    [Parameter(Mandatory=$false)]
    [switch]$ListAvailableRoles,
    
    [Parameter(Mandatory=$false)]
    [switch]$RemoveAllRoles,
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceNewScope,
    
    [Parameter(Mandatory=$false)]
    [switch]$RenameExistingScope,
    
    [Parameter(Mandatory=$false)]
    [switch]$VerifyOnly
)

# Define all available Exchange application roles
$AvailableRoles = @{
    "Application Mail.Read" = @{
        Protocol = "MS Graph"
        Permissions = "Mail.Read"
        Description = "Allows the app to read email in all mailboxes without a signed-in user."
    }
    "Application Mail.ReadBasic" = @{
        Protocol = "MS Graph" 
        Permissions = "Mail.ReadBasic"
        Description = "Allows the app to read email except the body, previewBody, attachments, and any extended properties in all mailboxes without a signed-in user"
    }
    "Application Mail.ReadWrite" = @{
        Protocol = "MS Graph"
        Permissions = "Mail.ReadWrite"
        Description = "Allows the app to create, read, update, and delete email in all mailboxes without a signed-in user. Doesn't include permission to send mail."
    }
    "Application Mail.Send" = @{
        Protocol = "MS Graph"
        Permissions = "Mail.Send"
        Description = "Allows the app to send mail as any user without a signed-in user."
    }
    "Application MailboxSettings.Read" = @{
        Protocol = "MS Graph"
        Permissions = "MailboxSettings.Read"
        Description = "Allows the app to read user's mailbox settings in all mailboxes without a signed-in user."
    }
    "Application MailboxSettings.ReadWrite" = @{
        Protocol = "MS Graph"
        Permissions = "MailboxSettings.ReadWrite"
        Description = "Allows the app to create, read, update, and delete user's mailbox settings in all mailboxes without a signed-in user."
    }
    "Application Calendars.Read" = @{
        Protocol = "MS Graph"
        Permissions = "Calendars.Read"
        Description = "Allows the app to read events of all calendars without a signed-in user."
    }
    "Application Calendars.ReadWrite" = @{
        Protocol = "MS Graph"
        Permissions = "Calendars.ReadWrite"
        Description = "Allows the app to create, read, update, and delete events of all calendars without a signed-in user."
    }
    "Application Contacts.Read" = @{
        Protocol = "MS Graph"
        Permissions = "Contacts.Read"
        Description = "Allows the app to read all contacts in all mailboxes without a signed-in user."
    }
    "Application Contacts.ReadWrite" = @{
        Protocol = "MS Graph"
        Permissions = "Contacts.ReadWrite"
        Description = "Allows the app to create, read, update, and delete all contacts in all mailboxes without a signed-in user."
    }
    "Application Mail Full Access" = @{
        Protocol = "MS Graph"
        Permissions = "Mail.ReadWrite, Mail.Send"
        Description = "Allows the app to create, read, update, and delete email in all mailboxes and send mail as any user without a signed-in user."
    }
    "Application Exchange Full Access" = @{
        Protocol = "MS Graph"
        Permissions = "Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite, Calendars.ReadWrite, Contacts.ReadWrite"
        Description = "Full access to email, mailbox settings, calendars, and contacts without a signed-in user."
    }
    "Application EWS.AccessAsApp" = @{
        Protocol = "EWS"
        Permissions = "EWS.AccessAsApp"
        Description = "Allows the app to use Exchange Web Services with full access to all mailboxes."
    }
}

# Display available roles if requested
if ($ListAvailableRoles) {
    Write-Host "=== Available Exchange Application Roles ===" -ForegroundColor Green
    Write-Host ""
    foreach ($Role in $AvailableRoles.Keys | Sort-Object) {
        $Info = $AvailableRoles[$Role]
        Write-Host "Role: $Role" -ForegroundColor Cyan
        Write-Host "  Protocol: $($Info.Protocol)" -ForegroundColor White
        Write-Host "  Permissions: $($Info.Permissions)" -ForegroundColor White
        Write-Host "  Description: $($Info.Description)" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "Usage Examples:" -ForegroundColor Yellow
    Write-Host "  # Add specific roles:"
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -RolesToAdd @('Application Mail.Read', 'Application Calendars.ReadWrite')" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  # Remove specific roles:"
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -RolesToRemove @('Application Mail.Send')" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  # Full access (all permissions):"
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -RolesToAdd @('Application Exchange Full Access')" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  # Force app-specific scope creation:"
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -ForceNewScope" -ForegroundColor Cyan
    Write-Host ""  
    Write-Host "  # Rename existing scope to generic name:"
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -RenameExistingScope" -ForegroundColor Cyan
    return
}

# ===== INTERACTIVE MODE =====
# Prompt for missing required parameters when not provided via command line
if (-not $AppId -or -not $TargetMailbox -or ((-not $VerifyOnly) -and ($RolesToAdd.Count -eq 0 -and $RolesToRemove.Count -eq 0 -and -not $RemoveAllRoles))) {
    if ($VerifyOnly) {
        Write-Host "=== Interactive Verification Mode ===" -ForegroundColor Green
    } else {
        Write-Host "=== Interactive Configuration Mode ===" -ForegroundColor Green
    }
    Write-Host ""
    
    # Prompt for AppId if not provided
    if (-not $AppId) {
        do {
            $AppId = Read-Host "Enter Application ID (GUID from Azure AD app registration)"
            if (-not $AppId -or $AppId.Trim() -eq "") {
                Write-Host "❌ Application ID is required!" -ForegroundColor Red
            }
        } while (-not $AppId -or $AppId.Trim() -eq "")
        $AppId = $AppId.Trim()
    }
    
    # Prompt for TargetMailbox if not provided
    if (-not $TargetMailbox) {
        do {
            $TargetMailbox = Read-Host "Enter Target Mailbox email address (shared/user/resource mailbox)"
            if (-not $TargetMailbox -or $TargetMailbox.Trim() -eq "") {
                Write-Host "❌ Target Mailbox is required!" -ForegroundColor Red
            } elseif ($TargetMailbox -notmatch "^[^@]+@[^@]+\.[^@]+$") {
                Write-Host "❌ Please enter a valid email address!" -ForegroundColor Red
                $TargetMailbox = ""
            }
        } while (-not $TargetMailbox -or $TargetMailbox.Trim() -eq "")
        $TargetMailbox = $TargetMailbox.Trim()
    }
}
    
    # Show current permissions for context (like Verify-AppPermissions)
    if ($VerifyOnly) {
        # In verification mode, we skip the interactive role selection and go straight to verification
        Write-Host ""
        Write-Host "=== Verifying Application Permissions ===" -ForegroundColor Green
        Write-Host "App ID: $AppId" -ForegroundColor White
        Write-Host "Target Mailbox: $TargetMailbox" -ForegroundColor White
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "=== Current Permissions for $TargetMailbox ===" -ForegroundColor Cyan
    }
    
    try {
        # Import modules for verification
        Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
        
        # Connect to Exchange Online if needed
        $WasAlreadyConnected = $false
        try {
            $ExistingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($ExistingConnection -and $ExistingConnection.State -eq "Connected") {
                # Verify connection is still active by testing a simple command
                try {
                    Get-OrganizationConfig -ErrorAction Stop | Out-Null
                    Write-Host "✓ Using existing Exchange Online connection ($($ExistingConnection.UserPrincipalName))" -ForegroundColor Green
                    $WasAlreadyConnected = $true
                } catch {
                    Write-Host "⚠ Existing connection expired, reconnecting..." -ForegroundColor Yellow
                    $ExistingConnection = $null
                }
            }
            
            if (-not $ExistingConnection -or $ExistingConnection.State -ne "Connected") {
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false
                Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
            }
        } catch {
            Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
            Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
            if ($VerifyOnly) { exit 1 } else { return }
        }
        
        # Check service principal
        $ServicePrincipal = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
        if ($ServicePrincipal) {
            Write-Host "✓ Service Principal exists" -ForegroundColor Green
            $ObjectId = $ServicePrincipal.ObjectId
            
            # Check role assignments (using both AppId and ObjectId)
            Write-Host ""
            Write-Host "Role Assignments:" -ForegroundColor Yellow
            
            if ($VerifyOnly) {
                # In verify mode, show ALL assignments for this application
                $Assignments = Get-ManagementRoleAssignment | Where-Object {$_.RoleAssigneeName -eq $AppId -or $_.RoleAssigneeName -eq $ObjectId}
            } else {
                # In config mode, show only assignments for this specific mailbox
                $Assignments = Get-ManagementRoleAssignment | Where-Object {
                    ($_.RoleAssigneeName -eq $AppId -or $_.RoleAssigneeName -eq $ObjectId) -and 
                    $_.CustomResourceScope -and
                    (Get-ManagementScope -Identity $_.CustomResourceScope -ErrorAction SilentlyContinue).RecipientFilter -eq "PrimarySmtpAddress -eq '$TargetMailbox'"
                }
            }
            
            if ($Assignments) {
                $Assignments | Format-Table Role, CustomResourceScope -AutoSize
                $CurrentAssignments = $Assignments  # Store for later use in config mode
            } else {
                if ($VerifyOnly) {
                    Write-Host "✗ No role assignments found" -ForegroundColor Red
                } else {
                    Write-Host "No roles currently assigned to this application for this mailbox." -ForegroundColor Yellow
                }
            }
            
            # Test authorization
            Write-Host ""
            Write-Host "Authorization Test:" -ForegroundColor Yellow
            try {
                $TestResults = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $TargetMailbox
                $TestResults | Format-Table RoleName, InScope, GrantedPermissions -AutoSize
                
                $InScopeCount = ($TestResults | Where-Object {$_.InScope -eq $true}).Count
                if ($InScopeCount -gt 0) {
                    Write-Host "✓ $InScopeCount permission(s) are in scope" -ForegroundColor Green
                } else {
                    Write-Host "✗ No permissions are in scope" -ForegroundColor Red
                }
            } catch {
                Write-Host "✗ Authorization test failed: $($_.Exception.Message)" -ForegroundColor Red
            }
            
        } else {
            if ($VerifyOnly) {
                Write-Host "✗ Service Principal NOT found" -ForegroundColor Red
            } else {
                Write-Host "Service principal not found. Will be created during configuration." -ForegroundColor Yellow
            }
            $ObjectId = $null
        }
        
        # If this is verification mode only, we're done here
        if ($VerifyOnly) {
            Write-Host ""
            Write-Host "Verification completed" -ForegroundColor Green
            
            # Disconnect if we made the connection
            if (-not $WasAlreadyConnected) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "Disconnected from Exchange Online" -ForegroundColor Gray
            }
            return
        }
        
    } catch {
        if ($VerifyOnly) {
            Write-Error "Verification failed: $($_.Exception.Message)"
            exit 1
        } else {
            Write-Host "Could not retrieve current permissions. Continuing with configuration..." -ForegroundColor Yellow
        }
    }
    
    # Prompt for operation type if no roles specified (default to viewing roles)
    if ($RolesToAdd.Count -eq 0 -and $RolesToRemove.Count -eq 0 -and -not $RemoveAllRoles) {
        Write-Host ""
        Write-Host "Select operation:" -ForegroundColor Yellow
        Write-Host "1. View available roles first (recommended)" -ForegroundColor Green
        Write-Host "2. Add roles (grant permissions)"
        Write-Host "3. Remove roles (revoke permissions)"
        Write-Host "4. Remove ALL roles"
        
        # Default to option 1 (View roles first)
        $OperationChoice = Read-Host "Enter choice (1-4) [default: 1]"
        if (-not $OperationChoice -or $OperationChoice.Trim() -eq "") {
            $OperationChoice = "1"
        }
        
        do {
            switch ($OperationChoice) {
                "1" {
                    # View available roles first (default action)
                    Write-Host ""
                    Write-Host "=== Available Exchange Application Roles ===" -ForegroundColor Green
                    Write-Host ""
                    $RoleNumber = 1
                    $RolesList = $AvailableRoles.Keys | Sort-Object
                    foreach ($Role in $RolesList) {
                        $Info = $AvailableRoles[$Role]
                        Write-Host "$RoleNumber. $Role" -ForegroundColor Cyan
                        Write-Host "   $($Info.Description)" -ForegroundColor Gray
                        Write-Host ""
                        $RoleNumber++
                    }
                    Write-Host "Now select what you want to do:" -ForegroundColor Yellow
                    Write-Host "2. Add roles (grant permissions)"
                    Write-Host "3. Remove roles (revoke permissions)"
                    Write-Host "4. Remove ALL roles"
                    Write-Host "0. Exit without changes"
                    
                    $OperationChoice = Read-Host "Enter choice (2-4, 0 to exit)"
                }
                "2" {
                    # Add roles
                    Write-Host ""
                    Write-Host "=== Select Roles to Add ===" -ForegroundColor Green
                    $RoleNumber = 1
                    $RolesList = $AvailableRoles.Keys | Sort-Object
                    foreach ($Role in $RolesList) {
                        $Info = $AvailableRoles[$Role]
                        Write-Host "$RoleNumber. $Role" -ForegroundColor Cyan
                        Write-Host "   $($Info.Description)" -ForegroundColor Gray
                        $RoleNumber++
                    }
                    Write-Host ""
                    Write-Host "Examples:" -ForegroundColor Yellow
                    Write-Host "• Single role: 3"
                    Write-Host "• Multiple roles: 1,3,5"
                    Write-Host "• Range: 1-5"
                    Write-Host "• Mixed: 1,3-5,8"
                    Write-Host "• Default (recommended): press Enter for Mail.ReadWrite + Mail.Send + Calendars.ReadWrite"
                    
                    $RoleSelection = Read-Host "Enter role numbers"
                    
                    if (-not $RoleSelection -or $RoleSelection.Trim() -eq "") {
                        # Use default roles
                        $RolesToAdd = @("Application Mail.ReadWrite", "Application Mail.Send", "Application Calendars.ReadWrite")
                        Write-Host "Using default roles: $($RolesToAdd -join ', ')" -ForegroundColor Green
                    } else {
                        $SelectedRoles = @()
                        $Selections = $RoleSelection -split ","
                        foreach ($Selection in $Selections) {
                            $Selection = $Selection.Trim()
                            if ($Selection -match "^(\d+)-(\d+)$") {
                                # Range selection
                                $Start = [int]$Matches[1]
                                $End = [int]$Matches[2]
                                for ($i = $Start; $i -le $End; $i++) {
                                    if ($i -ge 1 -and $i -le $RolesList.Count) {
                                        $SelectedRoles += $RolesList[$i - 1]
                                    }
                                }
                            } elseif ($Selection -match "^\d+$") {
                                # Single selection
                                $Index = [int]$Selection
                                if ($Index -ge 1 -and $Index -le $RolesList.Count) {
                                    $SelectedRoles += $RolesList[$Index - 1]
                                }
                            }
                        }
                        $RolesToAdd = $SelectedRoles | Select-Object -Unique
                        if ($RolesToAdd.Count -eq 0) {
                            Write-Host "❌ No valid roles selected. Using default roles." -ForegroundColor Red
                            $RolesToAdd = @("Application Mail.ReadWrite", "Application Mail.Send", "Application Calendars.ReadWrite")
                        }
                    }
                    break
                }
                "3" {
                    # Remove roles
                    Write-Host ""
                    Write-Host "=== Select Roles to Remove ===" -ForegroundColor Yellow
                    
                    # Show only currently assigned roles for removal
                    if ($CurrentAssignments) {
                        Write-Host "Currently assigned roles you can remove:" -ForegroundColor Cyan
                        $CurrentRoles = $CurrentAssignments.Role | Sort-Object
                        $RoleNumber = 1
                        foreach ($Role in $CurrentRoles) {
                            Write-Host "$RoleNumber. $Role" -ForegroundColor Cyan
                            $RoleNumber++
                        }
                        Write-Host ""
                        Write-Host "Examples: 1,3,5 or 1-3 or 1,3-5,8" -ForegroundColor Gray
                        
                        do {
                            $RoleSelection = Read-Host "Enter role numbers to remove"
                            if (-not $RoleSelection -or $RoleSelection.Trim() -eq "") {
                                Write-Host "❌ Please select at least one role to remove!" -ForegroundColor Red
                            }
                        } while (-not $RoleSelection -or $RoleSelection.Trim() -eq "")
                        
                        $SelectedRoles = @()
                        $Selections = $RoleSelection -split ","
                        foreach ($Selection in $Selections) {
                            $Selection = $Selection.Trim()
                            if ($Selection -match "^(\d+)-(\d+)$") {
                                # Range selection
                                $Start = [int]$Matches[1]
                                $End = [int]$Matches[2]
                                for ($i = $Start; $i -le $End; $i++) {
                                    if ($i -ge 1 -and $i -le $CurrentRoles.Count) {
                                        $SelectedRoles += $CurrentRoles[$i - 1]
                                    }
                                }
                            } elseif ($Selection -match "^\d+$") {
                                # Single selection
                                $Index = [int]$Selection
                                if ($Index -ge 1 -and $Index -le $CurrentRoles.Count) {
                                    $SelectedRoles += $CurrentRoles[$Index - 1]
                                }
                            }
                        }
                        $RolesToRemove = $SelectedRoles | Select-Object -Unique
                        if ($RolesToRemove.Count -eq 0) {
                            Write-Host "❌ No valid roles selected!" -ForegroundColor Red
                            $OperationChoice = ""
                        } else {
                            break
                        }
                    } else {
                        Write-Host "❌ No roles are currently assigned to remove!" -ForegroundColor Red
                        Write-Host "Choose a different operation:" -ForegroundColor Yellow
                        $OperationChoice = ""
                    }
                }
                "4" {
                    # Remove all roles
                    if ($CurrentAssignments) {
                        $RemoveAllRoles = $true
                        Write-Host "Selected: Remove ALL roles" -ForegroundColor Red
                        break
                    } else {
                        Write-Host "❌ No roles are currently assigned to remove!" -ForegroundColor Red
                        Write-Host "Choose a different operation:" -ForegroundColor Yellow
                        $OperationChoice = ""
                    }
                }
                "0" {
                    Write-Host "Exiting without changes." -ForegroundColor Yellow
                    return
                }
                default {
                    Write-Host "❌ Invalid choice. Please enter 2, 3, 4, or 0." -ForegroundColor Red
                    $OperationChoice = ""
                }
            }
        } while (-not $OperationChoice -or $OperationChoice -notin @("2", "3", "4", "0"))
    }
    
    # Only show configuration summary and ask for confirmation in configuration mode
    if (-not $VerifyOnly) {
        Write-Host ""
        Write-Host "=== Configuration Summary ===" -ForegroundColor Green
        Write-Host "Application ID: $AppId" -ForegroundColor White
        Write-Host "Target Mailbox: $TargetMailbox" -ForegroundColor White
        Write-Host "Display Name: $AppDisplayName" -ForegroundColor White
        if ($RemoveAllRoles) {
            Write-Host "Operation: Remove ALL roles" -ForegroundColor Red
        } elseif ($RolesToAdd.Count -gt 0) {
            Write-Host "Roles to Add: $($RolesToAdd -join ', ')" -ForegroundColor Green
        } elseif ($RolesToRemove.Count -gt 0) {
            Write-Host "Roles to Remove: $($RolesToRemove -join ', ')" -ForegroundColor Yellow
        }
        Write-Host ""
        
        $Confirmation = Read-Host "Proceed with this configuration? (y/N)"
        if ($Confirmation -notmatch "^[Yy]") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return
        }
    }

# Validate required parameters when not just listing roles or in verify mode
if (-not $AppId -or -not $TargetMailbox) {
    Write-Host "❌ AppId and TargetMailbox are required when not using -ListAvailableRoles" -ForegroundColor Red
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com'" -ForegroundColor Cyan
    Write-Host "  .\Configure-AppPermissions.ps1 -AppId 'your-app-id' -TargetMailbox 'mailbox@domain.com' -VerifyOnly" -ForegroundColor Cyan
    Write-Host "  .\Configure-AppPermissions.ps1 -ListAvailableRoles" -ForegroundColor Cyan
    Write-Host "  .\Configure-AppPermissions.ps1    # Interactive mode" -ForegroundColor Cyan
    Write-Host "  .\Configure-AppPermissions.ps1 -VerifyOnly    # Interactive verification mode" -ForegroundColor Cyan
    return
}

# If this is verification mode, skip all the configuration logic
if ($VerifyOnly) {
    # The verification logic has already been executed in the interactive section above
    # If we reach here, it means parameters were provided via command line
    Write-Host "=== Verifying Application Permissions ===" -ForegroundColor Green
    Write-Host "App ID: $AppId" -ForegroundColor White
    Write-Host "Target Mailbox: $TargetMailbox" -ForegroundColor White
    Write-Host ""

    # Import required modules
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
    } catch {
        Write-Error "Failed to import ExchangeOnlineManagement module: $($_.Exception.Message)"
        Write-Host "Please install: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
        exit 1
    }

    # Check if already connected to Exchange Online
    $WasAlreadyConnected = $false
    try {
        $ExistingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ExistingConnection -and $ExistingConnection.State -eq "Connected") {
            # Verify connection is still active by testing a simple command
            try {
                Get-OrganizationConfig -ErrorAction Stop | Out-Null
                Write-Host "✓ Using existing Exchange Online connection ($($ExistingConnection.UserPrincipalName))" -ForegroundColor Green
                $WasAlreadyConnected = $true
            } catch {
                Write-Host "⚠ Existing connection expired, reconnecting..." -ForegroundColor Yellow
                $ExistingConnection = $null
            }
        }
        
        if (-not $ExistingConnection -or $ExistingConnection.State -ne "Connected") {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
            Connect-ExchangeOnline -ShowBanner:$false
            Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
        }
    } catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
        exit 1
    }

    try {
        # Check service principal
        $ServicePrincipal = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
        if ($ServicePrincipal) {
            Write-Host "✓ Service Principal exists" -ForegroundColor Green
            $ObjectId = $ServicePrincipal.ObjectId
        } else {
            Write-Host "✗ Service Principal NOT found" -ForegroundColor Red
            $ObjectId = $null
        }
        
        # Check role assignments (using both AppId and ObjectId)
        Write-Host ""
        Write-Host "Role Assignments:" -ForegroundColor Yellow
        $Assignments = Get-ManagementRoleAssignment | Where-Object {$_.RoleAssigneeName -eq $AppId -or $_.RoleAssigneeName -eq $ObjectId}
        if ($Assignments) {
            $Assignments | Format-Table Role, CustomResourceScope -AutoSize
        } else {
            Write-Host "✗ No role assignments found" -ForegroundColor Red
        }
        
        # Test authorization
        Write-Host ""
        Write-Host "Authorization Test:" -ForegroundColor Yellow
        try {
            $TestResults = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $TargetMailbox
            $TestResults | Format-Table RoleName, InScope, GrantedPermissions -AutoSize
            
            $InScopeCount = ($TestResults | Where-Object {$_.InScope -eq $true}).Count
            if ($InScopeCount -gt 0) {
                Write-Host "✓ $InScopeCount permission(s) are in scope" -ForegroundColor Green
            } else {
                Write-Host "✗ No permissions are in scope" -ForegroundColor Red
            }
        } catch {
            Write-Host "✗ Authorization test failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
    } catch {
        Write-Error "Verification failed: $($_.Exception.Message)"
    } finally {
        # Only disconnect if we made the connection
        if (-not $WasAlreadyConnected) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "Disconnected from Exchange Online" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "Verification completed" -ForegroundColor Green
    return
}

# Use default roles if no specific roles are specified and not removing all
if ($RolesToAdd.Count -eq 0 -and $RolesToRemove.Count -eq 0 -and -not $RemoveAllRoles) {
    $RolesToAdd = @("Application Mail.ReadWrite", "Application Mail.Send", "Application Calendars.ReadWrite")
}

Write-Host "=== Exchange Online Application Permissions Setup ===" -ForegroundColor Green
Write-Host "Application ID: $AppId" -ForegroundColor White
Write-Host "Target Mailbox: $TargetMailbox" -ForegroundColor White
Write-Host "Display Name: $AppDisplayName" -ForegroundColor White

if ($RemoveAllRoles) {
    Write-Host "Operation: Remove ALL roles" -ForegroundColor Red
} else {
    if ($RolesToAdd.Count -gt 0) {
        Write-Host "Roles to Add: $($RolesToAdd -join ', ')" -ForegroundColor Green
    }
    if ($RolesToRemove.Count -gt 0) {
        Write-Host "Roles to Remove: $($RolesToRemove -join ', ')" -ForegroundColor Yellow
    }
}
Write-Host ""

# Validate requested roles
$InvalidRoles = @()
if (-not $RemoveAllRoles) {
    foreach ($Role in ($RolesToAdd + $RolesToRemove)) {
        if (-not $AvailableRoles.ContainsKey($Role)) {
            $InvalidRoles += $Role
        }
    }
}

if ($InvalidRoles.Count -gt 0) {
    Write-Host "❌ Invalid roles specified:" -ForegroundColor Red
    $InvalidRoles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Use -ListAvailableRoles to see all available roles." -ForegroundColor Yellow
    return
}

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
} catch {
    Write-Error "Failed to import ExchangeOnlineManagement module: $($_.Exception.Message)"
    Write-Host "Please install: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
    exit 1
}

# Check if already connected to Exchange Online
$WasAlreadyConnected = $false
try {
    $ExistingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ExistingConnection -and $ExistingConnection.State -eq "Connected") {
        # Verify connection is still active by testing a simple command
        try {
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
            Write-Host "✓ Using existing Exchange Online connection ($($ExistingConnection.UserPrincipalName))" -ForegroundColor Green
            $WasAlreadyConnected = $true
        } catch {
            Write-Host "⚠ Existing connection expired, reconnecting..." -ForegroundColor Yellow
            $ExistingConnection = $null
        }
    }
    
    if (-not $ExistingConnection -or $ExistingConnection.State -ne "Connected") {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
    exit 1
}

try {
    # Step 1: Create management scope
    $CleanAppName = $AppDisplayName -replace '[^a-zA-Z0-9]', ''  # Remove special characters
    $MailboxPrefix = $TargetMailbox.Split('@')[0]
    $GenericScopeName = "MailboxScope-$MailboxPrefix"
    $AppSpecificScopeName = "AppScope-$CleanAppName-$MailboxPrefix"
    
    # Check if user wants to rename existing scope to generic name
    if ($RenameExistingScope) {
        $ExistingScope = Get-ManagementScope | Where-Object {$_.RecipientFilter -eq "PrimarySmtpAddress -eq '$TargetMailbox'"}
        if ($ExistingScope -and $ExistingScope.Name -ne $GenericScopeName) {
            try {
                Write-Host "Renaming existing scope '$($ExistingScope.Name)' to '$GenericScopeName'..." -ForegroundColor Yellow
                Set-ManagementScope -Identity $ExistingScope.Name -Name $GenericScopeName
                Write-Host "✓ Renamed management scope to: $GenericScopeName" -ForegroundColor Green
                $ScopeName = $GenericScopeName
            } catch {
                Write-Host "⚠ Failed to rename scope, using existing name: $($ExistingScope.Name)" -ForegroundColor Yellow
                $ScopeName = $ExistingScope.Name
            }
        } else {
            $ScopeName = $GenericScopeName
            Write-Host "✓ Using generic scope name: $ScopeName" -ForegroundColor Green
        }
    } else {
        $ScopeName = $AppSpecificScopeName
        Write-Host "Creating management scope: $ScopeName" -ForegroundColor Yellow
    }
    
    # Check if this exact scope exists first
    $ExistingScope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue
    if ($ExistingScope) {
        Write-Host "✓ Using existing management scope: $ScopeName" -ForegroundColor Green
    } else {
        # Check if a generic scope with the same filter already exists
        $GenericScope = Get-ManagementScope | Where-Object {$_.RecipientFilter -eq "PrimarySmtpAddress -eq '$TargetMailbox'"}
        if ($GenericScope -and $GenericScope.Count -eq 1 -and -not $ForceNewScope) {
            # Ask user if they want to reuse the existing scope or create a new one
            $ExistingScopeName = $GenericScope.Name
            Write-Host "⚠ Found existing scope '$ExistingScopeName' for the same mailbox." -ForegroundColor Yellow
            Write-Host "  Reusing existing scope (use -ForceNewScope to create app-specific scope)" -ForegroundColor White
            
            $ScopeName = $ExistingScopeName
            Write-Host "✓ Reusing existing management scope: $ScopeName" -ForegroundColor Green
        } else {
            if ($ForceNewScope -and $GenericScope) {
                Write-Host "  -ForceNewScope specified: Creating app-specific scope" -ForegroundColor Yellow
            }
            try {
                New-ManagementScope -Name $ScopeName -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$TargetMailbox'"
                Write-Host "✓ Created new management scope: $ScopeName" -ForegroundColor Green
            } catch {
                # Check if error is about duplicate filter
                if ($_.Exception.Message -like "*same RecipientRoot, RecipientFilter*" -or $_.Exception.Message -like "*same*management scope*") {
                    Write-Host "❌ Cannot create app-specific scope: Exchange limitation" -ForegroundColor Red
                    Write-Host "   Exchange Online allows only one scope per mailbox filter" -ForegroundColor Yellow
                    
                    # Fall back to existing scope
                    $AnyExistingScope = Get-ManagementScope | Where-Object {$_.RecipientFilter -eq "PrimarySmtpAddress -eq '$TargetMailbox'"} | Select-Object -First 1
                    if ($AnyExistingScope) {
                        $ScopeName = $AnyExistingScope.Name
                        Write-Host "✓ Using existing management scope: $ScopeName" -ForegroundColor Green
                    } else {
                        throw $_
                    }
                } else {
                    throw $_
                }
            }
        }
    }

    # Step 2: Create service principal
    Write-Host "Creating service principal..." -ForegroundColor Yellow
    
    $ExistingServicePrincipal = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
    if (-not $ExistingServicePrincipal) {
        # We need ObjectId for New-ServicePrincipal - try to get it from Graph
        try {
            Import-Module Microsoft.Graph.Applications -ErrorAction Stop
            Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome -ErrorAction Stop
            $GraphSP = Get-MgServicePrincipal -Filter "AppId eq '$AppId'"
            if ($GraphSP) {
                New-ServicePrincipal -AppId $AppId -ObjectId $GraphSP.Id -DisplayName $AppDisplayName
                Write-Host "✓ Created service principal" -ForegroundColor Green
            } else {
                throw "Service principal not found in Azure AD"
            }
        } catch {
            Write-Host "⚠ Could not create service principal automatically" -ForegroundColor Yellow
            Write-Host "  Please create manually or provide ObjectId" -ForegroundColor Yellow
            Write-Host "  Command: New-ServicePrincipal -AppId $AppId -ObjectId <ObjectId> -DisplayName '$AppDisplayName'" -ForegroundColor Cyan
        }
    } else {
        Write-Host "⚠ Service principal already exists" -ForegroundColor Yellow
    }

    # Step 3: Manage permissions
    Write-Host "Managing permissions..." -ForegroundColor Yellow
    
    # Get the service principal ObjectId for checking existing assignments
    $ServicePrincipal = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
    $ServicePrincipalObjectId = if ($ServicePrincipal) { $ServicePrincipal.ObjectId } else { $null }

    # Remove all roles if requested
    if ($RemoveAllRoles) {
        Write-Host "Removing all role assignments..." -ForegroundColor Red
        $ExistingAssignments = Get-ManagementRoleAssignment | Where-Object {
            $_.RoleAssigneeName -eq $ServicePrincipalObjectId -and $_.CustomResourceScope -eq $ScopeName
        }
        
        foreach ($Assignment in $ExistingAssignments) {
            try {
                Remove-ManagementRoleAssignment -Identity $Assignment.Name -Confirm:$false
                Write-Host "  ✓ Removed $($Assignment.Role)" -ForegroundColor Green
            } catch {
                Write-Host "  ✗ Failed to remove $($Assignment.Role): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        # Remove specific roles
        if ($RolesToRemove.Count -gt 0) {
            Write-Host "Removing specified roles..." -ForegroundColor Yellow
            foreach ($Role in $RolesToRemove) {
                $ExistingAssignment = if ($ServicePrincipalObjectId) {
                    Get-ManagementRoleAssignment | Where-Object {
                        $_.Role -eq $Role -and $_.RoleAssigneeName -eq $ServicePrincipalObjectId -and $_.CustomResourceScope -eq $ScopeName
                    }
                } else {
                    Get-ManagementRoleAssignment | Where-Object {
                        $_.Role -eq $Role -and $_.Name -like "*$AppId*" -and $_.CustomResourceScope -eq $ScopeName
                    }
                }
                
                if ($ExistingAssignment) {
                    try {
                        Remove-ManagementRoleAssignment -Identity $ExistingAssignment.Name -Confirm:$false
                        Write-Host "  ✓ Removed $Role" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ Failed to remove ${Role}: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "  ⚠ $Role was not assigned" -ForegroundColor Yellow
                }
            }
        }

        # Add specific roles
        if ($RolesToAdd.Count -gt 0) {
            Write-Host "Adding specified roles..." -ForegroundColor Yellow
            foreach ($Role in $RolesToAdd) {
                # Check for existing assignment using ObjectId (which is stored in RoleAssigneeName)
                $ExistingAssignment = if ($ServicePrincipalObjectId) {
                    Get-ManagementRoleAssignment | Where-Object {
                        $_.Role -eq $Role -and $_.RoleAssigneeName -eq $ServicePrincipalObjectId -and $_.CustomResourceScope -eq $ScopeName
                    }
                } else {
                    # Fallback: check by partial name match
                    Get-ManagementRoleAssignment | Where-Object {
                        $_.Role -eq $Role -and $_.Name -like "*$AppId*" -and $_.CustomResourceScope -eq $ScopeName
                    }
                }
                
                if (-not $ExistingAssignment) {
                    try {
                        New-ManagementRoleAssignment -Role $Role -App $AppId -CustomResourceScope $ScopeName
                        Write-Host "  ✓ Assigned $Role" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ Failed to assign $Role : $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "  ⚠ $Role already assigned" -ForegroundColor Yellow
                }
            }
        }
    }

    # Step 4: Test permissions
    Write-Host ""
    Write-Host "Testing permissions..." -ForegroundColor Yellow
    try {
        $TestResults = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $TargetMailbox
        Write-Host "✓ Permission test completed" -ForegroundColor Green
        $TestResults | Format-Table RoleName, InScope, GrantedPermissions -AutoSize
    } catch {
        Write-Host "⚠ Could not test permissions: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "=== CONFIGURATION COMPLETE ===" -ForegroundColor Green
    
    if (-not $RemoveAllRoles -and $RolesToAdd.Count -gt 0) {
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "1. Add Microsoft Graph API permissions in Azure Portal:" -ForegroundColor White
        
        # Show unique Graph permissions needed
        $GraphPermissions = @()
        foreach ($Role in $RolesToAdd) {
            if ($AvailableRoles[$Role].Protocol -eq "MS Graph") {
                $Permissions = $AvailableRoles[$Role].Permissions -split ", "
                $GraphPermissions += $Permissions
            }
        }
        $UniquePermissions = $GraphPermissions | Select-Object -Unique
        foreach ($Permission in $UniquePermissions) {
            Write-Host "   - $Permission (Application)" -ForegroundColor White
        }
        
        Write-Host "2. Grant admin consent" -ForegroundColor White
        Write-Host "3. Wait 30-60 minutes for changes to take effect" -ForegroundColor White
    } elseif ($RemoveAllRoles -or $RolesToRemove.Count -gt 0) {
        Write-Host "Permissions have been removed from Exchange Online." -ForegroundColor Yellow
        Write-Host "You may also want to review Azure Portal for corresponding Graph API permissions." -ForegroundColor White
    }

} catch {
    Write-Error "Configuration failed: $($_.Exception.Message)"
} finally {
    # Only disconnect if we made the connection
    if (-not $WasAlreadyConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Configuration completed for $AppDisplayName" -ForegroundColor Green
