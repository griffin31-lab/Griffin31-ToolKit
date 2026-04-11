<#
.SYNOPSIS
    Usage Examples for Exchange Online Application Permissions Script

.DESCRIPTION
    This file contains practical examples for using the Configure-AppPermissions.ps1 script.
    Examples cover all functionality including configuration, verification, and management.

.NOTES
    Author: Exchange Admin Script
    Version: 3.0 (Unified Script)
    Last Updated: September 2025
    
    TABLE OF CONTENTS:
    ==================
    1. Prerequisites & Setup
    2. Interactive Mode Examples
    3. Basic Configuration Examples
    4. Role Management Examples
    5. Verification Examples (replaces Verify-AppPermissions.ps1)
    6. Management Scope Examples
    7. Different Mailbox Types
    8. Troubleshooting & Tips
    
    Instructions:
    1. Load variables into your PowerShell session FIRST:
       Option A: Run these commands in PowerShell:
         $AppId = "your-app-id"
         $TargetMailbox = "mailbox@domain.com"  # Any mailbox type
         $AppDisplayName = "YourAppName"
       
       Option B: Source the CONFIG.ps1 file:
         . .\CONFIG.ps1
    
    2. Then copy and run any example below (uses $TargetMailbox variable)
    3. Note: Configure-AppPermissions.ps1 now includes verification functionality
    4. See README.md for complete documentation
#>

# Example Usage - Exchange App Permissions with Role Management

# ==================================================================================
# 1. PREREQUISITES & SETUP - LOAD VARIABLES INTO YOUR POWERSHELL SESSION
# ==================================================================================
# Copy and paste these commands into PowerShell BEFORE running examples:
   $AppId = "00000000-0000-0000-0000-000000000000"
   $TargetMailbox = "mailbox@yourdomain.com"  # Any mailbox: shared, user, resource, etc.
   $AppDisplayName = "YourAppName"
#
# OR source the CONFIG.ps1 file:
#   . .\CONFIG.ps1

# ==================================================================================
# 2. INTERACTIVE MODE EXAMPLES
# ==================================================================================

# Run script with no parameters - prompts for all inputs
.\Configure-AppPermissions.ps1

# ==================================================================================  
# 3. BASIC CONFIGURATION EXAMPLES
# ==================================================================================

# NEW: Interactive Mode (recommended for beginners)
.\Configure-AppPermissions.ps1
# This will:
# 1. Prompt for App ID and Target Mailbox
# 2. Show current permissions (like Verify-AppPermissions)
# 3. Display available roles by default
# 4. Guide you through role selection

# NEW: Interactive Verification Mode (replaces Verify-AppPermissions.ps1)
.\Configure-AppPermissions.ps1 -VerifyOnly
# This will:
# 1. Prompt for App ID and Target Mailbox
# 2. Show current permissions and role assignments
# 3. Test authorization without making changes

# Show all available roles
.\Configure-AppPermissions.ps1 -ListAvailableRoles

# Connect once (optional - scripts will check if already connected)
# Connect-ExchangeOnline

# ==================================================================================
# 4. ROLE MANAGEMENT EXAMPLES  
# ==================================================================================

# ----- Adding Roles (-RolesToAdd) -----

# Default configuration (Mail.ReadWrite + Mail.Send + Calendars.ReadWrite)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -AppDisplayName $AppDisplayName

# Add single role - Mail Read Only
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Mail.Read")

# Add multiple specific roles
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Mail.Read", "Application Contacts.ReadWrite")

# Full Exchange access (all permissions)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Exchange Full Access")

# Mail-only access (read + write + send)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Mail Full Access")

# Read-only access (mail + calendar + contacts)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Mail.Read", "Application Calendars.Read", "Application Contacts.Read")

# EWS access (for legacy applications)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application EWS.AccessAsApp")

# Calendar-only access
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Calendars.ReadWrite")

# ----- Removing Roles (-RolesToRemove) -----

# Remove specific single role
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToRemove @("Application Mail.Send")

# Remove multiple roles
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToRemove @("Application Mail.Send", "Application Calendars.ReadWrite")

# Add and remove roles in one command
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application MailboxSettings.ReadWrite") -RolesToRemove @("Application Calendars.ReadWrite")

# ----- Remove All Roles (-RemoveAllRoles) -----

# Remove all roles from the application
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RemoveAllRoles

# ==================================================================================
# 5. VERIFICATION EXAMPLES (replaces Verify-AppPermissions.ps1)
# ==================================================================================

# Verify current configuration (command line)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -VerifyOnly

# Verify current configuration (interactive)
.\Configure-AppPermissions.ps1 -VerifyOnly

# ==================================================================================
# 6. MANAGEMENT SCOPE EXAMPLES
# ==================================================================================

# Rename existing scope to generic name (fixes app-specific naming)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RenameExistingScope

# ==================================================================================
# 7. DIFFERENT MAILBOX TYPES EXAMPLES
# ==================================================================================

# Example A: Shared mailbox (most common)
# $TargetMailbox = "shared.finance@domain.com"
# .\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Mail.ReadWrite")

# Example B: User mailbox  
# $TargetMailbox = "john.doe@domain.com"
# .\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Calendars.ReadWrite")

# Example C: Conference room resource
# $TargetMailbox = "conference.room1@domain.com"
# .\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Calendars.ReadWrite")

# Example D: Equipment resource
# $TargetMailbox = "projector.main@domain.com" 
# .\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RolesToAdd @("Application Calendars.Read")

# ==================================================================================
# 8. TROUBLESHOOTING & TIPS
# ==================================================================================

# ----- Common Issues and Solutions -----

# Issue: "InScope=False" during verification
# Solution: Wait 30-60 minutes for permissions to propagate, or check Azure AD app permissions

# Issue: Management scope naming conflicts  
# Solution: Use -RenameExistingScope to rename existing scope to generic name
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -RenameExistingScope

# Issue: Need to force new scope (will fail due to Exchange limitation)
.\Configure-AppPermissions.ps1 -AppId $AppId -TargetMailbox $TargetMailbox -ForceNewScope

# ----- Available Roles Quick Reference -----
# Application Mail.Read: Read email only
# Application Mail.ReadBasic: Read email without body/attachments  
# Application Mail.ReadWrite: Full email access (no send)
# Application Mail.Send: Send email as any user
# Application MailboxSettings.Read: Read mailbox settings
# Application MailboxSettings.ReadWrite: Manage mailbox settings
# Application Calendars.Read: Read calendar events
# Application Calendars.ReadWrite: Manage calendar events
# Application Contacts.Read: Read contacts
# Application Contacts.ReadWrite: Manage contacts
# Application Mail Full Access: Mail.ReadWrite + Mail.Send
# Application Exchange Full Access: All permissions combined
# Application EWS.AccessAsApp: EWS protocol access

# ----- Script Benefits Summary -----
# ✓ Supports all 13 Exchange application roles
# ✓ Unified script (configuration + verification in one)
# ✓ Interactive mode for user-friendly operation
# ✓ Smart connection reuse and role validation
# ✓ Automatic Graph permission mapping
# ✓ Bulk operations (add/remove multiple roles)
# ✓ Complete authorization testing and scope management
# ✓ Works with all mailbox types (Shared, User, Room, Equipment)
