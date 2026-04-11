# Exchange App Permissions Configuration
# Source this file to load variables into your PowerShell session
# Usage: . .\CONFIG.ps1

# ===== UPDATE THESE VALUES FOR YOUR ENVIRONMENT =====
$AppId = "00000000-0000-0000-0000-000000000000"     # Your Azure AD App Registration ID
$TargetMailbox = "your-mailbox@yourdomain.com"      # Target mailbox (any type)

# Examples of different mailbox types:
# $TargetMailbox = "shared.mailbox@domain.com"     # Shared mailbox
# $TargetMailbox = "john.doe@domain.com"           # User mailbox  
# $TargetMailbox = "conference.room1@domain.com"   # Room resource mailbox
# $TargetMailbox = "projector1@domain.com"         # Equipment resource mailbox

# Verify variables are loaded
Write-Host "✓ Configuration loaded:" -ForegroundColor Green
Write-Host "  AppId: $AppId" -ForegroundColor White
Write-Host "  Target Mailbox: $TargetMailbox" -ForegroundColor White
Write-Host "  AppDisplayName: $AppDisplayName" -ForegroundColor White
Write-Host ""
Write-Host "Supported mailbox types: Shared, User, Room, Equipment, Distribution Groups" -ForegroundColor Cyan
Write-Host "You can now use variables in commands like:" -ForegroundColor Cyan
Write-Host "  .\Configure-AppPermissions.ps1 -AppId `$AppId -TargetMailbox `$TargetMailbox" -ForegroundColor Yellow
