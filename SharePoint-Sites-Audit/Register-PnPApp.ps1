[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [Parameter(Mandatory)]
    [string]$ConfigPath,
    [string]$ApplicationName = "Griffin31 SPO Audit"
)

# One-time registration of an Entra ID app for PnP.PowerShell.
# Creates a delegated-only app (no secret), grants consent, saves ClientId to config.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  Griffin31 SPO Audit — First-time setup" -ForegroundColor Cyan
Write-Host "  This registers a delegated-only Entra ID app in your tenant." -ForegroundColor Gray
Write-Host "  No client secret. No certificate. You'll be asked to consent once." -ForegroundColor Gray
Write-Host ""
Write-Host "  Required: Global Administrator role (to grant consent)." -ForegroundColor Yellow
Write-Host ""

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "  Launching registration — a browser window will open." -ForegroundColor Gray
Write-Host ""
try {
    $result = Register-PnPEntraIDApp `
        -ApplicationName $ApplicationName `
        -Tenant $TenantDomain `
        -Interactive `
        -GraphDelegatePermissions @('Group.Read.All','Directory.Read.All','InformationProtectionPolicy.Read','User.Read.All') `
        -SharePointDelegatePermissions @('AllSites.FullControl','User.Read.All') `
        -ErrorAction Stop

    # Result returns a hashtable with "AzureAppId/ClientId"
    $clientId = $null
    if ($result -is [hashtable] -or $result -is [System.Collections.IDictionary]) {
        foreach ($k in @('AzureAppId','ClientId','ApplicationId')) {
            if ($result.ContainsKey($k)) { $clientId = $result[$k]; break }
        }
    } elseif ($result.AzureAppId) {
        $clientId = $result.AzureAppId
    } elseif ($result.ClientId) {
        $clientId = $result.ClientId
    } elseif ($result.ApplicationId) {
        $clientId = $result.ApplicationId
    }

    if (-not $clientId) {
        # Fallback: ask user to paste it
        Write-Host ""
        Write-Host "  Registration completed but I could not read the ClientId from the return value." -ForegroundColor Yellow
        $clientId = Read-Host "  Paste the AppId shown in the output above"
    }

    if (-not $clientId -or $clientId -notmatch '^[0-9a-fA-F-]{36}$') {
        throw "No valid ClientId captured."
    }

    # Save
    $config = @{
        TenantDomain    = $TenantDomain
        ClientId        = $clientId
        ApplicationName = $ApplicationName
        RegisteredAt    = (Get-Date).ToString("o")
    }
    $configDir = Split-Path $ConfigPath -Parent
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
    $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8

    Write-Host ""
    Write-Host "  Setup complete." -ForegroundColor Green
    Write-Host "  Application: $ApplicationName" -ForegroundColor Gray
    Write-Host "  ClientId:    $clientId" -ForegroundColor Gray
    Write-Host "  Saved to:    $ConfigPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  You can now run the full pipeline (menu option 3 or 4)." -ForegroundColor Cyan
    exit 0
} catch {
    Write-Host ""
    Write-Host "  [!] Registration failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Common causes:" -ForegroundColor Yellow
    Write-Host "    - You are not a Global Administrator (required to consent to new apps)" -ForegroundColor Yellow
    Write-Host "    - Your tenant blocks user consent to unverified apps" -ForegroundColor Yellow
    Write-Host "    - Browser window was closed before consent was granted" -ForegroundColor Yellow
    exit 1
}
