[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [Parameter(Mandatory)]
    [string]$ConfigPath,
    [string]$ApplicationName
)

# First-time setup: register an Entra ID app with APP-ONLY permissions and a self-signed cert.
# No client secret. No repeated interactive login after this. One-time Global Admin consent only.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "     SharePoint Sites Audit — First-time setup" -ForegroundColor Yellow
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  I'll do the following — this runs ONCE per tenant:" -ForegroundColor White
Write-Host "    1. Prompt you to sign in as Global Administrator" -ForegroundColor Gray
Write-Host "    2. Create a new Entra ID app in your tenant" -ForegroundColor Gray
Write-Host "    3. Generate a self-signed certificate on this machine" -ForegroundColor Gray
Write-Host "    4. Upload the cert's public key to the app" -ForegroundColor Gray
Write-Host "    5. Grant app-only permissions (SharePoint + Graph)" -ForegroundColor Gray
Write-Host "    6. Save the config to: $ConfigPath" -ForegroundColor Gray
Write-Host ""
Write-Host "  After setup, every subsequent run is silent — no login needed." -ForegroundColor Green
Write-Host ""

if (-not $ApplicationName) {
    Write-Host -NoNewline "  App name (press ENTER for 'Griffin31 SPO Audit'): " -ForegroundColor Yellow
    $ApplicationName = Read-Host
    if (-not $ApplicationName) { $ApplicationName = "Griffin31 SPO Audit" }
}

Write-Host ""
Write-Host "  Press ENTER to start. A browser window will open for Global Admin sign-in." -ForegroundColor Yellow
$null = Read-Host

Import-Module PnP.PowerShell -ErrorAction Stop

# Cert output folder (same folder as config.json)
$configDir = Split-Path $ConfigPath -Parent
if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
$certDir = Join-Path $configDir "cert"
if (-not (Test-Path $certDir)) { New-Item -ItemType Directory -Path $certDir -Force | Out-Null }

# Generate a random cert password (user never sees or enters it)
Add-Type -AssemblyName System.Security
$bytes = New-Object byte[] 32
[System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
$plainPassword = [Convert]::ToBase64String($bytes)
$securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

Write-Host ""
Write-Host "  Running Register-PnPEntraIDApp (this opens a browser)..." -ForegroundColor Cyan
Write-Host ""

try {
    $result = Register-PnPEntraIDApp `
        -ApplicationName $ApplicationName `
        -Tenant $TenantDomain `
        -OutPath $certDir `
        -CertificatePassword $securePassword `
        -GraphApplicationPermissions @('Group.Read.All','Directory.Read.All','InformationProtectionPolicy.Read.All','Sites.Read.All') `
        -SharePointApplicationPermissions @('Sites.FullControl.All','User.Read.All') `
        -Interactive `
        -ErrorAction Stop

    # Extract ClientId from result
    $clientId = $null
    foreach ($k in @('AzureAppId','ClientId','ApplicationId')) {
        if ($result.$k) { $clientId = $result.$k; break }
    }
    if (-not $clientId -and ($result -is [hashtable] -or $result -is [System.Collections.IDictionary])) {
        foreach ($k in @('AzureAppId','ClientId','ApplicationId')) {
            if ($result.ContainsKey($k)) { $clientId = $result[$k]; break }
        }
    }
    if (-not $clientId -or $clientId -notmatch '^[0-9a-fA-F-]{36}$') {
        Write-Host ""
        Write-Host "  Could not auto-detect ClientId from the return value." -ForegroundColor Yellow
        $clientId = Read-Host "  Paste the App (client) ID shown above"
        if ($clientId -notmatch '^[0-9a-fA-F-]{36}$') { throw "Invalid ClientId." }
    }

    # Find the generated PFX in $certDir
    $pfx = Get-ChildItem -Path $certDir -Filter "*.pfx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $pfx) { throw "Certificate .pfx was not found in $certDir" }

    # Get the thumbprint by loading the cert
    $certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $pfx.FullName, $plainPassword
    $thumbprint = $certObj.Thumbprint
    $certObj.Dispose()

    # Encrypt the password using PS SecureString (machine- and user-bound on Windows via DPAPI;
    # on macOS/Linux PS uses AES key stored per-user).
    $encryptedPassword = $securePassword | ConvertFrom-SecureString

    # Save config
    $config = @{
        TenantDomain       = $TenantDomain
        ClientId           = $clientId
        ApplicationName    = $ApplicationName
        CertificatePath    = $pfx.FullName
        CertificateThumbprint = $thumbprint
        EncryptedCertPassword = $encryptedPassword
        RegisteredAt       = (Get-Date).ToString("o")
        AuthMode           = "CertificateAppOnly"
    }
    $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8

    Write-Host ""
    Write-Host "  Setup complete." -ForegroundColor Green
    Write-Host "    App name:    $ApplicationName" -ForegroundColor Gray
    Write-Host "    ClientId:    $clientId" -ForegroundColor Gray
    Write-Host "    Cert:        $($pfx.Name)" -ForegroundColor Gray
    Write-Host "    Thumbprint:  $thumbprint" -ForegroundColor Gray
    Write-Host "    Config:      $ConfigPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Note: admin consent may take a few minutes to propagate." -ForegroundColor Yellow
    Write-Host "  If the first export fails with 'insufficient privileges', wait 2-3 minutes and retry." -ForegroundColor Yellow
    exit 0
} catch {
    Write-Host ""
    Write-Host "  [!] Setup failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Common causes:" -ForegroundColor Yellow
    Write-Host "    - Signed-in user is not a Global Administrator" -ForegroundColor Yellow
    Write-Host "    - Tenant blocks user-initiated app registrations" -ForegroundColor Yellow
    Write-Host "    - Browser was closed before consent was granted" -ForegroundColor Yellow
    exit 1
}
