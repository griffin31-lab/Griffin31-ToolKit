[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [Parameter(Mandatory)]
    [string]$ConfigPath,
    [string]$ApplicationName
)

# First-time setup: create an Entra ID app via Graph API with APP-ONLY permissions
# and a self-signed cert. No dependency on PnP-version-specific cmdlet params.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "     SharePoint Sites Audit — First-time setup" -ForegroundColor Yellow
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  I'll do the following — this runs ONCE per tenant:" -ForegroundColor White
Write-Host "    1. Prompt you to sign in as Global Administrator (device login)" -ForegroundColor Gray
Write-Host "    2. Generate a self-signed certificate on this machine" -ForegroundColor Gray
Write-Host "    3. Create a new Entra ID app in your tenant via Graph API" -ForegroundColor Gray
Write-Host "    4. Attach the cert's public key to the app" -ForegroundColor Gray
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
Write-Host "  Press ENTER to start (a browser window / device-code prompt will appear)." -ForegroundColor Yellow
$null = Read-Host

# ── Prepare cert output folder ──
$configDir = Split-Path $ConfigPath -Parent
if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
$certDir = Join-Path $configDir "cert"
if (-not (Test-Path $certDir)) { New-Item -ItemType Directory -Path $certDir -Force | Out-Null }

# ── Step 1: Sign in to Graph with admin scopes needed to create apps + consent ──
Write-Host ""
Write-Host "  [1/6] Signing in to Microsoft Graph as Global Admin..." -ForegroundColor Cyan
$adminScopes = @(
    'Application.ReadWrite.All',          # create apps
    'AppRoleAssignment.ReadWrite.All',    # grant admin consent
    'Directory.ReadWrite.All'             # service principals
)
try {
    Connect-MgGraph -Scopes $adminScopes -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Graph context not established." }
    Write-Host "        Signed in as $($ctx.Account)" -ForegroundColor Green
} catch {
    Write-Host "        [!] Graph sign-in failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Resolve tenant ID
$tenantId = $ctx.TenantId
if (-not $tenantId) {
    try {
        $disc = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomain/v2.0/.well-known/openid-configuration"
        if ($disc.issuer -match 'sts\.windows\.net/([0-9a-f-]{36})') { $tenantId = $Matches[1] }
    } catch {}
}
if (-not $tenantId) {
    Write-Host "        [!] Could not determine tenant ID." -ForegroundColor Red
    exit 1
}

# ── Step 2: Generate self-signed cert ──
Write-Host ""
Write-Host "  [2/6] Generating self-signed certificate..." -ForegroundColor Cyan

Add-Type -AssemblyName System.Security
$pwBytes = New-Object byte[] 32
[System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($pwBytes)
$plainPassword  = [Convert]::ToBase64String($pwBytes)
$securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

$certSubject = "CN=$ApplicationName"
$notAfter    = (Get-Date).AddYears(2)

try {
    if ($IsWindows -or $env:OS -match 'Windows') {
        # Windows: use New-SelfSignedCertificate (CurrentUser\My, then export PFX)
        $cert = New-SelfSignedCertificate -Subject $certSubject `
                    -CertStoreLocation "Cert:\CurrentUser\My" `
                    -KeyExportPolicy Exportable `
                    -KeySpec Signature `
                    -KeyLength 2048 `
                    -HashAlgorithm SHA256 `
                    -NotAfter $notAfter `
                    -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"
        $pfxPath = Join-Path $certDir "$($ApplicationName -replace '[^\w\-]','_').pfx"
        Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath $pfxPath -Password $securePassword | Out-Null
    } else {
        # macOS / Linux: use .NET APIs directly (no cert store needed for PFX generation)
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $req = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest `
                    $certSubject, $rsa, 'SHA256', 'Pkcs1'
        $cert = $req.CreateSelfSigned((Get-Date).AddMinutes(-5), $notAfter)
        $pfxBytes = $cert.Export('Pkcs12', $plainPassword)
        $pfxPath = Join-Path $certDir "$($ApplicationName -replace '[^\w\-]','_').pfx"
        [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)
    }
    $thumbprint = $cert.Thumbprint
    $certBase64 = [Convert]::ToBase64String($cert.RawData)
    Write-Host "        Cert thumbprint: $thumbprint" -ForegroundColor Green
    Write-Host "        PFX saved to:    $pfxPath" -ForegroundColor Green
} catch {
    Write-Host "        [!] Cert generation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ── Step 3: Resolve resource service principals (Graph + SPO) and their app-role IDs ──
Write-Host ""
Write-Host "  [3/6] Resolving API app-role IDs..." -ForegroundColor Cyan

# Graph resource: 00000003-0000-0000-c000-000000000000
# SPO resource:   00000003-0000-0ff1-ce00-000000000000
$graphAppId = '00000003-0000-0000-c000-000000000000'
$spoAppId   = '00000003-0000-0ff1-ce00-000000000000'

try {
    $graphSP = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$graphAppId'"
    $graphSP = $graphSP.value[0]
    $spoSP   = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$spoAppId'"
    $spoSP   = $spoSP.value[0]
    if (-not $graphSP -or -not $spoSP) { throw "Could not resolve Graph or SPO service principal." }
} catch {
    Write-Host "        [!] Failed to resolve SPs: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

function Get-AppRoleId {
    param($servicePrincipal, [string]$roleName)
    $role = $servicePrincipal.appRoles | Where-Object { $_.value -eq $roleName }
    if (-not $role) { throw "App role '$roleName' not found on service principal $($servicePrincipal.appId)" }
    return $role.id
}

try {
    $graphPermNames = @('Group.Read.All','Directory.Read.All','InformationProtectionPolicy.Read.All','Sites.Read.All')
    $spoPermNames   = @('Sites.FullControl.All','User.Read.All')
    $graphRoleIds   = $graphPermNames | ForEach-Object { Get-AppRoleId -servicePrincipal $graphSP -roleName $_ }
    $spoRoleIds     = $spoPermNames   | ForEach-Object { Get-AppRoleId -servicePrincipal $spoSP   -roleName $_ }
    Write-Host "        Resolved $($graphRoleIds.Count) Graph + $($spoRoleIds.Count) SPO role IDs" -ForegroundColor Green
} catch {
    Write-Host "        [!] Role resolution failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ── Step 4: Create the app with cert + requiredResourceAccess ──
Write-Host ""
Write-Host "  [4/6] Creating Entra ID app '$ApplicationName'..." -ForegroundColor Cyan

$appBody = @{
    displayName     = $ApplicationName
    signInAudience  = 'AzureADMyOrg'
    keyCredentials  = @(
        @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = $certBase64
            displayName = "Griffin31 SPO Audit cert"
            endDateTime = $notAfter.ToUniversalTime().ToString('o')
        }
    )
    requiredResourceAccess = @(
        @{
            resourceAppId  = $graphAppId
            resourceAccess = @($graphRoleIds | ForEach-Object { @{ id = $_; type = 'Role' } })
        },
        @{
            resourceAppId  = $spoAppId
            resourceAccess = @($spoRoleIds | ForEach-Object { @{ id = $_; type = 'Role' } })
        }
    )
} | ConvertTo-Json -Depth 6

try {
    $app = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications" -Body $appBody -ContentType "application/json"
    $clientId = $app.appId
    $appObjectId = $app.id
    Write-Host "        Created app $clientId" -ForegroundColor Green
} catch {
    Write-Host "        [!] Create-app failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ── Step 5: Create the service principal for the app ──
Write-Host ""
Write-Host "  [5/6] Creating service principal + granting admin consent..." -ForegroundColor Cyan
try {
    $spBody = @{ appId = $clientId } | ConvertTo-Json
    $sp = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" -Body $spBody -ContentType "application/json"
    $spObjectId = $sp.id
    Write-Host "        Service principal: $spObjectId" -ForegroundColor Green
} catch {
    Write-Host "        [!] SP creation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Grant admin consent: create appRoleAssignment for each permission (app-only)
function Grant-AdminConsent {
    param([string]$principalId, [string]$resourceId, [string[]]$roleIds, [string]$label)
    foreach ($rid in $roleIds) {
        $body = @{
            principalId = $principalId
            resourceId  = $resourceId
            appRoleId   = $rid
        } | ConvertTo-Json
        try {
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments" `
                -Body $body -ContentType "application/json" | Out-Null
        } catch {
            # Assignment may already exist — log & continue
            Write-Host "        (note) $label role $rid : $($_.Exception.Message)" -ForegroundColor DarkGray
        }
    }
}

Grant-AdminConsent -principalId $spObjectId -resourceId $graphSP.id -roleIds $graphRoleIds -label "Graph"
Grant-AdminConsent -principalId $spObjectId -resourceId $spoSP.id   -roleIds $spoRoleIds   -label "SPO"
Write-Host "        Admin consent granted for all app roles" -ForegroundColor Green

# ── Step 6: Save config ──
Write-Host ""
Write-Host "  [6/6] Saving config..." -ForegroundColor Cyan

$encryptedPassword = $securePassword | ConvertFrom-SecureString

$config = @{
    TenantDomain          = $TenantDomain
    TenantId              = $tenantId
    ClientId              = $clientId
    ApplicationObjectId   = $appObjectId
    ServicePrincipalId    = $spObjectId
    ApplicationName       = $ApplicationName
    CertificatePath       = $pfxPath
    CertificateThumbprint = $thumbprint
    EncryptedCertPassword = $encryptedPassword
    RegisteredAt          = (Get-Date).ToString("o")
    AuthMode              = "CertificateAppOnly"
}
$config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8

Write-Host ""
Write-Host "  Setup complete." -ForegroundColor Green
Write-Host "    App name:    $ApplicationName" -ForegroundColor Gray
Write-Host "    ClientId:    $clientId" -ForegroundColor Gray
Write-Host "    Cert:        $pfxPath" -ForegroundColor Gray
Write-Host "    Thumbprint:  $thumbprint" -ForegroundColor Gray
Write-Host "    Config:      $ConfigPath" -ForegroundColor Gray
Write-Host ""
Write-Host "  Note: token propagation takes ~60 seconds. If the first export fails with" -ForegroundColor Yellow
Write-Host "  'unauthorized' or 'insufficient privileges', wait 1-2 minutes and retry." -ForegroundColor Yellow

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
exit 0
