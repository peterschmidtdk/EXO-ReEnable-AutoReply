<#
.SYNOPSIS
  Creates the App Registration needed for ReEnabled-ExO-AutoReply and configures permissions + certificate.

.DESCRIPTION
  Uses Microsoft Graph PowerShell to:
   - Create an app registration + service principal
   - Add the application permission "Office 365 Exchange Online" -> Exchange.ManageAsApp
   - Grant admin consent (app role assignment)
   - CREATE a self-signed certificate in the local certificate store (default: LocalMachine\My)
   - Export CER (public) and PFX (private key) to disk
   - Upload the public certificate to the App Registration (KeyCredentials)

  OPTIONAL:
   - Connects to Exchange Online interactively and adds the service principal to an EXO role group
     using Add-RoleGroupMember -MemberType ServicePrincipal (per Microsoft app-only auth guidance).

.REQUIREMENTS
  - Run as an admin who can create apps and grant admin consent (typically Global Admin / Application Admin)
  - Microsoft.Graph module installed.
  - For certificate creation: Windows with New-SelfSignedCertificate cmdlet available.

.VERSION
  Version : 1.1.0
  Updated : 2025-12-15
  Author  : Peter Schmidt
#>

# ----------------------------
# Settings (EDIT THESE)
# ----------------------------
$ScriptVersion = "1.1.0"

$Config = [ordered]@{
    AppDisplayName              = "ReEnabled-ExO-AutoReply"

    # Exchange app-only permission
    ExoResourceAppId            = "00000002-0000-0ff1-ce00-000000000000" # Office 365 Exchange Online

    # Optional secret (fallback auth). Certificate auth is recommended for unattended.
    CreateClientSecret          = $true
    SecretDisplayName           = "ReEnabled-ExO-AutoReply Secret"
    SecretValidityDays          = 365

    # Certificate settings
    CreateCertificate           = $true
    CertSubjectCN              = "ReEnabled-ExO-AutoReply"
    CertValidityDays           = 730
    CertStoreLocation           = "Cert:\LocalMachine\My" # recommended for scheduled tasks running as SYSTEM
    CertExportDirectory         = ".\Cert"
    CertPfxFileName             = "ReEnabled-ExO-AutoReply.pfx"
    CertCerFileName             = "ReEnabled-ExO-AutoReply.cer"
    CertPfxPasswordEnvVar       = "EXO_AUTOREPLY_CERT_PFX_PASSWORD"  # optional (if empty, prompt)
    CertPfxPasswordPlain        = ""                                  # optional fallback (NOT recommended)

    # Optional: Add app SP to this EXO role group (high privilege if Organization Management)
    ConfigureExchangeRoleGroup  = $true
    ExchangeRoleGroupName       = "Organization Management"

    # For EXO interactive connection (only used if ConfigureExchangeRoleGroup=$true)
    ExoAdminUPN                 = ""  # admin@contoso.com (leave blank to prompt)
}

function Test-ModuleInstalled {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Ensure-GraphModule {
    if (-not (Test-ModuleInstalled -Name "Microsoft.Graph")) {
        Write-Host "Missing module: Microsoft.Graph"
        Write-Host "Install suggestion: Install-Module Microsoft.Graph -Scope AllUsers"
        throw "Microsoft.Graph module not installed."
    }
}

function Ensure-EXOModuleIfNeeded {
    if ($Config.ConfigureExchangeRoleGroup -and -not (Test-ModuleInstalled -Name "ExchangeOnlineManagement")) {
        Write-Host "Missing module: ExchangeOnlineManagement"
        Write-Host "Install suggestion: Install-Module ExchangeOnlineManagement -Scope AllUsers"
        throw "ExchangeOnlineManagement module not installed."
    }
}

function Ensure-CertExportDir {
    if (-not (Test-Path $Config.CertExportDirectory)) {
        New-Item -Path $Config.CertExportDirectory -ItemType Directory -Force | Out-Null
    }
}

function Get-PfxPasswordSecureString {
    if (-not [string]::IsNullOrWhiteSpace($Config.CertPfxPasswordPlain)) {
        return (ConvertTo-SecureString -String $Config.CertPfxPasswordPlain -AsPlainText -Force)
    }

    $envPwd = [Environment]::GetEnvironmentVariable($Config.CertPfxPasswordEnvVar, "Machine")
    if ([string]::IsNullOrWhiteSpace($envPwd)) {
        $envPwd = [Environment]::GetEnvironmentVariable($Config.CertPfxPasswordEnvVar, "User")
    }
    if (-not [string]::IsNullOrWhiteSpace($envPwd)) {
        return (ConvertTo-SecureString -String $envPwd -AsPlainText -Force)
    }

    # Prompt if not set
    return (Read-Host "Enter PFX export password (will not echo)" -AsSecureString)
}

function New-AndExportCertificate {
    # Creates self-signed cert and exports CER+PFX
    Ensure-CertExportDir

    $subject = "CN=$($Config.CertSubjectCN)"
    $notAfter = (Get-Date).AddDays([int]$Config.CertValidityDays)

    Write-Host "Creating self-signed certificate in $($Config.CertStoreLocation): $subject"
    $cert = New-SelfSignedCertificate `
        -Subject $subject `
        -CertStoreLocation $Config.CertStoreLocation `
        -KeySpec Signature `
        -KeyExportPolicy Exportable `
        -NotAfter $notAfter `
        -HashAlgorithm "SHA256" `
        -KeyLength 2048

    if (-not $cert) { throw "Failed to create certificate." }

    $cerPath = Join-Path $Config.CertExportDirectory $Config.CertCerFileName
    $pfxPath = Join-Path $Config.CertExportDirectory $Config.CertPfxFileName

    Write-Host "Exporting CER (public) to: $cerPath"
    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null

    $pfxPwd = Get-PfxPasswordSecureString
    Write-Host "Exporting PFX (private key) to: $pfxPath"
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPwd -Force | Out-Null

    return [pscustomobject]@{
        Cert        = $cert
        CerPath     = $cerPath
        PfxPath     = $pfxPath
        Thumbprint  = $cert.Thumbprint
        NotAfter    = $cert.NotAfter
    }
}

function Add-CertificateToAppKeyCredentials {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)]$CertObject,
        [Parameter(Mandatory)][string]$DisplayName
    )

    # Build a KeyCredential payload (public cert only)
    $keyId = [guid]::NewGuid()
    $keyB64 = [Convert]::ToBase64String($CertObject.RawData)

    # Preserve existing keyCredentials (don’t overwrite)
    $appFull = Get-MgApplication -ApplicationId $Application.Id -Property "keyCredentials"
    $existing = @()
    if ($appFull.KeyCredentials) { $existing = @($appFull.KeyCredentials) }

    $newKey = @{
        keyId = $keyId
        type  = "AsymmetricX509Cert"
        usage = "Verify"
        key   = $keyB64
        displayName   = $DisplayName
        startDateTime = (Get-Date).ToUniversalTime()
        endDateTime   = ($CertObject.NotAfter).ToUniversalTime()
    }

    $all = @($existing + $newKey)

    Write-Host "Uploading public certificate to App Registration (KeyCredentials)..."
    Update-MgApplication -ApplicationId $Application.Id -KeyCredentials $all
}

# ----------------------------
# Main
# ----------------------------
Write-Host "=== START Install-ReEnabled-ExO-AutoReplyAppRegistration v$ScriptVersion ==="

Ensure-GraphModule
Ensure-EXOModuleIfNeeded
Import-Module Microsoft.Graph -ErrorAction Stop

# Scopes needed to create apps + assign app roles (admin consent)
$scopes = @(
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "Directory.ReadWrite.All"
)

Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes $scopes | Out-Null

# 1) Create App Registration
Write-Host "Creating app registration: $($Config.AppDisplayName)"
$app = New-MgApplication -DisplayName $Config.AppDisplayName

# 2) Create Service Principal for the app
Write-Host "Creating service principal..."
$sp = New-MgServicePrincipal -AppId $app.AppId

# 3) Find Exchange Online resource SP + the Exchange.ManageAsApp app role
$exoSp = Get-MgServicePrincipal -Filter "appId eq '$($Config.ExoResourceAppId)'"
if (-not $exoSp) { throw "Could not find the Office 365 Exchange Online service principal (appId $($Config.ExoResourceAppId))." }

$appRole = $exoSp.AppRoles | Where-Object {
    $_.Value -eq "Exchange.ManageAsApp" -and $_.AllowedMemberTypes -contains "Application"
}
if (-not $appRole) { throw "Could not find app role Exchange.ManageAsApp on the Exchange Online service principal." }

# 4) Add requiredResourceAccess to the application
$required = @(
    @{
        resourceAppId  = $exoSp.AppId
        resourceAccess = @(
            @{
                id   = $appRole.Id
                type = "Role"
            }
        )
    }
)

Write-Host "Updating app requiredResourceAccess (Exchange.ManageAsApp)..."
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $required

# 5) Grant admin consent (app role assignment)
Write-Host "Granting admin consent via appRoleAssignment..."
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $exoSp.Id -AppRoleId $appRole.Id | Out-Null

# 6) Create and upload certificate
$certInfo = $null
if ($Config.CreateCertificate) {
    $certInfo = New-AndExportCertificate
    Add-CertificateToAppKeyCredentials -Application $app -CertObject $certInfo.Cert -DisplayName "ReEnabled-ExO-AutoReply Cert"
}

# 7) Create client secret (optional)
$clientSecret = $null
$secretEnd = $null
if ($Config.CreateClientSecret) {
    Write-Host "Creating client secret (valid $($Config.SecretValidityDays) days)..."
    $secretEnd = (Get-Date).AddDays([int]$Config.SecretValidityDays)
    $pwd = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        displayName = $Config.SecretDisplayName
        endDateTime = $secretEnd
    }
    $clientSecret = $pwd.SecretText
}

# OPTIONAL: Add SP to an Exchange role group
if ($Config.ConfigureExchangeRoleGroup) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $adminUpn = $Config.ExoAdminUPN
    if ([string]::IsNullOrWhiteSpace($adminUpn)) {
        $adminUpn = Read-Host "Enter EXO admin UPN to connect for role group assignment"
    }

    Write-Host "Connecting to Exchange Online interactively as $adminUpn..."
    Connect-ExchangeOnline -UserPrincipalName $adminUpn -ShowBanner:$false

    $exoSpName = "SP for Azure AD App $($Config.AppDisplayName)"
    Write-Host "Creating EXO service principal object: $exoSpName"
    New-ServicePrincipal -AppId $app.AppId -ObjectId $sp.Id -DisplayName $exoSpName | Out-Null

    Write-Host "Adding service principal to EXO role group: $($Config.ExchangeRoleGroupName)"
    Add-RoleGroupMember -Identity $Config.ExchangeRoleGroupName -MemberType ServicePrincipal -Member $sp.Id

    Disconnect-ExchangeOnline -Confirm:$false
}

Disconnect-MgGraph | Out-Null

Write-Host ""
Write-Host "=== OUTPUT ==="
Write-Host "App Display Name : $($Config.AppDisplayName)"
Write-Host "Client ID (AppId): $($app.AppId)"
Write-Host "SP Object ID     : $($sp.Id)"
if ($certInfo) {
    Write-Host "Certificate Thumbprint : $($certInfo.Thumbprint)"
    Write-Host "Certificate Expires    : $($certInfo.NotAfter.ToString('yyyy-MM-dd'))"
    Write-Host "Exported CER            : $($certInfo.CerPath)"
    Write-Host "Exported PFX            : $($certInfo.PfxPath)"
    Write-Host "NOTE: For scheduled tasks, keep the cert in LocalMachine\\My on the server that runs the main script."
}
if ($clientSecret) {
    Write-Host "Client Secret    : $clientSecret"
    Write-Host "Secret Expires   : $($secretEnd.ToString('yyyy-MM-dd'))"
}
Write-Host ""
Write-Host "=== END Install-ReEnabled-ExO-AutoReplyAppRegistration ==="
