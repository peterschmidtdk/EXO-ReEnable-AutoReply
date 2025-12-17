<#
.SYNOPSIS
  ReEnabled-ExO-AutoReply - "re-enable" (toggle) Out of Office (Automatic Replies) for mailboxes that ALREADY have OOF enabled.

.DESCRIPTION
  This script connects to Exchange Online using App-Only authentication and then “re-enables” OOF by toggling:
     Enabled/Scheduled -> Disabled -> Enabled/Scheduled (restoring the same config)

  IMPORTANT: The script NEVER enables OOF for a mailbox that is Disabled at the time of processing.

  Target modes (set in $Settings.TargetMode):
    - 'AllEnabled' : scans all mailboxes (UserMailbox, optionally SharedMailbox) and processes ONLY those with OOF enabled/scheduled
    - 'CsvEnabled' : reads identities from CSV and processes ONLY those that currently have OOF enabled/scheduled

  CSV format:
    Must contain one of these columns: Identity OR UserPrincipalName OR PrimarySmtpAddress

.AUTHENTICATION MODES
  Set $Settings.AuthMode:
    - "Certificate"  (recommended for scheduled/unattended)
      Uses: Connect-ExchangeOnline -AppId -CertificateThumbprint -Organization
      Requirements:
        - A certificate in the configured certificate store (default LocalMachine\My)
        - The public certificate must be uploaded to the App Registration

    - "ClientSecret" (fallback)
      Uses: OAuth2 client credentials to obtain an access token and then:
        Connect-ExchangeOnline -AccessToken

.APP REGISTRATION REQUIREMENTS (Entra ID / Azure AD)
  You must create an App Registration and configure:
   - API Permission: "Office 365 Exchange Online" -> Application permission -> Exchange.ManageAsApp
   - Grant admin consent for the tenant
   - Add Certificate credential (recommended) OR create a Client Secret (fallback)
   - Assign permissions/roles:
       Option A: assign a supported Entra role to the app (e.g., Exchange Administrator)
       Option B: create an Exchange service principal and add it to an Exchange role group

.SCHEDULING RECOMMENDATION (Task Scheduler)
  - Recommended: run as a dedicated service account OR SYSTEM (if cert is in LocalMachine\My)
  - Create a scheduled task running:
      powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\ReEnabled-ExO-AutoReply.ps1"
  - For scheduled runs set:
      $Settings.ScheduledMode = $true
    This suppresses interactive progress bars and keeps output minimal (logs still written).
  - Secret/cert handling:
      - Certificate auth is preferred for scheduled tasks.
      - If using ClientSecret, store it as a Machine environment variable (default behavior):
          EXO_AUTOREPLY_CLIENT_SECRET

.LOGGING
  - Writes to a single append-only log file:
      .\Logs\ReEnabled-ExO-AutoReply.log
  - Never overwrites the logfile.
  - Every entry is timestamped.

.VERSION
  Version : 1.1.1
  Updated : 2025-12-16
  Author  : Peter Schmidt
#>

# ----------------------------
# Settings (EDIT THESE)
# ----------------------------
$ScriptVersion = "1.1.0"

$Settings = [ordered]@{
    # Run mode
    ScheduledMode          = $false   # $false = interactive (progress on screen); $true = scheduled (quiet)

    # Target selection
    TargetMode             = "AllEnabled"  # "AllEnabled" or "CsvEnabled"
    IncludeSharedMailboxes = $false
    CsvPath                = ".\OOF-Users.csv"   # used when TargetMode=CsvEnabled

    # Auth mode
    AuthMode               = "Certificate" # "Certificate" or "ClientSecret"

    # Tenant / Org
    TenantId               = "YOUR-TENANT-ID-GUID-OR-domain.onmicrosoft.com"
    Organization           = "YOURTENANT.onmicrosoft.com"   # recommended for -Organization
    ClientId               = "YOUR-APP-CLIENT-ID-GUID"

    # Certificate auth
    CertificateThumbprint  = "PASTE-THUMBPRINT-HERE"
    CertificateStorePath   = "Cert:\LocalMachine\My"

    # ClientSecret auth (fallback)
    ClientSecretEnvVar     = "EXO_AUTOREPLY_CLIENT_SECRET"  # preferred for scheduled tasks
    ClientSecretPlain      = ""                             # optional fallback (NOT recommended)

    # Behavior
    ToggleDelaySeconds     = 2          # time between disabling and restoring OOF
    PerMailboxDelayMs      = 150        # light throttling delay
    WhatIfMode             = $false     # set $true to simulate changes

    # Logging
    LogDirectory           = ".\Logs"
    LogFileName            = "ReEnabled-ExO-AutoReply.log"   # single file, append-only
}

# ----------------------------
# Helpers
# ----------------------------
function Initialize-Logging {
    if (-not (Test-Path $Settings.LogDirectory)) {
        New-Item -Path $Settings.LogDirectory -ItemType Directory -Force | Out-Null
    }
    $script:LogPath = Join-Path $Settings.LogDirectory $Settings.LogFileName
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $script:LogPath -Value $line

    if (-not $Settings.ScheduledMode) {
        switch ($Level) {
            "ERROR" { Write-Host $line -ForegroundColor Red }
            "WARN"  { Write-Host $line -ForegroundColor Yellow }
            default { Write-Host $line }
        }
    }
}

function Test-ModuleInstalled {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Get-ModuleHighestVersion {
    param([Parameter(Mandatory)][string]$Name)
    $m = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    return $m.Version
}

function Ensure-Modules {
    if (-not (Test-ModuleInstalled -Name "ExchangeOnlineManagement")) {
        Write-Log "Required module missing: ExchangeOnlineManagement" "ERROR"
        Write-Log "Install suggestion: Install-Module ExchangeOnlineManagement -Scope AllUsers" "INFO"
        throw "ExchangeOnlineManagement is not installed."
    }

    $ver = Get-ModuleHighestVersion -Name "ExchangeOnlineManagement"
    Write-Log "ExchangeOnlineManagement detected. Version: $ver" "INFO"
}

function Get-ClientSecret {
    if (-not [string]::IsNullOrWhiteSpace($Settings.ClientSecretPlain)) {
        return $Settings.ClientSecretPlain
    }

    $envSecret = [Environment]::GetEnvironmentVariable($Settings.ClientSecretEnvVar, "Machine")
    if ([string]::IsNullOrWhiteSpace($envSecret)) {
        $envSecret = [Environment]::GetEnvironmentVariable($Settings.ClientSecretEnvVar, "User")
    }

    if (-not [string]::IsNullOrWhiteSpace($envSecret)) {
        return $envSecret
    }

    if (-not $Settings.ScheduledMode) {
        $sec = Read-Host "Client Secret not found in env var '$($Settings.ClientSecretEnvVar)'. Enter Client Secret" -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }

    throw "Client secret not provided. Set env var '$($Settings.ClientSecretEnvVar)' or populate Settings.ClientSecretPlain."
}

function Get-AccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantIdOrDomain,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $scope = "https://outlook.office365.com/.default"
    $tokenUri = "https://login.microsoftonline.com/$TenantIdOrDomain/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
        scope         = $scope
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
        return $resp.access_token
    }
    catch {
        throw "Failed to obtain access token: $($_.Exception.Message)"
    }
}

function Test-CertificateByThumbprint {
    param(
        [Parameter(Mandatory)][string]$Thumbprint,
        [Parameter(Mandatory)][string]$StorePath
    )

    try {
        $cert = Get-ChildItem -Path $StorePath -ErrorAction Stop | Where-Object { $_.Thumbprint -eq $Thumbprint } | Select-Object -First 1
        return $cert
    }
    catch {
        throw "Could not read certificate store path '$StorePath': $($_.Exception.Message)"
    }
}

function Connect-EXOAppOnly {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if ($Settings.AuthMode -eq "Certificate") {
        if ([string]::IsNullOrWhiteSpace($Settings.CertificateThumbprint)) {
            throw "AuthMode=Certificate requires Settings.CertificateThumbprint."
        }

        $cert = Test-CertificateByThumbprint -Thumbprint $Settings.CertificateThumbprint -StorePath $Settings.CertificateStorePath
        if (-not $cert) {
            throw "Certificate with thumbprint '$($Settings.CertificateThumbprint)' not found in '$($Settings.CertificateStorePath)'."
        }

        Write-Log "Connecting to Exchange Online (app-only certificate)..." "INFO"
        if ($Settings.WhatIfMode) {
            Write-Log "WhatIfMode: would Connect-ExchangeOnline using certificate thumbprint." "INFO"
            return
        }

        Connect-ExchangeOnline `
            -Organization $Settings.Organization `
            -AppId $Settings.ClientId `
            -CertificateThumbprint $Settings.CertificateThumbprint `
            -ShowBanner:$false `
            -ErrorAction Stop

        return
    }

    if ($Settings.AuthMode -eq "ClientSecret") {
        $secret = Get-ClientSecret
        $token  = Get-AccessToken -TenantIdOrDomain $Settings.TenantId -ClientId $Settings.ClientId -ClientSecret $secret

        Write-Log "Connecting to Exchange Online (app-only client secret via access token)..." "INFO"
        if ($Settings.WhatIfMode) {
            Write-Log "WhatIfMode: would Connect-ExchangeOnline using access token." "INFO"
            return
        }

        Connect-ExchangeOnline `
            -Organization $Settings.Organization `
            -AppId $Settings.ClientId `
            -AccessToken $token `
            -ShowBanner:$false `
            -ErrorAction Stop

        return
    }

    throw "Invalid AuthMode: $($Settings.AuthMode). Use Certificate or ClientSecret."
}

function Disconnect-EXOQuiet {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
}

function Get-Targets_AllEnabled {
    $types = @("UserMailbox")
    if ($Settings.IncludeSharedMailboxes) { $types += "SharedMailbox" }

    Write-Log "Resolving mailboxes (RecipientTypeDetails: $($types -join ', '))..." "INFO"
    $mbxs = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $types -Properties UserPrincipalName

    $targets = New-Object System.Collections.Generic.List[string]
    $i = 0
    $total = $mbxs.Count

    foreach ($m in $mbxs) {
        $i++
        if (-not $Settings.ScheduledMode) {
            Write-Progress -Activity "Scanning for Enabled/Scheduled OOF" -Status "$i / $total" -PercentComplete ([int](($i/$total)*100))
        }

        $id = $m.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        try {
            $cfg = Get-MailboxAutoReplyConfiguration -Identity $id -ErrorAction Stop
            if ($cfg.AutoReplyState -ne "Disabled") {
                $targets.Add($id)
            }
        }
        catch {
            Write-Log "Scan failed for $id: $($_.Exception.Message)" "WARN"
        }

        if ($Settings.PerMailboxDelayMs -gt 0) { Start-Sleep -Milliseconds $Settings.PerMailboxDelayMs }
    }

    if (-not $Settings.ScheduledMode) { Write-Progress -Activity "Scanning for Enabled/Scheduled OOF" -Completed }
    return $targets
}

function Get-Targets_CsvEnabled {
    if (-not (Test-Path $Settings.CsvPath)) {
        throw "CSV not found: $($Settings.CsvPath)"
    }

    $rows = Import-Csv -Path $Settings.CsvPath
    if (-not $rows -or $rows.Count -eq 0) {
        throw "CSV is empty: $($Settings.CsvPath)"
    }

    $ids = foreach ($r in $rows) {
        if ($r.Identity) { $r.Identity }
        elseif ($r.UserPrincipalName) { $r.UserPrincipalName }
        elseif ($r.PrimarySmtpAddress) { $r.PrimarySmtpAddress }
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

    if (-not $ids -or $ids.Count -eq 0) {
        throw "No usable Identity/UserPrincipalName/PrimarySmtpAddress values found in CSV."
    }

    $targets = New-Object System.Collections.Generic.List[string]
    $i = 0
    $total = $ids.Count

    foreach ($id in $ids) {
        $i++
        if (-not $Settings.ScheduledMode) {
            Write-Progress -Activity "Validating CSV targets (only Enabled/Scheduled)" -Status "$i / $total" -PercentComplete ([int](($i/$total)*100))
        }

        try {
            $cfg = Get-MailboxAutoReplyConfiguration -Identity $id -ErrorAction Stop
            if ($cfg.AutoReplyState -ne "Disabled") {
                $targets.Add($id)
            }
            else {
                Write-Log "Skipping $id (OOF is Disabled; script will not enable it)." "INFO"
            }
        }
        catch {
            Write-Log "CSV target check failed for $id: $($_.Exception.Message)" "WARN"
        }

        if ($Settings.PerMailboxDelayMs -gt 0) { Start-Sleep -Milliseconds $Settings.PerMailboxDelayMs }
    }

    if (-not $Settings.ScheduledMode) { Write-Progress -Activity "Validating CSV targets (only Enabled/Scheduled)" -Completed }
    return $targets
}

function Toggle-ReEnableOOF {
    param([Parameter(Mandatory)][string]$Identity)

    $cfg = Get-MailboxAutoReplyConfiguration -Identity $Identity -ErrorAction Stop
    $state = [string]$cfg.AutoReplyState

    if ($state -eq "Disabled") {
        Write-Log "Skipping $Identity (OOF is Disabled; will not enable)." "INFO"
        return
    }

    # Snapshot (do not log message contents)
    $snap = [ordered]@{
        AutoReplyState  = $cfg.AutoReplyState
        StartTime       = $cfg.StartTime
        EndTime         = $cfg.EndTime
        ExternalAudience= $cfg.ExternalAudience
        InternalMessage = $cfg.InternalMessage
        ExternalMessage = $cfg.ExternalMessage

        DeclineAllEventsForScheduledOOF = $cfg.DeclineAllEventsForScheduledOOF
        DeclineEventsForScheduledOOF    = $cfg.DeclineEventsForScheduledOOF
        DeclineMeetingMessage           = $cfg.DeclineMeetingMessage
    }

    Write-Log "Re-enabling OOF for $Identity (current state: $state)..." "INFO"

    if ($Settings.WhatIfMode) {
        Write-Log "WhatIfMode: would set Disabled then restore state '$state' for $Identity" "INFO"
        return
    }

    # 1) Disable
    Set-MailboxAutoReplyConfiguration -Identity $Identity -AutoReplyState Disabled -Confirm:$false -ErrorAction Stop

    if ($Settings.ToggleDelaySeconds -gt 0) { Start-Sleep -Seconds $Settings.ToggleDelaySeconds }

    # 2) Restore
    $restoreParams = @{
        Identity         = $Identity
        AutoReplyState   = $snap.AutoReplyState
        InternalMessage  = $snap.InternalMessage
        ExternalMessage  = $snap.ExternalMessage
        ExternalAudience = $snap.ExternalAudience
        Confirm          = $false
        ErrorAction      = "Stop"
    }

    if ($snap.AutoReplyState -eq "Scheduled") {
        if ($snap.StartTime) { $restoreParams.StartTime = [datetime]$snap.StartTime }
        if ($snap.EndTime)   { $restoreParams.EndTime   = [datetime]$snap.EndTime }

        if ($null -ne $snap.DeclineAllEventsForScheduledOOF -and $snap.DeclineAllEventsForScheduledOOF -ne "") {
            $restoreParams.DeclineAllEventsForScheduledOOF = [bool]$snap.DeclineAllEventsForScheduledOOF
        }
        if ($null -ne $snap.DeclineEventsForScheduledOOF -and $snap.DeclineEventsForScheduledOOF -ne "") {
            $restoreParams.DeclineEventsForScheduledOOF = [bool]$snap.DeclineEventsForScheduledOOF
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$snap.DeclineMeetingMessage)) {
            $restoreParams.DeclineMeetingMessage = $snap.DeclineMeetingMessage
        }
    }

    Set-MailboxAutoReplyConfiguration @restoreParams
    Write-Log "OOF re-enabled for $Identity (restored state: $state)." "INFO"
}

# ----------------------------
# Main
# ----------------------------
Initialize-Logging

# scheduled mode: suppress progress globally
if ($Settings.ScheduledMode) {
    $ProgressPreference = "SilentlyContinue"
}

Write-Log "=== START ReEnabled-ExO-AutoReply v$ScriptVersion ===" "INFO"
Write-Log "Mode: TargetMode=$($Settings.TargetMode), ScheduledMode=$($Settings.ScheduledMode), WhatIf=$($Settings.WhatIfMode), AuthMode=$($Settings.AuthMode)" "INFO"

try {
    Ensure-Modules
    Connect-EXOAppOnly

    $targets =
        if ($Settings.TargetMode -eq "AllEnabled") { Get-Targets_AllEnabled }
        elseif ($Settings.TargetMode -eq "CsvEnabled") { Get-Targets_CsvEnabled }
        else { throw "Invalid TargetMode: $($Settings.TargetMode). Use AllEnabled or CsvEnabled." }

    Write-Log "Targets to process (Enabled/Scheduled only): $($targets.Count)" "INFO"

    $i = 0
    $total = $targets.Count

    foreach ($id in $targets) {
        $i++
        if (-not $Settings.ScheduledMode) {
            Write-Progress -Activity "Re-enabling OOF (toggle)" -Status "$i / $total : $id" -PercentComplete ([int](($i/$total)*100))
        }

        try {
            Toggle-ReEnableOOF -Identity $id
        }
        catch {
            Write-Log "FAILED for $id: $($_.Exception.Message)" "ERROR"
        }

        if ($Settings.PerMailboxDelayMs -gt 0) { Start-Sleep -Milliseconds $Settings.PerMailboxDelayMs }
    }

    if (-not $Settings.ScheduledMode) { Write-Progress -Activity "Re-enabling OOF (toggle)" -Completed }

    Write-Log "Run complete. Processed: $total" "INFO"
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    throw
}
finally {
    Disconnect-EXOQuiet
    Write-Log "=== END ReEnabled-ExO-AutoReply ===" "INFO"
}
