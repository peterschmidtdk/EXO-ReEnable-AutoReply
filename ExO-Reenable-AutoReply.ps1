<# 
.SYNOPSIS
  Bulk disable/reset/restore Automatic Replies (Out of Office) for Exchange Online mailboxes.

.DESCRIPTION
  Supports:
    - Disable: Disables auto-reply for target mailboxes and exports a snapshot (only Enabled/Scheduled mailboxes).
    - Restore: Restores auto-reply settings from a snapshot CSV.
    - Reset: Disables then restores in the same run (also exports a snapshot).

  Auth:
    - Certificate (recommended by Microsoft for unattended EXO): Connect-ExchangeOnline -AppId -CertificateThumbprint -Organization
    - ClientSecret (token + AccessToken): Get OAuth token via client credentials and pass -AccessToken to Connect-ExchangeOnline.

.NOTES
  Author  : Peter
  Script  : Reset-EXOAutoReply.ps1
  Version : 1.0.0
  Updated : 2025-12-15
  Requires: ExchangeOnlineManagement module (v3.1.0+ for -AccessToken)

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Disable','Restore','Reset')]
    [string]$Action,

    # Targeting
    [Parameter(Mandatory=$false)]
    [switch]$AllMailboxes,

    [Parameter(Mandatory=$false)]
    [string]$InputCsvPath,   # for selected users (Identity/UserPrincipalName)

    # Snapshot
    [Parameter(Mandatory=$false)]
    [string]$SnapshotCsvPath = ".\AutoReply-Snapshot.csv",

    # Optional delay between disable and restore (Reset action)
    [Parameter(Mandatory=$false)]
    [ValidateRange(0,300)]
    [int]$ResetDelaySeconds = 2,

    # Mailbox scope
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSharedMailboxes,

    # If set, AllMailboxes will process every mailbox (not just enabled/scheduled) for Disable action
    # (Restore still only restores from snapshot)
    [Parameter(Mandatory=$false)]
    [switch]$DisableEvenIfAlreadyDisabled,

    # Authentication
    [Parameter(Mandatory=$true)]
    [ValidateSet('Certificate','ClientSecret')]
    [string]$AuthMode,

    # Organization: primary .onmicrosoft.com domain or tenant ID is recommended for app-only auth
    [Parameter(Mandatory=$true)]
    [string]$Organization,

    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    # Certificate mode
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint,

    # ClientSecret mode
    [Parameter(Mandatory=$false)]
    [string]$TenantId, # GUID or contoso.onmicrosoft.com
    [Parameter(Mandatory=$false)]
    [string]$ClientSecretPlain
)

# ----------------------------
# Globals / Logging
# ----------------------------
$ScriptVersion = "1.0.0"
$StartTime = Get-Date
$LogDir = ".\Logs"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }

$RunStamp = $StartTime.ToString("yyyyMMdd-HHmmss")
$LogPath = Join-Path $LogDir "EXO-AutoReply-$Action-$RunStamp.log"
$ReportCsvPath = Join-Path $LogDir "EXO-AutoReply-Report-$Action-$RunStamp.csv"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

function Add-ReportRow {
    param(
        [string]$Identity,
        [string]$Result,
        [string]$Details = "",
        [string]$PreviousState = "",
        [string]$NewState = ""
    )
    [pscustomobject]@{
        Timestamp     = (Get-Date).ToString("s")
        Action        = $Action
        Identity      = $Identity
        Result        = $Result
        Details       = $Details
        PreviousState = $PreviousState
        NewState      = $NewState
    }
}

$Report = New-Object System.Collections.Generic.List[object]

# ----------------------------
# Auth helpers
# ----------------------------
function Get-ClientSecretAccessToken {
    param(
        [Parameter(Mandatory=$true)][string]$TenantIdOrDomain,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$ClientSecret,
        [Parameter(Mandatory=$false)][string]$Scope = "https://outlook.office365.com/.default"
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantIdOrDomain/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
        scope         = $Scope
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
    return $resp.access_token
}

function Connect-EXO {
    Write-Log "Connecting to Exchange Online... (AuthMode=$AuthMode, Organization=$Organization, ClientId=$ClientId)"
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if ($AuthMode -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            throw "CertificateThumbprint is required when AuthMode=Certificate."
        }

        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        return
    }

    if ($AuthMode -eq 'ClientSecret') {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            throw "TenantId is required when AuthMode=ClientSecret."
        }
        if ([string]::IsNullOrWhiteSpace($ClientSecretPlain)) {
            throw "ClientSecretPlain is required when AuthMode=ClientSecret."
        }

        $token = Get-ClientSecretAccessToken -TenantIdOrDomain $TenantId -ClientId $ClientId -ClientSecret $ClientSecretPlain
        Connect-ExchangeOnline -Organization $Organization -AppId $ClientId -AccessToken $token -ShowBanner:$false -ErrorAction Stop
        return
    }

    throw "Unknown AuthMode."
}

function Disconnect-EXO {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
}

# ----------------------------
# Target resolution
# ----------------------------
function Resolve-TargetIdentities {
    # Returns array of mailbox identities (UPNs/Smtp addresses)
    if ($Action -eq 'Restore') {
        # Restore always driven by snapshot
        if (-not (Test-Path $SnapshotCsvPath)) {
            throw "SnapshotCsvPath not found: $SnapshotCsvPath"
        }
        $snap = Import-Csv -Path $SnapshotCsvPath
        if (-not $snap -or $snap.Count -eq 0) { throw "SnapshotCsvPath is empty: $SnapshotCsvPath" }
        return $snap.Identity | Sort-Object -Unique
    }

    if ($AllMailboxes) {
        $types = @('UserMailbox')
        if ($IncludeSharedMailboxes) { $types += 'SharedMailbox' }

        Write-Log "Resolving mailboxes (RecipientTypeDetails: $($types -join ', '))..."
        $mbxs = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $types -Properties PrimarySmtpAddress,UserPrincipalName
        return ($mbxs | ForEach-Object { $_.UserPrincipalName }) | Where-Object { $_ } | Sort-Object -Unique
    }

    if (-not [string]::IsNullOrWhiteSpace($InputCsvPath)) {
        if (-not (Test-Path $InputCsvPath)) { throw "InputCsvPath not found: $InputCsvPath" }
        $rows = Import-Csv -Path $InputCsvPath
        if (-not $rows -or $rows.Count -eq 0) { throw "InputCsvPath is empty: $InputCsvPath" }

        $ids = foreach ($r in $rows) {
            if ($r.Identity) { $r.Identity }
            elseif ($r.UserPrincipalName) { $r.UserPrincipalName }
        }

        $ids = $ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
        if (-not $ids -or $ids.Count -eq 0) { throw "No Identity/UserPrincipalName values found in $InputCsvPath" }
        return $ids
    }

    throw "You must specify either -AllMailboxes or -InputCsvPath (except for Restore, which uses -SnapshotCsvPath)."
}

# ----------------------------
# AutoReply snapshot helpers
# ----------------------------
function New-SnapshotRowFromConfig {
    param(
        [Parameter(Mandatory=$true)][string]$Identity,
        [Parameter(Mandatory=$true)]$Config
    )
    [pscustomobject]@{
        SnapshotTakenAt                 = (Get-Date).ToString("s")
        Identity                        = $Identity
        AutoReplyState                  = [string]$Config.AutoReplyState
        StartTime                       = if ($Config.StartTime) { (Get-Date $Config.StartTime).ToString("o") } else { "" }
        EndTime                         = if ($Config.EndTime)   { (Get-Date $Config.EndTime).ToString("o") } else { "" }
        ExternalAudience                = [string]$Config.ExternalAudience
        InternalMessage                 = [string]$Config.InternalMessage
        ExternalMessage                 = [string]$Config.ExternalMessage
        DeclineAllEventsForScheduledOOF = [string]$Config.DeclineAllEventsForScheduledOOF
        DeclineEventsForScheduledOOF    = [string]$Config.DeclineEventsForScheduledOOF
        DeclineMeetingMessage           = [string]$Config.DeclineMeetingMessage
    }
}

function Restore-AutoReplyFromSnapshotRow {
    param(
        [Parameter(Mandatory=$true)]$Row
    )

    $id = $Row.Identity
    $targetState = $Row.AutoReplyState

    $params = @{
        Identity        = $id
        AutoReplyState  = $targetState
        InternalMessage = $Row.InternalMessage
        ExternalMessage = $Row.ExternalMessage
        ExternalAudience= $Row.ExternalAudience
        Confirm         = $false
        ErrorAction     = 'Stop'
    }

    if ($targetState -eq 'Scheduled') {
        if ($Row.StartTime) { $params.StartTime = [datetime]::Parse($Row.StartTime) }
        if ($Row.EndTime)   { $params.EndTime   = [datetime]::Parse($Row.EndTime) }
        if ($Row.DeclineAllEventsForScheduledOOF -ne "") { $params.DeclineAllEventsForScheduledOOF = [bool]::Parse($Row.DeclineAllEventsForScheduledOOF) }
        if ($Row.DeclineEventsForScheduledOOF -ne "")    { $params.DeclineEventsForScheduledOOF    = [bool]::Parse($Row.DeclineEventsForScheduledOOF) }
        if ($Row.DeclineMeetingMessage -ne "")           { $params.DeclineMeetingMessage           = $Row.DeclineMeetingMessage }
    }

    Set-MailboxAutoReplyConfiguration @params
}

# ----------------------------
# Main
# ----------------------------
try {
    Write-Log "=== EXO AutoReply Tool v$ScriptVersion starting (Action=$Action) ==="
    Write-Log "Log: $LogPath"
    Write-Log "Report CSV: $ReportCsvPath"
    Write-Log "Snapshot CSV: $SnapshotCsvPath"

    Connect-EXO

    $targets = Resolve-TargetIdentities
    Write-Log "Targets resolved: $($targets.Count)"

    $SnapshotRows = New-Object System.Collections.Generic.List[object]

    if ($Action -in @('Disable','Reset')) {
        foreach ($id in $targets) {
            try {
                $cfg = Get-MailboxAutoReplyConfiguration -Identity $id -ErrorAction Stop

                $prevState = [string]$cfg.AutoReplyState
                $wasEnabled = ($prevState -ne 'Disabled')

                if (-not $wasEnabled -and -not $DisableEvenIfAlreadyDisabled) {
                    $Report.Add((Add-ReportRow -Identity $id -Result "SKIPPED" -Details "Already Disabled" -PreviousState $prevState -NewState $prevState))
                    continue
                }

                if ($wasEnabled) {
                    $SnapshotRows.Add((New-SnapshotRowFromConfig -Identity $id -Config $cfg))
                }

                Set-MailboxAutoReplyConfiguration -Identity $id -AutoReplyState Disabled -Confirm:$false -ErrorAction Stop

                $Report.Add((Add-ReportRow -Identity $id -Result "OK" -Details "Disabled auto-reply" -PreviousState $prevState -NewState "Disabled"))
            }
            catch {
                $Report.Add((Add-ReportRow -Identity $id -Result "ERROR" -Details $_.Exception.Message))
                Write-Log "Failed on $id (Disable/Reset step): $($_.Exception.Message)" -Level ERROR
            }
        }

        # Export snapshot (only those that had Enabled/Scheduled)
        if ($SnapshotRows.Count -gt 0) {
            $SnapshotRows | Export-Csv -Path $SnapshotCsvPath -NoTypeInformation -Encoding UTF8
            Write-Log "Snapshot exported: $SnapshotCsvPath (rows=$($SnapshotRows.Count))"
        }
        else {
            Write-Log "No Enabled/Scheduled auto-reply configs found to snapshot."
        }
    }

    if ($Action -eq 'Reset' -and $ResetDelaySeconds -gt 0) {
        Write-Log "Waiting $ResetDelaySeconds seconds before restore..."
        Start-Sleep -Seconds $ResetDelaySeconds
    }

    if ($Action -in @('Restore','Reset')) {
        # Load snapshot
        if (-not (Test-Path $SnapshotCsvPath)) {
            throw "SnapshotCsvPath not found: $SnapshotCsvPath"
        }
        $snap = Import-Csv -Path $SnapshotCsvPath
        if (-not $snap -or $snap.Count -eq 0) {
            throw "SnapshotCsvPath is empty: $SnapshotCsvPath"
        }

        foreach ($row in $snap) {
            $id = $row.Identity
            try {
                Restore-AutoReplyFromSnapshotRow -Row $row
                $Report.Add((Add-ReportRow -Identity $id -Result "OK" -Details "Restored from snapshot" -PreviousState "Disabled" -NewState $row.AutoReplyState))
            }
            catch {
                $Report.Add((Add-ReportRow -Identity $id -Result "ERROR" -Details $_.Exception.Message))
                Write-Log "Failed on $id (Restore step): $($_.Exception.Message)" -Level ERROR
            }
        }
    }

    # Export report
    $Report | Export-Csv -Path $ReportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Report exported: $ReportCsvPath"
}
catch {
    Write-Log $_.Exception.Message -Level ERROR
    throw
}
finally {
    Disconnect-EXO
    $dur = New-TimeSpan -Start $StartTime -End (Get-Date)
    Write-Log "=== Completed in $([int]$dur.TotalSeconds)s ==="
}
