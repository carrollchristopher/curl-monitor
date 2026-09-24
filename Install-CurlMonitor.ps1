<#
.SYNOPSIS
    Self-contained installer that deploys and starts an always-on curl monitor for any URL on a site server.

.DESCRIPTION
    Run this one script as administrator on a site server. It does everything with no second file to manage.
    1. Confirms it is running elevated, relaunching as administrator if not, and checks curl.exe and the
       ScheduledTasks module are present.
    2. Asks what to call this monitor, for example "GitHub Monitor", which names its folder, its scheduled task, and
       its alerts. One server can run several monitors, one per URL, side by side. Re-running with the same name
       upgrades that monitor in place and leaves the others alone.
    3. Asks for the URL to watch, the site name, and optional text that must appear on the page, all prefilled
       from the last run. Poll interval, timeouts, and alert thresholds live in the config block below.
    4. Walks through a mail setup wizard. Microsoft Graph is recommended (HTTPS only). On the first site the
       wizard creates the tenant side itself: signs you in, creates the app registration, creates the shared
       sender mailbox (no license), limits the app to sending only as that mailbox (Exchange RBAC for
       Applications, with the legacy application access policy as automatic fallback), and creates a 24-month
       secret. Later sites just paste the Tenant ID, Client ID, and secret. M365 Direct Send, an internal relay, and username and password SMTP
       remain available. A test email is sent and the wizard loops until you confirm it arrived.
    5. Stores the Graph client secret or SMTP password encrypted with machine-scope DPAPI in a file locked to
       SYSTEM and Administrators, so the monitor running as SYSTEM can read it. No plaintext secret is written.
    6. Generates the monitor from an embedded template with every value baked in and validates it parses.
    7. Stops any running copy, writes the monitor, registers a Scheduled Task (SYSTEM, at startup, highest
       privileges, no time limit, single instance, restart on failure), starts it, and verifies it is running.
    8. Runs a preflight probe and sends an install confirmation email.
    An older HST-only install (C:\ProgramData\DIT\HSTProbe with its own task) is migrated on first run: its
    settings and history move into the new layout and the old task and folder are removed.
    Set $NonInteractive to $true and fill in the config block to deploy silently through an RMM.

.REQUIREMENTS
    Administrative rights. Windows 10 1803+ or Windows Server 2019+ (ships with curl.exe and the ScheduledTasks module).
    Outbound SMTP from this server to the chosen mail host.

.OUTPUTS
    Deployed monitor script, saved settings JSON, optional encrypted credential file, and a registered running
    Scheduled Task. The monitor then produces monthly latency CSVs, an outages CSV, a drops log that records only
    failures and outage events, and daily transcript logs.

.NOTES
    Author:      Christopher Carroll
    Created:     09/04/2026
    Idempotency: Safe to re-run. Existing task, monitor, settings, and credential file are replaced cleanly.
    Context:     Written for an HST eChart slowness investigation across shared-resource clinic sites, then
                 generalised: one installer, any URL, one folder and task per monitor. Scheduled Task chosen
                 over a service wrapper to avoid a third-party binary at a healthcare client.
                 Direct Send is recommended because Microsoft disables Basic auth SMTP AUTH by default at the end
                 of December 2026; Direct Send and IP-based relay are not affected.

.LINK
    https://curl.se/docs/manpage.html
#>

Remove-Variable * -ErrorAction SilentlyContinue

# Deployment locations. Each monitor gets its own folder under the root and its own scheduled task.
$InstallRoot           = "C:\ProgramData\CurlMonitor"
$InstallDir            = $InstallRoot                  # Replaced with the monitor's own folder once its name is known
$MonitorFileName       = "Watch-CurlMonitor.ps1"
$SettingsFileName      = "install-settings.json"
$CredentialFileName    = "credential.bin"
$TaskNamePrefix        = "Curl Monitor - "
$MaxTaskNameLength     = 238                           # Task Scheduler limit, checked before anything is written
$TaskName              = $TaskNamePrefix               # Completed with the monitor name once it is known
$TaskPath              = "\CurlMonitor\"
$RunAsUser             = "SYSTEM"
$RestartCount          = 3
$RestartMinutes        = 1

# Earlier layouts, migrated into the one above and then removed
$LegacyInstallDir      = "C:\ProgramData\DIT\HSTProbe"
$LegacyTaskName        = "HST eChart Monitor"
$LegacyMonitorName     = "HST eChart"
$PreviousRoot          = "C:\ProgramData\DIT\CurlMonitor"   # Where monitors lived before the DIT folder was dropped
$PreviousTaskPath      = "\DIT\"

# Prompted every run, prefilled from the last one. In non-interactive mode these are used as-is.
$MonitorName           = ""                            # For example "GitHub Monitor". Names the folder, the task, and the alerts.
$Url                   = ""                            # The URL to watch, http:// or https://
$ExpectedContentMarker = ""                            # Text that must appear on the page. Blank checks only the status and the size.

# Monitor settings baked into the deployed script. Not prompted: edit here before deploying.
$IntervalSeconds       = 10
$TimeoutSeconds        = 15
$MaxRedirects          = 5                             # Redirects the probe follows before giving up
$MinPopulatedBytes     = 1000                          # A 200 with a smaller body counts as a failed poll
$MaxBodyBytes          = 8MB                           # curl stops downloading past this, so one huge reply cannot fill memory
$SlowThresholdMs       = 3000
$DownThreshold         = 3
$ReAlertMinutes        = 30
$AlertsEnabled         = $true                         # $false installs a telemetry-only monitor: no mail wizard, no alerts, data still recorded
$AlertOnRecovery       = $true
$AlertOnSlow           = $true                         # Email when the URL stays slow or keeps failing without a full outage
$SlowWindowMinutes     = 5                             # Rolling window the slow alert looks at
$SlowAlertPercent      = 50                            # Share of polls in the window, slower than SlowThresholdMs or failed, that declares SLOW
$SlowClearPercent      = 10                            # Share at or below which the slow period is over
$DailySummaryHour      = 7                             # Local hour of the daily summary of slow and failed polls, sent only when something happened. -1 turns it off.
$LogRetentionDays      = 30

# Mail defaults. The wizard prompts for these interactively; in non-interactive mode they are used as-is.
$MailMethod            = "Graph"                       # Graph, DirectSend, Relay, or Authenticated
$GraphTenantId         = ""                            # Printed by the tenant setup step on the first site
$GraphClientId         = ""                            # Printed by the tenant setup step on the first site
$GraphSecretExpires    = ""                            # yyyy-MM-dd. Monitor warns 30 days before.
$GraphAppDisplayName   = "Curl Monitor"                # Tenant setup step: app registration name
$GraphPolicyGroupAlias = "curl-monitor-senders"        # Tenant setup step: mail-enabled security group alias, created in the sender's domain
$GraphSecretMonths     = 24                            # Tenant setup step: secret lifetime, Entra maximum is 24
$SmtpServer            = ""                            # Leave blank for DirectSend to auto-discover from the sender domain MX
$SmtpPort              = 25
$SmtpUseSsl            = $false
$MailFrom              = ""                            # Prompted. For Graph this is the shared sender mailbox in the alerting tenant
$MailTo                = @()                           # Prompted. One or more recipients
$SmtpAuthUser          = ""                            # Authenticated only. Password is prompted, never stored here.

# Installer behavior
$Action                = "Install"                     # Install or Uninstall. Interactive runs ask; non-interactive runs use this.
$KeepHistoryOnUninstall = $true                        # Uninstall: $true moves the latency CSVs, outages, and drops log aside first
$HistoryKeepRoot       = "C:\ProgramData\CurlMonitor-history"
$TelemetryTaskName     = "Curl Monitor telemetry publisher"
$NonInteractive        = $false                        # $true = no prompts, use the config block, for RMM deployment
$SiteNameOverride      = ""                            # Used when non-interactive, otherwise prompted
$MonitorNameOverride   = ""                            # Used when non-interactive, otherwise prompted
$SendInstallTestEmail  = $true

# Internal state for the installer's own Graph calls, not settings
$script:LoginBase = 'https://login.microsoftonline.com'   # Overridable so a harness can drive the token path against a local listener
$script:GraphBase = 'https://graph.microsoft.com'
$script:FreshSecretClientId = ''                          # App whose secret this run minted, and the window in which sign-in rejections are retried
$script:FreshSecretCreatedUtc = [datetime]::MinValue
$script:FreshSecretUntilUtc = [datetime]::MinValue
$script:FreshSecretRetrySeconds = 15
$script:PastedSecretRetrySeconds = 60
$script:FreshSecretNoteSeconds = 120                      # A "still waiting" line this often during a long sign-in wait
$script:InstallerToken = $null                            # Last app-only token issued to the installer, reused until near expiry
$script:LastGraphSendError = ''                           # InvalidSecret, AccessDenied, or Other after a failed Graph send
$script:MailTestSkipped = $false                          # Set when the test email failed and the operator chose to finish anyway
$script:RetryTenantSetup = $false                         # Set when a secret minted this run was rejected, so the app prompt defaults to N

# Alert settings normalised once, so the install summary and the generated monitor describe the same behaviour
$DownThreshold         = [math]::Max(1, [int]$DownThreshold)
$SlowWindowMinutes     = [math]::Max(1, [int]$SlowWindowMinutes)
$SlowAlertPercent      = [math]::Min(100, [math]::Max(1, [int]$SlowAlertPercent))
$SlowClearPercent      = [math]::Max(0, [math]::Min([int]$SlowClearPercent, $SlowAlertPercent - 1))
$DailySummaryHour      = [math]::Min(23, [math]::Max(-1, [int]$DailySummaryHour))

$HashTable_PowerShellTranscriptLogsbyKeyword = [Ordered]@{
    ADDED          = "ADDED |"
    CREATED        = "CREATED |"
    ERROR          = "ERROR |"
    FAILED         = "FAILED |"
    FINISHED       = "FINISHED |"
    FOUND          = "FOUND |"
    INFORMATIONAL  = "INFORMATIONAL |"
    PROMPT         = "PROMPT |"
    "SANITY CHECK" = "SANITY CHECK |"
    STARTED        = "STARTED |"
    SUCCESS        = "SUCCESS |"
    TOTAL          = "TOTAL |"
    WARNING        = "WARNING |"
}

function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet('ADDED','CREATED','ERROR','FAILED','FINISHED','FOUND','INFORMATIONAL','PROMPT','SANITY CHECK','STARTED','SUCCESS','TOTAL','WARNING')][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    $prefix = $HashTable_PowerShellTranscriptLogsbyKeyword[$Level]
    Write-Host ("{0} {1} {2}" -f (Get-Date -Format "MM/dd/yy - hh:mm:ss tt"), $prefix, $Message)
}

function Test-ScriptBeingRanAsAdministrator {
    # Ensures the installer runs elevated, relaunching in a window that stays open if needed
    $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
            Write-Log -Level FAILED -Message "Cannot self-elevate because the script path is unknown. Save this as a .ps1 file and run it from disk."
            exit 1
        }
        Write-Log -Level WARNING -Message "Not elevated. Relaunching as administrator."
        Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList "-NoProfile -NoExit -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
    Write-Log -Level "SANITY CHECK" -Message "Running with administrator rights."
}

function Read-Setting {
    # Prompts with a default. In non-interactive mode the default is returned without prompting.
    param([Parameter(Mandatory)][string]$Prompt, [string]$Default = "")
    if ($NonInteractive) { return $Default }
    $shown = if ([string]::IsNullOrWhiteSpace($Default)) { $Prompt } else { "$Prompt [$Default]" }
    $entry = Read-Host $shown
    if ([string]::IsNullOrWhiteSpace($entry)) { return $Default }
    return $entry.Trim()
}

function Read-Choice {
    # Prompts for one of a set of allowed single-character answers. Returns the default when non-interactive.
    param([Parameter(Mandatory)][string]$Prompt, [Parameter(Mandatory)][string[]]$Allowed, [Parameter(Mandatory)][string]$Default)
    if ($NonInteractive) { return $Default }
    while ($true) {
        $entry = Read-Host "$Prompt [$($Allowed -join '/')] (default $Default)"
        if ([string]::IsNullOrWhiteSpace($entry)) { return $Default }
        $entry = $entry.Trim().ToUpper()
        if ($Allowed -contains $entry) { return $entry }
        Write-Log -Level WARNING -Message "Enter one of: $($Allowed -join ', ')."
    }
}

function Write-Question {
    # One wizard question: a blank line, the question, and at most a couple of short lines under it
    param([Parameter(Mandatory)][string]$Question, [string[]]$Hint = @())
    if ($NonInteractive) { return }
    Write-Host ""
    Write-Host $Question -ForegroundColor White
    foreach ($line in @($Hint)) { if ($line) { Write-Host "  $line" -ForegroundColor DarkGray } }
}

function Read-PortSetting {
    # Prompts for a TCP port until a valid one is entered. Returns $null only in non-interactive mode with a bad default.
    param([Parameter(Mandatory)][string]$Prompt, [Parameter(Mandatory)][int]$Default)
    while ($true) {
        $raw = Read-Setting -Prompt $Prompt -Default ([string]$Default)
        $n = 0
        if ([int]::TryParse(([string]$raw).Trim(), [ref]$n) -and $n -ge 1 -and $n -le 65535) { return $n }
        if ($NonInteractive) { Write-Log -Level FAILED -Message "Port '$raw' is not valid in non-interactive mode."; return $null }
        Write-Log -Level WARNING -Message "Port must be a whole number from 1 to 65535."
    }
}

function ConvertTo-SafeSiteName {
    # Keeps letters, numbers, space, dash, underscore, then trims and collapses spaces
    param([string]$Name)
    if ($null -eq $Name) { return "" }
    return (($Name -replace '[^A-Za-z0-9 _-]', '').Trim() -replace '\s+', ' ')
}

function Test-EmailAddress {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    return [bool]($Address.Trim() -match '^[^\s@]+@[^\s@]+\.[^\s@]+$')
}

function ConvertTo-RecipientList {
    # Splits a comma, semicolon, or space separated string into a deduplicated list of valid addresses
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $seen = @{}
    $list = @()
    foreach ($piece in ($Text -split '[,;\s]+')) {
        $p = $piece.Trim()
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (-not (Test-EmailAddress $p)) { continue }
        $key = $p.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $list += $p
    }
    return ,$list
}

function ConvertTo-DerivedDirectSendHost {
    # Derives the standard M365 MX host from a sender domain, used when MX lookup is unavailable
    param([string]$Domain)
    if ([string]::IsNullOrWhiteSpace($Domain)) { return "" }
    return ($Domain.Trim().ToLowerInvariant().Replace('.', '-') + '.mail.protection.outlook.com')
}

function Get-DirectSendHost {
    # Looks up the sender domain's MX record and falls back to the derived M365 host
    param([Parameter(Mandatory)][string]$FromAddress)
    $domain = ($FromAddress -split '@')[-1]
    try {
        $mx = Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop | Where-Object { $_.Type -eq 'MX' } | Sort-Object Preference | Select-Object -First 1
        if ($mx -and $mx.NameExchange) {
            return [PSCustomObject]@{ Host = $mx.NameExchange.TrimEnd('.'); Source = "MX record for $domain" }
        }
    }
    catch { }
    return [PSCustomObject]@{ Host = (ConvertTo-DerivedDirectSendHost $domain); Source = "derived from $domain (MX lookup unavailable)" }
}

function Get-SiteName {
    # Returns the site name from the override, saved settings, or an interactive prompt
    param([string]$SavedDefault = "")
    if (-not [string]::IsNullOrWhiteSpace($SiteNameOverride)) {
        $clean = ConvertTo-SafeSiteName $SiteNameOverride
        if ($clean) { Write-Log -Level FOUND -Message "Site name set from override: '$clean'."; return $clean }
    }
    $default = if ($SavedDefault) { $SavedDefault } else { $env:COMPUTERNAME }
    if ($NonInteractive) {
        $clean = ConvertTo-SafeSiteName $default
        if (-not $clean) { $clean = ConvertTo-SafeSiteName $env:COMPUTERNAME }
        Write-Log -Level FOUND -Message "Site name set to '$clean' (non-interactive)."
        return $clean
    }
    while ($true) {
        Write-Question -Question "Site name" -Hint "Where this server sits. It appears in every alert subject."
        $typed = Read-Setting -Prompt "Site" -Default $default
        $clean = ConvertTo-SafeSiteName $typed
        if (-not $clean) { Write-Log -Level WARNING -Message "Site name cannot be empty after cleanup."; continue }
        if ($clean -ne $typed.Trim()) {
            if ((Read-Choice -Prompt "Cleaned to '$clean'. Use it?" -Allowed @('Y','N') -Default 'Y') -ne 'Y') { continue }
        }
        return $clean
    }
}

function Get-InstallSettings {
    # Loads settings saved by a previous run, or returns $null
    $path = Join-Path $InstallDir $SettingsFileName
    if (-not (Test-Path $path)) { return $null }
    try {
        $json = Get-Content -Path $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Write-Log -Level FOUND -Message "Loaded settings from a previous install. Prompts will default to those values."
        return $json
    }
    catch {
        Write-Log -Level WARNING -Message "Previous settings file is unreadable and will be replaced. $($_.Exception.Message)"
        return $null
    }
}

function Save-InstallSettings {
    # Persists non-secret settings so a re-run prefills every prompt
    param([Parameter(Mandatory)][hashtable]$Settings)
    $path = Join-Path $InstallDir $SettingsFileName
    try {
        $Settings | ConvertTo-Json -Depth 3 | Set-Content -Path $path -Encoding UTF8 -Force
        Write-Log -Level CREATED -Message "Saved settings to '$path'."
    }
    catch {
        Write-Log -Level WARNING -Message "Could not save settings. Re-runs will not prefill. $($_.Exception.Message)"
    }
}

function ConvertTo-SecureText {
    # Builds a read-only SecureString from text that is already in memory, for cmdlets that accept nothing else
    param([Parameter(Mandatory)][string]$Text)
    $secure = New-Object System.Security.SecureString
    foreach ($ch in $Text.ToCharArray()) { $secure.AppendChar($ch) }
    $secure.MakeReadOnly()
    return $secure
}

function Protect-Secret {
    # Encrypts a secret with machine-scope DPAPI so SYSTEM can decrypt it on this machine only. Takes the SecureString
    # from a hidden prompt, or text that already exists in memory such as a secret Graph just minted.
    param(
        [Parameter(Mandatory, ParameterSetName = 'Secure')][securestring]$Password,
        [Parameter(Mandatory, ParameterSetName = 'Plain')][string]$PlainText
    )
    Add-Type -AssemblyName System.Security
    $plain   = if ($PSCmdlet.ParameterSetName -eq 'Secure') { [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)) } else { $PlainText }
    $bytes   = [System.Text.Encoding]::UTF8.GetBytes($plain)
    $entropy = [System.Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1')
    $cipher  = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $entropy, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    return [Convert]::ToBase64String($cipher)
}

function Save-SmtpCredential {
    # Writes the encrypted secret to a file that is locked to SYSTEM and Administrators before any content lands in it
    param([Parameter(Mandatory)][string]$CipherText)
    $path = Join-Path $InstallDir $CredentialFileName
    if (-not (Test-Path $path)) { New-Item -Path $path -ItemType File -Force | Out-Null }
    & icacls.exe "$path" /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "icacls could not restrict '$path' (exit $LASTEXITCODE)." }
    Set-Content -Path $path -Value $CipherText -Encoding ASCII -Force
    Write-Log -Level CREATED -Message "Stored encrypted mail secret at '$path' (readable by SYSTEM and Administrators only)."
}

function Remove-SmtpCredential {
    # Deletes any stored credential when the chosen method does not use one. Says so only when the file is really gone.
    $path = Join-Path $InstallDir $CredentialFileName
    if (-not (Test-Path $path)) { return }
    try {
        Remove-Item -Path $path -Force -ErrorAction Stop
        Write-Log -Level INFORMATIONAL -Message "Removed previous SMTP credential file because the selected method does not use one."
    }
    catch {
        Write-Log -Level WARNING -Message "Could not remove the previous SMTP credential file '$path'. Delete it by hand. $($_.Exception.Message)"
    }
}

function Protect-InstallFolder {
    # Locks a folder to SYSTEM and Administrators with read-only access for Users, and takes ownership away from
    # anyone else. Without this a standard user who pre-creates the folder owns the script that runs as SYSTEM.
    # -OwnerOnly is for the shared parent folder: ownership and stray grants are fixed, inheritance is left alone.
    param([Parameter(Mandatory)][string]$Path, [switch]$OwnerOnly)
    $trusted = @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators', 'NT SERVICE\TrustedInstaller')
    $owner = (Get-Acl -Path $Path).Owner
    if ($owner -notin $trusted) {
        & icacls.exe "$Path" /setowner "*S-1-5-32-544" /T /C | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "icacls could not take ownership of '$Path' (exit $LASTEXITCODE)." }
        if ($OwnerOnly) { & icacls.exe "$Path" /remove:g "$owner" | Out-Null }
        Write-Log -Level WARNING -Message "'$Path' was owned by '$owner'. Ownership moved to Administrators."
    }
    if ($OwnerOnly) { return }
    & icacls.exe "$Path" /inheritance:r /grant:r "*S-1-5-18:(OI)(CI)(F)" "*S-1-5-32-544:(OI)(CI)(F)" "*S-1-5-32-545:(OI)(CI)(RX)" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "icacls could not set permissions on '$Path' (exit $LASTEXITCODE)." }
    # Files already inside may carry explicit grants from whoever created them. Make them inherit from the folder.
    # Credential files and other monitors' folders are left alone: resetting into them would drop the
    # SYSTEM and Administrators only lock that Save-SmtpCredential puts on every stored secret.
    foreach ($item in (Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue)) {
        if ($item.PSIsContainer) { continue }
        if ($item.Name -eq $CredentialFileName) { continue }
        & icacls.exe "$($item.FullName)" /reset /C | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "icacls could not reset permissions on '$($item.FullName)' (exit $LASTEXITCODE)." }
    }
}

function Protect-StoredSecret {
    # Re-applies the SYSTEM and Administrators only lock to every stored secret under the root. Folder hardening
    # runs on every install, so without this a second monitor would leave the first one's secret readable by Users.
    param([Parameter(Mandatory)][string]$Root)
    foreach ($file in @(Get-ChildItem -Path $Root -Filter "*$CredentialFileName" -Recurse -File -Force -ErrorAction SilentlyContinue)) {
        & icacls.exe "$($file.FullName)" /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Log -Level WARNING -Message "Could not re-lock '$($file.FullName)' (icacls exit $LASTEXITCODE). Check its permissions by hand." }
    }
}

function Get-StoredSecret {
    # Decrypts the secret a previous install stored on this machine. Returns $null when there is none or it cannot be read.
    $path = Join-Path $InstallDir $CredentialFileName
    if (-not (Test-Path $path)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $cipher  = [Convert]::FromBase64String((Get-Content -Path $path -Raw -ErrorAction Stop).Trim())
        $entropy = [System.Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1')
        $plain   = [System.Text.Encoding]::UTF8.GetString([System.Security.Cryptography.ProtectedData]::Unprotect($cipher, $entropy, [System.Security.Cryptography.DataProtectionScope]::LocalMachine))
        if ([string]::IsNullOrEmpty($plain)) { return $null }
        return $plain
    }
    catch { return $null }
}

function Test-GuidLike {
    param([string]$Value)
    return [bool]($Value -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')
}

function Test-DateInput {
    # Accepts yyyy-MM-dd or blank. Returns the normalized string, "" for blank, or $null when invalid.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Value.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
    return $null
}

function ConvertTo-GraphMailBody {
    # Builds the JSON body for POST /users/{sender}/sendMail
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string[]]$To)
    $rcpts = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    return (@{ message = @{ subject = $Subject; body = @{ contentType = 'HTML'; content = $Body }; toRecipients = $rcpts }; saveToSentItems = $false } | ConvertTo-Json -Depth 6)
}

function New-AlertBody {
    # Builds a compact HTML body: a heading, a two-column table of details, and a footer naming the sending server
    param([Parameter(Mandatory)][string]$Heading, [Parameter(Mandatory)][System.Collections.IDictionary]$Details, [string]$Footer = "Sent by the curl monitor on $env:COMPUTERNAME.")
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
    $rows = foreach ($k in $Details.Keys) { "<tr><td style='padding:3px 16px 3px 0;color:#555;white-space:nowrap;vertical-align:top'>$(& $enc $k)</td><td style='padding:3px 0'>$(& $enc $Details[$k])</td></tr>" }
    return "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'><p style='font-size:16px;font-weight:600;margin:0 0 10px'>$(& $enc $Heading)</p><table style='border-collapse:collapse'>$($rows -join '')</table><p style='margin:14px 0 0;color:#777;font-size:12px'>$(& $enc $Footer)</p></body></html>"
}

function Get-AlertLabel {
    # "{Monitor} at {Site} ({HOST})" for subjects, dropping any part that is not set
    param([string]$MonitorName, [string]$SiteName, [string]$HostName)
    $label = if ($MonitorName -and $SiteName) { "$MonitorName at $SiteName" } elseif ($MonitorName) { $MonitorName } else { $SiteName }
    if ($HostName) { return "$label ($HostName)" }
    return $label
}

function Get-RestErrorDetail {
    # Pulls the service's own error code and message out of a web exception, so failures name their real cause
    param($ErrorRecord)
    $detail = ""
    try { $detail = $ErrorRecord.ErrorDetails.Message } catch { }
    if (-not $detail) { try { $stream = $ErrorRecord.Exception.Response.GetResponseStream(); $detail = (New-Object System.IO.StreamReader($stream)).ReadToEnd() } catch { } }
    if (-not $detail) { return "" }
    $detail = ($detail -replace '\s*\r?\n\s*', ' ').Trim()   # PowerShell 7 pretty-prints JSON bodies; keep one log line
    if ($detail.Length -gt 1000) { $detail = $detail.Substring(0, 1000) }
    try { $j = $detail | ConvertFrom-Json; if ($j.error.code -or $j.error.message) { return ("[{0}] {1}" -f $j.error.code, $j.error.message) } } catch { }
    return $detail
}

function Get-InstallerGraphToken {
    # Returns an app-only Graph token for the installer's own sends, reusing the last one until 5 minutes before it
    # expires. Microsoft's sign-in service replicates a new secret asynchronously and app-only requests have no replica
    # affinity, so for minutes after creation some replicas still answer AADSTS7000215 (secret unknown) or AADSTS700016
    # (app unknown) while others issue tokens. For a secret minted in this run those answers, and transient failures
    # (429, 5xx, timeouts), are retried until 20 minutes after minting. A pasted secret gets 60 seconds, because one
    # created moments ago elsewhere replicates the same way. MaxWaitSeconds caps the wait for callers that must not
    # block. Anything else, or a minted secret still rejected once its window has passed, is thrown to the caller.
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string]$Secret, [int]$MaxWaitSeconds = 0)
    $cached = $script:InstallerToken
    if ($cached -and $cached.TenantId -eq $TenantId -and $cached.ClientId -eq $ClientId -and [string]::Equals($cached.Secret, $Secret, [StringComparison]::Ordinal) -and (Get-Date).ToUniversalTime() -lt $cached.ExpiresUtc.AddMinutes(-5)) { return $cached.AccessToken }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $body = @{ client_id = $ClientId; client_secret = $Secret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' }
    $deadlineUtc = $null; $noteUtc = $null
    while ($true) {
        try {
            $resp = Invoke-RestMethod -Method Post -Uri "$script:LoginBase/$TenantId/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30 -ErrorAction Stop
            $script:InstallerToken = @{ TenantId = $TenantId; ClientId = $ClientId; Secret = $Secret; AccessToken = $resp.access_token; ExpiresUtc = (Get-Date).ToUniversalTime().AddSeconds([int]$resp.expires_in) }
            return $resp.access_token
        }
        catch {
            $detail = "$(Get-RestErrorDetail $_)"
            $status = 0; try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            $nowUtc = (Get-Date).ToUniversalTime()
            $minted = ($ClientId -eq $script:FreshSecretClientId)
            $fresh = $minted -and ($nowUtc -lt $script:FreshSecretUntilUtc)
            $reason = if ($detail -match 'AADSTS700016') { 'AADSTS700016' } elseif ($detail -match 'AADSTS7000215') { 'AADSTS7000215' } elseif ($status -in @(429, 500, 502, 503, 504)) { "HTTP $status" } elseif ("$($_.Exception.Message)" -match 'timed out|timeout') { 'timeout' } else { '' }
            $lag = $reason -match '^AADSTS'
            if (-not $reason) { throw }
            if (-not $fresh -and (-not $lag -or $minted)) { throw }   # transient outside the window, or a minted secret whose window is spent
            if ($null -eq $deadlineUtc) {
                if ($fresh) {
                    $deadlineUtc = $script:FreshSecretUntilUtc
                    Write-Log -Level INFORMATIONAL -Message "The sign-in service has not accepted the new secret yet ($reason). Usually a replication delay that clears within minutes. Retrying every $script:FreshSecretRetrySeconds s until $($deadlineUtc.ToLocalTime().ToString('HH:mm'))."
                }
                else {
                    $deadlineUtc = $nowUtc.AddSeconds($script:PastedSecretRetrySeconds)
                    Write-Log -Level INFORMATIONAL -Message "The sign-in service rejected the client secret ($reason). A secret created in the last few minutes can still be replicating, so this is retried every $script:FreshSecretRetrySeconds s for $script:PastedSecretRetrySeconds s."
                }
                if ($MaxWaitSeconds -gt 0 -and $deadlineUtc -gt $nowUtc.AddSeconds($MaxWaitSeconds)) { $deadlineUtc = $nowUtc.AddSeconds($MaxWaitSeconds) }
                $noteUtc = $nowUtc.AddSeconds($script:FreshSecretNoteSeconds)
            }
            if ($nowUtc -ge $deadlineUtc) { throw }
            if ($nowUtc -ge $noteUtc) {
                Write-Log -Level INFORMATIONAL -Message "Still waiting for the sign-in service to accept the secret ($reason). Giving up at $($deadlineUtc.ToLocalTime().ToString('HH:mm'))."
                $noteUtc = $nowUtc.AddSeconds($script:FreshSecretNoteSeconds)
            }
            Start-Sleep -Seconds $script:FreshSecretRetrySeconds
        }
    }
}

function Send-GraphMail {
    # Sends via Microsoft Graph with client credentials. Returns $true on success. Never throws. The failure class is
    # left in $script:LastGraphSendError (InvalidSecret, AccessDenied, Other); the caller that owns the prompts says
    # what to do about it.
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string]$Secret, [Parameter(Mandatory)][string]$SenderAddress, [Parameter(Mandatory)][string[]]$To, [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body, [int]$MaxWaitSeconds = 0)
    $script:LastGraphSendError = ''
    try {
        $token = Get-InstallerGraphToken -TenantId $TenantId -ClientId $ClientId -Secret $Secret -MaxWaitSeconds $MaxWaitSeconds
        Invoke-RestMethod -Method Post -Uri "$script:GraphBase/v1.0/users/$SenderAddress/sendMail" -Headers @{ Authorization = "Bearer $token" } -Body (ConvertTo-GraphMailBody -Subject $Subject -Body $Body -To $To) -ContentType 'application/json; charset=utf-8' -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        $script:InstallerToken = $null
        $detail = "$(Get-RestErrorDetail $_)"
        Write-Log -Level FAILED -Message "Graph send failed as $SenderAddress. $($_.Exception.Message) $detail"
        if ($detail -match 'AADSTS7000215|AADSTS700016') {
            $script:LastGraphSendError = 'InvalidSecret'
            if ($ClientId -eq $script:FreshSecretClientId) {
                $age = [int][math]::Round(((Get-Date).ToUniversalTime() - $script:FreshSecretCreatedUtc).TotalMinutes)
                Write-Log -Level WARNING -Message "The secret created in this run is still rejected $age minutes after it was created, which is past any replication delay."
            }
            else { Write-Log -Level WARNING -Message "The sign-in service does not accept this client secret for app $ClientId in tenant $TenantId." }
        }
        elseif ($detail -match 'ErrorAccessDenied') {
            $script:LastGraphSendError = 'AccessDenied'
            if (-not $script:QuietAccessDeniedHint) {
                Write-Log -Level INFORMATIONAL -Message "Access denied right after tenant setup usually means the Exchange sending grant has not reached Graph yet. Microsoft documents 30 minutes to 2 hours for permission changes. Wait at least 30 minutes without retrying (each attempt keeps the old authorization cached), then retry once."
            }
        }
        else { $script:LastGraphSendError = 'Other' }
        return $false
    }
}

function Test-RequiredModule {
    # Ensures a module is present, installing it silently for the current user when missing. Returns $true when available.
    param([Parameter(Mandatory)][string]$Name)
    if (Get-Module -ListAvailable -Name $Name) { Write-Log -Level "SANITY CHECK" -Message "Module $Name present."; return $true }
    Write-Log -Level INFORMATIONAL -Message "Installing $Name for the current user (needed only for the one-time tenant setup, removed afterwards)."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null }
        if ((Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy -ne 'Trusted') { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue }
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log -Level ADDED -Message "Installed $Name."
        return $true
    }
    catch { Write-Log -Level FAILED -Message "Could not install $Name. $($_.Exception.Message)"; return $false }
}

function Remove-TenantSetupModules {
    # Removes the setup-only module from this server once the tenant step is done
    try {
        Remove-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
        Uninstall-Module -Name ExchangeOnlineManagement -AllVersions -Force -ErrorAction Stop
        Write-Log -Level INFORMATIONAL -Message "Removed module ExchangeOnlineManagement from this server."
    }
    catch { Write-Log -Level WARNING -Message "Could not remove ExchangeOnlineManagement. $($_.Exception.Message)" }
}

function ConvertTo-Base64Url {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-JwtClaims {
    # Decodes the payload of a JWT without validation, for reading tid and upn
    param([Parameter(Mandatory)][string]$Token)
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    while ($payload.Length % 4) { $payload += '=' }
    return ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json)
}

function Get-FreeTcpPort {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $l.Start(); $port = ($l.LocalEndpoint).Port; $l.Stop()
    return $port
}

function Get-BrowserToken {
    # Interactive sign-in with PKCE in the default browser. The response arrives on a loopback listener. Returns @{ Token; RefreshToken; TenantId; Upn } or $null.
    param([Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string]$Scope, [string]$Prompt = 'select_account', [string]$Title = 'Sign in')
    $listener = $null
    try {
        $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        $vbytes = New-Object byte[] 32; $rng.GetBytes($vbytes)
        $verifier  = ConvertTo-Base64Url $vbytes
        $challenge = ConvertTo-Base64Url ([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)))
        $state = [guid]::NewGuid().ToString('N')
        for ($bindTry = 0; $bindTry -lt 5 -and $null -eq $listener; $bindTry++) {
            $port = Get-FreeTcpPort
            $candidate = "http://localhost:$port/"
            $hl = New-Object System.Net.HttpListener
            $hl.Prefixes.Add($candidate)
            try { $hl.Start(); $listener = $hl; $redirect = $candidate }
            catch { try { $hl.Close() } catch { }; Start-Sleep -Milliseconds 150 }
        }
        if ($null -eq $listener) { throw "Could not bind a loopback listener for the sign-in redirect." }

        $authUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?client_id=$ClientId&response_type=code&redirect_uri=$([uri]::EscapeDataString($redirect))&response_mode=query&scope=$([uri]::EscapeDataString($Scope))&state=$state&code_challenge=$challenge&code_challenge_method=S256"
        if ($Prompt) { $authUrl += "&prompt=$Prompt" }

        Start-Process $authUrl -ErrorAction Stop
        Write-Log -Level PROMPT -Message "$Title in your browser. Waiting up to 5 minutes."

        $deadline = (Get-Date).AddMinutes(5)
        while ($true) {
            $task = $listener.GetContextAsync()
            while (-not $task.Wait(500)) { if ((Get-Date) -gt $deadline) { throw "Sign-in timed out." } }
            $ctx = $task.Result
            $code = $ctx.Request.QueryString['code']; $err = $ctx.Request.QueryString['error']
            if (-not $code -and -not $err) {
                # Not the redirect (favicon or pre-connect probe). Answer 204 and keep waiting.
                try { $ctx.Response.StatusCode = 204; $ctx.Response.OutputStream.Close() } catch { }
                continue
            }
            $gotState = $ctx.Request.QueryString['state']; $errDesc = $ctx.Request.QueryString['error_description']
            break
        }
        if ($code -and $gotState -ne $state) { $code = $null; $err = 'state_mismatch'; $errDesc = 'Sign-in response did not match this session.' }
        $html = if ($code) { "<html><body style='font-family:Segoe UI;padding:40px'><h2>Signed in.</h2><p>You can close this tab and return to the installer.</p></body></html>" } else { "<html><body style='font-family:Segoe UI;padding:40px'><h2>Sign-in failed.</h2><p>$errDesc</p></body></html>" }
        $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
        $ctx.Response.ContentType = 'text/html'; $ctx.Response.ContentLength64 = $buf.Length
        $ctx.Response.OutputStream.Write($buf, 0, $buf.Length); $ctx.Response.OutputStream.Close()
        if (-not $code) { throw "Microsoft returned: $err $errDesc" }

        $tok = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -Body @{ grant_type = 'authorization_code'; client_id = $ClientId; code = $code; redirect_uri = $redirect; code_verifier = $verifier; scope = $Scope } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $claims = ConvertFrom-JwtClaims $tok.access_token
        Write-Log -Level FOUND -Message "Signed in to tenant $($claims.tid) as $($claims.upn)."
        return @{ Token = $tok.access_token; RefreshToken = $tok.refresh_token; TenantId = $claims.tid; Upn = $claims.upn }
    }
    catch {
        Write-Log -Level WARNING -Message "Sign-in did not complete ($($_.Exception.Message))."
        return $null
    }
    finally { if ($listener) { try { $listener.Stop(); $listener.Close() } catch { } } }
}

function Get-DeviceCodeToken {
    # Fallback sign-in for servers with no browser. Returns @{ Token; RefreshToken; TenantId; Upn } or $null.
    param([Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string]$Scope)
    try {
        $dc = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode' -Body @{ client_id = $ClientId; scope = $Scope } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    }
    catch { Write-Log -Level FAILED -Message "Could not start the sign-in. $($_.Exception.Message)"; return $null }
    $signInUrl = "$($dc.verification_uri)?otc=$($dc.user_code)"
    try { Set-Clipboard -Value $dc.user_code } catch { }
    Write-Host ""
    Write-Host $dc.message
    Write-Log -Level PROMPT -Message "On any device open $signInUrl and sign in as a Global Administrator. The code is on the clipboard. Waiting up to $([int]($dc.expires_in / 60)) minutes."
    Write-Host ""
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [int]$dc.interval
    if ($interval -lt 1) { $interval = 5 }
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $ClientId; device_code = $dc.device_code } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $claims = ConvertFrom-JwtClaims $tok.access_token
            Write-Log -Level FOUND -Message "Signed in to tenant $($claims.tid) as $($claims.upn)."
            return @{ Token = $tok.access_token; RefreshToken = $tok.refresh_token; TenantId = $claims.tid; Upn = $claims.upn }
        }
        catch {
            $err = ""
            try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch { }
            if ($err -eq 'slow_down') { $interval += 5; continue }
            if ($err -eq 'authorization_pending') { continue }
            Write-Log -Level FAILED -Message "Sign-in failed. $($_.Exception.Message)"
            return $null
        }
    }
    Write-Log -Level FAILED -Message "Sign-in timed out."
    return $null
}

function Get-DelegatedGraphToken {
    # Admin sign-in for the one-time tenant setup: browser first, device code fallback. No modules.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $clientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'   # Azure CLI first-party public client. Its Graph token carries Directory.AccessAsUser.All and its refresh token also yields Exchange Online tokens.
    $scope = 'https://graph.microsoft.com/.default offline_access'   # .default only. Naming individual Graph scopes on a first-party client returns AADSTS65002.
    $auth = Get-BrowserToken -ClientId $clientId -Scope $scope -Title 'Sign in as a Global Administrator'
    if (-not $auth) { Write-Log -Level WARNING -Message "Falling back to a device code."; $auth = Get-DeviceCodeToken -ClientId $clientId -Scope $scope }
    return $auth
}

function Get-ExchangeToken {
    # Redeems the sign-in's refresh token for an Exchange Online token with the same client. No second sign-in. Returns the token or $null.
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$RefreshToken)
    try {
        $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{ grant_type = 'refresh_token'; client_id = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'; refresh_token = $RefreshToken; scope = 'https://outlook.office365.com/.default offline_access' } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        return $r.access_token
    }
    catch { Write-Log -Level WARNING -Message "Could not obtain an Exchange Online token from the sign-in. $($_.Exception.Message) $(Get-RestErrorDetail $_)"; return $null }
}

function Invoke-ExoCommand {
    # Runs one Exchange Online cmdlet through the Exchange admin REST endpoint (what the EXO module uses internally). Returns the value array.
    # Throttling and server errors are retried a few times so a blip never changes which setup path runs.
    param([Parameter(Mandatory)][string]$Token, [Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$Upn, [Parameter(Mandatory)][string]$Cmdlet, [hashtable]$Parameters = @{})
    $headers = @{ Authorization = "Bearer $Token"; 'X-AnchorMailbox' = "UPN:$Upn"; 'X-ResponseFormat' = 'json'; 'Prefer' = 'odata.maxpagesize=1000'; 'X-ClientApplication' = 'CurlMonitorInstaller' }
    $body = @{ CmdletInput = @{ CmdletName = $Cmdlet; Parameters = $Parameters } } | ConvertTo-Json -Depth 6 -Compress
    $r = $null
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $r = Invoke-RestMethod -Method Post -Uri "https://outlook.office365.com/adminapi/beta/$TenantId/InvokeCommand" -Headers $headers -Body $body -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
            break
        }
        catch {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            if ($attempt -lt 4 -and $status -in @(429, 500, 502, 503, 504)) {
                Write-Log -Level WARNING -Message "$Cmdlet returned HTTP $status. Retrying in $(10 * $attempt) s (attempt $attempt of 4)."
                Start-Sleep -Seconds (10 * $attempt)
                continue
            }
            # Surface the server's own error text, not just the HTTP status. Keep the raw body; the parsed summary alone loses the detail that names the failing input.
            $detail = ""
            try { $detail = $_.ErrorDetails.Message } catch { }
            if (-not $detail) { try { $stream = $_.Exception.Response.GetResponseStream(); $detail = (New-Object System.IO.StreamReader($stream)).ReadToEnd() } catch { } }
            if ($detail) {
                if ($detail.Length -gt 1500) { $detail = $detail.Substring(0, 1500) }
                $parsed = ""
                try { $j = $detail | ConvertFrom-Json; if ($j.error.message) { $parsed = $j.error.message } } catch { }
                if ($parsed -and $parsed -ne $detail) { $detail = "$parsed | Full response: $detail" }
            }
            throw "$Cmdlet failed: $($_.Exception.Message) $detail"
        }
    }
    if ($null -ne $r.value) { return @($r.value) }
    return @()
}

function Invoke-ExoCmdlet {
    # Dispatches an Exchange cmdlet to REST (default) or to the loaded module (fallback). Throws on error.
    # With -NullOnNotFound only a genuine "object not found" becomes $null; anything else still throws, so a
    # throttled or failed lookup can never be mistaken for "does not exist yet".
    param([Parameter(Mandatory)][string]$Name, [hashtable]$Parameters = @{}, [switch]$NullOnNotFound)
    try {
        if ($script:ExoMode -eq 'module') { return @(& $Name @Parameters -ErrorAction Stop) }
        return (Invoke-ExoCommand -Token $script:ExoToken -TenantId $script:ExoTenantId -Upn $script:ExoUpn -Cmdlet $Name -Parameters $Parameters)
    }
    catch {
        if ($NullOnNotFound -and ($_.Exception.Message -match "couldn.t be found|could not be found|isn.t found|not found|doesn.t exist|does not exist|ObjectNotFound|\(404\)")) { return $null }
        throw
    }
}

function Invoke-GraphRequest {
    # Minimal Graph REST wrapper for the tenant setup
    param([Parameter(Mandatory)][string]$Token, [Parameter(Mandatory)][string]$Method, [Parameter(Mandatory)][string]$Path, [object]$Body)
    $p = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0/$Path"; Headers = @{ Authorization = "Bearer $Token" }; ContentType = 'application/json; charset=utf-8'; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 8) }
    return Invoke-RestMethod @p
}

function New-TenantMailApp {
    # One-time tenant setup. Returns a hashtable with TenantId, ClientId, Secret, Expires, Sender, or $null on failure.
    param([Parameter(Mandatory)][string]$SenderAddress, [Parameter(Mandatory)][string]$SiteName)

    $auth = Get-DelegatedGraphToken
    if (-not $auth) { return $null }
    $t = $auth.Token
    $tenantId = $auth.TenantId

    try {
        $graphSp = (Invoke-GraphRequest -Token $t -Method GET -Path "servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value | Select-Object -First 1
        $mailSendRole = $graphSp.appRoles | Where-Object { $_.value -eq 'Mail.Send' -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $mailSendRole) { Write-Log -Level FAILED -Message "Mail.Send application role not found on the Graph service principal."; return $null }

        $nameLiteral = $GraphAppDisplayName.Replace("'", "''")
        $app = (Invoke-GraphRequest -Token $t -Method GET -Path "applications?`$filter=displayName eq '$nameLiteral'").value | Select-Object -First 1
        if ($app) { Write-Log -Level FOUND -Message "App registration '$GraphAppDisplayName' exists (AppId $($app.appId)). Reusing it." }
        else {
            $body = @{ displayName = $GraphAppDisplayName; signInAudience = 'AzureADMyOrg'; notes = "Curl monitor alert sender. Restricted to sending only as $SenderAddress in Exchange Online."; requiredResourceAccess = @(@{ resourceAppId = $graphSp.appId; resourceAccess = @(@{ id = $mailSendRole.id; type = 'Role' }) }) }
            $app = Invoke-GraphRequest -Token $t -Method POST -Path 'applications' -Body $body
            Write-Log -Level CREATED -Message "Created app registration '$GraphAppDisplayName' (AppId $($app.appId))."
        }

        $sp = (Invoke-GraphRequest -Token $t -Method GET -Path "servicePrincipals?`$filter=appId eq '$($app.appId)'").value | Select-Object -First 1
        if (-not $sp) {
            # A brand-new app registration can take a few seconds to be visible to the service principal endpoint
            for ($i = 0; $i -lt 6 -and -not $sp; $i++) {
                try { $sp = Invoke-GraphRequest -Token $t -Method POST -Path 'servicePrincipals' -Body @{ appId = $app.appId } }
                catch { if ($i -eq 5) { throw }; Start-Sleep -Seconds 5 }
            }
            Write-Log -Level CREATED -Message "Created service principal."
        }

        # Consent and the client secret are handled after the Exchange step: the sending grant depends on which
        # scoping mechanism applies, and minting the secret last keeps failed runs from piling up secrets on the app.
    }
    catch { Write-Log -Level FAILED -Message "Graph configuration failed. $($_.Exception.Message) $(Get-RestErrorDetail $_)"; return $null }

    # Exchange Online: shared mailbox, policy group, application access policy. REST first, module only as fallback.
    $script:ExoMode = 'rest'; $script:ExoTenantId = $tenantId; $script:ExoUpn = $auth.Upn
    $script:ExoToken = Get-ExchangeToken -TenantId $tenantId -RefreshToken $auth.RefreshToken
    $restOk = $false
    if ($script:ExoToken) {
        try { Invoke-ExoCmdlet -Name 'Get-OrganizationConfig' | Out-Null; $restOk = $true; Write-Log -Level FOUND -Message "Connected to Exchange Online with the same sign-in (no module needed)." }
        catch { Write-Log -Level WARNING -Message "Exchange REST call failed ($($_.Exception.Message)). Falling back to the Exchange Online module." }
    }
    else { Write-Log -Level WARNING -Message "Falling back to the Exchange Online module." }
    if (-not $restOk) {
        if (-not (Test-RequiredModule -Name 'ExchangeOnlineManagement')) { Write-Log -Level FAILED -Message "Exchange Online fallback needs the ExchangeOnlineManagement module."; return $null }
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Write-Log -Level PROMPT -Message "Sign in to Exchange Online as an Exchange Administrator in the browser window."
            Connect-ExchangeOnline -UserPrincipalName $auth.Upn -ShowBanner:$false -ErrorAction Stop
            $script:ExoMode = 'module'
        }
        catch { Write-Log -Level FAILED -Message "Exchange Online sign-in failed. $($_.Exception.Message)"; return $null }
    }

    try {
        if (Invoke-ExoCmdlet -Name 'Get-Mailbox' -Parameters @{ Identity = $SenderAddress } -NullOnNotFound) { Write-Log -Level FOUND -Message "Sender mailbox $SenderAddress exists." }
        else {
            $alias = ($SenderAddress -split '@')[0]
            Invoke-ExoCmdlet -Name 'New-Mailbox' -Parameters @{ Shared = $true; Name = $alias; DisplayName = $GraphAppDisplayName; PrimarySmtpAddress = $SenderAddress } | Out-Null
            Write-Log -Level CREATED -Message "Created shared mailbox $SenderAddress (no license required)."
        }
        # Fresh tenants start dehydrated and reject organization-level writes until customization is enabled
        $org = @(Invoke-ExoCmdlet -Name 'Get-OrganizationConfig') | Select-Object -First 1
        if ($org -and "$($org.IsDehydrated)" -eq 'True') {
            try { Invoke-ExoCmdlet -Name 'Enable-OrganizationCustomization' | Out-Null; Write-Log -Level ADDED -Message "Enabled organization customization (one-time tenant step). This can take several minutes to apply." }
            catch { Write-Log -Level WARNING -Message "Could not enable organization customization. $($_.Exception.Message)" }
        }

        # Limit the app to sending only as the shared mailbox. RBAC for Applications is the current mechanism;
        # the legacy application access policy is the automatic fallback (Microsoft marks it legacy, replaced by RBAC).
        $scopedBy = ''
        try {
            if (-not (Invoke-ExoCmdlet -Name 'Get-ServicePrincipal' -Parameters @{ Identity = $sp.id } -NullOnNotFound)) {
                Invoke-ExoCmdlet -Name 'New-ServicePrincipal' -Parameters @{ AppId = $app.appId; ObjectId = $sp.id; DisplayName = $GraphAppDisplayName } | Out-Null
                Write-Log -Level CREATED -Message "Registered the app's service principal in Exchange Online."
            }
            $scopeName = "$GraphAppDisplayName Send Scope"
            $senderLiteral = $SenderAddress.Replace("'", "''")
            $scope = @(Invoke-ExoCmdlet -Name 'Get-ManagementScope' -Parameters @{ Identity = $scopeName } -NullOnNotFound) | Select-Object -First 1
            if (-not $scope) {
                Invoke-ExoCmdlet -Name 'New-ManagementScope' -Parameters @{ Name = $scopeName; RecipientRestrictionFilter = "PrimarySmtpAddress -eq '$senderLiteral'" } | Out-Null
                Write-Log -Level CREATED -Message "Created management scope '$scopeName' limited to $SenderAddress."
            }
            elseif ("$($scope.RecipientFilter)" -notmatch [regex]::Escape($SenderAddress)) {
                Invoke-ExoCmdlet -Name 'Set-ManagementScope' -Parameters @{ Identity = $scopeName; RecipientRestrictionFilter = "PrimarySmtpAddress -eq '$senderLiteral'" } | Out-Null
                Write-Log -Level ADDED -Message "Updated management scope '$scopeName' to target $SenderAddress."
            }
            $assignmentName = "$GraphAppDisplayName Mail Send"
            $assignmentExisted = [bool](Invoke-ExoCmdlet -Name 'Get-ManagementRoleAssignment' -Parameters @{ Identity = $assignmentName } -NullOnNotFound)
            if ($assignmentExisted) { Write-Log -Level FOUND -Message "Role assignment '$assignmentName' exists." }
            else {
                Invoke-ExoCmdlet -Name 'New-ManagementRoleAssignment' -Parameters @{ Name = $assignmentName; App = $sp.id; Role = 'Application Mail.Send'; CustomResourceScope = $scopeName } | Out-Null
                Write-Log -Level CREATED -Message "Granted 'Application Mail.Send' to the app, limited to $SenderAddress."
            }
            $chk = @(Invoke-ExoCmdlet -Name 'Test-ServicePrincipalAuthorization' -Parameters @{ Identity = $sp.id; Resource = $SenderAddress } -NullOnNotFound)
            $inScope = [bool]($chk | Where-Object { $_.RoleName -eq 'Application Mail.Send' -and "$($_.InScope)" -eq 'True' })
            if (-not $inScope -and $assignmentExisted) {
                # The authorization test reads the configuration directly, so an older assignment that does not cover this
                # service principal (app deleted and re-created) is replaced instead of trusted by name.
                Write-Log -Level WARNING -Message "Role assignment '$assignmentName' does not cover this app's service principal. Re-creating it."
                Invoke-ExoCmdlet -Name 'Remove-ManagementRoleAssignment' -Parameters @{ Identity = $assignmentName; Confirm = $false } | Out-Null
                Invoke-ExoCmdlet -Name 'New-ManagementRoleAssignment' -Parameters @{ Name = $assignmentName; App = $sp.id; Role = 'Application Mail.Send'; CustomResourceScope = $scopeName } | Out-Null
                Write-Log -Level CREATED -Message "Re-created role assignment '$assignmentName' for this app."
                $chk = @(Invoke-ExoCmdlet -Name 'Test-ServicePrincipalAuthorization' -Parameters @{ Identity = $sp.id; Resource = $SenderAddress } -NullOnNotFound)
                $inScope = [bool]($chk | Where-Object { $_.RoleName -eq 'Application Mail.Send' -and "$($_.InScope)" -eq 'True' })
            }
            if ($inScope) { Write-Log -Level "SANITY CHECK" -Message "Authorization test: $SenderAddress = InScope." }
            else { Write-Log -Level WARNING -Message "Authorization test not confirmed yet. Permission changes can take up to 30 minutes to apply." }
            $scopedBy = 'rbac'
        }
        catch { Write-Log -Level WARNING -Message "RBAC scoping failed ($($_.Exception.Message)). Falling back to the legacy application access policy." }

        if ($scopedBy -eq 'rbac') {
            # The Exchange role grant is the sending authority. Tenant-wide Mail.Send consent is a union with it and
            # would defeat the scope, so remove the consent if an earlier run granted it.
            $assigned = (Invoke-GraphRequest -Token $t -Method GET -Path "servicePrincipals/$($sp.id)/appRoleAssignments").value | Where-Object { $_.appRoleId -eq $mailSendRole.id -and $_.resourceId -eq $graphSp.id }
            $removedConsent = $false
            foreach ($a in @($assigned)) {
                try { Invoke-GraphRequest -Token $t -Method DELETE -Path "servicePrincipals/$($sp.id)/appRoleAssignments/$($a.id)" | Out-Null; $removedConsent = $true; Write-Log -Level INFORMATIONAL -Message "Removed tenant-wide Mail.Send consent. Sending rights now come only from the Exchange scope." }
                catch { Write-Log -Level WARNING -Message "Could not remove the tenant-wide Mail.Send consent. Until it is removed the app can send as any mailbox. $($_.Exception.Message)" }
            }
            if ($removedConsent) { Write-Log -Level WARNING -Message "The Exchange grant can take 30 minutes to 2 hours to reach live sends. If the test email is denied, wait at least 30 minutes without retrying, then try once." }
        }
        else {
            $assigned = (Invoke-GraphRequest -Token $t -Method GET -Path "servicePrincipals/$($sp.id)/appRoleAssignments").value | Where-Object { $_.appRoleId -eq $mailSendRole.id -and $_.resourceId -eq $graphSp.id }
            if ($assigned) { Write-Log -Level FOUND -Message "Mail.Send admin consent already granted." }
            else {
                $granted = $false
                for ($i = 0; $i -lt 6 -and -not $granted; $i++) {
                    try { Invoke-GraphRequest -Token $t -Method POST -Path "servicePrincipals/$($sp.id)/appRoleAssignments" -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $mailSendRole.id } | Out-Null; $granted = $true }
                    catch { if ($i -eq 5) { throw }; Start-Sleep -Seconds 5 }
                }
                Write-Log -Level ADDED -Message "Granted admin consent for Mail.Send (application)."
            }
            $policyGroup = "$GraphPolicyGroupAlias@" + (($SenderAddress -split '@')[-1])
            $grp = @(Invoke-ExoCmdlet -Name 'Get-DistributionGroup' -Parameters @{ Identity = $policyGroup } -NullOnNotFound) | Select-Object -First 1
            if (-not $grp) {
                Invoke-ExoCmdlet -Name 'New-DistributionGroup' -Parameters @{ Type = 'Security'; Name = $GraphPolicyGroupAlias; DisplayName = "$GraphAppDisplayName Senders"; PrimarySmtpAddress = $policyGroup } | Out-Null
                Write-Log -Level CREATED -Message "Created mail-enabled security group $policyGroup."
                $grp = @(Invoke-ExoCmdlet -Name 'Get-DistributionGroup' -Parameters @{ Identity = $policyGroup } -NullOnNotFound) | Select-Object -First 1
            }
            $members = Invoke-ExoCmdlet -Name 'Get-DistributionGroupMember' -Parameters @{ Identity = $policyGroup } -NullOnNotFound
            if (-not ($members | Where-Object { $_.PrimarySmtpAddress -eq $SenderAddress })) {
                try { Invoke-ExoCmdlet -Name 'Add-DistributionGroupMember' -Parameters @{ Identity = $policyGroup; Member = $SenderAddress } | Out-Null; Write-Log -Level ADDED -Message "Added $SenderAddress to the policy group." }
                catch { Write-Log -Level WARNING -Message "Could not add $SenderAddress to the policy group yet (new objects can take a minute). $($_.Exception.Message)" }
            }
            $policies = Invoke-ExoCmdlet -Name 'Get-ApplicationAccessPolicy' -NullOnNotFound
            if ($policies | Where-Object { $_.AppId -eq $app.appId }) { Write-Log -Level FOUND -Message "Application access policy already exists for this app." }
            else {
                # AppId is declared String[] on the cmdlet, so the wire value must be an array. Later attempts switch to the group's directory id.
                $policyDone = $false
                for ($i = 1; $i -le 8 -and -not $policyDone; $i++) {
                    $scopeId = $policyGroup
                    if ($i -ge 5 -and $grp -and $grp.ExternalDirectoryObjectId) { $scopeId = "$($grp.ExternalDirectoryObjectId)" }
                    try {
                        Invoke-ExoCmdlet -Name 'New-ApplicationAccessPolicy' -Parameters @{ AppId = [string[]]@($app.appId); PolicyScopeGroupId = $scopeId; AccessRight = 'RestrictAccess'; Description = "$GraphAppDisplayName may send only as $SenderAddress" } | Out-Null
                        $policyDone = $true
                        Write-Log -Level CREATED -Message "Applied application access policy: app restricted to $SenderAddress."
                    }
                    catch {
                        if ($i -eq 8) { throw }
                        Write-Log -Level WARNING -Message "Policy creation attempt $i failed ($($_.Exception.Message)). Waiting 15 s before retrying."
                        Start-Sleep -Seconds 15
                    }
                }
            }
            $chk = Invoke-ExoCmdlet -Name 'Test-ApplicationAccessPolicy' -Parameters @{ AppId = $app.appId; Identity = $SenderAddress } -NullOnNotFound
            if ($chk -and ($chk | Select-Object -First 1).AccessCheckResult -eq 'Granted') { Write-Log -Level "SANITY CHECK" -Message "Policy test: $SenderAddress = Granted." }
            else { Write-Log -Level WARNING -Message "Policy test not Granted yet. Policies can take up to 30 minutes to apply." }
        }
    }
    catch { Write-Log -Level FAILED -Message "Exchange configuration failed. $($_.Exception.Message)"; return $null }
    finally { if ($script:ExoMode -eq 'module') { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { } }; $script:ExoToken = $null }

    # Mint the secret only now, so failed runs never leave orphan secrets on the app
    $end = (Get-Date).AddMonths([math]::Min(24, [math]::Max(1, [int]$GraphSecretMonths)))
    try {
        $pw = Invoke-GraphRequest -Token $t -Method POST -Path "applications/$($app.id)/addPassword" -Body @{ passwordCredential = @{ displayName = "$GraphAppDisplayName $(Get-Date -Format 'yyyy-MM-dd') from $SiteName"; endDateTime = $end.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } }
        $secret = $pw.secretText
        $script:FreshSecretClientId = $app.appId
        $script:FreshSecretCreatedUtc = (Get-Date).ToUniversalTime()
        $script:FreshSecretUntilUtc = $script:FreshSecretCreatedUtc.AddMinutes(20)
        $script:InstallerToken = $null
        Write-Log -Level CREATED -Message "Created client secret expiring $($end.ToString('yyyy-MM-dd'))."
    }
    catch { Write-Log -Level FAILED -Message "Could not create the client secret. $($_.Exception.Message) $(Get-RestErrorDetail $_)"; return $null }

    $summary = "$GraphAppDisplayName app registration`nCreated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') on $SiteName`nTenant ID: $tenantId`nClient ID: $($app.appId)`nSender mailbox: $SenderAddress`nSecret expires: $($end.ToString('yyyy-MM-dd'))`nClient secret: shown once on screen, copied to clipboard, not stored here"
    $summaryPath = Join-Path $InstallDir 'AppRegistration.txt'
    try { Set-Content -Path $summaryPath -Value $summary -Encoding UTF8; Write-Log -Level CREATED -Message "Saved summary (without secret) to $summaryPath." } catch { }
    try { Set-Clipboard -Value $secret } catch { }

    Write-Host ""
    Write-Host "Keep these for the other site installs (also in $summaryPath):"
    Write-Host "  Tenant ID      : $tenantId"
    Write-Host "  Client ID      : $($app.appId)"
    [Console]::Out.WriteLine("  Client secret  : $secret   (shown only now, copied to clipboard, kept out of PowerShell transcripts)")
    Write-Host "  Sender mailbox : $SenderAddress"
    Write-Host "  Secret expires : $($end.ToString('yyyy-MM-dd'))"
    Write-Host ""
    if ($script:ExoMode -eq 'module') { Remove-TenantSetupModules }
    Write-Log -Level INFORMATIONAL -Message "Waiting for the sign-in service to accept the new secret (usually under a minute)."
    try {
        $null = Get-InstallerGraphToken -TenantId $tenantId -ClientId $app.appId -Secret $secret
        Write-Log -Level "SANITY CHECK" -Message "The sign-in service issued a token with the new secret. The test email reuses it, and later sends retry on their own while the secret finishes replicating."
    }
    catch { Write-Log -Level WARNING -Message "No token was issued with the new secret. $($_.Exception.Message) $(Get-RestErrorDetail $_)" }
    return @{ TenantId = $tenantId; ClientId = $app.appId; Secret = $secret; Expires = $end.ToString('yyyy-MM-dd'); Sender = $SenderAddress }
}

function Send-MailWithConfig {
    # Sends one message using a mail config hashtable. Returns $true on success. Never throws.
    param([Parameter(Mandatory)][hashtable]$Mail, [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body, [pscredential]$Credential, [string]$GraphSecret, [int]$MaxWaitSeconds = 0)
    if ($Mail.MailMethod -eq 'Graph') {
        if ([string]::IsNullOrEmpty($GraphSecret)) { Write-Log -Level FAILED -Message "No Graph client secret is available for this send."; return $false }
        return (Send-GraphMail -TenantId $Mail.GraphTenantId -ClientId $Mail.GraphClientId -Secret $GraphSecret -SenderAddress $Mail.MailFrom -To $Mail.MailTo -Subject $Subject -Body $Body -MaxWaitSeconds $MaxWaitSeconds)
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $p = @{ SmtpServer = $Mail.SmtpServer; Port = $Mail.SmtpPort; From = $Mail.MailFrom; To = $Mail.MailTo; Subject = $Subject; Body = $Body; BodyAsHtml = $true }
    if ($Mail.SmtpUseSsl) { $p['UseSsl'] = $true }
    if ($Credential) { $p['Credential'] = $Credential }
    try {
        Send-MailMessage @p -ErrorAction Stop -WarningAction SilentlyContinue
        return $true
    }
    catch {
        Write-Log -Level FAILED -Message "Mail send failed via $($Mail.SmtpServer):$($Mail.SmtpPort). $($_.Exception.Message)"
        return $false
    }
}

function Get-MailConfiguration {
    # Interactive mail wizard. Returns a hashtable with SmtpServer, SmtpPort, SmtpUseSsl, MailFrom, MailTo, MailMethod, SmtpAuthUser, CipherText.
    param([object]$Saved, [Parameter(Mandatory)][string]$SiteName, [string]$MonitorName)

    $defMethod = if ($Saved -and $Saved.MailMethod) { $Saved.MailMethod } else { $MailMethod }
    $defFrom   = if ($Saved -and $Saved.MailFrom)   { $Saved.MailFrom } else { $MailFrom }
    $defTo     = if ($Saved -and $Saved.MailTo)     { @($Saved.MailTo) -join ',' } else { $MailTo -join ',' }
    $defTenant = if ($Saved -and $Saved.GraphTenantId) { $Saved.GraphTenantId } else { $GraphTenantId }
    $defClient = if ($Saved -and $Saved.GraphClientId) { $Saved.GraphClientId } else { $GraphClientId }
    $defExpiry = if ($Saved -and $Saved.GraphSecretExpires) { $Saved.GraphSecretExpires } else { $GraphSecretExpires }
    # Host, port, TLS, and user are remembered per method, so trying one method never erases what was saved for another
    $remembered = @{}
    $savedServer = if ($Saved -and $Saved.SmtpServer) { $Saved.SmtpServer } else { $SmtpServer }
    $savedPort   = if ($Saved -and $Saved.SmtpPort)   { [int]$Saved.SmtpPort } else { $SmtpPort }
    $savedSsl    = if ($Saved -and $null -ne $Saved.SmtpUseSsl) { [bool]$Saved.SmtpUseSsl } else { $SmtpUseSsl }
    $savedUser   = if ($Saved -and $Saved.SmtpAuthUser) { $Saved.SmtpAuthUser } else { $SmtpAuthUser }
    if ($defMethod -ne 'Graph') { $remembered[$defMethod] = @{ Server = $savedServer; Port = $savedPort; Ssl = $savedSsl; User = $savedUser } }
    # Set when this run creates the app, so a later Graph pass reuses the new secret and still gets the propagation wait
    $createdClient = ''; $createdSecret = $null

    while ($true) {
        if (-not $NonInteractive) {
            Write-Question -Question "How should alerts be sent?"
            Write-Host "  1  Microsoft 365 Graph. HTTPS only, no port 25. The first site creates the app and the shared mailbox; later sites paste its three values."
            Write-Host "  2  Microsoft 365 direct send. Port 25 to the tenant, recipients inside the tenant, this server's IP in SPF."
            Write-Host "  3  Internal relay. An on-prem Exchange or relay that accepts this server's IP."
            Write-Host "  4  SMTP with a username and password."
        }
        $methodMap = @{ '1' = 'Graph'; '2' = 'DirectSend'; '3' = 'Relay'; '4' = 'Authenticated' }
        $defKey = ($methodMap.GetEnumerator() | Where-Object { $_.Value -eq $defMethod } | Select-Object -First 1).Key
        if (-not $defKey) { $defKey = '1' }
        $method = $methodMap[(Read-Choice -Prompt "Method" -Allowed @('1','2','3','4') -Default $defKey)]

        $from = ""
        while (-not (Test-EmailAddress $from)) {
            $from = Read-Setting -Prompt $(if ($method -eq 'Graph') { "Sender shared mailbox" } else { "Sender address (From)" }) -Default $defFrom
            if (-not (Test-EmailAddress $from)) {
                if ($NonInteractive) { Write-Log -Level FAILED -Message "Sender address '$from' is invalid in non-interactive mode."; return $null }
                Write-Log -Level WARNING -Message "Enter a valid email address."
            }
        }

        $to = @()
        while ($to.Count -eq 0) {
            $to = ConvertTo-RecipientList (Read-Setting -Prompt "Recipient address(es), comma separated" -Default $defTo)
            if ($to.Count -eq 0) {
                if ($NonInteractive) { Write-Log -Level FAILED -Message "No valid recipients in non-interactive mode."; return $null }
                Write-Log -Level WARNING -Message "Enter at least one valid email address."
            }
        }

        $server = ""; $port = 25; $ssl = $false; $user = ""; $cipher = ""; $cred = $null
        $tenant = ""; $client = ""; $expiry = ""; $graphSecret = ""; $restart = $false
        $mem = if ($remembered.ContainsKey($method)) { $remembered[$method] } else { @{} }
        switch ($method) {
            'Graph' {
                if ($NonInteractive) { Write-Log -Level FAILED -Message "Graph cannot be configured non-interactively because the client secret must be entered at the console."; return $null }
                $hasApp = if ($script:RetryTenantSetup) { 'N' } elseif ($defTenant -and $defClient) { 'Y' } else { 'N' }
                $script:RetryTenantSetup = $false
                $created = $null
                if ((Read-Choice -Prompt "Has the '$GraphAppDisplayName' app registration already been created in the tenant? Y = enter its details, N = run the tenant setup now (first site, or to mint a new secret for the existing app)" -Allowed @('Y','N') -Default $hasApp) -eq 'N') {
                    $created = New-TenantMailApp -SenderAddress $from -SiteName $SiteName
                    if (-not $created) {
                        Write-Log -Level WARNING -Message "Tenant setup did not complete."
                        if ((Read-Choice -Prompt "R = back to mail method selection, X = abort install" -Allowed @('R','X') -Default 'R') -eq 'X') { return $null }
                        $defMethod = $method; $defFrom = $from; $defTo = ($to -join ',')
                        $restart = $true
                        break
                    }
                    $tenant = $created.TenantId; $client = $created.ClientId; $expiry = $created.Expires
                    $defTenant = $tenant; $defClient = $client; $defExpiry = $expiry
                    $createdClient = $client; $createdSecret = $created.Secret
                    # Persist immediately: if this run is abandoned before install completes, the next run
                    # prefills Y with these IDs instead of re-creating and minting another secret
                    Save-InstallSettings -Settings @{
                        MonitorName = $MonitorName
                        ContentMarker = $ExpectedContentMarker
                        SiteName = $SiteName
                        MailMethod = 'Graph'
                        SmtpServer = 'graph.microsoft.com'
                        SmtpPort = 443
                        SmtpUseSsl = $true
                        MailFrom = $from
                        MailTo = @($to)
                        SmtpAuthUser = ''
                        GraphTenantId = $tenant
                        GraphClientId = $client
                        GraphSecretExpires = $expiry
                        Url = $Url
                    }
                }
                while (-not (Test-GuidLike $tenant)) {
                    $tenant = Read-Setting -Prompt "Tenant ID" -Default $defTenant
                    if (-not (Test-GuidLike $tenant)) { Write-Log -Level WARNING -Message "Tenant ID must be a GUID." }
                }
                while (-not (Test-GuidLike $client)) {
                    $client = Read-Setting -Prompt "Client ID (Application ID)" -Default $defClient
                    if (-not (Test-GuidLike $client)) { Write-Log -Level WARNING -Message "Client ID must be a GUID." }
                }
                if ($created) {
                    $graphSecret = $created.Secret
                    Write-Log -Level FOUND -Message "Using the client secret just created."
                }
                elseif ($createdSecret -and $client -eq $createdClient) {
                    $graphSecret = $createdSecret
                    Write-Log -Level FOUND -Message "Using the client secret created earlier in this run."
                }
                else {
                    # A re-run on this server can keep the secret already stored for this same app, so Enter is enough
                    $stored = if ($Saved -and $Saved.MailMethod -eq 'Graph' -and $Saved.CredentialFor -eq $client) { Get-StoredSecret } else { $null }
                    $graphSecret = ""
                    while ([string]::IsNullOrEmpty($graphSecret)) {
                        $secretPrompt = if ($stored) { "Client secret (Enter = keep the one stored on this server, or paste a new one, input hidden)" } else { "Client secret (paste, input hidden)" }
                        $secure = Read-Host $secretPrompt -AsSecureString
                        if ($secure -and $secure.Length -gt 0) { $graphSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)) }
                        elseif ($stored) { $graphSecret = $stored; Write-Log -Level FOUND -Message "Keeping the client secret stored on this server." }
                        else { Write-Log -Level WARNING -Message "Client secret is required." }
                        if ($graphSecret -and (Test-GuidLike $graphSecret)) { Write-Log -Level WARNING -Message "That is a GUID, so it is the secret ID. Paste the secret value, the longer string shown once when the secret was created."; $graphSecret = "" }
                    }
                }
                $cipher = Protect-Secret -PlainText $graphSecret
                if (-not $created) { $expiry = $null }
                while ($null -eq $expiry) {
                    $expiryPrompt = if ($defExpiry) { "Secret expiry date yyyy-MM-dd (Enter = keep, - = no expiry warnings)" } else { "Secret expiry date yyyy-MM-dd (blank = no expiry warnings)" }
                    $typed = Read-Setting -Prompt $expiryPrompt -Default $defExpiry
                    if ($typed -eq '-') { $typed = '' }
                    $expiry = Test-DateInput $typed
                    if ($null -eq $expiry) { Write-Log -Level WARNING -Message "Use the format yyyy-MM-dd." }
                }
                $server = "graph.microsoft.com"; $port = 443; $ssl = $true
            }
            'DirectSend' {
                $disc = Get-DirectSendHost -FromAddress $from
                Write-Log -Level FOUND -Message "Direct Send host: $($disc.Host) ($($disc.Source))."
                $server = Read-Setting -Prompt "SMTP host" -Default $(if ($mem.Server) { $mem.Server } else { $disc.Host })
                $port = 25; $ssl = $false
            }
            'Relay' {
                $server = ""
                while ([string]::IsNullOrWhiteSpace($server)) {
                    $server = Read-Setting -Prompt "Relay host name or IP" -Default $(if ($mem.Server) { $mem.Server } else { "" })
                    if ([string]::IsNullOrWhiteSpace($server)) {
                        if ($NonInteractive) { Write-Log -Level FAILED -Message "Relay host required in non-interactive mode."; return $null }
                        Write-Log -Level WARNING -Message "Relay host is required."
                    }
                }
                $port = Read-PortSetting -Prompt "Port" -Default $(if ($mem.Port) { [int]$mem.Port } else { 25 })
                if ($null -eq $port) { return $null }
                $ssl  = ((Read-Choice -Prompt "Use TLS (STARTTLS)?" -Allowed @('Y','N') -Default $(if ($mem.ContainsKey('Ssl') -and $mem.Ssl) { 'Y' } else { 'N' })) -eq 'Y')
            }
            'Authenticated' {
                $server = Read-Setting -Prompt "SMTP host" -Default $(if ($mem.Server) { $mem.Server } else { "smtp.office365.com" })
                $port   = Read-PortSetting -Prompt "Port" -Default $(if ($mem.Port) { [int]$mem.Port } else { 587 })
                if ($null -eq $port) { return $null }
                $ssl    = ((Read-Choice -Prompt "Use TLS (STARTTLS)?" -Allowed @('Y','N') -Default $(if ($mem.ContainsKey('Ssl')) { if ($mem.Ssl) { 'Y' } else { 'N' } } else { 'Y' })) -eq 'Y')
                $user   = Read-Setting -Prompt "Username" -Default $(if ($mem.User) { $mem.User } else { $from })
                if ($NonInteractive) { Write-Log -Level FAILED -Message "Authenticated SMTP cannot be configured non-interactively because the password must be entered at the console."; return $null }
                # A re-run on this server can keep the password already stored for this same user, so Enter is enough
                $storedPassword = if ($Saved -and $Saved.MailMethod -eq 'Authenticated' -and $Saved.CredentialFor -eq $user) { Get-StoredSecret } else { $null }
                if ($storedPassword) {
                    Write-Question -Question "Password for $user" -Hint "Enter keeps the password already stored on this server. Type a new one to replace it."
                    if ((Read-Choice -Prompt "Keep the stored password" -Allowed @('Y','N') -Default 'Y') -eq 'Y') {
                        $cipher = Protect-Secret -PlainText $storedPassword
                        $cred = New-Object System.Management.Automation.PSCredential($user, (ConvertTo-SecureText -Text $storedPassword))
                        Write-Log -Level FOUND -Message "Keeping the SMTP password stored on this server."
                        $storedPassword = $null
                    }
                }
                if (-not $cipher) { $cred = Get-Credential -UserName $user -Message "Password for $user (stored encrypted, machine-scope, readable by SYSTEM on this server only)" }
                if (-not $cipher -and (-not $cred -or $cred.Password.Length -eq 0)) {
                    Write-Log -Level WARNING -Message $(if ($cred) { "A password is required." } else { "No credential entered." })
                    $defMethod = $method; $defFrom = $from; $defTo = ($to -join ',')
                    $remembered[$method] = @{ Server = $server; Port = $port; Ssl = $ssl; User = $user }
                    $restart = $true
                    break
                }
                if ($cred) {
                    $user   = $cred.UserName
                    $cipher = Protect-Secret -Password $cred.Password
                }
            }
        }

        if ($restart) { continue }

        $mail = @{ MailMethod = $method; SmtpServer = $server; SmtpPort = $port; SmtpUseSsl = $ssl; MailFrom = $from; MailTo = $to; SmtpAuthUser = $user; CipherText = $cipher; GraphTenantId = $tenant; GraphClientId = $client; GraphSecretExpires = $expiry }

        Write-Log -Level INFORMATIONAL -Message "Sending a test email from $from to $($to -join ', ') via $server`:$port (TLS=$ssl, method=$method)."
        $testSubject = "[MONITOR TEST] $(Get-AlertLabel -MonitorName $MonitorName -SiteName $SiteName -HostName $env:COMPUTERNAME)"
        $testBody = New-AlertBody -Heading "Alert delivery test" -Details ([ordered]@{ 'Monitor' = $MonitorName; 'Site' = $SiteName; 'Server' = $env:COMPUTERNAME; 'Mail method' = "$method via $server`:$port"; 'Sent' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); 'Result' = 'If you can read this, alert delivery from this server works.' }) -Footer "Sent by the curl monitor installer on $env:COMPUTERNAME."
        $sent = Send-MailWithConfig -Mail $mail -Subject $testSubject -Body $testBody -Credential $cred -GraphSecret $graphSecret

        if ($NonInteractive) {
            if (-not $sent) { Write-Log -Level WARNING -Message "Test email failed in non-interactive mode. Continuing with the configured settings."; $script:MailTestSkipped = $true }
            return $mail
        }

        if ($sent) {
            Write-Log -Level SUCCESS -Message "The mail host accepted the test message. Check the inbox (and junk folder) now."
            $answer = Read-Choice -Prompt "Did the test email arrive? Y = continue, N = change settings, S = skip mail setup and continue anyway" -Allowed @('Y','N','S') -Default 'Y'
        }
        elseif ($method -eq 'Graph' -and $script:LastGraphSendError -eq 'AccessDenied' -and ($created -or ($createdClient -and $client -eq $createdClient))) {
            # Exchange applies a new send grant to live sends 30 minutes to 2 hours after it is created, and every failed
            # attempt keeps the old authorization cached. Waiting quietly is the only thing that shortens it. A rejected
            # secret never lands here: sign-in retries are handled inside Get-InstallerGraphToken.
            Write-Log -Level INFORMATIONAL -Message "This is expected right after first-time setup: Exchange has not applied the send grant yet (ErrorAccessDenied). Microsoft documents 30 minutes to 2 hours."
            $answer = Read-Choice -Prompt "W = wait and retry automatically (quiet 30 min, then every 10 min, up to 2 h), R = change settings, S = skip and finish (the monitor will send once it propagates)" -Allowed @('W','R','S') -Default 'W'
            if ($answer -eq 'W') {
                $deadline = (Get-Date).AddHours(2); $ok = $false; $waitMinutes = 30
                $script:QuietAccessDeniedHint = $true
                while ((Get-Date) -lt $deadline -and -not $ok) {
                    Write-Log -Level INFORMATIONAL -Message "Waiting $waitMinutes min without retrying. Next attempt at $((Get-Date).AddMinutes($waitMinutes).ToString('HH:mm'))."
                    Start-Sleep -Seconds ($waitMinutes * 60)
                    $testBody = New-AlertBody -Heading "Alert delivery test" -Details ([ordered]@{ 'Monitor' = $MonitorName; 'Site' = $SiteName; 'Server' = $env:COMPUTERNAME; 'Mail method' = "$method via $server`:$port"; 'Sent' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); 'Result' = 'If you can read this, alert delivery from this server works.' }) -Footer "Sent by the curl monitor installer on $env:COMPUTERNAME."
                    $ok = Send-MailWithConfig -Mail $mail -Subject $testSubject -Body $testBody -Credential $cred -GraphSecret $graphSecret
                    $waitMinutes = 10
                }
                $script:QuietAccessDeniedHint = $false
                if ($ok) { $sent = $true; Write-Log -Level SUCCESS -Message "The grant propagated and the test message was accepted. Check the inbox."; $answer = Read-Choice -Prompt "Did the test email arrive? Y = continue, N = change settings, S = skip and continue anyway" -Allowed @('Y','N','S') -Default 'Y' }
                else { Write-Log -Level WARNING -Message "Still not propagated after 2 hours. Finishing anyway; the monitor will send once it clears. If it persists, check the Exchange role assignment and scope."; $answer = 'S' }
            }
            elseif ($answer -eq 'R') { $answer = 'N' }
        }
        else {
            if ($method -eq 'Graph' -and $script:LastGraphSendError -eq 'InvalidSecret') {
                if ($createdClient -and $client -eq $createdClient) {
                    Write-Log -Level INFORMATIONAL -Message "Choose R and take the default N at the app prompt: that signs you in again as Global Administrator, reuses the app, and prints a new secret. The secret shown earlier is no longer valid. Or choose R, answer Y, and paste a secret created for this app in Entra."
                    $script:RetryTenantSetup = $true; $createdSecret = $null
                }
                else { Write-Log -Level INFORMATIONAL -Message "Paste the secret value, not the secret ID. A secret created in the last few minutes may still be replicating: wait a minute, choose R, and paste it again." }
            }
            $answer = Read-Choice -Prompt "Test email failed. R = change settings and retry, S = skip and continue anyway" -Allowed @('R','S') -Default 'R'
            if ($answer -eq 'R') { $answer = 'N' }
        }

        if ($answer -eq 'S' -and -not $sent) { $script:MailTestSkipped = $true }
        if ($answer -eq 'Y' -or $answer -eq 'S') { return $mail }
        # Remember only what this method used, so a detour through another method never erases saved values
        $defMethod = $method; $defFrom = $from; $defTo = ($to -join ',')
        if ($method -eq 'Graph') { $defTenant = $tenant; $defClient = $client; $defExpiry = $expiry }
        else { $remembered[$method] = @{ Server = $server; Port = $port; Ssl = $ssl; User = $user } }
    }
}

function ConvertTo-MonitorSlug {
    # Folder-safe form of the monitor name: letters, digits and dashes, no runs, no leading or trailing dash
    param([string]$Name)
    if ($null -eq $Name) { return "" }
    $slug = ($Name -replace '[^A-Za-z0-9]+', '-').Trim('-')
    if ($slug.Length -gt 60) { $slug = $slug.Substring(0, 60).Trim('-') }
    return $slug
}

function Get-InstalledMonitor {
    # Monitor names already installed under the root, read from each folder's saved settings
    param([string]$Root = $InstallRoot)
    $found = @()
    if (-not (Test-Path $Root)) { return $found }
    foreach ($dir in @(Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue)) {
        $settingsPath = Join-Path $dir.FullName $SettingsFileName
        if (-not (Test-Path $settingsPath)) { continue }
        try {
            $json = Get-Content -Path $settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $name = if ($json.MonitorName) { [string]$json.MonitorName } else { $dir.Name }
            $found += [PSCustomObject]@{ Name = $name; Slug = $dir.Name; Path = $dir.FullName; Url = [string]$json.Url }
        }
        catch { $found += [PSCustomObject]@{ Name = $dir.Name; Slug = $dir.Name; Path = $dir.FullName; Url = '' } }
    }
    return $found
}

function Get-MonitorName {
    # Returns the monitor name from the override, the only existing monitor, a legacy install, or a prompt.
    # $Elsewhere holds monitors installed somewhere this run is not writing to, such as ones left in the old
    # location when the move was declined. Reusing one of those names would leave two monitors on the same URL.
    param([string]$SavedDefault = "", [object[]]$Existing = @(), [object[]]$Elsewhere = @())
    # PowerShell 7 hands an empty result to an [object[]] parameter as one $null, which would list a blank monitor
    $Existing = @($Existing | Where-Object { $null -ne $_ })
    $Elsewhere = @($Elsewhere | Where-Object { $null -ne $_ })
    if (-not [string]::IsNullOrWhiteSpace($MonitorNameOverride)) {
        $clean = ConvertTo-SafeSiteName $MonitorNameOverride
        if ($clean -and ($TaskNamePrefix + $clean).Length -gt $MaxTaskNameLength) {
            Write-Log -Level FAILED -Message "Monitor name override '$clean' is too long for a task name."
            return ""
        }
        if ($clean -and @(@($Elsewhere) | Where-Object { $_.Name -eq $clean -or ($_.Slug -and $_.Slug -eq (ConvertTo-MonitorSlug $clean)) }).Count -gt 0 -and @(@($Existing) | Where-Object { $_.Name -eq $clean -or $_.Slug -eq (ConvertTo-MonitorSlug $clean) }).Count -eq 0) {
            Write-Log -Level FAILED -Message "'$clean' is still installed outside '$InstallRoot' and could not be moved. Installing it here as well would leave two monitors on the same URL. Nothing was installed."
            return ""
        }
        if ($clean) { Write-Log -Level FOUND -Message "Monitor name set from override: '$clean'."; return $clean }
    }
    $default = $SavedDefault
    if (-not $default) { $default = ConvertTo-SafeSiteName $MonitorName }
    if ($NonInteractive) {
        $clean = ConvertTo-SafeSiteName $default
        if (-not $clean) { Write-Log -Level FAILED -Message "No monitor name. Set `$MonitorNameOverride or `$MonitorName in the config block."; return "" }
        $awaySilent = @(@($Elsewhere) | Where-Object { $_.Name -eq $clean -or ($_.Slug -and $_.Slug -eq (ConvertTo-MonitorSlug $clean)) })
        $hereSilent = @(@($Existing) | Where-Object { $_.Name -eq $clean -or $_.Slug -eq (ConvertTo-MonitorSlug $clean) })
        if (@($awaySilent).Count -gt 0 -and @($hereSilent).Count -eq 0) {
            Write-Log -Level FAILED -Message "'$clean' is still installed outside '$InstallRoot' and could not be moved. Installing it here as well would leave two monitors on the same URL. Nothing was installed."
            return ""
        }
        if (@($awaySilent).Count -gt 0) {
            Write-Log -Level WARNING -Message "'$clean' is installed here and a copy is still at '$(@($awaySilent)[0].Dir)'. This run upgrades the one here; remove the other with the uninstall option or it keeps polling the same URL."
        }
        if (($TaskNamePrefix + $clean).Length -gt $MaxTaskNameLength) { Write-Log -Level FAILED -Message "Monitor name '$clean' is too long for a task name."; return "" }
        $clash = Get-SlugClash -Name $clean -Existing $Existing
        if ($clash) {
            Write-Log -Level WARNING -Message "'$clean' uses the same folder as the installed monitor '$($clash.Name)'. Upgrading that monitor instead of overwriting it."
            return $clash.Name
        }
        Write-Log -Level FOUND -Message "Monitor name set to '$clean' (non-interactive)."
        return $clean
    }
    # Only ever a prompt default: a silent run has to be told which monitor it is installing
    if (-not $default -and @($Existing).Count -eq 1) { $default = @($Existing)[0].Name }
    if (-not $default -and @($Existing).Count -gt 1) { $default = @(@($Existing) | Sort-Object { (Get-Item $_.Path).LastWriteTime } -Descending)[0].Name }
    if (@($Existing).Count -gt 0 -or @($Elsewhere).Count -gt 0) {
        Write-Host ""
        Write-Host "Already installed on this server"
        foreach ($m in @($Existing)) { Write-Host ("  {0}  {1}" -f $m.Name, $(if ($m.Url) { $m.Url } else { 'URL not recorded' })) }
        foreach ($m in @($Elsewhere)) { Write-Host ("  {0}  {1}  (still in {2})" -f $m.Name, $(if ($m.Url) { $m.Url } else { 'URL not recorded' }), (Split-Path $m.Dir -Parent)) }
    }
    $hint = @("Names its folder, its task, and every alert subject. For example: GitHub Monitor")
    if (@($Existing).Count -gt 0) { $hint += "One of the names above upgrades that monitor. A new name adds another beside it." }
    while ($true) {
        Write-Question -Question "Monitor name" -Hint $hint
        $typed = Read-Setting -Prompt "Name" -Default $default
        $clean = ConvertTo-SafeSiteName $typed
        if (-not $clean) { Write-Log -Level WARNING -Message "Monitor name cannot be empty after cleanup."; continue }
        if (-not (ConvertTo-MonitorSlug $clean)) { Write-Log -Level WARNING -Message "Monitor name needs at least one letter or digit."; continue }
        if (-not (Test-MonitorNameLength $clean)) { continue }
        if ($clean -ne $typed.Trim()) {
            if ((Read-Choice -Prompt "Cleaned to '$clean'. Use it?" -Allowed @('Y','N') -Default 'Y') -ne 'Y') { continue }
        }
        # Two names can clean to one folder ('HST eChart' and 'HST_eChart', or a case-only change). Adopting the
        # installed name upgrades that monitor; anything else would overwrite its files while leaving its task behind.
        # A monitor still installed somewhere else would keep polling beside the new one
        $slugWanted = ConvertTo-MonitorSlug $clean
        $away = @(@($Elsewhere) | Where-Object { $_.Name -eq $clean -or ($_.Slug -and $_.Slug -eq $slugWanted) })
        $here = @(@($Existing) | Where-Object { $_.Name -eq $clean -or $_.Slug -eq $slugWanted })
        if (@($away).Count -gt 0 -and @($here).Count -gt 0) {
            # The folder here is already taken, so the other copy cannot move onto it. Upgrade this one and say so.
            Write-Log -Level WARNING -Message "'$clean' is installed here and a copy is still at '$(@($away)[0].Dir)'. This run upgrades the one here. Remove the other with the uninstall option or it keeps polling the same URL."
        }
        elseif (@($away).Count -gt 0) {
            $one = @($away)[0]
            $taskNote = if ($one.HasTask) { "its task '$($one.TaskPath)$($one.TaskName)' is still polling" } else { "it has no task of its own" }
            Write-Log -Level WARNING -Message "'$($one.Name)' is still installed at '$($one.Dir)' and $taskNote. Installing under this name as well would leave two monitors on the same URL."
            Write-Question -Question "Move it here first?" -Hint @(
                "M  move '$($one.Name)' into $InstallRoot and upgrade it in this run",
                "N  type a different name for the new monitor"
            )
            if ((Read-Choice -Prompt "Choose" -Allowed @('M','N') -Default 'M') -ne 'M') { continue }
            if (-not (Move-MonitorToNewRoot -Monitor $one)) {
                Write-Log -Level WARNING -Message "'$($one.Name)' could not be moved, so it is still where it was. Give this run a different name, or fix the move and run again."
                continue
            }
            return $one.Name
        }
        $clash = Get-SlugClash -Name $clean -Existing $Existing
        if ($clash) {
            Write-Log -Level WARNING -Message "'$clean' uses the same folder as the installed monitor '$($clash.Name)' ($($clash.Path))."
            $answer = Read-Choice -Prompt "U = upgrade '$($clash.Name)' in place, N = type another name" -Allowed @('U','N') -Default 'U'
            if ($answer -ne 'U') { continue }
            Write-Log -Level FOUND -Message "Upgrading the installed monitor '$($clash.Name)'."
            return $clash.Name
        }
        return $clean
    }
}

function Test-MonitorNameLength {
    # Task Scheduler rejects a name past 238 characters, so the wizard refuses one before anything is written
    param([Parameter(Mandatory)][string]$Name)
    if (($TaskNamePrefix + $Name).Length -le $MaxTaskNameLength) { return $true }
    $max = [math]::Max(1, $MaxTaskNameLength - $TaskNamePrefix.Length)
    Write-Log -Level WARNING -Message "Monitor name is too long. Keep it to $max characters or fewer, so the task name '$TaskNamePrefix<name>' stays inside the Task Scheduler limit."
    return $false
}

function Get-SlugClash {
    # The installed monitor whose folder this name would land in, when that monitor goes by another name
    param([Parameter(Mandatory)][string]$Name, [object[]]$Existing = @())
    $Existing = @($Existing | Where-Object { $null -ne $_ })
    $slug = ConvertTo-MonitorSlug $Name
    foreach ($m in @($Existing)) {
        # Slugs match case-insensitively the way NTFS does; a name that differs only in case still has to adopt
        # the installed spelling, or the folder and the task would drift apart.
        if ($m.Slug -and $m.Slug -eq $slug -and $m.Name -cne $Name) { return $m }
    }
    return $null
}

function Test-MonitorUrl {
    # True for an absolute http or https URL
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $u = $null
    if (-not [uri]::TryCreate($Text.Trim(), [UriKind]::Absolute, [ref]$u)) { return $false }
    return ($u.Scheme -eq 'http' -or $u.Scheme -eq 'https')
}

function Get-MonitorUrl {
    # Returns the URL to watch from the config block, saved settings, or a prompt
    param([string]$SavedDefault = "")
    # A silent re-deploy has to be able to repoint a monitor, so the config block wins over the saved value there
    $default = if ($NonInteractive -and (Test-MonitorUrl $Url)) { $Url } elseif ($SavedDefault) { $SavedDefault } elseif (Test-MonitorUrl $Url) { $Url } else { "" }
    if ($NonInteractive) {
        if (Test-MonitorUrl $default) { Write-Log -Level FOUND -Message "URL set to '$default' (non-interactive)."; return $default.Trim() }
        Write-Log -Level FAILED -Message "No valid URL. Set `$Url in the config block."
        return ""
    }
    while ($true) {
        Write-Question -Question "URL to watch" -Hint "Requested every $IntervalSeconds seconds, following up to $MaxRedirects redirects."
        $typed = Read-Setting -Prompt "URL" -Default $default
        if (Test-MonitorUrl $typed) { return $typed.Trim() }
        Write-Log -Level WARNING -Message "Enter a full URL starting with http:// or https://."
    }
}

function Get-ContentMarker {
    # Returns the text that must appear on the page, or an empty string to check only the status and the size
    param([string]$SavedDefault = "")
    if ($NonInteractive) { return [string]$(if ($ExpectedContentMarker) { $ExpectedContentMarker } elseif ($SavedDefault) { $SavedDefault } else { '' }) }
    $default = if ($SavedDefault) { $SavedDefault } else { $ExpectedContentMarker }
    $hint = @("Words from the page, so a page that loads but comes back wrong still counts as down.",
              "Blank accepts any page that returns HTTP 200 and is at least $MinPopulatedBytes bytes.")
    $label = "Text"
    if ($default) {
        $hint += "Enter keeps '$default'. A single - drops the check."
        $label = "Text (- to drop)"
    }
    Write-Question -Question "Text the page must contain (optional)" -Hint $hint
    $typed = Read-Setting -Prompt $label -Default $default
    $typed = "$typed".Trim()
    if ($typed -eq '-') { $typed = '' }
    return $typed
}

function Get-LegacyInstall {
    # Finds an older HST-only install, returning its folder, task, and saved settings, or $null. That layout
    # always kept its task in the old folder, so $PreviousTaskPath is where it is looked for.
    param([string]$Dir = $LegacyInstallDir, [string]$TaskName = $LegacyTaskName, [string]$Path = $PreviousTaskPath)
    $settingsPath = Join-Path $Dir $SettingsFileName
    $task = Get-ScheduledTask -TaskName $TaskName -TaskPath $Path -ErrorAction SilentlyContinue
    $hasFiles = (Test-Path $settingsPath) -or (Test-Path (Join-Path $Dir 'Watch-HSTeChartUptime.ps1'))
    if (-not $hasFiles -and -not $task) { return $null }
    $settings = $null
    if (Test-Path $settingsPath) {
        try { $settings = Get-Content -Path $settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { }
    }
    return [PSCustomObject]@{
        Dir       = $Dir
        TaskName  = $TaskName
        TaskPath  = $Path
        HasTask   = [bool]$task
        Settings  = $settings
        Url       = [string]$(if ($settings -and $settings.Url) { $settings.Url } else { '' })
        SiteName  = [string]$(if ($settings -and $settings.SiteName) { $settings.SiteName } else { '' })
    }
}

function Invoke-LegacyMigration {
    # Moves an older HST-only install into the per-monitor layout: stops and unregisters its task, copies its
    # history under the new names, verifies every copy, then removes the old folder. Returns $true when the old
    # folder is gone. On any copy problem the old folder is left alone and reported.
    param([Parameter(Mandatory)][object]$Legacy, [Parameter(Mandatory)][string]$Destination)
    if ($Legacy.HasTask) {
        try {
            Stop-ScheduledTask -TaskName $Legacy.TaskName -TaskPath $Legacy.TaskPath -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Unregister-ScheduledTask -TaskName $Legacy.TaskName -TaskPath $Legacy.TaskPath -Confirm:$false -ErrorAction Stop
            Write-Log -Level INFORMATIONAL -Message "Stopped and removed the old task '$($Legacy.TaskPath)$($Legacy.TaskName)'."
        }
        catch { Write-Log -Level WARNING -Message "Could not remove the old task '$($Legacy.TaskPath)$($Legacy.TaskName)'. Remove it by hand or it keeps polling. $($_.Exception.Message)" }
    }
    if (-not (Test-Path $Legacy.Dir)) { return $true }
    $copied = 0; $failed = 0; $kept = 0
    foreach ($file in @(Get-ChildItem -Path $Legacy.Dir -File -Force -Recurse -ErrorAction SilentlyContinue)) {
        $target = switch -Regex ($file.Name) {
            '^HST-eChart-Latency_(.+)\.csv$'  { "Latency_$($Matches[1]).csv" ; break }
            '^HST-eChart-Outages\.csv$'       { 'Outages.csv' ; break }
            '^HST-eChart-Drops\.log$'         { 'Drops.log' ; break }
            '^HST-eChart-Drops_(.+)\.log$'    { "Drops_$($Matches[1]).log" ; break }
            '^HST-eChart-Monitor_(.+)\.log$'  { "Transcript_$($Matches[1]).log" ; break }
            '^monitor-heartbeat\.json$'       { 'heartbeat.json' ; break }
            '^smtp-credential\.bin$'          { $CredentialFileName ; break }
            '^daily-summary-sent\.txt$'       { 'summary-sent.txt' ; break }
            '^install-settings\.json$'        { $SettingsFileName ; break }
            '^HST-Monitor-AppRegistration\.txt$' { 'AppRegistration.txt' ; break }
            default { '' }
        }
        # Anything unrecognised, including files in subfolders, is carried under its own name rather than deleted
        $relative = $file.FullName.Substring($Legacy.Dir.Length).TrimStart('\')
        if (-not $target) { $target = 'legacy-' + ($relative -replace '[\\/]', '-') }
        $targetPath = Join-Path $Destination $target
        # Never write over history the new monitor has already recorded: a second migration must not roll it back
        $existing = Get-Item -LiteralPath $targetPath -Force -ErrorAction SilentlyContinue
        if ($existing) {
            if ($existing.Length -eq $file.Length -and $existing.LastWriteTimeUtc -ge $file.LastWriteTimeUtc) { $kept++; continue }
            $targetPath = Join-Path $Destination ("legacy-" + (Get-Date -Format 'yyyyMMdd_HHmmss') + "-" + $target)
            Write-Log -Level INFORMATIONAL -Message "'$target' already exists here, so '$($file.Name)' came in as '$(Split-Path $targetPath -Leaf)'."
        }
        try {
            Copy-Item -LiteralPath $file.FullName -Destination $targetPath -Force -ErrorAction Stop
            # -Force on Get-Item so a hidden file, which Test-Path skips, still verifies
            $landed = Get-Item -LiteralPath $targetPath -Force -ErrorAction SilentlyContinue
            if (-not $landed -or $landed.Length -ne $file.Length) { throw "copy is not the same size" }
            if ($target -eq $CredentialFileName) {
                # The stored secret arrives with inherited permissions, so lock it before any prompt can abort the run
                & icacls.exe "$targetPath" /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "icacls could not restrict the carried credential file (exit $LASTEXITCODE)" }
            }
            $copied++
        }
        catch { $failed++; Write-Log -Level WARNING -Message "Could not carry '$($file.Name)' over. $($_.Exception.Message)" }
    }
    Write-Log -Level CREATED -Message "Carried $copied file(s) from '$($Legacy.Dir)' into '$Destination'$(if ($kept) { ", left $kept already here alone" } else { '' })."
    # The old monitor was stopped on purpose, so mark the heartbeat: no restart notice, any outage still carries over
    $hb = Join-Path $Destination 'heartbeat.json'
    if (Test-Path $hb) {
        try {
            $doc = Get-Content -Path $hb -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $doc | Add-Member -NotePropertyName Stopped -NotePropertyValue $true -Force
            $doc | ConvertTo-Json -Compress | Set-Content -Path $hb -Encoding UTF8 -Force
        }
        catch { Remove-Item -Path $hb -Force -ErrorAction SilentlyContinue }
    }
    try { Add-Content -Path (Join-Path $Destination 'Drops.log') -Value ("{0} | {1,-9} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), 'STOP', "Monitor stopped by the installer on $env:COMPUTERNAME and moved to '$Destination'.") -Encoding UTF8 -ErrorAction Stop } catch { }
    if ($failed -gt 0) {
        Write-Log -Level WARNING -Message "$failed file(s) did not copy, so '$($Legacy.Dir)' is left in place. Move what you need by hand and delete it when you are done."
        return $false
    }
    try {
        Remove-Item -Path $Legacy.Dir -Recurse -Force -ErrorAction Stop
        Write-Log -Level INFORMATIONAL -Message "Removed the old folder '$($Legacy.Dir)'."
        return $true
    }
    catch {
        Write-Log -Level WARNING -Message "History was carried over but '$($Legacy.Dir)' could not be removed. Delete it by hand. $($_.Exception.Message)"
        return $false
    }
}

function Register-MonitorTask {
    # Registers or replaces a monitor's task. Shared by the install and by the move out of the old location.
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$ScriptPath)
    $action      = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
    $trigger     = New-ScheduledTaskTrigger -AtStartup
    $settingsSet = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -RestartCount $RestartCount -RestartInterval (New-TimeSpan -Minutes $RestartMinutes) -ExecutionTimeLimit ([TimeSpan]::Zero) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd
    $principal   = New-ScheduledTaskPrincipal -UserId $RunAsUser -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $Name -TaskPath $Path -Action $action -Trigger $trigger -Settings $settingsSet -Principal $principal -Force -ErrorAction Stop | Out-Null
    # CIM cmdlets can report failure without throwing, so confirm the task actually exists before claiming success
    if (-not (Get-ScheduledTask -TaskName $Name -TaskPath $Path -ErrorAction Stop)) { throw "The task was not found after registration." }
}

function Get-PreviousRootMonitor {
    # Monitors still installed where they lived before the parent folder was dropped
    param([string]$Root = $PreviousRoot, [string]$Path = $PreviousTaskPath, [string]$Prefix = $TaskNamePrefix, [string]$NewRoot = $InstallRoot)
    $found = @()
    if (-not $Root -or $Root -eq $NewRoot) { return $found }
    foreach ($m in @(Get-InstalledMonitor -Root $Root)) {
        $taskName = $Prefix + $m.Name
        $task = Get-ScheduledTask -TaskName $taskName -TaskPath $Path -ErrorAction SilentlyContinue
        $found += [PSCustomObject]@{
            Name = $m.Name; Slug = $m.Slug; Dir = $m.Path; Url = $m.Url
            TaskName = $taskName; TaskPath = $Path; HasTask = [bool]$task
        }
    }
    return $found
}

function Move-MonitorToNewRoot {
    # Moves one monitor out of the old location: stops its task, moves the folder, repoints the paths baked into
    # its script, and registers the task again under the new folder. If the move fails the old task goes back, so
    # the server is never left with nothing watching.
    param([Parameter(Mandatory)]$Monitor, [string]$NewRoot = $InstallRoot, [string]$NewTaskPath = $TaskPath)
    $dest = Join-Path $NewRoot $Monitor.Slug
    if (Test-Path -LiteralPath $dest) {
        Write-Log -Level WARNING -Message "'$dest' already exists, so '$($Monitor.Name)' stays at '$($Monitor.Dir)'. Remove one of the two with the uninstall option, then run this again."
        return $false
    }
    $oldScript = Join-Path $Monitor.Dir $MonitorFileName
    if (-not (Test-Path -LiteralPath $oldScript)) {
        Write-Log -Level WARNING -Message "'$($Monitor.Dir)' holds no '$MonitorFileName', so '$($Monitor.Name)' is an older layout this move does not understand. It stays where it is: run the installer again and answer Y when it offers to move the older install."
        return $false
    }
    $hadTask = $false
    if ($Monitor.HasTask -and (Get-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue)) {
        $hadTask = $true
        try { Stop-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue } catch { }
        for ($i = 0; $i -lt 15; $i++) {
            $t = Get-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue
            if (-not $t -or "$($t.State)" -ne 'Running') { break }
            Start-Sleep -Seconds 1
        }
        try { Unregister-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -Confirm:$false -ErrorAction Stop }
        catch {
            Write-Log -Level WARNING -Message "Could not remove the old task '$($Monitor.TaskPath)$($Monitor.TaskName)', so '$($Monitor.Name)' stays where it is. $($_.Exception.Message)"
            return $false
        }
    }
    $null = Stop-MonitorProcess -Dir $Monitor.Dir
    # The monitor reports its own restarts, so record that this stop was deliberate
    $hb = Join-Path $Monitor.Dir 'heartbeat.json'
    if (Test-Path -LiteralPath $hb) {
        try {
            $doc = Get-Content -Path $hb -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $doc | Add-Member -NotePropertyName Stopped -NotePropertyValue $true -Force
            $doc | ConvertTo-Json -Compress | Set-Content -Path $hb -Encoding UTF8 -Force
        }
        catch { Remove-Item -Path $hb -Force -ErrorAction SilentlyContinue }
    }
    try { Move-Item -LiteralPath $Monitor.Dir -Destination $dest -ErrorAction Stop }
    catch {
        Write-Log -Level WARNING -Message "Could not move '$($Monitor.Dir)'. $($_.Exception.Message)"
        if ($hadTask) {
            try {
                Register-MonitorTask -Name $Monitor.TaskName -Path $Monitor.TaskPath -ScriptPath $oldScript
                Start-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue
                Write-Log -Level INFORMATIONAL -Message "'$($Monitor.Name)' is running again from '$($Monitor.Dir)'."
            }
            catch { Write-Log -Level FAILED -Message "'$($Monitor.Name)' is not running any more. Re-run the installer and give it the same name." }
        }
        return $false
    }
    $script = Join-Path $dest $MonitorFileName
    if (Test-Path -LiteralPath $script) {
        try {
            $text = Get-Content -LiteralPath $script -Raw -ErrorAction Stop
            $moved = $text.Replace($Monitor.Dir, $dest)
            [ScriptBlock]::Create($moved) | Out-Null
            Set-Content -LiteralPath $script -Value $moved -Encoding UTF8 -Force -ErrorAction Stop
        }
        catch { Write-Log -Level WARNING -Message "Could not repoint '$script' at its new folder. Re-run the installer with this monitor's name to rebuild it. $($_.Exception.Message)" }
    }
    try { Protect-InstallFolder -Path $dest } catch { Write-Log -Level WARNING -Message "Moved '$($Monitor.Name)' but could not lock '$dest'. $($_.Exception.Message)" }
    $cred = Join-Path $dest $CredentialFileName
    if (Test-Path -LiteralPath $cred) { & icacls.exe "$cred" /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null }
    try {
        Register-MonitorTask -Name $Monitor.TaskName -Path $NewTaskPath -ScriptPath $script
        Start-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $NewTaskPath -ErrorAction SilentlyContinue
        $state = 'Unknown'
        for ($i = 0; $i -lt 10; $i++) {
            Start-Sleep -Seconds 1
            $state = "$((Get-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $NewTaskPath -ErrorAction SilentlyContinue).State)"
            if ($state -eq 'Running') { break }
        }
        if ($state -eq 'Running') {
            Write-Log -Level CREATED -Message "Moved '$($Monitor.Name)' to '$dest', now running from '$NewTaskPath$($Monitor.TaskName)'."
            return $true
        }
        Write-Log -Level FAILED -Message "Moved '$($Monitor.Name)' to '$dest' and registered '$NewTaskPath$($Monitor.TaskName)', but it is '$state' rather than running. It starts at the next boot; start it now with Start-ScheduledTask, or let this run reinstall it."
        return $true
    }
    catch {
        Write-Log -Level FAILED -Message "Moved '$($Monitor.Name)' to '$dest' but could not register its task. Re-run the installer and give it the same name. $($_.Exception.Message)"
        return $true
    }
}

function Invoke-RootMove {
    # Offers to move every monitor out of the old location. Returns how many moved.
    param([object[]]$Monitors, [string]$NewRoot = $InstallRoot, [string]$OldRoot = $PreviousRoot)
    # A call that returned nothing arrives here as one $null under PowerShell 7, so empty really means empty
    $list = @($Monitors | Where-Object { $null -ne $_ })
    if (@($list).Count -eq 0) { return 0 }
    Write-Host ""
    Write-Host "Found in the old location $OldRoot"
    foreach ($m in $list) { Write-Host ("  {0}  {1}" -f $m.Name, $(if ($m.Url) { $m.Url } else { 'URL not recorded' })) }
    if (-not $NonInteractive) {
        Write-Question -Question "Move to ${NewRoot}?" -Hint @(
            "Each one keeps its history and its settings and starts again from the new folder.",
            "N leaves them running where they are."
        )
        if ((Read-Choice -Prompt "Move" -Allowed @('Y','N') -Default 'Y') -ne 'Y') {
            Write-Log -Level INFORMATIONAL -Message "Left $(@($list).Count) monitor(s) in '$OldRoot'."
            return 0
        }
    }
    $moved = 0
    foreach ($m in $list) { if (Move-MonitorToNewRoot -Monitor $m -NewRoot $NewRoot) { $moved++ } }
    if ($moved) {
        $telemetry = Get-ScheduledTask -TaskName $TelemetryTaskName -TaskPath $PreviousTaskPath -ErrorAction SilentlyContinue
        if (-not $telemetry) { $telemetry = Get-ScheduledTask -TaskName $TelemetryTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue }
        if ($telemetry) { Write-Log -Level WARNING -Message "The telemetry publisher still reads '$OldRoot'. Re-run Install-TelemetryPublisher.ps1 so it reads '$NewRoot'." }
    }
    if ((Test-Path -LiteralPath $OldRoot) -and @(Get-ChildItem -LiteralPath $OldRoot -Force -ErrorAction SilentlyContinue).Count -eq 0) {
        try { Remove-Item -LiteralPath $OldRoot -Force -ErrorAction Stop } catch { }
    }
    return $moved
}

function Get-FolderSizeText {
    # Human-sized total of a folder, for the removal listing
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return 'no folder' }
    $bytes = 0
    try { $bytes = [int64](Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum } catch { }
    if ($bytes -ge 1MB) { return "{0:N1} MB" -f ($bytes / 1MB) }
    if ($bytes -ge 1KB) { return "{0:N0} KB" -f ($bytes / 1KB) }
    return "$bytes bytes"
}

function Get-RemovableMonitor {
    # Everything an uninstall could remove: installed monitors, a task whose folder was deleted by hand, and the
    # older HST-only layout. Paths and task names are parameters so tests drive scratch roots.
    param(
        [string]$Root = $InstallRoot,
        [string]$Path = $TaskPath,
        [string]$Prefix = $TaskNamePrefix,
        [string]$LegacyDir = $LegacyInstallDir,
        [string]$LegacyTask = $LegacyTaskName,
        [string]$LegacyName = $LegacyMonitorName,
        [string]$PrevRoot = $PreviousRoot,
        [string]$PrevPath = $PreviousTaskPath
    )
    $items = New-Object System.Collections.ArrayList
    $seen = @()
    foreach ($m in @(Get-InstalledMonitor -Root $Root)) {
        $taskName = $Prefix + $m.Name
        $task = Get-ScheduledTask -TaskName $taskName -TaskPath $Path -ErrorAction SilentlyContinue
        if (-not $task) {
            # The settings file may not name the monitor, so fall back to the task that runs out of this folder
            $owned = @(Get-ScheduledTask -TaskPath $Path -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -and $_.TaskName.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase) -and "$($_.Actions[0].Arguments)" -like "*$($m.Path)\*" })
            if (@($owned).Count -eq 1) { $task = @($owned)[0]; $taskName = @($owned)[0].TaskName }
        }
        $seen += $taskName
        [void]$items.Add([PSCustomObject]@{
            Name = $m.Name; Slug = $m.Slug; Dir = $m.Path; Url = $m.Url; Kind = 'Monitor'
            TaskName = $taskName; TaskPath = $Path; TaskState = $(if ($task) { "$($task.State)" } else { 'no task' })
        })
    }
    # A removal that could not finish leaves a folder with no settings file behind. It still has to be listed, or
    # the operator can never finish the job through the installer.
    $seenDirs = @(@($items) | ForEach-Object { $_.Dir })
    foreach ($dir in @(Get-ChildItem -Path $Root -Directory -Force -ErrorAction SilentlyContinue)) {
        if ($seenDirs -contains $dir.FullName) { continue }
        $taskName = $Prefix + $dir.Name
        $task = Get-ScheduledTask -TaskName $taskName -TaskPath $Path -ErrorAction SilentlyContinue
        $seen += $taskName
        [void]$items.Add([PSCustomObject]@{
            Name = $dir.Name; Slug = $dir.Name; Dir = $dir.FullName; Url = ''; Kind = 'Leftover'
            TaskName = $taskName; TaskPath = $Path; TaskState = $(if ($task) { "$($task.State)" } else { 'no task' })
        })
    }
    foreach ($task in @(Get-ScheduledTask -TaskPath $Path -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -and $_.TaskName.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase) })) {
        if ($seen -contains $task.TaskName) { continue }
        $name = $task.TaskName.Substring($Prefix.Length)
        [void]$items.Add([PSCustomObject]@{
            Name = $name; Slug = (ConvertTo-MonitorSlug $name); Dir = (Join-Path $Root (ConvertTo-MonitorSlug $name)); Url = ''; Kind = 'TaskOnly'
            TaskName = $task.TaskName; TaskPath = $Path; TaskState = "$($task.State)"
        })
    }
    foreach ($m in @(Get-PreviousRootMonitor -Root $PrevRoot -Path $PrevPath -Prefix $Prefix -NewRoot $Root)) {
        $seen += $m.TaskName
        [void]$items.Add([PSCustomObject]@{
            Name = "$($m.Name) (old location)"; Slug = $m.Slug; Dir = $m.Dir; Url = $m.Url; Kind = 'OldLocation'
            TaskName = $m.TaskName; TaskPath = $m.TaskPath; TaskState = $(if ($m.HasTask) { "$((Get-ScheduledTask -TaskName $m.TaskName -TaskPath $m.TaskPath -ErrorAction SilentlyContinue).State)" } else { 'no task' })
        })
    }
    if ($PrevRoot -and $PrevRoot -ne $Root) {
        $seenPrevDirs = @(@($items) | Where-Object { $_.TaskPath -eq $PrevPath } | ForEach-Object { $_.Dir })
        foreach ($dir in @(Get-ChildItem -Path $PrevRoot -Directory -Force -ErrorAction SilentlyContinue)) {
            if ($seenPrevDirs -contains $dir.FullName) { continue }
            $taskName = $Prefix + $dir.Name
            $task = Get-ScheduledTask -TaskName $taskName -TaskPath $PrevPath -ErrorAction SilentlyContinue
            $seen += $taskName
            [void]$items.Add([PSCustomObject]@{
                Name = "$($dir.Name) (old location)"; Slug = $dir.Name; Dir = $dir.FullName; Url = ''; Kind = 'Leftover'
                TaskName = $taskName; TaskPath = $PrevPath; TaskState = $(if ($task) { "$($task.State)" } else { 'no task' })
            })
        }
        foreach ($task in @(Get-ScheduledTask -TaskPath $PrevPath -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -and $_.TaskName.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase) })) {
            if (@(@($items) | Where-Object { $_.TaskPath -eq $PrevPath -and $_.TaskName -eq $task.TaskName }).Count) { continue }
            $name = $task.TaskName.Substring($Prefix.Length)
            [void]$items.Add([PSCustomObject]@{
                Name = "$name (old location)"; Slug = (ConvertTo-MonitorSlug $name); Dir = (Join-Path $PrevRoot (ConvertTo-MonitorSlug $name)); Url = ''; Kind = 'TaskOnly'
                TaskName = $task.TaskName; TaskPath = $PrevPath; TaskState = "$($task.State)"
            })
        }
    }
    $legacy = Get-LegacyInstall -Dir $LegacyDir -TaskName $LegacyTask -Path $PrevPath
    if ($legacy) {
        $lt = Get-ScheduledTask -TaskName $LegacyTask -TaskPath $legacy.TaskPath -ErrorAction SilentlyContinue
        [void]$items.Add([PSCustomObject]@{
            Name = "$LegacyName (older layout)"; Slug = ''; Dir = $legacy.Dir; Url = $legacy.Url; Kind = 'Legacy'
            TaskName = $LegacyTask; TaskPath = $legacy.TaskPath; TaskState = $(if ($lt) { "$($lt.State)" } else { 'no task' })
        })
    }
    return @($items)
}

function Show-RemovableMonitor {
    # Prints the numbered listing the operator picks from
    param([object[]]$Items)
    $Items = @($Items | Where-Object { $null -ne $_ })
    $i = 0
    foreach ($m in @($Items)) {
        $i++
        Write-Host ("  {0}  {1}" -f $i, $m.Name)
        Write-Host ("     {0}" -f $(if ($m.Url) { $m.Url } else { 'URL not recorded' })) -ForegroundColor DarkGray
        Write-Host ("     task {0}, {1} in {2}" -f $m.TaskState, (Get-FolderSizeText $m.Dir), $m.Dir) -ForegroundColor DarkGray
    }
}

function Select-RemovableMonitor {
    # Resolves a typed name to one monitor. Exact spelling wins, then case-insensitive, then the folder slug.
    param([Parameter(Mandatory)][object[]]$Items, [Parameter(Mandatory)][string]$Requested)
    $list = @($Items | Where-Object { $null -ne $_ })
    $wanted = "$Requested".Trim()
    if (-not $wanted) { Write-Log -Level WARNING -Message "Type a number, a name, or X to cancel."; return $null }
    $match = @($list | Where-Object { $_.Name -ceq $wanted })
    if (@($match).Count -eq 0) { $match = @($list | Where-Object { $_.Name -eq $wanted }) }
    if (@($match).Count -eq 0) {
        $slug = ConvertTo-MonitorSlug $wanted
        if ($slug) { $match = @($list | Where-Object { $_.Slug -and $_.Slug -eq $slug }) }
    }
    if (@($match).Count -eq 1) { return @($match)[0] }
    if (@($match).Count -gt 1) {
        Write-Log -Level WARNING -Message "'$wanted' matches more than one monitor: $((@($match) | ForEach-Object { $_.Name }) -join ', '). Use its number instead."
        return $null
    }
    Write-Log -Level WARNING -Message "No monitor called '$wanted'. Installed: $((@($list) | ForEach-Object { $_.Name }) -join ', ')."
    return $null
}

function Stop-MonitorProcess {
    # A monitor process outliving its task keeps the folder locked, so it is stopped by the path it runs from
    param([Parameter(Mandatory)][string]$Dir)
    $pattern = "*" + [Management.Automation.WildcardPattern]::Escape($Dir.TrimEnd('\') + '\') + "*"
    $found = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like $pattern })
    $stopped = 0
    foreach ($p in $found) {
        try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; $stopped++ } catch { }
    }
    for ($i = 0; $i -lt 20 -and $stopped -gt 0; $i++) {
        Start-Sleep -Milliseconds 300
        $still = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like $pattern })
        if (@($still).Count -eq 0) { break }
    }
    if ($stopped) { Write-Log -Level INFORMATIONAL -Message "Stopped $stopped monitor process(es) still running from '$Dir'." }
    return $stopped
}

function Move-MonitorHistory {
    # Moves the measurement history out of the folder before the rest goes. Returns the destination, or $null.
    param([Parameter(Mandatory)]$Monitor, [string]$KeepRoot = $HistoryKeepRoot, [System.Collections.ArrayList]$Problems = $null)
    if (-not (Test-Path -LiteralPath $Monitor.Dir)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $Monitor.Dir -File -Force -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -like 'Latency_*.csv' -or $_.Name -like 'Outages*.csv' -or $_.Name -like 'Drops*.log' -or $_.Name -like 'HST-eChart-*'
    })
    if (@($files).Count -eq 0) { return $null }
    $slug = if ($Monitor.Slug) { $Monitor.Slug } else { ConvertTo-MonitorSlug $Monitor.Name }
    # A hand-edited settings file can name a folder that climbs out of the root, so only the leaf is ever used
    $slug = Split-Path -Path $slug -Leaf
    if (-not $slug -or $slug -eq '..' -or $slug -eq '.' -or $slug -match '[\\/:]') { $slug = 'monitor' }
    $dest = Join-Path $KeepRoot ("{0}_{1}" -f $slug, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try { if (-not (Test-Path -LiteralPath $dest)) { New-Item -Path $dest -ItemType Directory -Force -ErrorAction Stop | Out-Null } }
    catch {
        Write-Log -Level WARNING -Message "Could not create '$dest' for the kept history. $($_.Exception.Message)"
        if ($null -ne $Problems) { [void]$Problems.Add("the history could not be kept in '$dest': $($_.Exception.Message)") }
        return $null
    }
    $moved = 0
    foreach ($f in $files) {
        $target = Join-Path $dest $f.Name
        if (Test-Path -LiteralPath $target) { $target = Join-Path $dest ("{0}_{1}{2}" -f $f.BaseName, (Get-Date -Format 'HHmmssfff'), $f.Extension) }
        try { Move-Item -LiteralPath $f.FullName -Destination $target -Force -ErrorAction Stop; $moved++ }
        catch {
            Write-Log -Level WARNING -Message "Could not keep '$($f.Name)'. $($_.Exception.Message)"
            if ($null -ne $Problems) { [void]$Problems.Add("'$($f.Name)' could not be kept: $($_.Exception.Message)") }
        }
    }
    if ($moved -eq 0) {
        Remove-Item -LiteralPath $dest -Recurse -Force -ErrorAction SilentlyContinue
        return $null
    }
    Write-Log -Level CREATED -Message "Kept $moved history file(s) in '$dest'."
    return $dest
}

function Remove-MonitorInstall {
    # Stops the task and its process, keeps the history when asked, then deletes the folder. Never throws, and
    # never touches another monitor. Returns what was and was not removed.
    param([Parameter(Mandatory)]$Monitor, [bool]$KeepHistory = $true, [string]$KeepRoot = $HistoryKeepRoot)
    $problems = New-Object System.Collections.ArrayList
    $taskRemoved = $false
    $task = Get-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue
    if ($task) {
        try { Stop-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue } catch { }
        for ($i = 0; $i -lt 15; $i++) {
            $t = Get-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -ErrorAction SilentlyContinue
            if (-not $t -or "$($t.State)" -ne 'Running') { break }
            Start-Sleep -Seconds 1
        }
        try {
            Unregister-ScheduledTask -TaskName $Monitor.TaskName -TaskPath $Monitor.TaskPath -Confirm:$false -ErrorAction Stop
            $taskRemoved = $true
            Write-Log -Level INFORMATIONAL -Message "Removed the task '$($Monitor.TaskPath)$($Monitor.TaskName)'."
        }
        catch { [void]$problems.Add("the task '$($Monitor.TaskPath)$($Monitor.TaskName)' could not be removed: $($_.Exception.Message)") }
    }
    else {
        $taskRemoved = $true
        Write-Log -Level INFORMATIONAL -Message "No task '$($Monitor.TaskPath)$($Monitor.TaskName)' to remove."
    }
    if (Test-Path -LiteralPath $Monitor.Dir) { $null = Stop-MonitorProcess -Dir $Monitor.Dir }
    $historyPath = $null
    if ($KeepHistory) { $historyPath = Move-MonitorHistory -Monitor $Monitor -KeepRoot $KeepRoot -Problems $problems }
    $folderRemoved = $false
    if (Test-Path -LiteralPath $Monitor.Dir) {
        try { Remove-Item -LiteralPath $Monitor.Dir -Recurse -Force -ErrorAction Stop; $folderRemoved = $true }
        catch {
            Start-Sleep -Seconds 2
            try { Remove-Item -LiteralPath $Monitor.Dir -Recurse -Force -ErrorAction Stop; $folderRemoved = $true }
            catch { [void]$problems.Add("the folder '$($Monitor.Dir)' could not be deleted: $($_.Exception.Message)") }
        }
        if ($folderRemoved) { Write-Log -Level INFORMATIONAL -Message "Deleted '$($Monitor.Dir)'." }
    }
    else {
        $folderRemoved = $true
        Write-Log -Level INFORMATIONAL -Message "No folder '$($Monitor.Dir)' to delete."
    }
    return [PSCustomObject]@{
        Name = $Monitor.Name; TaskRemoved = $taskRemoved; FolderRemoved = $folderRemoved
        HistoryPath = $historyPath; Problems = @($problems)
    }
}

function Invoke-UninstallFlow {
    # Lists what is installed, removes the chosen monitor, and says exactly what went and what stayed.
    param(
        [string]$Requested = $MonitorNameOverride,
        [bool]$KeepHistory = $KeepHistoryOnUninstall,
        [string]$Root = $InstallRoot,
        [string]$Path = $TaskPath,
        [string]$KeepRoot = $HistoryKeepRoot,
        [string]$PrevRoot = $PreviousRoot,
        [string]$PrevPath = $PreviousTaskPath
    )
    $items = @(Get-RemovableMonitor -Root $Root -Path $Path -PrevRoot $PrevRoot -PrevPath $PrevPath | Where-Object { $null -ne $_ })
    if (@($items).Count -eq 0) {
        Write-Log -Level FOUND -Message "No monitor is installed under '$Root' and no monitor task exists under '$Path'. Nothing to remove."
        return 0
    }
    Write-Host ""
    Write-Host "Installed monitors"
    Show-RemovableMonitor -Items $items
    Write-Host ""
    $target = $null
    if ($Requested) {
        $target = Select-RemovableMonitor -Items $items -Requested $Requested
        if (-not $target) { Write-Log -Level FAILED -Message "Nothing was removed."; return 1 }
    }
    elseif ($NonInteractive) {
        if (@($items).Count -eq 1) { $target = @($items)[0] }
        else {
            Write-Log -Level FAILED -Message "$(@($items).Count) monitors are installed. Set `$MonitorNameOverride to the one to remove. Nothing was removed."
            return 1
        }
    }
    else {
        $default = if (@($items).Count -eq 1) { '1' } else { '' }
        while (-not $target) {
            Write-Question -Question "Which one?" -Hint "Its number or its name. X cancels."
            $typed = "$(Read-Setting -Prompt "Remove" -Default $default)".Trim()
            if (-not $typed) { Write-Log -Level WARNING -Message "Type a number, a name, or X to cancel."; continue }
            if ($typed -eq 'X' -or $typed -eq 'x') { Write-Log -Level FOUND -Message "Cancelled. Nothing was removed."; return 0 }
            $n = 0
            if ([int]::TryParse($typed, [ref]$n)) {
                if ($n -ge 1 -and $n -le @($items).Count) { $target = @($items)[$n - 1] }
                else { Write-Log -Level WARNING -Message "Pick a number between 1 and $(@($items).Count)." }
                continue
            }
            $target = Select-RemovableMonitor -Items $items -Requested $typed
        }
    }
    $others = @($items | Where-Object { $_.Name -cne $target.Name })
    Write-Host ""
    Write-Host "About to remove"
    Write-Host "  Monitor     : $($target.Name)"
    Write-Host "  Task        : $($target.TaskPath)$($target.TaskName) ($($target.TaskState))"
    Write-Host "  Folder      : $($target.Dir) ($(Get-FolderSizeText $target.Dir))"
    Write-Host "  Left alone  : $(if (@($others).Count) { (@($others) | ForEach-Object { $_.Name }) -join ', ' } else { 'nothing else is installed' })"
    if (-not $NonInteractive) {
        Write-Question -Question "Remove it?"
        if ((Read-Choice -Prompt "Remove" -Allowed @('Y','N') -Default 'N') -ne 'Y') {
            Write-Log -Level FOUND -Message "Cancelled. Nothing was removed."
            return 0
        }
        Write-Question -Question "Keep what it measured?" -Hint @(
            "Y moves the latency CSVs, the outages CSV, and the drops log to $KeepRoot.",
            "N deletes them with the folder, and that cannot be undone."
        )
        $KeepHistory = (Read-Choice -Prompt "Keep" -Allowed @('Y','N') -Default 'Y') -eq 'Y'
        if (-not $KeepHistory) { Write-Log -Level WARNING -Message "The history goes with the folder. That cannot be undone." }
    }
    $result = Remove-MonitorInstall -Monitor $target -KeepHistory $KeepHistory -KeepRoot $KeepRoot
    $left = @(Get-RemovableMonitor -Root $Root -Path $Path -PrevRoot $PrevRoot -PrevPath $PrevPath | Where-Object { $null -ne $_ })
    if (@($left).Count -eq 0 -and (Test-Path -LiteralPath $Root)) {
        if (@(Get-ChildItem -LiteralPath $Root -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            try { Remove-Item -LiteralPath $Root -Force -ErrorAction Stop; Write-Log -Level INFORMATIONAL -Message "Removed the empty install root '$Root'." } catch { }
        }
    }
    Write-Host ""
    Write-Host "Removed"
    Write-Host "  Monitor     : $($target.Name)"
    Write-Host "  Task        : $($target.TaskPath)$($target.TaskName) ($(if ($result.TaskRemoved) { 'removed' } else { 'still there' }))"
    Write-Host "  Folder      : $($target.Dir) ($(if ($result.FolderRemoved) { 'deleted' } else { 'still there' }))"
    Write-Host "  History     : $(if ($result.HistoryPath) { "kept in $($result.HistoryPath)" } elseif ($KeepHistory) { 'none to keep' } else { 'deleted with the folder' })"
    Write-Host "  Left here   : $(if (@($left).Count) { (@($left) | ForEach-Object { $_.Name }) -join ', ' } else { 'nothing' })"
    Write-Host ""
    if (@($left).Count -eq 0) {
        $telemetryPath = @($Path, $PrevPath) | Where-Object { $_ } | Where-Object { Get-ScheduledTask -TaskName $TelemetryTaskName -TaskPath $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
        if ($telemetryPath) {
            Write-Log -Level WARNING -Message "The telemetry publisher task '$telemetryPath$TelemetryTaskName' is still scheduled and has nothing left to publish. Remove it with: Unregister-ScheduledTask -TaskPath '$telemetryPath' -TaskName '$TelemetryTaskName' -Confirm:`$false"
        }
    }
    if (@($result.Problems).Count) {
        foreach ($p in @($result.Problems)) { Write-Log -Level FAILED -Message "Not fully removed: $p" }
        Write-Log -Level FAILED -Message "Finish by hand, then run the uninstall again to confirm."
        return 1
    }
    Write-Log -Level FINISHED -Message "Removed '$($target.Name)'."
    return 0
}

function New-MonitorContent {
    # Builds the monitor script text from the embedded template with values baked in. One regex pass over the
    # template means a value that happens to contain a token is inserted as-is, never substituted again.
    param([Parameter(Mandatory)][string]$SiteName, [Parameter(Mandatory)][hashtable]$Mail)

    $SiteName = ConvertTo-SafeSiteName $SiteName
    if ([string]::IsNullOrWhiteSpace($SiteName)) { $SiteName = ConvertTo-SafeSiteName $env:COMPUTERNAME }
    $q = { param($s) if ($null -eq $s) { '' } else { ([string]$s).Replace("'", "''") } }
    $b = { param($v) if ($v) { '$true' } else { '$false' } }
    $toLiteral = "@(" + (($Mail.MailTo | ForEach-Object { "'" + (& $q $_) + "'" }) -join ",") + ")"

    $values = @{
        SITENAME        = (& $q $SiteName)
        MONITORNAME     = (& $q $MonitorName)
        URL             = (& $q $Url)
        INSTALLDIR      = (& $q $InstallDir)
        CREDFILE        = (& $q $CredentialFileName)
        INTERVAL        = [string][math]::Max(1, [int]$IntervalSeconds)
        TIMEOUT         = [string][math]::Max(1, [int]$TimeoutSeconds)
        MAXREDIRS       = [string][math]::Max(0, [int]$MaxRedirects)
        MARKER          = (& $q $ExpectedContentMarker)
        MINBYTES        = [string][math]::Max(1, [int]$MinPopulatedBytes)
        MAXBODYBYTES    = [string][math]::Max([math]::Max(1048576, 4 * [int]$MinPopulatedBytes), [int64]$MaxBodyBytes)
        SLOWMS          = [string][math]::Max(0, [int]$SlowThresholdMs)
        ALERTONSLOW     = (& $b $AlertOnSlow)
        SLOWWINDOW      = [string][math]::Max(1, [int]$SlowWindowMinutes)
        SLOWALERTPCT    = [string][math]::Min(100, [math]::Max(1, [int]$SlowAlertPercent))
        SLOWCLEARPCT    = [string][math]::Max(0, [math]::Min([int]$SlowClearPercent, [math]::Min(100, [math]::Max(1, [int]$SlowAlertPercent)) - 1))
        SUMMARYHOUR     = [string][math]::Min(23, [math]::Max(-1, [int]$DailySummaryHour))
        DOWNTHRESHOLD   = [string][math]::Max(1, [int]$DownThreshold)
        REALERT         = [string][math]::Max(0, [int]$ReAlertMinutes)
        RETENTIONDAYS   = [string][math]::Max(1, [int]$LogRetentionDays)
        ALERTONRECOVERY = (& $b $AlertOnRecovery)
        SENDEMAIL       = (& $b $AlertsEnabled)
        MAILMETHOD      = (& $q $Mail.MailMethod)
        SMTPSERVER      = (& $q $Mail.SmtpServer)
        SMTPPORT        = [string][int]$Mail.SmtpPort
        SMTPUSESSL      = (& $b $Mail.SmtpUseSsl)
        MAILFROM        = (& $q $Mail.MailFrom)
        MAILTO          = $toLiteral
        SMTPAUTHUSER    = (& $q $Mail.SmtpAuthUser)
        GRAPHTENANT     = (& $q $Mail.GraphTenantId)
        GRAPHCLIENT     = (& $q $Mail.GraphClientId)
        SECRETEXPIRES   = (& $q $Mail.GraphSecretExpires)
    }
    $evaluator = { param($m) $key = $m.Groups[1].Value; if ($values.ContainsKey($key)) { $values[$key] } else { $m.Value } }
    return [regex]::Replace($Template_MonitorScript, '@@([A-Z]+)@@', $evaluator)
}

function Test-EndpointReachable {
    # One-shot reachability check used at install time. Follows redirects exactly as the monitor does.
    $format = 'CODE=%{http_code}\nREDIRECTS=%{num_redirects}\nFINAL=%{url_effective}'
    $lines = @(& curl.exe -s -L --max-redirs $MaxRedirects -o NUL -A "CurlMonitor/1.0" -H "Cache-Control: no-cache" -w $format --max-time $TimeoutSeconds $Url 2>$null)
    $code = $null; $redirects = $null; $final = $null
    foreach ($line in $lines) {
        if ($line -match '^CODE=(\d{3})$') { $code = $Matches[1] }
        elseif ($line -match '^REDIRECTS=(\d+)$') { $redirects = [int]$Matches[1] }
        elseif ($line -match '^FINAL=(.+)$') { $final = $Matches[1] }
    }
    [PSCustomObject]@{ HttpCode = $code; Redirects = $redirects; FinalUrl = $final; Reachable = ($code -eq '200') }
}

function Wait-TaskStopped {
    # Waits up to 15 seconds for the task to leave the Running state
    for ($i = 0; $i -lt 15; $i++) {
        $t = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
        if (-not $t -or $t.State -ne 'Running') { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}

$Template_MonitorScript = @'
<#
.SYNOPSIS
    Continuous latency, availability, and timeout monitor for one URL.

.DESCRIPTION
    Deployed by Install-CurlMonitor.ps1. Runs continuously as SYSTEM, one instance per monitored URL.
    1. Validates curl.exe and the output folder, starts a daily transcript log, and prunes old logs.
    2. Probes the endpoint every interval, following redirects to the sign-in page, and captures DNS, connect,
       TLS, TTFB, redirect, and total timings plus HTTP code, redirect count, final URL, size, answering backend
       IP, and the curl exit code with a plain-English reason.
    3. Confirms the sign-in page populated: the content marker is present and the body meets the minimum size.
    4. Appends one timestamped row (local and UTC) to a monthly CSV on every poll, and writes failed or slow polls,
       outage transitions, alert delivery problems, and starts and restarts to Drops.log, never a healthy poll.
    5. Tracks up and down state with true outage onset, alerts on state change with the site in the subject,
       re-alerts while still down, and on recovery emails the outage duration and writes an outage record.
       Short of an outage, it emails SLOW when most polls over a rolling window are slow or failing, re-alerts
       while still slow, emails SLOW RESOLVED when responses are back to normal, and once a day sends a
       summary of slow and failed polls when there were any.
    6. Writes a heartbeat every poll. On start it reports how long it was not running, whether the server
       rebooted, and carries over any outage that was in progress. Alerts that cannot be sent are retried
       once a minute for an hour, and a recovery email says when the original DOWN alert never got out.

.REQUIREMENTS
    Windows 10 1803+ or Windows Server 2019+ (ships with curl.exe).

.OUTPUTS
    Monthly latency CSV, outages CSV, drops log, daily transcript logs, state-change email alerts, and a daily summary.

.NOTES
    Author:      Christopher Carroll
    Created:     09/04/2026
    Idempotency: Safe to run repeatedly. Each poll appends one row and never modifies prior rows.
    Context:     Generated file. Edit the template in Install-CurlMonitor.ps1 and re-run the installer, do not
                 hand-edit this deployed copy. Runs headless, so it never auto-opens any file.

.LINK
    https://curl.se/docs/manpage.html
#>

Remove-Variable * -ErrorAction SilentlyContinue

# What this monitor watches
$MonitorName           = '@@MONITORNAME@@'
$SiteName              = '@@SITENAME@@'
$Url                   = '@@URL@@'

# Output locations
$InstallDir            = '@@INSTALLDIR@@'
$OutageCsv             = '@@INSTALLDIR@@\Outages.csv'
$HeartbeatFile         = '@@INSTALLDIR@@\heartbeat.json'
$DropLog               = '@@INSTALLDIR@@\Drops.log'
$SummaryStateFile      = '@@INSTALLDIR@@\summary-sent.txt'
$LogRetentionDays      = @@RETENTIONDAYS@@

# Probe behavior
$IntervalSeconds       = @@INTERVAL@@
$TimeoutSeconds        = @@TIMEOUT@@
$MaxRedirects          = @@MAXREDIRS@@
$ExpectedContentMarker = '@@MARKER@@'
$MinPopulatedBytes     = @@MINBYTES@@
$MaxBodyBytes          = @@MAXBODYBYTES@@
$SlowThresholdMs       = @@SLOWMS@@

# Alerting behavior
$SendEmail             = @@SENDEMAIL@@
$DownThreshold         = @@DOWNTHRESHOLD@@
$ReAlertMinutes        = @@REALERT@@
$AlertOnRecovery       = @@ALERTONRECOVERY@@
$AlertOnSlow           = @@ALERTONSLOW@@
$SlowWindowMinutes     = @@SLOWWINDOW@@
$SlowAlertPercent      = @@SLOWALERTPCT@@
$SlowClearPercent      = @@SLOWCLEARPCT@@
$DailySummaryHour      = @@SUMMARYHOUR@@

# Mail
$MailMethod            = '@@MAILMETHOD@@'
$SmtpServer            = '@@SMTPSERVER@@'
$SmtpPort              = @@SMTPPORT@@
$SmtpUseSsl            = @@SMTPUSESSL@@
$MailFrom              = '@@MAILFROM@@'
$MailTo                = @@MAILTO@@
$SmtpAuthUser          = '@@SMTPAUTHUSER@@'
$SmtpCredentialFile    = '@@INSTALLDIR@@\@@CREDFILE@@'
$GraphTenantId         = '@@GRAPHTENANT@@'
$GraphClientId         = '@@GRAPHCLIENT@@'
$GraphSecretExpires    = '@@SECRETEXPIRES@@'

$HashTable_PowerShellTranscriptLogsbyKeyword = [Ordered]@{
    ADDED          = "ADDED |"
    CREATED        = "CREATED |"
    ERROR          = "ERROR |"
    FAILED         = "FAILED |"
    FINISHED       = "FINISHED |"
    FOUND          = "FOUND |"
    INFORMATIONAL  = "INFORMATIONAL |"
    PROMPT         = "PROMPT |"
    "SANITY CHECK" = "SANITY CHECK |"
    STARTED        = "STARTED |"
    SUCCESS        = "SUCCESS |"
    TOTAL          = "TOTAL |"
    WARNING        = "WARNING |"
}

function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet('ADDED','CREATED','ERROR','FAILED','FINISHED','FOUND','INFORMATIONAL','PROMPT','SANITY CHECK','STARTED','SUCCESS','TOTAL','WARNING')][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    $prefix = $HashTable_PowerShellTranscriptLogsbyKeyword[$Level]
    Write-Host ("{0} {1} {2}" -f (Get-Date -Format "MM/dd/yy - hh:mm:ss tt"), $prefix, $Message)
}

function ConvertTo-Ms {
    # Converts a curl timing value in seconds to whole milliseconds, culture-invariant
    param([string]$Seconds)
    if ([string]::IsNullOrWhiteSpace($Seconds)) { return $null }
    $d = 0.0
    if (-not [double]::TryParse($Seconds, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $null }
    return [math]::Round($d * 1000, 0)
}

function Format-Duration {
    # Formats a timespan as a compact human string, for example "1h 04m 30s"
    param([TimeSpan]$Span)
    if ($Span.Ticks -lt 0) { $Span = [TimeSpan]::Zero }
    $parts = @()
    if ($Span.Days -gt 0)  { $parts += ("{0}d" -f $Span.Days) }
    if ($Span.Hours -gt 0) { $parts += ("{0}h" -f $Span.Hours) }
    $parts += ("{0:00}m" -f $Span.Minutes)
    $parts += ("{0:00}s" -f $Span.Seconds)
    return ($parts -join ' ')
}

function Get-CurlReason {
    # Maps common curl exit codes to a short reason for the CSV and alert body
    param([int]$ExitCode)
    switch ($ExitCode) {
        0  { return 'OK' }
        6  { return 'DNS resolution failed' }
        7  { return 'Connection refused or unreachable' }
        18 { return 'Transfer ended early (partial body)' }
        28 { return 'Timed out' }
        35 { return 'TLS handshake failed' }
        47 { return 'Too many redirects' }
        52 { return 'Empty reply from server' }
        63 { return 'Response larger than the size limit' }
        55 { return 'Send failed' }
        60 { return 'TLS certificate not trusted' }
        56 { return 'Connection reset during receive' }
        default { return "curl exit $ExitCode" }
    }
}

function Get-LatencyCsvPath {
    # Monthly latency file so no single CSV grows without bound
    return (Join-Path $InstallDir ("Latency_" + (Get-Date -Format 'yyyyMM') + ".csv"))
}

function Write-CsvRow {
    # Appends one row. A file whose header no longer matches the row is moved aside and a fresh file started.
    # Any other write problem (for example the file open in Excel) is logged and that row skipped. Never throws.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Row)
    try {
        $folder = Split-Path -Path $Path -Parent
        if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
        $expected = ($Row.PSObject.Properties.Name | ForEach-Object { '"' + $_ + '"' }) -join ','
        if (Test-Path $Path) {
            $header = $null
            try { $header = [string](Get-Content -Path $Path -TotalCount 1 -ErrorAction Stop) } catch { }
            if ($header -and $header -ne $expected) {
                $aside = $Path -replace '\.csv$', ('_schema-' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
                Move-Item -Path $Path -Destination $aside -Force -ErrorAction Stop
                Write-Log -Level WARNING -Message "CSV header changed on '$Path'. Old file moved to '$aside' and a new file started."
            }
        }
        $Row | Export-Csv -Path $Path -NoTypeInformation -Append -ErrorAction Stop
    }
    catch {
        Write-Log -Level WARNING -Message "Could not append to '$Path' (open in another program?). This row was not recorded. $($_.Exception.Message)"
    }
}

function Get-TranscriptPath {
    return (Join-Path $InstallDir ("Transcript_" + (Get-Date -Format 'yyyyMMdd') + ".log"))
}

function Update-TranscriptIfDayChanged {
    # Rolls the transcript daily and prunes transcripts older than the retention window. CSVs and the drops log are kept.
    param([Parameter(Mandatory)][string]$CurrentPath)
    $wanted = Get-TranscriptPath
    if ($wanted -eq $CurrentPath) { return $CurrentPath }
    try { Stop-Transcript | Out-Null } catch { }
    Start-Transcript -Path $wanted -Append | Out-Null
    Write-Log -Level INFORMATIONAL -Message "Rolled transcript to '$wanted'."
    $cutoff = (Get-Date).AddDays(-$LogRetentionDays)
    Get-ChildItem -Path $InstallDir -Filter 'Transcript_*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Force -ErrorAction SilentlyContinue
    return $wanted
}

function Get-ProtectedSecret {
    # Reads and decrypts the machine-scope DPAPI secret file. Returns $null when missing or unreadable.
    if (-not (Test-Path $SmtpCredentialFile)) {
        Write-Log -Level WARNING -Message "Mail secret file is missing. Re-run the installer."
        return $null
    }
    try {
        Add-Type -AssemblyName System.Security
        $cipher  = [Convert]::FromBase64String((Get-Content -Path $SmtpCredentialFile -Raw).Trim())
        $entropy = [System.Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1')
        return [System.Text.Encoding]::UTF8.GetString([System.Security.Cryptography.ProtectedData]::Unprotect($cipher, $entropy, [System.Security.Cryptography.DataProtectionScope]::LocalMachine))
    }
    catch {
        Write-Log -Level WARNING -Message "Could not decrypt the mail secret. Re-run the installer. $($_.Exception.Message)"
        return $null
    }
}

function ConvertTo-SecureText {
    # Builds a read-only SecureString from text that is already in memory, for cmdlets that accept nothing else
    param([Parameter(Mandatory)][string]$Text)
    $secure = New-Object System.Security.SecureString
    foreach ($ch in $Text.ToCharArray()) { $secure.AppendChar($ch) }
    $secure.MakeReadOnly()
    return $secure
}

function Get-SmtpCredential {
    # Rebuilds the PSCredential for authenticated SMTP. Returns $null when not using auth.
    if ([string]::IsNullOrWhiteSpace($SmtpAuthUser)) { return $null }
    $plain = Get-ProtectedSecret
    if ($null -eq $plain) { return $null }
    if ($plain.Length -eq 0) { Write-Log -Level WARNING -Message "The stored SMTP password is empty. Re-run the installer."; return $null }
    return (New-Object System.Management.Automation.PSCredential($SmtpAuthUser, (ConvertTo-SecureText -Text $plain)))
}

function New-AlertBody {
    # Builds a compact HTML body: a heading, a two-column table of details, and a footer naming the sending server
    param([Parameter(Mandatory)][string]$Heading, [Parameter(Mandatory)][System.Collections.IDictionary]$Details, [string]$Footer = "Sent by the curl monitor on $env:COMPUTERNAME.")
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
    $rows = foreach ($k in $Details.Keys) { "<tr><td style='padding:3px 16px 3px 0;color:#555;white-space:nowrap;vertical-align:top'>$(& $enc $k)</td><td style='padding:3px 0'>$(& $enc $Details[$k])</td></tr>" }
    return "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'><p style='font-size:16px;font-weight:600;margin:0 0 10px'>$(& $enc $Heading)</p><table style='border-collapse:collapse'>$($rows -join '')</table><p style='margin:14px 0 0;color:#777;font-size:12px'>$(& $enc $Footer)</p></body></html>"
}

function Get-RestErrorDetail {
    # Pulls the service's own error code and message out of a web exception, so failures name their real cause
    param($ErrorRecord)
    $detail = ""
    try { $detail = $ErrorRecord.ErrorDetails.Message } catch { }
    if (-not $detail) { try { $stream = $ErrorRecord.Exception.Response.GetResponseStream(); $detail = (New-Object System.IO.StreamReader($stream)).ReadToEnd() } catch { } }
    if (-not $detail) { return "" }
    if ($detail.Length -gt 1000) { $detail = $detail.Substring(0, 1000) }
    try { $j = $detail | ConvertFrom-Json; if ($j.error.code -or $j.error.message) { return ("[{0}] {1}" -f $j.error.code, $j.error.message) } } catch { }
    return $detail
}

function Get-GraphToken {
    # Returns a cached client-credentials token, refreshing when within 5 minutes of expiry
    if ($script:GraphToken -and $script:GraphTokenExpiresUtc -and ((Get-Date).ToUniversalTime() -lt $script:GraphTokenExpiresUtc.AddMinutes(-5))) { return $script:GraphToken }
    $secret = Get-ProtectedSecret
    if ($null -eq $secret) { return $null }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $resp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$GraphTenantId/oauth2/v2.0/token" -Body @{ client_id = $GraphClientId; client_secret = $secret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30 -ErrorAction Stop
    $script:GraphToken = $resp.access_token
    $script:GraphTokenExpiresUtc = (Get-Date).ToUniversalTime().AddSeconds([int]$resp.expires_in)
    return $script:GraphToken
}

function ConvertTo-GraphMailBody {
    # Builds the JSON body for POST /users/{sender}/sendMail
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string[]]$To)
    $rcpts = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    return (@{ message = @{ subject = $Subject; body = @{ contentType = 'HTML'; content = $Body }; toRecipients = $rcpts }; saveToSentItems = $false } | ConvertTo-Json -Depth 6)
}

function Get-SecretExpiryWarning {
    # Pure. Returns a warning when the secret expires within 30 days or has expired, else $null.
    param([string]$ExpiresOn, [datetime]$Today)
    if ([string]::IsNullOrWhiteSpace($ExpiresOn)) { return $null }
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($ExpiresOn, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $null }
    $days = [int]($d.Date - $Today.Date).TotalDays
    if ($days -lt 0)  { return "Graph client secret EXPIRED on $ExpiresOn. Alerts are not being delivered. Re-run Install-CurlMonitor.ps1 and choose N at the app registration prompt to mint a new secret." }
    if ($days -le 30) { return "Graph client secret expires in $days day(s) on $ExpiresOn. Re-run Install-CurlMonitor.ps1 and choose N at the app registration prompt to mint a new secret." }
    return $null
}

function Get-ProbeResult {
    # Runs one timed request, following redirects to the sign-in page the way a browser would. stderr is discarded so parsing stays clean.
    # The body lands in the install folder, not the global temp folder, and is removed after the size and marker checks.
    $tempBody = Join-Path $InstallDir 'probe-body.tmp'
    $format = 'DNS=%{time_namelookup}\nCONNECT=%{time_connect}\nTLS=%{time_appconnect}\nTTFB=%{time_starttransfer}\nTOTAL=%{time_total}\nREDIRTIME=%{time_redirect}\nREDIRECTS=%{num_redirects}\nCODE=%{http_code}\nSIZE=%{size_download}\nIP=%{remote_ip}\nFINAL=%{url_effective}'
    $lines = @()
    $exit = -1
    try {
        $lines = @(& curl.exe -s -L --max-redirs $MaxRedirects -A "CurlMonitor/1.0" -H "Cache-Control: no-cache" -o $tempBody -w $format --max-time $TimeoutSeconds --max-filesize $MaxBodyBytes $Url 2>$null)
        $exit = $LASTEXITCODE
    }
    catch {
        $lines = @()
    }

    $parsed = @{}
    foreach ($line in $lines) {
        $kv = ([string]$line) -split '=', 2
        if ($kv.Count -eq 2) { $parsed[$kv[0].Trim()] = $kv[1].Trim() }
    }

    # Populated means the final body is at least the minimum size and, when a marker is configured, contains it
    $sizeOk = $false
    $markerOk = $true
    try {
        if (Test-Path $tempBody) {
            $bodyBytes = (Get-Item $tempBody).Length
            $sizeOk = ($bodyBytes -ge $MinPopulatedBytes)
            if (-not [string]::IsNullOrWhiteSpace($ExpectedContentMarker)) {
                # An oversized body is never loaded: curl stops at MaxBodyBytes, and a file past it is not scanned
                if ($bodyBytes -gt $MaxBodyBytes) { $markerOk = $false }
                else { $markerOk = [bool](Select-String -Path $tempBody -SimpleMatch -Pattern $ExpectedContentMarker -Quiet) }
            }
        }
    }
    finally {
        Remove-Item $tempBody -ErrorAction SilentlyContinue
    }

    $code = $parsed['CODE']
    if ([string]::IsNullOrWhiteSpace($code)) { $code = '000' }
    $reason = Get-CurlReason $exit
    if ($exit -eq 0 -and $code -eq '200' -and -not ($sizeOk -and $markerOk)) {
        $reason = "Page not populated ($($parsed['SIZE']) bytes, marker " + $(if ($markerOk) { 'found' } else { 'missing' }) + ")"
    }
    elseif ($exit -eq 0 -and $code -ne '200') { $reason = "HTTP $code" }

    [PSCustomObject]@{
        Timestamp_Local = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Timestamp_UTC   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        SiteName        = $SiteName
        Url             = $Url
        HttpCode        = $code
        CurlExit        = $exit
        Reason          = $reason
        DnsMs           = ConvertTo-Ms $parsed['DNS']
        ConnectMs       = ConvertTo-Ms $parsed['CONNECT']
        TlsMs           = ConvertTo-Ms $parsed['TLS']
        TtfbMs          = ConvertTo-Ms $parsed['TTFB']
        TotalMs         = ConvertTo-Ms $parsed['TOTAL']
        SizeBytes       = $parsed['SIZE']
        RemoteIp        = $parsed['IP']
        Redirects       = $parsed['REDIRECTS']
        RedirectMs      = ConvertTo-Ms $parsed['REDIRTIME']
        FinalUrl        = $parsed['FINAL']
        ContentOk       = ($sizeOk -and $markerOk)
    }
}

function Get-AlertLabel {
    # "{Monitor} at {Site} ({HOST})" for subjects, dropping any part that is not set
    param([string]$MonitorName, [string]$SiteName, [string]$HostName)
    $label = if ($MonitorName -and $SiteName) { "$MonitorName at $SiteName" } elseif ($MonitorName) { $MonitorName } else { $SiteName }
    if ($HostName) { return "$label ($HostName)" }
    return $label
}

function Update-MonitorState {
    # Pure state transition. No side effects. Returns new state plus any email, outage record, and transition log.
    param(
        [Parameter(Mandatory)][hashtable]$State,
        [Parameter(Mandatory)][bool]$Failed,
        [Parameter(Mandatory)][object]$Result,
        [Parameter(Mandatory)][datetime]$NowUtc,
        [int]$DownThreshold,
        [int]$ReAlertMinutes,
        [bool]$AlertOnRecovery,
        [string]$SiteName,
        [string]$Url,
        [string]$HostName,
        [string]$MonitorName
    )

    if ($DownThreshold -lt 1) { $DownThreshold = 1 }
    $siteLabel = Get-AlertLabel -MonitorName $MonitorName -SiteName $SiteName -HostName $HostName
    $what = if ($MonitorName) { $MonitorName } else { 'The endpoint' }

    $s = @{
        ConsecutiveFailures = [int]$State.ConsecutiveFailures
        IsDown              = [bool]$State.IsDown
        LastAlertUtc        = $State.LastAlertUtc
        OutageStartUtc      = $State.OutageStartUtc
        OutageStartLocalStr = $State.OutageStartLocalStr
        OutageStartUtcStr   = $State.OutageStartUtcStr
        AlertDelivered      = $State.AlertDelivered
    }
    # A clock that went backwards leaves these in the future, which would write a negative outage duration and
    # hold off the reminders. Update-SlowState clamps the same way.
    if ($s.OutageStartUtc -and $s.OutageStartUtc -gt $NowUtc) { $s.OutageStartUtc = $NowUtc }
    if ($s.LastAlertUtc -and $s.LastAlertUtc -gt $NowUtc) { $s.LastAlertUtc = $NowUtc }
    $emailSubject = $null; $emailBody = $null; $outageRecord = $null; $transitionLog = $null; $emailKind = $null
    $deliveryNote = if ($s.AlertDelivered) { 'Delivered' } else { 'Not delivered. This server had no working mail path earlier in the outage, so this is the first notice.' }
    $reason = if ($Result.PSObject.Properties['Reason']) { $Result.Reason } else { '' }

    if ($Failed) {
        if ($s.ConsecutiveFailures -eq 0 -or $null -eq $s.OutageStartUtc) {
            $s.OutageStartUtc      = $NowUtc
            $s.OutageStartLocalStr = $Result.Timestamp_Local
            $s.OutageStartUtcStr   = $Result.Timestamp_UTC
        }
        $s.ConsecutiveFailures++

        if ((-not $s.IsDown) -and ($s.ConsecutiveFailures -ge $DownThreshold)) {
            $s.IsDown         = $true
            $s.LastAlertUtc   = $NowUtc
            $s.AlertDelivered = $false
            $emailKind        = 'Down'
            $emailSubject     = "[DOWN] $siteLabel - unreachable"
            $emailBody      = New-AlertBody -Heading "$what is unreachable from $SiteName" -Details ([ordered]@{ 'Status' = 'DOWN'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Outage started' = "$($s.OutageStartLocalStr) local ($($s.OutageStartUtcStr) UTC)"; 'Failed polls' = $s.ConsecutiveFailures; 'Last HTTP code' = $Result.HttpCode; 'Reason' = $reason; 'Page populated' = $Result.ContentOk; 'Backend IP' = $Result.RemoteIp })
            $transitionLog  = "Declared DOWN for $SiteName after $($s.ConsecutiveFailures) consecutive failures ($reason)."
        }
        elseif ($s.IsDown -and ($ReAlertMinutes -gt 0) -and ($null -ne $s.LastAlertUtc) -and (($NowUtc - $s.LastAlertUtc).TotalMinutes -ge $ReAlertMinutes)) {
            $s.LastAlertUtc = $NowUtc
            $elapsed        = Format-Duration ($NowUtc - $s.OutageStartUtc)
            $emailKind      = 'Reminder'
            $emailSubject   = "[STILL DOWN] $siteLabel - down for $elapsed"
            $emailBody      = New-AlertBody -Heading "$what is still unreachable from $SiteName" -Details ([ordered]@{ 'Status' = 'STILL DOWN'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Outage started' = "$($s.OutageStartLocalStr) local ($($s.OutageStartUtcStr) UTC)"; 'Down for' = $elapsed; 'Failed polls' = $s.ConsecutiveFailures; 'Last HTTP code' = $Result.HttpCode; 'Reason' = $reason; 'Backend IP' = $Result.RemoteIp; 'Earlier alerts' = $deliveryNote })
            $transitionLog  = "Reminder raised for $SiteName, down for $elapsed."
        }
    }
    else {
        if ($s.IsDown) {
            $span            = $NowUtc - $s.OutageStartUtc
            $duration        = Format-Duration $span
            $durationSeconds = [int][math]::Floor($span.TotalSeconds)

            $outageRecord = [PSCustomObject]@{
                SiteName          = $SiteName
                OutageStart_Local = $s.OutageStartLocalStr
                OutageStart_UTC   = $s.OutageStartUtcStr
                OutageEnd_Local   = $Result.Timestamp_Local
                OutageEnd_UTC     = $Result.Timestamp_UTC
                DurationSeconds   = $durationSeconds
                Duration          = $duration
                FailedPolls       = $s.ConsecutiveFailures
                RecoveryCode      = $Result.HttpCode
                RecoveryIp        = $Result.RemoteIp
            }
            if ($AlertOnRecovery) {
                $emailKind    = 'Resolved'
                $emailSubject = "[RESOLVED] $siteLabel - outage lasted $duration"
                $emailBody    = New-AlertBody -Heading "$what is reachable again from $SiteName" -Details ([ordered]@{ 'Status' = 'RESOLVED'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Outage started' = "$($s.OutageStartLocalStr) local ($($s.OutageStartUtcStr) UTC)"; 'Outage ended' = "$($Result.Timestamp_Local) local ($($Result.Timestamp_UTC) UTC)"; 'Duration' = $duration; 'Failed polls' = $s.ConsecutiveFailures; 'Recovery HTTP code' = $Result.HttpCode; 'Backend IP' = $Result.RemoteIp; 'DOWN alert' = $deliveryNote })
            }
            $transitionLog = "Outage record written: $SiteName lasted $duration over $($s.ConsecutiveFailures) failed polls."
            $s.IsDown = $false
        }
        $s.ConsecutiveFailures = 0
        $s.OutageStartUtc      = $null
        $s.OutageStartLocalStr = $null
        $s.OutageStartUtcStr   = $null
        $s.LastAlertUtc        = $null
        $s.AlertDelivered      = $null
    }

    [PSCustomObject]@{ State = $s; EmailSubject = $emailSubject; EmailBody = $emailBody; EmailKind = $emailKind; OutageRecord = $outageRecord; TransitionLog = $transitionLog }
}

function Update-SlowState {
    # Pure. Tracks slow or failed polls over a rolling window and decides the SLOW, STILL SLOW, and SLOW RESOLVED
    # emails. SLOW needs the window at least half covered, three bad polls, and the alert share. The period ends when
    # the share drops to the clear share. An outage owns its polls: while DOWN the window is emptied and a slow period
    # in progress is closed without an email, so the outage emails tell that story.
    param(
        [Parameter(Mandatory)][hashtable]$State,
        [Parameter(Mandatory)][object]$Result,
        [Parameter(Mandatory)][bool]$Failed,
        [Parameter(Mandatory)][bool]$IsDown,
        [Parameter(Mandatory)][datetime]$NowUtc,
        [int]$SlowThresholdMs = 3000,
        [int]$WindowMinutes = 5,
        [int]$AlertPercent = 50,
        [int]$ClearPercent = 10,
        [int]$ReAlertMinutes = 30,
        [bool]$AlertOnRecovery = $true,
        [string]$SiteName,
        [string]$Url,
        [string]$HostName,
        [string]$MonitorName
    )
    $WindowMinutes = [math]::Max(1, $WindowMinutes)
    $siteLabel = Get-AlertLabel -MonitorName $MonitorName -SiteName $SiteName -HostName $HostName
    $what = if ($MonitorName) { $MonitorName } else { 'The endpoint' }
    $s = @{
        Samples        = New-Object System.Collections.ArrayList
        IsSlow         = [bool]$State.IsSlow
        StartUtc       = $State.StartUtc
        StartLocalStr  = $State.StartLocalStr
        LastAlertUtc   = $State.LastAlertUtc
        Polls          = [int]$State.Polls
        SlowPolls      = [int]$State.SlowPolls
        FailedPolls    = [int]$State.FailedPolls
        WorstMs        = [int]$State.WorstMs
        AlertDelivered = $State.AlertDelivered
    }
    $emailKind = $null; $emailSubject = $null; $emailBody = $null; $transitionLog = $null; $dropKind = $null
    $windowStart = $NowUtc.AddMinutes(-$WindowMinutes)
    # A clock that steps back leaves samples and timestamps in the future. Drop them and pull the period back to now,
    # or the window never reads as covered again and the period can neither remind nor clear.
    foreach ($x in @($State.Samples)) { if ($x -and $x.Utc -gt $windowStart -and $x.Utc -le $NowUtc) { [void]$s.Samples.Add($x) } }
    if ($s.StartUtc -and $s.StartUtc -gt $NowUtc) { $s.StartUtc = $NowUtc }
    if ($s.LastAlertUtc -and $s.LastAlertUtc -gt $NowUtc) { $s.LastAlertUtc = $NowUtc }
    $closePeriod = { $s.IsSlow = $false; $s.StartUtc = $null; $s.StartLocalStr = $null; $s.LastAlertUtc = $null; $s.Polls = 0; $s.SlowPolls = 0; $s.FailedPolls = 0; $s.WorstMs = 0; $s.AlertDelivered = $null }

    if ($IsDown) {
        if ($s.IsSlow) {
            $transitionLog = "Slow period for $SiteName since $($s.StartLocalStr) local ended in an outage."
            $dropKind = 'SLOWCLEAR'
        }
        & $closePeriod
        $s.Samples.Clear()
        return [PSCustomObject]@{ State = $s; EmailKind = $null; EmailSubject = $null; EmailBody = $null; TransitionLog = $transitionLog; DropKind = $dropKind }
    }

    $ms = if ($null -ne $Result.TotalMs -and "$($Result.TotalMs)" -ne '') { [int]$Result.TotalMs } else { $null }
    $slowPoll = (-not $Failed) -and ($null -ne $ms) -and ($ms -ge $SlowThresholdMs)
    [void]$s.Samples.Add([PSCustomObject]@{ Utc = $NowUtc; LocalStr = [string]$Result.Timestamp_Local; Slow = $slowPoll; Failed = $Failed; Ms = $ms })

    $total = $s.Samples.Count
    $bad = @($s.Samples | Where-Object { $_.Slow -or $_.Failed })
    $share = [int][math]::Floor(100 * $bad.Count / $total)
    $okMs = @($s.Samples | Where-Object { -not $_.Failed -and $null -ne $_.Ms } | ForEach-Object { [int]$_.Ms } | Sort-Object)
    $timing = if ($okMs.Count) { "median $($okMs[[int][math]::Floor(($okMs.Count - 1) / 2)]) ms, worst $($okMs[-1]) ms" } else { 'no successful polls' }
    $slowCount = @($s.Samples | Where-Object { $_.Slow }).Count
    $failedCount = @($s.Samples | Where-Object { $_.Failed }).Count
    $windowText = "$total polls in $WindowMinutes min: $slowCount slower than $SlowThresholdMs ms, $failedCount failed"
    $lastPoll = "HTTP $($Result.HttpCode), $(if ($null -ne $ms) { "$ms ms" } else { 'no timing' }), $($Result.Reason)"
    $deliveryNote = if ($s.AlertDelivered) { 'Delivered' } else { 'Not delivered. This server had no working mail path earlier, so this is the first notice.' }

    if ($s.IsSlow) {
        $s.Polls++
        if ($slowPoll) { $s.SlowPolls++ }
        if ($Failed) { $s.FailedPolls++ }
        if ($slowPoll -or (-not $Failed -and $null -ne $ms)) { if ($ms -gt $s.WorstMs) { $s.WorstMs = $ms } }
        # A period carried over a restart starts with an empty window, so nothing is decided until it is half covered again
        $covered = ($NowUtc - $s.Samples[0].Utc).TotalSeconds -ge ($WindowMinutes * 30)
        if ($covered -and $share -le $ClearPercent) {
            # The period ended at the first good poll after the last bad one, not when the window finally cleared
            $lastBad = -1
            for ($k = $s.Samples.Count - 1; $k -ge 0; $k--) { if ($s.Samples[$k].Slow -or $s.Samples[$k].Failed) { $lastBad = $k; break } }
            if ($lastBad -eq $s.Samples.Count - 1) { $normalUtc = $NowUtc; $normalLocal = [string]$Result.Timestamp_Local; $tail = 0 }
            else { $normalUtc = $s.Samples[$lastBad + 1].Utc; $normalLocal = $s.Samples[$lastBad + 1].LocalStr; $tail = $s.Samples.Count - 1 - $lastBad }
            $span = $normalUtc - $s.StartUtc
            if ($span.Ticks -lt 0) { $span = [TimeSpan]::Zero }
            $duration = Format-Duration $span
            $pollsWhileSlow = [math]::Max(0, $s.Polls - $tail)
            $transitionLog = "Slow period over for ${SiteName}: lasted $duration, $pollsWhileSlow polls, $($s.SlowPolls) slow, $($s.FailedPolls) failed, worst $($s.WorstMs) ms."
            $dropKind = 'SLOWCLEAR'
            if ($AlertOnRecovery) {
                $emailKind = 'SlowResolved'
                $emailSubject = "[SLOW RESOLVED] $siteLabel - slow period lasted $duration"
                $emailBody = New-AlertBody -Heading "$what is responding normally again from $SiteName" -Details ([ordered]@{ 'Status' = 'SLOW RESOLVED'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Slow from' = "$($s.StartLocalStr) local"; 'Normal from' = "$normalLocal local"; 'Duration' = $duration; 'Polls while slow' = "${pollsWhileSlow}: $($s.SlowPolls) slower than $SlowThresholdMs ms, $($s.FailedPolls) failed"; 'Worst response' = "$($s.WorstMs) ms"; 'Last poll' = $lastPoll; 'SLOW alert' = $deliveryNote })
            }
            & $closePeriod
        }
        elseif ($covered -and $ReAlertMinutes -gt 0 -and $null -ne $s.LastAlertUtc -and ($NowUtc - $s.LastAlertUtc).TotalMinutes -ge $ReAlertMinutes) {
            $s.LastAlertUtc = $NowUtc
            $elapsed = Format-Duration ($NowUtc - $s.StartUtc)
            $emailKind = 'SlowReminder'
            $emailSubject = "[STILL SLOW] $siteLabel - slow for $elapsed"
            $emailBody = New-AlertBody -Heading "$what is still slow from $SiteName" -Details ([ordered]@{ 'Status' = 'STILL SLOW'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Slow since' = "$($s.StartLocalStr) local"; 'Slow for' = $elapsed; 'Polls while slow' = "$($s.Polls): $($s.SlowPolls) slower than $SlowThresholdMs ms, $($s.FailedPolls) failed"; 'Worst response' = "$($s.WorstMs) ms"; "Last $WindowMinutes min" = "$windowText, $timing"; 'Last poll' = $lastPoll; 'Earlier alerts' = $deliveryNote })
            $transitionLog = "Reminder raised for $SiteName, slow for $elapsed."
            $dropKind = 'SLOWSTILL'
        }
    }
    elseif ($bad.Count -ge 3 -and $share -ge $AlertPercent -and ($NowUtc - $s.Samples[0].Utc).TotalSeconds -ge ($WindowMinutes * 30)) {
        $first = $bad[0]
        $s.IsSlow = $true
        $s.StartUtc = $first.Utc
        $s.StartLocalStr = $first.LocalStr
        $s.LastAlertUtc = $NowUtc
        $s.AlertDelivered = $false
        # Counters start where the period starts, so the totals never include the healthy polls before it
        $fromFirst = @($s.Samples | Where-Object { $_.Utc -ge $first.Utc })
        $s.Polls = $fromFirst.Count
        $s.SlowPolls = @($fromFirst | Where-Object { $_.Slow }).Count
        $s.FailedPolls = @($fromFirst | Where-Object { $_.Failed }).Count
        $s.WorstMs = if ($okMs.Count) { [int]$okMs[-1] } else { 0 }
        $emailKind = 'Slow'
        $emailSubject = "[SLOW] $siteLabel - $($bad.Count) of $total polls slow or failed in $WindowMinutes min"
        $emailBody = New-AlertBody -Heading "$what is slow from $SiteName" -Details ([ordered]@{ 'Status' = 'SLOW'; 'Site' = $SiteName; 'Server' = $HostName; 'Endpoint' = $Url; 'Slow since' = "$($first.LocalStr) local"; "Last $WindowMinutes min" = $windowText; 'Response time' = $timing; 'Last poll' = $lastPoll; 'Backend IP' = $Result.RemoteIp; 'Clears when' = "$ClearPercent% or fewer of the polls in $WindowMinutes min are slow or failed" })
        $transitionLog = "Declared SLOW for ${SiteName}: $windowText ($timing)."
        $dropKind = 'SLOWSTART'
    }

    [PSCustomObject]@{ State = $s; EmailKind = $emailKind; EmailSubject = $emailSubject; EmailBody = $emailBody; TransitionLog = $transitionLog; DropKind = $dropKind }
}

function Test-DailySummaryDue {
    # Pure. True once a day at or after the summary hour, when no summary was recorded for that date
    param([Parameter(Mandatory)][datetime]$NowLocal, [int]$Hour, [string]$LastSentDate)
    if ($Hour -lt 0 -or $NowLocal.Hour -lt $Hour) { return $false }
    return ($LastSentDate -ne $NowLocal.ToString('yyyy-MM-dd'))
}

function Get-DailySummary {
    # Pure. Summarises the drops log lines from the 24 hours before NowLocal. Returns $null when nothing went wrong.
    param([string[]]$Lines, [Parameter(Mandatory)][datetime]$NowLocal, [int]$SlowThresholdMs = 3000, [string]$SiteName, [string]$HostName, [string]$Url, [string]$MonitorName)
    $from = $NowLocal.AddHours(-24)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $failed = 0; $slow = 0; $worst = 0; $downs = 0; $slowPeriods = 0; $restarts = 0; $errors = 0; $open = $false
    $reasons = @{}; $hours = @{}
    $outages = New-Object System.Collections.ArrayList
    foreach ($line in @($Lines)) {
        if ("$line" -notmatch '^\uFEFF?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d) \| (\w+) +\| (.*)$') { continue }
        $kind = $Matches[2]; $text = $Matches[3]
        $t = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', $inv, [System.Globalization.DateTimeStyles]::None, [ref]$t)) { continue }
        if ($t -le $from -or $t -gt $NowLocal) { continue }
        if ($kind -eq 'FAIL' -or $kind -eq 'SLOW') { $h = $t.ToString('HH'); $hours[$h] = 1 + [int]$hours[$h] }
        if ($kind -eq 'FAIL') {
            $failed++
            $r = if ($text -match 'Reason=(.+)$') { $Matches[1].Trim() } else { 'unknown' }
            $reasons[$r] = 1 + [int]$reasons[$r]
        }
        elseif ($kind -eq 'SLOW') { $slow++; if ($text -match 'Total=(\d+)ms' -and [int]$Matches[1] -gt $worst) { $worst = [int]$Matches[1] } }
        elseif ($kind -eq 'DOWN') { $downs++; $open = $true }
        elseif ($kind -eq 'RESOLVED') {
            # An outage whose DOWN line is older than the window still counts once, with its length
            if (-not $open) { $downs++ }
            if ($text -match 'lasted (.+?) over') { [void]$outages.Add($Matches[1]) }
            $open = $false
        }
        elseif ($kind -eq 'REMINDER' -or ($kind -eq 'CARRYOVER' -and $text -like 'Outage in progress*')) { if (-not $open) { $downs++; $open = $true } }
        elseif ($kind -eq 'SLOWSTART') { $slowPeriods++ }
        elseif ($kind -eq 'RESTART') { $restarts++ }
        elseif ($kind -eq 'ERROR') { $errors++ }
    }
    if (($failed + $slow + $downs + $slowPeriods + $restarts + $errors) -eq 0) { return $null }
    $count = { param([int]$n, [string]$word) if ($n -eq 1) { "1 $word" } else { "$n ${word}s" } }
    $details = [ordered]@{
        'Monitor'      = $MonitorName
        'Site'         = $SiteName
        'Server'       = $HostName
        'Period'       = "$($from.ToString('yyyy-MM-dd HH:mm')) to $($NowLocal.ToString('yyyy-MM-dd HH:mm')) local"
        'Slow polls'   = if ($slow) { "$slow slower than $SlowThresholdMs ms, worst $worst ms" } else { 'None' }
        'Failed polls' = if ($failed) { "$failed (" + ((@($reasons.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key) x$($_.Value)" })) -join ', ') + ")" } else { 'None' }
        'Outages'      = if ($downs) { "$downs$(if ($outages.Count) { ', lasting ' + ($outages -join ', ') })$(if ($open) { ', one still in progress' })" } else { 'None' }
        'Slow periods' = if ($slowPeriods) { "$slowPeriods" } else { 'None' }
    }
    if ($hours.Count) {
        $top = $hours.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
        $details['Busiest hour'] = "$($top.Key):00 to $($top.Key):59, $($top.Value) slow or failed polls"
    }
    if ($restarts) { $details['Monitor restarts'] = "$restarts" }
    if ($errors) { $details['Monitor errors'] = "$errors" }
    $details['Endpoint'] = $Url
    $details['Details'] = "Drops.log on $HostName"
    [PSCustomObject]@{
        Subject      = "[DAILY] $(Get-AlertLabel -MonitorName $MonitorName -SiteName $SiteName -HostName $HostName) - $(& $count $slow 'slow poll'), $(& $count $failed 'failed poll'), $(& $count $downs 'outage') in 24 hours"
        Body         = (New-AlertBody -Heading "Daily summary for $(if ($MonitorName) { $MonitorName } else { 'the endpoint' }) at $SiteName" -Details $details)
        SlowPolls    = $slow
        FailedPolls  = $failed
        Outages      = $downs
        SlowPeriods  = $slowPeriods
        WorstMs      = $worst
    }
}

function Send-AlertEmail {
    # Sends one alert. Returns $true when the mail host accepted it. Failure to send never stops the monitor.
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body)
    if (-not $SendEmail) { return $true }
    if ($MailMethod -eq 'Graph') {
        try {
            $token = Get-GraphToken
            if ($null -eq $token) { Write-Log -Level FAILED -Message "Could not obtain a Graph token for alert '$Subject'."; return $false }
            Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$MailFrom/sendMail" -Headers @{ Authorization = "Bearer $token" } -Body (ConvertTo-GraphMailBody -Subject $Subject -Body $Body -To $MailTo) -ContentType 'application/json; charset=utf-8' -TimeoutSec 30 -ErrorAction Stop | Out-Null
            Write-Log -Level INFORMATIONAL -Message "Alert email sent via Graph: $Subject"
            return $true
        }
        catch {
            $script:GraphToken = $null
            Write-Log -Level FAILED -Message "Could not send alert email via Graph '$Subject'. $($_.Exception.Message) $(Get-RestErrorDetail $_)"
            return $false
        }
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $p = @{ SmtpServer = $SmtpServer; Port = $SmtpPort; From = $MailFrom; To = $MailTo; Subject = $Subject; Body = $Body; BodyAsHtml = $true }
    if ($SmtpUseSsl) { $p['UseSsl'] = $true }
    try {
        $cred = Get-SmtpCredential
        if ($SmtpAuthUser -and -not $cred) { Write-Log -Level FAILED -Message "No usable SMTP credential for alert '$Subject'."; return $false }
        if ($cred) { $p['Credential'] = $cred }
        Send-MailMessage @p -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log -Level INFORMATIONAL -Message "Alert email sent: $Subject"
        return $true
    }
    catch {
        Write-Log -Level FAILED -Message "Could not send alert email '$Subject'. $($_.Exception.Message)"
        return $false
    }
}

function Send-AlertOrQueue {
    # Sends now. When the mail path is down the message is queued and retried once a minute for up to an hour.
    # A newer message about the same outage makes older undelivered ones stale, so they are dropped.
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string]$Kind)
    $stale = switch ($Kind) { 'Resolved' { @('Down','Reminder') } 'Reminder' { @('Down','Reminder') } 'Down' { @('Slow','SlowReminder') } 'SlowResolved' { @('Slow','SlowReminder') } 'SlowReminder' { @('Slow','SlowReminder') } 'Restart' { @('Restart') } 'Summary' { @('Summary') } default { @() } }
    foreach ($old in @($script:PendingAlerts | Where-Object { $_.Kind -in $stale })) {
        Write-Log -Level WARNING -Message "Dropping undelivered '$($old.Subject)', superseded by '$Subject'."
        Write-DropLog -Kind 'ALERT' -Message "Dropped undelivered '$($old.Subject)', superseded by '$Subject'."
        $script:PendingAlerts.Remove($old)
    }
    if (Send-AlertEmail -Subject $Subject -Body $Body) { Write-DropLog -Kind 'ALERT' -Message "Sent: $Subject"; return $true }
    $now = (Get-Date).ToUniversalTime()
    while ($script:PendingAlerts.Count -ge 5) {
        $dropped = $script:PendingAlerts[0]
        Write-Log -Level WARNING -Message "Dropping undelivered '$($dropped.Subject)': the retry queue is full."
        Write-DropLog -Kind 'ALERT' -Message "Dropped undelivered, retry queue full: $($dropped.Subject)"
        $script:PendingAlerts.RemoveAt(0)
    }
    [void]$script:PendingAlerts.Add(@{ Subject = $Subject; Body = $Body; Kind = $Kind; FirstUtc = $now; LastUtc = $now })
    Write-Log -Level WARNING -Message "'$Subject' will be retried every minute for up to 60 minutes."
    Write-DropLog -Kind 'ALERT' -Message "Not sent, retrying every minute for up to 60 minutes: $Subject"
    return $false
}

function Send-PendingAlert {
    # Retries queued messages, one attempt per message per minute. Returns the kinds delivered on this pass.
    $delivered = @()
    $now = (Get-Date).ToUniversalTime()
    foreach ($item in @($script:PendingAlerts)) {
        # A clock that went backwards leaves these in the future, which would freeze the retries and the give-up
        if ($item.FirstUtc -gt $now) { $item.FirstUtc = $now }
        if ($item.LastUtc -gt $now) { $item.LastUtc = $now }
        if (($now - $item.LastUtc).TotalSeconds -lt 60) { continue }
        $item.LastUtc = $now
        if (Send-AlertEmail -Subject $item.Subject -Body $item.Body) {
            Write-Log -Level INFORMATIONAL -Message "Delivered on retry: $($item.Subject)"
            Write-DropLog -Kind 'ALERT' -Message "Delivered on retry: $($item.Subject)"
            $delivered += $item.Kind
            $script:PendingAlerts.Remove($item)
        }
        elseif (($now - $item.FirstUtc).TotalMinutes -ge 60) {
            Write-Log -Level FAILED -Message "Giving up on '$($item.Subject)' after 60 minutes of failed sends."
            Write-DropLog -Kind 'ALERT' -Message "Gave up after 60 minutes of failed sends: $($item.Subject)"
            $script:PendingAlerts.Remove($item)
        }
    }
    return $delivered
}

function Write-DropLog {
    # Appends one line to the drops log, which holds what went wrong and how the alerts about it fared: failed and slow
    # polls, outage transitions, probe errors, every alert delivery outcome, and monitor starts, restarts, and clean
    # stops. Never a healthy poll. Never throws. Rotated files past 10 MB are kept beside it. Lines that could not be
    # written while the file was locked are counted and reported once it is writable again.
    param([Parameter(Mandatory)][string]$Kind, [Parameter(Mandatory)][string]$Message)
    try {
        if ((Test-Path $DropLog) -and ((Get-Item $DropLog).Length -gt 10MB)) {
            $aside = $DropLog -replace '\.log$', ('_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
            Move-Item -Path $DropLog -Destination $aside -Force -ErrorAction Stop
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        if ($script:DropLogLost -gt 0) {
            Add-Content -Path $DropLog -Value ("{0} | {1,-9} | {2}" -f $stamp, 'LOST', "$($script:DropLogLost) line(s) were not recorded while this file was locked by another program.") -Encoding UTF8 -ErrorAction Stop
            $script:DropLogLost = 0
        }
        Add-Content -Path $DropLog -Value ("{0} | {1,-9} | {2}" -f $stamp, $Kind, $Message) -Encoding UTF8 -ErrorAction Stop
        $script:DropLogWarned = $false
    }
    catch {
        $script:DropLogLost = [int]$script:DropLogLost + 1
        if (-not $script:DropLogWarned) {
            $script:DropLogWarned = $true
            Write-Log -Level WARNING -Message "Could not write the drops log '$DropLog'. $($_.Exception.Message)"
        }
    }
}

function Write-Heartbeat {
    # Records that the monitor is alive plus the outage state, so a restart can measure its gap and carry an outage over
    param([Parameter(Mandatory)][hashtable]$State, [Parameter(Mandatory)][datetime]$NowUtc, [hashtable]$SlowState)
    $iso = { param($d) if ($null -eq $d) { $null } else { ([datetime]$d).ToUniversalTime().ToString('o') } }
    $doc = [ordered]@{
        Beat                = $NowUtc.ToUniversalTime().ToString('o')
        IsDown              = [bool]$State.IsDown
        ConsecutiveFailures = [int]$State.ConsecutiveFailures
        OutageStartUtc      = (& $iso $State.OutageStartUtc)
        OutageStartLocalStr = $State.OutageStartLocalStr
        OutageStartUtcStr   = $State.OutageStartUtcStr
        LastAlertUtc        = (& $iso $State.LastAlertUtc)
        AlertDelivered      = $State.AlertDelivered
        IsSlow              = [bool]$SlowState.IsSlow
        SlowStartUtc        = (& $iso $SlowState.StartUtc)
        SlowStartLocalStr   = $SlowState.StartLocalStr
        SlowLastAlertUtc    = (& $iso $SlowState.LastAlertUtc)
        SlowPolls           = [int]$SlowState.Polls
        SlowSlowPolls       = [int]$SlowState.SlowPolls
        SlowFailedPolls     = [int]$SlowState.FailedPolls
        SlowWorstMs         = [int]$SlowState.WorstMs
        SlowAlertDelivered  = $SlowState.AlertDelivered
        Stopped             = $false
    }
    try {
        # Written to a temp file and renamed into place, so a crash or kill mid-write can never leave a truncated heartbeat
        $tmp = "$HeartbeatFile.tmp"
        $doc | ConvertTo-Json -Compress | Set-Content -Path $tmp -Encoding UTF8 -Force
        # A reader holding the file open can block the replace for a moment, so try again before giving up
        $moved = $false
        for ($attempt = 1; -not $moved; $attempt++) {
            try { Move-Item -Path $tmp -Destination $HeartbeatFile -Force -ErrorAction Stop; $moved = $true }
            catch { if ($attempt -ge 3) { throw }; Start-Sleep -Milliseconds 200 }
        }
        $script:HeartbeatWarned = $false
    }
    catch {
        if (-not $script:HeartbeatWarned) {
            $script:HeartbeatWarned = $true
            Write-Log -Level WARNING -Message "Could not write the heartbeat file '$HeartbeatFile'. Restart gaps will not be reported until this clears. $($_.Exception.Message)"
        }
    }
}

function Read-Heartbeat {
    # Returns the last heartbeat with real UTC [datetime] values, or $null when the file is missing or unreadable
    if (-not (Test-Path $HeartbeatFile)) { return $null }
    try {
        $j = Get-Content -Path $HeartbeatFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        # PowerShell 7 hands ISO strings back as [datetime] already; Windows PowerShell leaves them as strings
        $parse = { param($s)
            if ($s -is [datetime]) { if ($s.Kind -eq 'Unspecified') { return [datetime]::SpecifyKind($s, 'Utc') } else { return $s.ToUniversalTime() } }
            if ([string]::IsNullOrWhiteSpace($s)) { return $null }
            return [datetime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
        }
        $beat = & $parse $j.Beat
        if ($null -eq $beat) { throw "no Beat value" }
        return [PSCustomObject]@{
            BeatUtc             = $beat
            IsDown              = [bool]$j.IsDown
            ConsecutiveFailures = [int]$j.ConsecutiveFailures
            OutageStartUtc      = (& $parse $j.OutageStartUtc)
            OutageStartLocalStr = [string]$j.OutageStartLocalStr
            OutageStartUtcStr   = [string]$j.OutageStartUtcStr
            LastAlertUtc        = (& $parse $j.LastAlertUtc)
            AlertDelivered      = $j.AlertDelivered
            IsSlow              = [bool]$j.IsSlow
            SlowStartUtc        = (& $parse $j.SlowStartUtc)
            SlowStartLocalStr   = [string]$j.SlowStartLocalStr
            SlowLastAlertUtc    = (& $parse $j.SlowLastAlertUtc)
            SlowPolls           = [int]$j.SlowPolls
            SlowSlowPolls       = [int]$j.SlowSlowPolls
            SlowFailedPolls     = [int]$j.SlowFailedPolls
            SlowWorstMs         = [int]$j.SlowWorstMs
            SlowAlertDelivered  = $j.SlowAlertDelivered
            Stopped             = [bool]$j.Stopped
        }
    }
    catch {
        Write-Log -Level WARNING -Message "Heartbeat file '$HeartbeatFile' is unreadable and will be replaced. $($_.Exception.Message)"
        return $null
    }
}

function Get-RestartNotice {
    # Pure. Compares the previous heartbeat with now. Returns $null when there was no previous heartbeat, else an object
    # with Subject and Body (both $null when the gap is under the threshold), RestoredState (an outage to carry over, or
    # $null), and GapSeconds. A clock that went backwards reads as a gap of zero.
    param([object]$Previous, [Parameter(Mandatory)][datetime]$NowUtc, [datetime]$BootTimeUtc = [datetime]::MinValue, [Parameter(Mandatory)][int]$GapThresholdSeconds, [Parameter(Mandatory)][int]$IntervalSeconds, [string]$SiteName, [string]$HostName, [string]$Url, [string]$MonitorName, [int]$SlowCarryOverSeconds = 0)
    if ($null -eq $Previous) { return $null }
    $restored = $null
    if ($Previous.IsDown -and $null -ne $Previous.OutageStartUtc) {
        $restored = @{ ConsecutiveFailures = [int]$Previous.ConsecutiveFailures; IsDown = $true; LastAlertUtc = $Previous.LastAlertUtc; OutageStartUtc = $Previous.OutageStartUtc; OutageStartLocalStr = $Previous.OutageStartLocalStr; OutageStartUtcStr = $Previous.OutageStartUtcStr; AlertDelivered = [bool]$Previous.AlertDelivered }
    }
    $gap = $NowUtc - $Previous.BeatUtc
    if ($gap.Ticks -lt 0) { $gap = [TimeSpan]::Zero }
    # The installer marks the heartbeat when it stops the monitor on purpose: no notice, but any outage still carries over
    if ($Previous.Stopped -or $gap.TotalSeconds -lt $GapThresholdSeconds) {
        return [PSCustomObject]@{ Subject = $null; Body = $null; Cause = $null; RestoredState = $restored; GapSeconds = [int]$gap.TotalSeconds; Stopped = [bool]$Previous.Stopped }
    }
    $gapText = Format-Duration $gap
    $fmt = { param([datetime]$d) "$($d.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')) local ($($d.ToString('yyyy-MM-dd HH:mm:ss')) UTC)" }
    $cause = if ($BootTimeUtc -eq [datetime]::MinValue) { "Could not read the server boot time." }
             elseif ($BootTimeUtc -gt $Previous.BeatUtc) { "Server restarted at $(& $fmt $BootTimeUtc), after the last heartbeat." }
             else { "Server did not restart (up since $(& $fmt $BootTimeUtc)). The monitor process or its scheduled task stopped, or the server was suspended." }
    $outage = if ($restored) { "Yes. It has been down since $($Previous.OutageStartLocalStr) local. Tracking continues from that onset." } else { "None at the last heartbeat." }
    $details = [ordered]@{
        'Monitor'            = $MonitorName
        'Site'               = $SiteName
        'Server'             = $HostName
        'Last heartbeat'     = (& $fmt $Previous.BeatUtc)
        'Monitor restarted'  = (& $fmt $NowUtc)
        'Not running for'    = $gapText
        'Polls missed'       = [int][math]::Floor($gap.TotalSeconds / [math]::Max(1, $IntervalSeconds))
        'Cause'              = $cause
        'Outage in progress' = $outage
        'Endpoint'           = $Url
    }
    if ($Previous.IsSlow) {
        $slowKept = ($SlowCarryOverSeconds -gt 0 -and [int]$gap.TotalSeconds -le $SlowCarryOverSeconds)
        $details['Slow period in progress'] = "Yes, since $($Previous.SlowStartLocalStr) local. $(if ($slowKept) { 'Tracking continues from that start.' } else { 'Slow tracking starts fresh.' })"
    }
    [PSCustomObject]@{
        Subject       = "[MONITOR RESTARTED] $(Get-AlertLabel -MonitorName $MonitorName -SiteName $SiteName -HostName $HostName) - not running for $gapText"
        Body          = (New-AlertBody -Heading "$(if ($MonitorName) { "The $MonitorName monitor" } else { 'The monitor' }) restarted on $SiteName" -Details $details)
        Cause         = $cause
        RestoredState = $restored
        Stopped       = $false
        GapSeconds    = [int]$gap.TotalSeconds
    }
}

function Wait-NetworkReady {
    # After a boot the task can start before name resolution works, which would look like an outage. Waits until the
    # endpoint host resolves, up to the timeout. Returns $true when it resolved.
    param([Parameter(Mandatory)][string]$Url, [int]$TimeoutSeconds = 180)
    $endpointHost = ([uri]$Url).Host
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $waited = $false
    while ($true) {
        try {
            [void][System.Net.Dns]::GetHostAddresses($endpointHost)
            if ($waited) { Write-Log -Level FOUND -Message "Name resolution for $endpointHost is available. Starting polls." }
            return $true
        }
        catch { }
        if ((Get-Date) -ge $deadline) { Write-Log -Level WARNING -Message "Name resolution for $endpointHost still fails after $TimeoutSeconds s. Starting polls anyway."; return $false }
        if (-not $waited) { Write-Log -Level INFORMATIONAL -Message "Waiting for name resolution of $endpointHost (up to $TimeoutSeconds s) before polling." }
        $waited = $true
        Start-Sleep -Seconds 5
    }
}

Write-Log -Level STARTED -Message "Curl monitor '$MonitorName' starting for site '$SiteName'. Watching $Url every $IntervalSeconds s, timeout $TimeoutSeconds s, down threshold $DownThreshold, mail via $MailMethod."

if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
    Write-Log -Level FAILED -Message "curl.exe not found. This host needs Windows 10 1803+ or Server 2019+. Exiting."
    exit 1
}
Write-Log -Level "SANITY CHECK" -Message "curl.exe present."

if (-not (Test-Path $InstallDir)) {
    try {
        New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
        Write-Log -Level CREATED -Message "Created output folder '$InstallDir'."
    }
    catch {
        Write-Log -Level FAILED -Message "Could not create output folder '$InstallDir'. $($_.Exception.Message). Exiting."
        exit 1
    }
}

$transcriptPath = Get-TranscriptPath
Start-Transcript -Path $transcriptPath -Append | Out-Null
$startupCutoff = (Get-Date).AddDays(-$LogRetentionDays)
Get-ChildItem -Path $InstallDir -Filter 'Transcript_*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $startupCutoff } | Remove-Item -Force -ErrorAction SilentlyContinue
$lastExpiryCheckDate = [datetime]::MinValue

$state = @{ ConsecutiveFailures = 0; IsDown = $false; LastAlertUtc = $null; OutageStartUtc = $null; OutageStartLocalStr = $null; OutageStartUtcStr = $null; AlertDelivered = $null }
$script:PendingAlerts = New-Object System.Collections.ArrayList
$script:HeartbeatWarned = $false

# A previous heartbeat means this is a restart. Measure the gap, report it, and carry any outage in progress over.
$bootUtc = [datetime]::MinValue
try { $bootUtc = ([datetime](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime).ToUniversalTime() }
catch { Write-Log -Level WARNING -Message "Could not read the server boot time. $($_.Exception.Message)" }
$startUtc = (Get-Date).ToUniversalTime()
$previous = Read-Heartbeat
$notice = Get-RestartNotice -Previous $previous -NowUtc $startUtc -BootTimeUtc $bootUtc -GapThresholdSeconds ([math]::Max(60, 2 * ($IntervalSeconds + $TimeoutSeconds))) -IntervalSeconds $IntervalSeconds -SiteName $SiteName -HostName $env:COMPUTERNAME -Url $Url -MonitorName $MonitorName -SlowCarryOverSeconds ($SlowWindowMinutes * 60)
Write-DropLog -Kind 'START' -Message "Monitor started on $env:COMPUTERNAME for site '$SiteName' (poll every $IntervalSeconds s, timeout $TimeoutSeconds s, down after $DownThreshold failures, slow at $SlowAlertPercent% of polls over $SlowThresholdMs ms or failed in $SlowWindowMinutes min)."
if ($notice) {
    if ($notice.Subject) {
        Write-Log -Level WARNING -Message "The monitor was not running for $(Format-Duration ([TimeSpan]::FromSeconds($notice.GapSeconds))). Sending a restart notice."
        Write-DropLog -Kind 'RESTART' -Message "Monitor was not running for $(Format-Duration ([TimeSpan]::FromSeconds($notice.GapSeconds))). $($notice.Cause)"
    }
    elseif ($notice.Stopped) { Write-Log -Level INFORMATIONAL -Message "Monitor restarted after a deliberate stop (reinstall) $($notice.GapSeconds) s ago. No restart notice." }
    else { Write-Log -Level INFORMATIONAL -Message "Monitor restarted after a $($notice.GapSeconds) s gap, below the notice threshold." }
    if ($notice.RestoredState) {
        $state = $notice.RestoredState
        Write-Log -Level WARNING -Message "An outage was in progress at the last heartbeat (since $($state.OutageStartLocalStr) local). Tracking continues from that onset."
        Write-DropLog -Kind 'CARRYOVER' -Message "Outage in progress since $($state.OutageStartLocalStr) local carried over from before the restart ($($state.ConsecutiveFailures) failed polls so far)."
    }
}
$slowState = @{ Samples = @(); IsSlow = $false; StartUtc = $null; StartLocalStr = $null; LastAlertUtc = $null; Polls = 0; SlowPolls = 0; FailedPolls = 0; WorstMs = 0; AlertDelivered = $null }
if ($previous -and $previous.IsSlow -and $null -ne $previous.SlowStartUtc -and -not $state.IsDown) {
    # Only a short gap keeps the period: after a long one nobody watched the gap, and the restart notice says so
    if ($notice -and $notice.GapSeconds -le ($SlowWindowMinutes * 60)) {
        $slowState.IsSlow = $true
        $slowState.StartUtc = $previous.SlowStartUtc
        $slowState.StartLocalStr = $previous.SlowStartLocalStr
        $slowState.LastAlertUtc = $previous.SlowLastAlertUtc
        $slowState.Polls = $previous.SlowPolls
        $slowState.SlowPolls = $previous.SlowSlowPolls
        $slowState.FailedPolls = $previous.SlowFailedPolls
        $slowState.WorstMs = $previous.SlowWorstMs
        $slowState.AlertDelivered = [bool]$previous.SlowAlertDelivered
        Write-Log -Level WARNING -Message "A slow period was in progress at the last heartbeat (since $($slowState.StartLocalStr) local). Tracking continues from that start."
        Write-DropLog -Kind 'CARRYOVER' -Message "Slow period since $($slowState.StartLocalStr) local carried over from before the restart."
    }
    else {
        $gapText = Format-Duration ([TimeSpan]::FromSeconds([int]$notice.GapSeconds))
        Write-Log -Level WARNING -Message "A slow period since $($previous.SlowStartLocalStr) local was not carried over because the monitor was not running for $gapText. Slow tracking starts fresh."
        Write-DropLog -Kind 'SLOWCLEAR' -Message "Slow period since $($previous.SlowStartLocalStr) local not carried over: the monitor was not running for $gapText. Slow tracking starts fresh."
    }
}
$summarySent = ''
try { if (Test-Path $SummaryStateFile) { $summarySent = "$(Get-Content -Path $SummaryStateFile -TotalCount 1 -ErrorAction Stop)".Trim() } } catch { }
if (-not $summarySent -and $DailySummaryHour -ge 0 -and (Get-Date).Hour -ge $DailySummaryHour) {
    # A first start after today's summary hour waits for tomorrow instead of summarising a day it did not watch
    $summarySent = (Get-Date).ToString('yyyy-MM-dd')
    try { Set-Content -Path $SummaryStateFile -Value $summarySent -Encoding ASCII -ErrorAction Stop } catch { }
}
Write-Heartbeat -State $state -NowUtc $startUtc -SlowState $slowState
Wait-NetworkReady -Url $Url | Out-Null
if ($SendEmail -and $notice -and $notice.Subject) { Send-AlertOrQueue -Subject $notice.Subject -Body $notice.Body -Kind 'Restart' | Out-Null }

try {
    while ($true) {
        try {
            $transcriptPath = Update-TranscriptIfDayChanged -CurrentPath $transcriptPath
            if ((Get-Date).Date -ne $lastExpiryCheckDate) {
                $lastExpiryCheckDate = (Get-Date).Date
                $w = Get-SecretExpiryWarning -ExpiresOn $GraphSecretExpires -Today (Get-Date)
                if ($w) { Write-Log -Level WARNING -Message $w }
            }

            Write-Heartbeat -State $state -NowUtc ((Get-Date).ToUniversalTime()) -SlowState $slowState
            $result = Get-ProbeResult

            # A non-zero curl exit means the transfer did not complete, even when a 200 status had already arrived
            $failed  = ($result.CurlExit -ne 0) -or ($result.HttpCode -ne '200') -or (-not $result.ContentOk) -or ($null -eq $result.TotalMs)
            $summary = "Site=$($result.SiteName) Code=$($result.HttpCode) Redirects=$($result.Redirects) TTFB=$($result.TtfbMs)ms Total=$($result.TotalMs)ms Populated=$($result.ContentOk) IP=$($result.RemoteIp) Reason=$($result.Reason)"

            if ($failed) { Write-Log -Level FAILED -Message $summary; Write-DropLog -Kind 'FAIL' -Message $summary }
            elseif ($result.TotalMs -ge $SlowThresholdMs) { Write-Log -Level WARNING -Message "Slow response. $summary"; Write-DropLog -Kind 'SLOW' -Message $summary }
            else { Write-Log -Level SUCCESS -Message $summary }

            $decision = Update-MonitorState -State $state -Failed $failed -Result $result -NowUtc ((Get-Date).ToUniversalTime()) -DownThreshold $DownThreshold -ReAlertMinutes $ReAlertMinutes -AlertOnRecovery $AlertOnRecovery -SiteName $SiteName -Url $Url -HostName $env:COMPUTERNAME -MonitorName $MonitorName
            $state = $decision.State

            if ($decision.TransitionLog) { Write-Log -Level ADDED -Message $decision.TransitionLog }
            if ($decision.OutageRecord) { Write-DropLog -Kind 'RESOLVED' -Message "$($decision.TransitionLog) Recovery HTTP $($result.HttpCode) from $($result.RemoteIp)." }
            elseif ($decision.EmailKind -eq 'Down') { Write-DropLog -Kind 'DOWN' -Message $decision.TransitionLog }
            elseif ($decision.EmailKind -eq 'Reminder') { Write-DropLog -Kind 'REMINDER' -Message $decision.TransitionLog }
            if ($decision.OutageRecord)  { Write-CsvRow -Path $OutageCsv -Row $decision.OutageRecord }
            if ($SendEmail -and $decision.EmailSubject) {
                $ok = Send-AlertOrQueue -Subject $decision.EmailSubject -Body $decision.EmailBody -Kind $decision.EmailKind
                if ($ok -and $decision.EmailKind -in @('Down','Reminder')) { $state.AlertDelivered = $true }
            }

            $slow = Update-SlowState -State $slowState -Result $result -Failed $failed -IsDown $state.IsDown -NowUtc ((Get-Date).ToUniversalTime()) -SlowThresholdMs $SlowThresholdMs -WindowMinutes $SlowWindowMinutes -AlertPercent $SlowAlertPercent -ClearPercent $SlowClearPercent -ReAlertMinutes $ReAlertMinutes -AlertOnRecovery $AlertOnRecovery -SiteName $SiteName -Url $Url -HostName $env:COMPUTERNAME -MonitorName $MonitorName
            $slowState = $slow.State
            if ($slow.TransitionLog) { Write-Log -Level ADDED -Message $slow.TransitionLog; Write-DropLog -Kind $slow.DropKind -Message $slow.TransitionLog }
            if ($SendEmail -and $AlertOnSlow -and $slow.EmailSubject) {
                $okSlow = Send-AlertOrQueue -Subject $slow.EmailSubject -Body $slow.EmailBody -Kind $slow.EmailKind
                if ($okSlow -and $slow.EmailKind -in @('Slow','SlowReminder')) { $slowState.AlertDelivered = $true }
            }

            if (Test-DailySummaryDue -NowLocal (Get-Date) -Hour $DailySummaryHour -LastSentDate $summarySent) {
                $summarySent = (Get-Date).ToString('yyyy-MM-dd')
                try { Set-Content -Path $SummaryStateFile -Value $summarySent -Encoding ASCII -ErrorAction Stop }
                catch { Write-Log -Level WARNING -Message "Could not record the daily summary date in '$SummaryStateFile'. $($_.Exception.Message)" }
                $dropLines = @()
                try {
                    # A rotation past the size limit splits the day across files, so read every log the window touches
                    $cutoff = (Get-Date).AddHours(-24)
                    foreach ($f in @(Get-ChildItem -Path $InstallDir -Filter 'Drops_*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $cutoff } | Sort-Object Name)) {
                        $dropLines += @(Get-Content -Path $f.FullName -ErrorAction Stop)
                    }
                    if (Test-Path $DropLog) { $dropLines += @(Get-Content -Path $DropLog -ErrorAction Stop) }
                }
                catch { Write-Log -Level WARNING -Message "Could not read the drops log for the daily summary. $($_.Exception.Message)" }
                $summary = Get-DailySummary -Lines $dropLines -NowLocal (Get-Date) -SlowThresholdMs $SlowThresholdMs -SiteName $SiteName -HostName $env:COMPUTERNAME -Url $Url -MonitorName $MonitorName
                if ($summary) {
                    Write-Log -Level INFORMATIONAL -Message "Daily summary: $($summary.Subject)"
                    if ($SendEmail) { Send-AlertOrQueue -Subject $summary.Subject -Body $summary.Body -Kind 'Summary' | Out-Null }
                }
                else { Write-Log -Level INFORMATIONAL -Message "Daily summary: no slow or failed polls, outages, or restarts in the last 24 hours. No email." }
            }

            $deliveredKinds = @(Send-PendingAlert)
            if ($state.IsDown -and (@($deliveredKinds | Where-Object { $_ -in @('Down','Reminder') }).Count -gt 0)) { $state.AlertDelivered = $true }
            if ($slowState.IsSlow -and (@($deliveredKinds | Where-Object { $_ -in @('Slow','SlowReminder') }).Count -gt 0)) { $slowState.AlertDelivered = $true }
            Write-CsvRow -Path (Get-LatencyCsvPath) -Row $result
            Write-Heartbeat -State $state -NowUtc ((Get-Date).ToUniversalTime()) -SlowState $slowState
        }
        catch {
            Write-Log -Level ERROR -Message "Probe cycle error: $($_.Exception.Message)"
            Write-DropLog -Kind 'ERROR' -Message "Probe cycle error: $($_.Exception.Message)"
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
}
finally {
    Write-Log -Level FINISHED -Message "Curl monitor '$MonitorName' stopping for site '$SiteName'."
    Write-DropLog -Kind 'STOP' -Message "Monitor stopping for site '$SiteName'."
    try { Stop-Transcript | Out-Null } catch { }
}
'@

Write-Log -Level STARTED -Message "Curl monitor self-installer starting."

# A PowerShell 7 window puts its own module folders first in PSModulePath. Windows PowerShell then finds the wrong
# Microsoft.PowerShell.Security and cannot load it, so ConvertTo-SecureString goes missing. Keep only its own paths.
if ($PSVersionTable.PSVersion.Major -le 5) {
    $env:PSModulePath = (@($env:PSModulePath -split ';') | Where-Object { $_ -and ($_ -notmatch '\\PowerShell\\Modules$|\\PowerShell\\7\\Modules$') }) -join ';'
}

Test-ScriptBeingRanAsAdministrator

if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
    Write-Log -Level FAILED -Message "curl.exe not found. This host needs Windows 10 1803+ or Server 2019+. Exiting."
    exit 1
}
Write-Log -Level "SANITY CHECK" -Message "curl.exe present."

if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
    Write-Log -Level FAILED -Message "The ScheduledTasks module is not available on this OS. Exiting."
    exit 1
}
Write-Log -Level "SANITY CHECK" -Message "ScheduledTasks module present."

if ("$Action" -ne 'Install' -and "$Action" -ne 'Uninstall') {
    Write-Log -Level FAILED -Message "Unknown action '$Action'. Set `$Action to Install or Uninstall. Exiting."
    exit 1
}
if (-not $NonInteractive) {
    Write-Question -Question "Install or uninstall?" -Hint @("I  install a monitor, or upgrade one already here", "U  remove a monitor from this server")
    $Action = if ((Read-Choice -Prompt "Choose" -Allowed @('I','U') -Default 'I') -eq 'U') { 'Uninstall' } else { 'Install' }
}
if ("$Action" -eq 'Uninstall') { exit (Invoke-UninstallFlow) }

if (-not (Test-Path $InstallRoot)) {
    try {
        New-Item -Path $InstallRoot -ItemType Directory -Force | Out-Null
        Write-Log -Level CREATED -Message "Created install root '$InstallRoot'."
    }
    catch {
        Write-Log -Level FAILED -Message "Could not create install root '$InstallRoot'. $($_.Exception.Message). Exiting."
        exit 1
    }
}

# Monitors from the old location move first, so the rest of this run sees them where everything else looks
$null = Invoke-RootMove -Monitors (Get-PreviousRootMonitor)
# Anything still there was declined or could not be moved, and is still polling
$elsewhereMonitors = @(Get-PreviousRootMonitor | Where-Object { $null -ne $_ })

# An older HST-only install is offered a migration before the prompts, so its settings prefill them
$existingMonitors = @(Get-InstalledMonitor | Where-Object { $null -ne $_ })
$legacy = Get-LegacyInstall
$migrate = $false
if ($legacy) {
    Write-Question -Question "Move the older install at $($legacy.Dir) into ${InstallRoot}?" -Hint @(
        "Its settings and history come across, then its task and folder go. Nothing is removed until this run is finished.",
        "N leaves it running where it is."
    )
    $migrate = if ($NonInteractive) { $true } else { (Read-Choice -Prompt "Move" -Allowed @('Y','N') -Default 'Y') -eq 'Y' }
}

if ($legacy -and -not $migrate) {
    $elsewhereMonitors += [PSCustomObject]@{
        Name = $LegacyMonitorName; Slug = (ConvertTo-MonitorSlug $LegacyMonitorName); Dir = $legacy.Dir; Url = $legacy.Url
        TaskName = $legacy.TaskName; TaskPath = $legacy.TaskPath; HasTask = $legacy.HasTask
    }
}
$MonitorName = Get-MonitorName -SavedDefault $(if ($migrate) { $LegacyMonitorName } else { "" }) -Existing $existingMonitors -Elsewhere $elsewhereMonitors
if (-not $MonitorName) {
    Write-Log -Level FAILED -Message "No monitor name. Nothing was installed. Exiting."
    exit 1
}
$InstallDir = Join-Path $InstallRoot (ConvertTo-MonitorSlug $MonitorName)
$TaskName   = "$TaskNamePrefix$MonitorName"

if (-not (Test-Path $InstallDir)) {
    try {
        New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
        Write-Log -Level CREATED -Message "Created install directory '$InstallDir'."
    }
    catch {
        Write-Log -Level FAILED -Message "Could not create install directory '$InstallDir'. $($_.Exception.Message). Exiting."
        exit 1
    }
}
try {
    Protect-InstallFolder -Path $InstallRoot
    Protect-InstallFolder -Path $InstallDir
    Protect-StoredSecret -Root $InstallRoot
    Write-Log -Level "SANITY CHECK" -Message "Install folder locked to SYSTEM and Administrators (Users read-only), stored secrets locked to SYSTEM and Administrators only."
}
catch {
    Write-Log -Level FAILED -Message "Could not secure the install folder. $($_.Exception.Message) Exiting."
    exit 1
}

$saved = Get-InstallSettings
if (-not $saved -and $legacy -and $legacy.Settings) {
    $saved = $legacy.Settings
    Write-Log -Level FOUND -Message "Prompts default to the older install's settings."
}
$site  = Get-SiteName -SavedDefault $(if ($saved -and $saved.SiteName) { $saved.SiteName } else { "" })
$Url   = Get-MonitorUrl -SavedDefault $(if ($saved -and $saved.Url) { [string]$saved.Url } else { "" })
if (-not $Url) {
    Write-Log -Level FAILED -Message "No URL to watch. Nothing was installed. Exiting."
    exit 1
}
$ExpectedContentMarker = Get-ContentMarker -SavedDefault $(if ($saved -and $saved.ContentMarker) { [string]$saved.ContentMarker } else { "" })

if ($AlertsEnabled) {
    $mail = Get-MailConfiguration -Saved $saved -SiteName $site -MonitorName $MonitorName
    if (-not $mail) {
        Write-Log -Level FAILED -Message "Mail configuration incomplete. Nothing was installed. Exiting."
        exit 1
    }
}
else {
    # Telemetry only: the monitor records everything it measures and sends nothing
    Write-Log -Level INFORMATIONAL -Message "Alerts are turned off in this install. No mail is configured and none is sent. Outages, slow periods, and every poll are still recorded."
    $mail = @{ MailMethod = 'None'; SmtpServer = ''; SmtpPort = 25; SmtpUseSsl = $false; MailFrom = ''; MailTo = @(); SmtpAuthUser = ''; CipherText = ''; GraphTenantId = ''; GraphClientId = ''; GraphSecretExpires = '' }
}

$monitorPath = Join-Path -Path $InstallDir -ChildPath $MonitorFileName
$content = New-MonitorContent -SiteName $site -Mail $mail

if ($content -cmatch '@@[A-Z]+@@') {
    Write-Log -Level FAILED -Message "Generated monitor still contains an unresolved placeholder ($($Matches[0])). Not writing anything. Exiting."
    exit 1
}
try {
    [ScriptBlock]::Create($content) | Out-Null
    Write-Log -Level "SANITY CHECK" -Message "Generated monitor script parses cleanly."
}
catch {
    Write-Log -Level FAILED -Message "Generated monitor failed to parse. Not writing anything. $($_.Exception.Message). Exiting."
    exit 1
}

if ($mail.MailMethod -in @('Authenticated','Graph') -and $mail.CipherText) {
    try { Save-SmtpCredential -CipherText $mail.CipherText }
    catch {
        Write-Log -Level FAILED -Message "Could not store the SMTP credential. Nothing was installed. $($_.Exception.Message). Exiting."
        exit 1
    }
}
else {
    Remove-SmtpCredential
}

# Saved now so a failure below still prefills the next run. InstalledAt is added only once the task is running.
$settings = @{
    MonitorName  = $MonitorName
    SiteName     = $site
    MailMethod   = $mail.MailMethod
    SmtpServer   = $mail.SmtpServer
    SmtpPort     = $mail.SmtpPort
    SmtpUseSsl   = $mail.SmtpUseSsl
    MailFrom     = $mail.MailFrom
    MailTo       = @($mail.MailTo)
    SmtpAuthUser = $mail.SmtpAuthUser
    GraphTenantId      = $mail.GraphTenantId
    GraphClientId      = $mail.GraphClientId
    GraphSecretExpires = $mail.GraphSecretExpires
    Url           = $Url
    ContentMarker = $ExpectedContentMarker
}
Save-InstallSettings -Settings $settings

# The new monitor is written beside the old one and swapped in only after it verified, so nothing here can leave a broken file
$stagedPath = "$monitorPath.new"
try {
    Set-Content -Path $stagedPath -Value $content -Encoding UTF8 -Force
    if (-not (Test-Path $stagedPath) -or ((Get-Item $stagedPath).Length -eq 0)) {
        Write-Log -Level FAILED -Message "Monitor file was not written correctly to '$stagedPath'. Exiting."
        exit 1
    }
    $written = Get-Content -Path $stagedPath -Raw
    if ($written.Trim() -ne $content.Trim()) {
        Write-Log -Level FAILED -Message "Monitor file on disk does not match the generated content. Exiting."
        exit 1
    }
}
catch {
    Write-Log -Level FAILED -Message "Could not write monitor script. $($_.Exception.Message). Exiting."
    exit 1
}

$existing = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
if ($existing) {
    # The running copy is stopped and the task definition replaced in place. It is never unregistered, so a
    # failure below still leaves a registered monitor that starts at the next boot.
    $wasRunning = ([string]$existing.State -eq 'Running')
    try {
        Stop-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
        if (-not (Wait-TaskStopped)) { Write-Log -Level WARNING -Message "Existing task is still reported running after 15 s. Continuing with reinstall." }
        elseif ($wasRunning) {
            Write-Log -Level INFORMATIONAL -Message "Stopped the existing monitor for a clean reinstall."
            # This stop is deliberate: mark the heartbeat so the new copy sends no restart notice, while any outage in
            # progress stays recorded and carries over. The drops log gets the stop the monitor could not write itself.
            $hbPath = Join-Path $InstallDir 'heartbeat.json'
            if (Test-Path $hbPath) {
                try {
                    $hb = Get-Content -Path $hbPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    $hb | Add-Member -NotePropertyName Stopped -NotePropertyValue $true -Force
                    $hb | ConvertTo-Json -Compress | Set-Content -Path $hbPath -Encoding UTF8 -Force
                }
                catch { Remove-Item -Path $hbPath -Force -ErrorAction SilentlyContinue }
            }
            try { Add-Content -Path (Join-Path $InstallDir 'Drops.log') -Value ("{0} | {1,-9} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), 'STOP', "Monitor stopped by the installer on $env:COMPUTERNAME for a reinstall.") -Encoding UTF8 -ErrorAction Stop } catch { }
        }
        else { Write-Log -Level WARNING -Message "The existing monitor task was not running. The new copy will report how long it was down." }
    }
    catch { Write-Log -Level WARNING -Message "Could not stop the existing task. $($_.Exception.Message)" }
}

try {
    Move-Item -Path $stagedPath -Destination $monitorPath -Force -ErrorAction Stop
    Write-Log -Level CREATED -Message "Wrote and verified monitor script at '$monitorPath'."
}
catch {
    Write-Log -Level FAILED -Message "Could not replace the monitor script at '$monitorPath'. $($_.Exception.Message). Exiting."
    exit 1
}

# Only now, with mail settled and the monitor written, is the old install taken apart. An abort before this point
# leaves it running, so the server is never left with nothing watching.
if ($migrate) {
    if (Invoke-LegacyMigration -Legacy $legacy -Destination $InstallDir) {
        Write-Log -Level FINISHED -Message "Migration complete: '$($legacy.Dir)' is gone and its history lives in '$InstallDir'."
    }
    else {
        Write-Log -Level WARNING -Message "Migration did not finish. The old folder '$($legacy.Dir)' is still there and its task is not running. Nothing was lost, and this monitor takes over from here."
    }
    Protect-StoredSecret -Root $InstallRoot
}

try {
    Register-MonitorTask -Name $TaskName -Path $TaskPath -ScriptPath $monitorPath
    Write-Log -Level CREATED -Message "Registered scheduled task '$TaskPath$TaskName' running as $RunAsUser at startup."
}
catch {
    Write-Log -Level FAILED -Message "Could not register the scheduled task. $($_.Exception.Message.Trim()) Exiting."
    exit 1
}

$taskState = "Unknown"
try {
    Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop
    for ($i = 0; $i -lt 10; $i++) {
        Start-Sleep -Seconds 1
        $taskState = [string](Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue).State
        if ($taskState -eq 'Running') { break }
    }
    if (-not $taskState) { $taskState = 'Missing' }
    if ($taskState -eq 'Running') { Write-Log -Level SUCCESS -Message "Task is running." }
    else { Write-Log -Level FAILED -Message "Task registered but state is '$taskState' after 10 s. Check Task Scheduler history. It will also start on next boot." }
}
catch {
    $taskState = 'NotStarted'
    Write-Log -Level FAILED -Message "Task registered but could not be started now. It will start on next boot. $($_.Exception.Message)"
}

if ($taskState -eq 'Running') {
    $settings['InstalledAt'] = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $settings['InstalledBy'] = "$env:USERDOMAIN\$env:USERNAME"
    # Records which credential the stored file belongs to, so a later re-run only offers to keep it for that same app or user
    $settings['CredentialFor'] = switch ($mail.MailMethod) { 'Graph' { $mail.GraphClientId } 'Authenticated' { $mail.SmtpAuthUser } default { '' } }
    Save-InstallSettings -Settings $settings
}

$reach = Test-EndpointReachable
if ($reach.Reachable) { Write-Log -Level FOUND -Message "Preflight probe reached '$Url' from '$site' (HTTP $($reach.HttpCode) after $($reach.Redirects) redirect(s))." }
else { Write-Log -Level WARNING -Message "Preflight probe did not get HTTP 200 from '$site' (got '$($reach.HttpCode)' after $($reach.Redirects) redirect(s), final URL '$($reach.FinalUrl)'). The monitor is installed and will keep trying." }

if ($SendInstallTestEmail -and $AlertsEnabled) {
    $cred = $null; $graphSecret = ""
    if ($mail.MailMethod -eq 'Graph' -and $mail.CipherText) {
        $graphSecret = Get-StoredSecret
        if (-not $graphSecret) { Write-Log -Level WARNING -Message "Could not read back the stored Graph secret for the install email." }
    }
    $canSend = $true
    if ($mail.MailMethod -eq 'Authenticated' -and $mail.SmtpAuthUser) {
        $plain = Get-StoredSecret
        if ($plain) { $cred = New-Object System.Management.Automation.PSCredential($mail.SmtpAuthUser, (ConvertTo-SecureText -Text $plain)) }
        else { $canSend = $false; Write-Log -Level WARNING -Message "Could not read back the stored SMTP password. Skipping the install confirmation email." }
    }
    $body = New-AlertBody -Heading "Curl monitor '$MonitorName' installed" -Details ([ordered]@{ 'Monitor' = $MonitorName; 'Site' = $site; 'Server' = $env:COMPUTERNAME; 'Endpoint' = $Url; 'Task' = "$TaskPath$TaskName ($taskState)"; 'Preflight' = "HTTP $($reach.HttpCode) after $($reach.Redirects) redirect(s)"; 'Alerts via' = "$($mail.MailMethod) from $($mail.MailFrom)"; 'Down alert' = "After $DownThreshold failed polls in a row"; 'Slow alert' = $(if ($AlertOnSlow) { "When $SlowAlertPercent% of polls in $SlowWindowMinutes min are slower than $SlowThresholdMs ms or fail" } else { 'Off' }); 'Daily summary' = $(if ($DailySummaryHour -ge 0) { "At $('{0:00}' -f $DailySummaryHour):00 when anything was slow or failed" } else { 'Off' }); 'Installed' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }) -Footer "Sent by the curl monitor installer on $env:COMPUTERNAME."
    if ($script:MailTestSkipped) { Write-Log -Level WARNING -Message "Install confirmation email skipped: the test email failed and was skipped. The monitor sends alerts on its own once the mail path works." }
    elseif ($canSend -and (Send-MailWithConfig -Mail $mail -Subject "[MONITOR INSTALLED] $(Get-AlertLabel -MonitorName $MonitorName -SiteName $site -HostName $env:COMPUTERNAME)" -Body $body -Credential $cred -GraphSecret $graphSecret -MaxWaitSeconds 120)) {
        Write-Log -Level INFORMATIONAL -Message "Install confirmation email sent."
    }
    elseif ($canSend) { Write-Log -Level WARNING -Message "Install confirmation email not sent. The monitor sends alerts on its own once the mail path works." }
}

Write-Host ""
Write-Host "Install summary"
Write-Host "  Monitor         : $MonitorName"
Write-Host "  Site            : $site"
Write-Host "  URL             : $Url"
Write-Host "  Healthy poll    : HTTP 200, at least $MinPopulatedBytes bytes$(if ($ExpectedContentMarker) { ", containing '$ExpectedContentMarker'" } else { '' })"
Write-Host "  Task            : $TaskPath$TaskName ($taskState, runs as $RunAsUser at startup)"
Write-Host "  Folder          : $InstallDir  (the monitor, its settings, and its history)"
if ($AlertsEnabled) {
    Write-Host "  Alerts          : $($mail.MailMethod) from $($mail.MailFrom) to $($mail.MailTo -join ', ')"
    Write-Host "  Down alert      : after $DownThreshold failed polls in a row"
    Write-Host "  Slow alert      : $(if ($AlertOnSlow) { "when $SlowAlertPercent% of polls in $SlowWindowMinutes min are slower than $SlowThresholdMs ms or fail" } else { 'off (slow periods are still logged)' })"
    Write-Host "  Daily summary   : $(if ($DailySummaryHour -ge 0) { "at $('{0:00}' -f $DailySummaryHour):00 when anything was slow or failed" } else { 'off' })"
}
else {
    Write-Host "  Alerts          : off, telemetry only. Outages and slow periods are recorded, no mail is sent."
}
if ($AlertsEnabled -and $mail.MailMethod -eq 'Graph') {
    Write-Host "  Tenant ID       : $($mail.GraphTenantId)"
    Write-Host "  Client ID       : $($mail.GraphClientId)"
    Write-Host "  Secret expires  : $(if ($mail.GraphSecretExpires) { $mail.GraphSecretExpires } else { 'not recorded' })"
}
Write-Host "  Preflight       : HTTP $($reach.HttpCode) after $($reach.Redirects) redirect(s)"
Write-Host ""
if ($taskState -ne 'Running') {
    Write-Log -Level FAILED -Message "Install finished but the monitor task is not running (state '$taskState'). Fix the task in Task Scheduler or re-run the installer."
    exit 1
}
Write-Log -Level FINISHED -Message "Install complete: '$MonitorName' at site '$site' watching $Url."
