<#
.SYNOPSIS
    One-shot diagnostic for the Curl Monitor Graph mail path. Proves whether the app can send and, if not, exactly why.

.DESCRIPTION
    Run on the site server (or any machine) after the installer has run tenant setup.
    1. Loads Tenant ID, Client ID, sender, and recipient from the installer's saved settings when present,
       and the client secret from the DPAPI credential file when readable, so no typing is needed on the server.
    2. Requests a client-credentials token and prints the identity claims it carries (appid, oid, tid, roles).
       The oid claim is the service principal object id Exchange authorizes against.
    3. Attempts one sendMail as the shared mailbox and prints the full server response body.
    4. Prints what the outcome means and the exact Exchange Online checks to run if it failed.

.REQUIREMENTS
    Windows PowerShell 5.1. Network access to login.microsoftonline.com and graph.microsoft.com.
    Reading the stored secret needs administrator rights; otherwise the secret is prompted.

.OUTPUTS
    Console output only. Sends at most one email.

.NOTES
    Author:      Christopher Carroll
    Created:     09/04/2026
    Idempotency: Read-only apart from the single test email. Safe to re-run.
    Context:     Compass HST monitor. Discriminates a stale Exchange authorization cache from a
                 structural misconfiguration after RBAC for Applications setup.

.LINK
    https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac
#>

Remove-Variable * -ErrorAction SilentlyContinue

# Fill these to skip all prompting, or leave blank to use the installer's saved settings on this machine
$GraphTenantId  = ""
$GraphClientId  = ""
$ClientSecret   = ""                       # Plaintext secret for ad hoc use. Left blank, the DPAPI file or a hidden prompt is used.
$MailFrom       = ""
$MailTo         = ""

$InstallRoot         = "C:\ProgramData\CurlMonitor"
$MonitorName         = ""                  # Which monitor's settings to read. Blank picks the only one installed.
$SettingsFileName    = "install-settings.json"
$CredentialFileName  = "credential.bin"

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

function ConvertFrom-JwtPart {
    # Decodes one base64url JWT segment to an object
    param([Parameter(Mandatory)][string]$Part)
    $s = $Part.Replace('-', '+').Replace('_', '/')
    switch ($s.Length % 4) { 2 { $s += '==' } 3 { $s += '=' } }
    return ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($s)) | ConvertFrom-Json)
}

function Get-RestErrorDetail {
    # Pulls the service's own error code and message out of a web exception
    param($ErrorRecord)
    $detail = ""
    try { $detail = $ErrorRecord.ErrorDetails.Message } catch { }
    if (-not $detail) { try { $stream = $ErrorRecord.Exception.Response.GetResponseStream(); $detail = (New-Object System.IO.StreamReader($stream)).ReadToEnd() } catch { } }
    return "$detail"
}

Write-Log -Level STARTED -Message "Curl Monitor mail diagnostic."
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Prefill from the installer's saved settings
# One folder per monitor under the root, each with its own settings and its own stored secret
$candidates = @(Get-ChildItem -Path $InstallRoot -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path (Join-Path $_.FullName $SettingsFileName) })
if ($MonitorName) {
    $wantSlug = ($MonitorName -replace '[^A-Za-z0-9]+', '-').Trim('-')
    if ($wantSlug.Length -gt 60) { $wantSlug = $wantSlug.Substring(0, 60).Trim('-') }
    $wanted = @($candidates | Where-Object { $_.Name -eq $wantSlug -or $_.Name -eq $MonitorName })
    if (-not $wanted.Count) { Write-Log -Level FAILED -Message "No monitor called '$MonitorName' under '$InstallRoot'. Installed: $(if ($candidates.Count) { (@($candidates | ForEach-Object { $_.Name }) -join ', ') } else { 'none' })."; exit 1 }
    $candidates = $wanted
}
if (@($candidates).Count -gt 1 -and -not ($GraphTenantId -and $GraphClientId -and $ClientSecret -and $MailFrom -and $MailTo)) {
    Write-Log -Level FAILED -Message "More than one monitor is installed under '$InstallRoot': $((@($candidates | ForEach-Object { $_.Name })) -join ', '). Set `$MonitorName at the top of this script to pick one."
    exit 1
}
$InstallDir = if (@($candidates).Count -eq 1) { @($candidates)[0].FullName } else { $InstallRoot }
$settingsPath = Join-Path $InstallDir $SettingsFileName
if (Test-Path $settingsPath) {
    try {
        $saved = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
        if (-not $GraphTenantId -and $saved.GraphTenantId) { $GraphTenantId = $saved.GraphTenantId }
        if (-not $GraphClientId -and $saved.GraphClientId) { $GraphClientId = $saved.GraphClientId }
        if (-not $MailFrom -and $saved.MailFrom)           { $MailFrom = $saved.MailFrom }
        if (-not $MailTo -and $saved.MailTo)               { $MailTo = @($saved.MailTo) | Select-Object -First 1 }
        Write-Log -Level FOUND -Message "Loaded saved settings from $settingsPath."
    }
    catch { Write-Log -Level WARNING -Message "Could not read saved settings. $($_.Exception.Message)" }
}
if (-not $GraphTenantId) { $GraphTenantId = (Read-Host "Tenant ID").Trim() }
if (-not $GraphClientId) { $GraphClientId = (Read-Host "Client ID (Application ID)").Trim() }
if (-not $MailFrom)      { $MailFrom      = (Read-Host "Sender mailbox").Trim() }
if (-not $MailTo)        { $MailTo        = (Read-Host "Send the test to").Trim() }

# Secret: config value, then the DPAPI file, then a hidden prompt
if (-not $ClientSecret) {
    $credPath = Join-Path $InstallDir $CredentialFileName
    if (Test-Path $credPath) {
        try {
            Add-Type -AssemblyName System.Security
            $cipher  = [Convert]::FromBase64String((Get-Content -Path $credPath -Raw).Trim())
            $entropy = [System.Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1')
            $ClientSecret = [System.Text.Encoding]::UTF8.GetString([System.Security.Cryptography.ProtectedData]::Unprotect($cipher, $entropy, [System.Security.Cryptography.DataProtectionScope]::LocalMachine))
            Write-Log -Level FOUND -Message "Decrypted the stored client secret."
        }
        catch { Write-Log -Level WARNING -Message "Could not read the stored secret (needs admin on the install server). $($_.Exception.Message)" }
    }
}
if (-not $ClientSecret) {
    $secure = Read-Host "Client secret (paste, input hidden)" -AsSecureString
    $ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
}

# Step 1: token
try {
    $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$GraphTenantId/oauth2/v2.0/token" -Body @{ client_id = $GraphClientId; client_secret = $ClientSecret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    Write-Log -Level SUCCESS -Message "Token issued."
}
catch {
    Write-Log -Level FAILED -Message "Token request failed. $($_.Exception.Message)"
    Write-Host (Get-RestErrorDetail $_)
    Write-Host ""
    Write-Host "A token failure is credentials, not permissions: wrong Tenant ID, wrong Client ID, or wrong/expired secret."
    exit 1
}

$claims = ConvertFrom-JwtPart (($tok.access_token -split '\.')[1])
$roles = "(none)"
if ($claims.roles) { $roles = ($claims.roles -join ', ') }
Write-Host ""
Write-Host "Token identity:"
Write-Host "  appid (application)        : $($claims.appid)"
Write-Host "  oid (service principal)    : $($claims.oid)"
Write-Host "  tid (tenant)               : $($claims.tid)"
Write-Host "  roles (Entra permissions)  : $roles"
Write-Host ""
Write-Host "The oid above is what Exchange authorizes against. It must equal BOTH of these:"
Write-Host "  1. Entra portal: Enterprise applications -> Curl Monitor -> Overview -> Object ID"
Write-Host "  2. Exchange:     Get-ServicePrincipal | fl DisplayName,AppId,ObjectId"
Write-Host ""

# Step 2: one send
$body = @{ message = @{ subject = "[MONITOR TEST] diagnostic $(Get-Date -Format 'HH:mm:ss')"; body = @{ contentType = 'Text'; content = "Diagnostic send from Test-CurlMonitorMailSend.ps1 at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." }; toRecipients = @(@{ emailAddress = @{ address = $MailTo } }) }; saveToSentItems = $false } | ConvertTo-Json -Depth 6
try {
    Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$MailFrom/sendMail" -Headers @{ Authorization = "Bearer $($tok.access_token)" } -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    Write-Log -Level SUCCESS -Message "Send accepted. Check $MailTo. The mail path works; nothing is broken."
    Write-Log -Level FINISHED -Message "Diagnostic complete."
    exit 0
}
catch {
    $raw = Get-RestErrorDetail $_
    Write-Log -Level FAILED -Message "Send failed. $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Full server response:"
    Write-Host $raw
    Write-Host ""
    Write-Host "What this means:"
    if ($raw -match 'ErrorAccessDenied') {
        Write-Host "  ErrorAccessDenied = Exchange refused this app for this mailbox. The token reached Exchange;"
        Write-Host "  this is not a Graph consent problem. More than 2 hours after the role assignment this should"
        Write-Host "  no longer be cache. Run these from Connect-ExchangeOnline as the admin and compare:"
        Write-Host ""
        Write-Host "    Get-ServicePrincipal | fl DisplayName,AppId,ObjectId"
        Write-Host "        ObjectId must equal the token oid printed above. Mismatch = the Exchange pointer is wrong:"
        Write-Host "        Remove-ServicePrincipal -Identity <bad>; New-ServicePrincipal -AppId $GraphClientId -ObjectId <token oid> -DisplayName 'Curl Monitor'"
        Write-Host ""
        Write-Host "    Test-ServicePrincipalAuthorization -Identity <ObjectId> -Resource $MailFrom | ft"
        Write-Host "        Expect RoleName 'Application Mail.Send' with InScope True."
        Write-Host ""
        Write-Host "    Get-Mailbox -Identity $MailFrom | fl PrimarySmtpAddress"
        Write-Host "        Must exactly equal the address in the management scope filter."
        Write-Host ""
        Write-Host "    Get-OrganizationConfig | fl *Rbac*"
        Write-Host "        Note any EnforceExoAppRbacPermissions value for the escalation."
    }
    elseif ($raw -match 'Authorization_RequestDenied|Insufficient privileges') {
        Write-Host "  This is a Graph gateway consent rejection, which contradicts the RBAC design. Capture this"
        Write-Host "  output; the fallback is re-granting Mail.Send consent plus the application access policy."
    }
    elseif ($raw -match 'ErrorInvalidUser|ResourceNotFound|MailboxNotEnabledForRESTAPI') {
        Write-Host "  The sender address did not resolve to a licensed-or-shared mailbox. Verify $MailFrom exists:"
        Write-Host "    Get-Mailbox -Identity $MailFrom | fl RecipientTypeDetails,PrimarySmtpAddress"
    }
    else {
        Write-Host "  Unrecognized failure. Capture this full output for escalation."
    }
    Write-Log -Level FINISHED -Message "Diagnostic complete."
    exit 1
}
