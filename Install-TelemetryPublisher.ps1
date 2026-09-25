<#
.SYNOPSIS
    Sets up the twice-daily job that publishes curl monitor telemetry to a git repository.

.DESCRIPTION
    Run once from an elevated Windows PowerShell console on the machine that holds the monitors. It asks for the
    repository, the working copy, the two run times, and the commit identity, takes a personal access token at a
    hidden prompt, and stores the token encrypted with machine-scope DPAPI in a file only SYSTEM and
    Administrators can read. It then clones or updates the working copy, registers a SYSTEM scheduled task with
    one trigger per run time, and proves the pipeline with a dry run.

    The token is never written to the console, a transcript, the task definition, or the git remote.

.PARAMETER PublisherPath
    Publish-UptimeTelemetry.ps1 to schedule. Defaults to the copy beside this script.

.PARAMETER StateDir
    Folder for the token, the publish state, and the endpoint map. Never published.

.PARAMETER TaskPath
    Task Scheduler folder. Default \CurlMonitor\.

.PARAMETER NonInteractive
    Uses the parameter values without prompting. Requires -RepoUrl, -AuthorName, -AuthorEmail, and -Token.

.EXAMPLE
    .\Install-TelemetryPublisher.ps1

.NOTES
    Author:      Christopher Carroll
    Created:     09/21/2026
    Idempotency: Safe to re-run. The task, the stored token, and the settings are replaced cleanly.
#>
[CmdletBinding()]
param(
    [string]$PublisherPath = '',
    [string]$StateDir = 'C:\ProgramData\CurlMonitor-telemetry',
    [string]$RepoPath = 'C:\ProgramData\CurlMonitor-telemetry\curl-monitor',
    [string]$MonitorRoot = 'C:\ProgramData\CurlMonitor',
    [string]$TaskPath = '\CurlMonitor\',
    [string]$PreviousTaskPath = '\DIT\',
    [string]$PreviousStateDir = 'C:\ProgramData\DIT\Telemetry',
    [string]$TaskName = 'Curl Monitor telemetry publisher',
    [string]$RepoUrl = '',
    [string]$AuthorName = '',
    [string]$AuthorEmail = '',
    [string]$MorningTime = '07:15',
    [string]$EveningTime = '17:15',
    [string]$Token = '',
    [switch]$NonInteractive,
    [switch]$SkipClone
)

$TokenFileName = 'github-token.bin'
$SettingsFileName = 'publisher-settings.json'
$CarriedStateFiles = @($TokenFileName, $SettingsFileName, 'publish-state.json', 'endpoint-map.json', 'git-credentials')
$Entropy = 'DIT-CurlMonitor-Telemetry-v1'

function Move-PreviousState {
    # Carries an earlier install's token, settings, publish state and endpoint map into the new state folder.
    # Without the endpoint map the next run reassigns ENDPOINT-01 upward by first-seen order, so one endpoint can
    # end up published under two codes and one code can hold two URLs.
    param([Parameter(Mandatory)][string]$From, [Parameter(Mandatory)][string]$To, [Parameter(Mandatory)][string[]]$Names)
    if (-not $From -or $From -eq $To -or -not (Test-Path -LiteralPath $From)) { return 0 }
    $moved = 0
    foreach ($name in $Names) {
        $src = Join-Path $From $name
        $dst = Join-Path $To $name
        if (-not (Test-Path -LiteralPath $src)) { continue }
        if (Test-Path -LiteralPath $dst) { continue }
        try {
            Copy-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
            if ($name -in @($TokenFileName, 'git-credentials')) { Set-SecretFileAcl -Path $dst }
            $moved++
        }
        catch { Write-Log -Level WARNING -Message "Could not carry '$name' over from '$From'. $($_.Exception.Message)" }
    }
    if ($moved) { Write-Log -Level CREATED -Message "Carried $moved file(s) from the earlier state folder '$From' into '$To'." }
    return $moved
}

function Grant-RepoAccess {
    # The publisher runs as SYSTEM while an administrator created the working copy. Git rejects a repository whose
    # owner differs from the account running it, so the path is recorded as safe for SYSTEM and for the installing
    # administrator, and SYSTEM is given full control of the tree.
    param([Parameter(Mandatory)][string]$RepoPath)
    $gitPath = ($RepoPath -replace '\\', '/')
    $systemConfig = Join-Path $env:SystemRoot 'System32\config\systemprofile\.gitconfig'
    foreach ($target in @(@{ Args = @('--global'); Env = $null }, @{ Args = @('--file', $systemConfig); Env = $systemConfig })) {
        $existing = @(& git.exe config @($target.Args) --get-all safe.directory 2>$null)
        if ($existing -notcontains $gitPath) { & git.exe config @($target.Args) --add safe.directory $gitPath 2>&1 | Out-Null }
    }
    & icacls.exe "$RepoPath" /grant "*S-1-5-18:(OI)(CI)F" /T /C 2>&1 | Out-Null
    Write-Log -Level CREATED -Message "Recorded '$RepoPath' as a safe repository for SYSTEM and granted it full control."
}

function Write-Log {
    param([Parameter(Mandatory)][ValidateSet('STARTED', 'PROMPT', 'FOUND', 'CREATED', 'SANITY CHECK', 'INFORMATIONAL', 'WARNING', 'FAILED', 'FINISHED')][string]$Level, [Parameter(Mandatory)][string]$Message)
    Write-Host ("{0} {1} | {2}" -f (Get-Date -Format 'MM/dd/yy - hh:mm:ss tt'), $Level, $Message)
}

function Test-ElevatedSession {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Read-Setting {
    param([Parameter(Mandatory)][string]$Prompt, [string]$Default = '')
    $label = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
    $answer = Read-Host $label
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim()
}

function Test-TimeOfDay {
    param([string]$Text)
    $t = [datetime]::MinValue
    return [datetime]::TryParseExact("$Text".Trim(), 'HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$t)
}

function Test-RepoUrl {
    param([string]$Text)
    return ("$Text".Trim() -match '^https://[^/\s]+/[^/\s]+/[^/\s]+?(\.git)?$')
}

function Test-EmailLike {
    param([string]$Text)
    return ("$Text".Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Protect-Token {
    # Machine-scope DPAPI, so the SYSTEM task can read it and another machine cannot
    param([Parameter(Mandatory)][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { throw 'Refusing to store an empty token.' }
    Add-Type -AssemblyName System.Security
    $bytes = [Text.Encoding]::UTF8.GetBytes($PlainText)
    $entropy = [Text.Encoding]::UTF8.GetBytes($Entropy)
    return [Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($bytes, $entropy, [Security.Cryptography.DataProtectionScope]::LocalMachine))
}

function Unprotect-Token {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $cipher = [Convert]::FromBase64String((Get-Content -LiteralPath $Path -Raw -ErrorAction Stop).Trim())
        $entropy = [Text.Encoding]::UTF8.GetBytes($Entropy)
        return [Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect($cipher, $entropy, [Security.Cryptography.DataProtectionScope]::LocalMachine))
    }
    catch { return $null }
}

function Set-SecretFileAcl {
    # SYSTEM and Administrators only. A file created where nothing is inheritable picks up an ACE for whoever
    # created it from the account's default permissions, so anything else is removed afterwards.
    param([Parameter(Mandatory)][string]$Path)
    & icacls.exe "$Path" /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "icacls could not restrict '$Path' (exit $LASTEXITCODE)." }
    $keep = @('S-1-5-18', 'S-1-5-32-544')
    $acl = Get-Acl -Path $Path
    foreach ($rule in @($acl.Access)) {
        $sid = try { $rule.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } catch { '' }
        if ($sid -notin $keep) { [void]$acl.RemoveAccessRule($rule) }
    }
    Set-Acl -Path $Path -AclObject $acl
}

function Save-ProtectedToken {
    # Locks the file to SYSTEM and Administrators before any content lands in it
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$CipherText)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -ItemType File -Force | Out-Null }
    Set-SecretFileAcl -Path $Path
    Set-Content -LiteralPath $Path -Value $CipherText -Encoding ASCII -Force
}

function New-AuthenticatedUrl {
    # Only ever held in memory for the clone and the credential helper, never written to git config
    param([Parameter(Mandatory)][string]$RepoUrl, [Parameter(Mandatory)][string]$Token)
    return ($RepoUrl -replace '^https://', "https://x-access-token:$Token@")
}

function Set-RepoCredential {
    # Stores the token where git finds it for this repository only, outside the repository itself
    param([Parameter(Mandatory)][string]$RepoPath, [Parameter(Mandatory)][string]$RepoUrl, [Parameter(Mandatory)][string]$Token, [Parameter(Mandatory)][string]$StoreFile)
    $uri = [uri]$RepoUrl
    $line = "https://x-access-token:$Token@$($uri.Host)"
    if (-not (Test-Path -LiteralPath $StoreFile)) { New-Item -Path $StoreFile -ItemType File -Force | Out-Null }
    Set-SecretFileAcl -Path $StoreFile
    [IO.File]::WriteAllText($StoreFile, ($line + "`n"), [Text.Encoding]::ASCII)
    # A machine-wide credential manager would run first and wait on a prompt no one can answer under SYSTEM, so the
    # helper list is cleared for this working copy and only the stored file is used. The empty first entry clears
    # what git inherits; PowerShell drops an empty argument, so both entries are written into the config directly.
    # The path is given with forward slashes: a value containing a backslash and a space is run through sh by git,
    # which strips the backslashes and writes the token to a drive-relative path inside the working copy.
    $storeForGit = ($StoreFile -replace '\\', '/')

    & git.exe -C $RepoPath config --unset-all credential.helper 2>&1 | Out-Null
    Add-Content -LiteralPath (Join-Path $RepoPath '.git\config') -Value "[credential]" -Encoding ASCII
    Add-Content -LiteralPath (Join-Path $RepoPath '.git\config') -Value "`thelper = " -Encoding ASCII
    Add-Content -LiteralPath (Join-Path $RepoPath '.git\config') -Value "`thelper = store --file='$storeForGit'" -Encoding ASCII
    # Ask git what it would actually use, rather than trusting the file we just wrote
    # Windows PowerShell prefixes a native command's stdin with a BOM when the console is UTF-8, and git would
    # read it as part of the first key, so the first line is one git discards.
    $probe = "capability=`nprotocol=https`nhost=$($uri.Host)`n`n"
    $answer = ''
    try {
        $env:GIT_TERMINAL_PROMPT = '0'
        $answer = ($probe | & git.exe -C $RepoPath credential fill 2>&1) -join "`n"
    }
    catch { $answer = "$($_.Exception.Message)" }
    finally { Remove-Item Env:\GIT_TERMINAL_PROMPT -ErrorAction SilentlyContinue }
    if ($answer -match '(?m)^password=.+$') {
        Write-Log -Level CREATED -Message "git reads the token for $($uri.Host) from '$StoreFile' for this working copy only."
    }
    else {
        Write-Log -Level FAILED -Message "git could not read the stored token for $($uri.Host) from '$StoreFile', so the scheduled push will fail. Helpers: $(@(& git.exe -C $RepoPath config --get-all credential.helper 2>$null) -join ' | ')"
    }
}

function Install-PublisherTask {
    param(
        [Parameter(Mandatory)][string]$TaskPath,
        [string]$EarlierTaskPath = '',
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$RepoPath,
        [Parameter(Mandatory)][string]$MonitorRoot,
        [Parameter(Mandatory)][string]$StateDir,
        [Parameter(Mandatory)][string]$MorningTime,
        [Parameter(Mandatory)][string]$EveningTime
    )
    $argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -RepoPath `"$RepoPath`" -MonitorRoot `"$MonitorRoot`" -StatePath `"$StateDir`""
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument
    $triggers = @(
        (New-ScheduledTaskTrigger -Daily -At ([datetime]::ParseExact($MorningTime, 'HH:mm', [Globalization.CultureInfo]::InvariantCulture))),
        (New-ScheduledTaskTrigger -Daily -At ([datetime]::ParseExact($EveningTime, 'HH:mm', [Globalization.CultureInfo]::InvariantCulture)))
    )
    $settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::FromHours(1))
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    # This task used to live under \DIT\. Register-ScheduledTask -Force only replaces the one at $TaskPath, so the
    # old one has to go or both fire at the same times, against two working copies, into the same repository.
    foreach ($old in @($EarlierTaskPath) | Where-Object { $_ -and $_ -ne $TaskPath }) {
        if (Get-ScheduledTask -TaskName $TaskName -TaskPath $old -ErrorAction SilentlyContinue) {
            try {
                Unregister-ScheduledTask -TaskName $TaskName -TaskPath $old -Confirm:$false -ErrorAction Stop
                Write-Log -Level INFORMATIONAL -Message "Removed the earlier publisher task '$old$TaskName'."
            }
            catch { Write-Log -Level WARNING -Message "'$old$TaskName' is still registered and will publish as well. Remove it by hand. $($_.Exception.Message)" }
        }
    }
    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force -ErrorAction Stop | Out-Null
    return (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop)
}

if ($MyInvocation.InvocationName -eq '.') { return }

# $PSScriptRoot is not set while parameter defaults bind under Windows PowerShell, so the publisher is resolved here
if (-not $PublisherPath) { $PublisherPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Publish-UptimeTelemetry.ps1' }

Write-Log -Level STARTED -Message 'Telemetry publisher setup starting.'
if (-not (Test-ElevatedSession)) {
    Write-Log -Level FAILED -Message 'Run this from an elevated console. Nothing was changed.'
    exit 1
}
if (-not (Test-Path -LiteralPath $PublisherPath)) {
    Write-Log -Level FAILED -Message "Publish-UptimeTelemetry.ps1 was not found at '$PublisherPath'. Nothing was changed."
    exit 1
}
if (-not (Get-Command git.exe -ErrorAction SilentlyContinue)) {
    Write-Log -Level FAILED -Message 'git.exe is not on the PATH. Install Git for Windows first.'
    exit 1
}

if (-not (Test-Path -LiteralPath $StateDir)) { New-Item -Path $StateDir -ItemType Directory -Force | Out-Null }
& icacls.exe "$StateDir" /inheritance:r /grant:r "*S-1-5-18:(OI)(CI)(F)" "*S-1-5-32-544:(OI)(CI)(F)" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Log -Level FAILED -Message "Could not restrict '$StateDir' to SYSTEM and Administrators (icacls exit $LASTEXITCODE). The stored token would be readable by others, so nothing was set up."
    exit 1
}
$null = Move-PreviousState -From $PreviousStateDir -To $StateDir -Names $CarriedStateFiles

$settingsPath = Join-Path $StateDir $SettingsFileName
$saved = $null
if (Test-Path -LiteralPath $settingsPath) {
    try { $saved = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json; Write-Log -Level FOUND -Message 'Loaded settings from the last setup. Prompts default to those values.' } catch { }
}

if (-not $NonInteractive) {
    Write-Log -Level PROMPT -Message 'Where the telemetry is published, and who it is committed as.'
    while (-not (Test-RepoUrl $RepoUrl)) {
        $RepoUrl = Read-Setting -Prompt 'Repository URL (https://github.com/owner/repo)' -Default $(if ($RepoUrl) { $RepoUrl } elseif ($saved) { [string]$saved.RepoUrl } else { '' })
        if (-not (Test-RepoUrl $RepoUrl)) { Write-Log -Level WARNING -Message 'Enter an https repository URL, for example https://github.com/owner/repo.' }
    }
    $RepoPath = Read-Setting -Prompt 'Working copy path' -Default $(if ($saved -and $saved.RepoPath) { [string]$saved.RepoPath } else { $RepoPath })
    $MonitorRoot = Read-Setting -Prompt 'Monitor root to read' -Default $(if ($saved -and $saved.MonitorRoot) { [string]$saved.MonitorRoot } else { $MonitorRoot })
    while (-not $AuthorName) { $AuthorName = Read-Setting -Prompt 'Commit author name' -Default $(if ($saved) { [string]$saved.AuthorName } else { '' }) }
    Write-Log -Level INFORMATIONAL -Message 'The author email has to be one attached to the GitHub account, or the commits will not count as contributions.'
    while (-not (Test-EmailLike $AuthorEmail)) {
        $AuthorEmail = Read-Setting -Prompt 'Commit author email' -Default $(if ($saved) { [string]$saved.AuthorEmail } else { '' })
        if (-not (Test-EmailLike $AuthorEmail)) { Write-Log -Level WARNING -Message 'Enter a valid email address.' }
    }
    while (-not (Test-TimeOfDay $MorningTime)) { $MorningTime = Read-Setting -Prompt 'First run time (HH:mm)' -Default $(if ($saved -and $saved.MorningTime) { [string]$saved.MorningTime } else { $MorningTime }) }
    while (-not (Test-TimeOfDay $EveningTime)) { $EveningTime = Read-Setting -Prompt 'Second run time (HH:mm)' -Default $(if ($saved -and $saved.EveningTime) { [string]$saved.EveningTime } else { $EveningTime }) }
}

$tokenPath = Join-Path $StateDir $TokenFileName
if (-not $Token) {
    $existing = Unprotect-Token -Path $tokenPath
    if ($NonInteractive) { $Token = $existing }
    else {
        $prompt = if ($existing) { 'Personal access token (Enter = keep the one stored on this machine, input hidden)' } else { 'Personal access token with content write on that repository (input hidden)' }
        while (-not $Token) {
            $secure = Read-Host $prompt -AsSecureString
            if ($secure -and $secure.Length -gt 0) { $Token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)) }
            elseif ($existing) { $Token = $existing; Write-Log -Level FOUND -Message 'Keeping the token stored on this machine.' }
            else { Write-Log -Level WARNING -Message 'A token is required.' }
        }
    }
}
if (-not $Token) { Write-Log -Level FAILED -Message 'No token. Nothing was changed.'; exit 1 }
if (-not (Test-RepoUrl $RepoUrl) -or -not (Test-EmailLike $AuthorEmail) -or -not $AuthorName -or -not (Test-TimeOfDay $MorningTime) -or -not (Test-TimeOfDay $EveningTime)) {
    Write-Log -Level FAILED -Message 'Repository URL, author name, author email, and both run times are all required. Nothing was changed.'
    exit 1
}

try {
    Save-ProtectedToken -Path $tokenPath -CipherText (Protect-Token -PlainText $Token)
    Write-Log -Level CREATED -Message "Stored the token at '$tokenPath' (readable by SYSTEM and Administrators only)."
}
catch {
    Write-Log -Level FAILED -Message "Could not store the token. $($_.Exception.Message). Nothing else was changed."
    exit 1
}

if (-not $SkipClone) {
    $authUrl = New-AuthenticatedUrl -RepoUrl $RepoUrl -Token $Token
    if (Test-Path -LiteralPath (Join-Path $RepoPath '.git')) {
        & git.exe -C $RepoPath remote set-url origin $RepoUrl | Out-Null
        $pull = & git.exe -C $RepoPath pull --rebase --autostash 2>&1
        if ($LASTEXITCODE -ne 0) { Write-Log -Level WARNING -Message "Could not update the working copy. $(($pull | Select-Object -Last 1))" }
        else { Write-Log -Level FOUND -Message "Working copy at '$RepoPath' updated." }
    }
    else {
        $parent = Split-Path -Path $RepoPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
        $clone = & git.exe clone $authUrl $RepoPath 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Level FAILED -Message "Clone failed. $((($clone | ForEach-Object { "$_" }) -join ' ') -replace [regex]::Escape($Token), '***')"
            exit 1
        }
        & git.exe -C $RepoPath remote set-url origin $RepoUrl | Out-Null
        Write-Log -Level CREATED -Message "Cloned the repository into '$RepoPath'."
    }
    Grant-RepoAccess -RepoPath $RepoPath
    Set-RepoCredential -RepoPath $RepoPath -RepoUrl $RepoUrl -Token $Token -StoreFile (Join-Path $StateDir 'git-credentials')
    & git.exe -C $RepoPath config user.name $AuthorName | Out-Null
    & git.exe -C $RepoPath config user.email $AuthorEmail | Out-Null
    Write-Log -Level CREATED -Message "Commits from this working copy are authored as $AuthorName <$AuthorEmail>."
}

@{ RepoUrl = $RepoUrl; RepoPath = $RepoPath; MonitorRoot = $MonitorRoot; AuthorName = $AuthorName; AuthorEmail = $AuthorEmail; MorningTime = $MorningTime; EveningTime = $EveningTime; ConfiguredAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } |
    ConvertTo-Json | Set-Content -LiteralPath $settingsPath -Encoding UTF8

try {
    $task = Install-PublisherTask -TaskPath $TaskPath -EarlierTaskPath $PreviousTaskPath -TaskName $TaskName -ScriptPath $PublisherPath -RepoPath $RepoPath -MonitorRoot $MonitorRoot -StateDir $StateDir -MorningTime $MorningTime -EveningTime $EveningTime
    Write-Log -Level CREATED -Message "Registered '$TaskPath$TaskName' to run at $MorningTime and $EveningTime as SYSTEM ($($task.Triggers.Count) triggers)."
}
catch {
    Write-Log -Level FAILED -Message "Could not register the task. $($_.Exception.Message)"
    exit 1
}

Write-Log -Level INFORMATIONAL -Message 'Proving the pipeline with a dry run. Nothing is committed.'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $PublisherPath -RepoPath $RepoPath -MonitorRoot $MonitorRoot -StatePath $StateDir -DryRun
if ($LASTEXITCODE -ne 0) { Write-Log -Level WARNING -Message 'The dry run reported a problem. Fix it before the first scheduled run.' }
else { Write-Log -Level "SANITY CHECK" -Message 'Dry run produced the files it would commit.' }

Write-Host ''
Write-Host 'Telemetry publisher'
Write-Host "  Repository      : $RepoUrl"
Write-Host "  Working copy    : $RepoPath"
Write-Host "  Monitors read   : $MonitorRoot"
Write-Host "  Commits as      : $AuthorName <$AuthorEmail>"
Write-Host "  Runs at         : $MorningTime and $EveningTime daily, catching up a missed run"
Write-Host "  Task            : $TaskPath$TaskName (SYSTEM)"
Write-Host "  Token           : $tokenPath (SYSTEM and Administrators only)"
Write-Host ''
Write-Log -Level FINISHED -Message 'Telemetry publisher ready.'
