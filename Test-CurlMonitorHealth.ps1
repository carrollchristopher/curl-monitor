<#
.SYNOPSIS
    Checks that the curl monitors on this server are running, polling, recording, and able to alert.

.DESCRIPTION
    Run from an elevated Windows PowerShell console on a server where Install-CurlMonitor.ps1 was run. Every
    monitor under C:\ProgramData\CurlMonitor is checked, one section each, with a single verdict at the end.
    Use -Monitor to check just one.
    By default it changes nothing. It checks the scheduled task and the monitor process, the heartbeat, the polls
    recorded in the last hour and the last 24 hours, problems in the monitor's own logs, the stored mail secret,
    and probes the endpoint once the same way the monitor does.

    -OutageDrill proves alerting end to end. For about a minute it makes the monitor's own probes fail, by placing
    a small curl settings file in the profile of the account the monitor runs as. The file sends curl.exe requests
    for the HST host to a local port that refuses connections. Browsers and every other program keep working, and
    curl.exe requests to other hosts are unaffected. It waits for the monitor to declare DOWN and send the alert,
    removes the file, and waits for RESOLVED. The file is removed when the drill ends or is stopped with Ctrl+C. If
    the drill window is closed, a one-time SYSTEM scheduled task removes it after DrillCleanupMinutes. The drill
    sends a real DOWN and RESOLVED email and leaves one short outage in Outages.csv and the drops log.

.PARAMETER OutageDrill
    Runs the outage drill after the checks.

.PARAMETER InstallDir
    One monitor folder to check instead of every monitor under the root.

.PARAMETER Monitor
    Check only the monitor with this name. Required with -OutageDrill when the server runs more than one.

.PARAMETER InstallRoot
    Where the installer keeps its monitors. Default C:\ProgramData\CurlMonitor.

.PARAMETER DrillCleanupMinutes
    Minutes after which a one-time scheduled task removes the drill file if the drill was interrupted. Default 10.

.EXAMPLE
    .\Test-CurlMonitorHealth.ps1

.EXAMPLE
    .\Test-CurlMonitorHealth.ps1 -OutageDrill

.NOTES
    Author:      Christopher Carroll
    Created:     09/15/2026
    Exit code:   0 all checks passed, 1 a check failed, 2 warnings only.
#>
[CmdletBinding()]
param(
    [switch]$OutageDrill,
    [string]$Monitor = '',
    [string]$InstallRoot = 'C:\ProgramData\CurlMonitor',
    [string]$InstallDir = '',
    [string]$TaskPath = '\CurlMonitor\',
    [string]$PreviousRoot = 'C:\ProgramData\DIT\CurlMonitor',
    [string]$TaskName = '',
    [int]$DrillCleanupMinutes = 10
)

$script:Results = New-Object System.Collections.ArrayList
$script:DrillMarker = '# Curl monitor outage drill. Safe to delete.'
$script:DrillCleanupTaskName = 'Curl Monitor drill cleanup'
$script:DrillMaxMinutes = 15
$script:DrillOffset = 0
$script:LoginBase = 'https://login.microsoftonline.com'

function Write-Check {
    param([Parameter(Mandatory)][ValidateSet('PASS', 'WARN', 'FAIL', 'INFO')][string]$Status, [Parameter(Mandatory)][string]$Message)
    $color = @{ PASS = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; INFO = 'Gray' }[$Status]
    Write-Host ("{0,-5} {1}" -f $Status, $Message) -ForegroundColor $color
    [void]$script:Results.Add([PSCustomObject]@{ Status = $Status; Message = $Message })
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
}

function Test-IsAdministrator {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-Age {
    param([Parameter(Mandatory)][double]$Seconds)
    $t = [TimeSpan]::FromSeconds([math]::Max(0, $Seconds))
    if ($t.TotalDays -ge 1) { return "{0}d {1}h" -f [int][math]::Floor($t.TotalDays), $t.Hours }
    if ($t.TotalHours -ge 1) { return "{0}h {1}m" -f [int][math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return "{0}m {1}s" -f [int][math]::Floor($t.TotalMinutes), $t.Seconds }
    return "{0} s" -f [int]$t.TotalSeconds
}

function ConvertFrom-LocalStamp {
    # The monitor writes Timestamp_Local and drops log stamps as yyyy-MM-dd HH:mm:ss
    param([string]$Text)
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact("$Text".Trim([char]0xFEFF, ' '), 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Get-MonitorConfig {
    # Reads the values the installer baked into the generated monitor, without running any of it
    param([Parameter(Mandatory)][string]$Path)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    $cfg = @{}
    foreach ($st in $ast.EndBlock.Statements) {
        if ($st -is [System.Management.Automation.Language.FunctionDefinitionAst]) { break }
        if ($st -isnot [System.Management.Automation.Language.AssignmentStatementAst]) { continue }
        if ($st.Left -isnot [System.Management.Automation.Language.VariableExpressionAst]) { continue }
        try { $cfg[$st.Left.VariablePath.UserPath] = $st.Right.Expression.SafeGetValue() } catch { }
    }
    return $cfg
}

function Get-Percentile {
    param([int[]]$Sorted, [double]$Percent)
    if (-not $Sorted -or $Sorted.Count -eq 0) { return $null }
    $i = [int][math]::Ceiling($Percent / 100 * $Sorted.Count) - 1
    return $Sorted[[math]::Min($Sorted.Count - 1, [math]::Max(0, $i))]
}

function Get-RecentPoll {
    # Latency rows from the last $Hours hours, oldest first. Reads only the tail of the two newest monthly files,
    # so a month of polls does not have to be loaded.
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][int]$Hours, [Parameter(Mandatory)][int]$IntervalSeconds)
    $cutoff = (Get-Date).AddHours(-$Hours)
    $want = [int][math]::Ceiling($Hours * 3600 / [math]::Max(1, $IntervalSeconds)) + 200
    $rows = New-Object System.Collections.ArrayList
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $files = @(Get-ChildItem -Path $Folder -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^Latency_\d{6}\.csv$' } | Sort-Object Name | Select-Object -Last 2)
    foreach ($f in $files) {
        $header = "$(Get-Content -Path $f.FullName -TotalCount 1 -ErrorAction SilentlyContinue)".Trim([char]0xFEFF)
        if (-not $header) { continue }
        $tail = @(Get-Content -Path $f.FullName -Tail $want -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim([char]0xFEFF) -ne $header })
        if ($tail.Count -eq 0) { continue }
        foreach ($r in (@($header) + $tail | ConvertFrom-Csv)) {
            $t = [datetime]::MinValue
            if (-not [datetime]::TryParseExact("$($r.Timestamp_Local)", 'yyyy-MM-dd HH:mm:ss', $inv, [Globalization.DateTimeStyles]::None, [ref]$t) -or $t -lt $cutoff) { continue }
            $failed = ($r.HttpCode -ne '200' -or $r.ContentOk -ne 'True' -or [string]::IsNullOrEmpty($r.TotalMs))
            [void]$rows.Add([PSCustomObject]@{ When = $t; Timestamp_Local = $r.Timestamp_Local; HttpCode = $r.HttpCode; TotalMs = $r.TotalMs; Reason = $r.Reason; Failed = $failed })
        }
    }
    return @($rows | Sort-Object When)
}

function Get-RestErrorDetail {
    param($ErrorRecord)
    $detail = ""
    try { $detail = $ErrorRecord.ErrorDetails.Message } catch { }
    if (-not $detail) { try { $detail = (New-Object System.IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())).ReadToEnd() } catch { } }
    $detail = ("$detail" -replace '\s*\r?\n\s*', ' ').Trim()
    if ($detail -match '(AADSTS\d+)[^"]*') { return $Matches[0] }
    if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) }
    return $detail
}

function Get-StoredMailSecret {
    # Decrypts the machine-scope DPAPI file the monitor reads. Returns $null when missing or unreadable.
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        Add-Type -AssemblyName System.Security
        $cipher = [Convert]::FromBase64String((Get-Content -Path $Path -Raw -ErrorAction Stop).Trim())
        $entropy = [Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1')
        return [Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect($cipher, $entropy, [Security.Cryptography.DataProtectionScope]::LocalMachine))
    }
    catch { return $null }
}

function Test-TcpPort {
    param([Parameter(Mandatory)][string]$HostName, [Parameter(Mandatory)][int]$Port, [int]$TimeoutMs = 5000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync($HostName, $Port)
        if (-not $task.Wait($TimeoutMs)) { return "no answer within $([int]($TimeoutMs / 1000)) s" }
        if ($client.Connected) { return $null }
        return "not connected"
    }
    catch { return "$($_.Exception.GetBaseException().Message)" }
    finally { $client.Dispose() }
}

function Invoke-EndpointProbe {
    # One request with the same curl options as the monitor
    param([Parameter(Mandatory)][hashtable]$Config)
    $body = Join-Path $env:TEMP ("hst-health-probe-{0}.tmp" -f $PID)
    $format = 'CODE=%{http_code}\nTOTAL=%{time_total}\nREDIRECTS=%{num_redirects}\nSIZE=%{size_download}\nIP=%{remote_ip}'
    $lines = @(& curl.exe -q -s -L --max-redirs ([int]$Config.MaxRedirects) -A "CurlMonitor-Health/1.0" -H "Cache-Control: no-cache" -o $body -w $format --max-time ([int]$Config.TimeoutSeconds) $Config.Url 2>$null)
    $exit = $LASTEXITCODE
    $parsed = @{}
    foreach ($l in $lines) { if ("$l" -match '^([A-Z]+)=(.*)$') { $parsed[$Matches[1]] = $Matches[2] } }
    $marker = $true
    $size = 0
    if (Test-Path $body) {
        $size = (Get-Item $body).Length
        if ($Config.ExpectedContentMarker) { $marker = [bool](Select-String -Path $body -SimpleMatch -Pattern $Config.ExpectedContentMarker -Quiet) }
        Remove-Item $body -Force -ErrorAction SilentlyContinue
    }
    $totalMs = $null
    $sec = 0.0
    if ([double]::TryParse("$($parsed['TOTAL'])", [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$sec)) { $totalMs = [int][math]::Round($sec * 1000) }
    $ok = ($exit -eq 0 -and $parsed['CODE'] -eq '200' -and $marker -and $size -ge [int]$Config.MinPopulatedBytes)
    return [PSCustomObject]@{ Ok = $ok; Exit = $exit; Code = "$($parsed['CODE'])"; TotalMs = $totalMs; Redirects = "$($parsed['REDIRECTS'])"; Size = $size; Marker = $marker; Ip = "$($parsed['IP'])" }
}

function Read-NewDropLine {
    # Returns complete lines appended to the drops log since the last call. Starts over if the file was rotated.
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    $bytes = $null
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    try {
        if ($fs.Length -lt $script:DrillOffset) { $script:DrillOffset = 0 }
        $count = [int]($fs.Length - $script:DrillOffset)
        if ($count -le 0) { return @() }
        [void]$fs.Seek($script:DrillOffset, [IO.SeekOrigin]::Begin)
        $bytes = New-Object byte[] $count
        $read = 0
        while ($read -lt $count) { $n = $fs.Read($bytes, $read, $count - $read); if ($n -le 0) { break }; $read += $n }
    }
    finally { $fs.Dispose() }
    $lastNl = [Array]::LastIndexOf($bytes, [byte]10)
    if ($lastNl -lt 0) { return @() }
    $script:DrillOffset += $lastNl + 1
    $text = [Text.Encoding]::UTF8.GetString($bytes, 0, $lastNl + 1).TrimStart([char]0xFEFF)
    return @($text -split "\r?\n" | Where-Object { $_ })
}

function Get-CurlConfigPath {
    # The settings file curl.exe reads for the account the monitor runs as: the first of CURL_HOME, XDG_CONFIG_HOME,
    # and HOME set for the machine, otherwise that account's profile folder. Without a process, the SYSTEM profile.
    param($Process)
    foreach ($name in @('CURL_HOME', 'XDG_CONFIG_HOME', 'HOME')) {
        $v = [Environment]::GetEnvironmentVariable($name, 'Machine')
        if ($v -and (Test-Path -LiteralPath $v -PathType Container)) { return (Join-Path $v $(if ($name -eq 'XDG_CONFIG_HOME') { 'curlrc' } else { '.curlrc' })) }
    }
    $sid = 'S-1-5-18'
    if ($Process) {
        $sid = $null
        try { $sid = (Invoke-CimMethod -InputObject $Process -MethodName GetOwnerSid -ErrorAction Stop).Sid } catch { }
    }
    if (-not $sid) { return $null }
    if ($sid -eq 'S-1-5-18') { return (Join-Path $env:SystemRoot 'System32\config\systemprofile\.curlrc') }
    $userProfile = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$sid'" -ErrorAction SilentlyContinue
    if ($userProfile -and $userProfile.LocalPath) { return (Join-Path $userProfile.LocalPath '.curlrc') }
    return $null
}

function Test-DrillFile {
    param([Parameter(Mandatory)][string]$ConfigPath)
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $false }
    return ("$(Get-Content -LiteralPath $ConfigPath -TotalCount 1 -ErrorAction SilentlyContinue)".Trim([char]0xFEFF) -eq $script:DrillMarker)
}

function Get-DrillOwner {
    # Reads the owner line of a drill file. The drill is live while its process still runs, that process started no later
    # than the recorded time, and the drill is younger than the longest a drill can last.
    param([Parameter(Mandatory)][string]$ConfigPath)
    $owner = [PSCustomObject]@{ ProcessId = 0; StartedUtc = $null; Live = $false }
    if (-not (Test-DrillFile -ConfigPath $ConfigPath)) { return $owner }
    $line = "$(@(Get-Content -LiteralPath $ConfigPath -TotalCount 2 -ErrorAction SilentlyContinue)[1])"
    if ($line -notmatch '^# pid (\d+) started (\S+)$') { return $owner }
    $owner.ProcessId = [int]$Matches[1]
    $started = [datetime]::MinValue
    if (-not [datetime]::TryParse($Matches[2], [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$started)) { return $owner }
    $owner.StartedUtc = $started.ToUniversalTime()
    if (((Get-Date).ToUniversalTime() - $owner.StartedUtc).TotalMinutes -ge $script:DrillMaxMinutes) { return $owner }
    $proc = Get-Process -Id $owner.ProcessId -ErrorAction SilentlyContinue
    if (-not $proc) { return $owner }
    $procStart = $null
    try { $procStart = $proc.StartTime.ToUniversalTime() } catch { }
    $owner.Live = ($null -eq $procStart -or $procStart -le $owner.StartedUtc.AddSeconds(1))
    return $owner
}

function Add-DrillBlock {
    # Sends curl.exe requests for the HST host, under the monitor's account, to a local port that refuses connections
    param([Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][string]$Url)
    $u = [uri]$Url
    $text = "$script:DrillMarker`r`n# pid $PID started $((Get-Date).ToUniversalTime().ToString('o'))`r`nconnect-to = `"$($u.Host):$($u.Port):127.0.0.1:9`"`r`n"
    [IO.File]::WriteAllText($ConfigPath, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function Remove-DrillBlock {
    # Deletes the drill file only when it carries the drill marker and belongs to this process or to no live drill.
    # Reports whether it is gone.
    param([Parameter(Mandatory)][string]$ConfigPath)
    if (Test-DrillFile -ConfigPath $ConfigPath) {
        $owner = Get-DrillOwner -ConfigPath $ConfigPath
        if ($owner.ProcessId -eq $PID -or -not $owner.Live) { Remove-Item -LiteralPath $ConfigPath -Force -ErrorAction SilentlyContinue }
    }
    return (-not (Test-DrillFile -ConfigPath $ConfigPath))
}

function Register-DrillCleanup {
    # One-time SYSTEM task that removes the drill file if the drill never gets to, then removes itself
    param([Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][int]$Minutes, [Parameter(Mandatory)][string]$TaskPath)
    $q = { param($t) ([string]$t).Replace("'", "''") }
    $command = "`$f = '$(& $q $ConfigPath)'; if ((Test-Path -LiteralPath `$f) -and ((Get-Content -LiteralPath `$f -TotalCount 1) -eq '$(& $q $script:DrillMarker)')) { Remove-Item -LiteralPath `$f -Force }; Unregister-ScheduledTask -TaskPath '$(& $q $TaskPath)' -TaskName '$(& $q $script:DrillCleanupTaskName)' -Confirm:`$false"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes([math]::Max(1, $Minutes))
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
    Register-ScheduledTask -TaskPath $TaskPath -TaskName $script:DrillCleanupTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
}

function Unregister-DrillCleanup {
    param([Parameter(Mandatory)][string]$TaskPath)
    if (Get-ScheduledTask -TaskPath $TaskPath -TaskName $script:DrillCleanupTaskName -ErrorAction SilentlyContinue) {
        try { Unregister-ScheduledTask -TaskPath $TaskPath -TaskName $script:DrillCleanupTaskName -Confirm:$false -ErrorAction Stop } catch { }
    }
}

function Test-MailHost {
    # Connects to the mail host this monitor sends through. A missing host, or a probe that throws, is a failure:
    # the point of this section is that alerts can leave the server.
    param([Parameter(Mandatory)][hashtable]$Config)
    if ([string]::IsNullOrWhiteSpace("$($Config.SmtpServer)")) {
        Write-Check FAIL "This monitor has no mail host recorded, so alerts cannot be sent. Re-run the installer."
        return
    }
    $port = 0
    if (-not [int]::TryParse("$($Config.SmtpPort)", [ref]$port) -or $port -le 0) {
        Write-Check FAIL "This monitor has no usable mail port recorded ('$($Config.SmtpPort)'). Re-run the installer."
        return
    }
    $tcp = try { Test-TcpPort -HostName "$($Config.SmtpServer)" -Port $port } catch { "$($_.Exception.Message)" }
    if ($tcp) { Write-Check FAIL "Cannot connect to $($Config.SmtpServer):$port`: $tcp." }
    else { Write-Check PASS "Connected to $($Config.SmtpServer):$port." }
}

function Invoke-OutageDrill {
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$DropLogPath, $Process, [string]$TaskPath = '\CurlMonitor\', [int]$CleanupMinutes = 10)
    $cycle = [int]$Config.IntervalSeconds + [int]$Config.TimeoutSeconds
    $firstFailBudget = $cycle * 2 + 20
    $downBudget = $cycle * ([int]$Config.DownThreshold + 1) + 60
    $upBudget = $cycle * 2 + 60
    $alertWait = 150   # longer than the 100 s Send-MailMessage timeout, so a stalled send still reports its outcome
    $expectMail = [bool]$Config.SendEmail
    $expectResolvedMail = $expectMail -and [bool]$Config.AlertOnRecovery

    $cfgPath = Get-CurlConfigPath -Process $Process
    if (-not $cfgPath) { Write-Check FAIL "Could not tell which account the monitor runs as, so the drill did not start."; return }
    if ((Test-Path -LiteralPath $cfgPath) -and -not (Test-DrillFile -ConfigPath $cfgPath)) { Write-Check FAIL "A curl settings file already exists at '$cfgPath'. The drill does not change it, so it did not start."; return }
    $running = Get-DrillOwner -ConfigPath $cfgPath
    if ($running.Live) { Write-Check FAIL "Another outage drill, started at $($running.StartedUtc.ToLocalTime().ToString('HH:mm')) by process $($running.ProcessId), is still running. This drill did not start."; return }
    $drillHost = ([uri]$Config.Url).Host

    $script:DrillOffset = if (Test-Path $DropLogPath) { (Get-Item $DropLogPath).Length } else { 0 }
    $down = $null; $downAlert = $null; $fails = 0; $firstFail = $null; $blockedAt = $null; $downAt = $null
    try {
        try { Register-DrillCleanup -ConfigPath $cfgPath -Minutes $CleanupMinutes -TaskPath $TaskPath }
        catch { Write-Check WARN "Could not register the cleanup task, so a drill window closed early would leave the drill file until this script runs again. $($_.Exception.Message)" }
        Add-DrillBlock -ConfigPath $cfgPath -Url $Config.Url
        $blockedAt = Get-Date
        Write-Check INFO "Simulated outage started at $($blockedAt.ToString('HH:mm:ss')). Only requests to $drillHost made by curl.exe under the monitor's account are affected. Waiting up to $(Format-Age $downBudget) for DOWN."
        $nextNote = (Get-Date).AddSeconds(20)
        $downDeadline = $blockedAt.AddSeconds($downBudget)
        while ((Get-Date) -lt $downDeadline) {
            foreach ($l in Read-NewDropLine -Path $DropLogPath) {
                Write-Host "      $l" -ForegroundColor DarkGray
                if ($l -match '\| FAIL +\|') { $fails++; if (-not $firstFail) { $firstFail = Get-Date } }
                if (-not $down -and $l -match '\| DOWN +\|') { $down = $l; $downAt = Get-Date }
                if ($down -and $l -match '\| ALERT +\| (Sent|Not sent|Delivered on retry|Gave up)[^\r\n]*\[DOWN\]') { $downAlert = $l }
            }
            # Once DOWN lands, keep waiting for its send outcome even if the budget for DOWN itself has run out
            if ($down -and $downAt.AddSeconds($alertWait) -gt $downDeadline) { $downDeadline = $downAt.AddSeconds($alertWait) }
            if ($down -and ($downAlert -or -not $expectMail -or (Get-Date) -gt $downAt.AddSeconds($alertWait))) { break }
            if (-not $firstFail -and (Get-Date) -gt $blockedAt.AddSeconds($firstFailBudget)) {
                Write-Check FAIL "The monitor's polls did not start failing within $(Format-Age $firstFailBudget), so the drill cannot simulate an outage on this server. The monitor's curl.exe may read its settings from a folder other than '$(Split-Path $cfgPath -Parent)'."
                $blockedAt = $null
                break
            }
            if ((Get-Date) -ge $nextNote) { Write-Host "      waiting, $([int]((Get-Date) - $blockedAt).TotalSeconds) s so far, $fails failed poll(s) seen" -ForegroundColor DarkGray; $nextNote = (Get-Date).AddSeconds(20) }
            Start-Sleep -Seconds 2
        }
    }
    catch {
        Write-Check FAIL "The drill stopped with an error: $($_.Exception.Message)"
        $blockedAt = $null
    }
    finally {
        $gone = Remove-DrillBlock -ConfigPath $cfgPath
        Unregister-DrillCleanup -TaskPath $TaskPath
        if ($gone) { Write-Check INFO "Simulated outage ended at $((Get-Date).ToString('HH:mm:ss'))." }
        else { Write-Check FAIL "Could not delete the drill file. Delete '$cfgPath' now, or every poll keeps failing." }
    }

    if (-not $blockedAt) { return }
    if (-not $down) {
        Write-Check FAIL "The monitor did not declare DOWN within $(Format-Age $downBudget) of the simulated outage starting ($fails failed poll(s) recorded)."
        return
    }
    Write-Check PASS "DOWN declared $([int]($downAt - $blockedAt).TotalSeconds) s after the simulated outage started, after $fails failed poll(s)."
    if (-not $expectMail) { Write-Check WARN "Email alerts are turned off in this monitor, so no DOWN email was sent." }
    elseif ($downAlert -match '\| (Sent|Delivered on retry): ') { Write-Check PASS "DOWN alert email sent." }
    elseif ($downAlert) { Write-Check FAIL "The DOWN alert email was not sent. The monitor retries every minute for an hour. $($downAlert -replace '^.*?\| ALERT +\| ', '')" }
    else { Write-Check FAIL "No DOWN alert outcome was recorded within $alertWait s of DOWN." }

    $resolved = $null; $resolvedAlert = $null; $resolvedAt = $null
    $clearedAt = Get-Date
    $upDeadline = $clearedAt.AddSeconds($upBudget)
    while ((Get-Date) -lt $upDeadline) {
        foreach ($l in Read-NewDropLine -Path $DropLogPath) {
            Write-Host "      $l" -ForegroundColor DarkGray
            if (-not $resolved -and $l -match '\| RESOLVED +\|') { $resolved = $l; $resolvedAt = Get-Date }
            if ($resolved -and $l -match '\| ALERT +\| (Sent|Not sent|Delivered on retry|Gave up)[^\r\n]*\[RESOLVED\]') { $resolvedAlert = $l }
        }
        if ($resolved -and $resolvedAt.AddSeconds($alertWait) -gt $upDeadline) { $upDeadline = $resolvedAt.AddSeconds($alertWait) }
        if ($resolved -and ($resolvedAlert -or -not $expectResolvedMail -or (Get-Date) -gt $resolvedAt.AddSeconds($alertWait))) { break }
        Start-Sleep -Seconds 2
    }
    if (-not $resolved) { Write-Check FAIL "The monitor did not record RESOLVED within $(Format-Age $upBudget) of the simulated outage ending."; return }
    Write-Check PASS "RESOLVED recorded $([int]($resolvedAt - $clearedAt).TotalSeconds) s after the simulated outage ended."
    if (-not $expectResolvedMail) { Write-Check INFO "No RESOLVED email is configured for this monitor." }
    elseif ($resolvedAlert -match '\| (Sent|Delivered on retry): ') { Write-Check PASS "RESOLVED alert email sent. Check the inbox for both emails from $($Config.SiteName) ($env:COMPUTERNAME)." }
    elseif ($resolvedAlert) { Write-Check FAIL "The RESOLVED alert email was not sent. $($resolvedAlert -replace '^.*?\| ALERT +\| ', '')" }
    else { Write-Check FAIL "No RESOLVED alert outcome was recorded within $alertWait s of RESOLVED." }
}

function Invoke-MonitorCheck {
    # Runs every check against one monitor folder. Findings go to $script:Results, so the caller can sum them up.
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$TaskPath,
        [string]$TaskName = '',
        [bool]$IsAdmin = $false,
        [bool]$RunDrill = $false,
        [int]$DrillCleanupMinutes = 10
    )
    $InstallDir = $Folder
    $monitorPath = Join-Path $InstallDir 'Watch-CurlMonitor.ps1'
    if (-not (Test-Path $monitorPath)) {
        Write-Check FAIL "No monitor script at '$monitorPath'. The installer has not been run for this monitor."
        return
    }
    $cfg = Get-MonitorConfig -Path $monitorPath
    if (-not $TaskName) { $TaskName = "Curl Monitor - $($cfg.MonitorName)" }
    foreach ($k in @('IntervalSeconds', 'TimeoutSeconds', 'DownThreshold', 'SlowThresholdMs', 'MaxRedirects', 'MinPopulatedBytes')) { if ($null -eq $cfg[$k]) { $cfg[$k] = @{ IntervalSeconds = 10; TimeoutSeconds = 15; DownThreshold = 3; SlowThresholdMs = 3000; MaxRedirects = 5; MinPopulatedBytes = 1000 }[$k] } }
    $staleAfter = [int]$cfg.IntervalSeconds + [int]$cfg.TimeoutSeconds + 30
    $busyAfter = $staleAfter + 600
    Write-Check INFO "Monitor '$($cfg.MonitorName)' at site '$($cfg.SiteName)' polls $($cfg.Url) every $($cfg.IntervalSeconds) s with a $($cfg.TimeoutSeconds) s timeout, declares DOWN after $($cfg.DownThreshold) failures in a row, and alerts by $($cfg.MailMethod) to $(@($cfg.MailTo) -join ', ')."
    if ($null -eq $cfg.SlowWindowMinutes) { Write-Check WARN "This monitor was installed before slow alerts and daily summaries existed. Re-run Install-CurlMonitor.ps1 to add them." }
    else {
        $slowRule = if ($cfg.AlertOnSlow) { "SLOW email when $($cfg.SlowAlertPercent)% of polls in $($cfg.SlowWindowMinutes) min are slower than $($cfg.SlowThresholdMs) ms or fail" } else { 'slow alerts off' }
        $summaryRule = if ([int]$cfg.DailySummaryHour -ge 0) { "daily summary at $('{0:00}' -f [int]$cfg.DailySummaryHour):00 when anything went wrong" } else { 'daily summary off' }
        Write-Check INFO "Alert rules: $slowRule, $summaryRule."
    }

    Write-Section "Scheduled task and process"
    $task = $null
    try { $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop } catch { }
    $taskRunning = $false
    if (-not $task) { Write-Check FAIL "Scheduled task '$TaskPath$TaskName' was not found." }
    else {
        $taskRunning = ("$($task.State)" -eq 'Running')
        if ($taskRunning) { Write-Check PASS "Task '$TaskPath$TaskName' is Running." }
        else { Write-Check FAIL "Task '$TaskPath$TaskName' is $($task.State), not Running. Start it with: Start-ScheduledTask -TaskPath '$TaskPath' -TaskName '$TaskName'" }
        $user = "$($task.Principal.UserId)"
        if ($user -match '(^|\\)SYSTEM$') { Write-Check PASS "Task runs as SYSTEM." } else { Write-Check WARN "Task runs as '$user'. The installer sets SYSTEM." }
        if (@($task.Triggers | Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskBootTrigger' }).Count) { Write-Check PASS "Task starts at boot." }
        else { Write-Check WARN "Task has no at-startup trigger, so the monitor will not come back after a reboot." }
        $args0 = "$(@($task.Actions)[0].Arguments)"
        if ($args0 -notlike "*$monitorPath*") { Write-Check WARN "Task runs '$args0', not '$monitorPath'." }
    }
    $pattern = "*$([Management.Automation.WildcardPattern]::Escape($monitorPath))*"
    $monitorProc = $null
    $procs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like $pattern })
    if ($procs.Count -eq 0) { Write-Check FAIL "No monitor process is running$(if (-not $isAdmin) { ', or it is hidden because this console is not elevated' })." }
    else {
        if ($procs.Count -gt 1) { Write-Check WARN "$($procs.Count) monitor processes are running. There should be one." }
        $p = $procs | Sort-Object CreationDate | Select-Object -First 1
        $monitorProc = $p
        $owner = $null
        try { $owner = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop } catch { }
        $ownerText = if ($owner -and $owner.User) { "$($owner.Domain)\$($owner.User)" } else { 'an unknown account' }
        $status = if ($owner -and $owner.User -eq 'SYSTEM') { 'PASS' } else { 'WARN' }
        Write-Check $status "Monitor process $($p.ProcessId) runs as $ownerText, started $($p.CreationDate.ToString('yyyy-MM-dd HH:mm:ss')) ($(Format-Age ((Get-Date) - $p.CreationDate).TotalSeconds) ago)."
    }

    if ($isAdmin) {
        $leftover = Get-CurlConfigPath -Process $monitorProc
        $liveDrill = $false
        if ($leftover -and (Test-DrillFile -ConfigPath $leftover)) {
            $owner = Get-DrillOwner -ConfigPath $leftover
            if ($owner.Live) {
                $liveDrill = $true
                Write-Check INFO "An outage drill started at $($owner.StartedUtc.ToLocalTime().ToString('HH:mm')) by process $($owner.ProcessId) is in progress. Failing polls are expected until it ends."
            }
            elseif (Remove-DrillBlock -ConfigPath $leftover) { Write-Check WARN "A file left by an interrupted outage drill was making every poll fail. It was removed from '$leftover', and the monitor recovers on its next poll." }
            else { Write-Check FAIL "A file left by an interrupted outage drill is making every poll fail and could not be removed. Delete '$leftover'." }
        }
        if (-not $liveDrill) { Unregister-DrillCleanup -TaskPath $TaskPath }
    }

    Write-Section "Heartbeat"
    $hbPath = Join-Path $InstallDir 'heartbeat.json'
    $hb = $null
    $heartbeatFresh = $false
    if (-not (Test-Path $hbPath)) { Write-Check FAIL "No heartbeat file. The monitor has never polled on this server." }
    else {
        $age = ((Get-Date) - (Get-Item $hbPath).LastWriteTime).TotalSeconds
        $heartbeatFresh = ($age -le $staleAfter)
        if ($heartbeatFresh) { Write-Check PASS "Heartbeat written $([int]$age) s ago." }
        elseif ($monitorProc -and $age -le $busyAfter) { Write-Check WARN "Last heartbeat was $(Format-Age $age) ago, but the monitor process is running. It is probably waiting on an alert send, which can take up to 100 s. Run this check again in a minute." }
        else { Write-Check FAIL "Last heartbeat was $(Format-Age $age) ago. The monitor is not polling." }
        try {
            # Opened with delete sharing so this read never blocks the monitor replacing the file
            $fs = [IO.File]::Open($hbPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
            try { $hb = (New-Object IO.StreamReader($fs)).ReadToEnd() | ConvertFrom-Json } finally { $fs.Dispose() }
        }
        catch { Write-Check WARN "The heartbeat file is unreadable. $($_.Exception.Message)" }
        if ($hb) {
            if ($hb.IsDown) { Write-Check WARN "'$($cfg.MonitorName)' at $($cfg.SiteName) is DOWN right now, since $($hb.OutageStartLocalStr), after $($hb.ConsecutiveFailures) failed polls in a row." }
            elseif ($hb.IsSlow) { Write-Check WARN "'$($cfg.MonitorName)' at $($cfg.SiteName) is SLOW right now, since $($hb.SlowStartLocalStr), worst response $($hb.SlowWorstMs) ms." }
            elseif ([int]$hb.ConsecutiveFailures -gt 0) { Write-Check INFO "The last $($hb.ConsecutiveFailures) poll(s) failed. DOWN is declared at $($cfg.DownThreshold)." }
            if ($hb.Stopped -and -not $heartbeatFresh) { Write-Check WARN "The heartbeat is marked as stopped by the installer, and no monitor has polled since." }
        }
    }

    Write-Section "Polls recorded"
    $day = @(Get-RecentPoll -Folder $InstallDir -Hours 24 -IntervalSeconds ([int]$cfg.IntervalSeconds))
    $lastPollOk = $null
    if ($day.Count -eq 0) { Write-Check FAIL "No polls recorded in the last 24 hours in Latency_*.csv." }
    else {
        $last = $day[-1]
        $lastPollOk = -not $last.Failed
        $lastAge = ((Get-Date) - $last.When).TotalSeconds
        $lastText = "HTTP $($last.HttpCode), $(if ($last.TotalMs) { "$($last.TotalMs) ms" } else { 'no timing' }), $($last.Reason)"
        if ($lastAge -le $staleAfter) { Write-Check PASS "Last poll recorded $([int]$lastAge) s ago: $lastText." }
        elseif ($monitorProc -and $lastAge -le $busyAfter) { Write-Check WARN "Last poll was recorded $(Format-Age $lastAge) ago, at $($last.Timestamp_Local), but the monitor process is running. It is probably waiting on an alert send, which can take up to 100 s. Run this check again in a minute." }
        else { Write-Check FAIL "Last poll was recorded $(Format-Age $lastAge) ago, at $($last.Timestamp_Local): $lastText." }
        $hourAgo = (Get-Date).AddHours(-1)
        $slowMs = [int]$cfg.SlowThresholdMs
        foreach ($set in @(@{ Label = 'Last hour'; Rows = @($day.Where({ $_.When -ge $hourAgo })) }, @{ Label = 'Last 24 hours'; Rows = $day })) {
            $rows = @($set.Rows)
            if ($rows.Count -eq 0) { Write-Check INFO "$($set.Label): no polls."; continue }
            $failedCount = 0; $slow = 0
            $msList = New-Object System.Collections.Generic.List[int]
            foreach ($r in $rows) { if ($r.Failed) { $failedCount++ } else { $v = [int]$r.TotalMs; $msList.Add($v); if ($v -ge $slowMs) { $slow++ } } }
            $msList.Sort()
            $ms = $msList.ToArray()
            $timing = if ($ms.Count) { " Response time median $(Get-Percentile $ms 50) ms, 95th percentile $(Get-Percentile $ms 95) ms, worst $($ms[-1]) ms." } else { '' }
            Write-Check INFO "$($set.Label): $($rows.Count) polls, $failedCount failed, $slow slower than $slowMs ms.$timing"
        }
        $gaps = New-Object System.Collections.ArrayList
        for ($i = 1; $i -lt $day.Count; $i++) {
            $s = ($day[$i].When - $day[$i - 1].When).TotalSeconds
            if ($s -gt $staleAfter) { [void]$gaps.Add([PSCustomObject]@{ From = $day[$i - 1].When; To = $day[$i].When; Seconds = $s }) }
        }
        if ($gaps.Count -eq 0) { Write-Check PASS "Polls recorded without gaps since $($day[0].When.ToString('yyyy-MM-dd HH:mm'))." }
        else {
            $g = $gaps | Sort-Object Seconds -Descending | Select-Object -First 1
            Write-Check WARN "$($gaps.Count) gap(s) with no polls recorded in the last 24 hours, the longest $(Format-Age $g.Seconds) from $($g.From.ToString('HH:mm:ss')) to $($g.To.ToString('HH:mm:ss')). A gap means the monitor was stopped, the server was down, or the CSV was open in another program."
        }
        $problems = @(@($day.Where({ $_.Failed -or [int]"0$($_.TotalMs)" -ge $slowMs })) | Select-Object -Last 10)
        if ($problems.Count) {
            Write-Host "      Latest failed or slow polls:" -ForegroundColor DarkGray
            foreach ($r in $problems) { Write-Host ("      {0}  HTTP {1}  {2,6} ms  {3}" -f $r.Timestamp_Local, $r.HttpCode, $r.TotalMs, $r.Reason) -ForegroundColor DarkGray }
        }
    }

    Write-Section "Monitor logs"
    $dropPath = Join-Path $InstallDir 'Drops.log'
    if (-not (Test-Path $dropPath)) { Write-Check WARN "No drops log. The monitor writes one when it starts, so this monitor predates it or has never started." }
    else {
        $entries = @(Get-Content -Path $dropPath -Tail 3000 -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_ -match '^\uFEFF?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d) \| (\w+) +\| (.*)$') { [PSCustomObject]@{ When = (ConvertFrom-LocalStamp $Matches[1]); Kind = $Matches[2]; Text = $Matches[3] } }
            } | Where-Object { $_.When })
        $since24 = (Get-Date).AddHours(-24); $since7 = (Get-Date).AddDays(-7)
        $recent = @($entries | Where-Object { $_.When -ge $since24 })
        $counts = foreach ($k in @('FAIL', 'SLOW', 'SLOWSTART', 'DOWN', 'RESOLVED', 'RESTART', 'ERROR')) { "$(@($recent | Where-Object { $_.Kind -eq $k }).Count) $k" }
        Write-Check INFO "Drops log, last 24 hours: $($counts -join ', ')."
        $start = $entries | Where-Object { $_.Kind -eq 'START' } | Select-Object -Last 1
        if ($start) { Write-Check INFO "Monitor last started $($start.When.ToString('yyyy-MM-dd HH:mm:ss'))." }
        $restart = $entries | Where-Object { $_.Kind -eq 'RESTART' -and $_.When -ge $since7 } | Select-Object -Last 1
        if ($restart) { Write-Check WARN "The monitor was not running for a while before $($restart.When.ToString('yyyy-MM-dd HH:mm')): $($restart.Text)" }
        $undelivered = $entries | Where-Object { $_.Kind -eq 'ALERT' -and $_.When -ge $since7 -and $_.Text -match '^(Not sent|Gave up)' } | Select-Object -Last 1
        if ($undelivered) { Write-Check WARN "An alert could not be sent in the last 7 days, at $($undelivered.When.ToString('yyyy-MM-dd HH:mm')): $($undelivered.Text)" }
        $err = $recent | Where-Object { $_.Kind -in @('ERROR', 'LOST') } | Select-Object -Last 1
        if ($err) { Write-Check WARN "The monitor logged a problem at $($err.When.ToString('HH:mm:ss')): $($err.Text)" }
    }
    $transcripts = @(Get-ChildItem -Path $InstallDir -Filter 'Transcript_*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge (Get-Date).AddHours(-24) })
    if ($transcripts.Count) {
        $noise = 'FAILED \| Site=|WARNING \| Slow response\.'
        $issues = @(Select-String -Path ($transcripts | Sort-Object Name | ForEach-Object FullName) -Pattern '^\d\d\D\d\d\D\d\d - \d\d\D\d\d\D\d\d.{0,12}? (WARNING|FAILED|ERROR) \| ' -ErrorAction SilentlyContinue | Where-Object { $_.Line -notmatch $noise })
        if ($issues.Count -eq 0) { Write-Check PASS "No warnings or errors in the monitor's own log in the last 24 hours, apart from poll results." }
        else {
            Write-Check WARN "$($issues.Count) warning or error line(s) in the monitor's own log in the last 24 hours. Latest:"
            foreach ($m in ($issues | Select-Object -Last 5)) { Write-Host "      $($m.Line)" -ForegroundColor DarkGray }
        }
    }

    Write-Section "Alert delivery"
    if ("$($cfg.MailMethod)" -eq 'None') { Write-Check INFO "This monitor was installed without alerts. Nothing is emailed, and outages, slow periods and every poll are still recorded." }
    elseif (-not $cfg.SendEmail) { Write-Check WARN "Mail is configured but sending is switched off in this monitor, so no alert leaves the server." }
    $credPath = "$($cfg.SmtpCredentialFile)"
    switch ("$($cfg.MailMethod)") {
        'None' { }
        'Graph' {
            $exp = [datetime]::MinValue
            if ([datetime]::TryParseExact("$($cfg.GraphSecretExpires)", 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$exp)) {
                $days = [int][math]::Floor(($exp - (Get-Date).Date).TotalDays)
                if ($days -lt 0) { Write-Check FAIL "The client secret expired on $($cfg.GraphSecretExpires)." }
                elseif ($days -le 30) { Write-Check WARN "The client secret expires on $($cfg.GraphSecretExpires), in $days day(s)." }
                else { Write-Check PASS "The client secret expires on $($cfg.GraphSecretExpires), in $days days." }
            }
            if (-not $isAdmin) { Write-Check WARN "The stored client secret can only be tested from an elevated console." }
            else {
                $secret = Get-StoredMailSecret -Path $credPath
                if (-not $secret) { Write-Check FAIL "The stored client secret at '$credPath' is missing or cannot be decrypted. Re-run the installer." }
                else {
                    try {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
                        $tok = Invoke-RestMethod -Method Post -Uri "$script:LoginBase/$($cfg.GraphTenantId)/oauth2/v2.0/token" -Body @{ client_id = $cfg.GraphClientId; client_secret = $secret; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30 -ErrorAction Stop
                        if ($tok.access_token) { Write-Check PASS "Microsoft sign-in accepted the stored client secret for app $($cfg.GraphClientId)." }
                        else { Write-Check FAIL "Microsoft sign-in answered without a token." }
                    }
                    catch { Write-Check FAIL "Microsoft sign-in did not issue a token with the stored client secret. $(Get-RestErrorDetail $_)" }
                    finally { $secret = $null }
                }
            }
            Write-Check INFO "Permission to send as $($cfg.MailFrom) is proven only by a real email. Run with -OutageDrill to send one through the monitor itself."
        }
        'Authenticated' {
            if (-not $isAdmin) { Write-Check WARN "The stored SMTP password can only be checked from an elevated console." }
            elseif (-not (Get-StoredMailSecret -Path $credPath)) { Write-Check FAIL "The stored SMTP password at '$credPath' is missing or cannot be decrypted. Re-run the installer." }
            else { Write-Check PASS "The stored SMTP password decrypts." }
            Test-MailHost -Config $cfg
        }
        default {
            Test-MailHost -Config $cfg
        }
    }

    Write-Section "Endpoint from this server right now"
    $probe = Invoke-EndpointProbe -Config $cfg
    if ($probe.Ok) { Write-Check PASS "Reached the sign-in page: HTTP 200 after $($probe.Redirects) redirect(s) in $($probe.TotalMs) ms, $($probe.Size) bytes, from $($probe.Ip)." }
    else {
        $why = if ($probe.Exit -ne 0) { "curl exit $($probe.Exit)" } elseif ($probe.Code -ne '200') { "HTTP $($probe.Code)" } else { "page not populated ($($probe.Size) bytes, marker $(if ($probe.Marker) { 'found' } else { 'missing' }))" }
        Write-Check WARN "This probe failed: $why. One failure does not alert. The monitor declares DOWN after $($cfg.DownThreshold) in a row."
    }
    if ($null -ne $lastPollOk -and $probe.Ok -ne $lastPollOk) { Write-Check INFO "The monitor's last poll $(if ($lastPollOk) { 'succeeded' } else { 'failed' }) and this probe $(if ($probe.Ok) { 'succeeded' } else { 'failed' }). A single difference is normal when the site is flapping." }

    if ($RunDrill) {
        Write-Section "Outage drill"
        if (-not $isAdmin) { Write-Check FAIL "The outage drill needs an elevated console." }
        elseif (-not ($taskRunning -and $heartbeatFresh -and $monitorProc)) { Write-Check FAIL "The drill was not started because the monitor is not running and polling." }
        elseif ($hb -and $hb.IsDown) { Write-Check FAIL "The drill was not started because the monitor already considers the site DOWN." }
        else { Invoke-OutageDrill -Config $cfg -DropLogPath $dropPath -Process $monitorProc -TaskPath $TaskPath -CleanupMinutes $DrillCleanupMinutes }
    }
}


$isAdmin = Test-IsAdministrator
Write-Host ""
Write-Host "curl monitor health check on $env:COMPUTERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
if (-not $isAdmin) { Write-Check WARN "Not running as administrator. The task, the process owner, and the mail secret can only be checked from an elevated console." }

function Get-MonitorFolder {
    # Monitor folders under the root, newest install layout only: a folder with a generated monitor in it
    param([Parameter(Mandatory)][string]$Root)
    if (-not (Test-Path $Root)) { return @() }
    return @(Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName 'Watch-CurlMonitor.ps1') } |
        ForEach-Object {
            $name = $_.Name
            $settings = Join-Path $_.FullName 'install-settings.json'
            if (Test-Path $settings) {
                try { $j = Get-Content -Path $settings -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop; if ($j.MonitorName) { $name = [string]$j.MonitorName } } catch { }
            }
            [PSCustomObject]@{ Name = $name; Slug = $_.Name; Path = $_.FullName }
        })
}

$folders = @()
if ($InstallDir) { $folders = @([PSCustomObject]@{ Name = (Split-Path $InstallDir -Leaf); Slug = (Split-Path $InstallDir -Leaf); Path = $InstallDir }) }
else {
    $folders = @(Get-MonitorFolder -Root $InstallRoot)
    if ($Monitor) {
        $wanted = @($folders | Where-Object { $_.Name -eq $Monitor -or $_.Slug -eq $Monitor })
        if (-not $wanted.Count) {
            Write-Check FAIL "No monitor called '$Monitor' under '$InstallRoot'. Installed: $(if ($folders.Count) { (@($folders | ForEach-Object { $_.Name }) -join ', ') } else { 'none' })."
            exit 1
        }
        $folders = $wanted
    }
}
if (-not $folders.Count) {
    # A server set up before the parent folder was dropped still has its monitors one level deeper
    $older = @(Get-MonitorFolder -Root $PreviousRoot)
    if ($older.Count) {
        Write-Check FAIL "No monitor under '$InstallRoot', but $($older.Count) is installed in the old location '$PreviousRoot': $((@($older | ForEach-Object { $_.Name })) -join ', '). Run Install-CurlMonitor.ps1 and let it move them, or check that one with -InstallDir."
    }
    else {
        Write-Check FAIL "No monitor found under '$InstallRoot'. Run Install-CurlMonitor.ps1 on this server first."
    }
    exit 1
}
if ($OutageDrill -and $folders.Count -gt 1) {
    Write-Check FAIL "The outage drill needs -Monitor when more than one monitor is installed: $((@($folders | ForEach-Object { $_.Name })) -join ', ')."
    exit 1
}
foreach ($f in $folders) {
    if ($folders.Count -gt 1) {
        Write-Host ""
        Write-Host "===== $($f.Name) ($($f.Path))" -ForegroundColor Cyan
    }
    Invoke-MonitorCheck -Folder $f.Path -TaskPath $TaskPath -TaskName $TaskName -IsAdmin $isAdmin -RunDrill ([bool]$OutageDrill) -DrillCleanupMinutes $DrillCleanupMinutes
}

$failCount = @($script:Results | Where-Object { $_.Status -eq 'FAIL' }).Count
$warnCount = @($script:Results | Where-Object { $_.Status -eq 'WARN' }).Count
Write-Host ""
if ($failCount) { Write-Host "Result: NOT WORKING CORRECTLY. $failCount check(s) failed, $warnCount warning(s)." -ForegroundColor Red; exit 1 }
if ($warnCount) { Write-Host "Result: WORKING, with $warnCount warning(s) to review." -ForegroundColor Yellow; exit 2 }
Write-Host "Result: WORKING. Every check passed." -ForegroundColor Green
exit 0
