<#
.SYNOPSIS
    Summarises what the curl monitors on this machine measured and commits the result to a git working copy.

.DESCRIPTION
    Reads every monitor under the curl-monitor install root, takes the polls recorded since the last published
    run, and writes three things into a git working copy: an append-only CSV per endpoint, a dated report, and a
    refreshed table in the repository README between marker comments. It then commits and pushes.

    Endpoints appear in the published data as a code (ENDPOINT-01 upward) and the first 12 characters of the
    SHA-256 of the URL. The code map lives beside the state file on this machine and is never published, so the
    data says what an endpoint did without saying what it is.

.PARAMETER RepoPath
    Working copy to publish into.

.PARAMETER MonitorRoot
    Root holding one folder per monitor. Default C:\ProgramData\CurlMonitor.

.PARAMETER StatePath
    Folder holding publish-state.json and endpoint-map.json. Never published. Default C:\ProgramData\CurlMonitor-telemetry.

.PARAMETER Window
    Morning, Evening, or Auto. Auto picks Morning before noon.

.PARAMETER DryRun
    Computes the window and writes the files, then stops without touching git.

.EXAMPLE
    .\Publish-UptimeTelemetry.ps1 -RepoPath C:\ProgramData\CurlMonitor-telemetry\curl-monitor

.NOTES
    Author:      Christopher Carroll
    Created:     09/21/2026
    Exit code:   0 published or nothing to publish, 1 a step failed.
#>
[CmdletBinding()]
param(
    [string]$RepoPath = 'C:\ProgramData\CurlMonitor-telemetry\curl-monitor',
    [string]$MonitorRoot = 'C:\ProgramData\CurlMonitor',
    [string]$StatePath = 'C:\ProgramData\CurlMonitor-telemetry',
    [ValidateSet('Morning', 'Evening', 'Auto')][string]$Window = 'Auto',
    [switch]$DryRun,
    [string]$LogPath = ''
)
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'never'
$env:GIT_ASKPASS = ''
if (-not $LogPath) { $LogPath = Join-Path $StatePath ("publisher_" + (Get-Date -Format 'yyyyMM') + ".log") }

$ReadmeStartMarker = '<!-- telemetry:start -->'
$ReadmeEndMarker = '<!-- telemetry:end -->'
$DataColumns = @('WindowEnd_Local', 'Window', 'Endpoint', 'UrlHash', 'Polls', 'AvailabilityPercent', 'FailedPolls', 'SlowPolls', 'P50Ms', 'P95Ms', 'MaxMs', 'Outages', 'LongestOutageSeconds', 'TotalOutageSeconds', 'SlowPeriods', 'BackendAddresses', 'FailureReasons')

function Write-Line {
    param([Parameter(Mandatory)][ValidateSet('INFO', 'DONE', 'WARN', 'FAIL')][string]$Level, [Parameter(Mandatory)][string]$Message)
    $color = @{ INFO = 'Gray'; DONE = 'Green'; WARN = 'Yellow'; FAIL = 'Red' }[$Level]
    $line = "{0} {1,-4} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $color
    try {
        $folder = Split-Path -Path $LogPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch { }
}

function ConvertTo-UrlHash {
    # First 12 hex characters of the SHA-256 of the URL, so a published row can be tied to an endpoint without naming it
    param([Parameter(Mandatory)][string]$Url)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return (-join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Url.Trim())) | ForEach-Object { $_.ToString('x2') })).Substring(0, 12) }
    finally { $sha.Dispose() }
}

function Get-EndpointCode {
    # Stable code per URL hash. The map is a local file, never published, so codes survive reinstalls and reorders.
    param([Parameter(Mandatory)][hashtable]$Map, [Parameter(Mandatory)][string]$UrlHash)
    if ($Map.ContainsKey($UrlHash)) { return $Map[$UrlHash] }
    $next = 1 + @($Map.Values | ForEach-Object { if ("$_" -match 'ENDPOINT-(\d+)$') { [int]$Matches[1] } else { 0 } } | Measure-Object -Maximum).Maximum
    $code = 'ENDPOINT-{0:00}' -f $next
    $Map[$UrlHash] = $code
    return $code
}

function ConvertFrom-LocalStamp {
    param([string]$Text)
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact("$Text".Trim([char]0xFEFF, ' '), 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Get-Percentile {
    # Nearest-rank percentile over a sorted array
    param([int[]]$Sorted, [double]$Percent)
    if (-not $Sorted -or $Sorted.Count -eq 0) { return $null }
    $i = [int][math]::Ceiling($Percent / 100 * $Sorted.Count) - 1
    return $Sorted[[math]::Min($Sorted.Count - 1, [math]::Max(0, $i))]
}

function Get-MonitorSetting {
    # Values the installer baked into a deployed monitor, read without running any of it
    param([Parameter(Mandatory)][string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $cfg = @{}
    foreach ($st in $ast.EndBlock.Statements) {
        if ($st -is [System.Management.Automation.Language.FunctionDefinitionAst]) { break }
        if ($st -isnot [System.Management.Automation.Language.AssignmentStatementAst]) { continue }
        if ($st.Left -isnot [System.Management.Automation.Language.VariableExpressionAst]) { continue }
        try { $cfg[$st.Left.VariablePath.UserPath] = $st.Right.Expression.SafeGetValue() } catch { }
    }
    return $cfg
}

function Get-TelemetryRow {
    # Latency rows inside the window, read from the monthly files the window can touch
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][datetime]$From, [Parameter(Mandatory)][datetime]$To)
    $rows = New-Object System.Collections.ArrayList
    $months = @($From.ToString('yyyyMM'), $To.ToString('yyyyMM')) | Sort-Object -Unique
    foreach ($m in $months) {
        $file = Join-Path $Folder "Latency_$m.csv"
        if (-not (Test-Path -LiteralPath $file)) { continue }
        foreach ($r in (Import-Csv -LiteralPath $file)) {
            $t = ConvertFrom-LocalStamp $r.Timestamp_Local
            if (-not $t -or $t -le $From -or $t -gt $To) { continue }
            [void]$rows.Add($r)
        }
    }
    return @($rows)
}

function Get-TelemetryStat {
    # Pure. Turns latency rows, outage rows, and drops lines from one window into the published figures.
    param(
        [object[]]$Rows,
        [object[]]$Outages,
        [string[]]$DropLines,
        [Parameter(Mandatory)][datetime]$From,
        [Parameter(Mandatory)][datetime]$To,
        [int]$SlowThresholdMs = 3000
    )
    $ok = New-Object System.Collections.Generic.List[int]
    $failed = 0; $slow = 0
    $reasons = @{}; $backends = @{}
    foreach ($r in @($Rows)) {
        $good = ($r.HttpCode -eq '200' -and $r.ContentOk -eq 'True' -and -not [string]::IsNullOrEmpty($r.TotalMs))
        if ($good) {
            $ms = [int]$r.TotalMs
            $ok.Add($ms)
            if ($ms -ge $SlowThresholdMs) { $slow++ }
            if ($r.RemoteIp) { $backends[[string]$r.RemoteIp] = $true }
        }
        else {
            $failed++
            $reason = if ($r.Reason) { [string]$r.Reason } else { 'unknown' }
            $reasons[$reason] = 1 + [int]$reasons[$reason]
        }
    }
    $polls = @($Rows).Count
    $ok.Sort()
    $sorted = $ok.ToArray()
    $outageCount = 0; $longest = 0; $total = 0
    foreach ($o in @($Outages)) {
        $start = ConvertFrom-LocalStamp $o.OutageStart_Local
        if (-not $start -or $start -le $From -or $start -gt $To) { continue }
        $outageCount++
        $seconds = 0
        [void][int]::TryParse("$($o.DurationSeconds)", [ref]$seconds)
        $total += $seconds
        if ($seconds -gt $longest) { $longest = $seconds }
    }
    $slowPeriods = 0
    foreach ($line in @($DropLines)) {
        if ("$line" -notmatch '^\uFEFF?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d) \| (\w+) +\|') { continue }
        if ($Matches[2] -ne 'SLOWSTART') { continue }
        $t = ConvertFrom-LocalStamp $Matches[1]
        if ($t -and $t -gt $From -and $t -le $To) { $slowPeriods++ }
    }
    $availability = if ($polls -gt 0) { [math]::Round(100 * ($polls - $failed) / $polls, 2) } else { 0 }
    [PSCustomObject]@{
        Polls                = $polls
        AvailabilityPercent  = $availability
        FailedPolls          = $failed
        SlowPolls            = $slow
        P50Ms                = (Get-Percentile $sorted 50)
        P95Ms                = (Get-Percentile $sorted 95)
        MaxMs                = $(if ($sorted.Count) { $sorted[-1] } else { $null })
        Outages              = $outageCount
        LongestOutageSeconds = $longest
        TotalOutageSeconds   = $total
        SlowPeriods          = $slowPeriods
        BackendAddresses     = $backends.Keys.Count
        FailureReasons       = (@($reasons.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key) x$($_.Value)" }) -join '; ')
    }
}

function Get-DropLine {
    # Current drops log plus any rotated file whose last write falls inside the window
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][datetime]$From)
    $lines = New-Object System.Collections.ArrayList
    foreach ($f in @(Get-ChildItem -LiteralPath $Folder -Filter 'Drops*.log' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Drops.log' -or $_.LastWriteTime -ge $From })) {
        foreach ($l in @(Get-Content -LiteralPath $f.FullName -ErrorAction SilentlyContinue)) { [void]$lines.Add($l) }
    }
    return @($lines)
}

function New-CommitSubject {
    param([Parameter(Mandatory)][datetime]$Date, [Parameter(Mandatory)][string]$Window, [Parameter(Mandatory)][double]$Availability)
    return "Feature Improvement: telemetry reporting publisher, data sampling $($Date.ToString('yyyy-MM-dd')) $($Window.ToLower()), availability $('{0:N2}' -f $Availability)%"
}

function New-CommitBody {
    param([Parameter(Mandatory)][object[]]$Measurements)
    $Measurements = @($Measurements | Where-Object { $null -ne $_ })
    $lines = foreach ($m in $Measurements) {
        $outage = if ($m.Stats.Outages -gt 0) { ", $($m.Stats.Outages) outage$(if ($m.Stats.Outages -eq 1) { '' } else { 's' }) totalling $($m.Stats.TotalOutageSeconds)s" } else { '' }
        "$($m.Code): $($m.Stats.Polls) polls, $('{0:N2}' -f $m.Stats.AvailabilityPercent)% available, p95 $(if ($null -ne $m.Stats.P95Ms) { "$($m.Stats.P95Ms) ms" } else { 'no successful polls' })$outage"
    }
    return ($lines -join "`n")
}

function New-TelemetryReport {
    # The dated report, rebuilt from every row already published for that date so a second window keeps the first
    param([Parameter(Mandatory)][datetime]$Date, [Parameter(Mandatory)][object[]]$Rows)
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine("# Endpoint availability, $($Date.ToString('yyyy-MM-dd'))")
    [void]$sb.AppendLine()
    foreach ($window in @('Morning', 'Evening')) {
        $all = @($Rows | Where-Object { $_.Window -eq $window })
        if ($all.Count -eq 0) { continue }
        # Every row published for this window, oldest first per endpoint: a window published twice shows both
        $set = @($all | Sort-Object Endpoint, WindowEnd_Local)
        $measuredTo = @(@($all | ForEach-Object { "$($_.WindowEnd_Local)" }) | Sort-Object)[-1]
        [void]$sb.AppendLine("## $window window, measured to $measuredTo")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine('| Endpoint | Polls | Available | p50 | p95 | Max | Failed | Slow | Outages |')
        [void]$sb.AppendLine('|---|---:|---:|---:|---:|---:|---:|---:|---:|')
        foreach ($r in $set) {
            $p50 = if ($r.P50Ms) { "$($r.P50Ms) ms" } else { 'n/a' }
            $p95 = if ($r.P95Ms) { "$($r.P95Ms) ms" } else { 'n/a' }
            $max = if ($r.MaxMs) { "$($r.MaxMs) ms" } else { 'n/a' }
            [void]$sb.AppendLine("| $($r.Endpoint) | $($r.Polls) | $($r.AvailabilityPercent)% | $p50 | $p95 | $max | $($r.FailedPolls) | $($r.SlowPolls) | $($r.Outages) |")
        }
        [void]$sb.AppendLine()
        foreach ($r in $set) {
            if ($r.FailureReasons) { [void]$sb.AppendLine("$($r.Endpoint) failures: $($r.FailureReasons).") }
            if ([int]$r.Outages -gt 0) { [void]$sb.AppendLine("$($r.Endpoint) was unreachable for $($r.TotalOutageSeconds)s in $($r.Outages) outage(s), longest $($r.LongestOutageSeconds)s.") }
            if ([int]$r.SlowPeriods -gt 0) { [void]$sb.AppendLine("$($r.Endpoint) had $($r.SlowPeriods) sustained slow period(s).") }
        }
        [void]$sb.AppendLine()
    }
    [void]$sb.AppendLine('Measured by curl-monitor. Endpoints are published as codes; the code map is kept off the repository.')
    return $sb.ToString()
}

function Update-ReadmeTable {
    # Replaces the block between the markers. Returns the new text, or the original when the markers are missing.
    param([Parameter(Mandatory)][string]$Text, [Parameter(Mandatory)][string]$Table, [string]$StartMarker = $ReadmeStartMarker, [string]$EndMarker = $ReadmeEndMarker)
    $start = $Text.IndexOf($StartMarker)
    $end = $Text.IndexOf($EndMarker)
    if ($start -lt 0 -or $end -lt $start) { return $Text }
    return $Text.Substring(0, $start + $StartMarker.Length) + "`n" + $Table.TrimEnd() + "`n" + $Text.Substring($end)
}

function New-ReadmeTable {
    param([Parameter(Mandatory)][object[]]$Rows, [Parameter(Mandatory)][datetime]$AsOf)
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine("Last 24 hours, measured to $($AsOf.ToString('yyyy-MM-dd HH:mm')) local.")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| Endpoint | Polls | Available | p50 | p95 | Outages |')
    [void]$sb.AppendLine('|---|---:|---:|---:|---:|---:|')
    foreach ($r in @($Rows | Sort-Object Endpoint)) {
        $p50 = if ($r.P50Ms) { "$($r.P50Ms) ms" } else { 'n/a' }
        $p95 = if ($r.P95Ms) { "$($r.P95Ms) ms" } else { 'n/a' }
        [void]$sb.AppendLine("| $($r.Endpoint) | $($r.Polls) | $($r.AvailabilityPercent)% | $p50 | $p95 | $($r.Outages) |")
    }
    return $sb.ToString()
}

function Add-DataRow {
    # Appends one row per endpoint, writing the header exactly once per monthly file
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Row)
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    $ordered = [ordered]@{}
    foreach ($c in $DataColumns) { $ordered[$c] = $Row.$c }
    [PSCustomObject]$ordered | Export-Csv -LiteralPath $Path -NoTypeInformation -Append
}

function Invoke-Git {
    param([Parameter(Mandatory)][string]$RepoPath, [Parameter(Mandatory)][string[]]$Arguments)
    $out = & git.exe -C $RepoPath @Arguments 2>&1
    return [PSCustomObject]@{ ExitCode = $LASTEXITCODE; Output = ($out | ForEach-Object { "$_" }) -join "`n" }
}

function New-OutputSnapshot {
    # Copies aside only the paths this job writes, so they can be put back without touching anything else
    param([Parameter(Mandatory)][string]$RepoPath, [Parameter(Mandatory)][string[]]$Paths)
    $root = Join-Path ([IO.Path]::GetTempPath()) ('curlmon-publish-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -Path $root -ItemType Directory -Force | Out-Null
    $items = New-Object System.Collections.ArrayList
    foreach ($rel in $Paths) {
        $live = Join-Path $RepoPath $rel
        $existed = Test-Path -LiteralPath $live
        [void]$items.Add([PSCustomObject]@{ Rel = $rel; Existed = $existed })
        if (-not $existed) { continue }
        $saved = Join-Path $root $rel
        $parent = Split-Path -Path $saved -Parent
        if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
        Copy-Item -LiteralPath $live -Destination $saved -Recurse -Force -ErrorAction SilentlyContinue
    }
    return [PSCustomObject]@{ Root = $root; Items = @($items); RepoPath = $RepoPath }
}

function Restore-OutputSnapshot {
    # Puts the snapshotted paths back exactly as they were, and removes anything this run created
    param([Parameter(Mandatory)][object]$Snapshot)
    $ok = $true
    foreach ($i in @($Snapshot.Items)) {
        $live = Join-Path $Snapshot.RepoPath $i.Rel
        try {
            if (Test-Path -LiteralPath $live) { Remove-Item -LiteralPath $live -Recurse -Force -ErrorAction Stop }
            if ($i.Existed) {
                $saved = Join-Path $Snapshot.Root $i.Rel
                $parent = Split-Path -Path $live -Parent
                if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
                Copy-Item -LiteralPath $saved -Destination $live -Recurse -Force -ErrorAction Stop
            }
        }
        catch { $ok = $false }
    }
    Remove-Item -LiteralPath $Snapshot.Root -Recurse -Force -ErrorAction SilentlyContinue
    return $ok
}

function Test-RepoMidRebase {
    # Both rebase backends, and a conflicted autostash reapply, leave the working copy unusable for later runs
    param([Parameter(Mandatory)][string]$RepoPath)
    if (Test-Path -LiteralPath (Join-Path $RepoPath '.git\rebase-merge')) { return $true }
    if (Test-Path -LiteralPath (Join-Path $RepoPath '.git\rebase-apply')) { return $true }
    $unmerged = @(& git.exe -C $RepoPath ls-files -u 2>$null)
    return (@($unmerged).Count -gt 0)
}

function Publish-Window {
    # One run: measure every monitor, write the files, then commit and push unless -DryRun was given
    param(
        [Parameter(Mandatory)][string]$RepoPath,
        [Parameter(Mandatory)][string]$MonitorRoot,
        [Parameter(Mandatory)][string]$StatePath,
        [Parameter(Mandatory)][string]$Window,
        [Parameter(Mandatory)][datetime]$Now,
        [switch]$DryRun
    )
    if (-not (Test-Path -LiteralPath $RepoPath)) { Write-Line FAIL "No working copy at '$RepoPath'."; return 1 }
    if (-not (Test-Path -LiteralPath $MonitorRoot)) { Write-Line FAIL "No monitors at '$MonitorRoot'."; return 1 }
    if (-not (Test-Path -LiteralPath $StatePath)) { New-Item -Path $StatePath -ItemType Directory -Force | Out-Null }

    $statefile = Join-Path $StatePath 'publish-state.json'
    $mapfile = Join-Path $StatePath 'endpoint-map.json'
    $from = $Now.AddHours(-12)
    if (Test-Path -LiteralPath $statefile) {
        try {
            $state = Get-Content -LiteralPath $statefile -Raw | ConvertFrom-Json
            $parsed = ConvertFrom-LocalStamp $state.LastWindowEnd
            if ($parsed -and $parsed -lt $Now) { $from = $parsed }
        }
        catch { Write-Line WARN "Could not read '$statefile'. Falling back to the last 12 hours." }
    }
    $map = @{}
    if (Test-Path -LiteralPath $mapfile) {
        try { (Get-Content -LiteralPath $mapfile -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $map[$_.Name] = [string]$_.Value } } catch { }
    }

    $measurements = New-Object System.Collections.ArrayList
    foreach ($folder in @(Get-ChildItem -LiteralPath $MonitorRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $settingsFile = Join-Path $folder.FullName 'install-settings.json'
        $monitorFile = Join-Path $folder.FullName 'Watch-CurlMonitor.ps1'
        if (-not (Test-Path -LiteralPath $settingsFile) -or -not (Test-Path -LiteralPath $monitorFile)) { continue }
        $url = ''
        try { $url = [string](Get-Content -LiteralPath $settingsFile -Raw | ConvertFrom-Json).Url } catch { }
        if (-not $url) { Write-Line WARN "Skipping '$($folder.Name)': no URL in its settings."; continue }
        $cfg = Get-MonitorSetting -Path $monitorFile
        $threshold = if ($cfg.SlowThresholdMs) { [int]$cfg.SlowThresholdMs } else { 3000 }
        $rows = Get-TelemetryRow -Folder $folder.FullName -From $from -To $Now
        $outages = @()
        $outFile = Join-Path $folder.FullName 'Outages.csv'
        if (Test-Path -LiteralPath $outFile) { $outages = @(Import-Csv -LiteralPath $outFile) }
        $stats = Get-TelemetryStat -Rows $rows -Outages $outages -DropLines (Get-DropLine -Folder $folder.FullName -From $from) -From $from -To $Now -SlowThresholdMs $threshold
        $hash = ConvertTo-UrlHash -Url $url
        if ([int]$stats.Polls -le 0) {
            Write-Line WARN "Skipping '$($folder.Name)': no polls recorded in this window, so there is nothing to publish for it."
            continue
        }
        [void]$measurements.Add([PSCustomObject]@{ Code = (Get-EndpointCode -Map $map -UrlHash $hash); UrlHash = $hash; Stats = $stats })
    }
    if ($measurements.Count -eq 0) { Write-Line WARN 'No monitors with data. Nothing published.'; return 0 }

    # Written as one object keyed by hash, which is how it is read back, so codes survive every later run
    $mapOut = [ordered]@{}
    foreach ($k in @($map.Keys | Sort-Object { $map[$_] })) { $mapOut[$k] = $map[$k] }
    ([PSCustomObject]$mapOut) | ConvertTo-Json -Compress | Set-Content -LiteralPath $mapfile -Encoding UTF8

    $outputPaths = @('data/telemetry', 'reports', 'README.md')
    $snapshot = New-OutputSnapshot -RepoPath $RepoPath -Paths $outputPaths
    $headBefore = "$(& git.exe -C $RepoPath rev-parse HEAD 2>$null)".Trim()
    $stateBefore = if (Test-Path -LiteralPath $statefile) { Get-Content -LiteralPath $statefile -Raw -ErrorAction SilentlyContinue } else { $null }
    $windowEnd = $Now.ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($m in $measurements) {
        $row = [PSCustomObject]@{
            WindowEnd_Local      = $windowEnd
            Window               = $Window
            Endpoint             = $m.Code
            UrlHash              = $m.UrlHash
            Polls                = $m.Stats.Polls
            AvailabilityPercent  = $m.Stats.AvailabilityPercent
            FailedPolls          = $m.Stats.FailedPolls
            SlowPolls            = $m.Stats.SlowPolls
            P50Ms                = $m.Stats.P50Ms
            P95Ms                = $m.Stats.P95Ms
            MaxMs                = $m.Stats.MaxMs
            Outages              = $m.Stats.Outages
            LongestOutageSeconds = $m.Stats.LongestOutageSeconds
            TotalOutageSeconds   = $m.Stats.TotalOutageSeconds
            SlowPeriods          = $m.Stats.SlowPeriods
            BackendAddresses     = $m.Stats.BackendAddresses
            FailureReasons       = $m.Stats.FailureReasons
        }
        Add-DataRow -Path (Join-Path $RepoPath "data\telemetry\$($m.Code)\$($Now.ToString('yyyy-MM')).csv") -Row $row
    }

    $published = New-Object System.Collections.ArrayList
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $RepoPath 'data\telemetry') -Filter '*.csv' -Recurse -File -ErrorAction SilentlyContinue)) {
        foreach ($r in @(Import-Csv -LiteralPath $f.FullName)) { [void]$published.Add($r) }
    }
    $today = @($published | Where-Object { (ConvertFrom-LocalStamp $_.WindowEnd_Local) -and (ConvertFrom-LocalStamp $_.WindowEnd_Local).Date -eq $Now.Date })
    $reportDir = Join-Path $RepoPath 'reports'
    if (-not (Test-Path -LiteralPath $reportDir)) { New-Item -Path $reportDir -ItemType Directory -Force | Out-Null }
    Set-Content -LiteralPath (Join-Path $reportDir "$($Now.ToString('yyyy-MM-dd')).md") -Value (New-TelemetryReport -Date $Now -Rows $today) -Encoding UTF8

    $dayAgo = $Now.AddHours(-24)
    $recent = @($published | Where-Object { $t = ConvertFrom-LocalStamp $_.WindowEnd_Local; $t -and $t -gt $dayAgo } | Group-Object Endpoint | ForEach-Object {
            $set = @($_.Group)
            [PSCustomObject]@{
                Endpoint            = $_.Name
                Polls               = (@($set | ForEach-Object { [int]$_.Polls }) | Measure-Object -Sum).Sum
                AvailabilityPercent = [math]::Round((@($set | ForEach-Object { [double]$_.AvailabilityPercent }) | Measure-Object -Average).Average, 2)
                P50Ms               = (@($set | ForEach-Object { [int]$_.P50Ms }) | Measure-Object -Maximum).Maximum
                P95Ms               = (@($set | ForEach-Object { [int]$_.P95Ms }) | Measure-Object -Maximum).Maximum
                Outages             = (@($set | ForEach-Object { [int]$_.Outages }) | Measure-Object -Sum).Sum
            }
        })
    $readme = Join-Path $RepoPath 'README.md'
    if (Test-Path -LiteralPath $readme) {
        $text = Get-Content -LiteralPath $readme -Raw
        $updated = Update-ReadmeTable -Text $text -Table (New-ReadmeTable -Rows $recent -AsOf $Now)
        if ($updated -ne $text) { Set-Content -LiteralPath $readme -Value $updated -Encoding UTF8 -NoNewline }
        else { Write-Line WARN "README.md has no telemetry markers, so its table was left alone." }
    }

    $totalPolls = (@($measurements | ForEach-Object { [int]$_.Stats.Polls }) | Measure-Object -Sum).Sum
    $totalFailed = (@($measurements | ForEach-Object { [int]$_.Stats.FailedPolls }) | Measure-Object -Sum).Sum
    $overall = if ($totalPolls -gt 0) { [math]::Round(100 * ($totalPolls - $totalFailed) / $totalPolls, 2) } else { 0 }
    $subject = New-CommitSubject -Date $Now -Window $Window -Availability $overall
    Write-Line INFO $subject

    if ($DryRun) {
        # These files were written into the real working copy, and a later run would commit them beside its own
        # rows for the same window, so exactly what this run wrote is put back.
        if (Restore-OutputSnapshot -Snapshot $snapshot) {
            Write-Line DONE 'Dry run: the files it would commit were produced and then put back, git untouched.'
        }
        else {
            Write-Line WARN "Dry run could not fully put the working copy back. Check 'git status' in '$RepoPath' before the next run."
        }
        return 0
    }

    $add = Invoke-Git -RepoPath $RepoPath -Arguments @('add', '--', 'data/telemetry', 'reports', 'README.md')
    if ($add.ExitCode -ne 0) {
        $null = Restore-OutputSnapshot -Snapshot $snapshot
        Write-Line FAIL "git add failed, so nothing was published and the files were put back. $($add.Output)"
        return 1
    }
    if ((Invoke-Git -RepoPath $RepoPath -Arguments @('diff', '--cached', '--quiet')).ExitCode -eq 0) {
        Write-Line DONE 'Nothing changed since the last run. No commit made.'
        Set-Content -LiteralPath $statefile -Value (@{ LastWindowEnd = $windowEnd; LastWindow = $Window } | ConvertTo-Json -Compress) -Encoding UTF8
        return 0
    }
    $commit = Invoke-Git -RepoPath $RepoPath -Arguments @('commit', '-m', $subject, '-m', (New-CommitBody -Measurements $measurements))
    if ($commit.ExitCode -ne 0) {
        $null = Invoke-Git -RepoPath $RepoPath -Arguments @('reset', '--mixed', $headBefore)
        $null = Restore-OutputSnapshot -Snapshot $snapshot
        Write-Line FAIL "git commit failed, so nothing was published and the files were put back. $($commit.Output)"
        return 1
    }
    # The commit records what has been measured, so the window moves on with it. A push that fails only delays
    # delivery, while leaving the window open would measure the same polls again and publish them twice.
    Set-Content -LiteralPath $statefile -Value (@{ LastWindowEnd = $windowEnd; LastWindow = $Window } | ConvertTo-Json -Compress) -Encoding UTF8

    $pushed = $false
    foreach ($attempt in 1..2) {
        $rebase = Invoke-Git -RepoPath $RepoPath -Arguments @('pull', '--rebase', '--autostash')
        if ($rebase.ExitCode -ne 0) {
            Write-Line WARN "git pull --rebase failed on attempt $attempt. $($rebase.Output)"
            if (Test-RepoMidRebase -RepoPath $RepoPath) {
                # Left in place this would break every run after this one
                $null = Invoke-Git -RepoPath $RepoPath -Arguments @('rebase', '--abort')
                $null = Invoke-Git -RepoPath $RepoPath -Arguments @('reset', '--mixed', $headBefore)
                $null = Restore-OutputSnapshot -Snapshot $snapshot
                if ($null -ne $stateBefore) { Set-Content -LiteralPath $statefile -Value $stateBefore -Encoding UTF8 -NoNewline }
                else { Remove-Item -LiteralPath $statefile -Force -ErrorAction SilentlyContinue }
                Write-Line FAIL 'The rebase stopped on a conflict, so this run was rolled back and the working copy put back. The next run measures the same window again.'
                return 1
            }
        }
        $push = Invoke-Git -RepoPath $RepoPath -Arguments @('push')
        if ($push.ExitCode -eq 0) { $pushed = $true; break }
        Write-Line WARN "git push was rejected on attempt $attempt. Rebasing and trying once more."
    }
    if (-not $pushed) { Write-Line FAIL 'Could not push after two attempts. The commit is waiting in the working copy and goes out with the next run.'; return 1 }

    Write-Line DONE "Published and pushed: $subject"
    return 0
}

if ($MyInvocation.InvocationName -ne '.') {
    $now = Get-Date
    $label = if ($Window -eq 'Auto') { if ($now.Hour -lt 12) { 'Morning' } else { 'Evening' } } else { $Window }
    exit (Publish-Window -RepoPath $RepoPath -MonitorRoot $MonitorRoot -StatePath $StatePath -Window $label -Now $now -DryRun:$DryRun)
}
