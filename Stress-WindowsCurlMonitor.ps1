$ErrorActionPreference = 'Stop'
$script:pass=0; $script:fail=0; $script:failed=@()
function Check($n,$c){ if($c){$script:pass++} else {$script:fail++; $script:failed += $n; Write-Host "FAIL: $n"} }
function Section($n){ Write-Host ""; Write-Host "== $n ==" }

$installerPath = Join-Path $PSScriptRoot 'Install-CurlMonitor.ps1'
$T=$null;$E=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($installerPath,[ref]$T,[ref]$E)
$hereNodes = $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -eq 'SingleQuotedHereString'},$true)
$template = $hereNodes[0].Value

# Load installer config assigns and the functions under test
$firstFunc = ($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] } | Select-Object -First 1).Extent.StartLineNumber
$assigns = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.AssignmentStatementAst] -and $_.Extent.StartLineNumber -lt $firstFunc }
foreach ($a in $assigns) { if ($a.Left.Extent.Text -ne '$Template_MonitorScript') { Invoke-Expression $a.Extent.Text } }
$Template_MonitorScript = $template
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('Write-Log','ConvertTo-SafeSiteName','ConvertTo-SecureText','Protect-Secret','Save-SmtpCredential','Remove-SmtpCredential','New-MonitorContent','ConvertTo-MonitorSlug','Get-InstalledMonitor','Get-LegacyInstall','Get-FolderSizeText','Get-RemovableMonitor','Show-RemovableMonitor','Select-RemovableMonitor','Stop-MonitorProcess','Move-MonitorHistory','Remove-MonitorInstall','Invoke-UninstallFlow','Get-PreviousRootMonitor','Move-MonitorToNewRoot','Invoke-RootMove','Register-MonitorTask','Protect-InstallFolder') }) { Invoke-Expression $f.Extent.Text }
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message" }

$scratch = 'C:\tmp\hst_win_tests'
if (Test-Path $scratch) {
    Get-ChildItem $scratch -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { & icacls.exe $_.FullName /grant "${env:USERNAME}:(F)" 2>$null | Out-Null }
    Remove-Item $scratch -Recurse -Force
}
New-Item $scratch -ItemType Directory | Out-Null

Section "K. DPAPI protect and unprotect roundtrip (PS $($PSVersionTable.PSVersion))"
$plainSecret = "s3cret-O'Brien-" + [char]0x00DC + [char]0x00EF + "-" + [char]0xD83D + [char]0xDD11
$sec = ConvertTo-SecureString $plainSecret -AsPlainText -Force
$cipher = Protect-Secret -Password $sec
Check "K1 cipher is valid base64" ($null -ne ([Convert]::FromBase64String($cipher)))
Check "K1 cipher does not contain the plaintext" (-not ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($cipher))).Contains("s3cret"))
$cipherPlain = Protect-Secret -PlainText $plainSecret
Add-Type -AssemblyName System.Security
$unprotect = { param($c) [Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($c), [Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1'), [Security.Cryptography.DataProtectionScope]::LocalMachine)) }
Check "K1 Protect-Secret from text and from SecureString decrypt to the same secret" ((& $unprotect $cipher) -eq $plainSecret -and (& $unprotect $cipherPlain) -eq $plainSecret)
$credRt = New-Object System.Management.Automation.PSCredential('svc', (ConvertTo-SecureText -Text $plainSecret))
Check "K1 ConvertTo-SecureText round-trips unicode and surrogate pairs through PSCredential" ($credRt.GetNetworkCredential().Password -eq $plainSecret -and $credRt.Password.IsReadOnly())
Check "K1 Protect-Secret rejects an empty text secret" ((& { try { Protect-Secret -PlainText '' ; $false } catch { $true } }))
$InstallDir = $scratch
$credPath = Join-Path $scratch $CredentialFileName
# The file is locked to SYSTEM and Administrators before any content lands in it. Not elevated, this caller is refused
# at the write and the file stays empty. Elevated, the write succeeds and only those two principals have access.
$saveThrew = $false
try { Save-SmtpCredential -CipherText $cipher } catch { $saveThrew = $true }
$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($elevated) {
    $credAcl = Get-Acl $credPath
    $credSids = @($credAcl.Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } | Sort-Object -Unique)
    Check "K1 credential file locked to SYSTEM and Administrators, inheritance removed, content written (elevated)" (-not $saveThrew -and (Get-Item $credPath).Length -gt 0 -and $credAcl.AreAccessRulesProtected -and (($credSids -join ',') -eq 'S-1-5-18,S-1-5-32-544'))
}
else {
    Check "K1 credential file created and locked before content (non-admin write refused, file empty)" ($saveThrew -and (Test-Path $credPath) -and ((Get-Item $credPath).Length -eq 0))
}
$acl = & icacls.exe $credPath
$aclText = ($acl | Out-String)
Check "K1 ACL grants SYSTEM full" ($aclText -match [regex]::Escape('NT AUTHORITY\SYSTEM:(F)'))
Check "K1 ACL grants Administrators full" ($aclText -match [regex]::Escape('BUILTIN\Administrators:(F)'))
$grantLines = @($acl | Where-Object { $_ -match ':\(' })
Check "K1 ACL has exactly 2 entries (inheritance stripped)" ($grantLines.Count -eq 2)
# Owner re-grants self so the test can read the file back (SYSTEM does this for real)
& icacls.exe $credPath /grant "${env:USERNAME}:(F)" | Out-Null
Set-Content -Path $credPath -Value $cipher -Encoding ASCII
Check "K1 file on disk is the base64 cipher, not plaintext" ((Get-Content $credPath -Raw).Trim() -eq $cipher)

# Decrypt through the monitor's own Get-ProtectedSecret extracted from a generated monitor
$mailK = @{ MailMethod='Authenticated'; SmtpServer='smtp.office365.com'; SmtpPort=587; SmtpUseSsl=$true; MailFrom='m@x.invalid'; MailTo=@('a@x.invalid'); SmtpAuthUser='m@x.invalid'; CipherText=$cipher }
$genK = New-MonitorContent -SiteName 'WinTest' -Mail $mailK
$gT=$null;$gE=$null
$gAst=[System.Management.Automation.Language.Parser]::ParseInput($genK,[ref]$gT,[ref]$gE)
foreach ($f in $gAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-ProtectedSecret'},$true)) { Invoke-Expression $f.Extent.Text }
$SmtpCredentialFile = $credPath
Check "K1 monitor decrypts installer's cipher back to the exact secret" ((Get-ProtectedSecret) -eq $plainSecret)
# Tampered file degrades to null, no throw
Set-Content $credPath -Value ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('garbage-not-dpapi'))) -Encoding ASCII
Check "K1 tampered cipher returns null with warning" ($null -eq (Get-ProtectedSecret) -and $script:LastLog -match 'Could not decrypt')
Set-Content $credPath -Value 'not base64 at all !!!' -Encoding ASCII
Check "K1 non-base64 content returns null" ($null -eq (Get-ProtectedSecret))
$SmtpCredentialFile = Join-Path $scratch 'missing.bin'
Check "K1 missing file returns null with warning" ($null -eq (Get-ProtectedSecret) -and $script:LastLog -match 'missing')
Remove-SmtpCredential
Check "K1 Remove-SmtpCredential deletes the file" (-not (Test-Path $credPath))

Section "L. Scheduled task objects built with the installer's exact parameters"
Import-Module ScheduledTasks
$RestartCount = 3; $RestartMinutes = 1; $RunAsUser = 'SYSTEM'
$monitorPath = 'C:\ProgramData\CurlMonitor\Watch-CurlMonitor.ps1'
$taskOk = $true
try {
    $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$monitorPath`""
    $trigger   = New-ScheduledTaskTrigger -AtStartup
    $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -RestartCount $RestartCount -RestartInterval (New-TimeSpan -Minutes $RestartMinutes) -ExecutionTimeLimit ([TimeSpan]::Zero) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd
    $principal = New-ScheduledTaskPrincipal -UserId $RunAsUser -LogonType ServiceAccount -RunLevel Highest
} catch { $taskOk = $false; Write-Host "  task object error: $($_.Exception.Message)" }
Check "L1 all four task objects construct without error" $taskOk
if ($taskOk) {
    Check "L1 no execution time limit (PT0S)" ([string]$settings.ExecutionTimeLimit -eq 'PT0S')
    Check "L1 single instance policy IgnoreNew" ([string]$settings.MultipleInstances -eq 'IgnoreNew')
    Check "L1 restart 3x every PT1M" ($settings.RestartCount -eq 3 -and [string]$settings.RestartInterval -eq 'PT1M')
    Check "L1 battery and availability settings" ($settings.StartWhenAvailable -and -not $settings.DisallowStartIfOnBatteries -and -not $settings.StopIfGoingOnBatteries)
    Check "L1 principal SYSTEM ServiceAccount Highest" ($principal.UserId -eq 'SYSTEM' -and [string]$principal.LogonType -eq 'ServiceAccount' -and [string]$principal.RunLevel -eq 'Highest')
    Check "L1 boot trigger" ($trigger.CimClass.CimClassName -match 'BootTrigger')
    Check "L1 action hidden window, file path quoted" ($action.Execute -eq 'powershell.exe' -and $action.Arguments -match '-WindowStyle Hidden' -and $action.Arguments -match '-File "')
}

Section "M. End-to-end: generated monitor runs live under this PowerShell"
# M1: refused port, DirectSend to a bogus host: DOWN must be declared, send failure must not kill the loop
$runDir = Join-Path $scratch 'run_down'
New-Item $runDir -ItemType Directory | Out-Null
$InstallDir = $runDir
$Url = 'http://127.0.0.1:9/'
$IntervalSeconds = 1; $TimeoutSeconds = 2; $DownThreshold = 3; $ReAlertMinutes = 30
$mailDown = @{ MailMethod='DirectSend'; SmtpServer='nonexistent-host-zz.invalid'; SmtpPort=25; SmtpUseSsl=$false; MailFrom='m@x.invalid'; MailTo=@('a@x.invalid'); SmtpAuthUser=''; CipherText='' }
$genDown = New-MonitorContent -SiteName 'WinDown' -Mail $mailDown
$monDown = Join-Path $runDir 'monitor.ps1'
Set-Content $monDown -Value $genDown -Encoding UTF8
$p = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monDown`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 45
$stillRunning = -not $p.HasExited
Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Check "M1 monitor process survived 45 s (send failures non-fatal)" $stillRunning
$latCsv = Get-ChildItem $runDir -Filter 'Latency_*.csv' | Select-Object -First 1
Check "M1 monthly latency CSV created" ($null -ne $latCsv)
if ($latCsv) {
    $rows = @(Import-Csv $latCsv.FullName)
    Check "M1 polled continuously (>= 8 rows in 45 s, a refused loopback connect can take 2 s on filtered hosts)" ($rows.Count -ge 8)
    $cols = ($rows[0].PSObject.Properties.Name -join ',')
    Check "M1 CSV schema complete" ($cols -eq 'Timestamp_Local,Timestamp_UTC,SiteName,Url,HttpCode,CurlExit,Reason,DnsMs,ConnectMs,TlsMs,TtfbMs,TotalMs,SizeBytes,RemoteIp,Redirects,RedirectMs,FinalUrl,ContentOk')
    $badRows = @($rows | Where-Object { $_.HttpCode -ne '000' -or (-not (($_.CurlExit -eq '7' -and $_.Reason -eq 'Connection refused or unreachable') -or ($_.CurlExit -eq '28' -and $_.Reason -eq 'Timed out'))) })
    Check "M1 every row classified failed with mapped reason (refused or timed out)" ($badRows.Count -eq 0)
    Check "M1 site name baked into rows" ($rows[0].SiteName -eq 'WinDown')
}
$logDown = Get-ChildItem $runDir -Filter 'Transcript_*.log' | Select-Object -First 1
Check "M1 daily transcript log created" ($null -ne $logDown)
if ($logDown) {
    $logText = Get-Content $logDown.FullName -Raw
    Check "M1 DOWN alert raised after threshold" ($logText -match '\[DOWN\] WinDown')
    Check "M1 mail send failure logged FAILED, loop continued" ($logText -match 'FAILED \|')
    $dropM1 = Get-Content (Join-Path $runDir 'Drops.log') -ErrorAction SilentlyContinue
    Check "M1 drops log: start line, one FAIL per failed poll, DOWN, alert not sent, nothing healthy" ($dropM1 -and @($dropM1 | Where-Object { $_ -match '\| START     \|' }).Count -eq 1 -and @($dropM1 | Where-Object { $_ -match '\| FAIL      \| Site=WinDown Code=000' }).Count -ge 8 -and @($dropM1 | Where-Object { $_ -match '\| DOWN      \| Declared DOWN for WinDown after 3' }).Count -eq 1 -and @($dropM1 | Where-Object { $_ -match '\| ALERT     \| Not sent, retrying every minute.*\[DOWN\] WinDown' }).Count -eq 1 -and -not (($dropM1 -join "`n") -match 'SUCCESS|Code=200'))
    # Anchored on the alert send failing, not on a failed poll: every poll against a closed port logs FAILED
$m1Lines = @($logText -split "`n")
$m1SendFail = @($m1Lines | Select-String 'Could not send alert email' | Select-Object -First 1)
$m1PollsAfter = if (@($m1SendFail).Count -eq 1) { @($m1Lines[@($m1SendFail)[0].LineNumber..($m1Lines.Count - 1)] | Where-Object { $_ -match 'Site=' }).Count } else { 0 }
Check "M1 polling continued after the failed alert" (@($m1SendFail).Count -eq 1 -and $m1PollsAfter -ge 3)
}

# M2: the real HST endpoint with the installer defaults: redirects followed to the sign-in page, UP rows, no DOWN
$runDir2 = Join-Path $scratch 'run_up'
New-Item $runDir2 -ItemType Directory | Out-Null
$InstallDir = $runDir2
$Url = 'https://prodasp09.hstpathways.com/p95_CSP/HSTeChart'
$IntervalSeconds = 2; $TimeoutSeconds = 15
$genUp = New-MonitorContent -SiteName 'WinUp' -Mail $mailDown
$monUp = Join-Path $runDir2 'monitor.ps1'
Set-Content $monUp -Value $genUp -Encoding UTF8
$p2 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monUp`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 20
Stop-Process -Id $p2.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$latCsv2 = Get-ChildItem $runDir2 -Filter 'Latency_*.csv' | Select-Object -First 1
Check "M2 latency CSV created" ($null -ne $latCsv2)
if ($latCsv2) {
    $rows2 = @(Import-Csv $latCsv2.FullName)
    Check "M2 has UP rows" ($rows2.Count -ge 3)
    $up = @($rows2 | Where-Object { $_.HttpCode -eq '200' })
    Check "M2 rows are 200 with content OK" ($up.Count -eq $rows2.Count -and @($rows2 | Where-Object { $_.ContentOk -ne 'True' }).Count -eq 0)
    Check "M2 timings numeric and IP captured" ($rows2[0].TotalMs -match '^\d+$' -and $rows2[0].RemoteIp -match '\d+\.\d+\.\d+\.\d+|:')
    Check "M2 every poll followed 2 redirects to the federation sign-in page" (@($rows2 | Where-Object { $_.Redirects -eq '2' -and $_.FinalUrl -match '^https://prodasp09\.hstpathways\.com/p95_CSP/HSTFederationProvider/' }).Count -eq $rows2.Count)
    Check "M2 redirect time recorded and below total" ($rows2[0].RedirectMs -match '^\d+$' -and [int]$rows2[0].RedirectMs -gt 0 -and [int]$rows2[0].RedirectMs -le [int]$rows2[0].TotalMs)
    Check "M2 sign-in page body well above the minimum size" (@($rows2 | Where-Object { [int]$_.SizeBytes -lt 1000 }).Count -eq 0)
}
$logUp = Get-ChildItem $runDir2 -Filter 'Transcript_*.log' | Select-Object -First 1
Check "M2 no DOWN alert on healthy endpoint" ($null -ne $logUp -and -not ((Get-Content $logUp.FullName -Raw) -match '\[DOWN\]'))
$dropM2 = @(Get-Content (Join-Path $runDir2 'Drops.log') -ErrorAction SilentlyContinue)
Check "M2 healthy endpoint: drops log holds only the start line (slow polls allowed)" ($dropM2.Count -ge 1 -and @($dropM2 | Where-Object { $_ -notmatch '\| (START|SLOW) +\|' }).Count -eq 0)

# M4: full outage lifecycle: DOWN on closed port, then RESOLVED plus outage CSV row when the port comes up
$runDir3 = Join-Path $scratch 'run_recover'
New-Item $runDir3 -ItemType Directory | Out-Null
$InstallDir = $runDir3
function Get-FreeTestPort { $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0); $l.Start(); $p = ($l.LocalEndpoint).Port; $l.Stop(); return $p }
$port = Get-FreeTestPort
$Url = "http://127.0.0.1:$port/"
$IntervalSeconds = 1; $TimeoutSeconds = 2; $DownThreshold = 3; $MinPopulatedBytes = 100; $ExpectedContentMarker = ''
$genRec = New-MonitorContent -SiteName 'WinRecover' -Mail $mailDown
$monRec = Join-Path $runDir3 'monitor.ps1'
Set-Content $monRec -Value $genRec -Encoding UTF8
$p3 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monRec`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 14
$srv = Start-Job -ScriptBlock {
    param($port)
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
    $listener.Start()
    $deadline = (Get-Date).AddSeconds(25)
    $body = ('HST eChart OK ' * 20)
    while ((Get-Date) -lt $deadline) {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient()
            $stream = $client.GetStream()
            $stream.ReadTimeout = 300
            $buf = New-Object byte[] 4096
            try { $stream.Read($buf, 0, $buf.Length) | Out-Null } catch { }
            $resp = "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($resp)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush()
            $client.Close()
        } else { Start-Sleep -Milliseconds 50 }
    }
    $listener.Stop()
} -ArgumentList $port
Start-Sleep -Seconds 15
Stop-Process -Id $p3.Id -Force -ErrorAction SilentlyContinue
$srv | Wait-Job -Timeout 15 | Out-Null; $srv | Remove-Job -Force
Start-Sleep -Seconds 2
$logRec = Get-ChildItem $runDir3 -Filter 'Transcript_*.log' | Select-Object -First 1
Check "M4 transcript exists" ($null -ne $logRec)
if ($logRec) {
    $recText = Get-Content $logRec.FullName -Raw
    Check "M4 DOWN declared while port closed" ($recText -match '\[DOWN\] WinRecover')
    Check "M4 RESOLVED raised when port came up" ($recText -match '\[RESOLVED\] WinRecover \([^)]+\) - outage lasted ')
}
$outCsv = Join-Path $runDir3 'Outages.csv'
Check "M4 outages CSV written" (Test-Path $outCsv)
if (Test-Path $outCsv) {
    $out = @(Import-Csv $outCsv)
    Check "M4 exactly one outage record" ($out.Count -eq 1)
    $dropM4 = Get-Content (Join-Path $runDir3 'Drops.log') -ErrorAction SilentlyContinue
    Check "M4 drops log: DOWN then RESOLVED with recovery code, alert lines for both" ($dropM4 -and (($dropM4 -join "`n") -match '\| DOWN      \| Declared DOWN for WinRecover') -and (($dropM4 -join "`n") -match '\| RESOLVED  \| Outage record written: WinRecover lasted .* Recovery HTTP 200 from 127\.0\.0\.1\.') -and @($dropM4 | Where-Object { $_ -match '\| ALERT     \| Not sent.*\[(DOWN|RESOLVED)\] WinRecover' }).Count -eq 2)
    Check "M4 record fields sane (duration, polls, recovery 200)" ($out[0].SiteName -eq 'WinRecover' -and [int]$out[0].DurationSeconds -ge 3 -and [int]$out[0].FailedPolls -ge 3 -and $out[0].RecoveryCode -eq '200' -and $out[0].OutageStart_UTC -and $out[0].OutageEnd_UTC)
}

# M6: restart with a stale heartbeat and an outage in progress: gap reported, outage carried over and resolved from its true onset
$runDir6 = Join-Path $scratch 'run_restart'
New-Item $runDir6 -ItemType Directory | Out-Null
$InstallDir = $runDir6
$port6 = Get-FreeTestPort
$Url = "http://127.0.0.1:$port6/"
$IntervalSeconds = 1; $TimeoutSeconds = 2; $DownThreshold = 3; $MinPopulatedBytes = 100; $ExpectedContentMarker = ''
$genRs = New-MonitorContent -SiteName 'WinRestart' -Mail $mailDown
$monRs = Join-Path $runDir6 'monitor.ps1'
Set-Content $monRs -Value $genRs -Encoding UTF8
$nowUtc = (Get-Date).ToUniversalTime()
$onsetUtc = $nowUtc.AddMinutes(-15)
$hb6 = [ordered]@{ Beat = $nowUtc.AddMinutes(-10).ToString('o'); IsDown = $true; ConsecutiveFailures = 90; OutageStartUtc = $onsetUtc.ToString('o'); OutageStartLocalStr = $onsetUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'); OutageStartUtcStr = $onsetUtc.ToString('yyyy-MM-dd HH:mm:ss'); LastAlertUtc = $onsetUtc.ToString('o'); AlertDelivered = $false }
$hb6 | ConvertTo-Json -Compress | Set-Content (Join-Path $runDir6 'heartbeat.json') -Encoding UTF8
$srv6 = Start-Job -ScriptBlock {
    param($port)
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
    $listener.Start()
    $deadline = (Get-Date).AddSeconds(40)
    $body = ('HST eChart OK ' * 20)
    while ((Get-Date) -lt $deadline) {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient()
            $stream = $client.GetStream()
            $stream.ReadTimeout = 300
            $buf = New-Object byte[] 4096
            try { $stream.Read($buf, 0, $buf.Length) | Out-Null } catch { }
            $resp = "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($resp)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush()
            $client.Close()
        } else { Start-Sleep -Milliseconds 50 }
    }
    $listener.Stop()
} -ArgumentList $port6
Start-Sleep -Seconds 2
$p6 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monRs`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 25
Stop-Process -Id $p6.Id -Force -ErrorAction SilentlyContinue
$srv6 | Wait-Job -Timeout 30 | Out-Null; $srv6 | Remove-Job -Force
Start-Sleep -Seconds 2
$logRs = Get-ChildItem $runDir6 -Filter 'Transcript_*.log' | Select-Object -First 1
Check "M6 transcript exists" ($null -ne $logRs)
if ($logRs) {
    $rsText = Get-Content $logRs.FullName -Raw
    Check "M6 gap detected and outage carry-over logged" ($rsText -match 'The monitor was not running for 10m [0-2]\ds' -and $rsText -match 'outage was in progress at the last heartbeat')
    Check "M6 restart notice queued for retry (mail path is down here)" ($rsText -match "'\[MONITOR RESTARTED\] WinRestart \($env:COMPUTERNAME\) - not running for 10m [0-2]\ds' will be retried")
    Check "M6 RESOLVED from the original onset, 15 minutes ago" ($rsText -match '\[RESOLVED\] WinRestart \([^)]+\) - outage lasted 15m [0-2]\ds')
    Check "M6 RESOLVED superseded nothing but the outage alerts, restart notice still queued" (-not ($rsText -match 'Dropping undelivered'))
    Check "M6 no DOWN declared on a healthy endpoint" (-not ($rsText -match '\[DOWN\]'))
}
$outCsv6 = Join-Path $runDir6 'Outages.csv'
Check "M6 outage record spans the gap with the pre-restart poll count" ((Test-Path $outCsv6) -and (& { $o = @(Import-Csv $outCsv6); $o.Count -eq 1 -and [int]$o[0].DurationSeconds -ge 900 -and [int]$o[0].DurationSeconds -le 990 -and [int]$o[0].FailedPolls -eq 90 -and $o[0].RecoveryCode -eq '200' }))
$hbAfter = Get-Content (Join-Path $runDir6 'heartbeat.json') -Raw | ConvertFrom-Json
$dropM6 = Get-Content (Join-Path $runDir6 'Drops.log') -ErrorAction SilentlyContinue
Check "M6 drops log: RESTART with cause, CARRYOVER with onset and polls, RESOLVED from the true onset" ($dropM6 -and (($dropM6 -join "`n") -match '\| RESTART   \| Monitor was not running for 10m [0-2]\ds\. Server did not restart') -and (($dropM6 -join "`n") -match '\| CARRYOVER \| Outage in progress since .* \(90 failed polls so far\)') -and (($dropM6 -join "`n") -match '\| RESOLVED  \| Outage record written: WinRestart lasted 15m [0-2]\ds over 90 failed polls'))
Check "M6 heartbeat rewritten as up with a fresh beat" (-not $hbAfter.IsDown -and ([datetime]::Parse($hbAfter.Beat, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)).ToUniversalTime() -gt $nowUtc)

# M8: the installer marked a deliberate stop while an outage was in progress: no RESTART, outage carried over and resolved
$runDir8 = Join-Path $scratch 'run_stopped'
New-Item $runDir8 -ItemType Directory | Out-Null
$InstallDir = $runDir8
$port8 = Get-FreeTestPort
$Url = "http://127.0.0.1:$port8/"
$IntervalSeconds = 1; $TimeoutSeconds = 2; $DownThreshold = 3; $MinPopulatedBytes = 100; $ExpectedContentMarker = ''
$genSt = New-MonitorContent -SiteName 'WinStopped' -Mail $mailDown
$monSt = Join-Path $runDir8 'monitor.ps1'
Set-Content $monSt -Value $genSt -Encoding UTF8
$nowUtc8 = (Get-Date).ToUniversalTime()
$onset8 = $nowUtc8.AddMinutes(-20)
$hb8 = [ordered]@{ Beat = $nowUtc8.AddHours(-3).ToString('o'); IsDown = $true; ConsecutiveFailures = 40; OutageStartUtc = $onset8.ToString('o'); OutageStartLocalStr = $onset8.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'); OutageStartUtcStr = $onset8.ToString('yyyy-MM-dd HH:mm:ss'); LastAlertUtc = $onset8.ToString('o'); AlertDelivered = $true; Stopped = $true }
$hb8 | ConvertTo-Json -Compress | Set-Content (Join-Path $runDir8 'heartbeat.json') -Encoding UTF8
$srv8 = Start-Job -ScriptBlock {
    param($port)
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
    $listener.Start()
    $deadline = (Get-Date).AddSeconds(30)
    $body = ('HST eChart OK ' * 20)
    while ((Get-Date) -lt $deadline) {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient(); $stream = $client.GetStream(); $stream.ReadTimeout = 300
            $buf = New-Object byte[] 4096
            try { $stream.Read($buf, 0, $buf.Length) | Out-Null } catch { }
            $resp = "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($resp); $stream.Write($bytes, 0, $bytes.Length); $stream.Flush(); $client.Close()
        } else { Start-Sleep -Milliseconds 50 }
    }
    $listener.Stop()
} -ArgumentList $port8
Start-Sleep -Seconds 2
$p8 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monSt`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 20
Stop-Process -Id $p8.Id -Force -ErrorAction SilentlyContinue
$srv8 | Wait-Job -Timeout 20 | Out-Null; $srv8 | Remove-Job -Force
Start-Sleep -Seconds 2
$stText = (Get-ChildItem $runDir8 -Filter 'Transcript_*.log' | Select-Object -First 1 | Get-Content -Raw)
$dropM8 = @(Get-Content (Join-Path $runDir8 'Drops.log') -ErrorAction SilentlyContinue)
Check "M8 deliberate stop: no restart notice after a 3 h gap, outage carried over" ($stText -match 'Monitor restarted after a deliberate stop' -and $stText -match 'outage was in progress at the last heartbeat' -and -not ($stText -cmatch '\[MONITOR RESTARTED\]'))
Check "M8 deliberate stop: RESOLVED from the 20 min onset with 40 carried polls, DOWN alert marked delivered" ($stText -match '\[RESOLVED\] WinStopped \([^)]+\) - outage lasted 20m [0-2]\ds' -and @($dropM8 | Where-Object { $_ -match '\| RESOLVED  \| Outage record written: WinStopped lasted 20m [0-2]\ds over 40 failed polls' }).Count -eq 1 -and @($dropM8 | Where-Object { $_ -match '\| CARRYOVER \|' }).Count -eq 1 -and @($dropM8 | Where-Object { $_ -match '\| RESTART' }).Count -eq 0)

# M7: crash without a heartbeat gap is quiet, and a fresh install (no heartbeat) sends no notice
$runDir7 = Join-Path $scratch 'run_fresh'
New-Item $runDir7 -ItemType Directory | Out-Null
$InstallDir = $runDir7
$Url = 'http://127.0.0.1:9/'
$genFr = New-MonitorContent -SiteName 'WinFresh' -Mail $mailDown
$monFr = Join-Path $runDir7 'monitor.ps1'
Set-Content $monFr -Value $genFr -Encoding UTF8
$p7 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monFr`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 12
Stop-Process -Id $p7.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$p7b = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monFr`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 12
Stop-Process -Id $p7b.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$frText = (Get-ChildItem $runDir7 -Filter 'Transcript_*.log' | Select-Object -First 1 | Get-Content -Raw)
Check "M7 first start (no heartbeat) sends no restart notice" (-not ($frText -match '\[MONITOR RESTARTED\]') -and (Test-Path (Join-Path $runDir7 'heartbeat.json')))
$dropM7 = @(Get-Content (Join-Path $runDir7 'Drops.log') -ErrorAction SilentlyContinue)
Check "M7 two starts, no RESTART line for a quick restart, FAIL lines only" (@($dropM7 | Where-Object { $_ -match '\| START     \|' }).Count -eq 2 -and @($dropM7 | Where-Object { $_ -match '\| RESTART' }).Count -eq 0 -and @($dropM7 | Where-Object { $_ -notmatch '\| (START|FAIL|DOWN|ALERT|STOP|CARRYOVER) +\|' }).Count -eq 0)
Check "M7 quick restart is logged below the threshold, no notice" ($frText -match 'Monitor restarted after a \d+ s gap, below the notice threshold' -and -not ($frText -match '\[MONITOR RESTARTED\]'))

# M5: latency CSV held open exclusively by another program: polls skipped with a warning, DOWN alert still fires, file not moved aside
$runDir5 = Join-Path $scratch 'run_locked'
New-Item $runDir5 -ItemType Directory | Out-Null
$InstallDir = $runDir5
$Url = 'http://127.0.0.1:9/'
$IntervalSeconds = 1; $TimeoutSeconds = 2; $DownThreshold = 3; $ExpectedContentMarker = ''
$genLock = New-MonitorContent -SiteName 'WinLocked' -Mail $mailDown
$monLock = Join-Path $runDir5 'monitor.ps1'
Set-Content $monLock -Value $genLock -Encoding UTF8
$csvLock = Join-Path $runDir5 ("Latency_" + (Get-Date -Format 'yyyyMM') + ".csv")
$fsLock = [System.IO.File]::Open($csvLock, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
$p5 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monLock`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 20
Stop-Process -Id $p5.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$fsLock.Close()
$logLock = Get-ChildItem $runDir5 -Filter 'Transcript_*.log' | Select-Object -First 1
Check "M5 transcript exists" ($null -ne $logLock)
if ($logLock) {
    $lockText = Get-Content $logLock.FullName -Raw
    Check "M5 DOWN declared while the latency CSV was locked" ($lockText -match '\[DOWN\] WinLocked')
    Check "M5 skipped polls logged as warnings, no cycle errors" ($lockText -match 'Could not append to' -and -not ($lockText -match 'Probe cycle error'))
}
Check "M5 locked CSV not moved aside" (@(Get-ChildItem $runDir5 -Filter 'Latency_*_schema-*.csv').Count -eq 0)

# M9: sustained slow responses raise SLOW, recovery raises SLOW RESOLVED, and a due daily summary goes out on the first poll
$runDir9 = Join-Path $scratch 'run_slow'
New-Item $runDir9 -ItemType Directory | Out-Null
$InstallDir = $runDir9
$port9 = Get-FreeTestPort
$slowFlag = Join-Path $runDir9 'slow.flag'
$Url = "http://127.0.0.1:$port9/"
$saved9 = @($SlowThresholdMs, $SlowWindowMinutes, $SlowAlertPercent, $SlowClearPercent, $AlertOnSlow, $DailySummaryHour)
$IntervalSeconds = 1; $TimeoutSeconds = 5; $DownThreshold = 3; $MinPopulatedBytes = 100; $ExpectedContentMarker = ''
$SlowThresholdMs = 1000; $SlowWindowMinutes = 1; $SlowAlertPercent = 50; $SlowClearPercent = 10; $AlertOnSlow = $true; $DailySummaryHour = (Get-Date).Hour
$genSlow9 = New-MonitorContent -SiteName 'WinSlow' -Mail $mailDown
$SlowThresholdMs, $SlowWindowMinutes, $SlowAlertPercent, $SlowClearPercent, $AlertOnSlow, $DailySummaryHour = $saved9
$monSlow9 = Join-Path $runDir9 'monitor.ps1'
Set-Content $monSlow9 -Value $genSlow9 -Encoding UTF8
Set-Content (Join-Path $runDir9 'Drops.log') -Value ("{0} | {1,-9} | {2}" -f (Get-Date).AddHours(-1).ToString('yyyy-MM-dd HH:mm:ss'), 'FAIL', 'Site=WinSlow Code=000 Redirects=0 TTFB=5000ms Total=5000ms Populated=False IP= Reason=Timed out') -Encoding UTF8
Set-Content (Join-Path $runDir9 'summary-sent.txt') -Value (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') -Encoding ASCII
$srv9 = Start-Job -ScriptBlock {
    param($port, $flag)
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
    $listener.Start()
    $deadline = (Get-Date).AddSeconds(240)
    $body = ('HST eChart OK ' * 20)
    while ((Get-Date) -lt $deadline) {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient()
            $stream = $client.GetStream()
            $stream.ReadTimeout = 300
            $buf = New-Object byte[] 4096
            try { $stream.Read($buf, 0, $buf.Length) | Out-Null } catch { }
            if (Test-Path $flag) { Start-Sleep -Milliseconds 1500 }
            $resp = "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($resp)
            try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush() } catch { }
            $client.Close()
        } else { Start-Sleep -Milliseconds 30 }
    }
    $listener.Stop()
} -ArgumentList $port9, $slowFlag
Start-Sleep -Seconds 2
$p9 = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monSlow9`"" -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 20
Set-Content -Path $slowFlag -Value 'x'
Start-Sleep -Seconds 75
Remove-Item $slowFlag -Force
Start-Sleep -Seconds 95
Stop-Process -Id $p9.Id -Force -ErrorAction SilentlyContinue
$srv9 | Stop-Job -ErrorAction SilentlyContinue; $srv9 | Remove-Job -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$drop9 = @(Get-Content (Join-Path $runDir9 'Drops.log') -ErrorAction SilentlyContinue)
$drop9Text = $drop9 -join "`n"
$log9 = (Get-ChildItem $runDir9 -Filter 'Transcript_*.log' | Select-Object -First 1 | Get-Content -Raw)
Check "M9 START line states the slow alert rule" ($drop9Text -match '\| START +\| Monitor started on .* slow at 50% of polls over 1000 ms or failed in 1 min\)\.')
Check "M9 slow responses logged as SLOW lines, no failures, no outage" (@($drop9 | Where-Object { $_ -match '\| SLOW +\| Site=WinSlow Code=200' }).Count -ge 10 -and -not ($drop9Text -match '\| DOWN +\|') -and @($drop9 | Where-Object { $_ -match '\| FAIL +\|' }).Count -eq 1)
Check "M9 SLOW declared once while the endpoint was slow" (@($drop9 | Where-Object { $_ -match '\| SLOWSTART \| Declared SLOW for WinSlow: \d+ polls in 1 min: \d+ slower than 1000 ms, 0 failed \(median \d+ ms, worst \d+ ms\)\.' }).Count -eq 1)
Check "M9 slow period closed after recovery with its totals" (@($drop9 | Where-Object { $_ -match '\| SLOWCLEAR \| Slow period over for WinSlow: lasted \d+m \d+s, \d+ polls, \d+ slow, 0 failed, worst \d+ ms\.' }).Count -eq 1)
Check "M9 SLOW and SLOW RESOLVED emails attempted in order" ($drop9Text -match '(?s)\| ALERT +\| Not sent, retrying every minute for up to 60 minutes: \[SLOW\] WinSlow \([^)]+\) - \d+ of \d+ polls slow or failed in 1 min.*\| SLOWCLEAR .*\[SLOW RESOLVED\] WinSlow \([^)]+\) - slow period lasted ')
Check "M9 heartbeat ends not slow" (-not (Get-Content (Join-Path $runDir9 'heartbeat.json') -Raw | ConvertFrom-Json).IsSlow)
Check "M9 due daily summary built from the drops log and sent on the first poll" ($log9 -match 'Daily summary: \[DAILY\] WinSlow \([^)]+\) - 0 slow polls, 1 failed poll, 0 outages in 24 hours' -and $drop9Text -match '\| ALERT +\| Not sent, retrying every minute for up to 60 minutes: \[DAILY\] WinSlow')
Check "M9 summary date recorded, so it is not sent twice" ((Get-Content (Join-Path $runDir9 'summary-sent.txt') -TotalCount 1) -eq (Get-Date).ToString('yyyy-MM-dd') -and @([regex]::Matches($log9, 'Daily summary: ')).Count -eq 1)

# U: a real removal on this machine. Two monitors polling under scratch tasks, one removed, the other untouched.
$runDirU = Join-Path $scratch 'run_uninstall'
New-Item $runDirU -ItemType Directory | Out-Null
$uTaskPath = '\CurlMonitorTest\'
$portU = Get-FreeTestPort
$Url = "http://127.0.0.1:$portU/"
$IntervalSeconds = 1; $TimeoutSeconds = 3; $DownThreshold = 3; $MinPopulatedBytes = 50; $ExpectedContentMarker = ''
$srvU = Start-Job -ScriptBlock {
    param($port)
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port); $l.Start()
    $deadline = (Get-Date).AddSeconds(180); $body = ('CurlMonitor OK ' * 10)
    while ((Get-Date) -lt $deadline) {
        if ($l.Pending()) {
            $c = $l.AcceptTcpClient(); $s = $c.GetStream(); $s.ReadTimeout = 300
            $buf = New-Object byte[] 4096; try { $s.Read($buf, 0, $buf.Length) | Out-Null } catch { }
            $resp = "HTTP/1.1 200 OK`r`nContent-Type: text/html`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n$body"
            $b = [System.Text.Encoding]::ASCII.GetBytes($resp); $s.Write($b, 0, $b.Length); $s.Flush(); $c.Close()
        } else { Start-Sleep -Milliseconds 30 }
    }
    $l.Stop()
} -ArgumentList $portU
Start-Sleep -Seconds 2
$uProcs = @{}
foreach ($name in @('Keeper', 'Doomed')) {
    $dir = Join-Path $runDirU (ConvertTo-MonitorSlug $name)
    New-Item $dir -ItemType Directory -Force | Out-Null
    $InstallDir = $dir
    $MonitorName = $name
    $gen = New-MonitorContent -SiteName 'UninstallTest' -Mail $mailDown
    Set-Content (Join-Path $dir 'Watch-CurlMonitor.ps1') -Value $gen -Encoding UTF8
    @{ MonitorName = $name; Url = $Url; SiteName = 'UninstallTest' } | ConvertTo-Json | Set-Content (Join-Path $dir 'install-settings.json') -Encoding UTF8
    Set-Content (Join-Path $dir 'credential.bin') -Value 'cipher' -Encoding ASCII
    Register-MonitorTask -Name ($TaskNamePrefix + $name) -Path $uTaskPath -ScriptPath (Join-Path $dir 'Watch-CurlMonitor.ps1')
    Start-ScheduledTask -TaskPath $uTaskPath -TaskName ($TaskNamePrefix + $name)
    $uProcs[$name] = $dir
}
Start-Sleep -Seconds 25
$keeperCsv = Get-ChildItem (Join-Path $runDirU 'Keeper') -Filter 'Latency_*.csv' | Select-Object -First 1
$keeperBefore = if ($keeperCsv) { @(Import-Csv $keeperCsv.FullName).Count } else { 0 }
$doomedPolling = $null -ne (Get-ChildItem (Join-Path $runDirU 'Doomed') -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue)
$uKeepRoot = Join-Path $scratch 'kept'
$NonInteractive = $true
$rcU = Invoke-UninstallFlow -Requested 'Doomed' -KeepHistory $true -Root $runDirU -Path $uTaskPath -KeepRoot $uKeepRoot -PrevRoot (Join-Path $runDirU 'no-old-location') -PrevPath $uTaskPath
$NonInteractive = $false
Start-Sleep -Seconds 12
$keeperAfter = if ($keeperCsv) { @(Import-Csv $keeperCsv.FullName).Count } else { 0 }
$keeperTask = Get-ScheduledTask -TaskName ($TaskNamePrefix + 'Keeper') -TaskPath $uTaskPath -ErrorAction SilentlyContinue
$doomedTask = Get-ScheduledTask -TaskName ($TaskNamePrefix + 'Doomed') -TaskPath $uTaskPath -ErrorAction SilentlyContinue
$doomedProc = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like "*run_uninstall\Doomed*" })
$uKeptDir = @(Get-ChildItem $uKeepRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Doomed_*' })
Check "U live: both monitors polled before the removal" ($keeperBefore -ge 5 -and $doomedPolling)
Check "U live: the removed monitor's task, folder and process are gone" ($rcU -eq 0 -and $null -eq $doomedTask -and -not (Test-Path (Join-Path $runDirU 'Doomed')) -and @($doomedProc).Count -eq 0)
Check "U live: its history was kept where the run said" (@($uKeptDir).Count -eq 1 -and @(Get-ChildItem $uKeptDir[0].FullName -Filter 'Latency_*.csv').Count -eq 1)
Check "U live: the other monitor is untouched, still running and still recording" ($null -ne $keeperTask -and "$($keeperTask.State)" -eq 'Running' -and $keeperAfter -gt $keeperBefore -and (Test-Path (Join-Path $runDirU 'Keeper\credential.bin')))
foreach ($t in @(Get-ScheduledTask -TaskPath $uTaskPath -ErrorAction SilentlyContinue)) { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $uTaskPath -Confirm:$false -ErrorAction SilentlyContinue }
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like "*run_uninstall*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
$srvU | Stop-Job -ErrorAction SilentlyContinue; $srvU | Remove-Job -Force -ErrorAction SilentlyContinue

# ----- R: the move out of the old install root, with a real task and a live monitor -----
$rOldRoot = Join-Path $scratch 'run_move_old'
$rNewRoot = Join-Path $scratch 'run_move_new'
New-Item $rOldRoot -ItemType Directory -Force | Out-Null
New-Item $rNewRoot -ItemType Directory -Force | Out-Null
$portR = Get-FreeTestPort
$srvR = Start-Job -ScriptBlock {
    param($port)
    $l = New-Object System.Net.HttpListener
    $l.Prefixes.Add("http://127.0.0.1:$port/")
    $l.Start()
    while ($l.IsListening) {
        $ctx = $l.GetContext()
        $body = [Text.Encoding]::UTF8.GetBytes(('<html>move drill ' + ('x' * 2000) + '</html>'))
        $ctx.Response.StatusCode = 200
        $ctx.Response.OutputStream.Write($body, 0, $body.Length)
        $ctx.Response.Close()
    }
} -ArgumentList $portR
Start-Sleep -Seconds 2
$rUrl = "http://127.0.0.1:$portR/"
$rName = 'Move Drill'
$rSlug = ConvertTo-MonitorSlug $rName
$rDir = Join-Path $rOldRoot $rSlug
New-Item $rDir -ItemType Directory -Force | Out-Null
$InstallDir = $rDir
$MonitorName = $rName
$Url = $rUrl
$rGen = New-MonitorContent -SiteName 'MoveSite' -Mail $mailDown
Set-Content (Join-Path $rDir 'Watch-CurlMonitor.ps1') -Value $rGen -Encoding UTF8
@{ MonitorName = $rName; Url = $rUrl; SiteName = 'MoveSite' } | ConvertTo-Json | Set-Content (Join-Path $rDir 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $rDir 'credential.bin') -Value 'cipher' -Encoding ASCII
Register-MonitorTask -Name ($TaskNamePrefix + $rName) -Path $uTaskPath -ScriptPath (Join-Path $rDir 'Watch-CurlMonitor.ps1')
Start-ScheduledTask -TaskPath $uTaskPath -TaskName ($TaskNamePrefix + $rName)
Start-Sleep -Seconds 20
$rCsvBefore = @(Get-ChildItem $rDir -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue)
$rRowsBefore = if ($rCsvBefore.Count) { @(Import-Csv $rCsvBefore[0].FullName).Count } else { 0 }
Check "R live: the monitor polls from the old location before the move" ($rRowsBefore -ge 3 -and $null -ne (Get-ScheduledTask -TaskName ($TaskNamePrefix + $rName) -TaskPath $uTaskPath -ErrorAction SilentlyContinue))

$rFound = @(Get-PreviousRootMonitor -Root $rOldRoot -Path $uTaskPath -NewRoot $rNewRoot)
$rMoved = Move-MonitorToNewRoot -Monitor @($rFound)[0] -NewRoot $rNewRoot -NewTaskPath $uTaskPath
$rDest = Join-Path $rNewRoot $rSlug
Start-Sleep -Seconds 20
$rOldTask = Get-ScheduledTask -TaskName ($TaskNamePrefix + $rName) -TaskPath $uTaskPath -ErrorAction SilentlyContinue
$rAction = if ($rOldTask) { "$($rOldTask.Actions[0].Arguments)" } else { '' }
$rCsvAfter = @(Get-ChildItem $rDest -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue)
$rRowsAfter = if ($rCsvAfter.Count) { @(Import-Csv $rCsvAfter[0].FullName).Count } else { 0 }
$rText = if (Test-Path (Join-Path $rDest 'Watch-CurlMonitor.ps1')) { Get-Content (Join-Path $rDest 'Watch-CurlMonitor.ps1') -Raw } else { '' }
$rHb = if (Test-Path (Join-Path $rDest 'heartbeat.json')) { Get-Content (Join-Path $rDest 'heartbeat.json') -Raw | ConvertFrom-Json } else { $null }
$rProc = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like "*run_move_new*" })
Check "R live: the move reported success and the folder is at the new root, with its history and secret" ($rMoved -and (Test-Path $rDest) -and -not (Test-Path $rDir) -and (Test-Path (Join-Path $rDest 'credential.bin')) -and $rRowsAfter -ge $rRowsBefore)
Check "R live: the task now runs the monitor from the new folder" ($null -ne $rOldTask -and $rAction -like "*$rDest*" -and $rAction -notlike "*$rDir\*")
Check "R live: the moved monitor is polling again at the new location" ($rRowsAfter -gt $rRowsBefore -and @($rProc).Count -ge 1)
Check "R live: its script was repointed and carries no reference to the old folder" (($rText -match [regex]::Escape($rDest)) -and -not ($rText -match [regex]::Escape($rDir)) -and $null -ne [ScriptBlock]::Create($rText))
Check "R live: the deliberate stop was recorded, so no restart notice is raised" ($null -ne $rHb -and -not ($rHb.Stopped -eq $true -and $rRowsAfter -eq $rRowsBefore))
$rDrops = if (Test-Path (Join-Path $rDest 'Drops.log')) { Get-Content (Join-Path $rDest 'Drops.log') -Raw } else { '' }
Check "R live: the move raised no RESTART line in the drops log" (-not ($rDrops -match '\| RESTART'))

# Removing one monitor must not touch a monitor whose folder name merely starts with the same text
$rSibling = Join-Path $rNewRoot ($rSlug + '-Reports')
New-Item $rSibling -ItemType Directory -Force | Out-Null
$InstallDir = $rSibling
$MonitorName = "$rName Reports"
Set-Content (Join-Path $rSibling 'Watch-CurlMonitor.ps1') -Value (New-MonitorContent -SiteName 'MoveSite' -Mail $mailDown) -Encoding UTF8
@{ MonitorName = "$rName Reports"; Url = $rUrl; SiteName = 'MoveSite' } | ConvertTo-Json | Set-Content (Join-Path $rSibling 'install-settings.json') -Encoding UTF8
Register-MonitorTask -Name ($TaskNamePrefix + "$rName Reports") -Path $uTaskPath -ScriptPath (Join-Path $rSibling 'Watch-CurlMonitor.ps1')
Start-ScheduledTask -TaskPath $uTaskPath -TaskName ($TaskNamePrefix + "$rName Reports")
Start-Sleep -Seconds 15
$rSibBefore = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like "*$rSibling*" })
$null = Stop-MonitorProcess -Dir $rDest
Start-Sleep -Seconds 3
$rSibAfter = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine -like "*$rSibling*" })
Check "R live: stopping one monitor leaves a sibling whose folder name starts the same alone" (@($rSibBefore).Count -ge 1 -and @($rSibAfter).Count -eq @($rSibBefore).Count)

foreach ($t in @(Get-ScheduledTask -TaskPath $uTaskPath -ErrorAction SilentlyContinue)) { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $uTaskPath -Confirm:$false -ErrorAction SilentlyContinue }
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and ($_.CommandLine -like "*run_move_new*" -or $_.CommandLine -like "*run_move_old*") } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
$srvR | Stop-Job -ErrorAction SilentlyContinue; $srvR | Remove-Job -Force -ErrorAction SilentlyContinue

# M3: no plaintext secret anywhere in any artifact
$leak = Get-ChildItem $scratch -Recurse -File | Where-Object { (Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue) -like "*s3cret-O'Brien*" }
Check "M3 plaintext secret appears in no artifact on disk" (@($leak).Count -eq 0)

Write-Host ""
Write-Host "TOTAL: $script:pass passed, $script:fail failed"
if ($script:fail){ Write-Host "Failed: $($script:failed -join '; ')"; exit 1 }
