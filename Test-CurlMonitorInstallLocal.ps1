<#
.SYNOPSIS
    Runs Install-CurlMonitor.ps1 for real on this machine, twice, and verifies everything it leaves behind.

.DESCRIPTION
    Run from an elevated PowerShell window. The wizard answers are typed into this console's input buffer, so the
    unmodified installer runs exactly as an operator would run it:
    1. Fresh install with the Graph app details supplied here. The installer sends its real test email.
    2. After a settle period the live output is checked: the task runs as SYSTEM at startup with no time limit,
       the monitor process is owned by SYSTEM, the credential file is locked to SYSTEM and Administrators and
       decrypts to the secret supplied here, the install folder is locked down, and the latency CSV shows polls
       that followed the endpoint's redirects to a populated sign-in page.
    3. Re-install over the running task with saved settings (Enter through every prompt except the secret) and
       confirm the old monitor stopped, the task is running again, and the CSV kept its schema.
    4. Removes the task and install folder unless -KeepInstalled. Copies of the logs and CSVs go to -OutDir.

.REQUIREMENTS
    Elevated PowerShell 5.1 or 7. Internet access to the tenant and the HST endpoint.

.OUTPUTS
    PASS/FAIL lines, a summary count, result.json and the installer logs in -OutDir. Exit code 1 on any failure.

.NOTES
    Author:      Christopher Carroll
    Created:     09/04/2026
    Idempotency: Safe to re-run. Each run starts from whatever the previous one left and cleans up after itself.
    Context:     Static checks on the installer passed while three production defects went unnoticed. This test
                 exists so the installer is exercised on a real machine before it goes to a site.
#>

[CmdletBinding()]
param(
    [string]$InstallerPath = (Join-Path $PSScriptRoot 'Install-CurlMonitor.ps1'),
    [string]$SiteName = 'LocalTest',
    [Parameter(Mandatory)][string]$Sender,
    [Parameter(Mandatory)][string]$Recipient,
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [securestring]$ClientSecret,
    [string]$SecretExpires = '',
    [int]$SettleSeconds = 45,
    [switch]$KeepInstalled,
    [string]$OutDir = (Join-Path $env:TEMP 'HSTMonitorLocalTest')
)

$ErrorActionPreference = 'Stop'
$InstallDir = 'C:\ProgramData\CurlMonitor'
$TaskName = 'HST eChart Monitor'
$TaskPath = '\CurlMonitor\'
$script:pass = 0; $script:fail = 0; $script:failed = @()
$R = [ordered]@{ Started = (Get-Date).ToString('o'); Checks = @() }

function Check([string]$Name, [bool]$Condition, [string]$Detail = '') {
    if ($Condition) { $script:pass++; Write-Host "PASS  $Name" }
    else { $script:fail++; $script:failed += $Name; Write-Host "FAIL  $Name  $Detail" }
    $script:R.Checks += [ordered]@{ Name = $Name; Pass = $Condition; Detail = $Detail }
}
function Save { $script:R.Ended = (Get-Date).ToString('o'); $script:R | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $OutDir 'result.json') -Encoding UTF8 }

$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $elevated) { Write-Host "Run this from an elevated PowerShell window."; exit 1 }
if (-not (Test-Path $InstallerPath)) { Write-Host "Installer not found at '$InstallerPath'."; exit 1 }
if ($SiteName -notmatch '^[A-Za-z0-9][A-Za-z0-9 _-]*[A-Za-z0-9]$') { Write-Host "SiteName must be letters, numbers, spaces, dashes, or underscores with no leading or trailing space."; exit 1 }
if (-not $ClientSecret) { $ClientSecret = Read-Host "Client secret (input hidden)" -AsSecureString }
$plainSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))
if ([string]::IsNullOrWhiteSpace($plainSecret)) { Write-Host "A client secret is required."; exit 1 }
New-Item -Path $OutDir -ItemType Directory -Force | Out-Null

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ConIn {
    [StructLayout(LayoutKind.Sequential)]
    public struct KEY_EVENT_RECORD { public int bKeyDown; public ushort wRepeatCount; public ushort wVirtualKeyCode; public ushort wVirtualScanCode; public ushort UnicodeChar; public uint dwControlKeyState; }
    [StructLayout(LayoutKind.Explicit)]
    public struct INPUT_RECORD { [FieldOffset(0)] public ushort EventType; [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent; }
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr CreateFileW(string name, uint access, uint share, IntPtr sec, uint disp, uint flags, IntPtr tmpl);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteConsoleInputW(IntPtr h, INPUT_RECORD[] recs, int n, out int written);
    [DllImport("user32.dll")] public static extern short VkKeyScanW(char ch);
    [DllImport("user32.dll")] public static extern uint MapVirtualKeyW(uint code, uint mapType);
    public static int Send(string text) {
        IntPtr h = CreateFileW("CONIN$", 0xC0000000u, 3u, IntPtr.Zero, 3u, 0u, IntPtr.Zero);
        if (h == IntPtr.Zero || h.ToInt64() == -1) throw new Exception("CONIN$ open failed, error " + Marshal.GetLastWin32Error());
        var list = new System.Collections.Generic.List<INPUT_RECORD>();
        foreach (char c in text) {
            ushort vk; ushort ch = (ushort)c; uint ctrl = 0;
            if (c == '\r') continue;
            if (c == '\n') { vk = 0x0D; ch = 0x0D; }
            else { short scan = VkKeyScanW(c); vk = (ushort)(scan & 0xFF); if (((scan >> 8) & 1) == 1) ctrl = 0x10; }
            INPUT_RECORD down = new INPUT_RECORD();
            down.EventType = 1; down.KeyEvent.bKeyDown = 1; down.KeyEvent.wRepeatCount = 1; down.KeyEvent.wVirtualKeyCode = vk;
            down.KeyEvent.wVirtualScanCode = (ushort)MapVirtualKeyW(vk, 0); down.KeyEvent.UnicodeChar = ch; down.KeyEvent.dwControlKeyState = ctrl;
            INPUT_RECORD up = down; up.KeyEvent.bKeyDown = 0;
            list.Add(down); list.Add(up);
        }
        INPUT_RECORD[] arr = list.ToArray(); int total = 0;
        for (int i = 0; i < arr.Length; i += 200) {
            int n = Math.Min(200, arr.Length - i);
            INPUT_RECORD[] chunk = new INPUT_RECORD[n]; Array.Copy(arr, i, chunk, 0, n);
            int written;
            if (!WriteConsoleInputW(h, chunk, n, out written)) throw new Exception("WriteConsoleInput failed, error " + Marshal.GetLastWin32Error());
            total += written;
        }
        return total;
    }
}
"@

function Invoke-DrivenInstall([string[]]$Answers, [string]$LogName, [int]$TimeoutSec = 900) {
    # Types the answers first, then runs the installer in this console with its output captured to a log
    [void][ConIn]::Send((($Answers -join "`n") + "`n"))
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$InstallerPath`""
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.RedirectStandardInput = $false
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $p = [System.Diagnostics.Process]::Start($psi)
    $o = $p.StandardOutput.ReadToEndAsync(); $e = $p.StandardError.ReadToEndAsync()
    $done = $p.WaitForExit($TimeoutSec * 1000)
    if (-not $done) { try { $p.Kill() } catch { } }
    $log = Join-Path $OutDir $LogName
    Set-Content -Path $log -Value ($o.Result + "`n----- STDERR -----`n" + $e.Result) -Encoding UTF8
    return [ordered]@{ ExitCode = $p.ExitCode; TimedOut = (-not $done); Seconds = [int]$sw.Elapsed.TotalSeconds; Log = $log; Text = $o.Result }
}

function Get-MonitorProcesses {
    @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -match 'Watch-HSTeChartUptime\.ps1' } | ForEach-Object {
        $own = Invoke-CimMethod -InputObject $_ -MethodName GetOwner
        [PSCustomObject]@{ Pid = $_.ProcessId; Owner = "$($own.Domain)\$($own.User)" }
    })
}

function Get-Task { Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue }

try {
    $answers1 = @($SiteName, '1', $Sender, $Recipient, 'Y', $TenantId, $ClientId, $plainSecret, $SecretExpires, 'Y')
    $answers2 = @('', '', '', '', '', '', '', $plainSecret, '', 'Y')

    Write-Host ""
    Write-Host "Run 1: fresh install"
    $run1 = Invoke-DrivenInstall -Answers $answers1 -LogName 'install-run1.log'
    $R.Run1 = @{ ExitCode = $run1.ExitCode; TimedOut = $run1.TimedOut; Seconds = $run1.Seconds }
    Check "Run 1 exit code 0" ($run1.ExitCode -eq 0 -and -not $run1.TimedOut) "exit $($run1.ExitCode) timedOut=$($run1.TimedOut)"
    Check "Run 1 reports install complete with no FAILED lines" ($run1.Text -match 'Install complete for site' -and -not ($run1.Text -match 'FAILED \|'))
    Check "Run 1 test email accepted by Graph" ($run1.Text -match 'The mail host accepted the test message')
    Check "Run 1 preflight reached the sign-in page through 2 redirects" ($run1.Text -match 'reached the sign-in page.*HTTP 200 after 2 redirect')

    $t = Get-Task
    Check "Task exists" ($null -ne $t)
    if ($t) {
        $i = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath $TaskPath
        Check "Task is running" ([string]$t.State -eq 'Running') "state $($t.State)"
        Check "Task runs as SYSTEM, service account, highest" ($t.Principal.UserId -match 'SYSTEM|S-1-5-18' -and [string]$t.Principal.LogonType -eq 'ServiceAccount' -and [string]$t.Principal.RunLevel -eq 'Highest') "$($t.Principal.UserId) $($t.Principal.LogonType) $($t.Principal.RunLevel)"
        Check "Task has no time limit, single instance, restarts 3 x 1 min" ([string]$t.Settings.ExecutionTimeLimit -eq 'PT0S' -and [string]$t.Settings.MultipleInstances -eq 'IgnoreNew' -and $t.Settings.RestartCount -eq 3 -and [string]$t.Settings.RestartInterval -eq 'PT1M')
        Check "Task starts at boot and on battery" ((($t.Triggers | ForEach-Object { $_.CimClass.CimClassName }) -join ',') -match 'BootTrigger' -and -not $t.Settings.DisallowStartIfOnBatteries -and -not $t.Settings.StopIfGoingOnBatteries)
        Check "Task action runs the deployed monitor hidden" ($t.Actions[0].Execute -eq 'powershell.exe' -and $t.Actions[0].Arguments -match '-WindowStyle Hidden' -and $t.Actions[0].Arguments -match [regex]::Escape("$InstallDir\Watch-CurlMonitor.ps1"))
        Check "Task last result 0 or still running" ($i.LastTaskResult -in @(0, 267009)) "last result $($i.LastTaskResult)"
    }
    $procs1 = Get-MonitorProcesses
    Check "Monitor process running as SYSTEM" ($procs1.Count -eq 1 -and $procs1[0].Owner -match 'SYSTEM') ("$($procs1 | ForEach-Object { "$($_.Owner) pid=$($_.Pid)" })")

    $settingsPath = Join-Path $InstallDir 'install-settings.json'
    $credPath = Join-Path $InstallDir 'credential.bin'
    $monPath = Join-Path $InstallDir 'Watch-CurlMonitor.ps1'
    Check "Settings file present with InstalledAt and no secret" ((Test-Path $settingsPath) -and ((Get-Content $settingsPath -Raw) -match '"InstalledAt"') -and -not ((Get-Content $settingsPath -Raw).Contains($plainSecret)))
    if (Test-Path $settingsPath) {
        $s = Get-Content $settingsPath -Raw | ConvertFrom-Json
        Check "Settings hold the wizard answers" ($s.SiteName -eq $SiteName -and $s.MailMethod -eq 'Graph' -and $s.MailFrom -eq $Sender -and @($s.MailTo) -contains $Recipient -and $s.GraphTenantId -eq $TenantId -and $s.GraphClientId -eq $ClientId)
    }
    Check "Credential file present" (Test-Path $credPath)
    if (Test-Path $credPath) {
        $aclLines = @((& icacls.exe $credPath) | Where-Object { $_ -match ':\(' } | ForEach-Object { $_.Trim() })
        Check "Credential file ACL is exactly SYSTEM and Administrators" ($aclLines.Count -eq 2 -and ($aclLines -join ' ') -match 'NT AUTHORITY\\SYSTEM:\(F\)' -and ($aclLines -join ' ') -match 'BUILTIN\\Administrators:\(F\)') ($aclLines -join ' | ')
        $raw = (Get-Content $credPath -Raw).Trim()
        Add-Type -AssemblyName System.Security
        $decrypted = ''
        try { $decrypted = [Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($raw), [Text.Encoding]::UTF8.GetBytes('DIT-HSTMonitor-SMTP-v1'), [Security.Cryptography.DataProtectionScope]::LocalMachine)) } catch { }
        Check "Credential file decrypts to the supplied secret and holds no plaintext" ($decrypted -eq $plainSecret -and -not $raw.Contains($plainSecret))
    }
    Check "Monitor file present and parses" ((Test-Path $monPath) -and (& { $tk = $null; $er = $null; [void][System.Management.Automation.Language.Parser]::ParseFile($monPath, [ref]$tk, [ref]$er); $er.Count -eq 0 }))
    if (Test-Path $monPath) { Check "Monitor file holds no plaintext secret" (-not (Get-Content $monPath -Raw).Contains($plainSecret)) }
    $dirAcl = (& icacls.exe $InstallDir | Out-String)
    Check "Install folder locked: SYSTEM and Administrators full, Users read, owner Administrators" ($dirAcl -match 'NT AUTHORITY\\SYSTEM:\(OI\)\(CI\)\(F\)' -and $dirAcl -match 'BUILTIN\\Administrators:\(OI\)\(CI\)\(F\)' -and $dirAcl -match 'BUILTIN\\Users:\(OI\)\(CI\)\(RX\)' -and -not ($dirAcl -match '\(I\)') -and (Get-Acl $InstallDir).Owner -eq 'BUILTIN\Administrators')

    Write-Host ""
    Write-Host "Settling $SettleSeconds s for live polls"
    Start-Sleep -Seconds $SettleSeconds
    $csv = Get-ChildItem $InstallDir -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue | Select-Object -First 1
    Check "Latency CSV created" ($null -ne $csv)
    $rowsBefore = 0
    if ($csv) {
        $rows = @(Import-Csv $csv.FullName)
        $rowsBefore = $rows.Count
        Check "Latency CSV has the full schema" (($rows[0].PSObject.Properties.Name -join ',') -eq 'Timestamp_Local,Timestamp_UTC,SiteName,Url,HttpCode,CurlExit,Reason,DnsMs,ConnectMs,TlsMs,TtfbMs,TotalMs,SizeBytes,RemoteIp,Redirects,RedirectMs,FinalUrl,ContentOk')
        Check "Latency CSV has at least 3 polls" ($rows.Count -ge 3) "$($rows.Count) rows"
        $bad = @($rows | Where-Object { $_.HttpCode -ne '200' -or $_.CurlExit -ne '0' -or $_.ContentOk -ne 'True' -or $_.Redirects -ne '2' -or $_.FinalUrl -notmatch 'HSTFederationProvider' })
        Check "Every poll: 200, curl exit 0, populated, 2 redirects to the sign-in page" ($bad.Count -eq 0) "$($bad.Count) bad of $($rows.Count)"
        Check "Every poll carries timings and the backend IP" (@($rows | Where-Object { $_.TotalMs -notmatch '^\d+$' -or $_.RemoteIp -notmatch '\d' }).Count -eq 0)
        $R.SampleRow = $rows[-1]
    }
    $log = Get-ChildItem $InstallDir -Filter 'Transcript_*.log' -ErrorAction SilentlyContinue | Select-Object -First 1
    Check "Monitor transcript created" ($null -ne $log)
    if ($log) {
        $lt = Get-Content $log.FullName
        Check "Monitor logged SUCCESS polls and no failures, errors, or DOWN" ((@($lt | Where-Object { $_ -match 'SUCCESS \|' }).Count -ge 3) -and (@($lt | Where-Object { $_ -match 'FAILED \||ERROR \||\[DOWN\]' }).Count -eq 0))
    }
    Check "No probe body file left behind" (-not (Test-Path (Join-Path $InstallDir 'probe-body.tmp')))
    $drops = Join-Path $InstallDir 'Drops.log'
    Check "Drops log has the start line and no failures on a healthy endpoint" ((Test-Path $drops) -and ((Get-Content $drops -Raw) -match '\| START     \| Monitor started on') -and -not ((Get-Content $drops -Raw) -match '\| (FAIL|DOWN) '))
    Save

    Write-Host ""
    Write-Host "Run 2: re-install over the running task with saved settings"
    $run2 = Invoke-DrivenInstall -Answers $answers2 -LogName 'install-run2.log'
    $R.Run2 = @{ ExitCode = $run2.ExitCode; TimedOut = $run2.TimedOut; Seconds = $run2.Seconds }
    Check "Run 2 exit code 0" ($run2.ExitCode -eq 0 -and -not $run2.TimedOut) "exit $($run2.ExitCode)"
    Check "Run 2 loaded saved settings and stopped the old monitor" ($run2.Text -match 'Loaded settings from a previous install' -and $run2.Text -match 'Stopped the existing monitor')
    Check "Run 2 reports install complete with no FAILED lines" ($run2.Text -match 'Install complete for site' -and -not ($run2.Text -match 'FAILED \|'))
    Start-Sleep -Seconds 15
    $t2 = Get-Task
    Check "Task running again after re-install" ($t2 -and [string]$t2.State -eq 'Running') "state $($t2.State)"
    $procs2 = Get-MonitorProcesses
    Check "Exactly one monitor process, new pid, SYSTEM" ($procs2.Count -eq 1 -and $procs2[0].Owner -match 'SYSTEM' -and ($procs1.Count -eq 0 -or $procs2[0].Pid -ne $procs1[0].Pid)) ("$($procs2 | ForEach-Object { "$($_.Owner) pid=$($_.Pid)" })")
    $csvs = @(Get-ChildItem $InstallDir -Filter 'Latency_*.csv' -ErrorAction SilentlyContinue)
    Check "Still one latency CSV, schema unchanged, rows kept growing" ($csvs.Count -eq 1 -and @(Import-Csv $csvs[0].FullName).Count -gt $rowsBefore)
    Save
}
catch {
    Check "Test script ran to completion" $false "$($_.Exception.Message) @ line $($_.InvocationInfo.ScriptLineNumber)"
}
finally {
    if (-not $KeepInstalled) {
        Write-Host ""
        Write-Host "Cleanup"
        try { Stop-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue } catch { }
        Start-Sleep -Seconds 3
        Get-MonitorProcesses | ForEach-Object { Stop-Process -Id $_.Pid -Force -ErrorAction SilentlyContinue }
        try { Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop } catch { }
        if (Test-Path $InstallDir) {
            $art = Join-Path $OutDir 'artifacts'
            New-Item $art -ItemType Directory -Force | Out-Null
            Get-ChildItem $InstallDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'credential.bin' } | Copy-Item -Destination $art -Force -ErrorAction SilentlyContinue
            Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        Check "Cleanup removed the task and the install folder" ($null -eq (Get-Task) -and -not (Test-Path $InstallDir))
    }
    Save
    Write-Host ""
    Write-Host "TOTAL: $script:pass passed, $script:fail failed"
    if ($script:fail) { Write-Host "Failed: $($script:failed -join '; ')" }
    Write-Host "Logs and result.json in $OutDir"
}
if ($script:fail) { exit 1 }
