$ErrorActionPreference = 'Stop'
$script:pass=0; $script:fail=0; $script:failed=@()
function Check($n,$c){ if($c){$script:pass++} else {$script:fail++; $script:failed += $n; Write-Host "FAIL: $n"} }
function Section($n){ Write-Host ""; Write-Host "== $n ==" }

$installerPath = '/tmp/Install-CurlMonitor.ps1'
$src = Get-Content -Raw $installerPath
$T=$null;$E=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($installerPath,[ref]$T,[ref]$E)

Section "A. Installer parse and static checks"
Check "Installer parses with zero errors" ($E.Count -eq 0)
if ($E.Count){ $E | % { Write-Host ("  " + $_.Message + " @ " + $_.Extent.StartLineNumber) } }
Check "No em dashes anywhere" (-not ($src -match ([char]0x2014)))
Check "No PS7-only ?? operator" (-not ($src -match '\?\?'))
Check "No PS7-only && / || chain operators" (-not ($src -match '\s&&\s|\s\|\|\s'))
Check "No banner-style comment headers" (-not ($src -match '(?m)^\s*#\s*[-=#*]{5,}'))
Check "Remove-Variable * immediately after help (installer)" ($src -match '(?s)#>\s*\r?\n\s*\r?\nRemove-Variable \* -ErrorAction SilentlyContinue')
Check "Author name present" ($src -match 'Author:\s+Christopher Carroll')

# Approved verbs on every function in installer + template
$hereNodes = $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -eq 'SingleQuotedHereString'},$true)
Check "Exactly one embedded template" ($hereNodes.Count -eq 1)
$template = $hereNodes[0].Value
$tT=$null;$tE=$null
$tAst=[System.Management.Automation.Language.Parser]::ParseInput($template,[ref]$tT,[ref]$tE)
$approved = (Get-Verb).Verb
$allFuncs = @($ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)) + @($tAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true))
$badVerbs = $allFuncs | Where-Object { ($_.Name -split '-')[0] -notin $approved } | Select-Object -ExpandProperty Name -Unique
Check "All function verbs approved ($($allFuncs.Count) functions)" ($badVerbs.Count -eq 0)
if ($badVerbs){ Write-Host "  bad: $($badVerbs -join ', ')" }

# Every Write-Log -Level literal is in the approved keyword set
$levels = @('ADDED','CREATED','ERROR','FAILED','FINISHED','FOUND','INFORMATIONAL','PROMPT','SANITY CHECK','STARTED','SUCCESS','TOTAL','WARNING')
$badLevels = @()
foreach ($a in @($ast,$tAst)) {
  foreach ($c in $a.FindAll({param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Write-Log'},$true)) {
    for ($i=0;$i -lt $c.CommandElements.Count;$i++) {
      $el=$c.CommandElements[$i]
      if ($el -is [System.Management.Automation.Language.CommandParameterAst] -and $el.ParameterName -eq 'Level') {
        $v=$c.CommandElements[$i+1]
        if ($v -is [System.Management.Automation.Language.StringConstantExpressionAst]) { if ($v.Value -notin $levels){ $badLevels += $v.Value } }
      }
    }
  }
}
Check "All Write-Log levels approved" ($badLevels.Count -eq 0)

# Every email subject literal starts with [HST
$subjects = [regex]::Matches($src,'"\[[A-Z ]+\][^"]*"') | % { $_.Value }
Check "Found subject literals ($($subjects.Count))" ($subjects.Count -ge 5)
$subjectTypes = $subjects | % { ([regex]::Match($_,'\[([A-Z ]+)\]')).Groups[1].Value } | Sort-Object -Unique
Check "Subject set is exactly DAILY/DOWN/MONITOR INSTALLED/MONITOR RESTARTED/MONITOR TEST/RESOLVED/SLOW/SLOW RESOLVED/STILL DOWN/STILL SLOW" ((($subjectTypes -join '|')) -eq 'DAILY|DOWN|MONITOR INSTALLED|MONITOR RESTARTED|MONITOR TEST|RESOLVED|SLOW|SLOW RESOLVED|STILL DOWN|STILL SLOW')
Check "Every subject carries the site name variable" (($subjects | Where-Object { $_ -notmatch '\$SiteName|\$site' }).Count -eq 0)

Check "Single deliverable: only Install-CurlMonitor.ps1 among installer files" ((Get-ChildItem /mnt/user-data/outputs -Filter '*CurlMonitor*.ps1').Count -eq 1)
Check "No dedicatedit.com or DIT-specific addresses anywhere" (-not ($src -match 'dedicatedit'))
Check "Sender and recipients have no baked defaults" ($src -match '(?m)^\$MailFrom\s+=\s+""' -and $src -match '(?m)^\$MailTo\s+=\s+@\(\)')
Check "Policy group derived from sender domain" ($src -match '\$policyGroup = "\$GraphPolicyGroupAlias@" \+')
Check "No reference to a second script" (-not ($src -match 'New-HSTMonitorAppRegistration'))
Check "Tenant setup lives inside installer" ($src -match 'function New-TenantMailApp' -and $src -match 'function Test-RequiredModule')
Check "Graph wizard branch calls tenant setup" ($src -match [regex]::Escape('New-TenantMailApp -SenderAddress $from -SiteName $SiteName'))
Check "Access policy RestrictAccess applied" ($src -match 'New-ApplicationAccessPolicy' -and $src -match 'RestrictAccess')
Check "Mail.Send role looked up, not hardcoded" ($src -match "Value -eq 'Mail.Send'" -and -not ($src -match 'b633e1c5-b582-4048-a93e-9f11b44c7e96'))
Check "Secret never written to disk in plaintext" (-not ($src -match 'Set-Content[^\n]*\$secret\b') -and -not ($src -match 'Set-Content[^\n]*\$graphSecret'))
Check "Exchange module session disconnected in finally" ($src -match 'finally \{ if \(\$script:ExoMode -eq .module.\) \{ try \{ Disconnect-ExchangeOnline')
Check "Setup module removed automatically after setup" ($src -match 'function Remove-TenantSetupModules' -and $src.Contains('{ Remove-TenantSetupModules }'))
Check "No Microsoft.Graph module dependency" (-not ($src -match 'Microsoft\.Graph\.|Connect-MgGraph|Get-Mg|New-Mg|Add-Mg'))
Check "Only ExchangeOnlineManagement is installed" (([regex]::Matches($src,"Test-RequiredModule -Name '([^']+)'") | % { $_.Groups[1].Value } | Sort-Object -Unique) -join ',' -eq 'ExchangeOnlineManagement')
$moduleBody = $src.Substring($src.IndexOf('function Test-RequiredModule'), $src.IndexOf('function ConvertTo-Base64Url') - $src.IndexOf('function Test-RequiredModule'))
Check "Module install and removal have no Y/N prompts" (-not ($moduleBody -match 'Read-Choice') -and $moduleBody -match 'Install-Module' -and $moduleBody -match 'Uninstall-Module')
Check "Browser auth-code sign-in for Graph, device code as fallback" ($src -match 'oauth2/v2\.0/authorize' -and $src -match 'code_challenge_method=S256' -and $src -match 'oauth2/v2\.0/devicecode')
Check "Graph setup uses REST endpoints" ($src.Contains('/addPassword') -and $src.Contains('appRoleAssignments'))
$rawReadHost = [regex]::Matches($src,'(?m)^\s*\$\w+\s*=\s*Read-Host\b(?![^\n]*-AsSecureString)') | ? { $_.Value -notmatch '\$entry\s*=' }
Check "All plain prompts route through Read-Setting/Read-Choice" ($rawReadHost.Count -eq 0)
Check "No plaintext-to-SecureString cmdlet anywhere" (-not ($src -match 'ConvertTo-SecureString -String') -and -not ($src -match '-AsPlainText'))
Check "ConvertTo-SecureText defined once in the installer and once in the monitor" (([regex]::Matches($src, '(?m)^function ConvertTo-SecureText \{')).Count -eq 2)
Check "Protect-Secret takes the prompt SecureString or in-memory text" ($src -match "ParameterSetName = 'Secure'\)\]\[securestring\]\`$Password" -and $src -match "ParameterSetName = 'Plain'\)\]\[string\]\`$PlainText")
Check "Drops log gets failed and slow polls, transitions, alerts, starts, restarts, never a healthy poll" ($template -match "Write-DropLog -Kind 'FAIL'" -and $template -match "Write-DropLog -Kind 'SLOW'" -and $template -match "Write-DropLog -Kind 'DOWN'" -and $template -match "Write-DropLog -Kind 'REMINDER'" -and $template -match "Write-DropLog -Kind 'RESOLVED'" -and $template -match "Write-DropLog -Kind 'RESTART'" -and $template -match "Write-DropLog -Kind 'CARRYOVER'" -and $template -match "Write-DropLog -Kind 'ALERT'" -and $template -match "Write-DropLog -Kind 'START'" -and $template -match "Write-DropLog -Kind 'STOP'" -and $template -match '(?m)^\s*else \{ Write-Log -Level SUCCESS -Message \$summary \}\s*$')
Check "Drops log is not pruned with the daily transcripts" ($template -match "Filter 'Transcript_\*\.log'" -and $template -match "Drops\.log" -and -not ($template -match "Filter 'HST-eChart-\*"))
Check "Stored secret offered only for the same Graph app" ($src -match "if \(\`$Saved -and \`$Saved\.MailMethod -eq 'Graph' -and \`$Saved\.CredentialFor -eq \`$client\) \{ Get-StoredSecret \}")
Check "Installer decrypts the stored secret in one place" (([regex]::Matches($src, 'ProtectedData\]::Unprotect')).Count -eq 2 -and $src -match 'function Get-StoredSecret')
Check "Analyzer settings file present and names only warning-level style rules" ((Test-Path 'C:\Workspaces\HST Monitor\PSScriptAnalyzerSettings.psd1') -and -not ((Get-Content 'C:\Workspaces\HST Monitor\PSScriptAnalyzerSettings.psd1' -Raw) -match 'SecureString|PlainText|Credential|Password'))
Check "No parameter shadows an automatic variable" (-not ($src -match '(?i)\[string\]\$(Sender|Event|Args|Input|Matches|Error|Host|PID|Profile)'))
Check "Non-interactive Graph refused (secret needs console)" ($src -match "Graph cannot be configured non-interactively")
Check "Probe follows redirects with a bounded hop count" ($template -match '-L --max-redirs \$MaxRedirects' -and $src -match '(?m)^\$MaxRedirects\s+=\s+5')
Check "Install-time preflight follows redirects too" ($src -match 'function Test-EndpointReachable[\s\S]*?-L --max-redirs \$MaxRedirects')
Check "URL, monitor name, and content marker are prompted, not hardcoded" ($src -match '(?m)^\$Url\s+=\s+""\s' -and $src -match '(?m)^\$MonitorName\s+=\s+""\s' -and $src -match '(?m)^\$ExpectedContentMarker\s+=\s+""\s' -and $src -match 'function Get-MonitorUrl' -and $src -match 'function Get-MonitorName' -and $src -match 'function Get-ContentMarker')
Check "Task cmdlets stop on error and registration is verified" ($src -match 'Register-ScheduledTask[^\n]*-ErrorAction Stop' -and $src -match 'Get-ScheduledTask -TaskName \$TaskName -TaskPath \$TaskPath -ErrorAction Stop' -and $src -match 'Start-ScheduledTask -TaskName \$TaskName -TaskPath \$TaskPath -ErrorAction Stop')
$installTail = $src.Substring($src.IndexOf('$MonitorName = Get-MonitorName'))
Check "Existing task replaced in place, unregister only for the older install" ((-not ($installTail -match 'Unregister-ScheduledTask')) -and $src -match '(?s)function Invoke-LegacyMigration.*?Unregister-ScheduledTask' -and $src -match 'Register-ScheduledTask[^\n]*-Force')
Check "Monitor staged as .new and swapped in after verification" ($src -match '\$stagedPath = "\$monitorPath\.new"' -and $src -match 'Move-Item -Path \$stagedPath -Destination \$monitorPath -Force -ErrorAction Stop')
Check "InstalledAt recorded only after the task is running" ($src -match "(?s)if \(\`$taskState -eq 'Running'\) \{\s*\`$settings\['InstalledAt'\]")
Check "Monitor failure test includes the curl exit code" ($template -match '\$failed\s+=\s+\(\$result\.CurlExit -ne 0\) -or')
Check "Latency CSV written after alert dispatch" ($template -match "(?s)Send-AlertOrQueue -Subject \`$decision\.EmailSubject.*?Write-CsvRow -Path \(Get-LatencyCsvPath\) -Row \`$result")
Check "CSV write problems are logged, never rethrown" ($template -match "Could not append to")
Check "Graph JSON bodies declare UTF-8" ((([regex]::Matches($src, "application/json; charset=utf-8")).Count -ge 4) -and -not ($src -match "ContentType 'application/json'[^;]"))
Check "Graph send with no secret fails fast instead of falling through to SMTP" ($src -match 'No Graph client secret is available')
Check "Template substitution is a single pass" ($src -match "\[regex\]::Replace\(\`$Template_MonitorScript, '@@\(\[A-Z\]\+\)@@'" -and -not ($src -match "\.Replace\('@@SITENAME@@'"))
Check "Install ends non-zero when the task is not running" ($src -match "if \(\`$taskState -ne 'Running'\) \{[\s\S]*?exit 1")
$topReturns = @($ast.EndBlock.Statements | Where-Object { $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] } | ForEach-Object { $_.FindAll({param($n) $n -is [System.Management.Automation.Language.ReturnStatementAst]},$true) })
Check "Installer failure paths exit non-zero (no top-level return)" ($topReturns.Count -eq 0)
$tTopReturns = @($tAst.EndBlock.Statements | Where-Object { $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] } | ForEach-Object { $_.FindAll({param($n) $n -is [System.Management.Automation.Language.ReturnStatementAst]},$true) })
Check "Monitor failure paths exit non-zero (no top-level return)" ($tTopReturns.Count -eq 0)
Check "Windows PowerShell drops PowerShell 7 module paths before loading modules" ($src -match "if \(\`$PSVersionTable\.PSVersion\.Major -le 5\) \{\s*\`$env:PSModulePath = ")
Check "Probe body written under InstallDir, not the global temp folder" ($template -match "Join-Path \`$InstallDir 'probe-body\.tmp'" -and -not ($template -match 'GetTempFileName'))

Section "B. Load installer helpers and generate monitors across a value matrix"
# Define config vars from the installer's assignment statements (top-level scalar assigns only, before functions)
$firstFunc = ($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] } | Select-Object -First 1).Extent.StartLineNumber
$assigns = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.AssignmentStatementAst] -and $_.Extent.StartLineNumber -lt $firstFunc }
foreach ($a in $assigns) { if ($a.Left.Extent.Text -ne '$Template_MonitorScript') { Invoke-Expression $a.Extent.Text } }
$Template_MonitorScript = $template
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('Write-Log','ConvertTo-SafeSiteName','Test-EmailAddress','ConvertTo-RecipientList','ConvertTo-DerivedDirectSendHost','New-MonitorContent','Read-Setting','Read-Choice','Read-PortSetting','Test-GuidLike','Test-DateInput','ConvertTo-GraphMailBody','New-AlertBody') }) { Invoke-Expression $f.Extent.Text }
function Write-Log { param($Level,$Message) }   # silence during tests

function Gen($site,$mail){ New-MonitorContent -SiteName $site -Mail $mail }
function ParseOk($text){ $x=$null;$y=$null; [void][System.Management.Automation.Language.Parser]::ParseInput($text,[ref]$x,[ref]$y); return ($y.Count -eq 0) }
function ConfigValue($text,$var){ $x=$null;$y=$null; $a=[System.Management.Automation.Language.Parser]::ParseInput($text,[ref]$x,[ref]$y); $n=$a.FindAll({param($k) $k -is [System.Management.Automation.Language.AssignmentStatementAst] -and $k.Left.Extent.Text -eq $var},$false) | Select-Object -First 1; Invoke-Expression $n.Right.Extent.Text }

$mDirect = @{ MailMethod='DirectSend'; SmtpServer='dedicatedit-com.mail.protection.outlook.com'; SmtpPort=25; SmtpUseSsl=$false; MailFrom='hst-monitor@dedicatedit.com'; MailTo=@('alerts@dedicatedit.com'); SmtpAuthUser=''; CipherText='' }
$mRelay  = @{ MailMethod='Relay'; SmtpServer='10.1.1.5'; SmtpPort=2525; SmtpUseSsl=$true; MailFrom='noc@dedicatedit.com'; MailTo=@('a@dedicatedit.com','b@dedicatedit.com','c@dedicatedit.com'); SmtpAuthUser=''; CipherText='' }
$mGraph  = @{ MailMethod='Graph'; SmtpServer='graph.microsoft.com'; SmtpPort=443; SmtpUseSsl=$true; MailFrom='hst-monitor@dedicatedit.com'; MailTo=@('alerts@dedicatedit.com','noc@dedicatedit.com'); SmtpAuthUser=''; CipherText='AAAA'; GraphTenantId='11111111-2222-3333-4444-555555555555'; GraphClientId='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; GraphSecretExpires='2028-09-04' }
$mAuth   = @{ MailMethod='Authenticated'; SmtpServer='smtp.office365.com'; SmtpPort=587; SmtpUseSsl=$true; MailFrom="o'brien@dedicatedit.com"; MailTo=@("d'arcy@dedicatedit.com"); SmtpAuthUser="o'brien@dedicatedit.com"; CipherText='AAAA' }

$matrix = @(
  @{ site='CapCity'; mail=$mDirect; marker='' },
  @{ site='Jersey Shore'; mail=$mRelay; marker='HST eChart Login' },
  @{ site="O'Brien Site"; mail=$mAuth; marker="Bob's Login" },
  @{ site='  weird!!  name   '; mail=$mDirect; marker='' },
  @{ site='Café-Test_1'; mail=$mRelay; marker='<title>x</title>' },
  @{ site='X'; mail=$mAuth; marker='' },
  @{ site='GraphSite'; mail=$mGraph; marker='' }
)
$i=0
foreach ($m in $matrix) {
  $i++
  $ExpectedContentMarker = $m.marker
  $g = Gen $m.site $m.mail
  Check "M$i parses ($($m.site))" (ParseOk $g)
  Check "M$i no leftover tokens" (-not ($g -match '@@[A-Z]+@@'))
  Check "M$i site round-trips" ((ConfigValue $g '$SiteName') -eq (ConvertTo-SafeSiteName $m.site))
  Check "M$i marker round-trips" ((ConfigValue $g '$ExpectedContentMarker') -eq $m.marker)
  Check "M$i MailTo count" ((@(ConfigValue $g '$MailTo')).Count -eq $m.mail.MailTo.Count)
  Check "M$i MailFrom round-trips" ((ConfigValue $g '$MailFrom') -eq $m.mail.MailFrom)
  Check "M$i AuthUser round-trips" ((ConfigValue $g '$SmtpAuthUser') -eq $m.mail.SmtpAuthUser)
  Check "M$i SSL bool" ((ConfigValue $g '$SmtpUseSsl') -eq $m.mail.SmtpUseSsl)
  Check "M$i Port int" ((ConfigValue $g '$SmtpPort') -eq $m.mail.SmtpPort)
  if ($m.mail.MailMethod -eq 'Graph') {
    Check "M$i GraphTenant round-trips" ((ConfigValue $g '$GraphTenantId') -eq $m.mail.GraphTenantId)
    Check "M$i GraphClient round-trips" ((ConfigValue $g '$GraphClientId') -eq $m.mail.GraphClientId)
    Check "M$i Expiry round-trips"      ((ConfigValue $g '$GraphSecretExpires') -eq $m.mail.GraphSecretExpires)
    Check "M$i MailMethod Graph"        ((ConfigValue $g '$MailMethod') -eq 'Graph')
  }
}
$ExpectedContentMarker=''
# Config clamps
$DownThreshold=0; $IntervalSeconds=0; $TimeoutSeconds=-5
$g = Gen 'Clamp' $mDirect
Check "DownThreshold 0 clamps to 1" ((ConfigValue $g '$DownThreshold') -eq 1)
Check "Interval 0 clamps to 1"      ((ConfigValue $g '$IntervalSeconds') -eq 1)
Check "Timeout -5 clamps to 1"      ((ConfigValue $g '$TimeoutSeconds') -eq 1)
$DownThreshold=3; $IntervalSeconds=10; $TimeoutSeconds=15
$gen = Gen 'CapCity' $mDirect
Set-Content /tmp/generated_monitor.ps1 $gen
Check "Generated monitor: Remove-Variable * right after help" ($gen -match '(?s)#>\s*\r?\n\s*\r?\nRemove-Variable \* -ErrorAction SilentlyContinue')

Section "C. Site name, email, recipient, host helpers"
Check "Safe: strips punctuation" ((ConvertTo-SafeSiteName "O'Brien Site!!") -eq 'OBrien Site')
Check "Safe: collapses/trims spaces" ((ConvertTo-SafeSiteName '  a   b  ') -eq 'a b')
Check "Safe: keeps dash/underscore" ((ConvertTo-SafeSiteName 'Site_1-A') -eq 'Site_1-A')
Check "Safe: null -> empty" ((ConvertTo-SafeSiteName $null) -eq '')
Check "Safe: only junk -> empty" ((ConvertTo-SafeSiteName '!!!') -eq '')
Check "Email valid" (Test-EmailAddress 'a.b@c.io')
Check "Email invalid no @" (-not (Test-EmailAddress 'abc'))
Check "Email invalid no tld" (-not (Test-EmailAddress 'a@b'))
Check "Email invalid space" (-not (Test-EmailAddress 'a b@c.com'))
Check "Email null" (-not (Test-EmailAddress $null))
$r = ConvertTo-RecipientList 'a@x.com, B@x.com;a@X.COM  c@y.org bogus'
Check "Recipients: split, dedupe case-insensitive, drop invalid" (($r -join '|') -eq 'a@x.com|B@x.com|c@y.org')
Check "Recipients: empty -> 0" ((ConvertTo-RecipientList '').Count -eq 0)
Check "Recipients: single returns array" ((@(ConvertTo-RecipientList 'a@x.com')).Count -eq 1)
Check "DirectSend host derived" ((ConvertTo-DerivedDirectSendHost 'DedicatedIT.com') -eq 'dedicatedit-com.mail.protection.outlook.com')
Check "DirectSend subdomain" ((ConvertTo-DerivedDirectSendHost 'mail.foo.co.uk') -eq 'mail-foo-co-uk.mail.protection.outlook.com')
Check "Guid valid" (Test-GuidLike '11111111-2222-3333-4444-555555555555')
Check "Guid uppercase ok" (Test-GuidLike 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE')
Check "Guid invalid short" (-not (Test-GuidLike '1234'))
Check "Guid invalid braces" (-not (Test-GuidLike '{11111111-2222-3333-4444-555555555555}'))
Check "Date blank -> empty" ((Test-DateInput '') -eq '')
Check "Date valid" ((Test-DateInput ' 2028-09-04 ') -eq '2028-09-04')
Check "Date invalid -> null" ($null -eq (Test-DateInput '09/04/2028'))
Check "Date impossible -> null" ($null -eq (Test-DateInput '2028-13-40'))
$j = ConvertTo-GraphMailBody -Subject 'S' -Body "L1`nL2" -To @('a@x.com','b@x.com') | ConvertFrom-Json
Check "Graph body subject" ($j.message.subject -eq 'S')
Check "Graph body HTML type" ($j.message.body.contentType -eq 'HTML')
$ab = New-AlertBody -Heading "H <b>" -Details ([ordered]@{ 'Site' = "O'Brien & Co"; 'Server' = 'SRV1' })
Check "Alert body: HTML table, values encoded, footer names this server" ($ab -match '^<html>' -and $ab -match 'H &lt;b&gt;' -and $ab -match "O&#39;Brien &amp; Co" -and $ab -match '<td[^>]*>Server</td><td[^>]*>SRV1</td>' -and $ab -match [regex]::Escape("monitor on $env:COMPUTERNAME."))
Check "Both send paths deliver HTML" ((([regex]::Matches($src, 'BodyAsHtml = \$true')).Count -eq 2) -and -not ($src -match "contentType = 'Text'"))
Check "Graph body 2 recipients" ($j.message.toRecipients.Count -eq 2 -and $j.message.toRecipients[1].emailAddress.address -eq 'b@x.com')
Check "Graph body no sent items" ($j.saveToSentItems -eq $false)
$j1 = ConvertTo-GraphMailBody -Subject 'S' -Body 'B' -To @('a@x.com') | ConvertFrom-Json
Check "Graph body single recipient is array" (@($j1.message.toRecipients).Count -eq 1)
$NonInteractive=$true
Check "Read-Setting non-interactive returns default" ((Read-Setting -Prompt 'x' -Default 'dflt') -eq 'dflt')
Check "Read-Choice non-interactive returns default" ((Read-Choice -Prompt 'x' -Allowed @('Y','N') -Default 'N') -eq 'N')
$NonInteractive=$false

Section "D. Monitor functions: extract from generated monitor"
$gT=$null;$gE=$null
$gAst=[System.Management.Automation.Language.Parser]::ParseInput($gen,[ref]$gT,[ref]$gE)
foreach ($f in $gAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('ConvertTo-Ms','Format-Duration','Get-CurlReason','Get-LatencyCsvPath','Write-CsvRow','Get-TranscriptPath','Update-MonitorState','Get-ProbeResult','Get-SecretExpiryWarning','New-AlertBody','Write-Heartbeat','Read-Heartbeat','Get-RestartNotice','Send-AlertOrQueue','Send-PendingAlert','Wait-NetworkReady','Get-SmtpCredential','ConvertTo-SecureText','Write-DropLog','Update-SlowState','Test-DailySummaryDue','Get-DailySummary','Get-AlertLabel')},$true)) { Invoke-Expression $f.Extent.Text }
$InstallDir = '/tmp/hstprobe_test'; if (Test-Path $InstallDir){Remove-Item $InstallDir -Recurse -Force}; New-Item $InstallDir -ItemType Directory | Out-Null

Check "Ms: 0.123456 -> 123" ((ConvertTo-Ms '0.123456') -eq 123)
Check "Ms: garbage -> null" ($null -eq (ConvertTo-Ms 'abc'))
Check "Ms: empty -> null"   ($null -eq (ConvertTo-Ms ''))
Check "Ms: comma decimal rejected (invariant)" ($null -eq (ConvertTo-Ms '0,5'))
Check "Duration negative -> 00m 00s" ((Format-Duration ([TimeSpan]::FromSeconds(-5))) -eq '00m 00s')
Check "Duration 2d" ((Format-Duration ([TimeSpan]::FromSeconds(172800+3661))) -eq '2d 1h 01m 01s')
Check "CurlReason 28" ((Get-CurlReason 28) -eq 'Timed out')
Check "CurlReason 7" ((Get-CurlReason 7) -eq 'Connection refused or unreachable')
Check "CurlReason unknown" ((Get-CurlReason 99) -eq 'curl exit 99')
Check "CurlReason 47" ((Get-CurlReason 47) -eq 'Too many redirects')
Check "CurlReason 18" ((Get-CurlReason 18) -eq 'Transfer ended early (partial body)')
Check "Latency path monthly" ((Get-LatencyCsvPath) -match 'Latency_\d{6}\.csv$')
$today=[datetime]'2026-09-04'
Check "Expiry blank -> null" ($null -eq (Get-SecretExpiryWarning -ExpiresOn '' -Today $today))
Check "Expiry garbage -> null" ($null -eq (Get-SecretExpiryWarning -ExpiresOn 'soon' -Today $today))
Check "Expiry 31d -> null" ($null -eq (Get-SecretExpiryWarning -ExpiresOn '2026-10-05' -Today $today))
Check "Expiry 30d -> warn" ((Get-SecretExpiryWarning -ExpiresOn '2026-10-04' -Today $today) -match 'expires in 30 day')
Check "Expiry today -> warn 0" ((Get-SecretExpiryWarning -ExpiresOn '2026-09-04' -Today $today) -match 'expires in 0 day')
Check "Expiry yesterday -> EXPIRED" ((Get-SecretExpiryWarning -ExpiresOn '2026-09-03' -Today $today) -match 'EXPIRED')
Check "Transcript path daily"  ((Get-TranscriptPath) -match 'Transcript_\d{8}\.log$')

# CSV schema mismatch handling
$csv = Join-Path $InstallDir 'schema.csv'
Write-CsvRow -Path $csv -Row ([PSCustomObject]@{A=1;B=2})
Write-CsvRow -Path $csv -Row ([PSCustomObject]@{A=3;B=4})
Check "CSV appends 2 rows" ((Import-Csv $csv).Count -eq 2)
Write-CsvRow -Path $csv -Row ([PSCustomObject]@{A=5;C=6})
$aside = Get-ChildItem $InstallDir -Filter 'schema_schema-*.csv'
Check "CSV mismatch moves old file aside" ($aside.Count -eq 1)
Check "CSV mismatch starts new file with 1 row" (@(Import-Csv $csv).Count -eq 1)
Check "CSV creates missing folder" ( (Write-CsvRow -Path (Join-Path $InstallDir 'sub/x.csv') -Row ([PSCustomObject]@{A=1})) -eq $null -and (Test-Path (Join-Path $InstallDir 'sub/x.csv')) )
# A locked file: the row is skipped with a warning, nothing is moved aside, nothing throws
$lockCsv = Join-Path $InstallDir 'locked.csv'
Write-CsvRow -Path $lockCsv -Row ([PSCustomObject]@{A=1;B=2})
$fsLock = [System.IO.File]::Open($lockCsv, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message" }
$script:LastLog = ''
$threw = $false
try { Write-CsvRow -Path $lockCsv -Row ([PSCustomObject]@{A=3;B=4}) } catch { $threw = $true }
$fsLock.Close()
Check "CSV locked: no throw, warning logged, file kept in place" (-not $threw -and $script:LastLog -match '^WARNING\|Could not append' -and (Test-Path $lockCsv) -and @(Get-ChildItem $InstallDir -Filter 'locked_schema-*.csv').Count -eq 0)
Check "CSV locked: earlier rows intact, skipped row absent" (@(Import-Csv $lockCsv).Count -eq 1)
# A value containing a token is inserted verbatim, not substituted again
$gTok = New-MonitorContent -SiteName 'T' -Mail @{ MailMethod='Relay'; SmtpServer='host-with-@@URL@@-inside'; SmtpPort=25; SmtpUseSsl=$false; MailFrom='m@x.invalid'; MailTo=@('a@x.invalid'); SmtpAuthUser=''; CipherText='' }
Check "Template: value containing a token survives unchanged" ($gTok -match "SmtpServer\s+=\s+'host-with-@@URL@@-inside'" -and (ParseOk $gTok))
# Drops log
$DropLog = Join-Path $InstallDir 'drops.log'
$script:DropLogWarned = $false
Write-DropLog -Kind 'FAIL' -Message 'Code=000 Reason=Timed out'
Write-DropLog -Kind 'CARRYOVER' -Message 'carried'
$dl = @(Get-Content $DropLog)
Check "Drops log: one line per event, timestamp, padded kind, message" ($dl.Count -eq 2 -and $dl[0] -match '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d \| FAIL      \| Code=000 Reason=Timed out$' -and $dl[1] -match '\| CARRYOVER \| carried$')
$fsDrop = [System.IO.File]::Open($DropLog, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message" }
$script:LastLog = ''; $threw = $false
try { Write-DropLog -Kind 'FAIL' -Message 'locked' } catch { $threw = $true }
Write-DropLog -Kind 'FAIL' -Message 'locked again'
$fsDrop.Close()
Check "Drops log locked: no throw, one warning, lost lines counted once writable" (-not $threw -and $script:LastLog -match '^WARNING\|Could not write the drops log' -and (& { Write-DropLog -Kind 'FAIL' -Message 'after unlock'; $dl2 = @(Get-Content $DropLog); $dl2.Count -eq 4 -and $dl2[2] -match '\| LOST      \| 2 line\(s\) were not recorded while this file was locked' -and $dl2[3] -match '\| FAIL      \| after unlock$' }))
function Write-Log { param($Level,$Message) }
$big = New-Object byte[] (10MB + 1); [System.IO.File]::WriteAllBytes($DropLog, $big)
Write-DropLog -Kind 'STOP' -Message 'rotated'
$asideDrops = @(Get-ChildItem $InstallDir -Filter 'drops_*.log')
Check "Drops log rotates aside past 10 MB and starts fresh" ($asideDrops.Count -eq 1 -and $asideDrops[0].Length -gt 10MB -and @(Get-Content $DropLog).Count -eq 1 -and (Get-Content $DropLog) -match '\| STOP      \| rotated$')
function Write-Log { param($Level,$Message) }

Section "E. State machine"
function NewState { @{ ConsecutiveFailures=0; IsDown=$false; LastAlertUtc=$null; OutageStartUtc=$null; OutageStartLocalStr=$null; OutageStartUtcStr=$null } }
function Res($code,$ok,$reason='Timed out') { [PSCustomObject]@{ Timestamp_Local='L'; Timestamp_UTC='U'; HttpCode=$code; ContentOk=$ok; RemoteIp='1.2.3.4'; Reason=$reason } }
$T0 = [datetime]::SpecifyKind([datetime]'2026-09-04T12:00:00','Utc')
function Step { param([hashtable]$State,[bool]$Failed,[datetime]$Now,[int]$Dt=3,[int]$Re=30,[bool]$Aor=$true,$Result=$null)
  if ($null -eq $Result){ $Result = Res '000' $false }
  Update-MonitorState -State $State -Failed $Failed -Result $Result -NowUtc $Now -DownThreshold $Dt -ReAlertMinutes $Re -AlertOnRecovery $Aor -SiteName 'CapCity' -Url 'http://x' -HostName 'HOST1' -MonitorName 'eChart' }

$st=NewState; $any=$false
for($i=0;$i -le 5;$i++){ $d=Step $st $false $T0.AddSeconds($i*10); $st=$d.State; if($d.EmailSubject){$any=$true} }
Check "S1 steady up no emails" (-not $any -and -not $st.IsDown -and $st.ConsecutiveFailures -eq 0)

$st=NewState; $d=Step $st $true $T0; $st=$d.State
Check "S2 blip CF=1 not down no email" ($st.ConsecutiveFailures -eq 1 -and -not $st.IsDown -and $null -eq $d.EmailSubject)
$d=Step $st $false $T0.AddSeconds(10); $st=$d.State
Check "S2 blip recover: no record, no email, reset" ($null -eq $d.OutageRecord -and $null -eq $d.EmailSubject -and $st.ConsecutiveFailures -eq 0)

$st=NewState; $orig = $st.Clone()
$d1=Step $st $true $T0; $st=$d1.State; $d2=Step $st $true $T0.AddSeconds(10); $st=$d2.State; $d3=Step $st $true $T0.AddSeconds(20); $st=$d3.State
Check "S3 no email before threshold" ($null -eq $d1.EmailSubject -and $null -eq $d2.EmailSubject)
Check "S3 DOWN at 3rd, subject names site and server" ($d3.EmailSubject -eq '[DOWN] eChart at CapCity (HOST1) - unreachable')
Check "S3 DOWN body is HTML with server and reason rows" ($d3.EmailBody -match '^<html>' -and $d3.EmailBody -match '<td[^>]*>Server</td><td[^>]*>HOST1</td>' -and $d3.EmailBody -match '<td[^>]*>Reason</td><td[^>]*>Timed out</td>')
Check "S3 onset at first failure" ($st.OutageStartUtc -eq $T0)
Check "S3 input state not mutated" ($orig.ConsecutiveFailures -eq 0 -and -not $orig.IsDown)
$dr=Step $st $false $T0.AddSeconds(30) -Result (Res '200' $true 'OK'); $st=$dr.State
Check "S3 RESOLVED subject exact" ($dr.EmailSubject -eq '[RESOLVED] eChart at CapCity (HOST1) - outage lasted 00m 30s')
Check "S3 record fields" ($dr.OutageRecord.DurationSeconds -eq 30 -and $dr.OutageRecord.FailedPolls -eq 3 -and $dr.OutageRecord.RecoveryCode -eq '200' -and $dr.OutageRecord.Duration -eq '00m 30s')
Check "S3 state fully reset" (-not $st.IsDown -and $st.ConsecutiveFailures -eq 0 -and $null -eq $st.OutageStartUtc -and $null -eq $st.LastAlertUtc)

$st=NewState; for($i=0;$i -lt 3;$i++){ $d=Step $st $true $T0.AddSeconds($i*10); $st=$d.State }
$r1=Step $st $true $T0.AddSeconds(20).AddMinutes(29).AddSeconds(59); $st=$r1.State
Check "S4 no reminder at 29m59s" ($null -eq $r1.EmailSubject)
$r2=Step $st $true $T0.AddSeconds(20).AddMinutes(30); $st=$r2.State
Check "S4 reminder exactly at 30m" ($r2.EmailSubject -like '`[STILL DOWN`] eChart at CapCity (HOST1) - down for *')
Check "S4 reminder elapsed counts from onset (30m20s)" ($r2.EmailSubject -eq '[STILL DOWN] eChart at CapCity (HOST1) - down for 30m 20s')
$r3=Step $st $true $T0.AddSeconds(20).AddMinutes(31); $st=$r3.State
Check "S4 no premature second reminder" ($null -eq $r3.EmailSubject)
$r4=Step $st $true $T0.AddSeconds(20).AddMinutes(60); $st=$r4.State
Check "S4 second reminder at +60m" ($r4.EmailSubject -like '`[STILL DOWN`]*')
$fin=Step $st $false $T0.AddSeconds(20).AddMinutes(61) -Result (Res '200' $true 'OK'); $st=$fin.State
Check "S4 resolved 1h 01m 20s" ($fin.EmailSubject -eq '[RESOLVED] eChart at CapCity (HOST1) - outage lasted 1h 01m 20s')
Check "S4 failed polls counted through reminders (7)" ($fin.OutageRecord.FailedPolls -eq 7)

$st=NewState; for($i=0;$i -lt 3;$i++){ $d=Step $st $true $T0.AddSeconds($i*10) -Re 0; $st=$d.State }
$d=Step $st $true $T0.AddDays(3) -Re 0
Check "S5 ReAlert=0 never reminds" ($null -eq $d.EmailSubject)

$st=NewState; $d=Step $st $true $T0; $st=$d.State; $d=Step $st $false $T0.AddSeconds(5); $st=$d.State; $d=Step $st $true $T0.AddSeconds(10); $st=$d.State
Check "S6 flapping onset resets" ($st.OutageStartUtc -eq $T0.AddSeconds(10))

$d=Step (NewState) $true $T0 -Dt 1
Check "S7 threshold 1 immediate" ($d.EmailSubject -like '`[DOWN`]*')
$d=Step (NewState) $true $T0 -Dt 0
Check "S7 threshold 0 clamps to 1" ($d.EmailSubject -like '`[DOWN`]*')

$st=NewState; for($i=0;$i -lt 3;$i++){ $d=Step $st $true $T0.AddSeconds($i*10) -Aor $false; $st=$d.State }
$d=Step $st $false $T0.AddSeconds(30) -Aor $false
Check "S8 recovery alert off: record yes, email no" ($null -ne $d.OutageRecord -and $null -eq $d.EmailSubject)

$st=NewState; for($i=0;$i -lt 3;$i++){ $d=Step $st $true $T0.AddSeconds($i*10); $st=$d.State }
$d=Step $st $false $T0.AddDays(2).AddHours(3) -Result (Res '200' $true 'OK')
Check "S9 multi-day duration" ($d.EmailSubject -eq '[RESOLVED] eChart at CapCity (HOST1) - outage lasted 2d 3h 00m 00s')
Check "S9 multi-day seconds" ($d.OutageRecord.DurationSeconds -eq (2*86400+3*3600))

$corrupt = @{ ConsecutiveFailures=2; IsDown=$false; LastAlertUtc=$null; OutageStartUtc=$null; OutageStartLocalStr=$null; OutageStartUtcStr=$null }
$d=Step $corrupt $true $T0
Check "S10 corrupt state (CF>0, no onset) repairs onset" ($d.State.OutageStartUtc -eq $T0 -and $d.State.IsDown)

$d=Step (NewState) $true $T0 -Result ([PSCustomObject]@{Timestamp_Local='L';Timestamp_UTC='U';HttpCode='000';ContentOk=$false;RemoteIp=''})
Check "S11 result without Reason property tolerated" ($d.State.ConsecutiveFailures -eq 1)

$st=NewState; for($i=0;$i -lt 3;$i++){ $d=Step $st $true $T0.AddSeconds($i*10); $st=$d.State }
$d=Step $st $false $T0.AddSeconds(30).AddMilliseconds(600) -Result (Res '200' $true 'OK')
Check "S12 DurationSeconds truncates like the formatted duration (30.6 -> 30)" ($d.OutageRecord.DurationSeconds -eq 30)

# Long random soak: 20000 steps, invariants must always hold
$st=NewState; $rng=[Random]::new(42); $now=$T0; $emails=0; $records=0; $ok=$true
for($k=0;$k -lt 20000;$k++){
  $f = ($rng.NextDouble() -lt 0.3); $now=$now.AddSeconds(10)
  $d=Step $st $f $now; $st=$d.State
  if($d.EmailSubject){$emails++}; if($d.OutageRecord){$records++}
  if($st.IsDown -and $st.ConsecutiveFailures -lt 3){$ok=$false}
  if(-not $f -and $st.ConsecutiveFailures -ne 0){$ok=$false}
  if($st.ConsecutiveFailures -gt 0 -and $null -eq $st.OutageStartUtc){$ok=$false}
  if($d.OutageRecord -and $d.OutageRecord.FailedPolls -lt 3){$ok=$false}
  if($d.EmailSubject -and -not ($d.EmailSubject -match '^\[(DOWN|STILL DOWN|RESOLVED)\] eChart at CapCity \(HOST1\) - ')){$ok=$false}
}
Check "S13 20k-step random soak: invariants hold" $ok
Check "S13 soak produced outages and records" ($emails -gt 0 -and $records -gt 0)

Section "E2. Restart notice, heartbeat, delivery tracking, alert queue"
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message"; $script:Logs += "$Level|$Message" }
$script:Logs = @()
$HeartbeatFile = Join-Path $InstallDir 'heartbeat.json'
Check "R1 no previous heartbeat -> no notice" ($null -eq (Get-RestartNotice -Previous $null -NowUtc $T0 -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'))
Check "R1 missing heartbeat file -> null" ($null -eq (Read-Heartbeat))
$downState = @{ ConsecutiveFailures=7; IsDown=$true; LastAlertUtc=$T0.AddMinutes(5); OutageStartUtc=$T0; OutageStartLocalStr='2026-09-04 08:00:00'; OutageStartUtcStr='2026-09-04 12:00:00'; AlertDelivered=$false }
Write-Heartbeat -State $downState -NowUtc $T0.AddMinutes(10)
$hb = Read-Heartbeat
Check "R2 heartbeat roundtrip keeps every field and UTC kind" ($hb -and $hb.BeatUtc -eq $T0.AddMinutes(10) -and $hb.BeatUtc.Kind -eq 'Utc' -and $hb.IsDown -and $hb.ConsecutiveFailures -eq 7 -and $hb.OutageStartUtc -eq $T0 -and $hb.OutageStartUtc.Kind -eq 'Utc' -and $hb.LastAlertUtc -eq $T0.AddMinutes(5) -and $hb.OutageStartLocalStr -eq '2026-09-04 08:00:00' -and $hb.AlertDelivered -eq $false)
Check "R2 heartbeat written atomically: no temp file left, rename into place" (-not (Test-Path "$HeartbeatFile.tmp") -and $template -match 'Move-Item -Path \$tmp -Destination \$HeartbeatFile -Force -ErrorAction Stop')
Check "R2 heartbeat is compact JSON with ISO dates" (((Get-Content $HeartbeatFile -Raw) -match '"Beat":"2026-09-04T12:10:00\.0000000Z"') -and -not ((Get-Content $HeartbeatFile -Raw) -match 'Date\('))
$upState = @{ ConsecutiveFailures=0; IsDown=$false; LastAlertUtc=$null; OutageStartUtc=$null; OutageStartLocalStr=$null; OutageStartUtcStr=$null; AlertDelivered=$null }
Write-Heartbeat -State $upState -NowUtc $T0.AddMinutes(10)
$hbUp = Read-Heartbeat
Check "R2 up-state heartbeat has null outage fields" ($hbUp -and -not $hbUp.IsDown -and $null -eq $hbUp.OutageStartUtc -and $null -eq $hbUp.LastAlertUtc -and $null -eq $hbUp.AlertDelivered)
$n = Get-RestartNotice -Previous $hb -NowUtc $T0.AddMinutes(10).AddSeconds(30) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R3 short gap -> no email but the outage is carried over" ($n -and $null -eq $n.Subject -and $n.GapSeconds -eq 30 -and $n.RestoredState -and $n.RestoredState.IsDown -and $n.RestoredState.OutageStartUtc -eq $T0 -and $n.RestoredState.ConsecutiveFailures -eq 7 -and $n.RestoredState.AlertDelivered -eq $false -and $n.RestoredState.LastAlertUtc -eq $T0.AddMinutes(5))
$n = Get-RestartNotice -Previous $hb -NowUtc $T0.AddMinutes(10).AddSeconds(3840) -BootTimeUtc $T0.AddMinutes(40) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R4 long gap after a reboot -> subject names site, server, gap" ($n.Subject -eq '[MONITOR RESTARTED] CapCity (HOST1) - not running for 1h 04m 00s')
Check "R4 body says the server restarted after the heartbeat, outage carried, polls missed" ($n.Body -match '^<html>' -and $n.Body -match 'Server restarted at' -and $n.Body -match 'after the last heartbeat' -and $n.Body -match 'Yes\. It has been down since 2026-09-04 08:00:00 local' -and $n.Body -match '<td[^>]*>Polls missed</td><td[^>]*>384</td>' -and $n.Body -match '<td[^>]*>Server</td><td[^>]*>HOST1</td>' -and $n.RestoredState.IsDown)
$n = Get-RestartNotice -Previous $hbUp -NowUtc $T0.AddMinutes(10).AddSeconds(90) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R5 long gap without a reboot -> process stopped, nothing carried" ($n.Subject -eq '[MONITOR RESTARTED] CapCity (HOST1) - not running for 01m 30s' -and $n.Body -match 'Server did not restart' -and $n.Body -match 'None at the last heartbeat' -and $null -eq $n.RestoredState)
$n = Get-RestartNotice -Previous $hbUp -NowUtc $T0.AddMinutes(5) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R6 clock went backwards -> gap zero, no email" ($n -and $null -eq $n.Subject -and $n.GapSeconds -eq 0)
$n = Get-RestartNotice -Previous $hbUp -NowUtc $T0.AddMinutes(20) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R7 unknown boot time is stated, not guessed" ($n.Subject -and $n.Body -match 'Could not read the server boot time')
Check "R7 threshold boundary: gap equal to threshold reports" ((Get-RestartNotice -Previous $hbUp -NowUtc $T0.AddMinutes(11) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'C' -HostName 'H' -Url 'u').Subject -ne $null)
# Deliberate stop marked by the installer: no notice however long the gap, outage still carried over
Write-Heartbeat -State $downState -NowUtc $T0.AddMinutes(10)
$hbj = Get-Content $HeartbeatFile -Raw | ConvertFrom-Json
$hbj | Add-Member -NotePropertyName Stopped -NotePropertyValue $true -Force
$hbj | ConvertTo-Json -Compress | Set-Content -Path $HeartbeatFile -Encoding UTF8 -Force
$hbStopped = Read-Heartbeat
$n = Get-RestartNotice -Previous $hbStopped -NowUtc $T0.AddHours(30) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "R7b deliberate stop: heartbeat flag read back, no notice after 30 h, outage carried" ($hbStopped.Stopped -and $n -and $null -eq $n.Subject -and $n.Stopped -and $n.RestoredState -and $n.RestoredState.OutageStartUtc -eq $T0 -and $n.GapSeconds -eq 107400)
Write-Heartbeat -State $downState -NowUtc $T0.AddMinutes(10)
Check "R7b the monitor's own heartbeat writes Stopped false" (-not (Read-Heartbeat).Stopped)
Set-Content $HeartbeatFile -Value '{ not json' -Encoding UTF8
Check "R8 corrupt heartbeat -> null with warning" ($null -eq (Read-Heartbeat) -and $script:LastLog -match '^WARNING\|Heartbeat file .* is unreadable')
Set-Content $HeartbeatFile -Value '{"IsDown":true}' -Encoding UTF8
Check "R8 heartbeat without Beat -> null" ($null -eq (Read-Heartbeat))
Remove-Item $HeartbeatFile -Force
# Restored state flows through the state machine: first up poll resolves from the original onset and names the undelivered DOWN
$n = Get-RestartNotice -Previous $hb -NowUtc $T0.AddMinutes(70) -BootTimeUtc $T0.AddMinutes(40) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
$d = Step $n.RestoredState $false $T0.AddMinutes(70) -Result (Res '200' $true)
Check "R9 carried-over outage resolves with the true onset and duration" ($d.EmailKind -eq 'Resolved' -and $d.EmailSubject -eq '[RESOLVED] eChart at CapCity (HOST1) - outage lasted 1h 10m 00s' -and $d.OutageRecord.DurationSeconds -eq 4200 -and $d.OutageRecord.FailedPolls -eq 7 -and $d.OutageRecord.OutageStart_Local -eq '2026-09-04 08:00:00')
Check "R9 RESOLVED says the DOWN alert was never delivered" ($d.EmailBody -match '<td[^>]*>DOWN alert</td><td[^>]*>Not delivered\.' -and $null -eq $d.State.AlertDelivered)
$d = Step $n.RestoredState $true $T0.AddMinutes(70)
Check "R9 carried-over outage still down -> reminder from the old onset, first-notice row" ($d.EmailKind -eq 'Reminder' -and $d.EmailSubject -eq '[STILL DOWN] eChart at CapCity (HOST1) - down for 1h 10m 00s' -and $d.State.ConsecutiveFailures -eq 8 -and $d.EmailBody -match '<td[^>]*>Earlier alerts</td><td[^>]*>Not delivered\.')
$deliveredState = $n.RestoredState.Clone(); $deliveredState.AlertDelivered = $true
$d = Step $deliveredState $false $T0.AddMinutes(70) -Result (Res '200' $true)
Check "R9 RESOLVED after a delivered DOWN says so" ($d.EmailBody -match '<td[^>]*>DOWN alert</td><td[^>]*>Delivered</td>')
# Delivery tracking through DOWN
$st=NewState; $d=Step $st $true $T0; $st=$d.State; $d=Step $st $true $T0.AddSeconds(10); $st=$d.State; $d=Step $st $true $T0.AddSeconds(20)
Check "R10 DOWN decision carries kind and starts undelivered" ($d.EmailKind -eq 'Down' -and $d.State.AlertDelivered -eq $false)
$d2 = Step $d.State $false $T0.AddSeconds(30) -Result (Res '200' $true)
Check "R10 recovery clears delivery tracking" ($null -eq $d2.State.AlertDelivered -and $d2.EmailKind -eq 'Resolved')
# Alert queue with a stubbed sender
$script:PendingAlerts = New-Object System.Collections.ArrayList
$script:SendOk = $false; $script:Sent = @()
function Send-AlertEmail { param($Subject,$Body) $script:Sent += $Subject; return $script:SendOk }
$ok = Send-AlertOrQueue -Subject '[DOWN] x' -Body 'b' -Kind 'Down'
Check "Q1 failed send is queued and reported" (-not $ok -and $script:PendingAlerts.Count -eq 1 -and $script:LastLog -match 'will be retried every minute')
Check "Q1 retry within a minute does nothing" (@(Send-PendingAlert).Count -eq 0 -and $script:Sent.Count -eq 1)
$script:PendingAlerts[0].LastUtc = $script:PendingAlerts[0].LastUtc.AddSeconds(-61)
Check "Q1 retry after a minute while still failing keeps it queued" (@(Send-PendingAlert).Count -eq 0 -and $script:PendingAlerts.Count -eq 1 -and $script:Sent.Count -eq 2)
$script:PendingAlerts[0].LastUtc = $script:PendingAlerts[0].LastUtc.AddSeconds(-61); $script:SendOk = $true
$k = @(Send-PendingAlert)
Check "Q1 retry succeeds -> kind returned, queue empty" ($k.Count -eq 1 -and $k[0] -eq 'Down' -and $script:PendingAlerts.Count -eq 0 -and $script:LastLog -match 'Delivered on retry')
$script:SendOk = $false
Send-AlertOrQueue -Subject '[DOWN] x' -Body 'b' -Kind 'Down' | Out-Null
Send-AlertOrQueue -Subject '[MONITOR RESTARTED] x' -Body 'b' -Kind 'Restart' | Out-Null
Send-AlertOrQueue -Subject '[STILL DOWN] x' -Body 'b' -Kind 'Reminder' | Out-Null
Check "Q2 reminder supersedes the queued DOWN, restart notice kept" ($script:PendingAlerts.Count -eq 2 -and @($script:PendingAlerts | Where-Object { $_.Kind -eq 'Down' }).Count -eq 0 -and @($script:PendingAlerts | Where-Object { $_.Kind -eq 'Restart' }).Count -eq 1 -and ($script:Logs -join "`n") -match "Dropping undelivered '\[DOWN\] x', superseded by '\[STILL DOWN\] x'")
Send-AlertOrQueue -Subject '[RESOLVED] x' -Body 'b' -Kind 'Resolved' | Out-Null
Check "Q2 resolved supersedes the reminder, restart notice kept" ($script:PendingAlerts.Count -eq 2 -and @($script:PendingAlerts | Where-Object { $_.Kind -in @('Down','Reminder') }).Count -eq 0)
Send-AlertOrQueue -Subject '[MONITOR RESTARTED] y' -Body 'b' -Kind 'Restart' | Out-Null
Check "Q2 a newer restart notice replaces the older one" (@($script:PendingAlerts | Where-Object { $_.Kind -eq 'Restart' }).Count -eq 1 -and $script:PendingAlerts[-1].Subject -eq '[MONITOR RESTARTED] y')
foreach ($item in $script:PendingAlerts) { $item.FirstUtc = $item.FirstUtc.AddMinutes(-61); $item.LastUtc = $item.LastUtc.AddSeconds(-61) }
$k = @(Send-PendingAlert)
Check "Q3 gives up after an hour of failures" ($k.Count -eq 0 -and $script:PendingAlerts.Count -eq 0 -and ($script:Logs -join "`n") -match "Giving up on '\[RESOLVED\] x' after 60 minutes")
$script:SendOk = $false
1..7 | ForEach-Object { Send-AlertOrQueue -Subject "[DOWN] $_" -Body 'b' -Kind "K$_" | Out-Null }
Check "Q4 queue is capped at 5, oldest dropped" ($script:PendingAlerts.Count -eq 5 -and $script:PendingAlerts[0].Subject -eq '[DOWN] 3')
$script:PendingAlerts.Clear()
$script:SendOk = $true
Check "Q5 successful send is never queued" ((Send-AlertOrQueue -Subject 's' -Body 'b' -Kind 'Down') -and $script:PendingAlerts.Count -eq 0)
# SMTP credential rebuild never throws into the probe loop
function Get-ProtectedSecret { return $script:StoredSecret }
$SmtpAuthUser = 'svc@contoso.com'
$script:StoredSecret = ''
Check "C1 empty stored password -> null credential with a warning, no throw" ($null -eq (Get-SmtpCredential) -and $script:LastLog -match '^WARNING\|The stored SMTP password is empty')
$script:StoredSecret = $null
Check "C1 missing secret -> null credential" ($null -eq (Get-SmtpCredential))
$script:StoredSecret = "p@ss word'`u{00FC}"
$c1 = Get-SmtpCredential
Check "C1 stored password -> read-only credential with the exact password" ($c1 -and $c1.UserName -eq 'svc@contoso.com' -and $c1.GetNetworkCredential().Password -ceq $script:StoredSecret -and $c1.Password.IsReadOnly())
$SmtpAuthUser = ''
Check "C1 no auth user -> null credential" ($null -eq (Get-SmtpCredential))
Remove-Item function:Get-ProtectedSecret
Check "Monitor rebuilds the SMTP credential inside the send try block" ($template -match "(?s)try \{\s*\`$cred = Get-SmtpCredential\s*if \(\`$SmtpAuthUser -and -not \`$cred\)")
# Network wait
$sw = [Diagnostics.Stopwatch]::StartNew()
Check "N1 IP literal endpoint resolves instantly" ((Wait-NetworkReady -Url 'http://127.0.0.1:9/' -TimeoutSeconds 5) -and $sw.Elapsed.TotalSeconds -lt 3)
$sw.Restart()
Check "N1 unresolvable host gives up at the timeout with a warning" (-not (Wait-NetworkReady -Url 'http://nonexistent-host-zz.invalid/' -TimeoutSeconds 1) -and $sw.Elapsed.TotalSeconds -lt 40 -and $script:LastLog -match 'still fails after 1 s')
Check "N1 real endpoint resolves" (Wait-NetworkReady -Url 'https://prodasp09.hstpathways.com/p95_CSP/HSTeChart' -TimeoutSeconds 10)
# Static wiring
Check "Monitor writes a heartbeat before and after each poll" (([regex]::Matches($template, 'Write-Heartbeat -State \$state -NowUtc')).Count -eq 3)
Check "Monitor waits for name resolution before its first poll" ($template -match 'Wait-NetworkReady -Url \$Url')
Check "Monitor restart notice threshold is at least 60 s" ($template -match '\[math\]::Max\(60, 2 \* \(\$IntervalSeconds \+ \$TimeoutSeconds\)\)')
Check "Monitor Graph calls have a 30 s timeout" (([regex]::Matches($template, '-TimeoutSec 30')).Count -eq 2)
Check "Installer marks the heartbeat as a deliberate stop and logs STOP to the drops log" ($src -match "(?s)elseif \(\`$wasRunning\) \{.*?Add-Member -NotePropertyName Stopped -NotePropertyValue \`$true.*?Drops\.log.*?'STOP'")
Check "Probe cycle errors reach the drops log" ($template.Contains("Write-DropLog -Kind 'ERROR' -Message `"Probe cycle error"))
Check "Install email skipped when the SMTP password cannot be read back" ($src -match '\$canSend -and \(Send-MailWithConfig')
Check "Stored secret gated on the credential recorded at the last successful install" ($src -match "\`$Saved\.CredentialFor -eq \`$client\) \{ Get-StoredSecret \}" -and $src -match "\`$settings\['CredentialFor'\] = switch")
function Write-Log { param($Level,$Message) }

Section "E3. Slow periods and daily summary"
function SlowRes($ms, $local = 'L', $code = '200', $reason = 'OK') { [PSCustomObject]@{ Timestamp_Local = $local; Timestamp_UTC = 'U'; HttpCode = $code; ContentOk = ($code -eq '200'); RemoteIp = '1.2.3.4'; Reason = $reason; TotalMs = $ms } }
function NewSlow { @{ Samples = @(); IsSlow = $false; StartUtc = $null; StartLocalStr = $null; LastAlertUtc = $null; Polls = 0; SlowPolls = 0; FailedPolls = 0; WorstMs = 0; AlertDelivered = $null } }
function Pl($ms, $f = $false, $down = $false) { [PSCustomObject]@{ Ms = $ms; F = $f; D = $down } }
function RunSlow {
    param([hashtable]$State, [object[]]$Polls, [datetime]$Start, [int]$Spacing = 15, [int]$Re = 30, [bool]$Aor = $true)
    $steps = New-Object System.Collections.ArrayList; $st = $State; $i = 0
    foreach ($q in $Polls) {
        $now = $Start.AddSeconds($i * $Spacing)
        $res = if ($q.F) { SlowRes $null ("P{0:000}" -f $i) '000' 'Timed out' } else { SlowRes $q.Ms ("P{0:000}" -f $i) }
        $d = Update-SlowState -State $st -Result $res -Failed ([bool]$q.F) -IsDown ([bool]$q.D) -NowUtc $now -SlowThresholdMs 3000 -WindowMinutes 5 -AlertPercent 50 -ClearPercent 10 -ReAlertMinutes $Re -AlertOnRecovery $Aor -SiteName 'CapCity' -Url 'http://x' -HostName 'HOST1' -MonitorName 'eChart'
        $st = $d.State
        [void]$steps.Add([PSCustomObject]@{ I = $i; Now = $now; D = $d })
        $i++
    }
    [PSCustomObject]@{ State = $st; Steps = $steps; Emails = @($steps | Where-Object { $_.D.EmailSubject }) }
}
function Rep($poll, [int]$n) { @(1..$n | ForEach-Object { $poll }) }

$r = RunSlow -State (NewSlow) -Polls (Rep (Pl 300) 40) -Start $T0
Check "SL1 steady fast polls: no email, not slow, window holds at most 5 min of polls" ($r.Emails.Count -eq 0 -and -not $r.State.IsSlow -and $r.State.Samples.Count -eq 20)

$blips = @(); for ($i = 0; $i -lt 40; $i++) { $blips += $(if ($i % 10 -eq 5) { Pl 4500 } elseif ($i -eq 22) { Pl $null $true } else { Pl 300 }) }
$r = RunSlow -State (NewSlow) -Polls $blips -Start $T0
Check "SL2 isolated slow polls and one timeout: no email" ($r.Emails.Count -eq 0 -and -not $r.State.IsSlow)

$seq = @(Rep (Pl 300) 20) + @(Rep (Pl 4500) 140) + @(Rep (Pl 300) 25)
$seq[100] = Pl 9001
$r = RunSlow -State (NewSlow) -Polls $seq -Start $T0
$kinds = @($r.Emails | ForEach-Object { "$($_.I):$($_.D.EmailKind)" })
Check "SL3 sustained slowness: exactly SLOW, STILL SLOW, SLOW RESOLVED in that order" (($kinds -join ',') -eq '29:Slow,149:SlowReminder,177:SlowResolved')
$e = $r.Steps[29].D
Check "SL3 SLOW when half the 5 min window is slow, subject counts polls" ($e.EmailSubject -eq '[SLOW] eChart at CapCity (HOST1) - 10 of 20 polls slow or failed in 5 min' -and $e.DropKind -eq 'SLOWSTART' -and $e.State.StartLocalStr -eq 'P020' -and $e.State.StartUtc -eq $T0.AddSeconds(300))
Check "SL3 SLOW body: HTML table with window counts, timing, clear rule, server" ($e.EmailBody -match '^<html>' -and $e.EmailBody -match '<td[^>]*>Last 5 min</td><td[^>]*>20 polls in 5 min: 10 slower than 3000 ms, 0 failed</td>' -and $e.EmailBody -match 'median \d+ ms, worst 4500 ms' -and $e.EmailBody -match '<td[^>]*>Server</td><td[^>]*>HOST1</td>' -and $e.EmailBody -match '10% or fewer')
Check "SL3 log line for the drops log" ($e.TransitionLog -eq 'Declared SLOW for CapCity: 20 polls in 5 min: 10 slower than 3000 ms, 0 failed (median 300 ms, worst 4500 ms).')
$e = $r.Steps[149].D
Check "SL3 STILL SLOW after 30 min with elapsed since the first slow poll" ($e.EmailSubject -eq "[STILL SLOW] eChart at CapCity (HOST1) - slow for $(Format-Duration ([TimeSpan]::FromSeconds(149 * 15 - 300)))" -and $e.DropKind -eq 'SLOWSTILL')
Check "SL3 hysteresis: still slow at 40 percent, no email" ($r.Steps[171].D.State.IsSlow -and $null -eq $r.Steps[171].D.EmailSubject)
$e = $r.Steps[177].D
Check "SL3 SLOW RESOLVED once the share is 10 percent, timed from the first good poll after the last slow one" ($e.EmailSubject -eq "[SLOW RESOLVED] eChart at CapCity (HOST1) - slow period lasted $(Format-Duration ([TimeSpan]::FromSeconds(160 * 15 - 300)))" -and $e.DropKind -eq 'SLOWCLEAR' -and $e.EmailBody -match '<td[^>]*>Normal from</td><td[^>]*>P160 local</td>' -and $e.EmailBody -match "<td[^>]*>Duration</td><td[^>]*>$(Format-Duration ([TimeSpan]::FromSeconds(2100)))</td>" -and $e.EmailBody -match '<td[^>]*>Polls while slow</td><td[^>]*>140: 140 slower than 3000 ms, 0 failed</td>' -and $e.EmailBody -match '<td[^>]*>Worst response</td><td[^>]*>9001 ms</td>' -and $e.TransitionLog -match 'lasted 35m 00s, 140 polls, 140 slow, 0 failed, worst 9001 ms\.$')
Check "SL3 state reset after resolution" (-not $e.State.IsSlow -and $null -eq $e.State.StartUtc -and $e.State.Polls -eq 0 -and $e.State.WorstMs -eq 0)
Check "SL3 SLOW RESOLVED says whether the SLOW email was delivered" ($e.EmailBody -match '<td[^>]*>SLOW alert</td><td[^>]*>Not delivered')

$mixed = @(Rep (Pl 300) 20); for ($i = 0; $i -lt 30; $i++) { $mixed += $(switch ($i % 3) { 0 { Pl $null $true } 1 { Pl 4000 } 2 { Pl 300 } }) }
$r = RunSlow -State (NewSlow) -Polls $mixed -Start $T0
$first = @($r.Emails)[0]
Check "SL4 timeouts mixed with slow polls count toward SLOW" ($first.D.EmailKind -eq 'Slow' -and $first.D.State.FailedPolls -gt 0 -and $first.D.EmailBody -match '\d+ slower than 3000 ms, [1-9]\d* failed')

$seq = @(Rep (Pl 300) 20) + @(Rep (Pl 4500) 12) + @(Pl 300 $false $true) + @(Rep (Pl 300) 3) + @(Rep (Pl 4500) 3) + @(Rep (Pl 300) 20)
$r = RunSlow -State (NewSlow) -Polls $seq -Start $T0
$downStep = $r.Steps[32].D
Check "SL5 an outage closes the slow period with no email and empties the window" ($r.Steps[29].D.EmailKind -eq 'Slow' -and $null -eq $downStep.EmailSubject -and $downStep.DropKind -eq 'SLOWCLEAR' -and $downStep.TransitionLog -match 'ended in an outage' -and -not $downStep.State.IsSlow -and $downStep.State.Samples.Count -eq 0)
Check "SL5 slow polls right after the outage do not re-alert before half the window is covered" (@($r.Emails | Where-Object { $_.I -gt 32 }).Count -eq 0)

$seq = @(Rep (Pl 300) 20) + @(Rep (Pl 4500) 20) + @(Rep (Pl 300) 25)
$r = RunSlow -State (NewSlow) -Polls $seq -Start $T0 -Aor $false
Check "SL6 recovery alerts off: slow period still closes and is logged, no RESOLVED email" (((@($r.Emails | ForEach-Object { $_.D.EmailKind })) -join ',') -eq 'Slow' -and @($r.Steps | Where-Object { $_.D.DropKind -eq 'SLOWCLEAR' }).Count -eq 1)
$seq = @(Rep (Pl 300) 20) + @(Rep (Pl 4500) 200)
$r = RunSlow -State (NewSlow) -Polls $seq -Start $T0 -Re 0
Check "SL7 reminders off: one SLOW email in 50 minutes of slowness" ($r.Emails.Count -eq 1 -and $r.State.IsSlow)

$orig = NewSlow; $orig.Samples = @([PSCustomObject]@{ Utc = $T0; LocalStr = 'X'; Slow = $true; Failed = $false; Ms = 5000 })
$null = Update-SlowState -State $orig -Result (SlowRes 4000) -Failed $false -IsDown $false -NowUtc $T0.AddSeconds(15) -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "SL8 pure: the caller's state and window are not changed" (@($orig.Samples).Count -eq 1 -and -not $orig.IsSlow)

function NewCarried { $c = NewSlow; $c.IsSlow = $true; $c.StartUtc = $T0; $c.StartLocalStr = '2026-09-04 08:00:00'; $c.LastAlertUtc = $T0.AddMinutes(1); $c.Polls = 40; $c.SlowPolls = 30; $c.WorstMs = 7000; $c.AlertDelivered = $true; $c }
$r = RunSlow -State (NewCarried) -Polls (Rep (Pl 250) 12) -Start $T0.AddMinutes(20)
$e = @($r.Emails)[0]
Check "SL9 carried-over slow period waits for half the window, then resolves at the 11th fast poll with its true start" ($r.Emails.Count -eq 1 -and $e.I -eq 10 -and $e.D.EmailKind -eq 'SlowResolved' -and $e.D.EmailSubject -eq '[SLOW RESOLVED] eChart at CapCity (HOST1) - slow period lasted 20m 00s' -and $e.D.EmailBody -match '<td[^>]*>Slow from</td><td[^>]*>2026-09-04 08:00:00 local</td>' -and $e.D.EmailBody -match '<td[^>]*>Normal from</td><td[^>]*>P000 local</td>' -and $e.D.EmailBody -match '40: 30 slower than 3000 ms' -and $e.D.EmailBody -match '<td[^>]*>SLOW alert</td><td[^>]*>Delivered</td>')
$stillSlow = @(); for ($i = 0; $i -lt 40; $i++) { $stillSlow += $(if ($i % 3 -eq 0) { Pl 300 } else { Pl 4500 }) }
$r = RunSlow -State (NewCarried) -Polls $stillSlow -Start $T0.AddMinutes(20)
Check "SL9 carried-over period that is still slow: no SLOW RESOLVED and no second SLOW over 10 minutes, even with a fast first poll" ($r.Emails.Count -eq 0 -and $r.State.IsSlow -and $r.State.StartUtc -eq $T0)
$r = RunSlow -State (NewCarried) -Polls (@(Pl 300) + @(Rep (Pl 4500) 5)) -Start $T0.AddHours(3)
Check "SL9 carried-over period: no STILL SLOW before the window is covered again" ($r.Emails.Count -eq 0)
$prevSlow = [PSCustomObject]@{ BeatUtc = $T0; IsDown = $false; ConsecutiveFailures = 0; OutageStartUtc = $null; OutageStartLocalStr = ''; OutageStartUtcStr = ''; LastAlertUtc = $null; AlertDelivered = $null; Stopped = $false; IsSlow = $true; SlowStartLocalStr = '2026-09-04 07:55:00' }
$nSlow = Get-RestartNotice -Previous $prevSlow -NowUtc $T0.AddHours(2) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
$prevSlow.IsSlow = $false
$nFast = Get-RestartNotice -Previous $prevSlow -NowUtc $T0.AddHours(2) -BootTimeUtc $T0.AddHours(-5) -GapThresholdSeconds 60 -IntervalSeconds 10 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "SL9 restart notice after a long gap names the slow period that was in progress" ($nSlow.Body -match '<td[^>]*>Slow period in progress</td><td[^>]*>Yes, since 2026-09-04 07:55:00 local\. Slow tracking starts fresh\.</td>' -and -not ($nFast.Body -match 'Slow period in progress'))
Check "SL9 startup keeps a slow period only across a gap no longer than the window, and logs when it does not" ($template -match '(?s)if \(\$previous -and \$previous\.IsSlow -and \$null -ne \$previous\.SlowStartUtc -and -not \$state\.IsDown\) \{\s+# [^\n]*\n\s+if \(\$notice -and \$notice\.GapSeconds -le \(\$SlowWindowMinutes \* 60\)\) \{' -and $template.Contains("Write-DropLog -Kind 'SLOWCLEAR' -Message `"Slow period since `$(`$previous.SlowStartLocalStr) local not carried over"))

$slowHb = NewSlow; $slowHb.IsSlow = $true; $slowHb.StartUtc = $T0; $slowHb.StartLocalStr = '2026-09-04 08:00:00'; $slowHb.LastAlertUtc = $T0.AddMinutes(3); $slowHb.Polls = 25; $slowHb.SlowPolls = 14; $slowHb.FailedPolls = 2; $slowHb.WorstMs = 8123; $slowHb.AlertDelivered = $true
Write-Heartbeat -State $upState -NowUtc $T0.AddMinutes(10) -SlowState $slowHb
$hbS = Read-Heartbeat
Check "SL10 heartbeat roundtrip keeps the slow period" ($hbS.IsSlow -and $hbS.SlowStartUtc -eq $T0 -and $hbS.SlowStartUtc.Kind -eq 'Utc' -and $hbS.SlowStartLocalStr -eq '2026-09-04 08:00:00' -and $hbS.SlowLastAlertUtc -eq $T0.AddMinutes(3) -and $hbS.SlowPolls -eq 25 -and $hbS.SlowSlowPolls -eq 14 -and $hbS.SlowFailedPolls -eq 2 -and $hbS.SlowWorstMs -eq 8123 -and $hbS.SlowAlertDelivered -eq $true -and -not $hbS.IsDown)
Write-Heartbeat -State $upState -NowUtc $T0.AddMinutes(11)
Check "SL10 heartbeat without a slow state reads back as not slow" (-not (Read-Heartbeat).IsSlow)

$script:PendingAlerts = New-Object System.Collections.ArrayList; $script:SendOk = $false
Send-AlertOrQueue -Subject '[SLOW] a' -Body 'b' -Kind 'Slow' | Out-Null
Send-AlertOrQueue -Subject '[DAILY] a' -Body 'b' -Kind 'Summary' | Out-Null
Send-AlertOrQueue -Subject '[SLOW RESOLVED] a' -Body 'b' -Kind 'SlowResolved' | Out-Null
Check "SL11 SLOW RESOLVED supersedes an undelivered SLOW" (@($script:PendingAlerts | Where-Object { $_.Kind -eq 'Slow' }).Count -eq 0 -and @($script:PendingAlerts | Where-Object { $_.Kind -eq 'SlowResolved' }).Count -eq 1)
Send-AlertOrQueue -Subject '[SLOW] b' -Body 'b' -Kind 'Slow' | Out-Null
Send-AlertOrQueue -Subject '[DOWN] b' -Body 'b' -Kind 'Down' | Out-Null
Send-AlertOrQueue -Subject '[DAILY] b' -Body 'b' -Kind 'Summary' | Out-Null
Check "SL11 DOWN supersedes an undelivered SLOW, a newer summary replaces the older one" (@($script:PendingAlerts | Where-Object { $_.Kind -eq 'Slow' }).Count -eq 0 -and @($script:PendingAlerts | Where-Object { $_.Kind -eq 'Summary' }).Count -eq 1 -and @($script:PendingAlerts | Where-Object { $_.Kind -eq 'Summary' })[0].Subject -eq '[DAILY] b')
$script:PendingAlerts.Clear(); $script:SendOk = $true

$d7 = [datetime]'2026-09-15T07:00:00'
Check "DS1 summary due at the hour once per day, not before, not twice, never when off" ((Test-DailySummaryDue -NowLocal $d7 -Hour 7 -LastSentDate '2026-09-14') -and -not (Test-DailySummaryDue -NowLocal $d7.AddMinutes(-1) -Hour 7 -LastSentDate '2026-09-14') -and -not (Test-DailySummaryDue -NowLocal $d7.AddHours(5) -Hour 7 -LastSentDate '2026-09-15') -and (Test-DailySummaryDue -NowLocal $d7.AddHours(5) -Hour 7 -LastSentDate '') -and -not (Test-DailySummaryDue -NowLocal $d7 -Hour -1 -LastSentDate ''))

$lines = @(
    ([char]0xFEFF + '2026-09-14 06:59:59 | FAIL      | Site=CapCity Code=000 Redirects=0 TTFB=15013ms Total=15013ms Populated=False IP= Reason=Timed out'),
    '2026-09-14 07:00:00 | SLOW      | Site=CapCity Code=200 Redirects=2 TTFB=9900ms Total=9999ms Populated=True IP=1.2.3.4 Reason=OK',
    '2026-09-14 07:00:01 | START     | Monitor started on HOST1 for site CapCity.',
    '2026-09-15 06:16:14 | SLOW      | Site=CapCity Code=200 Redirects=2 TTFB=6007ms Total=6048ms Populated=True IP=50.19.13.93 Reason=OK',
    '2026-09-15 06:16:49 | SLOW      | Site=CapCity Code=200 Redirects=2 TTFB=8891ms Total=9001ms Populated=True IP=98.91.165.173 Reason=OK',
    '2026-09-15 06:18:49 | SLOW      | Site=CapCity Code=200 Redirects=2 TTFB=4303ms Total=4322ms Populated=True IP=184.73.88.34 Reason=OK',
    '2026-09-15 06:18:50 | SLOWSTART | Declared SLOW for CapCity: 20 polls in 5 min.',
    '2026-09-15 06:18:50 | ALERT     | Sent: [SLOW] eChart at CapCity (HOST1) - 10 of 20 polls slow or failed in 5 min',
    '2026-09-15 06:53:07 | FAIL      | Site=CapCity Code=000 Redirects=0 TTFB=15013ms Total=15013ms Populated=False IP= Reason=Timed out',
    '2026-09-15 06:55:00 | FAIL      | Site=CapCity Code=503 Redirects=2 TTFB=40ms Total=41ms Populated=False IP=1.2.3.4 Reason=HTTP 503',
    '2026-09-15 06:55:10 | DOWN      | Declared DOWN for CapCity after 3 consecutive failures (HTTP 503).',
    '2026-09-15 06:57:20 | RESOLVED  | Outage record written: CapCity lasted 02m 10s over 13 failed polls. Recovery HTTP 200 from 1.2.3.4.',
    '2026-09-15 06:58:00 | RESTART   | Monitor was not running for 5m 00s. Server restarted.',
    'garbage line without a stamp',
    '2026-09-15 07:00:01 | FAIL      | Site=CapCity Code=000 Reason=Timed out'
)
$sum = Get-DailySummary -Lines $lines -NowLocal $d7 -SlowThresholdMs 3000 -SiteName 'CapCity' -HostName 'HOST1' -Url 'http://x'
Check "DS2 summary counts only the 24 hours before the send, subject reads naturally" ($sum -and $sum.Subject -eq '[DAILY] CapCity (HOST1) - 3 slow polls, 2 failed polls, 1 outage in 24 hours' -and $sum.SlowPolls -eq 3 -and $sum.FailedPolls -eq 2 -and $sum.Outages -eq 1 -and $sum.SlowPeriods -eq 1 -and $sum.WorstMs -eq 9001)
Check "DS2 summary body: worst response, failure reasons, outage length, busiest hour, restarts" ($sum.Body -match '<td[^>]*>Slow polls</td><td[^>]*>3 slower than 3000 ms, worst 9001 ms</td>' -and $sum.Body -match '<td[^>]*>Failed polls</td><td[^>]*>2 \((Timed out x1, HTTP 503 x1|HTTP 503 x1, Timed out x1)\)</td>' -and $sum.Body -match '<td[^>]*>Outages</td><td[^>]*>1, lasting 02m 10s</td>' -and $sum.Body -match '<td[^>]*>Busiest hour</td><td[^>]*>06:00 to 06:59, 5 slow or failed polls</td>' -and $sum.Body -match '<td[^>]*>Monitor restarts</td><td[^>]*>1</td>' -and $sum.Body -match '<td[^>]*>Server</td><td[^>]*>HOST1</td>')
Check "DS3 nothing went wrong: no summary" ($null -eq (Get-DailySummary -Lines @('2026-09-15 06:00:00 | START     | x', '2026-09-15 06:30:00 | ALERT     | Sent: y', '2026-09-15 06:40:00 | STOP      | z') -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1') -and $null -eq (Get-DailySummary -Lines @() -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1'))
$ongoing = Get-DailySummary -Lines @('2026-09-14 06:00:00 | DOWN      | Declared DOWN for CapCity after 3 consecutive failures (Timed out).', '2026-09-15 06:30:00 | REMINDER  | Reminder raised for CapCity, down for 24h 30m.', '2026-09-15 06:40:00 | FAIL      | Site=CapCity Code=000 Reason=Timed out') -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1'
Check "DS5 an outage that began before the window and is still going counts once, marked in progress" ($ongoing.Outages -eq 1 -and $ongoing.Subject -match '1 outage in 24 hours$' -and $ongoing.Body -match '<td[^>]*>Outages</td><td[^>]*>1, one still in progress</td>')
$crossing = Get-DailySummary -Lines @('2026-09-14 06:50:00 | DOWN      | Declared DOWN for CapCity after 3 consecutive failures (Timed out).', '2026-09-14 07:20:00 | RESOLVED  | Outage record written: CapCity lasted 30m 00s over 180 failed polls. Recovery HTTP 200 from 1.2.3.4.') -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1'
Check "DS6 an outage that crossed the start of the window counts once with its length" ($crossing.Outages -eq 1 -and $crossing.Body -match '<td[^>]*>Outages</td><td[^>]*>1, lasting 30m 00s</td>')
$carriedOut = Get-DailySummary -Lines @('2026-09-15 05:00:00 | CARRYOVER | Outage in progress since 2026-09-14 05:00:00 local carried over from before the restart (40 failed polls so far).', '2026-09-15 05:10:00 | CARRYOVER | Slow period since 2026-09-15 04:55:00 local carried over from before the restart.', '2026-09-15 05:20:00 | RESOLVED  | Outage record written: CapCity lasted 24h 20m 00s over 900 failed polls. Recovery HTTP 200 from 1.2.3.4.') -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1'
Check "DS6 an outage carried over a restart counts once, a carried slow period is not an outage" ($carriedOut.Outages -eq 1 -and $carriedOut.Body -match '<td[^>]*>Outages</td><td[^>]*>1, lasting 24h 20m 00s</td>')
$one = Get-DailySummary -Lines @('2026-09-15 06:53:07 | FAIL      | Site=CapCity Code=000 Reason=Timed out') -NowLocal $d7 -SiteName 'CapCity' -HostName 'HOST1'
Check "DS4 a single timeout is enough for a summary, singular wording" ($one.Subject -eq '[DAILY] CapCity (HOST1) - 0 slow polls, 1 failed poll, 0 outages in 24 hours' -and $one.Body -match '<td[^>]*>Slow polls</td><td[^>]*>None</td>')

$genSlow = New-MonitorContent -SiteName 'S' -Mail $mRelay
Check "SL12 generated monitor bakes the slow alert and summary settings" ($genSlow -match '(?m)^\$AlertOnSlow\s+=\s+\$true$' -and $genSlow -match '(?m)^\$SlowWindowMinutes\s+=\s+5$' -and $genSlow -match '(?m)^\$SlowAlertPercent\s+=\s+50$' -and $genSlow -match '(?m)^\$SlowClearPercent\s+=\s+10$' -and $genSlow -match '(?m)^\$DailySummaryHour\s+=\s+7$' -and $genSlow -match "(?m)^\`$SummaryStateFile\s+=\s+'.*\\summary-sent\.txt'$")
$saveSlow = @($AlertOnSlow, $SlowWindowMinutes, $SlowAlertPercent, $SlowClearPercent, $DailySummaryHour)
$AlertOnSlow = $false; $SlowWindowMinutes = 0; $SlowAlertPercent = 5; $SlowClearPercent = 20; $DailySummaryHour = 30
$genClamp = New-MonitorContent -SiteName 'S' -Mail $mRelay
$DailySummaryHour = -9
$genOff = New-MonitorContent -SiteName 'S' -Mail $mRelay
$AlertOnSlow, $SlowWindowMinutes, $SlowAlertPercent, $SlowClearPercent, $DailySummaryHour = $saveSlow
Check "SL12 settings are clamped: window at least 1, clear below alert, hour 23 or off" ($genClamp -match '(?m)^\$AlertOnSlow\s+=\s+\$false$' -and $genClamp -match '(?m)^\$SlowWindowMinutes\s+=\s+1$' -and $genClamp -match '(?m)^\$SlowAlertPercent\s+=\s+5$' -and $genClamp -match '(?m)^\$SlowClearPercent\s+=\s+4$' -and $genClamp -match '(?m)^\$DailySummaryHour\s+=\s+23$' -and $genOff -match '(?m)^\$DailySummaryHour\s+=\s+-1$' -and (ParseOk $genClamp))
Check "SL13 loop evaluates slowness after the outage decision, gated emails, heartbeat carries it" ($template -match '(?s)\$decision = Update-MonitorState.*?\$slow = Update-SlowState -State \$slowState -Result \$result -Failed \$failed -IsDown \$state\.IsDown' -and $template -match 'if \(\$SendEmail -and \$AlertOnSlow -and \$slow\.EmailSubject\)' -and ([regex]::Matches($template, 'Write-Heartbeat -State \$state -NowUtc [^\r\n]*-SlowState \$slowState')).Count -eq 3 -and $template -match 'Write-DropLog -Kind \$slow\.DropKind -Message \$slow\.TransitionLog')
Check "SL13 daily summary reads the drops log once a day and records the date first" ($template -match '(?s)if \(Test-DailySummaryDue -NowLocal \(Get-Date\) -Hour \$DailySummaryHour -LastSentDate \$summarySent\) \{\s+\$summarySent = \(Get-Date\)\.ToString\(''yyyy-MM-dd''\)\s+try \{ Set-Content -Path \$SummaryStateFile' -and $template -match "Send-AlertOrQueue -Subject \`$summary\.Subject -Body \`$summary\.Body -Kind 'Summary'")
Check "SL13 slow period carried over at startup unless an outage is" ($template -match 'if \(\$previous -and \$previous\.IsSlow -and \$null -ne \$previous\.SlowStartUtc -and -not \$state\.IsDown\)')
Check "SL13 alert settings normalised before the first function, so the summary text matches the monitor" ($src -match '(?m)^\$DownThreshold\s+=\s+\[math\]::Max\(1, \[int\]\$DownThreshold\)$' -and $src -match '(?m)^\$SlowWindowMinutes\s+=\s+\[math\]::Max\(1, \[int\]\$SlowWindowMinutes\)$' -and $src -match '(?m)^\$SlowAlertPercent\s+=\s+\[math\]::Min\(100, \[math\]::Max\(1, \[int\]\$SlowAlertPercent\)\)$' -and $src -match '(?m)^\$SlowClearPercent\s+=\s+\[math\]::Max\(0, \[math\]::Min\(\[int\]\$SlowClearPercent, \$SlowAlertPercent - 1\)\)$' -and $src -match '(?m)^\$DailySummaryHour\s+=\s+\[math\]::Min\(23, \[math\]::Max\(-1, \[int\]\$DailySummaryHour\)\)$' -and $src.IndexOf('$DailySummaryHour      = [math]::Min(23') -lt $src.IndexOf('function Write-Log'))
Check "SL13 heartbeat replace retried three times before warning" ($template -match '(?s)for \(\$attempt = 1; -not \$moved; \$attempt\+\+\) \{\s+try \{ Move-Item -Path \$tmp -Destination \$HeartbeatFile -Force -ErrorAction Stop; \$moved = \$true \}\s+catch \{ if \(\$attempt -ge 3\) \{ throw \}; Start-Sleep -Milliseconds 200 \}')
Check "SL13 installer summary and install email describe the alert rules" ($src -match 'Write-Host "  Slow alert      : ' -and $src -match "'Slow alert' = " -and $src -match "'Daily summary' = ")

Section "F. Live probe via curl.exe shim"
$SiteName='T'; $ExpectedContentMarker=''; $MinPopulatedBytes=100; $TimeoutSeconds=10
$Url='https://raw.githubusercontent.com/PowerShell/PowerShell/master/README.md'
$p=Get-ProbeResult
Check "Probe healthy 200"          ($p.HttpCode -eq '200')
Check "Probe healthy exit 0 / OK"  ($p.CurlExit -eq 0 -and $p.Reason -eq 'OK')
Check "Probe timings numeric"      ($null -ne $p.TotalMs -and $null -ne $p.TtfbMs -and $null -ne $p.DnsMs)
Check "Probe IP + size"            (-not [string]::IsNullOrWhiteSpace($p.RemoteIp) -and [int]$p.SizeBytes -gt 100)
Check "Probe populated"            ($p.ContentOk)
Check "Probe classified UP"        (-not (($p.CurlExit -ne 0) -or ($p.HttpCode -ne '200') -or (-not $p.ContentOk) -or ($null -eq $p.TotalMs)))
$Url='http://127.0.0.1:9/'; $TimeoutSeconds=3
$p=Get-ProbeResult
Check "Refused: code 000"          ($p.HttpCode -eq '000')
Check "Refused: exit 7 reason"     ($p.CurlExit -eq 7 -and $p.Reason -eq 'Connection refused or unreachable')
Check "Refused: classified DOWN"   (($p.CurlExit -ne 0) -or ($p.HttpCode -ne '200') -or (-not $p.ContentOk) -or ($null -eq $p.TotalMs))
$Url='http://nonexistent-host-zz.invalid/'
$p=Get-ProbeResult
Check "DNS fail: exit 6"           ($p.CurlExit -eq 6 -and $p.Reason -eq 'DNS resolution failed')
Check "DNS fail: code 000"         ($p.HttpCode -eq '000')
$Url='http://192.0.2.1/'; $TimeoutSeconds=2
$p=Get-ProbeResult
Check "Blackhole: non-zero curl exit" ($p.CurlExit -ne 0)
Check "Blackhole: classified DOWN" (($p.CurlExit -ne 0) -or ($p.HttpCode -ne '200') -or (-not $p.ContentOk) -or ($null -eq $p.TotalMs))
Check "No probe body file left behind" (-not (Test-Path (Join-Path $InstallDir 'probe-body.tmp')))
$Url='https://raw.githubusercontent.com/PowerShell/PowerShell/master/README.md'; $TimeoutSeconds=10
$ExpectedContentMarker='PowerShell'; $p=Get-ProbeResult
Check "Marker present -> populated" ($p.ContentOk)
$ExpectedContentMarker='zzz-not-in-page-zzz'; $p=Get-ProbeResult
Check "Marker absent -> not populated -> DOWN with a reason that says so" (-not $p.ContentOk -and $p.Reason -match '^Page not populated \(\d+ bytes, marker missing\)$')
$ExpectedContentMarker='PowerShell'; $MinPopulatedBytes=10000000; $p=Get-ProbeResult
Check "Marker present but body too small -> not populated, reason says marker found" (-not $p.ContentOk -and $p.Reason -match '^Page not populated \(\d+ bytes, marker found\)$')

# The real endpoint with the installer defaults: two redirects land on the federation sign-in page
$Url='https://prodasp09.hstpathways.com/p95_CSP/HSTeChart'; $TimeoutSeconds=15; $MaxRedirects=5; $MinPopulatedBytes=1000; $ExpectedContentMarker='HST Federation Provider'
$p=Get-ProbeResult
Check "Live HST: 200 after following redirects" ($p.HttpCode -eq '200' -and $p.CurlExit -eq 0 -and $p.Reason -eq 'OK')
Check "Live HST: exactly 2 redirects to the federation sign-in page" ($p.Redirects -eq '2' -and $p.FinalUrl -match '^https://prodasp09\.hstpathways\.com/p95_CSP/HSTFederationProvider/')
Check "Live HST: redirect time captured and below total" ($null -ne $p.RedirectMs -and $p.RedirectMs -gt 0 -and $p.RedirectMs -le $p.TotalMs)
Check "Live HST: sign-in page populated (marker and size)" ($p.ContentOk -and [int]$p.SizeBytes -ge 1000)
Check "Live HST: classified UP" (-not (($p.CurlExit -ne 0) -or ($p.HttpCode -ne '200') -or (-not $p.ContentOk) -or ($null -eq $p.TotalMs)))
$MaxRedirects=0; $p=Get-ProbeResult
Check "Redirect cap hit: exit 47, last code 302, classified DOWN" ($p.CurlExit -eq 47 -and $p.Reason -eq 'Too many redirects' -and $p.HttpCode -eq '302' -and (($p.HttpCode -ne '200') -or (-not $p.ContentOk)))
$MaxRedirects=0; $Url='https://prodasp09.hstpathways.com/p95_CSP/HSTeChart/'; $TimeoutSeconds=15
$p=Get-ProbeResult
$MaxRedirects=5
Check "Clean non-200 without a curl error names the HTTP code as the reason" ($p.CurlExit -eq 47 -or ($p.CurlExit -eq 0 -and $p.Reason -eq "HTTP $($p.HttpCode)"))
$MaxRedirects=5; $ExpectedContentMarker=''; $MinPopulatedBytes=100

# Install-time preflight uses the same redirect handling
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-EndpointReachable'},$false)) { Invoke-Expression $f.Extent.Text }
$Url='https://prodasp09.hstpathways.com/p95_CSP/HSTeChart'; $TimeoutSeconds=15; $MaxRedirects=5
$r = Test-EndpointReachable
Check "Preflight: live HST reachable through 2 redirects" ($r.Reachable -and $r.HttpCode -eq '200' -and $r.Redirects -eq 2 -and $r.FinalUrl -match '^https://prodasp09\.hstpathways\.com/p95_CSP/HSTFederationProvider/')
$Url='http://127.0.0.1:9/'; $TimeoutSeconds=3
$r = Test-EndpointReachable
Check "Preflight: refused port -> not reachable, code 000" (-not $r.Reachable -and $r.HttpCode -eq '000')
$Url='https://raw.githubusercontent.com/PowerShell/PowerShell/master/README.md'; $TimeoutSeconds=10

Section "G. Wizard flows with scripted keystrokes"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('Get-MailConfiguration','Get-SiteName') }) { Invoke-Expression $f.Extent.Text }
$script:Q = [System.Collections.Queue]::new(); $script:PromptCount = 0; $script:PromptLog = @()
function Push($answers){ $script:Q.Clear(); foreach($a in $answers){ $script:Q.Enqueue($a) }; $script:PromptCount=0; $script:PromptLog=@() }
function NextAnswer($prompt){ $script:PromptCount++; $script:PromptLog += $prompt; if ($script:Q.Count -eq 0) { throw "Wizard asked an unexpected extra prompt: $prompt" }; return $script:Q.Dequeue() }
function Read-Setting { param([string]$Prompt,[string]$Default="") if ($NonInteractive){return $Default}; $a = NextAnswer $Prompt; if ([string]::IsNullOrWhiteSpace($a)) { return $Default }; return $a.Trim() }
function Read-Choice { param([string]$Prompt,[string[]]$Allowed,[string]$Default) if ($NonInteractive){return $Default}; while ($true) { $a = NextAnswer $Prompt; if ([string]::IsNullOrWhiteSpace($a)) { return $Default }; $a=$a.Trim().ToUpper(); if ($Allowed -contains $a) { return $a } } }
function Read-Host { param([string]$Prompt,[switch]$AsSecureString) $a = NextAnswer $Prompt; if ($AsSecureString) { if ($a) { return (ConvertTo-SecureString $a -AsPlainText -Force) } else { return (New-Object System.Security.SecureString) } }; return $a }
function Get-DirectSendHost { param($FromAddress) [PSCustomObject]@{ Host='contoso-com.mail.protection.outlook.com'; Source='stub' } }
function Protect-Secret { param($Password,$PlainText) $script:LastPlainText = $PlainText; return 'CIPHER' }
$script:StoredSecretForWizard = $null
function Get-StoredSecret { return $script:StoredSecretForWizard }
$script:SavedMid = $null
function Save-InstallSettings { param($Settings) $script:SavedMid = $Settings }
function Get-Credential { param($UserName,$Message) New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }
$script:SendResult = $true; $script:SentSubjects=@()
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) $script:SentSubjects += $Subject; $script:LastGraphSecret=$GraphSecret; return $script:SendResult }
$script:CreateResult = @{ TenantId='11111111-2222-3333-4444-555555555555'; ClientId='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; Secret='newsecret'; Expires='2028-09-04'; Sender='hst@contoso.com' }
function New-TenantMailApp { param($SenderAddress,$SiteName) return $script:CreateResult }
$NonInteractive=$false; $GraphAppDisplayName='HST Monitor'; $MailMethod='Graph'; $SmtpServer=''; $SmtpPort=25; $SmtpUseSsl=$false; $MailFrom=''; $MailTo=@(); $SmtpAuthUser=''; $GraphTenantId=''; $GraphClientId=''; $GraphSecretExpires=''
$T='11111111-2222-3333-4444-555555555555'; $C='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'

# W1 Graph, app exists, fresh box
Push @('1','hst@contoso.com','a@contoso.com','Y',$T,$C,'s3cret','2028-01-01','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W1 Graph existing app returns config" ($m.MailMethod -eq 'Graph' -and $m.GraphTenantId -eq $T -and $m.GraphClientId -eq $C -and $m.GraphSecretExpires -eq '2028-01-01' -and $m.CipherText -eq 'CIPHER' -and $m.SmtpServer -eq 'graph.microsoft.com')
Check "W1 secret passed to test send" ($script:LastGraphSecret -eq 's3cret')
Check "W1 pasted secret reaches Protect-Secret as text" ($script:LastPlainText -eq 's3cret')
Check "W1 exactly 9 prompts" ($script:PromptCount -eq 9)
Check "W1 test subject names site and server" ($script:SentSubjects[-1] -eq "[MONITOR TEST] S ($env:COMPUTERNAME)")

# W2 Graph re-run with saved settings: all Enter except secret
$saved = [PSCustomObject]@{ MailMethod='Graph'; SmtpServer='graph.microsoft.com'; SmtpPort=443; SmtpUseSsl=$true; MailFrom='hst@contoso.com'; MailTo=@('a@contoso.com','b@contoso.com'); SmtpAuthUser=''; GraphTenantId=$T; GraphClientId=$C; GraphSecretExpires='2028-01-01'; CredentialFor=$C }
Push @('','','','','','','s3cret','','')
$m = Get-MailConfiguration -Saved $saved -SiteName 'S'
Check "W2 saved settings prefill everything" ($m.MailFrom -eq 'hst@contoso.com' -and $m.MailTo.Count -eq 2 -and $m.GraphTenantId -eq $T -and $m.GraphClientId -eq $C -and $m.GraphSecretExpires -eq '2028-01-01')
Check "W2 hasApp defaulted to Y when saved IDs exist" ($script:PromptLog[3] -match 'already been created' -and $m.GraphTenantId -eq $T)
Check "W2 9 prompts, all Enter except secret" ($script:PromptCount -eq 9)
Check "W2 no stored secret -> plain paste prompt" ($script:PromptLog[6] -eq 'Client secret (paste, input hidden)')

# W20 re-run with a stored secret for the same app: Enter keeps it, nothing needs pasting
$script:StoredSecretForWizard = 'stored~secret'
Push @('','','','','','','','','')
$m = Get-MailConfiguration -Saved $saved -SiteName 'S'
Check "W20 Enter keeps the stored secret" ($m.MailMethod -eq 'Graph' -and $script:LastGraphSecret -eq 'stored~secret' -and $script:LastPlainText -eq 'stored~secret' -and $script:PromptCount -eq 9 -and $script:PromptLog[6] -match '^Client secret \(Enter = keep')
# W21 a pasted secret wins over the stored one
Push @('','','','','','','fresh~secret','','')
$m = Get-MailConfiguration -Saved $saved -SiteName 'S'
Check "W21 pasted secret replaces the stored one" ($script:LastGraphSecret -eq 'fresh~secret' -and $script:LastPlainText -eq 'fresh~secret')
# W22 stored secret belongs to a different app: not offered, blank is refused
$savedOther = [PSCustomObject]@{ MailMethod='Graph'; SmtpServer='graph.microsoft.com'; SmtpPort=443; SmtpUseSsl=$true; MailFrom='hst@contoso.com'; MailTo=@('a@contoso.com'); SmtpAuthUser=''; GraphTenantId=$T; GraphClientId='ffffffff-ffff-ffff-ffff-ffffffffffff'; GraphSecretExpires=''; CredentialFor='ffffffff-ffff-ffff-ffff-ffffffffffff' }
Push @('','','','','',$C,'','s3cret','','Y')
$m = Get-MailConfiguration -Saved $savedOther -SiteName 'S'
Check "W22 stored secret for another app is not offered" ($m.GraphClientId -eq $C -and $script:LastGraphSecret -eq 's3cret' -and $script:PromptCount -eq 10 -and $script:PromptLog[6] -eq 'Client secret (paste, input hidden)')
# W23 saved settings were not Graph: stored file (from an SMTP password) is never offered as a Graph secret
$savedSmtp = [PSCustomObject]@{ MailMethod='Authenticated'; SmtpServer='smtp.office365.com'; SmtpPort=587; SmtpUseSsl=$true; MailFrom='hst@contoso.com'; MailTo=@('a@contoso.com'); SmtpAuthUser='hst@contoso.com'; GraphTenantId=''; GraphClientId=''; GraphSecretExpires=''; CredentialFor='hst@contoso.com' }
Push @('1','','','Y',$T,$C,'','s3cret','','Y')
$m = Get-MailConfiguration -Saved $savedSmtp -SiteName 'S'
Check "W23 an SMTP password on disk is never offered as a Graph secret" ($script:LastGraphSecret -eq 's3cret' -and $script:PromptCount -eq 10 -and $script:PromptLog[6] -eq 'Client secret (paste, input hidden)')
# W24 settings saved by an abandoned run (no CredentialFor yet): the file on disk is not trusted for this app
$savedAbandoned = [PSCustomObject]@{ MailMethod='Graph'; SmtpServer='graph.microsoft.com'; SmtpPort=443; SmtpUseSsl=$true; MailFrom='hst@contoso.com'; MailTo=@('a@contoso.com'); SmtpAuthUser=''; GraphTenantId=$T; GraphClientId=$C; GraphSecretExpires='' }
Push @('','','','','','','','s3cret','','Y')
$m = Get-MailConfiguration -Saved $savedAbandoned -SiteName 'S'
Check "W24 no CredentialFor in saved settings -> stored secret not offered" ($script:LastGraphSecret -eq 's3cret' -and $script:PromptCount -eq 10 -and $script:PromptLog[6] -eq 'Client secret (paste, input hidden)')
$script:StoredSecretForWizard = $null

# W3 Graph, create path (N): no Tenant/Client/expiry/secret prompts afterwards
Push @('1','hst@contoso.com','a@contoso.com','N','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W3 create path uses created values" ($m.GraphTenantId -eq $T -and $m.GraphClientId -eq $C -and $m.GraphSecretExpires -eq '2028-09-04')
Check "W3 created secret used for test send" ($script:LastGraphSecret -eq 'newsecret')
Check "W3 exactly 5 prompts (method, from, to, create?, arrived?)" ($script:PromptCount -eq 5)
Check "W3 created IDs persisted immediately (abandoned runs prefill Y)" ($script:SavedMid -and $script:SavedMid.GraphClientId -eq $C -and $script:SavedMid.GraphTenantId -eq $T -and $script:SavedMid.MailMethod -eq 'Graph')

# W4 create fails then abort
$script:CreateResult = $null
Push @('1','hst@contoso.com','a@contoso.com','N','X')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W4 abort after failed tenant setup returns null" ($null -eq $m)
# W4b create fails, go back, choose Direct Send instead
Push @('1','hst@contoso.com','a@contoso.com','N','R','2','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W4b back to menu keeps from/to as defaults and lands on DirectSend" ($m.MailMethod -eq 'DirectSend' -and $m.MailFrom -eq 'hst@contoso.com' -and $m.SmtpServer -eq 'contoso-com.mail.protection.outlook.com' -and $m.SmtpPort -eq 25)
$script:CreateResult = @{ TenantId=$T; ClientId=$C; Secret='newsecret'; Expires='2028-09-04'; Sender='hst@contoso.com' }

# W5 Relay with bad inputs corrected, then N to change, then Y
Push @('3','notanemail','ok@contoso.com','','a@contoso.com, bad, b@contoso.com','','relay01','2525','y','N','','','','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W5 relay validation loops and defaults carry across N" ($m.MailMethod -eq 'Relay' -and $m.MailFrom -eq 'ok@contoso.com' -and $m.MailTo.Count -eq 2 -and $m.SmtpServer -eq 'relay01' -and $m.SmtpPort -eq 2525 -and $m.SmtpUseSsl)
Check "W5 second pass needed only Enters" ($script:PromptCount -eq 17)

# W6 Authenticated, skip on arrival prompt
Push @('4','ok@contoso.com','a@contoso.com','','','','','S')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W6 authenticated defaults: o365 587 TLS, user=from, cipher stored" ($m.MailMethod -eq 'Authenticated' -and $m.SmtpServer -eq 'smtp.office365.com' -and $m.SmtpPort -eq 587 -and $m.SmtpUseSsl -and $m.SmtpAuthUser -eq 'ok@contoso.com' -and $m.CipherText -eq 'CIPHER')

# W7 send failure: R retries with defaults, then success
$script:SendResult = $false
Push @('2','ok@contoso.com','a@contoso.com','','R')
try { $script:SendResult = $false; $m = $null
  # second pass will succeed: flip result when the retry prompt is consumed
  $orig = ${function:Send-MailWithConfig}
  function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) $script:SentSubjects += $Subject; $script:SendCalls++; return ($script:SendCalls -ge 2) }
  $script:SendCalls=0
  Push @('2','ok@contoso.com','a@contoso.com','','R','','','','','Y')
  $m = Get-MailConfiguration -Saved $null -SiteName 'S'
  Check "W7 failed send -> retry -> success returns config" ($m.MailMethod -eq 'DirectSend' -and $script:SendCalls -eq 2)
} finally { ${function:Send-MailWithConfig} = $orig; $script:SendResult = $true }

# W8 skip after failed send
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) return $false }
Push @('2','ok@contoso.com','a@contoso.com','','S')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W8 skip after failed send still returns config" ($m.MailMethod -eq 'DirectSend')

# W10 first-site Graph: test send denied while the grant propagates, W waits and retries until it succeeds
$script:LastGraphSendError = 'AccessDenied'   # a stubbed send that fails stands for the Exchange grant not applied yet
function Start-Sleep { param($Seconds,$Milliseconds) $script:SleptSeconds += [int]$Seconds }
$script:SleptSeconds=0; $script:SendCalls=0; $script:SentSubjects=@()
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) $script:SentSubjects += $Subject; $script:SendCalls++; return ($script:SendCalls -ge 3) }
Push @('1','hst@contoso.com','a@contoso.com','N','W','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W10 W waits, retries, succeeds on 3rd send, then asks arrived" ($m -and $m.MailMethod -eq 'Graph' -and $script:SendCalls -eq 3 -and $script:PromptCount -eq 6 -and $script:PromptLog[4] -match '^W = wait' -and $script:PromptLog[5] -match 'Did the test email arrive')
Check "W10 quiet 30 min first, then 10 min between retries" ($script:SleptSeconds -eq 2400)
Check "W10 returned the created app config" ($m.GraphClientId -eq $C -and $m.CipherText -eq 'CIPHER' -and $m.GraphSecretExpires -eq '2028-09-04')

# W11 W never propagates: gives up at the 40 minute deadline and finishes as S, no arrived prompt
$script:Clock = [datetime]'2026-09-04T12:00:00'
function Get-Date { param($Format) $script:Clock = $script:Clock.AddSeconds(30); if ($Format) { return $script:Clock.ToString($Format) }; return $script:Clock }
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) $script:SendCalls++; return $false }
$script:SendCalls=0
Push @('1','hst@contoso.com','a@contoso.com','N','W')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Remove-Item function:Get-Date
Check "W11 gave up after the window and still returned config" ($m -and $m.MailMethod -eq 'Graph' -and $script:PromptCount -eq 5)
Check "W11 retried repeatedly before giving up" ($script:SendCalls -ge 10 -and $script:SendCalls -le 100)

# W12 first-site Graph denial, S finishes immediately without waiting
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) return $false }
$script:SleptSeconds=0
Push @('1','hst@contoso.com','a@contoso.com','N','S')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W12 S after propagation denial returns config without waiting" ($m -and $m.MailMethod -eq 'Graph' -and $script:PromptCount -eq 5 -and $script:SleptSeconds -eq 0)

# W25 first-site Graph, the secret itself is rejected: no propagation wait is offered, the plain retry prompt appears
$script:LastGraphSendError = 'InvalidSecret'
Push @('1','hst@contoso.com','a@contoso.com','N','S')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W25 rejected secret gets the plain retry prompt, not the propagation wait" ($m -and $m.MailMethod -eq 'Graph' -and $script:PromptCount -eq 5 -and $script:PromptLog[4] -match '^Test email failed\. R = change settings')
Check "W25 finishing after a failed test marks the install email to be skipped" ($script:MailTestSkipped -eq $true)
$script:MailTestSkipped = $false

# W26 a secret minted this run is rejected: R, then the app prompt defaults to N and the tenant setup runs again with a new secret
function New-TenantMailApp { param($SenderAddress,$SiteName) $script:CreateCalls++; return $script:CreateResult }
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret,$MaxWaitSeconds) $script:SendCalls++; $script:LastGraphSecret=$GraphSecret; return ($script:SendCalls -ge 2) }
$script:SendCalls = 0; $script:CreateCalls = 0; $script:LastGraphSendError = 'InvalidSecret'
Push @('1','hst@contoso.com','a@contoso.com','N','R','','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W26 after a rejected minted secret the app prompt defaults to N and mints again" ($m -and $m.GraphClientId -eq $C -and $script:PromptCount -eq 10 -and $script:CreateCalls -eq 2 -and $script:PromptLog[8] -match 'already been created' -and $script:LastGraphSecret -eq 'newsecret' -and $script:MailTestSkipped -eq $false)
# W27 same start, but Y at the app prompt now asks for a pasted secret instead of silently reusing the rejected one
$script:SendCalls = 0; $script:CreateCalls = 0; $script:LastGraphSendError = 'InvalidSecret'
Push @('1','hst@contoso.com','a@contoso.com','N','R','','','','Y','','','pasted~secret','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W27 Y after a rejected minted secret prompts for a paste, rejected secret never reused" ($m -and $script:PromptCount -eq 14 -and $script:CreateCalls -eq 1 -and (@($script:PromptLog | Where-Object { $_ -match '^Client secret \(paste' }).Count -eq 1) -and $script:LastGraphSecret -eq 'pasted~secret')
# W28 a GUID pasted as the secret is refused and asked again
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret,$MaxWaitSeconds) $script:LastGraphSecret=$GraphSecret; return $true }
$script:LastGraphSendError = ''
Push @('1','hst@contoso.com','a@contoso.com','Y',$T,$C,'11111111-2222-3333-4444-555555555555','s3cret','2028-01-01','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W28 a GUID is refused as the secret and the real value is taken on the second prompt" ($m -and $script:PromptCount -eq 10 -and (@($script:PromptLog | Where-Object { $_ -match '^Client secret' }).Count -eq 2) -and $script:LastGraphSecret -eq 's3cret')
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret,$MaxWaitSeconds) return $false }
$script:LastGraphSendError = 'AccessDenied'
Remove-Item function:Start-Sleep

# W13 port validation loops until a valid port is typed
$script:SendPlan = [System.Collections.Queue]::new()
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) $script:SentSubjects += $Subject; $script:LastGraphSecret=$GraphSecret; $script:SendCalls++; if ($script:SendPlan.Count -gt 0) { return $script:SendPlan.Dequeue() }; return $true }
Push @('3','ok@contoso.com','a@contoso.com','relay01','abc','70000','2525','N','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W13 bad ports re-prompt, valid port accepted" ($m.MailMethod -eq 'Relay' -and $m.SmtpPort -eq 2525 -and $script:PromptCount -eq 9)

# W14 saved Authenticated with TLS off: all Enter keeps TLS off and port 25
$savedAuth = [PSCustomObject]@{ MailMethod='Authenticated'; SmtpServer='mail.internal.local'; SmtpPort=25; SmtpUseSsl=$false; MailFrom='svc@contoso.com'; MailTo=@('a@contoso.com'); SmtpAuthUser='svc@contoso.com' }
Push @('','','','','','','','Y')
$m = Get-MailConfiguration -Saved $savedAuth -SiteName 'S'
Check "W14 saved TLS off survives an all-Enter re-run" ($m.MailMethod -eq 'Authenticated' -and $m.SmtpServer -eq 'mail.internal.local' -and $m.SmtpPort -eq 25 -and -not $m.SmtpUseSsl -and $m.SmtpAuthUser -eq 'svc@contoso.com' -and $script:PromptCount -eq 8)

# W15 saved Graph IDs survive a failed Direct Send detour: the app prompt still defaults to Y and nothing is re-created
$script:CreateCalls = 0
function New-TenantMailApp { param($SenderAddress,$SiteName) $script:CreateCalls++; return $script:CreateResult }
$savedGraph = [PSCustomObject]@{ MailMethod='Graph'; SmtpServer='graph.microsoft.com'; SmtpPort=443; SmtpUseSsl=$true; MailFrom='hst@contoso.com'; MailTo=@('a@contoso.com'); SmtpAuthUser=''; GraphTenantId=$T; GraphClientId=$C; GraphSecretExpires='2027-01-01' }
$script:SendPlan.Clear(); $script:SendPlan.Enqueue($false)
Push @('2','','','','R','1','','','','','','s3cret','','Y')
$m = Get-MailConfiguration -Saved $savedGraph -SiteName 'S'
Check "W15 Graph IDs kept after a Direct Send detour, app not re-created" ($m.MailMethod -eq 'Graph' -and $m.GraphTenantId -eq $T -and $m.GraphClientId -eq $C -and $m.GraphSecretExpires -eq '2027-01-01' -and $script:CreateCalls -eq 0 -and $script:PromptCount -eq 14 -and $script:PromptLog[8] -match 'already been created')

# W16 a saved expiry can be cleared with '-'
Push @('1','','','','','','s3cret','-','Y')
$m = Get-MailConfiguration -Saved $savedGraph -SiteName 'S'
Check "W16 dash clears the saved expiry" ($m.MailMethod -eq 'Graph' -and $m.GraphSecretExpires -eq '')

# W17 cancelled credential dialog keeps the host, port, TLS, and user just typed
$script:CredCalls = 0
function Get-Credential { param($UserName,$Message) $script:CredCalls++; if ($script:CredCalls -eq 1) { return $null }; New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }
Push @('4','ok@contoso.com','a@contoso.com','smtp.custom.local','2587','N','me@contoso.com','','','','','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W17 typed values survive a cancelled credential dialog" ($m.MailMethod -eq 'Authenticated' -and $m.SmtpServer -eq 'smtp.custom.local' -and $m.SmtpPort -eq 2587 -and -not $m.SmtpUseSsl -and $m.SmtpAuthUser -eq 'me@contoso.com' -and $script:PromptCount -eq 15)
function Get-Credential { param($UserName,$Message) New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }

# W19 a blank password in the credential dialog is refused like a cancel, then a real one is accepted
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message"; $script:Logs += "$Level|$Message" }
$script:Logs = @(); $script:CredCalls = 0
function Get-Credential { param($UserName,$Message) $script:CredCalls++; if ($script:CredCalls -eq 1) { return (New-Object System.Management.Automation.PSCredential($UserName, (New-Object System.Security.SecureString))) }; New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }
Push @('4','ok@contoso.com','a@contoso.com','','','','','','','','','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W19 blank password refused, wizard restarts with defaults kept, real password accepted" ($m.MailMethod -eq 'Authenticated' -and $m.CipherText -eq 'CIPHER' -and $script:CredCalls -eq 2 -and $script:PromptCount -eq 15 -and (($script:Logs -join "`n") -match 'WARNING\|A password is required'))
function Write-Log { param($Level,$Message) }
function Get-Credential { param($UserName,$Message) New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }

# W18 after creating the app and choosing R at the propagation prompt, the next Graph pass reuses the new secret and needs no paste
$script:SendPlan.Clear(); $script:SendPlan.Enqueue($false); $script:SendPlan.Enqueue($true)
function Start-Sleep { param($Seconds,$Milliseconds) }
Push @('1','hst@contoso.com','a@contoso.com','N','R','','','','','','','','Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Remove-Item function:Start-Sleep
Check "W18 second Graph pass reuses the secret created in this run" ($m.MailMethod -eq 'Graph' -and $m.GraphClientId -eq $C -and $script:LastGraphSecret -eq 'newsecret' -and $script:PromptCount -eq 13 -and (@($script:PromptLog | Where-Object { $_ -match 'Client secret' }).Count -eq 0) -and $script:PromptLog[4] -match '^W = wait')

# W9 Non-interactive: Graph refused, DirectSend proceeds with config block
$NonInteractive=$true; $MailMethod='Graph'; $MailFrom='x@contoso.com'; $MailTo=@('a@contoso.com')
Check "W9 non-interactive Graph refused" ($null -eq (Get-MailConfiguration -Saved $null -SiteName 'S'))
$MailMethod='DirectSend'
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W9 non-interactive DirectSend proceeds" ($m.MailMethod -eq 'DirectSend' -and $m.MailFrom -eq 'x@contoso.com')
$NonInteractive=$false

# Site name prompts
$env:COMPUTERNAME='SRV01'; $SiteNameOverride=''
Push @('CapCity'); $n = Get-SiteName; Check "Site: clean input needs 1 prompt" ($n -eq 'CapCity' -and $script:PromptCount -eq 1)
Push @(''); $n = Get-SiteName; Check "Site: Enter takes machine name" ($n -eq 'SRV01' -and $script:PromptCount -eq 1)
Push @("O'Brien!!",'Y'); $n = Get-SiteName; Check "Site: dirty input asks once to confirm cleanup" ($n -eq 'OBrien' -and $script:PromptCount -eq 2)
Push @('!!!','Jersey Shore'); $n = Get-SiteName; Check "Site: empty-after-cleanup re-prompts" ($n -eq 'Jersey Shore' -and $script:PromptCount -eq 2)
Push @(''); $n = Get-SiteName -SavedDefault 'Saved Site'; Check "Site: saved default wins over machine name" ($n -eq 'Saved Site')
$SiteNameOverride='Override Site'; Push @(); $n = Get-SiteName; Check "Site: override skips prompt" ($n -eq 'Override Site' -and $script:PromptCount -eq 0); $SiteNameOverride=''

Section "H. Sign-in and setup flow static checks"
Check "One sign-in: offline_access requested" ($src -match 'offline_access')
Check "Device code copied to clipboard" ($src -match 'Set-Clipboard -Value \$dc\.user_code')
Check "Device fallback still has pre-filled URL" ($src -match [regex]::Escape('$signInUrl = "$($dc.verification_uri)?otc=$($dc.user_code)"'))
Check "One sign-in: Azure CLI client with .default scope only (no AADSTS65002)" ($src -match "clientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'" -and $src -match "scope = 'https://graph\.microsoft\.com/\.default offline_access'" -and -not ($src -match 'graph\.microsoft\.com/Application\.ReadWrite'))
Check "Exchange token from the same sign-in's refresh token, same client" ($src -match "grant_type = 'refresh_token'; client_id = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'" -and $src -match 'outlook\.office365\.com/\.default offline_access')
Check "No second interactive sign-in for Exchange" ((([regex]::Matches($src,'Get-BrowserToken -ClientId')).Count) -eq 1)
Check "No embedded IE control, no other first-party client ids" (-not ($src -match 'WebBrowser|FEATURE_BROWSER_EMULATION|DoEvents|14d82eec|fb78d390'))
Check "Prompt text asks for Global Administrator" ($src -match "Sign in as a Global Administrator'" -and -not ($src -match 'Application Administrator'))
Check "Exchange work goes through admin REST endpoint" ($src -match 'outlook\.office365\.com/adminapi/beta/\$TenantId/InvokeCommand' -and $src -match 'CmdletInput')
Check "Module is fallback only, connects with -UserPrincipalName (all versions)" ($src -match 'Connect-ExchangeOnline -UserPrincipalName \$auth\.Upn' -and -not ($src -match 'Connect-ExchangeOnline -Device') -and -not ($src -match 'Connect-ExchangeOnline -AccessToken'))
Check "Module removed only if it was used" ($src.Contains("if (`$script:ExoMode -eq 'module') { Remove-TenantSetupModules }"))
Check "No background install job remains" (-not ($src -match 'Start-ModuleInstallJob|Start-Job'))
Check "Every Exchange cmdlet routed through the dispatcher" ((([regex]::Matches($src,"Invoke-ExoCmdlet -Name '([A-Za-z-]+)'") | % { $_.Groups[1].Value } | Sort-Object -Unique) -join ',') -eq 'Add-DistributionGroupMember,Enable-OrganizationCustomization,Get-ApplicationAccessPolicy,Get-DistributionGroup,Get-DistributionGroupMember,Get-Mailbox,Get-ManagementRoleAssignment,Get-ManagementScope,Get-OrganizationConfig,Get-ServicePrincipal,New-ApplicationAccessPolicy,New-DistributionGroup,New-Mailbox,New-ManagementRoleAssignment,New-ManagementScope,New-ServicePrincipal,Remove-ManagementRoleAssignment,Set-ManagementScope,Test-ApplicationAccessPolicy,Test-ServicePrincipalAuthorization')
Check "AppId sent as a string array (String[] parameter)" ($src -match '\bAppId = \[string\[\]\]@\(\$app\.appId\)')
Check "RBAC for Applications is the primary scoping" ($src -match "Role = 'Application Mail\.Send'" -and $src -match 'New-ManagementScope' -and $src -match 'Test-ServicePrincipalAuthorization')
Check "Legacy policy retained as fallback" ($src -match 'Falling back to the legacy application access policy')
Check "Tenant-wide consent removed when RBAC scope applies" ($src -match 'Method DELETE -Path "servicePrincipals/')
Check "Dehydrated tenant handled" ($src -match 'IsDehydrated' -and $src -match 'Enable-OrganizationCustomization')
Check "Secret minted only after Exchange configuration" ($src.IndexOf('/addPassword') -gt $src.IndexOf('Exchange configuration failed'))
Check "Raw REST error body preserved" ($src -match 'Full response')
Check "Web error bodies surfaced in installer and monitor" ((([regex]::Matches($src,'function Get-RestErrorDetail')).Count -eq 2) -and (([regex]::Matches($src,[regex]::Escape('Get-RestErrorDetail $_'))).Count -ge 4))
Check "Access-denied hint explains the propagation wait" ($src -match 'Wait at least 30 minutes without retrying')
Check "First-site Graph test send auto-waits only for an Exchange access denial" ($src -match "\`$method -eq 'Graph' -and \`$script:LastGraphSendError -eq 'AccessDenied' -and \(\`$created -or" -and $src -match 'wait and retry automatically')
Check "Ports validated through Read-PortSetting" ($src -match 'function Read-PortSetting' -and -not ($src -match '\[int\]\(Read-Setting -Prompt "Port"'))
Check "Install root and monitor folder hardened before anything is written" ($src -match 'function Protect-InstallFolder' -and $src -match "Protect-InstallFolder -Path \`$InstallDir\s*\r?\n" -and $src -match "Protect-InstallFolder -Path \`$InstallRoot\s*\r?\n" -and $src -match "Protect-InstallFolder -Path \(Split-Path -Path \`$InstallRoot -Parent\) -OwnerOnly")
Check "Credential file restricted before content is written, icacls checked" ($src -match "(?s)function Save-SmtpCredential \{.*?icacls\.exe.*?LASTEXITCODE.*?Set-Content.*?\n\}")
Check "Exchange not-found mapping is narrow and transient errors retried" ($src -match "couldn\.t be found\|could not be found" -and $src -match '429, 500, 502, 503, 504')
Check "Stale role assignment replaced instead of trusted by name" ($src -match "Remove-ManagementRoleAssignment")
Check "Created secret bypasses the transcript" (-not ($src -match 'Write-Host "  Client secret') -and $src -match '\[Console\]::Out\.WriteLine\("  Client secret')
Check "Device code honours slow_down" ($src -match "if \(\`$err -eq 'slow_down'\) \{ \`$interval \+= 5; continue \}")
Check "Browser sign-in checks state before rendering the page" ($src -match "if \(\`$code -and \`$gotState -ne \`$state\)")
Check "Policy retry can switch to the group directory id" ($src -match 'ExternalDirectoryObjectId')
Check "Scope filter quotes apostrophes in the sender" ($src -match [regex]::Escape('$SenderAddress.Replace("''", "''''")'))
Check "Created values skip re-prompts" ($src -match '\$tenant = \$created\.TenantId; \$client = \$created\.ClientId; \$expiry = \$created\.Expires')
Check "Site confirm only on cleanup change" ($src -match "Cleaned to '" -and -not ($src -match "Use site name '"))
$sw = $src.Substring($src.IndexOf('        switch ($method) {'), $src.IndexOf('        if ($restart) { continue }') - $src.IndexOf('        switch ($method) {'))
Check "No 'continue' inside the wizard switch (would exit switch, not loop)" (-not ($sw -match '\bcontinue\b'))
# W10 password branch: cancelled Get-Credential restarts cleanly with defaults kept
function Get-Credential { param($UserName,$Message) if ($script:CredCancelOnce) { $script:CredCancelOnce=$false; return $null }; New-Object System.Management.Automation.PSCredential($UserName,(ConvertTo-SecureString 'pw' -AsPlainText -Force)) }
function Send-MailWithConfig { param($Mail,$Subject,$Body,$Credential,$GraphSecret) return $true }
$script:CredCancelOnce=$true
Push @('4','ok@contoso.com','a@contoso.com','','','','',   '','','','','','','',   'Y')
$m = Get-MailConfiguration -Saved $null -SiteName 'S'
Check "W10 cancelled credential restarts wizard, defaults kept, completes" ($m.MailMethod -eq 'Authenticated' -and $m.MailFrom -eq 'ok@contoso.com' -and $m.CipherText -eq 'CIPHER' -and $script:PromptCount -eq 15)
Check "Install summary printed" ($src -match 'Write-Host "Install summary"')


Section "H2. Installer token path against a local sign-in mock (fresh secret replication)"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('Get-RestErrorDetail','Get-InstallerGraphToken','Send-GraphMail') }) { Invoke-Expression $f.Extent.Text }
Check "H2 seams and state declared once each" ((([regex]::Matches($src,'(?m)^\$script:LoginBase = ''https://login\.microsoftonline\.com''')).Count -eq 1) -and (([regex]::Matches($src,'(?m)^\$script:GraphBase = ''https://graph\.microsoft\.com''')).Count -eq 1) -and $src -match '(?m)^\$script:FreshSecretRetrySeconds = 15$' -and $src -match '(?m)^\$script:PastedSecretRetrySeconds = 60$')
Check "H2 installer app-only token and send go through the seams" ($src -match '"\$script:LoginBase/\$TenantId/oauth2/v2\.0/token"' -and $src -match '"\$script:GraphBase/v1\.0/users/\$SenderAddress/sendMail"')
Check "H2 install email and test email share the helper (single Graph send path)" ((([regex]::Matches($src,'Get-InstallerGraphToken -TenantId')).Count -eq 2) -and (([regex]::Matches($src,'client_secret = \$Secret;')).Count -eq 1))
$h2WriteLog = ${function:Write-Log}
$script:H2Log = @()
function Write-Log { param($Level,$Message) $script:H2Log += "$Level|$Message" }
function Get-FreeTestPort { $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0); $l.Start(); $p = ($l.LocalEndpoint).Port; $l.Stop(); return $p }
$mockLog = Join-Path ([IO.Path]::GetTempPath()) "hst_mock_$PID.log"
function Start-MockGraph {
    # Minimal HTTP responder: the token route fails the first TokenFailures requests with FailCode, then issues tokens;
    # the sendMail route answers SendStatus with SendBody. Every request is appended to LogPath as path|status.
    param([int]$Port, [int]$TokenFailures = 0, [string]$FailCode = 'AADSTS7000215', [int]$FailStatus = 401, [int]$SendStatus = 202, [string]$SendBody = '', [string]$LogPath)
    if (Test-Path $LogPath) { Remove-Item $LogPath -Force }
    Start-Job -ScriptBlock {
        param($Port, $TokenFailures, $FailCode, $FailStatus, $SendStatus, $SendBody, $LogPath)
        $reasons = @{ 200 = 'OK'; 202 = 'Accepted'; 400 = 'Bad Request'; 401 = 'Unauthorized'; 403 = 'Forbidden'; 404 = 'Not Found' }
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
        $listener.Start()
        $deadline = (Get-Date).AddSeconds(45)
        $tokenHits = 0
        while ((Get-Date) -lt $deadline) {
            if (-not $listener.Pending()) { Start-Sleep -Milliseconds 30; continue }
            $client = $listener.AcceptTcpClient()
            $stream = $client.GetStream(); $stream.ReadTimeout = 3000
            $buf = New-Object byte[] 65536; $raw = New-Object System.IO.MemoryStream
            $headerEnd = -1; $bodyLen = 0
            try {
                while ($true) {
                    $n = $stream.Read($buf, 0, $buf.Length)
                    if ($n -le 0) { break }
                    $raw.Write($buf, 0, $n)
                    $text = [Text.Encoding]::ASCII.GetString($raw.ToArray())
                    if ($headerEnd -lt 0) {
                        $headerEnd = $text.IndexOf("`r`n`r`n")
                        if ($headerEnd -ge 0) { $m = [regex]::Match($text.Substring(0, $headerEnd), '(?im)^Content-Length:\s*(\d+)'); if ($m.Success) { $bodyLen = [int]$m.Groups[1].Value } }
                    }
                    if ($headerEnd -ge 0 -and $raw.Length -ge ($headerEnd + 4 + $bodyLen)) { break }
                }
            } catch { }
            $text = [Text.Encoding]::ASCII.GetString($raw.ToArray())
            $path = (($text -split "`r`n")[0] -split ' ')[1]
            if ($path -match '/oauth2/v2\.0/token$') {
                $tokenHits++
                if ($tokenHits -le $TokenFailures) { $status = $FailStatus; $body = '{"error":"invalid_client","error_description":"' + $FailCode + ': Invalid client secret provided. Ensure the secret being sent in the request is the client secret value, not the client secret ID","error_codes":[7000215]}' }
                else { $status = 200; $body = '{"token_type":"Bearer","expires_in":3599,"access_token":"tok' + $tokenHits + '"}' }
            }
            elseif ($path -match '/sendMail$') { $status = $SendStatus; $body = $SendBody }
            else { $status = 404; $body = '' }
            Add-Content -Path $LogPath -Value "$path|$status" -Encoding ASCII
            $resp = "HTTP/1.1 $status $($reasons[$status])`r`nContent-Type: application/json`r`nContent-Length: $([Text.Encoding]::UTF8.GetByteCount($body))`r`nConnection: close`r`n`r`n$body"
            $bytes = [Text.Encoding]::UTF8.GetBytes($resp)
            try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush() } catch { }
            $client.Close()
        }
        $listener.Stop()
    } -ArgumentList $Port, $TokenFailures, $FailCode, $FailStatus, $SendStatus, $SendBody, $LogPath
}
function Wait-MockGraph {
    param([int]$Port)
    $deadline = (Get-Date).AddSeconds(25)
    while ((Get-Date) -lt $deadline) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient; $c.Connect('127.0.0.1', $Port); $s = $c.GetStream(); $s.ReadTimeout = 3000
            $req = [Text.Encoding]::ASCII.GetBytes("GET /ping HTTP/1.1`r`nHost: 127.0.0.1`r`nConnection: close`r`n`r`n"); $s.Write($req, 0, $req.Length); $s.Flush()
            $b = New-Object byte[] 512; $n = $s.Read($b, 0, $b.Length); $c.Close()
            if ($n -gt 0) { return $true }
        } catch { Start-Sleep -Milliseconds 250 }
    }
    return $false
}
function Stop-MockGraph { param($Job) $Job | Stop-Job -ErrorAction SilentlyContinue; $Job | Remove-Job -Force -ErrorAction SilentlyContinue }
function Get-MockHits { param($Path, $Kind) if (-not (Test-Path $Path)) { return 0 }; return @(Get-Content $Path | Where-Object { $_ -match $Kind }).Count }
function Use-MockGraph { param([int]$Port) $script:LoginBase = "http://127.0.0.1:$Port"; $script:GraphBase = "http://127.0.0.1:$Port" }
function Show-H2Log { if ($script:fail -ne $script:H2FailBefore) { $script:H2Log | ForEach-Object { Write-Host "  log: $_" } }; $script:H2FailBefore = $script:fail }
$script:H2FailBefore = $script:fail
function Set-FreshSecret { param([int]$AgeMinutes = 0) $script:FreshSecretClientId = 'app-fresh'; $script:FreshSecretCreatedUtc = (Get-Date).ToUniversalTime().AddMinutes(-$AgeMinutes); $script:FreshSecretUntilUtc = $script:FreshSecretCreatedUtc.AddMinutes(20); $script:InstallerToken = $null; $script:H2Log = @() }
$script:FreshSecretRetrySeconds = 1; $script:PastedSecretRetrySeconds = 3; $script:FreshSecretNoteSeconds = 2; $script:QuietAccessDeniedHint = $false
Check "H2 static: install email honours the skip flag and caps its wait, wizard refuses a GUID as the secret" ($src -match 'if \(\$script:MailTestSkipped\) \{ Write-Log -Level WARNING' -and $src -match '-GraphSecret \$graphSecret -MaxWaitSeconds 120\)' -and $src -match 'Test-GuidLike \$graphSecret' -and $src -match "if \(\`$answer -eq 'S' -and -not \`$sent\) \{ \`$script:MailTestSkipped = \`$true \}")
$sendArgs = @{ TenantId = 'tid'; ClientId = 'app-fresh'; Secret = 'S1'; SenderAddress = 'hst@contoso.com'; To = @('a@contoso.com'); Subject = 'T'; Body = '<p>x</p>' }

# H2-1 fresh secret: two replicas answer 7000215, the third issues a token, the send goes through, one explanatory line
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 2 -LogPath $mockLog; $up = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
Check "H2 mock sign-in listener answers" $up
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs; $sw.Stop()
Check "H2-1 fresh secret rejected twice then accepted: send succeeds" ($ok -eq $true -and (Get-MockHits $mockLog '/token\|401') -eq 2 -and (Get-MockHits $mockLog '/token\|200') -eq 1 -and (Get-MockHits $mockLog '/sendMail\|202') -eq 1)
Check "H2-1 waited the retry interval between attempts" ($sw.Elapsed.TotalSeconds -ge 1.8)
Check "H2-1 one replication line naming the code and the deadline, no failure class" (@($script:H2Log | Where-Object { $_ -match '^INFORMATIONAL\|The sign-in service has not accepted the new secret yet \(AADSTS7000215\)\. Usually a replication delay that clears within minutes\. Retrying every 1 s until \d\d:\d\d\.$' }).Count -eq 1 -and @($script:H2Log | Where-Object { $_ -match '^(FAILED|WARNING)\|' }).Count -eq 0 -and $script:LastGraphSendError -eq '')
Check "H2-1 bearer token cached with expiry" ($script:InstallerToken.AccessToken -eq 'tok3' -and $script:InstallerToken.ExpiresUtc -gt (Get-Date).ToUniversalTime().AddMinutes(50))
# H2-2 the token is reused: a second send makes no token request
$ok2 = Send-GraphMail @sendArgs
Check "H2-2 second send reuses the token" ($ok2 -eq $true -and (Get-MockHits $mockLog '/token\|') -eq 3 -and (Get-MockHits $mockLog '/sendMail\|202') -eq 2)
# H2-3 a different secret is a cache miss
$sa2 = $sendArgs.Clone(); $sa2['Secret'] = 'S2'
$ok3 = Send-GraphMail @sa2
Check "H2-3 a different secret requests a new token" ($ok3 -eq $true -and (Get-MockHits $mockLog '/token\|') -eq 4 -and $script:InstallerToken.Secret -eq 'S2')
# H2-3b an expired cached token is replaced
$script:InstallerToken.ExpiresUtc = (Get-Date).ToUniversalTime().AddMinutes(4)
$ok3b = Send-GraphMail @sa2
Check "H2-3b a token within 5 minutes of expiry is refreshed" ($ok3b -eq $true -and (Get-MockHits $mockLog '/token\|') -eq 5)
Show-H2Log; Stop-MockGraph $job

# H2-4 pasted secret always rejected: bounded retry, then InvalidSecret with the value-not-ID guidance, no send attempted
$script:FreshSecretClientId = ''; $script:InstallerToken = $null; $script:H2Log = @()
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 99 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs; $sw.Stop()
$hits = Get-MockHits $mockLog '/token\|401'
Check "H2-4 pasted secret rejected: bounded retry then InvalidSecret, nothing sent" ($ok -eq $false -and $script:LastGraphSendError -eq 'InvalidSecret' -and $hits -ge 3 -and $hits -le 6 -and (Get-MockHits $mockLog '/sendMail') -eq 0 -and $sw.Elapsed.TotalSeconds -ge 2.5 -and $sw.Elapsed.TotalSeconds -lt 15)
Check "H2-4 pasted secret messages: retry notice, then value-not-ID guidance" ((($script:H2Log -join "`n") -match 'INFORMATIONAL\|The sign-in service rejected the client secret \(AADSTS7000215\)\. A secret created in the last few minutes can still be replicating, so this is retried every 1 s for 3 s\.') -and (($script:H2Log -join "`n") -match '(?m)^WARNING\|The sign-in service does not accept this client secret for app app-fresh in tenant tid\.$') -and (($script:H2Log -join "`n") -match 'FAILED\|Graph send failed as hst@contoso\.com\. .*AADSTS7000215'))
Show-H2Log; Stop-MockGraph $job

# H2-5 secret minted in this run but past its 20 minute window: short retry, then the past-replication message with the age
Set-FreshSecret -AgeMinutes 25
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 99 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$ok = Send-GraphMail @sendArgs
Check "H2-5 fresh secret past its window: one attempt, InvalidSecret, age reported, no advice about prompts" ($ok -eq $false -and $script:LastGraphSendError -eq 'InvalidSecret' -and (Get-MockHits $mockLog '/token\|401') -eq 1 -and (($script:H2Log -join "`n") -match '(?m)^WARNING\|The secret created in this run is still rejected 2[56] minutes after it was created, which is past any replication delay\.$') -and -not (($script:H2Log -join "`n") -match 'Choose R|retried'))
Show-H2Log; Stop-MockGraph $job

# H2-6 an unrelated sign-in error during the fresh window is not retried
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 99 -FailCode 'AADSTS90002' -FailStatus 400 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs; $sw.Stop()
Check "H2-6 unrelated sign-in error: one attempt, class Other, no replication line" ($ok -eq $false -and $script:LastGraphSendError -eq 'Other' -and (Get-MockHits $mockLog '/token\|400') -eq 1 -and $sw.Elapsed.TotalSeconds -lt 2 -and (($script:H2Log -join "`n") -match 'FAILED\|Graph send failed as hst@contoso\.com\. .*AADSTS90002') -and -not (($script:H2Log -join "`n") -match 'replicating|retried'))
Show-H2Log; Stop-MockGraph $job

# H2-7 Graph refuses the send with ErrorAccessDenied: AccessDenied class, Exchange hint, cached token dropped
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -SendStatus 403 -SendBody '{"error":{"code":"ErrorAccessDenied","message":"Access is denied. Check credentials and try again."}}' -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$ok = Send-GraphMail @sendArgs
Check "H2-7 access denied from Graph: AccessDenied class, hint logged, token dropped" ($ok -eq $false -and $script:LastGraphSendError -eq 'AccessDenied' -and $null -eq $script:InstallerToken -and (($script:H2Log -join "`n") -match 'FAILED\|Graph send failed as hst@contoso\.com\. .*\[ErrorAccessDenied\] Access is denied') -and (($script:H2Log -join "`n") -match 'INFORMATIONAL\|Access denied right after tenant setup usually means the Exchange sending grant'))
$script:QuietAccessDeniedHint = $true; $script:H2Log = @()
$ok = Send-GraphMail @sendArgs
Check "H2-7 next send requests a fresh token, hint silenced during the quiet wait" ((Get-MockHits $mockLog '/token\|200') -eq 2 -and -not (($script:H2Log -join "`n") -match 'Access denied right after tenant setup'))
$script:QuietAccessDeniedHint = $false
Show-H2Log; Stop-MockGraph $job

# H2-8 a replica that does not know the app yet (AADSTS700016) is retried like a rejected secret
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 1 -FailCode 'AADSTS700016' -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$ok = Send-GraphMail @sendArgs
Check "H2-8 app-not-found from a lagging replica is retried and the send succeeds" ($ok -eq $true -and (Get-MockHits $mockLog '/token\|401') -eq 1 -and (Get-MockHits $mockLog '/token\|200') -eq 1 -and (($script:H2Log -join "`n") -match 'has not accepted the new secret yet \(AADSTS700016\)'))
Show-H2Log; Stop-MockGraph $job

# H2-10 a transient 503 inside the fresh window is retried, then the send succeeds
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 1 -FailCode 'ServerError' -FailStatus 503 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$ok = Send-GraphMail @sendArgs
Check "H2-10 transient 503 during the fresh window is retried" ($ok -eq $true -and (Get-MockHits $mockLog '/token\|503') -eq 1 -and (Get-MockHits $mockLog '/token\|200') -eq 1 -and (($script:H2Log -join "`n") -match 'has not accepted the new secret yet \(HTTP 503\)'))
Show-H2Log; Stop-MockGraph $job

# H2-11 the same 503 for a pasted secret is not retried
$script:FreshSecretClientId = ''; $script:InstallerToken = $null; $script:H2Log = @()
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 99 -FailCode 'ServerError' -FailStatus 503 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs; $sw.Stop()
Check "H2-11 transient error outside the fresh window: one attempt, class Other" ($ok -eq $false -and $script:LastGraphSendError -eq 'Other' -and (Get-MockHits $mockLog '/token\|503') -eq 1 -and $sw.Elapsed.TotalSeconds -lt 2)
Show-H2Log; Stop-MockGraph $job

# H2-12 MaxWaitSeconds caps the fresh window for a caller that must not block, and the waiting line appears
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -TokenFailures 99 -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs -MaxWaitSeconds 4; $sw.Stop()
$hits = Get-MockHits $mockLog '/token\|401'
Check "H2-12 capped wait: gives up after about 4 s inside a 20 min window, still classed InvalidSecret" ($ok -eq $false -and $script:LastGraphSendError -eq 'InvalidSecret' -and $hits -ge 4 -and $hits -le 7 -and $sw.Elapsed.TotalSeconds -ge 3.5 -and $sw.Elapsed.TotalSeconds -lt 12)
Check "H2-12 a still-waiting line with the give-up time appears during the wait" (@($script:H2Log | Where-Object { $_ -match '^INFORMATIONAL\|Still waiting for the sign-in service to accept the secret \(AADSTS7000215\)\. Giving up at \d\d:\d\d\.$' }).Count -ge 1)
Show-H2Log; Stop-MockGraph $job

# H2-13 a tenant correction is a cache miss even with the same client and secret
Set-FreshSecret
$port = Get-FreeTestPort; $job = Start-MockGraph -Port $port -LogPath $mockLog; $null = Wait-MockGraph -Port $port; Use-MockGraph -Port $port
$null = Send-GraphMail @sendArgs                      # tid, S1: first token
$sa4 = $sendArgs.Clone(); $sa4['Secret'] = 's1'
$null = Send-GraphMail @sa4                           # same tenant and client, secret differs only by case: must miss
$sa3 = $sendArgs.Clone(); $sa3['TenantId'] = 'tid2'
$null = Send-GraphMail @sa3                           # tenant differs: must miss
$null = Send-GraphMail @sa3                           # identical: hit
Check "H2-13 secret case change and tenant change each request a new token, identical call reuses it" ((Get-MockHits $mockLog '/tid/oauth2') -eq 2 -and (Get-MockHits $mockLog '/tid2/oauth2') -eq 1 -and (Get-MockHits $mockLog '/token\|200') -eq 3 -and (Get-MockHits $mockLog '/sendMail\|202') -eq 4)
Show-H2Log; Stop-MockGraph $job

# H2-9 a listener that never answers: the token call times out and is reported, not hung
Set-FreshSecret
$deadPort = Get-FreeTestPort
Use-MockGraph -Port $deadPort
$sw = [Diagnostics.Stopwatch]::StartNew(); $ok = Send-GraphMail @sendArgs; $sw.Stop()
Check "H2-9 refused connection: class Other, no retry loop" ($ok -eq $false -and $script:LastGraphSendError -eq 'Other' -and $sw.Elapsed.TotalSeconds -lt 10)
Show-H2Log

if (Test-Path $mockLog) { Remove-Item $mockLog -Force }
${function:Write-Log} = $h2WriteLog
$script:LoginBase = 'https://login.microsoftonline.com'; $script:GraphBase = 'https://graph.microsoft.com'; $script:FreshSecretClientId = ''; $script:InstallerToken = $null

Section "I. Browser sign-in flow (listener, PKCE, state, redemption)"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('ConvertTo-Base64Url','ConvertFrom-JwtClaims','Get-FreeTcpPort','Get-BrowserToken') }) { Invoke-Expression $f.Extent.Text }
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
# RFC 7636 test vector
$v='dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk'
Check "PKCE S256 matches RFC 7636 vector" ((ConvertTo-Base64Url ([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::ASCII.GetBytes($v)))) -eq 'E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM')
Check "Free port is ephemeral" ((Get-FreeTcpPort) -gt 1024)
# fake JWT
function NewJwt($claims){ $h=ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes('{"alg":"none"}')); $p=ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes(($claims|ConvertTo-Json -Compress))); "$h.$p.sig" }
$jwt = NewJwt @{ tid='11111111-2222-3333-4444-555555555555'; upn='admin@contoso.com' }
Check "JWT claims decode (base64url padding)" ((ConvertFrom-JwtClaims $jwt).upn -eq 'admin@contoso.com')

# Browser simulation: Start-Process stub parses the authorize URL and calls back the redirect like a browser would
$script:AuthUrl=$null; $script:TokenBody=$null; $script:BrowserMode='ok'
function Start-Process { param($FilePath,$ArgumentList,[switch]$Wait) $script:AuthUrl=$FilePath
  $q=[System.Web.HttpUtility]::ParseQueryString(([uri]$FilePath).Query); $redir=$q['redirect_uri']; $state=$q['state']
  $cb = switch ($script:BrowserMode) { 'ok' { "$redir`?code=AUTHCODE123&state=$state" } 'badstate' { "$redir`?code=AUTHCODE123&state=wrong" } 'denied' { "$redir`?error=access_denied&error_description=User+cancelled" } }
  Start-Job -ScriptBlock { param($u,$r) Start-Sleep -Milliseconds 300; try { Invoke-WebRequest -Uri ($r + 'favicon.ico') -UseBasicParsing -TimeoutSec 5 | Out-Null } catch { }; try { (Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 10).Content } catch { "ERR $_" } } -ArgumentList $cb,$redir | Out-Null }
function Invoke-RestMethod { param($Method,$Uri,$Body,$ContentType,$Headers,$ErrorAction) $script:TokenBody=$Body; [PSCustomObject]@{ access_token=$jwt; refresh_token='RT'; expires_in=3600 } }
function Write-Log { param($Level,$Message) $script:LastLog="$Level|$Message" }
$cid='04b07795-8ddb-461a-bbee-02f9e1bf7b46'; $scope='https://graph.microsoft.com/.default offline_access'
$r = Get-BrowserToken -ClientId $cid -Scope $scope
Get-Job | Wait-Job -Timeout 5 | Out-Null; Get-Job | Remove-Job -Force
Check "Browser flow: token returned with tenant and upn" ($r -and $r.TenantId -eq '11111111-2222-3333-4444-555555555555' -and $r.Upn -eq 'admin@contoso.com' -and $r.RefreshToken -eq 'RT')
$aq=[System.Web.HttpUtility]::ParseQueryString(([uri]$script:AuthUrl).Query)
Check "Authorize URL: PKCE S256, select_account, localhost redirect, scope, state" ($aq['code_challenge_method'] -eq 'S256' -and $aq['prompt'] -eq 'select_account' -and $aq['redirect_uri'] -like 'http://localhost:*/' -and $aq['scope'] -eq $scope -and $aq['client_id'] -eq $cid -and $aq['response_type'] -eq 'code')
Check "Token request: authorization_code with verifier and matching redirect" ($script:TokenBody.grant_type -eq 'authorization_code' -and $script:TokenBody.code -eq 'AUTHCODE123' -and $script:TokenBody.code_verifier -and $script:TokenBody.redirect_uri -eq $aq['redirect_uri'])
Check "Verifier matches challenge" ((ConvertTo-Base64Url ([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::ASCII.GetBytes($script:TokenBody.code_verifier)))) -eq $aq['code_challenge'])
$script:BrowserMode='badstate'; $r = Get-BrowserToken -ClientId $cid -Scope $scope; Get-Job | Wait-Job -Timeout 5 | Out-Null; Get-Job | Remove-Job -Force
Check "Browser flow: state mismatch rejected" ($null -eq $r -and $script:LastLog -match 'did not match')
$script:BrowserMode='denied'; $r = Get-BrowserToken -ClientId $cid -Scope $scope; Get-Job | Wait-Job -Timeout 5 | Out-Null; Get-Job | Remove-Job -Force
Check "Browser flow: user cancel surfaces error and returns null" ($null -eq $r -and $script:LastLog -match 'access_denied')
Check "Stray favicon request ignored, real redirect still accepted" ($src -match "Not the redirect \(favicon or pre-connect probe\)" -and $src -match 'StatusCode = 204')
Check "Listener bind retried" ($src -match '\$bindTry -lt 5')
Check "No System.Web dependency" (-not ($src -match 'System\.Web'))
Check "Role assignment retried on fresh SP" ($src -match 'for \(\$i = 0; \$i -lt 6 -and -not \$granted')
Check "Fresh secret window opened at minting, token obtained before the test email" ($src -match '\$script:FreshSecretUntilUtc = \$script:FreshSecretCreatedUtc\.AddMinutes\(20\)' -and $src -match 'Get-InstallerGraphToken -TenantId \$tenantId -ClientId \$app\.appId -Secret \$secret' -and -not ($src -match 'Wait-GraphAppReady'))
Check "Listener released after each attempt (new listener binds)" ((Get-FreeTcpPort) -gt 0)
Check "Fallback order: window/browser then device code" ($src -match 'Get-BrowserToken -ClientId \$clientId -Scope \$scope -Title' -and $src -match 'Falling back to a device code\."; \$auth = Get-DeviceCodeToken')
$script:BrowserMode='ok'; $r = Get-BrowserToken -ClientId $cid -Scope $scope -Prompt ''; Get-Job | Wait-Job -Timeout 5 | Out-Null; Get-Job | Remove-Job -Force
$aq2=[System.Web.HttpUtility]::ParseQueryString(([uri]$script:AuthUrl).Query)
Check "Silent second sign-in omits prompt parameter" ($r -and $null -eq $aq2['prompt'])
Check "Normal path never mentions a device code" (-not ($src.Substring($src.IndexOf('function Get-BrowserToken'), $src.IndexOf('function Get-DeviceCodeToken') - $src.IndexOf('function Get-BrowserToken')) -match 'user_code'))


Section "J. Exchange REST transport"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('Invoke-ExoCommand','Invoke-ExoCmdlet') }) { Invoke-Expression $f.Extent.Text }
$script:RestCalls=@()
function Invoke-RestMethod { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction) $script:RestCalls += [PSCustomObject]@{ Method=$Method; Uri=$Uri; Headers=$Headers; Body=($Body|ConvertFrom-Json) }
  if ($Body -match 'Get-Mailbox' -and $Body -match 'missing@') { throw "not found" }
  [PSCustomObject]@{ value = @([PSCustomObject]@{ PrimarySmtpAddress='hst@contoso.com'; AccessCheckResult='Granted' }) } }
$r = Invoke-ExoCommand -Token 'T' -TenantId 'tid' -Upn 'admin@contoso.com' -Cmdlet 'New-Mailbox' -Parameters @{ Shared=$true; Name='hst'; PrimarySmtpAddress='hst@contoso.com' }
$c = $script:RestCalls[-1]
Check "REST: endpoint and method" ($c.Method -eq 'Post' -and $c.Uri -eq 'https://outlook.office365.com/adminapi/beta/tid/InvokeCommand')
Check "REST: bearer + anchor mailbox headers" ($c.Headers.Authorization -eq 'Bearer T' -and $c.Headers['X-AnchorMailbox'] -eq 'UPN:admin@contoso.com')
Check "REST: CmdletInput payload carries cmdlet and params" ($c.Body.CmdletInput.CmdletName -eq 'New-Mailbox' -and $c.Body.CmdletInput.Parameters.Shared -eq $true -and $c.Body.CmdletInput.Parameters.PrimarySmtpAddress -eq 'hst@contoso.com')
Check "REST: returns value array" (@($r).Count -eq 1 -and $r[0].PrimarySmtpAddress -eq 'hst@contoso.com')
$script:ExoMode='rest'; $script:ExoToken='T'; $script:ExoTenantId='tid'; $script:ExoUpn='admin@contoso.com'
Check "Dispatch rest: NullOnNotFound swallows error" ($null -eq (Invoke-ExoCmdlet -Name 'Get-Mailbox' -Parameters @{ Identity='missing@contoso.com' } -NullOnNotFound))
$threw=$false; try { Invoke-ExoCmdlet -Name 'Get-Mailbox' -Parameters @{ Identity='missing@contoso.com' } | Out-Null } catch { $threw=$true }
Check "Dispatch rest: errors propagate without the switch" $threw
function Get-Mailbox { param($Identity,$ErrorAction) if ($Identity -like 'missing*') { throw "not found" }; [PSCustomObject]@{ PrimarySmtpAddress=$Identity; Via='module' } }
$script:ExoMode='module'
Check "Dispatch module: calls local cmdlet with splat" (((Invoke-ExoCmdlet -Name 'Get-Mailbox' -Parameters @{ Identity='x@contoso.com' })[0]).Via -eq 'module')
Check "Dispatch module: not-found swallowed" ($null -eq (Invoke-ExoCmdlet -Name 'Get-Mailbox' -Parameters @{ Identity='missing@contoso.com' } -NullOnNotFound))


Section "K. Monitor name, URL, content marker, and migration from the older install"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object { $_.Name -in @('ConvertTo-MonitorSlug','Test-MonitorUrl','Get-MonitorName','Get-MonitorUrl','Get-ContentMarker','Get-InstalledMonitor','Get-LegacyInstall','Invoke-LegacyMigration','Get-SlugClash','Test-MonitorNameLength','Protect-StoredSecret') }) { Invoke-Expression $f.Extent.Text }
$NonInteractive = $false; $MonitorNameOverride = ''; $Url = ''; $ExpectedContentMarker = ''
function Write-Log { param($Level,$Message) $script:LastLog = "$Level|$Message"; $script:Logs += "$Level|$Message" }
$script:Logs = @()
$kRoot = '/tmp/curlmon_test'
if (Test-Path $kRoot) { Remove-Item $kRoot -Recurse -Force }
New-Item $kRoot -ItemType Directory | Out-Null
$InstallRoot = $kRoot

Check "K1 slug keeps letters and digits, collapses the rest, trims dashes" ((ConvertTo-MonitorSlug 'HST eChart') -eq 'HST-eChart' -and (ConvertTo-MonitorSlug '  --HST  eChart!!  ') -eq 'HST-eChart' -and (ConvertTo-MonitorSlug 'Portal (prod) / v2') -eq 'Portal-prod-v2' -and (ConvertTo-MonitorSlug '***') -eq '' -and (ConvertTo-MonitorSlug $null) -eq '')
Check "K1 slug stays inside a sensible folder length" (((ConvertTo-MonitorSlug ('a' * 200)).Length -le 60) -and -not ((ConvertTo-MonitorSlug ('a' * 200)).EndsWith('-')))
Check "K2 URL check takes http and https only" ((Test-MonitorUrl 'https://example.com/x') -and (Test-MonitorUrl 'http://10.0.0.5:8080/health') -and -not (Test-MonitorUrl 'example.com') -and -not (Test-MonitorUrl 'ftp://example.com') -and -not (Test-MonitorUrl '') -and -not (Test-MonitorUrl 'not a url'))

Push @('', 'not-a-url', 'ftp://x.example', 'https://portal.example.com/health')
$u = Get-MonitorUrl -SavedDefault ''
Check "K2 prompt refuses anything that is not an http URL and keeps asking" ($u -eq 'https://portal.example.com/health' -and $script:PromptCount -eq 4)
Push @('')
Check "K2 Enter keeps the saved URL" ((Get-MonitorUrl -SavedDefault 'https://saved.example.com/a') -eq 'https://saved.example.com/a' -and $script:PromptCount -eq 1)

Push @('Site Portal')
$m1 = Get-MonitorName -SavedDefault '' -Existing @()
Check "K3 monitor name prompt returns the cleaned name" ($m1 -eq 'Site Portal' -and $script:PromptCount -eq 1)
Push @('***', 'HST eChart')
$m2 = Get-MonitorName -SavedDefault '' -Existing @()
Check "K3 a name with no letters or digits is refused" ($m2 -eq 'HST eChart')
Push @('')
Check "K3 Enter takes the only monitor already installed" ((Get-MonitorName -SavedDefault '' -Existing @([PSCustomObject]@{ Name = 'HST eChart'; Url = 'https://x' })) -eq 'HST eChart')

Push @('')
Check "K4 Enter keeps the saved content marker" ((Get-ContentMarker -SavedDefault 'Sign In') -eq 'Sign In')
Push @('-')
Check "K4 a single dash clears the marker" ((Get-ContentMarker -SavedDefault 'Sign In') -eq '')
Push @('Welcome back')
Check "K4 typed text becomes the marker" ((Get-ContentMarker -SavedDefault '') -eq 'Welcome back')

# Two monitors on one server: different folders, tasks, files, and subjects
$MonitorName = 'HST eChart'; $SiteName = 'CapCity'
$InstallDir = Join-Path $kRoot (ConvertTo-MonitorSlug $MonitorName); $Url = 'https://a.example.com/one'
$genA = New-MonitorContent -SiteName 'CapCity' -Mail $mRelay
$MonitorName = 'Patient Portal'
$InstallDir = Join-Path $kRoot (ConvertTo-MonitorSlug $MonitorName); $Url = 'https://b.example.com/two'
$genB = New-MonitorContent -SiteName 'CapCity' -Mail $mRelay
Check "K5 two monitors bake different names, URLs, folders, and data files" ($genA -match "(?m)^\`$MonitorName\s+=\s+'HST eChart'$" -and $genB -match "(?m)^\`$MonitorName\s+=\s+'Patient Portal'$" -and $genA -match "(?m)^\`$Url\s+=\s+'https://a\.example\.com/one'$" -and $genB -match "(?m)^\`$Url\s+=\s+'https://b\.example\.com/two'$" -and $genA -match "HST-eChart\\heartbeat\.json" -and $genB -match "Patient-Portal\\heartbeat\.json" -and (ParseOk $genA) -and (ParseOk $genB))
Check "K5 the installer names the folder and the task after the monitor" ($src -match '\$InstallDir = Join-Path \$InstallRoot \(ConvertTo-MonitorSlug \$MonitorName\)' -and $src -match '\$TaskName   = "\$TaskNamePrefix\$MonitorName"' -and $src -match '(?m)^\$TaskNamePrefix\s+=\s+"Curl Monitor - "')
$labA = Get-AlertLabel -MonitorName 'HST eChart' -SiteName 'CapCity' -HostName 'HOST1'
$labB = Get-AlertLabel -MonitorName 'Patient Portal' -SiteName 'CapCity' -HostName 'HOST1'
Check "K5 alert labels tell the two monitors apart on the same server" ($labA -eq 'HST eChart at CapCity (HOST1)' -and $labB -eq 'Patient Portal at CapCity (HOST1)' -and (Get-AlertLabel -MonitorName '' -SiteName 'CapCity' -HostName 'HOST1') -eq 'CapCity (HOST1)')

# Re-running with the same name upgrades in place
$upgradeDir = Join-Path $kRoot 'HST-eChart'
New-Item $upgradeDir -ItemType Directory -Force | Out-Null
Set-Content (Join-Path $upgradeDir 'Watch-CurlMonitor.ps1') -Value '# monitor' -Encoding UTF8
@{ MonitorName = 'HST eChart'; SiteName = 'CapCity'; Url = 'https://a.example.com/one'; ContentMarker = 'Sign In' } | ConvertTo-Json | Set-Content (Join-Path $upgradeDir 'install-settings.json') -Encoding UTF8
$installed = @(Get-InstalledMonitor -Root $kRoot)
Check "K6 installed monitors are listed with their names and URLs" ($installed.Count -eq 1 -and $installed[0].Name -eq 'HST eChart' -and $installed[0].Slug -eq 'HST-eChart' -and $installed[0].Url -eq 'https://a.example.com/one')
Push @('')
Check "K6 re-running defaults to that monitor, so the same folder is upgraded" (((ConvertTo-MonitorSlug (Get-MonitorName -SavedDefault 'HST eChart' -Existing $installed)) -eq 'HST-eChart') -and (Test-Path $upgradeDir))

# Migration from the older HST-only layout, against scratch folders and a task that does not exist
$legacyDir = Join-Path $kRoot 'old-HSTProbe'
New-Item $legacyDir -ItemType Directory | Out-Null
@{ SiteName = 'CapCity'; Url = 'https://legacy.example.com/echart'; MailMethod = 'Graph'; GraphClientId = 'cid'; CredentialFor = 'cid' } | ConvertTo-Json | Set-Content (Join-Path $legacyDir 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $legacyDir 'Watch-HSTeChartUptime.ps1') -Value '# old monitor' -Encoding UTF8
Set-Content (Join-Path $legacyDir 'HST-eChart-Latency_202609.csv') -Value 'a,b' -Encoding UTF8
Set-Content (Join-Path $legacyDir 'HST-eChart-Outages.csv') -Value 'c,d' -Encoding UTF8
Set-Content (Join-Path $legacyDir 'HST-eChart-Drops.log') -Value '2026-09-15 06:16:14 | SLOW      | x' -Encoding UTF8
Set-Content (Join-Path $legacyDir 'HST-eChart-Monitor_20260915.log') -Value 'transcript' -Encoding UTF8
Set-Content (Join-Path $legacyDir 'smtp-credential.bin') -Value 'Y2lwaGVy' -Encoding ASCII
Set-Content (Join-Path $legacyDir 'daily-summary-sent.txt') -Value '2026-09-15' -Encoding ASCII
@{ Beat = '2026-09-15T10:00:00.0000000Z'; IsDown = $true; ConsecutiveFailures = 5; Stopped = $false } | ConvertTo-Json | Set-Content (Join-Path $legacyDir 'monitor-heartbeat.json') -Encoding UTF8
$legacy = Get-LegacyInstall -Dir $legacyDir -TaskName 'No Such Curl Task' -Path '\DIT\'
Check "K7 the older install is found with its settings and URL" ($legacy -and -not $legacy.HasTask -and $legacy.Url -eq 'https://legacy.example.com/echart' -and $legacy.SiteName -eq 'CapCity')
Check "K7 no older install means nothing to migrate" ($null -eq (Get-LegacyInstall -Dir (Join-Path $kRoot 'missing') -TaskName 'No Such Curl Task' -Path '\DIT\'))
$newDir = Join-Path $kRoot 'HST-eChart-migrated'
New-Item $newDir -ItemType Directory | Out-Null
$migrated = Invoke-LegacyMigration -Legacy $legacy -Destination $newDir
$names = @(Get-ChildItem $newDir -File -Force | ForEach-Object { $_.Name } | Sort-Object)
Check "K8 history moves under the new names and the old folder is gone" ($migrated -and -not (Test-Path $legacyDir) -and (@($names | Where-Object { $_ -notlike 'legacy-*' }) -join ',') -eq 'credential.bin,Drops.log,heartbeat.json,install-settings.json,Latency_202609.csv,Outages.csv,summary-sent.txt,Transcript_20260915.log')
Check "K8 carried settings prefill the new prompts" ((Get-Content (Join-Path $newDir 'install-settings.json') -Raw | ConvertFrom-Json).Url -eq 'https://legacy.example.com/echart')
$hbM = Get-Content (Join-Path $newDir 'heartbeat.json') -Raw | ConvertFrom-Json
Check "K8 the carried heartbeat is marked stopped, so the new monitor sends no restart notice and still carries the outage" ($hbM.Stopped -eq $true -and $hbM.IsDown -eq $true -and [int]$hbM.ConsecutiveFailures -eq 5)
Check "K8 the drops log records the move" ((Get-Content (Join-Path $newDir 'Drops.log') -Raw) -match '\| STOP      \| Monitor stopped by the installer on .+ and moved to ')

# A file that cannot be copied leaves the old folder in place
$legacyDir2 = Join-Path $kRoot 'old-HSTProbe2'
New-Item $legacyDir2 -ItemType Directory | Out-Null
@{ SiteName = 'CapCity'; Url = 'https://legacy2.example.com/x' } | ConvertTo-Json | Set-Content (Join-Path $legacyDir2 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $legacyDir2 'HST-eChart-Outages.csv') -Value 'e,f' -Encoding UTF8
Set-Content (Join-Path $legacyDir2 'HST-eChart-Drops.log') -Value 'x' -Encoding UTF8
$newDir2 = Join-Path $kRoot 'HST-eChart-blocked'
New-Item $newDir2 -ItemType Directory | Out-Null
$blocker = [System.IO.File]::Open((Join-Path $legacyDir2 'HST-eChart-Outages.csv'), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
$legacy2 = Get-LegacyInstall -Dir $legacyDir2 -TaskName 'No Such Curl Task' -Path '\DIT\'
$migrated2 = Invoke-LegacyMigration -Legacy $legacy2 -Destination $newDir2
$blocker.Close()
Check "K9 a file that will not copy keeps the old folder and says so" (-not $migrated2 -and (Test-Path $legacyDir2) -and (($script:Logs -join "`n") -match "did not copy, so '.*old-HSTProbe2' is left in place"))
# A destination file that is locked or already newer is never overwritten: the legacy copy lands beside it
$legacyDir3 = Join-Path $kRoot 'old-HSTProbe3'
New-Item $legacyDir3 -ItemType Directory | Out-Null
@{ SiteName = 'CapCity'; Url = 'https://legacy3.example.com/x' } | ConvertTo-Json | Set-Content (Join-Path $legacyDir3 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $legacyDir3 'HST-eChart-Outages.csv') -Value 'old,short' -Encoding UTF8
$newDir3 = Join-Path $kRoot 'HST-eChart-kept'
New-Item $newDir3 -ItemType Directory | Out-Null
Set-Content (Join-Path $newDir3 'Outages.csv') -Value 'newer history that must survive' -Encoding UTF8
$legacy3 = Get-LegacyInstall -Dir $legacyDir3 -TaskName 'No Such Curl Task' -Path '\DIT\'
$migrated3 = Invoke-LegacyMigration -Legacy $legacy3 -Destination $newDir3
Check "K9 history already in the new folder survives and the legacy copy lands beside it" ($migrated3 -and (Get-Content (Join-Path $newDir3 'Outages.csv')) -eq 'newer history that must survive' -and @(Get-ChildItem $newDir3 -Filter 'legacy-*Outages.csv').Count -eq 1)
Remove-Item $kRoot -Recurse -Force -ErrorAction SilentlyContinue
$InstallRoot = 'C:\ProgramData\DIT\CurlMonitor'


Section "L. Fixes from the stress campaign"
$TaskNamePrefix = 'Curl Monitor - '; $MaxTaskNameLength = 238
$lExisting = @([PSCustomObject]@{ Name = 'HST eChart'; Slug = 'HST-eChart'; Path = 'C:\x\HST-eChart'; Url = 'u' }, [PSCustomObject]@{ Name = 'Portal'; Slug = 'Portal'; Path = 'C:\x\Portal'; Url = 'v' })
Check "L1 a name that cleans to another monitor's folder is reported, whatever the punctuation or case" ((Get-SlugClash -Name 'HST_eChart' -Existing $lExisting).Name -eq 'HST eChart' -and (Get-SlugClash -Name 'hst echart' -Existing $lExisting).Name -eq 'HST eChart' -and (Get-SlugClash -Name 'HST.eChart' -Existing $lExisting).Name -eq 'HST eChart')
Check "L1 the same name and an unrelated name are not clashes" ($null -eq (Get-SlugClash -Name 'HST eChart' -Existing $lExisting) -and $null -eq (Get-SlugClash -Name 'Billing Portal' -Existing $lExisting))
Check "L1 wizard offers to upgrade the clashing monitor and returns its installed name" ($src -match "U = upgrade '\`$\(\`$clash\.Name\)' in place, N = type another name")
Push @('HST_eChart', 'U')
$lName = Get-MonitorName -SavedDefault '' -Existing $lExisting
Check "L1 answering U adopts the installed name instead of overwriting its folder" ($lName -eq 'HST eChart' -and $script:PromptCount -eq 2 -and (($script:Logs -join "`n") -match "uses the same folder as the installed monitor 'HST eChart'"))
Push @('HST_eChart', 'N', 'Billing Portal')
$lName = Get-MonitorName -SavedDefault '' -Existing $lExisting
Check "L1 answering N asks again and takes a name with its own folder" ($lName -eq 'Billing Portal' -and $script:PromptCount -eq 3)

Check "L2 a monitor name too long for a task name is refused before anything is written" ((Test-MonitorNameLength 'HST eChart') -and -not (Test-MonitorNameLength ('x' * 230)) -and (($script:Logs -join "`n") -match 'Monitor name is too long'))
Push @(('y' * 240), 'Short Name')
$lName = Get-MonitorName -SavedDefault '' -Existing @()
Check "L2 the wizard keeps asking until the name fits the task name limit" ($lName -eq 'Short Name' -and $script:PromptCount -eq 2)

$NonInteractive = $true; $MonitorNameOverride = ''; $MonitorName = 'RMM Named'
$lName = Get-MonitorName -SavedDefault '' -Existing @()
Check "L3 a silent install falls back to the config block monitor name" ($lName -eq 'RMM Named')
$MonitorName = 'HST_eChart'
$lName = Get-MonitorName -SavedDefault '' -Existing $lExisting
Check "L3 a silent install upgrades the clashing monitor rather than overwriting it" ($lName -eq 'HST eChart' -and (($script:Logs -join "`n") -match 'Upgrading that monitor instead of overwriting it'))
$MonitorName = ('z' * 240)
Check "L3 a silent install refuses a name too long for a task name" ((Get-MonitorName -SavedDefault '' -Existing @()) -eq '')
$MonitorName = ''

$Url = 'https://config-block.example/health'
Check "L4 a silent re-deploy repoints the monitor from the config block" ((Get-MonitorUrl -SavedDefault 'https://old.example/health') -eq 'https://config-block.example/health')
$ExpectedContentMarker = 'From config'
Check "L4 a silent re-deploy takes the marker from the config block too" ((Get-ContentMarker -SavedDefault 'Saved text') -eq 'From config')
$ExpectedContentMarker = ''
Check "L4 a silent re-deploy with no config marker keeps the saved one" ((Get-ContentMarker -SavedDefault 'Saved text') -eq 'Saved text')
$NonInteractive = $false; $Url = ''

Push @('')
$m1 = Get-ContentMarker -SavedDefault 'Sign in'
Check "L5 Enter keeps the saved text and the prompt says so" ($m1 -eq 'Sign in' -and $script:PromptLog[0] -eq "Required text (Enter = keep 'Sign in', - = no text check)")
Push @('-')
$m2 = Get-ContentMarker -SavedDefault 'Sign in'
Check "L5 a dash drops the text check on a re-run" ($m2 -eq '' -and (($script:Logs -join "`n") -match 'No text check'))
Push @('')
$m3 = Get-ContentMarker -SavedDefault ''
Check "L5 with nothing saved the prompt still offers blank for none" ($m3 -eq '' -and $script:PromptLog[0] -eq 'Required text (blank for none)')

$lMig = Join-Path $kRoot 'mig'
$lLegacy = Join-Path $lMig 'legacy'; $lDest = Join-Path $lMig 'dest'
New-Item (Join-Path $lLegacy 'Archive') -ItemType Directory -Force | Out-Null
New-Item $lDest -ItemType Directory -Force | Out-Null
Set-Content (Join-Path $lLegacy 'HST-eChart-Outages.csv') -Value 'old'
Set-Content (Join-Path $lLegacy 'Archive/HST-eChart-Latency_202501.csv') -Value 'archived rows'
Set-Content (Join-Path $lLegacy 'notes-from-the-admin.txt') -Value 'keep me'
Set-Content (Join-Path $lDest 'Outages.csv') -Value 'newer history that must survive'
$lLegacyObj = [PSCustomObject]@{ Dir = $lLegacy; TaskName = 'x'; TaskPath = '\None\'; HasTask = $false; Settings = $null; Url = ''; SiteName = '' }
$lOk = Invoke-LegacyMigration -Legacy $lLegacyObj -Destination $lDest
Check "L6 migration reports success and removes the old folder" ($lOk -and -not (Test-Path $lLegacy))
Check "L6 history the new monitor already wrote is never overwritten" ((Get-Content (Join-Path $lDest 'Outages.csv')) -eq 'newer history that must survive' -and @(Get-ChildItem $lDest -Filter 'legacy-*Outages.csv').Count -eq 1)
Check "L6 files in subfolders and unrecognised files are carried, not deleted" (@(Get-ChildItem $lDest -Filter '*Latency_202501.csv').Count -eq 1 -and @(Get-ChildItem $lDest -Filter 'legacy-*notes-from-the-admin.txt').Count -eq 1)
Check "L6 the migrated credential is locked to SYSTEM and Administrators as it lands" ($src -match '(?s)Split-Path \$targetPath -Leaf\) -eq \$CredentialFileName.*?icacls\.exe .\$targetPath. /inheritance:r /grant:r')
Check "L6 a copy that fails keeps the old folder and says so" ($src -match 'file\(s\) did not copy, so .* is left in place')
Check "L7 migration runs after the monitor is written, just before the task is registered" ($src -match "(?s)Wrote and verified monitor script.*?if \(\`$migrate\) \{\s+if \(Invoke-LegacyMigration -Legacy \`$legacy -Destination \`$InstallDir\)" -and $src -match 'An abort before this point')
Check "L7 the installer acts on the migration result rather than ignoring it" ($src -match 'Migration did not finish' -and -not ($src -match '\$null = Invoke-LegacyMigration'))

Check "L8 folder hardening leaves stored secrets and other monitors' folders alone" ($src -match '(?m)^\s+if \(\$item\.PSIsContainer\) \{ continue \}' -and $src -match '(?m)^\s+if \(\$item\.Name -eq \$CredentialFileName\) \{ continue \}' -and -not ($src -match '/reset /T /C'))
Check "L8 every stored secret is re-locked after hardening, on both the normal and the migration path" ($src -match 'function Protect-StoredSecret' -and ([regex]::Matches($src, 'Protect-StoredSecret -Root \$InstallRoot')).Count -eq 2)

$lGen = New-MonitorContent -SiteName 'S' -Mail $mRelay
Check "L9 the probe caps what it downloads and the reason table names the cap" ($lGen -match '--max-filesize \$MaxBodyBytes' -and $lGen -match '(?m)^\$MaxBodyBytes\s+=\s+\d+$' -and $lGen -match "63 \{ return 'Response larger than the size limit' \}")
Check "L9 an oversized body is never scanned for the marker" ($lGen -match '\$bodyBytes -gt \$MaxBodyBytes' -and $lGen -match '\$bodyBytes = \(Get-Item \$tempBody\)\.Length')
Check "L10 transcripts are pruned at startup, not only on a day rollover" ($lGen -match '\$startupCutoff = \(Get-Date\)\.AddDays\(-\$LogRetentionDays\)' -and $lGen -match "(?s)Start-Transcript -Path \`$transcriptPath -Append.*?Transcript_\*\.log' -ErrorAction SilentlyContinue \| Where-Object \{ \`$_\.LastWriteTime -lt \`$startupCutoff \}")
Check "L10 the daily summary reads rotated drops logs inside the window" ($lGen -match "Get-ChildItem -Path \`$InstallDir -Filter 'Drops_\*\.log'" -and $lGen -match '\$_\.LastWriteTime -ge \$cutoff')

$lSeq = @(Rep (Pl 300) 20) + @(Rep (Pl 4500) 20) + @(Rep (Pl 300) 25)
$lRun = RunSlow -State (NewSlow) -Polls $lSeq -Start $T0
$lSlow = @($lRun.Emails | Where-Object { $_.D.EmailKind -eq 'Slow' })[0].D
Check "L11 a slow period counts only the polls from its first bad one" ($lSlow.State.Polls -eq 10 -and $lSlow.State.SlowPolls -eq 10 -and $lSlow.State.FailedPolls -eq 0)
$lResolved = @($lRun.Emails | Where-Object { $_.D.EmailKind -eq 'SlowResolved' })[0].D
Check "L11 SLOW RESOLVED reports the period's own polls, not the healthy ones before it" ($lResolved.EmailBody -match '<td[^>]*>Polls while slow</td><td[^>]*>20: 20 slower than 3000 ms, 0 failed</td>')

$lBack = NewSlow
$lBack.IsSlow = $true; $lBack.StartUtc = $T0.AddMinutes(30); $lBack.StartLocalStr = 'L'; $lBack.LastAlertUtc = $T0.AddMinutes(30); $lBack.Polls = 12; $lBack.SlowPolls = 12
$lBack.Samples = @(0..9 | ForEach-Object { [PSCustomObject]@{ Utc = $T0.AddMinutes(29).AddSeconds($_ * 15); LocalStr = 'F'; Slow = $true; Failed = $false; Ms = 5000 } })
$lStep = Update-SlowState -State $lBack -Result (SlowRes 250) -Failed $false -IsDown $false -NowUtc $T0 -SlowThresholdMs 3000 -WindowMinutes 5 -AlertPercent 50 -ClearPercent 10 -ReAlertMinutes 30 -AlertOnRecovery $true -SiteName 'CapCity' -Url 'http://x' -HostName 'HOST1' -MonitorName 'Demo'
Check "L12 a clock that steps back drops future samples and pulls the period to now" (@($lStep.State.Samples | Where-Object { $_.Utc -gt $T0 }).Count -eq 0 -and $lStep.State.Samples.Count -eq 1 -and ($null -eq $lStep.State.StartUtc -or $lStep.State.StartUtc -le $T0))
$lBack2 = NewSlow
$lBack2.IsSlow = $true; $lBack2.StartUtc = $T0.AddMinutes(-20); $lBack2.StartLocalStr = 'L'; $lBack2.LastAlertUtc = $T0.AddMinutes(40); $lBack2.Polls = 40; $lBack2.SlowPolls = 40
$lBack2.Samples = @(0..19 | ForEach-Object { [PSCustomObject]@{ Utc = $T0.AddMinutes(-4).AddSeconds($_ * 12); LocalStr = 'F'; Slow = $false; Failed = $false; Ms = 200 } })
$lStep2 = Update-SlowState -State $lBack2 -Result (SlowRes 200) -Failed $false -IsDown $false -NowUtc $T0 -SlowThresholdMs 3000 -WindowMinutes 5 -AlertPercent 50 -ClearPercent 10 -ReAlertMinutes 30 -AlertOnRecovery $true -SiteName 'CapCity' -Url 'http://x' -HostName 'HOST1' -MonitorName 'Demo'
Check "L12 a period whose last alert is in the future can still clear after the step back" ($lStep2.EmailKind -eq 'SlowResolved' -and -not $lStep2.State.IsSlow)

Section "M. Telemetry: alert-free installs and the uptime publisher"
$tRoot = Join-Path ([IO.Path]::GetTempPath()) "curlmon_telemetry_$PID"
if (Test-Path $tRoot) { Remove-Item $tRoot -Recurse -Force }
New-Item $tRoot -ItemType Directory | Out-Null

# The telemetry-only toggle
$saveAlerts = $AlertsEnabled
$AlertsEnabled = $false
$genQuiet = New-MonitorContent -SiteName 'S' -Mail @{ MailMethod='None'; SmtpServer=''; SmtpPort=25; SmtpUseSsl=$false; MailFrom=''; MailTo=@(); SmtpAuthUser=''; GraphTenantId=''; GraphClientId=''; GraphSecretExpires='' }
$AlertsEnabled = $saveAlerts
$genLoud = New-MonitorContent -SiteName 'S' -Mail $mRelay
Check "M1 alerts off bakes SendEmail false and still parses" ($genQuiet -match '(?m)^\$SendEmail\s+=\s+\$false$' -and (ParseOk $genQuiet) -and $genLoud -match '(?m)^\$SendEmail\s+=\s+\$true$')
Check "M1 every alert dispatch in the monitor is gated on SendEmail" (([regex]::Matches($template, 'if \(\$SendEmail -and')).Count -ge 3 -and $template -match 'if \(\$SendEmail\) \{ Send-AlertOrQueue -Subject \$summary\.Subject')
Check "M1 the monitor still records outages, slow periods, and the drops log when alerts are off" ($genQuiet -match "Write-DropLog -Kind 'DOWN'" -and $genQuiet -match 'Write-DropLog -Kind \$slow\.DropKind' -and $genQuiet -match "\$dropKind = 'SLOWSTART'" -and $genQuiet -match 'Write-CsvRow -Path \$OutageCsv' -and $genQuiet -match 'Write-CsvRow -Path \(Get-LatencyCsvPath\)')
Check "M1 installer skips the mail wizard, the credential, and the install email when alerts are off" ($src -match 'if \(\$AlertsEnabled\) \{\s*\r?\n\s*\$mail = Get-MailConfiguration' -and $src -match "MailMethod = 'None'" -and $src -match 'if \(\$SendInstallTestEmail -and \$AlertsEnabled\)' -and $src -match 'Alerts          : off, telemetry only')

# Publisher functions, loaded from the real script
$pubPath = 'C:\Workspaces\HST Monitor\Publish-UptimeTelemetry.ps1'
$pubAst = [System.Management.Automation.Language.Parser]::ParseFile($pubPath, [ref]$null, [ref]$null)
foreach ($f in $pubAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $false)) { Invoke-Expression $f.Extent.Text }
$ReadmeStartMarker = '<!-- telemetry:start -->'
$ReadmeEndMarker = '<!-- telemetry:end -->'
$DataColumns = @('WindowEnd_Local','Window','Endpoint','UrlHash','Polls','AvailabilityPercent','FailedPolls','SlowPolls','P50Ms','P95Ms','MaxMs','Outages','LongestOutageSeconds','TotalOutageSeconds','SlowPeriods','BackendAddresses','FailureReasons')
function Write-Line { param($Level, $Message) }

$W0 = [datetime]'2026-09-21T07:15:00'
$Wfrom = $W0.AddHours(-12)
function TRow($t, $ms, $code = '200', $ok = 'True', $reason = 'OK', $ip = '10.0.0.1') { [PSCustomObject]@{ Timestamp_Local = $t.ToString('yyyy-MM-dd HH:mm:ss'); HttpCode = $code; ContentOk = $ok; TotalMs = "$ms"; Reason = $reason; RemoteIp = $ip } }
$rows = @()
foreach ($i in 1..10) { $rows += TRow $W0.AddMinutes(-$i) (100 * $i) }
$rows += TRow $W0.AddMinutes(-11) '' '000' 'False' 'Timed out' ''
$rows += TRow $W0.AddMinutes(-12) '' '503' 'False' 'HTTP 503' '10.0.0.2'
$rows += TRow $W0.AddMinutes(-13) 4200
$outs = @([PSCustomObject]@{ OutageStart_Local = $W0.AddMinutes(-12).ToString('yyyy-MM-dd HH:mm:ss'); DurationSeconds = '95' }, [PSCustomObject]@{ OutageStart_Local = $Wfrom.AddHours(-3).ToString('yyyy-MM-dd HH:mm:ss'); DurationSeconds = '900' })
$drops = @("$($W0.AddMinutes(-13).ToString('yyyy-MM-dd HH:mm:ss')) | SLOWSTART | Declared SLOW", "$($Wfrom.AddHours(-2).ToString('yyyy-MM-dd HH:mm:ss')) | SLOWSTART | older", "garbage")
$st = Get-TelemetryStat -Rows $rows -Outages $outs -DropLines $drops -From $Wfrom -To $W0 -SlowThresholdMs 3000
Check "M2 statistics match the hand-computed window" ($st.Polls -eq 13 -and $st.FailedPolls -eq 2 -and $st.AvailabilityPercent -eq 84.62 -and $st.SlowPolls -eq 1 -and $st.MaxMs -eq 4200 -and $st.P50Ms -eq 600 -and $st.BackendAddresses -eq 1 -and $st.Outages -eq 1 -and $st.LongestOutageSeconds -eq 95 -and $st.TotalOutageSeconds -eq 95 -and $st.SlowPeriods -eq 1)
Check "M2 failure reasons are grouped and ordered" ($st.FailureReasons -match '^(Timed out x1; HTTP 503 x1|HTTP 503 x1; Timed out x1)$')
$empty = Get-TelemetryStat -Rows @() -Outages @() -DropLines @() -From $Wfrom -To $W0
Check "M2 an empty window reports zero polls and no percentiles rather than throwing" ($empty.Polls -eq 0 -and $empty.AvailabilityPercent -eq 0 -and $null -eq $empty.P95Ms -and $empty.FailureReasons -eq '')
$allBad = Get-TelemetryStat -Rows @((TRow $W0.AddMinutes(-1) '' '000' 'False' 'Timed out' ''), (TRow $W0.AddMinutes(-2) '' '000' 'False' 'Timed out' '')) -Outages @() -DropLines @() -From $Wfrom -To $W0
Check "M2 a window with no successful poll reports 0 percent and null percentiles" ($allBad.Polls -eq 2 -and $allBad.AvailabilityPercent -eq 0 -and $null -eq $allBad.P50Ms -and $null -eq $allBad.MaxMs -and $allBad.FailureReasons -eq 'Timed out x2')
Check "M2 percentile edges" ((Get-Percentile @(1) 95) -eq 1 -and (Get-Percentile @(1,2) 50) -eq 1 -and (Get-Percentile @(1,2,3,4) 95) -eq 4 -and $null -eq (Get-Percentile @() 50))

# Window reading across a month rollover, and rotated drops logs
$mDir = Join-Path $tRoot 'monitor-a'
New-Item $mDir -ItemType Directory | Out-Null
$hdr = '"Timestamp_Local","HttpCode","ContentOk","TotalMs","Reason","RemoteIp"'
$aug = @($hdr, '"2026-08-31 23:59:00","200","True","120","OK","10.0.0.1"', '"2026-08-31 20:00:00","200","True","120","OK","10.0.0.1"')
$sep = @($hdr, '"2026-09-01 00:01:00","200","True","140","OK","10.0.0.1"')
[IO.File]::WriteAllLines((Join-Path $mDir 'Latency_202608.csv'), [string[]]$aug)
[IO.File]::WriteAllLines((Join-Path $mDir 'Latency_202609.csv'), [string[]]$sep)
$roll = Get-TelemetryRow -Folder $mDir -From ([datetime]'2026-08-31T23:00:00') -To ([datetime]'2026-09-01T01:00:00')
Check "M3 a window spanning a month reads both monthly files and honours its edges" (@($roll).Count -eq 2 -and @($roll | Where-Object { $_.Timestamp_Local -eq '2026-08-31 20:00:00' }).Count -eq 0)
Set-Content (Join-Path $mDir 'Drops.log') -Value "2026-09-21 07:00:00 | SLOWSTART | current"
Set-Content (Join-Path $mDir 'Drops_20260920_120000.log') -Value "2026-09-21 02:00:00 | SLOWSTART | rotated"
$dl = Get-DropLine -Folder $mDir -From ([datetime]'2026-09-21T00:00:00')
Check "M3 rotated drops logs inside the window are read as well as the current one" (@($dl | Where-Object { $_ -match 'rotated' }).Count -eq 1 -and @($dl | Where-Object { $_ -match 'current' }).Count -eq 1)

# Codes, README markers, commit text, report
$map = @{}
$h1 = ConvertTo-UrlHash -Url 'https://example.com/a'
$h2 = ConvertTo-UrlHash -Url 'https://example.com/b'
$c1 = Get-EndpointCode -Map $map -UrlHash $h1
$c2 = Get-EndpointCode -Map $map -UrlHash $h2
Check "M4 codes are assigned in order, stable per URL, and the hash is 12 hex characters" ($c1 -eq 'ENDPOINT-01' -and $c2 -eq 'ENDPOINT-02' -and (Get-EndpointCode -Map $map -UrlHash $h1) -eq 'ENDPOINT-01' -and $h1 -match '^[0-9a-f]{12}$' -and $h1 -ne $h2 -and $h1 -eq (ConvertTo-UrlHash -Url 'https://example.com/a '))
Check "M4 no URL, site, or monitor name reaches a published code" ($c1 -notmatch 'example' -and $h1 -notmatch 'example')
$readme = "intro`n$ReadmeStartMarker`nold table`n$ReadmeEndMarker`ntail"
$newReadme = Update-ReadmeTable -Text $readme -Table "| a |`n| b |"
Check "M4 the README table is replaced between the markers and nothing else moves" ($newReadme -match '(?s)intro.*telemetry:start.*\| a \|.*\| b \|.*telemetry:end.*tail' -and $newReadme -notmatch 'old table' -and (Update-ReadmeTable -Text 'no markers' -Table 'x') -eq 'no markers')
$subject = New-CommitSubject -Date ([datetime]'2026-09-21') -Window 'Morning' -Availability 99.4
Check "M4 commit subject reads as a telemetry run" ($subject -eq 'Feature Improvement: telemetry reporting publisher, data sampling 2026-09-21 morning, availability 99.40%' -and (New-CommitSubject -Date ([datetime]'2026-09-21') -Window 'Evening' -Availability 100) -eq 'Feature Improvement: telemetry reporting publisher, data sampling 2026-09-21 evening, availability 100.00%')
$body = New-CommitBody -Measurements @([PSCustomObject]@{ Code = 'ENDPOINT-01'; Stats = $st })
Check "M4 commit body names each endpoint with its p95 and outage" ($body -match 'ENDPOINT-01: 13 polls, 84\.62% available, p95 \d+ ms, 1 outage totalling 95s')
$reportRows = @([PSCustomObject]@{ Window='Morning'; WindowEnd_Local='2026-09-21 07:15:00'; Endpoint='ENDPOINT-01'; Polls='120'; AvailabilityPercent='99.17'; P50Ms='210'; P95Ms='480'; MaxMs='900'; FailedPolls='1'; SlowPolls='0'; Outages='0'; TotalOutageSeconds='0'; LongestOutageSeconds='0'; SlowPeriods='0'; FailureReasons='Timed out x1' })
$report = New-TelemetryReport -Date ([datetime]'2026-09-21') -Rows $reportRows
Check "M4 the report carries the date, the window, the table, and the failure note" ($report -match '# Endpoint availability, 2026-09-21' -and $report -match '## Morning window, measured to 2026-09-21 07:15:00' -and $report -match '\| ENDPOINT-01 \| 120 \| 99\.17% \| 210 ms \| 480 ms \| 900 ms \| 1 \| 0 \| 0 \|' -and $report -match 'ENDPOINT-01 failures: Timed out x1\.')

# Full publish against a local bare repository
$savedEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'   # git writes ordinary progress to stderr, which would otherwise end the run
$bare = Join-Path $tRoot 'origin.git'
$work = Join-Path $tRoot 'work'
$state = Join-Path $tRoot 'state'
$mroot = Join-Path $tRoot 'monitors'
New-Item (Join-Path $mroot 'site-a') -ItemType Directory -Force | Out-Null
$now = Get-Date
Set-Content (Join-Path $mroot 'site-a\install-settings.json') -Value (@{ MonitorName='Site A'; SiteName='S'; Url='https://example.com/a' } | ConvertTo-Json)
Copy-Item (Join-Path $mDir 'Latency_202609.csv') (Join-Path $mroot "site-a\Latency_$($now.ToString('yyyyMM')).csv")
$recent = @($hdr) + @(1..40 | ForEach-Object { '"' + $now.AddMinutes(-$_).ToString('yyyy-MM-dd HH:mm:ss') + '","200","True","' + (100 + $_) + '","OK","10.0.0.1"' })
[IO.File]::WriteAllLines((Join-Path $mroot "site-a\Latency_$($now.ToString('yyyyMM')).csv"), [string[]]$recent)
Set-Content (Join-Path $mroot 'site-a\Watch-CurlMonitor.ps1') -Value "`$SlowThresholdMs       = 3000`nfunction x { }"
& git.exe init --bare -q $bare
& git.exe clone -q $bare $work 2>&1 | Out-Null
Set-Content (Join-Path $work 'README.md') -Value "# repo`n`n$ReadmeStartMarker`n$ReadmeEndMarker`n"
& git.exe -C $work add . | Out-Null
& git.exe -C $work -c user.name=T -c user.email=t@example.com commit -qm 'init' | Out-Null
& git.exe -C $work push -q origin HEAD 2>&1 | Out-Null
& git.exe -C $work config user.name 'Test Author' | Out-Null
& git.exe -C $work config user.email 'test@example.com' | Out-Null
$before = [int](& git.exe -C $work rev-list --count HEAD)
$rc1 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubPath -RepoPath $work -MonitorRoot $mroot -StatePath $state -Window Morning 2>&1
$after = [int](& git.exe -C $work rev-list --count HEAD)
$dataFile = Join-Path $work "data\telemetry\ENDPOINT-01\$($now.ToString('yyyy-MM')).csv"
Check "M5 a publish commits one change with the telemetry subject and pushes it" ($after -eq $before + 1 -and (& git.exe -C $work log -1 --pretty=%s) -match '^Feature Improvement: telemetry reporting publisher, data sampling \d{4}-\d\d-\d\d morning, availability ' -and (& git.exe -C $bare rev-list --count HEAD) -eq "$after")
Check "M5 it writes the data row, the dated report, and the README table" ((Test-Path $dataFile) -and @(Import-Csv $dataFile).Count -eq 1 -and (Import-Csv $dataFile)[0].Polls -eq '40' -and (Test-Path (Join-Path $work "reports\$($now.ToString('yyyy-MM-dd')).md")) -and (Get-Content (Join-Path $work 'README.md') -Raw) -match 'ENDPOINT-01')
Check "M5 the map and state stay out of the repository" ((Test-Path (Join-Path $state 'endpoint-map.json')) -and (Test-Path (Join-Path $state 'publish-state.json')) -and -not (Test-Path (Join-Path $work 'endpoint-map.json')) -and -not ((& git.exe -C $work log -1 --name-only --pretty=format:) -match 'endpoint-map'))
Check "M5 the committed data names no URL, and the author is the configured one" (-not ((Get-Content $dataFile -Raw) -match 'example\.com') -and (& git.exe -C $work log -1 --pretty=%an) -eq 'Test Author')

# A second run in the same window appends rather than duplicating the file, and the report keeps both
$rc2 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubPath -RepoPath $work -MonitorRoot $mroot -StatePath $state -Window Evening 2>&1
Check "M6 the same endpoint keeps its code, so a second run appends instead of starting a new folder" (@(Get-ChildItem (Join-Path $work 'data\telemetry') -Directory).Count -eq 1 -and (Get-Content (Join-Path $state 'endpoint-map.json') -Raw) -match 'ENDPOINT-01')
Check "M6 a second window appends one row under the same header and commits again" (@(Import-Csv $dataFile).Count -eq 2 -and @(Get-Content $dataFile | Where-Object { $_ -match 'WindowEnd_Local' }).Count -eq 1 -and [int](& git.exe -C $work rev-list --count HEAD) -eq $after + 1)
Check "M6 the day's report holds both windows" ((Get-Content (Join-Path $work "reports\$($now.ToString('yyyy-MM-dd')).md") -Raw) -match '(?s)## Morning window.*## Evening window')

# Nothing to publish makes no commit
$emptyRoot = Join-Path $tRoot 'no-monitors'
New-Item $emptyRoot -ItemType Directory | Out-Null
$countBefore = [int](& git.exe -C $work rev-list --count HEAD)
$rc3 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubPath -RepoPath $work -MonitorRoot $emptyRoot -StatePath $state 2>&1
Check "M7 no monitor data makes no commit and reports it" ([int](& git.exe -C $work rev-list --count HEAD) -eq $countBefore -and ($rc3 -join ' ') -match 'Nothing published')

# A push rejected by someone else's commit is recovered by the rebase
$other = Join-Path $tRoot 'other'
& git.exe clone -q $bare $other 2>&1 | Out-Null
Set-Content (Join-Path $other 'notes.md') -Value 'from elsewhere'
& git.exe -C $other add . | Out-Null
& git.exe -C $other -c user.name=O -c user.email=o@example.com commit -qm 'unrelated change' | Out-Null
& git.exe -C $other push -q origin HEAD 2>&1 | Out-Null
$rc4 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubPath -RepoPath $work -MonitorRoot $mroot -StatePath $state -Window Morning 2>&1
Check "M8 a non-fast-forward push is rebased and lands, keeping the other commit" ((& git.exe -C $bare log --pretty=%s) -match 'Feature Improvement: telemetry reporting publisher' -and (& git.exe -C $bare log --pretty=%s) -match 'unrelated change' -and ($rc4 -join ' ') -match 'Published and pushed|telemetry: ')

# Dry run writes files and leaves git alone
$dryWork = Join-Path $tRoot 'dry'
& git.exe clone -q $bare $dryWork 2>&1 | Out-Null
$dryBefore = [int](& git.exe -C $dryWork rev-list --count HEAD)
$rc5 = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubPath -RepoPath $dryWork -MonitorRoot $mroot -StatePath (Join-Path $tRoot 'state2') -Window Morning -DryRun 2>&1
Check "M9 a dry run writes the files but commits nothing" ([int](& git.exe -C $dryWork rev-list --count HEAD) -eq $dryBefore -and (Test-Path (Join-Path $dryWork 'data\telemetry')) -and ($rc5 -join ' ') -match 'Dry run' -and (& git.exe -C $dryWork status --porcelain) -match 'data/telemetry')

# The publisher installer, driven into scratch paths, never reaching GitHub
$pubInstaller = 'C:\Workspaces\HST Monitor\Install-TelemetryPublisher.ps1'
$instState = Join-Path $tRoot 'installer-state'
$taskPathTest = '\CurlMonitorTest\'
$taskNameTest = 'Curl Monitor telemetry publisher test'
$instOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pubInstaller -NonInteractive -SkipClone -StateDir $instState -RepoPath $work -MonitorRoot $mroot -TaskPath $taskPathTest -TaskName $taskNameTest -RepoUrl 'https://github.com/owner/repo' -AuthorName 'Test Author' -AuthorEmail 'test@example.com' -Token 'ghp_testtoken_value_0123456789' 2>&1
$tokenFile = Join-Path $instState 'github-token.bin'
$instTask = Get-ScheduledTask -TaskPath $taskPathTest -TaskName $taskNameTest -ErrorAction SilentlyContinue
$acl = if (Test-Path $tokenFile) { @((Get-Acl $tokenFile).Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } | Sort-Object -Unique) } else { @() }
foreach ($f in $pubAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $false)) { }
$instAst = [System.Management.Automation.Language.Parser]::ParseFile($pubInstaller, [ref]$null, [ref]$null)
foreach ($f in $instAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Unprotect-Token','New-AuthenticatedUrl','Test-RepoUrl','Test-TimeOfDay','Test-EmailLike')}, $false)) { Invoke-Expression $f.Extent.Text }
$Entropy = 'DIT-CurlMonitor-Telemetry-v1'
Check "M10 the token is stored encrypted, readable back, and locked to SYSTEM and Administrators" ((Test-Path $tokenFile) -and (Unprotect-Token -Path $tokenFile) -eq 'ghp_testtoken_value_0123456789' -and -not ((Get-Content $tokenFile -Raw) -match 'ghp_testtoken') -and (($acl -join ',') -eq 'S-1-5-18,S-1-5-32-544'))
Check "M10 the task runs as SYSTEM twice a day and catches up a missed run" ($instTask -and @($instTask.Triggers).Count -eq 2 -and "$($instTask.Principal.UserId)" -match 'SYSTEM' -and $instTask.Settings.StartWhenAvailable -and "$(@($instTask.Actions)[0].Arguments)" -match 'Publish-UptimeTelemetry\.ps1')
Check "M10 the token reaches neither the console, the task definition, nor the settings file" (-not (($instOut -join ' ') -match 'ghp_testtoken') -and -not ("$(@($instTask.Actions)[0].Arguments)" -match 'ghp_testtoken') -and -not ((Get-Content (Join-Path $instState 'publisher-settings.json') -Raw) -match 'ghp_testtoken'))
Check "M10 input validation refuses a bad repository URL, time, or email" ((Test-RepoUrl 'https://github.com/owner/repo') -and -not (Test-RepoUrl 'git@github.com:owner/repo.git') -and -not (Test-RepoUrl 'ftp://x/y/z') -and (Test-TimeOfDay '07:15') -and -not (Test-TimeOfDay '25:00') -and (Test-EmailLike 'a@b.co') -and -not (Test-EmailLike 'nope'))
Check "M10 the authenticated URL is only ever built in memory, never stored in the clone" ((New-AuthenticatedUrl -RepoUrl 'https://github.com/o/r' -Token 'T') -eq 'https://x-access-token:T@github.com/o/r' -and -not ((& git.exe -C $work config --get remote.origin.url) -match 'x-access-token'))
if ($instTask) { Unregister-ScheduledTask -TaskPath $taskPathTest -TaskName $taskNameTest -Confirm:$false -ErrorAction SilentlyContinue }
try { $svc = New-Object -ComObject Schedule.Service; $svc.Connect(); $svc.GetFolder('\').DeleteFolder('CurlMonitorTest', 0) } catch { }
Check "M10 the test task is gone again" ($null -eq (Get-ScheduledTask -TaskPath $taskPathTest -TaskName $taskNameTest -ErrorAction SilentlyContinue))
$ErrorActionPreference = $savedEap
Remove-Item $tRoot -Recurse -Force -ErrorAction SilentlyContinue


Section "U. Uninstall: listing, selection, removal, and the broken shapes"
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Get-FolderSizeText','Get-RemovableMonitor','Show-RemovableMonitor','Select-RemovableMonitor','Stop-MonitorProcess','Move-MonitorHistory','Remove-MonitorInstall','Invoke-UninstallFlow')},$false)) { Invoke-Expression $f.Extent.Text }
$uRoot = Join-Path $InstallDir 'uninstall'
$uKeep = Join-Path $InstallDir 'uninstall-kept'
$uTaskPath = '\NoSuchTaskPath\'
$script:ULogs = @()
function Write-Log { param($Level,$Message) $script:ULogs += "$Level|$Message" }
$script:UOut = @()
function Write-Host { param([Parameter(Position=0,ValueFromRemainingArguments=$true)]$Text,$ForegroundColor) $script:UOut += (@($Text) -join ' ') }
$script:UQ = [System.Collections.Queue]::new()
function UAnswers($a) { $script:UQ.Clear(); foreach ($x in $a) { $script:UQ.Enqueue($x) } }
function Read-Setting { param([string]$Prompt,[string]$Default="") if ($script:UQ.Count -eq 0) { throw "unexpected prompt: $Prompt" }; $v = $script:UQ.Dequeue(); if ([string]::IsNullOrWhiteSpace($v)) { return $Default }; return $v }
function Read-Choice { param([string]$Prompt,[string[]]$Allowed,[string]$Default) if ($script:UQ.Count -eq 0) { throw "unexpected choice: $Prompt" }; while ($true) { $v = $script:UQ.Dequeue(); if ([string]::IsNullOrWhiteSpace($v)) { return $Default }; $v = $v.Trim().ToUpper(); if ($Allowed -contains $v) { return $v }; if ($script:UQ.Count -eq 0) { return $Default } } }
function UMonitor { param([string]$Name,[string]$Root,[string]$Url='https://example.invalid/p')
  $dir = Join-Path $Root (ConvertTo-MonitorSlug $Name)
  New-Item $dir -ItemType Directory -Force | Out-Null
  @{ MonitorName=$Name; Url=$Url; SiteName='S' } | ConvertTo-Json | Set-Content (Join-Path $dir 'install-settings.json') -Encoding UTF8
  foreach ($f in @('Watch-CurlMonitor.ps1','credential.bin','Outages.csv','Drops.log','Transcript_20260923.log',("Latency_" + (Get-Date -Format 'yyyyMM') + ".csv"))) { Set-Content (Join-Path $dir $f) -Value 'x' -Encoding UTF8 }
  return $dir }
function UReset { param([string]$Dir) if (Test-Path $Dir) { Get-ChildItem $Dir -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object { try { $_.IsReadOnly = $false } catch { } }; Remove-Item $Dir -Recurse -Force -ErrorAction SilentlyContinue }; New-Item $Dir -ItemType Directory -Force | Out-Null }
$NonInteractive = $false; $MonitorNameOverride = ''

UReset $uRoot
$null = UMonitor -Name 'HST eChart' -Root $uRoot
$null = UMonitor -Name 'Patient Portal' -Root $uRoot -Url 'https://portal.invalid/'
New-Item (Join-Path $uRoot 'half-removed') -ItemType Directory -Force | Out-Null
Set-Content (Join-Path $uRoot 'half-removed\Drops.log') -Value 'x' -Encoding UTF8
$uItems = @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath -LegacyDir (Join-Path $uRoot 'none') -LegacyTask 'No Legacy Task')
Check "U1 listing covers installed monitors and a folder left by a part-finished removal" (@($uItems).Count -eq 3 -and @($uItems | Where-Object { $_.Kind -eq 'Leftover' -and $_.Name -eq 'half-removed' }).Count -eq 1 -and @($uItems | Where-Object { $_.Url -eq 'https://portal.invalid/' }).Count -eq 1 -and @($uItems | Where-Object { $_.TaskState -eq 'no task' }).Count -eq 3)
Check "U1 folder size reads as a size, not a byte dump" ((Get-FolderSizeText (@($uItems)[0].Dir)) -match '(bytes|KB|MB)$' -and (Get-FolderSizeText (Join-Path $uRoot 'gone')) -eq 'no folder')
Check "U2 a name resolves exactly, by case, and by folder spelling" ((Select-RemovableMonitor -Items $uItems -Requested 'HST eChart').Name -eq 'HST eChart' -and (Select-RemovableMonitor -Items $uItems -Requested 'hst echart').Name -eq 'HST eChart' -and (Select-RemovableMonitor -Items $uItems -Requested 'HST_eChart').Name -eq 'HST eChart')
$script:ULogs = @()
Check "U2 an unknown or empty name is refused and names what is installed" ($null -eq (Select-RemovableMonitor -Items $uItems -Requested 'Nope') -and $null -eq (Select-RemovableMonitor -Items $uItems -Requested '  ') -and ($script:ULogs -join "`n") -match "No monitor called 'Nope'\. Installed: ")
UAnswers @('X')
Check "U3 X cancels and removes nothing" ((Invoke-UninstallFlow -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 0 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath).Count -eq 3)
UAnswers @('', '9', 'nope', '1', '')
$script:ULogs = @()
Check "U3 blank, out of range and junk are refused, and the confirmation defaults to no" ((Invoke-UninstallFlow -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 0 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath).Count -eq 3 -and ($script:ULogs -join "`n") -match 'Pick a number between 1 and 3' -and ($script:ULogs -join "`n") -match 'Cancelled\. Nothing was removed')
UAnswers @('hst echart', 'Y', 'Y')
$rcU = Invoke-UninstallFlow -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep
$uLeft = @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath)
$uKept = @(Get-ChildItem $uKeep -Directory -ErrorAction SilentlyContinue)
Check "U4 picking by name removes that monitor only, and keeps its history" ($rcU -eq 0 -and @($uLeft).Count -eq 2 -and -not (Test-Path (Join-Path $uRoot 'HST-eChart')) -and (Test-Path (Join-Path $uRoot 'Patient-Portal\credential.bin')) -and @($uKept).Count -eq 1 -and @(Get-ChildItem $uKept[0].FullName -File).Count -eq 3)
UAnswers @('Patient Portal', 'Y', 'N')
$script:ULogs = @()
$rcU = Invoke-UninstallFlow -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep
Check "U4 deleting the history says it cannot be undone and keeps nothing" ($rcU -eq 0 -and @(Get-ChildItem $uKeep -Directory).Count -eq 1 -and ($script:ULogs -join "`n") -match 'cannot be undone')
Check "U4 the summary says what went and what is still installed" ((($script:UOut -join "`n") -match '(?s)Removal summary.*Folder\s+: .*deleted.*Still installed : half-removed'))

UReset $uRoot
$uLegacy = Join-Path $uRoot 'HSTProbe'
New-Item $uLegacy -ItemType Directory -Force | Out-Null
@{ Url='https://legacy.invalid/x'; SiteName='Old' } | ConvertTo-Json | Set-Content (Join-Path $uLegacy 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $uLegacy 'HST-eChart-Latency_202609.csv') -Value 'a,b' -Encoding UTF8
$uLeg = @(Get-RemovableMonitor -Root (Join-Path $uRoot 'empty') -Path $uTaskPath -LegacyDir $uLegacy -LegacyTask 'No Legacy Task')
$uLegRes = Remove-MonitorInstall -Monitor @($uLeg)[0] -KeepHistory $true -KeepRoot $uKeep
Check "U5 the older layout is listed with its URL, removed, and its HST-named history kept" (@($uLeg).Count -eq 1 -and @($uLeg)[0].Kind -eq 'Legacy' -and @($uLeg)[0].Url -eq 'https://legacy.invalid/x' -and $uLegRes.FolderRemoved -and -not (Test-Path $uLegacy) -and @(Get-ChildItem $uLegRes.HistoryPath -Filter 'HST-eChart-*').Count -eq 1)

UReset $uRoot
$uLock = UMonitor -Name 'Locked' -Root $uRoot
$uHandle = [System.IO.File]::Open((Join-Path $uLock 'Transcript_20260923.log'), 'Open', 'Read', 'None')
$uLockRes = Remove-MonitorInstall -Monitor @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath)[0] -KeepHistory $true -KeepRoot $uKeep
$uHandle.Close()
Check "U6 a file held open keeps the folder and names the failure, history still saved" (-not $uLockRes.FolderRemoved -and @($uLockRes.Problems).Count -eq 1 -and @($uLockRes.Problems)[0] -match 'could not be deleted' -and @(Get-ChildItem $uLockRes.HistoryPath -File).Count -eq 3)
$uLockRes2 = Remove-MonitorInstall -Monitor @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath)[0] -KeepHistory $true -KeepRoot $uKeep
Check "U6 running it again once the lock is gone finishes the removal" ($uLockRes2.FolderRemoved -and @($uLockRes2.Problems).Count -eq 0 -and -not (Test-Path $uLock))

UReset $uRoot
$null = UMonitor -Name 'Alpha' -Root $uRoot
$null = UMonitor -Name 'Beta' -Root $uRoot
$NonInteractive = $true
$script:ULogs = @()
Check "U7 non-interactive with several monitors and no name refuses and removes nothing" ((Invoke-UninstallFlow -Requested '' -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 1 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath).Count -eq 2 -and ($script:ULogs -join "`n") -match 'Set \$MonitorNameOverride')
Check "U7 non-interactive with an unknown name refuses and removes nothing" ((Invoke-UninstallFlow -Requested 'Gamma' -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 1 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath).Count -eq 2)
Check "U7 non-interactive with a name removes that one, then the last needs no name" ((Invoke-UninstallFlow -Requested 'Alpha' -KeepHistory $false -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 0 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath)[0].Name -eq 'Beta' -and (Invoke-UninstallFlow -Requested '' -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 0 -and @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath).Count -eq 0)
$NonInteractive = $false
Check "U7 a run with nothing installed says so and changes nothing" ((Invoke-UninstallFlow -Root $uRoot -Path $uTaskPath -KeepRoot $uKeep) -eq 0)

UReset $uRoot
$uEvilDir = Join-Path $uRoot 'evil'
New-Item $uEvilDir -ItemType Directory -Force | Out-Null
$uEvilName = '$(New-Item -Path C:\tmp\uninstall-pwned.txt -ItemType File -Force)`"; whoami #'
@{ MonitorName=$uEvilName; Url='https://x.invalid/' } | ConvertTo-Json | Set-Content (Join-Path $uEvilDir 'install-settings.json') -Encoding UTF8
Set-Content (Join-Path $uEvilDir 'Drops.log') -Value 'x' -Encoding UTF8
$uEvil = @(Get-RemovableMonitor -Root $uRoot -Path $uTaskPath)
$script:UOut = @()
Show-RemovableMonitor -Items $uEvil
$uEvilRes = Remove-MonitorInstall -Monitor @($uEvil)[0] -KeepHistory $true -KeepRoot $uKeep
Check "U8 a hostile monitor name is printed literally and runs nothing" ((($script:UOut -join "`n") -match [regex]::Escape('New-Item -Path C:\tmp\uninstall-pwned.txt')) -and -not (Test-Path 'C:\tmp\uninstall-pwned.txt') -and $uEvilRes.FolderRemoved -and (Test-Path $uRoot))
$uTraversal = [PSCustomObject]@{ Name='T'; Slug='..\..\Windows'; Dir=(Join-Path $uRoot 'safe'); Url=''; Kind='Monitor'; TaskName='none'; TaskPath=$uTaskPath; TaskState='no task' }
New-Item (Join-Path $uRoot 'safe') -ItemType Directory -Force | Out-Null
Set-Content (Join-Path $uRoot 'safe\Drops.log') -Value 'x' -Encoding UTF8
$uTrav = Move-MonitorHistory -Monitor $uTraversal -KeepRoot $uKeep
Check "U8 a folder name that climbs out cannot write outside the keep root" ($uTrav -and $uTrav.StartsWith($uKeep, [StringComparison]::OrdinalIgnoreCase))
Check "U9 the first question routes to the uninstall and exits before the install root is touched" ($src -match "Install or upgrade a monitor, or remove one\? I = install or upgrade, U = uninstall" -and $src -match "if \(""\`$Action"" -eq 'Uninstall'\) \{ exit \(Invoke-UninstallFlow\) \}" -and $src.IndexOf("exit (Invoke-UninstallFlow)") -lt $src.IndexOf("Created install root"))
Check "U9 the config block carries the action and the history toggle" ($src -match '(?m)^\$Action\s+=\s+"Install"' -and $src -match '(?m)^\$KeepHistoryOnUninstall = \$true' -and $src -match '(?m)^\$HistoryKeepRoot\s+=')
Check "U9 the last monitor going names the telemetry task rather than removing it" ($src -match 'telemetry publisher task .* is still scheduled' -and $src -match 'Unregister-ScheduledTask -TaskPath')
Remove-Item function:Write-Host -ErrorAction SilentlyContinue
Remove-Item $uRoot -Recurse -Force -ErrorAction SilentlyContinue
function Write-Log { param($Level,$Message) }


Write-Host ""
Write-Host "TOTAL: $script:pass passed, $script:fail failed"
if ($script:fail){ Write-Host "Failed: $($script:failed -join '; ')"; exit 1 }
