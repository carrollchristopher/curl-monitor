# curl-monitor

Always-on availability monitor for any URL, installed on a server at the site that depends on it. It runs as a SYSTEM scheduled task, polls the URL through its redirects with curl, and emails when the site loses access, when it stays slow, when it comes back, and when the monitor itself was not running.

One server can run several monitors side by side, one per URL. Each gets its own folder, task, data files, and alert subjects, named after what you called it at install time, for example "HST eChart" or "Patient Portal".

## What it does

- Polls the URL every 10 seconds, follows redirects, and checks the response is HTTP 200, at least a minimum size, and contains the text you asked for.
- Declares an outage after 3 failed polls in a row, emails DOWN, STILL DOWN every 30 minutes, and RESOLVED with the outage length.
- Emails SLOW when at least half the polls in the last 5 minutes were slower than 3 seconds or failed, STILL SLOW every 30 minutes, and SLOW RESOLVED once 10 percent or fewer are. A single slow poll or timeout never emails.
- Sends a daily summary at 7 AM of slow polls, failed polls with their reasons, outages, the worst response, and the busiest hour. It is sent only when something went wrong.
- Writes a monthly latency CSV, an outages CSV, and `Drops.log` with only the failures, slow polls, outage transitions, and alert delivery outcomes. Healthy polls never appear in the drops log.
- Keeps a heartbeat so a restart after a reboot, a stopped task, or a killed process sends a MONITOR RESTARTED notice with the gap and the cause. An outage or slow period in progress at the last heartbeat is carried over.
- Queues alerts that cannot be sent and retries them every minute for an hour.
- Sends mail through Microsoft Graph with an app registration scoped to one shared mailbox (RBAC for Applications), or through Microsoft 365 direct send, an internal relay, or authenticated SMTP.

Alert subjects read `[DOWN] HST eChart at Capital City (SERVER01) - unreachable`, so one inbox rule can sort by monitor, site, or server.

## Install

Run `Install-CurlMonitor.ps1` from an elevated Windows PowerShell 5.1 console on the site server. It asks for:

1. **Monitor name**, for example `HST eChart`. It names the folder, the scheduled task, and every alert.
2. **URL** to watch, http or https.
3. **Site name**, which appears in every alert subject beside the monitor name.
4. **Text that must appear on the page**, or blank to check only the status code and the page size.
5. **Mail settings**, through the wizard.

The installer then writes the monitor, registers the task, sends a test email, and sends an install confirmation. Poll interval, timeouts, and alert thresholds are not prompted: edit the config block at the top of the installer before deploying.

Re-running upgrades a monitor in place, with every prompt prefilled from its last run, so pressing Enter through the whole run is enough. The monitor name prefills with the one installed most recently, so on a server running several monitors, type the name of the one you are upgrading. Type a different name to add a second monitor beside it. An older HST-only install under `C:\ProgramData\DIT\HSTProbe` is offered a migration on the first run: its settings and history move into the new layout, and its task and folder are removed.

The first site in a tenant can create the Graph app registration, the shared sender mailbox, and the Exchange send scope during the install. It prints the tenant ID, client ID, and client secret once. Later sites paste those three values.

Requirements: Windows Server or Windows 10/11 with Windows PowerShell 5.1, curl.exe (built in since Windows 10 1803), administrator rights for the install, and outbound HTTPS.

## Files

| File | Purpose |
|---|---|
| `Install-CurlMonitor.ps1` | The installer. Contains the monitor script as an embedded template. |
| `Test-CurlMonitorHealth.ps1` | Checks the installed monitors: task, process, heartbeat, recorded polls, logs, mail secret, and a live probe. `-OutageDrill` simulates an outage end to end. |
| `Publish-UptimeTelemetry.ps1` | Summarises what the monitors measured and commits the result to a git working copy. |
| `Install-TelemetryPublisher.ps1` | Sets up the twice-daily scheduled task that runs the publisher. |
| `Test-CurlMonitorMailSend.ps1` | One-shot diagnostic for the Graph mail path on an installed server. |
| `Stress-InstallCurlMonitor.ps1` | Static and functional checks for the installer and the generated monitor, including the wizard, naming, migration, and the sign-in token path against a local mock. |
| `Stress-WindowsCurlMonitor.ps1` | Live checks on Windows PowerShell 5.1: DPAPI, ACLs, task objects, and the generated monitor running against local listeners and a real endpoint. |
| `Test-CurlMonitorInstallLocal.ps1` | End-to-end test that runs the real installer twice from an elevated console and verifies the result. |
| `PSScriptAnalyzerSettings.psd1` | Analyzer settings with the deliberate rule exclusions. |

## Checking a server

Copy `Test-CurlMonitorHealth.ps1` to the server and run it from an elevated Windows PowerShell console.

```powershell
.\Test-CurlMonitorHealth.ps1
```

It changes nothing and ends with WORKING or NOT WORKING CORRECTLY. Every monitor under the root gets its own section covering the task and process, the heartbeat age, the last poll, response times and failures over the last hour and 24 hours, gaps in recording, warnings in the monitor's own log, the mail secret, and a probe of the URL from the server. Use `-Monitor "HST eChart"` to check one.

```powershell
.\Test-CurlMonitorHealth.ps1 -Monitor "HST eChart" -OutageDrill
```

The drill makes only that monitor's own probes fail for about a minute, using a curl settings file in the profile of the account the monitor runs as, scoped to the monitored host. Browsers and other programs keep working. It confirms the monitor declares DOWN, sends the DOWN email, records RESOLVED, and sends the RESOLVED email. The drill leaves one short outage in the outage log. If the window is closed during the drill, a one-time cleanup task removes the file, and running the script again removes it immediately.

## Telemetry

The monitors measure a public endpoint continuously. `Publish-UptimeTelemetry.ps1` turns those measurements into a
published record twice a day, and `Install-TelemetryPublisher.ps1` schedules it.

```powershell
.\Install-TelemetryPublisher.ps1
```

It asks for the repository, the working copy, the two run times, and the commit identity, takes a personal access
token at a hidden prompt, and stores it encrypted with machine-scope DPAPI in a file only SYSTEM and Administrators
can read. The token never reaches the console, the task definition, or the git remote.

Each run publishes three things into the working copy:

- `data/telemetry/<code>/<yyyy-MM>.csv`, one appended row per run: polls, availability, failed and slow polls,
  p50, p95 and maximum response time, outages and their length, slow periods, and how many backend addresses
  answered.
- `reports/<yyyy-MM-dd>.md`, rewritten each run so the day holds both windows.
- A table in this README between the telemetry markers, covering the last 24 hours.

Endpoints are published as a code, `ENDPOINT-01` upward, alongside the first 12 characters of the SHA-256 of the
URL. The map from hash to code lives beside the publish state on the machine that runs the job and is never
committed, so the data shows what an endpoint did without naming it.

A commit reads like its contents:

```
Feature Improvement: telemetry reporting publisher, data sampling 2026-09-21 morning, availability 99.86%

ENDPOINT-01: 2519 polls, 99.86% available, p95 412 ms, 1 outage totalling 35s
```

Install a monitor with `$AlertsEnabled = $false` in the config block for a telemetry-only collector: no mail
wizard, no alerts, no stored credential, and every poll, outage and slow period still recorded.

<!-- telemetry:start -->
Last 24 hours, measured to 2026-09-22 08:49 local.

| Endpoint | Polls | Available | p50 | p95 | Outages |
|---|---:|---:|---:|---:|---:|
| ENDPOINT-01 | 7826 | 100% | 257 ms | 586 ms | 0 |
<!-- telemetry:end -->

## Testing

`Stress-InstallCurlMonitor.ps1` expects a copy of the installer at `C:\tmp\Install-CurlMonitor.ps1` and a copy under `C:\mnt\user-data\outputs\`. It runs under pwsh 7 or Windows PowerShell 5.1. `Stress-WindowsCurlMonitor.ps1` runs under Windows PowerShell 5.1. Run the live harness on its own; its timing checks are sensitive to CPU load.

```powershell
Invoke-ScriptAnalyzer -Path .\Install-CurlMonitor.ps1 -Settings .\PSScriptAnalyzerSettings.psd1
```

## Data written on the server

Each monitor keeps its own folder under `C:\ProgramData\DIT\CurlMonitor`, named after the monitor, locked to SYSTEM and Administrators with read access for Users:

- `Watch-CurlMonitor.ps1`, the generated monitor
- `install-settings.json` and `credential.bin` (DPAPI, machine scope)
- `Latency_yyyyMM.csv`, `Outages.csv`, `Drops.log`
- `Transcript_yyyyMMdd.log` daily transcripts, pruned after 30 days
- `heartbeat.json` and `summary-sent.txt`

The scheduled tasks live under `\DIT\`, named `Curl Monitor - <monitor name>`.
