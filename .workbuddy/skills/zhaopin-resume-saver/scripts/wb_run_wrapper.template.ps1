$ErrorActionPreference = 'Continue'
# WebBridge resume-download wrapper with AUTO-RESTART loop (#62) + FAST-START (#63, 2026-09-16)
# Reads run.ps1 path + log path from %TEMP%\wb_wrapper_paths.txt (line1=line2=).
# Round loop: probe extension -> (rebuild daemon only if needed) -> run run.ps1 -> check exit code.
#   exit 0 (DONE) / 1 (fatal config/job) / 2 (pool exhausted) -> stop
#   exit 3 (env fatal) / 4 (stall watchdog) -> restart daemon and re-run (resume)
# #63 fast-start changes (root cause: old wrapper ALWAYS rebuilt the daemon and slept 5s
#   BEFORE the first probe + 3s after ok => >=10s fixed overhead even when everything was
#   healthy; a daemon rebuild also forces the Chrome extension to reconnect, worst case 200s):
#   1) Round 1: quick probe (3 x 800ms) first - if the extension is already alive, skip
#      stop/start entirely and go straight to run.ps1. Rounds >= 2 (after exit 3/4) always
#      rebuild the daemon, preserving the wedge-recovery behavior.
#   2) Probe waits use adaptive intervals (0.5s -> 1s -> 2s -> 5s, ceiling ~153s) so a
#      fast-ready daemon is detected in <1s instead of always paying a 5s pre-sleep.
#   3) Removed the unconditional 3s sleep after probe ok (run.ps1 env-check re-probes anyway).
#   4) Inter-round pause reduced 5s -> 2s.
# Hard cap: 8 rounds. run.ps1 auto-computes remaining target from existing files.
$maxRounds = 8
$probeReq  = "$env:TEMP\wb-probe.json"
$probeRes  = "$env:TEMP\wb-probe-res.json"
$daemonExe = "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe"
$env:NO_PROXY = "127.0.0.1,localhost"
$env:no_PROXY = "127.0.0.1,localhost"
[System.IO.File]::WriteAllText($probeReq, '{"action":"list_tabs","args":{},"session":"probe"}', [System.Text.UTF8Encoding]::new($false))
function Probe-Ext {
    Remove-Item $probeRes -Force -ErrorAction SilentlyContinue
    curl.exe -sS --noproxy "*" -m 8 -X POST http://127.0.0.1:10086/command -H "Content-Type: application/json" --data-binary "@$probeReq" --output $probeRes 2>$null
    if (Test-Path $probeRes) {
        $c = [System.IO.File]::ReadAllText($probeRes)
        if ($c -match '"ok":true') { return $true }
    }
    return $false
}
# adaptive-interval readiness poll (true probe via list_tabs, same rationale as #55)
function Wait-ExtReady {
    $sched = @( @(6,500), @(10,1000), @(10,2000), @(24,5000) )
    foreach ($s in $sched) {
        for ($i = 0; $i -lt $s[0]; $i++) {
            if (Probe-Ext) { return $true }
            Start-Sleep -Milliseconds $s[1]
        }
    }
    return $false
}
# --- read paths table (line1=run.ps1, line2=log) ---
if (-not (Test-Path "$env:TEMP\wb_wrapper_paths.txt")) {
    [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_fail.txt", "PATHS_FILE_MISSING", [System.Text.UTF8Encoding]::new($false))
    exit 1
}
$lines = [System.IO.File]::ReadAllLines("$env:TEMP\wb_wrapper_paths.txt", [System.Text.Encoding]::UTF8)
$runPs = $lines[0].Trim()
$logPath = $lines[1].Trim()
if (-not (Test-Path $runPs)) {
    [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_fail.txt", "RUNPS_MISSING_$runPs", [System.Text.UTF8Encoding]::new($false))
    exit 1
}
if (-not $logPath) { $logPath = "$env:TEMP\wb_run_dl.log" }
$round = 0
while ($round -lt $maxRounds) {
    $round++
    $ready = $false
    if ($round -eq 1) {
        # #63 fast path: extension already alive -> reuse running daemon, skip stop/start
        for ($i = 0; $i -lt 3; $i++) {
            if (Probe-Ext) { $ready = $true; break }
            Start-Sleep -Milliseconds 800
        }
    }
    if (-not $ready) {
        & $daemonExe stop 2>$null
        Start-Sleep -Seconds 1
        & $daemonExe start 2>&1 | Out-String | Set-Content "$env:TEMP\wb-daemon-schtask.txt" -Encoding UTF8
        if (-not (Wait-ExtReady)) {
            [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_fail.txt", "EXT_NOT_READY_round_${round}", [System.Text.UTF8Encoding]::new($false))
            exit 1
        }
    }
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Add-Content -Path $logPath -Value "===== wrapper round $round start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====="
    & $runPs *>> $logPath
    $code = $LASTEXITCODE
    $status = "round=$round exit=$code at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_status.txt", $status, [System.Text.UTF8Encoding]::new($false))
    if ($code -eq 0 -or $code -eq 1 -or $code -eq 2) { break }
    Start-Sleep -Seconds 2
}
