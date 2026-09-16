$ErrorActionPreference = 'Continue'
# WebBridge resume-download wrapper with AUTO-RESTART loop (#62, 2026-09-16)
# Reads run.ps1 path + log path from %TEMP%\wb_wrapper_paths.txt (line1=line2=).
# Round loop: rebuild daemon -> list_tabs probe -> run run.ps1 -> check exit code.
#   exit 0 (DONE) / 1 (fatal config/job) / 2 (pool exhausted) -> stop
#   exit 3 (env fatal) / 4 (stall watchdog) -> restart daemon and re-run (resume)
# Hard cap: 8 rounds. run.ps1 auto-computes remaining target from existing files.
$maxRounds = 8
$round = 0
while ($round -lt $maxRounds) {
    $round++
    & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" stop 2>$null
    Start-Sleep -Seconds 2
    & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start 2>&1 | Out-String | Set-Content "$env:TEMP\wb-daemon-schtask.txt" -Encoding UTF8
    $ok = $false
    $i = 0
    for ($i = 1; $i -le 40; $i++) {
        Start-Sleep -Seconds 5
        $probe = '{"action":"list_tabs","args":{},"session":"probe"}'
        [System.IO.File]::WriteAllText("$env:TEMP\wb-probe.json", $probe, [System.Text.UTF8Encoding]::new($false))
        curl.exe -sS --noproxy "*" -X POST http://127.0.0.1:10086/command -H "Content-Type: application/json" --data-binary "@$env:TEMP\wb-probe.json" --output "$env:TEMP\wb-probe-res.json"
        if (Test-Path "$env:TEMP\wb-probe-res.json") {
            $c = [System.IO.File]::ReadAllText("$env:TEMP\wb-probe-res.json")
            if ($c -match '"ok":true') { $ok = $true; break }
        }
    }
    if (-not $ok) {
        [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_fail.txt", "EXT_NOT_READY_round_${round}_try_$i", [System.Text.UTF8Encoding]::new($false))
        exit 1
    }
    Start-Sleep -Seconds 3
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $env:NO_PROXY = "127.0.0.1,localhost"
    $env:no_PROXY = "127.0.0.1,localhost"
    Add-Content -Path $logPath -Value "===== wrapper round $round start $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====="
    & $runPs *>> $logPath
    $code = $LASTEXITCODE
    $status = "round=$round exit=$code at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    [System.IO.File]::WriteAllText("$env:TEMP\wb_wrapper_status.txt", $status, [System.Text.UTF8Encoding]::new($false))
    if ($code -eq 0 -or $code -eq 1 -or $code -eq 2) { break }
    Start-Sleep -Seconds 5
}
