# ============================================================
# wb-core.ps1 — WebBridge 客户端核心 + 环境自检 + 稳定性原语
# ============================================================
# 由 run.ps1 点源（dot-source）加载，函数运行于 run.ps1 作用域。
# 作用域契约：run.ps1 必须在点源【之前】定义：
#   $Config  （含 WebBridgeUrl / DownloadDir）
#   $Session （WebBridge 会话名）
#   $Utf8NoBom
# 本文件只放"与智联页面无关"的通用能力；页面交互见 zhaopin-page.ps1。
# 经验与踩坑记录见 SKILL.md（每个关键函数注释标注了对应问题编号）。
# ============================================================

# ------------------------------------------------------------
# 日志：结构化 + 时间戳 + UTF-8 无 BOM 落盘
# （修复"*> 重定向产生 UTF-16 文件导致日志无法直接读取"的历史痛点）
# 用法：Write-Log "消息" -Level WARN ; 文件 = $Config.DownloadDir\_run_log.txt
# ------------------------------------------------------------
$Script:LogFile = $null
function Initialize-LogFile {
    param([string]$Directory)
    if ($Directory) {
        $Script:LogFile = Join-Path $Directory '_run_log.txt'
        # 启动时清空旧日志（每次运行一份完整日志）
        [System.IO.File]::WriteAllText($Script:LogFile, "", $Utf8NoBom)
    }
}
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','FAIL','STEP')][string]$Level = 'INFO',
        [switch]$NoConsole
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    if (-not $NoConsole) {
        $color = switch ($Level) {
            'OK'   { 'Green' }
            'WARN' { 'Yellow' }
            'FAIL' { 'Red' }
            'STEP' { 'Cyan' }
            default { 'Gray' }
        }
        Write-Host $line -ForegroundColor $color
    }
    if ($Script:LogFile) {
        try { [System.IO.File]::AppendAllText($Script:LogFile, $line + "`r`n", $Utf8NoBom) } catch {}
    }
}

function Wait { param([int]$Ms) Start-Sleep -Milliseconds $Ms }

# ------------------------------------------------------------
# WebBridge HTTP 客户端（每次请求独立临时文件，#40）
# ------------------------------------------------------------
function Send-Web {
    param([string]$Action, $Payload, [string]$SessionLocal = $Session)
    $body = @{ action = $Action; args = $Payload; session = $SessionLocal } | ConvertTo-Json -Compress -Depth 5
    $reqFile = Join-Path $env:TEMP ("wb-req-" + [Guid]::NewGuid().ToString('N') + ".json")
    [System.IO.File]::WriteAllText($reqFile, $body, $Utf8NoBom)
    try {
        # #58：-m 20 硬超时——daemon 无响应（如清理期 stop 挂起）时不至于永久阻塞整个脚本
        return curl.exe -s -m 20 --noproxy '*' -X POST $Config.WebBridgeUrl -H 'Content-Type: application/json' --data-binary "@$reqFile"
    } finally {
        for ($d = 0; $d -lt 3; $d++) {
            try { Remove-Item $reqFile -Force -ErrorAction Stop; break }
            catch { Start-Sleep -Milliseconds 120 }
        }
    }
}

function Invoke-Eval {
    param([string]$Code)
    return Send-Web -Action 'evaluate' -Payload @{ code = $Code }
}

function Invoke-CDP {
    param([string]$Type, [int]$X, [int]$Y, [string]$Button = '', [int]$DeltaX = 0, [int]$DeltaY = 0, [int]$ClickCount = 1)
    $p = @{ type = $Type; x = $X; y = $Y }
    if ($Button)   { $p.button = $Button; $p.clickCount = $ClickCount }
    if ($DeltaY)   { $p.deltaX = $DeltaX; $p.deltaY = $DeltaY }
    $null = Send-Web -Action 'cdp' -Payload @{ method = 'Input.dispatchMouseEvent'; params = $p }
}

# ★ #44：遮挡/背景标签页不会收到真实输入事件，每次交互前必须置前
function Ensure-TabFocused {
    $null = Send-Web -Action 'cdp' -Payload @{ method = 'Page.bringToFront'; params = @{} }
}

# DOM 级 click（selector）；成功返回 $true
function Invoke-Click {
    param([string]$Selector)
    $r = Send-Web -Action 'click' -Payload @{ selector = $Selector }
    return ($r -match '"success":true')
}

# ------------------------------------------------------------
# 稳定性原语 1：通用重试（带退避与每次重试回调）
# ------------------------------------------------------------
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Handler,   # 返回 $true 表示成功
        [int]$MaxRetries = 3,
        [int]$DelayMs = 1000,
        [string]$Description = 'operation',
        [scriptblock]$OnRetry = $null                  # param($attempt)
    )
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        if (& $Handler) { return $true }
        if ($i -lt $MaxRetries - 1) {
            Write-Log "$Description failed (attempt $($i+1)/$MaxRetries) — retrying in ${DelayMs}ms" -Level WARN
            if ($OnRetry) { & $OnRetry $i }
            Wait $DelayMs
        }
    }
    return $false
}

# ------------------------------------------------------------
# 稳定性原语 2：deadline 制条件等待（替代"固定次数×固定间隔"，
#   单个等待永不过期、也永不超时失控）
# ------------------------------------------------------------
function Wait-Until {
    param(
        [Parameter(Mandatory)][scriptblock]$Condition,  # 返回 $true 表示条件满足
        [int]$TimeoutMs = 30000,
        [int]$PollMs = 500,
        [string]$Description = 'condition'
    )
    $deadline = [DateTime]::Now.AddMilliseconds($TimeoutMs)
    while ([DateTime]::Now -lt $deadline) {
        if (& $Condition) { return $true }
        Wait $PollMs
    }
    Write-Log "Wait-Until timeout: $Description (${TimeoutMs}ms)" -Level WARN
    return $false
}

# ------------------------------------------------------------
# 环境自检（把 SKILL.md「标准启动序列」固化成代码）
# ------------------------------------------------------------
function Test-PortListening {
    param([int]$Port = 10086)
    $conn = netstat -ano 2>$null | Select-String ":$Port\s.*LISTENING"
    return ($null -ne $conn -and $conn.Count -gt 0)
}

# ★ 真探测：必须用 list_tabs——扩展在线才返回 ok:true。
#   snapshot 在"daemon 活着但扩展未连接"时返回业务错误 "no tab"，
#   会把未就绪误判为就绪（#55 排障实测）。
function Test-ExtensionReady {
    $r = Send-Web -Action 'list_tabs' -Payload @{}
    return ($r -match '"ok":true')
}

# 返回 $true=环境就绪；$false=致命不可用（调用方应 exit 3）
function Initialize-WebBridgeEnv {
    param([string]$DaemonExe = "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe")

    # 1) 代理铁律：HTTP_PROXY 会劫持 curl 对 127.0.0.1 的请求（#47）
    $env:NO_PROXY  = '127.0.0.1,localhost'
    $env:no_PROXY  = '127.0.0.1,localhost'

    # 2) daemon：端口在听就复用；否则启动
    if (Test-PortListening -Port 10086) {
        Write-Log 'daemon already listening on 10086' -Level INFO
    } else {
        if (-not (Test-Path $DaemonExe)) {
            Write-Log "daemon exe not found: $DaemonExe" -Level FAIL
            return $false
        }
        Write-Log 'starting WebBridge daemon...' -Level STEP
        $null = & $DaemonExe 'start' 2>&1
        Wait 4000
        if (-not (Test-PortListening -Port 10086)) {
            Write-Log 'daemon failed to listen on 10086 after start' -Level FAIL
            return $false
        }
    }

    # 3) 扩展连接：list_tabs 真探测，耐心轮询等重连（反复 stop/start 后扩展掉线实测 round 6 才重连）
    Write-Log 'probing extension via list_tabs (true probe)...' -Level STEP
    $ready = $false
    for ($i = 1; $i -le 40; $i++) {
        if (Test-ExtensionReady) { $ready = $true; break }
        Wait 5000
    }
    if (-not $ready) {
        Write-Log 'extension not connected after 200s — open Chrome and check kimi-webbridge extension' -Level FAIL
        return $false
    }
    Write-Log "environment ready (extension connected, round $i)" -Level OK
    return $true
}

# ------------------------------------------------------------
# 标签页管理（#41/#42/#43 经验的代码化）
# ------------------------------------------------------------
function Switch-ToDetailTab {
    param([int]$MaxWaitMs = 9000, [string]$ExpectName = '', [string[]]$ExcludeUrls = @())
    $deadline = (Get-Date).AddMilliseconds($MaxWaitMs)
    while ((Get-Date) -lt $deadline) {
        $raw = Send-Web -Action 'list_tabs' -Payload @{}
        if ($raw -match '"tabs"') {
            # #42：按 tabId 降序取"最新的"详情页（历史遗留旧详情页排最前）
            $cands = @()
            $ms = [regex]::Matches($raw, '"tabId"\s*:\s*(\d+)\s*,\s*"url"\s*:\s*"([^"]*resumeNumber=[^"]*)"')
            foreach ($m in $ms) {
                $tid = [long]$m.Groups[1].Value
                $u   = ($m.Groups[2].Value -replace '\\/', '/')
                if ($ExcludeUrls -contains $u) { continue }
                $cands += [pscustomobject]@{ TabId = $tid; Url = $u }
            }
            foreach ($c in ($cands | Sort-Object TabId -Descending)) {
                $navRaw = Send-Web -Action 'navigate' -Payload @{ url = $c.Url; newTab = $false }
                if ($navRaw -match '"success":true') {
                    $okBtn = Wait-Until -TimeoutMs 7200 -PollMs 400 -Description 'detail page save-local button' -Condition {
                        (Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return"no";const r=b.getBoundingClientRect();return(r.width>0&&r.height>0)?"yes":"no"})()') -match '"value":"yes"'
                    }
                    return $true
                }
            }
        }
        Wait 500
    }
    return $false
}

# ★ #43：绝不关闭 active 标签页；只关 active=false 的残留详情页
function Clear-StaleDetailTabs {
    $raw = Send-Web -Action 'list_tabs' -Payload @{}
    if ($raw -notmatch '"tabs"') { return 0 }
    $closed = 0
    $objs = [regex]::Matches($raw, '"tabId"\s*:\s*(\d+)\s*,\s*"url"\s*:\s*"([^"]*)"[^}]*?"active"\s*:\s*(true|false)')
    foreach ($m in $objs) {
        $tid      = [long]$m.Groups[1].Value
        $u        = $m.Groups[2].Value
        $isActive = ($m.Groups[3].Value -eq 'true')
        if ($u -notmatch 'resumeNumber=') { continue }
        if ($isActive) { continue }
        $null = Send-Web -Action 'close_tab' -Payload @{ tabId = $tid }
        $closed++
        Wait 400
    }
    return $closed
}

function Test-OnDetailTab {
    $r = Invoke-Eval 'location.href.indexOf("resumeNumber=")>=0?"yes":"no"'
    return ($r -match '"value":"yes"')
}

# ------------------------------------------------------------
# 终止与清理（幂等；不要用 trap 做兜底——trap 会误杀正常流程）
# ------------------------------------------------------------
function Stop-BrowserAutomation {
    param([switch]$Quiet, [switch]$KeepDaemon)
    if ($Script:AutomationStopped) { return }
    $Script:AutomationStopped = $true
    if (-not $Quiet) { Write-Log 'stopping browser automation...' -Level STEP }

    try {
        $null = Send-Web -Action 'close_tab' -Payload @{}
        if (-not $Quiet) { Write-Log 'task tab closed' -Level OK }
    } catch {
        if (-not $Quiet) { Write-Log 'close_tab failed (tab may already be closed)' -Level WARN }
    }
    Wait 800

    try {
        $null = Send-Web -Action 'cdp_disable' -Payload @{}
        if (-not $Quiet) { Write-Log 'CDP automation disconnected' -Level OK }
    } catch {
        if (-not $Quiet) { Write-Log 'cdp_disable failed' -Level WARN }
    }

    if (-not $KeepDaemon) {
        try {
            $null = & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" stop 2>&1
            if (-not $Quiet) { Write-Log 'WebBridge daemon stopped' -Level OK }
        } catch {
            if (-not $Quiet) { Write-Log 'daemon stop failed' -Level WARN }
        }
    } else {
        if (-not $Quiet) { Write-Log 'daemon kept alive (-KeepDaemon)' -Level INFO }
    }

    # 清理历史遗留临时文件（Send-Web 用 GUID 文件用毕自删，这里只兜底 >10min 的残留）
    Get-ChildItem -Path $env:TEMP -Filter 'wb-req-*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-10) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# ------------------------------------------------------------
# JS 字符串安全化：非 ASCII 转 \uXXXX（编码铁律）
# ------------------------------------------------------------
function ConvertTo-JsSafeName {
    param([string]$Name)
    $chars = @()
    for ($ci = 0; $ci -lt $Name.Length; $ci++) {
        $c = $Name[$ci]
        if ([int]$c -gt 127) {
            $chars += '\u{0:X4}' -f [int]$c
        } else {
            $chars += $c
        }
    }
    return ($chars -join '')
}
