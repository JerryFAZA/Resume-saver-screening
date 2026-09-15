# ============================================================
# 智联招聘简历批量下载 — 主入口（重构版）
# 用法: 编辑 config.ps1 后直接运行，或通过参数覆盖
#   .\run.ps1 -Url "..." -JobName "..." -DownloadDir "..." -DownloadCount 30
# ============================================================
param(
    [string]$Url,
    [string]$JobName,
    [string]$DownloadDir,
    [int]$DownloadCount = 0,
    [ValidateSet('pdf','word')]
    [string]$FileFormat = ''
)

# ★ 强制 UTF-8 输出编码，防止中文乱码（尤其在 Agent 后台任务环境中）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- 加载配置（从 JSON 读取，避免中文编码问题）---
$sd = Split-Path -Parent $MyInvocation.MyCommand.Path
$jsonPath = Join-Path $sd 'config.json'
if (Test-Path $jsonPath) {
    $jsonConfig = Get-Content -Raw -Encoding UTF8 $jsonPath | ConvertFrom-Json
    $Config = @{
        Url            = $jsonConfig.Url
        JobName        = $jsonConfig.JobName
        DownloadDir    = $jsonConfig.DownloadDir
        DownloadCount  = $jsonConfig.DownloadCount
        FileFormat     = $jsonConfig.FileFormat
        WebBridgeUrl = 'http://127.0.0.1:10086/command'
        Session      = 'resume-screening'
        SaveLocalX   = 1079
        SaveLocalY   = 142
        SaveConfirmX = 996
        SaveConfirmY = 561
        MaskCloseX   = 30
        MaskCloseY   = 300
        CloseWaitMs        = 700
        ClickWaitMs        = 800
        MouseMoveWaitMs    = 900
        PressReleaseWaitMs = 400
        DialogCheckWaitMs  = 2000
        DownloadWaitMs     = 6000
        ScrollWaitMs       = 1200
        DownloadFilter = '*智联简历*'
        DownloadSource = "$env:USERPROFILE\Downloads"
    }
} else {
    Write-Host 'ERROR: config.json not found!' -ForegroundColor Red
    exit 1
}

if ($Url)          { $Config.Url          = $Url }
if ($JobName)      { $Config.JobName      = $JobName }
if ($DownloadDir)  { $Config.DownloadDir  = $DownloadDir }
if ($DownloadCount -gt 0) { $Config.DownloadCount = $DownloadCount }
if ($FileFormat)   { $Config.FileFormat   = $FileFormat.ToLower() }

# Word 文件通常更大，需要更长等待时间
if ($Config.FileFormat -eq 'word' -and $Config.DownloadWaitMs -lt 10000) {
    $Config.DownloadWaitMs = 10000
    Write-Host "[INFO] Word 模式：DownloadWaitMs 自动调整为 10000ms" -ForegroundColor Cyan
}

# 派生字段：根据 FileFormat 推断文件后缀
$Config.FileExt = if ($Config.FileFormat -eq 'word') { 'docx' } else { 'pdf' }
# DownloadFilter already set in config

# ============================================================
# 强制参数校验
# ============================================================
$missing = @()
if (-not $Config.Url)          { $missing += 'Url' }
if (-not $Config.JobName)      { $missing += 'JobName' }
if (-not $Config.DownloadDir)  { $missing += 'DownloadDir' }

if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Red
    Write-Host '  MISSING REQUIRED PARAMETERS' -ForegroundColor Red
    Write-Host '========================================' -ForegroundColor Red
    Write-Host ''
    Write-Host '请提供以下必要参数后重试：' -ForegroundColor Yellow
    Write-Host ''
    if ('Url'         -in $missing) { Write-Host '  1. Url          - 智联招聘推荐页完整 URL（含 jobNumber）' -ForegroundColor White }
    if ('JobName'     -in $missing) { Write-Host '  2. JobName      - 岗位名称（需与页面标签完全一致）' -ForegroundColor White }
    if ('DownloadDir' -in $missing) { Write-Host '  3. DownloadDir  - 简历保存目录（绝对路径）' -ForegroundColor White }
    Write-Host ''
    Write-Host '用法示例：' -ForegroundColor Gray
    Write-Host '  .\run.ps1 -Url "https://rd6.zhaopin.com/app/recommend?jobNumber=XXX" -JobName "销售部经理" -DownloadDir "C:\Resumes" -DownloadCount 30' -ForegroundColor Gray
    Write-Host ''
    Write-Host '或编辑 config.ps1 填写参数后直接运行：.\run.ps1' -ForegroundColor Gray
    Write-Host '========================================' -ForegroundColor Red
    exit 1
}

# --- 初始化 ---
$Utf8NoBom    = New-Object System.Text.UTF8Encoding $false
# 兼容保留：Send-Web 已改为每次请求生成独立临时文件（修复 #40），
# 这里仅用于启动时清理历史遗留的固定名临时文件。
$ReqFilePath  = Join-Path $env:TEMP 'wb-download.json'
Remove-Item $ReqFilePath -Force -ErrorAction SilentlyContinue
$Session      = 'resume-screening'

# 确保守护进程运行
$null = & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start 2>&1

# 确保目标目录存在
if (-not (Test-Path $Config.DownloadDir)) {
    $null = New-Item -ItemType Directory -Path $Config.DownloadDir -Force
}

# ============================================================
# 工具函数
# ============================================================

function Send-Web {
    param([string]$Action, $Payload, [string]$SessionLocal = $Session)
    $body = @{ action = $Action; args = $Payload; session = $SessionLocal } | ConvertTo-Json -Compress -Depth 5

    # ★ 修复 #40：原实现固定复用同一个 wb-download.json，curl.exe 可能仍持有
    #   该文件句柄（异步/未完全退出），下一次 WriteAllText 立刻抛 IOException
    #   "文件正由另一进程使用"，导致**首次请求之后的所有请求全部失败**
    #   （表现为 SaveLocal 探测成功但点击无响应、No save dialog、卡死）。
    #   修复：每次请求使用独立的临时文件（GUID 命名），写后立即交给 curl，
    #   用毕删除；既无锁竞争，也不会被并发调用互相覆盖。
    $reqFile = Join-Path $env:TEMP ("wb-req-" + [Guid]::NewGuid().ToString('N') + ".json")
    [System.IO.File]::WriteAllText($reqFile, $body, $Utf8NoBom)
    try {
        return curl.exe -s -X POST $Config.WebBridgeUrl -H 'Content-Type: application/json' --data-binary "@$reqFile"
    } finally {
        # 给 curl 一点时间释放句柄；重试删除，仍失败则留给系统清理（不阻断主流程）
        for ($d = 0; $d -lt 3; $d++) {
            try { Remove-Item $reqFile -Force -ErrorAction Stop; break }
            catch { Start-Sleep -Milliseconds 120 }
        }
    }
}

function Wait { param([int]$Ms) Start-Sleep -Milliseconds $Ms }

function Invoke-CDP {
    param([string]$Type, [int]$X, [int]$Y, [string]$Button = '', [int]$DeltaX = 0, [int]$DeltaY = 0, [int]$ClickCount = 1)
    $p = @{ type = $Type; x = $X; y = $Y }
    if ($Button)   { $p.button = $Button; $p.clickCount = $ClickCount }
    if ($DeltaY)   { $p.deltaX = $DeltaX; $p.deltaY = $DeltaY }
    $null = Send-Web -Action 'cdp' -Payload @{ method = 'Input.dispatchMouseEvent'; params = $p }
}

function Invoke-Eval {
    param([string]$Code)
    return Send-Web -Action 'evaluate' -Payload @{ code = $Code }
}

# ============================================================
# ★★★ 修复 #44：DOM 级点击（click action）与标签页置前（bringToFront）★★★
# ============================================================
# 【为什么必须用 click action 而不是 CDP 坐标点击】
#   实测（2026-09-15）WebBridge 的 `mouse_click` / CDP `Input.dispatchMouseEvent`
#   在下列情况会**静默失败**（返回 ok:true 但页面毫无反应）：
#     1) 标签页被其他标签页遮挡（backgrounded / occluded）时，
#        Chrome **不会把真实输入事件投递给隐藏的 render widget**。
#        官方报错原文：
#          "the click did not reach the page — no pointerdown/mousedown fired.
#           The tab is likely backgrounded and occluded (another tab in its window
#           is shown, so its render widget is hidden and real input isn't delivered)."
#     2) 元素在虚拟列表中位于视口外（如 y=1853 / 视口高 735），
#        普通 scrollIntoView 对该虚拟滚动容器无效。
#
# 【修复方案】
#   - `Ensure-TabFocused`：每次交互前调用 CDP `Page.bringToFront`，
#     把 session 的标签页提到窗口最前，保证真实输入能投递。
#   - `Invoke-Click`：改用 WebBridge 的 `click` action（DOM 级派发，
#     内部自动处理滚动与不可见元素），**不再依赖坐标**。
#
# 【实测验证】`Page.bringToFront` + `click .talent-basic-info__name` 之后：
#   - 详情面板**在同一标签页内**渲染（`.resume-detail-wrap` 出现）
#   - `.resume-button.position-r`（存至本地）出现 → 面板可操作
#   结论：**根本不需要切换标签页**，此前的 Switch-ToDetailTab 方案是误判。
function Ensure-TabFocused {
    $null = Send-Web -Action 'cdp' -Payload @{ method = 'Page.bringToFront'; params = @{} }
}

# 用 DOM 级 click（selector）点击元素；成功返回 $true。
# $Selector 支持 CSS 选择器或扩展生成的 @e 引用。
function Invoke-Click {
    param([string]$Selector)
    $r = Send-Web -Action 'click' -Payload @{ selector = $Selector }
    return ($r -match '"success":true')
}

# ============================================================
# ★★★ 修复 #41：把 session 切换到"详情标签页" ★★★
# ============================================================
# 【根因】智联推荐页点击候选人后，简历详情在**新标签页**中打开，该标签页的
#   URL 带有 `resumeNumber=<encoded>` 参数；而 WebBridge 的 session 始终绑定
#   在**列表标签页**（active=true）。两者是不同 tab：
#     - 列表页 tabId 1861309718  url: ...&tab=recommend#sortType=recommend
#     - 详情页 tabId 1861309715  url: ...&tab=recommend&resumeNumber=XXX#sortType=...
#   实测确认：`.resume-button.position-r`（存至本地）与 button[保存]
#   **只存在于详情页**，列表中根本不存在。
#   因此脚本在列表页里探测按钮永远失败 → 退回硬编码兜底坐标 →
#   所有 CDP 鼠标事件全部打在列表页上 → 点击无响应 →
#   "No save dialog after 4 attempts" → 空转卡死（这就是用户看到的
#   "一直在翻候选人列表，但没点进去任何一份详情"）。
#
# 【修复】点击候选人后，轮询 session 内所有标签页，找到 URL 含
#   `resumeNumber=` 的详情页，然后用 navigate 把 session 导航到该 URL，
#   使 session 的 active tab 变成详情页；此后所有 evaluate / CDP 事件
#   才会真正作用在详情页上。
#
# 【实测验证】（2026-09-15）
#   navigate 到详情 URL 后：
#     document.querySelector('.resume-button.position-r')  → 存在，(1079,142)
#     button[保存]                                          → 存在，(996,561)
#   与 Config.SaveLocalX/Y、SaveConfirmX/Y 完全一致；随后点击
#   "存至本地" → "word" → "保存"，成功产出
#   `260915_test\何先生_24岁_智联简历_31791.docx`（100694 字节）。
#
# 返回 $true 表示已成功切到详情页，$false 表示未找到（调用方需跳过或重试）。
function Switch-ToDetailTab {
    param([int]$MaxWaitMs = 9000, [string]$ExpectName = '', [string[]]$ExcludeUrls = @())

    $deadline = (Get-Date).AddMilliseconds($MaxWaitMs)
    while ((Get-Date) -lt $deadline) {
        $raw = Send-Web -Action 'list_tabs' -Payload @{}
        if ($raw -match '"tabs"') {
            # ★ 修复 #42：必须取"最新的"详情页，而不是"第一个"。
            #   实测 list_tabs 会返回多个 resumeNumber 标签页（历史遗留的旧详情页
            #   排在最前），若按首个匹配就会 navigate 到**上一位候选人**的简历，
            #   表现为"存至本地能点开但下错人"。这里按 tabId 降序取最大者
            #   （tabId 单调递增），并排除调用方显式传入的已处理 URL。
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
                    # 轮询确认详情页关键元素已就绪（最多 ~7.2s）
                    for ($vw = 0; $vw -lt 18; $vw++) {
                        $chk = Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return"no";const r=b.getBoundingClientRect();return(r.width>0&&r.height>0)?"yes":"no"})()'
                        if ($chk -match '"value":"yes"') { return $true }
                        Wait 400
                    }
                    return $true
                }
            }
        }
        Wait 500
    }
    return $false
}

# ★ 修复 #42：清理残留的详情标签页（历史运行遗留的 resumeNumber 标签页会让
#   Switch-ToDetailTab 误取旧简历）。在阶段 4 开始前调用一次。
#
# ⚠️ 修复 #43（实测 2026-09-15）：**绝不能关闭 session 当前 active 的标签页**。
#   实测把 active 的详情页 close_tab 掉之后，session 直接变成"无标签页"状态
#   （list_tabs 返回 `tabs: []`），后续所有 evaluate 返回 ok:false，
#   表现为主循环连续 [SKIP] Not found after full-list sweep。
#   因此这里只关 active=false 的残留详情页；若 active 的那个也是详情页，
#   则不关它，改由调用方用 navigate 导航回列表 URL（导航会复用当前标签页）。
function Clear-StaleDetailTabs {
    $raw = Send-Web -Action 'list_tabs' -Payload @{}
    if ($raw -notmatch '"tabs"') { return 0 }
    $closed = 0
    # 逐个 tab 对象解析：同时取出 tabId / url / active 三个字段
    $objs = [regex]::Matches($raw, '"tabId"\s*:\s*(\d+)\s*,\s*"url"\s*:\s*"([^"]*)"[^}]*?"active"\s*:\s*(true|false)')
    foreach ($m in $objs) {
        $tid    = [long]$m.Groups[1].Value
        $u      = $m.Groups[2].Value
        $isActive = ($m.Groups[3].Value -eq 'true')
        if ($u -notmatch 'resumeNumber=') { continue }   # 只处理详情页
        if ($isActive) { continue }                       # ★ 绝不关 active 标签页
        $null = Send-Web -Action 'close_tab' -Payload @{ tabId = $tid }
        $closed++
        Wait 400
    }
    return $closed
}

# 判断当前 session 的 active tab 是否已经是详情页（URL 含 resumeNumber）
function Test-OnDetailTab {
    $r = Invoke-Eval 'location.href.indexOf("resumeNumber=")>=0?"yes":"no"'
    return ($r -match '"value":"yes"')
}

# ★★★ 终止与清理：确保脚本任何退出路径都停止浏览器自动化并释放浏览器 ★★★
# 幂等设计：可重复调用；$Script:AutomationStopped 标记避免重复执行。
# 修复 #37：任务结束后残留的标签页/自动化连接会让浏览器"看起来还在自动下载"，
#           必须显式清理（close tab + CDP disable + daemon stop）。
#
# ⚠️ 不要用 trap 做兜底清理（实测 2026-09-15）：
#    PowerShell 的 trap 会捕获作用域内任何 terminating error，在长流程脚本中
#    极易被无关的瞬时错误触发，从而**误杀正常执行中的任务**（实测下载第 1 条时
#    就被 trap 中断）。清理只放在正常流程末尾的 Stop-BrowserAutomation 中，
#    并通过幂等标记 + 多处显式调用来保证覆盖。
function Stop-BrowserAutomation {
    param([switch]$Quiet, [switch]$KeepDaemon)
    if ($Script:AutomationStopped) { return }
    $Script:AutomationStopped = $true
    if (-not $Quiet) { Write-Host ''; Write-Host '[9] Stopping browser automation...' -ForegroundColor Cyan }

    # 1) 关闭为本次任务创建的标签页（避免残留会话继续触发页面行为）
    try {
        $null = Send-Web -Action 'close_tab' -Payload @{}
        if (-not $Quiet) { Write-Host '  [OK] Task tab closed' -ForegroundColor Green }
    } catch {
        if (-not $Quiet) { Write-Host '  [WARN] close_tab failed (tab may already be closed)' -ForegroundColor Yellow }
    }
    Wait 800

    # 2) 断开 CDP 自动化连接（释放浏览器控制权，用户可正常操作）
    try {
        $null = Send-Web -Action 'cdp_disable' -Payload @{}
        if (-not $Quiet) { Write-Host '  [OK] CDP automation disconnected' -ForegroundColor Green }
    } catch {
        if (-not $Quiet) { Write-Host '  [WARN] cdp_disable failed' -ForegroundColor Yellow }
    }

    # 3) 停止 daemon（真正终止自动化常驻进程）
    #    -KeepDaemon：同一批次连续跑多岗位时跳过，保持 daemon 存活以免反复重启
    if (-not $KeepDaemon) {
        try {
            $null = & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" stop 2>&1
            if (-not $Quiet) { Write-Host '  [OK] WebBridge daemon stopped' -ForegroundColor Green }
        } catch {
            if (-not $Quiet) { Write-Host '  [WARN] daemon stop failed' -ForegroundColor Yellow }
        }
    } else {
        if (-not $Quiet) { Write-Host '  [SKIP] daemon kept alive (-KeepDaemon)' -ForegroundColor DarkGray }
    }

    # 4) 清理历史遗留的固定名临时文件（Send-Web 已改用 GUID 独立文件，用毕自删）
    Remove-Item $ReqFilePath -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path $env:TEMP -Filter 'wb-req-*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-10) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# 从 evaluate 响应中解析 "cx,cy" 格式坐标（JS 返回简单字符串，避免复杂 JSON 正则）
function Parse-Coords {
    param([string]$Raw)
    if ($Raw -match '"value":"(\d+),(\d+)"') {
        $cx = [int]$Matches[1]; $cy = [int]$Matches[2]
        if ($cx -gt 10 -and $cy -gt 10) { return @($cx, $cy) }
    }
    return $null
}

<#
.SYNOPSIS 通过 getBoundingClientRect() 获取"保存"按钮的中心坐标（CSS 像素）
.DESCRIPTION CDP Input.dispatchMouseEvent 使用 CSS 像素坐标，与 getBoundingClientRect() 一致。
             之前用 DOM.getBoxModel 会因 nodeId=0（未注册）失败，且可能返回设备像素导致坐标偏移。
#>
function Get-SaveButtonCoords {
    $r = Invoke-Eval '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0){const r=x.getBoundingClientRect();if(r.width<=0||r.height<=0)continue;return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)}}return "0,0"})()'
    return Parse-Coords $r
}

# ★★★ CRITICAL: ConvertTo-JsSafeName 必须在此处（阶段1之前）定义 ★★★
# PowerShell 是顺序解释执行的，函数必须先定义后调用。
# 阶段1第188行就会调用此函数，原定义在551行导致 CommandNotFoundException。
# 将姓名转为 JS 安全的 Unicode 转义字符串，避免 PS 字符串插值时中文乱码
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

# ============================================================
# 阶段 1：导航与岗位选择
# ============================================================
Write-Host '[1] Navigate...' -ForegroundColor Cyan
# ★ 使用 newTab=$true 确保 WebBridge 创建新标签页并自动激活为 current tab
#    之前 newTab=$false 导致标签页可能未获得焦点，用户需手动点击
$null = Send-Web -Action 'navigate' -Payload @{
    url = $Config.Url
    newTab = $true
    group_title = 'Zhaopin Resume Screening'
}
# 等待标签页创建
Wait 3000

# ★ 修复：navigate(newTab=true) 创建标签页后，WebBridge 内部将其设为 current tab
#    且 CDP 已连接到新标签页。但 Chrome UI 层面可能未自动切换（实测 list_tabs
#    显示 active=false）。通过 CDP Page.bringToFront 强制浏览器切换到当前
#    CDP 连接的标签页（即刚创建的新标签页）。
#    ⚠️ 注意：不能用 find_tab（URL 模糊匹配会错误切回旧标签页），
#       也不能用 Target.activateTarget（Chrome 安全策略阻止）。
Write-Host '  Focusing tab via CDP Page.bringToFront...' -ForegroundColor Gray
$null = Send-Web -Action 'cdp' -Payload @{
    method = 'Page.bringToFront'
    params = @{}
}
Write-Host '  Tab focused' -ForegroundColor Green
# 等待页面加载完成（智联招聘页面较重，需较长时间）
Wait 3000

Write-Host '[2] Select job:' $Config.JobName -ForegroundColor Cyan

# ★ 修复 #28：岗位 Tab 查找改为双路径策略（snapshot 包含匹配 + JS 模糊匹配交叉进行）
# 路径 A：snapshot 包含匹配 — 支持岗位名带后缀（如"渠道部经理·协作未上线"）
# 路径 B：JS evaluate 模糊匹配 — 用 textContent.includes 兜底
# 两路径交叉进行，每 2s 一轮，最多 15 轮（30s），覆盖慢网络/慢渲染场景
$jobTabClicked = $false
$safeJobName = ConvertTo-JsSafeName -Name $Config.JobName

for ($retrySnap = 0; $retrySnap -lt 15; $retrySnap++) {
    # --- 路径 A: snapshot 包含匹配（用 .*? 允许 name 前后有额外字符） ---
    $snap = Send-Web -Action 'snapshot' -Payload @{}
    if ($snap -match '"name":"[^"]*' + [regex]::Escape($Config.JobName) + '[^"]*","ref":"(@e\d+)"') {
        $null = Send-Web -Action 'click' -Payload @{ selector = $Matches[1] }
        Write-Host '  [A] Snapshot matched & clicked:' $Matches[1] -ForegroundColor Green
        $jobTabClicked = $true
        break
    }

    # --- 路径 B: JS evaluate 模糊匹配（用 textContent.includes 查找 .job-pane__item） ---
    $jsFindTab = '(()=>{const links=document.querySelectorAll(''.job-pane__item'');for(const l of links){if(l.textContent.trim().includes("' + $safeJobName + '")){l.click();return"ok:"+l.textContent.trim().substring(0,40)}}return"no"})()'
    $ftResult = Invoke-Eval $jsFindTab
    if ($ftResult -match '"value":"ok:(.*?)"') {
        Write-Host '  [B] JS fuzzy matched & clicked:' $Matches[1] -ForegroundColor Green
        $jobTabClicked = $true
        break
    }

    if ($retrySnap -eq 0) {
        Write-Host '  Job tab not found yet, waiting for page render (retry every 2s)...' -ForegroundColor Gray
    }
    if ($retrySnap -gt 0 -and $retrySnap % 3 -eq 0) {
        Write-Host "  Still searching... (retry $retrySnap/15)" -ForegroundColor Gray
    }
    Wait 2000
}

if (-not $jobTabClicked) {
    # 失败时输出页面上所有岗位标签，帮助诊断
    Write-Host ''
    Write-Host '  Dumping all job tabs on page for diagnosis:' -ForegroundColor Yellow
    $dumpJs = '(()=>{const links=document.querySelectorAll(''.job-pane__item'');const names=[];for(const l of links){names.push(l.textContent.trim().substring(0,50))}return JSON.stringify(names)})()'
    $dumpR = Invoke-Eval $dumpJs
    if ($dumpR -match '"value":"(\[.*\])"') {
        try { $tabs = $Matches[1] -replace '\\"','"' | ConvertFrom-Json; Write-Host "  Tabs: $($tabs -join ' | ')" -ForegroundColor Gray } catch {}
    }
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Red
    Write-Host '  FATAL: Job tab not found on page' -ForegroundColor Red
    Write-Host '========================================' -ForegroundColor Red
    Write-Host "  JobName: $($Config.JobName)" -ForegroundColor Yellow
    Write-Host '  The specified job tab could not be located after 15 retries (30s).' -ForegroundColor Yellow
    Write-Host '  Possible reasons:' -ForegroundColor Yellow
    Write-Host '    1. Job name does not match any page label (check spelling/characters)' -ForegroundColor Yellow
    Write-Host '    2. The URL jobNumber does not correspond to this job' -ForegroundColor Yellow
    Write-Host '    3. The page did not fully render the job tabs in time' -ForegroundColor Yellow
    Write-Host '    4. Above "Tabs:" list shows all available job labels on page' -ForegroundColor Yellow
    Write-Host '========================================' -ForegroundColor Red
    Stop-BrowserAutomation -Quiet
    exit 1
}

Wait 3000

# 等待候选人列表加载
Write-Host '  Waiting for candidate list...' -ForegroundColor Gray
for ($w = 0; $w -lt 10; $w++) {
    $r = Invoke-Eval "String(document.querySelectorAll('.talent-basic-info__name').length)"
    if ($r -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) {
        Write-Host ('  Loaded ' + $Matches[1] + ' candidate(s)') -ForegroundColor Gray
        break
    }
    Wait 1000
}

# ============================================================
# ★ 岗位验证：确保当前激活的岗位与指令一致
# ============================================================
Write-Host '[2.1] Verifying active job...' -ForegroundColor Cyan

# 智联岗位标签栏中，当前激活/选中的 Tab 通常有特殊样式类名（如 .is-active, .active, [aria-selected="true"] 等）
# 尝试多种方式获取当前激活岗位的名称
$activeJobJs = @'
(()=>{
  // ★ 修复 #29：策略优先级重排，优先使用精确选择器 .job-pane__item--active
  // 避免全局 .active 选择器误匹配到导航栏"推荐"等无关元素
  // 智联岗位标签栏使用 class="job-pane__item job-pane__item--active" 标记当前激活岗位

  // 策略1: 直接匹配岗位标签栏的激活 class（最精确，优先使用）
  let el = document.querySelector('.job-pane__item--active');
  if (el && el.textContent.trim()) return el.textContent.trim();

  // 策略2: 在岗位容器内查找激活元素（限定作用域）
  const container = document.querySelector('.job-pane, [class*="job-pane"], [class*="job-list"]');
  if (container) {
    const activeEl = container.querySelector('.job-pane__item--active, .is-active, .active, [class*="is-active"], [class*="active"], [aria-selected="true"]');
    if (activeEl && activeEl.textContent.trim()) return activeEl.textContent.trim();
  }

  // 策略3: 查找带有 aria-selected="true" 的 tab/link
  el = document.querySelector('[aria-selected="true"]');
  if (el && el.textContent.trim()) return el.textContent.trim();

  // 策略4: 查找当前高亮的 link/tab（aria-current）
  el = document.querySelector('[aria-current="page"], [aria-current="true"]');
  if (el && el.textContent.trim()) return el.textContent.trim();

  // 策略5: 遍历岗位标签元素，找第一个包含高亮样式类的
  const candidates = document.querySelectorAll('.job-pane__item, [class*="tab"], [class*="Tab"], [class*="job-item"], [role="tab"]');
  for (const c of candidates) {
    if (c.classList.contains('job-pane__item--active') ||
        c.classList.contains('active') || c.classList.contains('is-active') ||
        c.getAttribute('aria-selected') === 'true' || c.getAttribute('aria-current') === 'page') {
      const txt = c.textContent.trim();
      if (txt) return txt;
    }
  }

  return '';
})()
'@

$activeJob = ''
$activeR = Invoke-Eval $activeJobJs
# ★ 修复 #38：原正则 '"value":"([^"]*)"' 会贪婪匹配到响应里任意 "value":"..." 字段。
#   JS 返回空串时响应可能是 {"value":""} 或含其它 value 键，正则可能捕获到
#   无关片段（实测捕获到 "no"）→ 被当成岗位名 → 误报 "Job mismatch" 并 exit 1。
#   修复：只接受"看起来像岗位名"的值（含中文、长度>=2、且不是常见 JS 返回值）。
if ($activeR -match '"value":"([^"]*)"') {
    $cand = $Matches[1].Trim()
    if ($cand -and $cand.Length -ge 2 -and
        $cand -notin @('no', 'ok', 'has', 'yes', 'none', 'null', 'true', 'false', '0', '1') -and
        $cand -match '[\u4e00-\u9fa5]') {
        $activeJob = $cand
    }
}

# ★ 修复 #38（续）：若 JS 探测失败，不要急于判定"不匹配"。
#   先重试几次（页面可能仍在渲染），仍失败则降级为"无法检测 → 警告并继续"，
#   而不是 exit 1 直接终止整个下载任务。
if (-not $activeJob) {
    Write-Host '  [WARN] Could not detect active job via JS — checking snapshot...' -ForegroundColor Yellow
    $snap = Send-Web -Action 'snapshot' -Payload @{}
    # 在快照中搜索带有 [active] 或 [selected] 等标记的 link
    if ($snap -match '"role":"link"[^}]*"name":"([^"]+)"[^}]*"active"') {
        $activeJob = $Matches[1]
    } elseif ($snap -match '"role":"link"[^}]*"name":"([^"]+)"[^}]*"selected"') {
        $activeJob = $Matches[1]
    }
}

if ($activeJob) {
    Write-Host "  Active job detected: [$activeJob]" -ForegroundColor Gray
    Write-Host "  Expected job:        [$($Config.JobName)]" -ForegroundColor Gray

    # 比较：支持模糊匹配（trim 后忽略首尾空白，可能的话也忽略数量标记如 "(30)" 后缀）
    $activeClean = $activeJob.Trim()
    $expectedClean = $Config.JobName.Trim()

    # 同时尝试去掉括号内数量信息，如 "销售部经理(30)" → "销售部经理"
    $activeNoParen = $activeClean -replace '\s*\(.*$', ''
    $expectedNoParen = $expectedClean -replace '\s*\(.*$', ''

    if ($activeClean -eq $expectedClean -or $activeNoParen -eq $expectedNoParen -or $activeClean -like "*$expectedClean*") {
        Write-Host '  [OK] Active job matches expected job — proceeding.' -ForegroundColor Green
    } else {
        # ★ 修复 #38（续）：仅在"确实读到了另一个岗位名"时才判定不匹配并终止。
        #   探测失败（空值）走降级路径（警告继续），不阻断整个下载任务。
        Write-Host ''
        Write-Host '========================================' -ForegroundColor Red
        Write-Host '  FATAL: Job mismatch detected!' -ForegroundColor Red
        Write-Host '========================================' -ForegroundColor Red
        Write-Host "  Expected job : $($Config.JobName)" -ForegroundColor Yellow
        Write-Host "  Active job   : $activeJob" -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  The page is currently showing candidates for a different job.' -ForegroundColor Yellow
        Write-Host '  Possible reasons:' -ForegroundColor Yellow
        Write-Host '    1. The specified job tab was not found or clickable on the page' -ForegroundColor Yellow
        Write-Host '    2. The browser restored a previously selected tab from last session' -ForegroundColor Yellow
        Write-Host '    3. The job name parameter does not exactly match the page label' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Action required: Verify the JobName parameter matches the page label exactly.' -ForegroundColor Yellow
        Write-Host '========================================' -ForegroundColor Red
        Stop-BrowserAutomation -Quiet
        exit 1
    }
} else {
    Write-Host '  [WARN] Unable to detect active job — skipping verification (proceed with caution)' -ForegroundColor Yellow
}

# ============================================================
# 阶段 2：收集候选人姓名（虚拟滚动累积）
# ============================================================
Write-Host '[3] Collecting names...' -ForegroundColor Cyan

# ★ 修复 #39：收集前先确认列表容器已渲染。
#   上一轮 abort/cleanup 会关闭标签页，重新 navigate 后页面需时间渲染；
#   若此时直接进循环，首个 evaluate 可能返回 ok:false（"no tab"）导致
#   整个收集循环静默失败 → Found 0 unique names → 误报 no names abort。
Write-Host '  Waiting for list container to render...' -ForegroundColor Gray
$listReady = $false
for ($w = 0; $w -lt 20; $w++) {
    $chk = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
    if ($chk -match '"ok":false') {
        Write-Host '  [WARN] Session has no tab yet — re-navigating...' -ForegroundColor Yellow
        $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
        Wait 3000
        $null = Send-Web -Action 'cdp' -Payload @{ method = 'Page.bringToFront'; params = @{} }
        Wait 2000
        continue
    }
    if ($chk -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) {
        Write-Host "  List ready ($($Matches[1]) nodes in viewport)" -ForegroundColor Gray
        $listReady = $true
        break
    }
    Wait 1000
}
if (-not $listReady) {
    Write-Host '  [WARN] List container not confirmed — proceeding anyway (will retry in loop)' -ForegroundColor Yellow
}

$null = Invoke-Eval 'window._SN = new Set()'
Wait 300

$minNeeded = [Math]::Max($Config.DownloadCount * 2, 40)  # 至少收集目标数×2 或 40 个
$last      = 0
$stag      = 0
$maxScrolls = 80  # 硬上限防止死循环（deep list needs many scrolls）
$stagnantLimit = 12  # ★ 修复：3 轮太激进，虚拟滚动偶发不增量会误判到底；提高到 12 轮
$evalErrors = 0    # ★ 修复 #39：连续 evaluate 失败计数，用于区分离线/到底

# ★ 关键修复：直接滚动列表容器（JS），而非 CDP mouseWheel 固定坐标。
#   CDP mouseWheel 在 (500,600) 对虚拟列表滚动不可靠，会在 ~31 个时误判停滞。
for ($i = 0; $i -lt $maxScrolls; $i++) {
    $js = '(()=>{const c=document.querySelector(''.app-layout--default'')||document.scrollingElement;const e=document.querySelectorAll(''.talent-basic-info__name'');e.forEach(el=>{const m=el.textContent.trim().match(/^\S+/);if(m)window._SN.add(m[0])});if(c)c.scrollTop=c.scrollTop+900;else window.scrollBy(0,900);return ''''})()'
    $null = Invoke-Eval $js
    Wait $Config.ScrollWaitMs

    $r = Invoke-Eval 'String(window._SN.size)'
    if ($r -match '"ok":false') {
        # ★ 修复 #39：evaluate 失败（tab 丢失/断连）不应静默跳过，否则循环空转到
        #   maxScrolls 后报 0 个名字。这里显式重连并计数。
        $evalErrors++
        Write-Host "  Scroll $i : [EVAL ERROR] session/connection issue (errors=$evalErrors)" -ForegroundColor Yellow
        if ($evalErrors -ge 3) {
            Write-Host '  [WARN] 3 consecutive eval errors — re-navigating to recover...' -ForegroundColor Yellow
            $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
            Wait 4000
            $null = Send-Web -Action 'cdp' -Payload @{ method = 'Page.bringToFront'; params = @{} }
            Wait 3000
            $null = Invoke-Eval 'window._SN = new Set()'
            Wait 500
            $evalErrors = 0
        }
        continue
    }
    if ($r -match '"value":"(\d+)"') {
        $evalErrors = 0
        $cur = [int]$Matches[1]
        if ($cur -eq $last) { $stag++ } else { $stag = 0; $last = $cur }
        Write-Host "  Scroll $i : size=$cur (stagnant=$stag)" -ForegroundColor Gray
        # 收集够了就停，或连续多轮无新增（确认到底）才停
        if ($cur -ge $minNeeded) { Write-Host "  Enough names collected, stopping." -ForegroundColor Green; break }
        if ($stag -ge $stagnantLimit) { Write-Host "  Reached list end (stagnant limit)." -ForegroundColor Green; break }
    } else {
        Write-Host "  Scroll $i : [WARN] unparseable response" -ForegroundColor Yellow
    }
}

# 解析姓名数组
$names = @()
$r = Invoke-Eval 'JSON.stringify(Array.from(window._SN))'
if ($r -match '"value":"(\[.*\])"') {
    try {
        $json = $Matches[1] -replace '\\"', '"'
        $raw  = $json | ConvertFrom-Json
        $names = @($raw | Where-Object { $_ } | ForEach-Object { $_ -replace '\s', '' } | Where-Object { $_ })
    } catch {
        Write-Host '  [ERROR] Name parse failed:' $_ -ForegroundColor Red
    }
}
Write-Host ('  Found ' + $names.Count + ' unique names') -ForegroundColor Green

if ($names.Count -eq 0) {
    Write-Host '[FAIL] No names — abort.' -ForegroundColor Red
    # ★ 修复 #39：abort 时保留 daemon（-KeepDaemon），便于 Agent 立即重试；
    #   仍关闭标签页与 CDP 连接，避免残留自动化。
    Stop-BrowserAutomation -Quiet -KeepDaemon
    exit 1
}

# ★ 关键修复：收集阶段把列表滚到了底部，进入下载阶段前必须重置回顶部。
#   否则虚拟滚动只保留视口内约 20 个节点，且已到底无法再往下滚 →
#   查找候选人时 15 次重试全在原地打转 → 大量 [SKIP] Not found in viewport 空转。
Write-Host '  Resetting list to top before download...' -ForegroundColor Gray
$null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return "ok"})()'
Wait 2500

# ============================================================
# ★★★ 修复 #42/#43：清理历史遗留的详情标签页 ★★★
# ============================================================
# 之前运行/手工调试留下的 resumeNumber 标签页会一直挂在浏览器里。
# Switch-ToDetailTab 若按"首个匹配"取，就会 navigate 到**上一位候选人**
# 的详情页 → 出现"能点开保存框但下载的是别人/旧简历"。
# 注意：只关 active=false 的残留详情页（详见 Clear-StaleDetailTabs 注释）；
# 若当前 active 就是详情页，则先导航回列表页，保证主循环从列表开始找候选人。
$staleClosed = Clear-StaleDetailTabs
if ($staleClosed -gt 0) {
    Write-Host "  [OK] Closed $staleClosed stale detail tab(s)" -ForegroundColor Green
    Wait 800
}
# 兜底：若当前 active tab 仍是详情页（或 session 已无标签页），导航回列表 URL
$null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $false }
for ($bw0 = 0; $bw0 -lt 20; $bw0++) {
    $chk0 = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
    if ($chk0 -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) { break }
    Wait 500
}
Wait 1000

# ============================================================
# 去重辅助函数
# ============================================================

<#
.SYNOPSIS 识别智联简历文件（纯 ASCII 安全的结构匹配，不依赖脚本内的中文字面量）。
.DESCRIPTION ★★★ 核心修复 ★★★
  在中文 Windows + PowerShell 5.x 环境下，UTF-8 无 BOM 的 .ps1 脚本文件会被
  GBK 解析，导致脚本内的中文字面量（如 "*智联简历*"）全部变成乱码。
  Get-ChildItem -Include "*智联简历*" 实际传入乱码，永远匹配不到中文文件名。
  
  解决方案：智联简历文件名格式为 {姓名}_{年龄}岁_智联简历_{5位数字}.{pdf|docx}，
  用纯 ASCII 的结构特征识别：
    - 以 .pdf 或 .docx 结尾
    - 至少 4 段（用 _ 分割）
    - 最后一段是纯数字 + 扩展名（如 "62595.docx"）
    - 倒数第二段包含数字（年龄段，如 "27岁"）
  不依赖任何中文字面量。
#>
function Get-ZhaopinFiles {
    param([string]$Dir)
    # 列出所有 pdf/docx，再用结构化特征过滤（纯 ASCII 安全，不依赖中文）
    $all = Get-ChildItem -Path "$Dir\*" -Include "*.pdf", "*.docx" -ErrorAction SilentlyContinue
    if (-not $all) { return @() }
    $result = $all | Where-Object {
        $n = $_.Name
        $parts = $n -split '_'
        # 智联文件名至少 4 段: 姓名_年龄岁_智联简历_数字.ext
        ($parts.Count -ge 4) -and ($n -match '_\d+\.(pdf|docx)$')
    }
    return @($result)
}

<#
.SYNOPSIS 从智联文件名中提取姓名（第一个 _ 之前的部分）。
#>
function Get-NameFromZhilianFile {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $idx = $base.IndexOf('_')
    if ($idx -gt 0) { return $base.Substring(0, $idx) }
    return ''
}

<#
.SYNOPSIS 从智联文件名中提取年龄（第二个 _ 之前、"岁"之前的数字）。
#>
function Get-AgeFromZhilianFile {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    # 格式: 姓名_27岁_...  → 取第二个 _ 前的内容中的数字
    $parts = $base -split '_'
    if ($parts.Count -ge 2) {
        if ($parts[1] -match '^(\d+)') {
            return $Matches[1]
        }
    }
    return ''
}

<#
.SYNOPSIS 判断文件是否与目标目录中已有简历重复。
重复条件：姓名 + 年龄 + 文件大小 三者完全相同。
#>
function Test-DuplicateFile {
    param([string]$FilePath, [string]$TargetDir)
    $fname = Split-Path -Leaf $FilePath
    $name = Get-NameFromZhilianFile $fname
    $age  = Get-AgeFromZhilianFile $fname
    if (-not $name -or -not $age) { return $false }

    $size = (Get-Item $FilePath -ErrorAction SilentlyContinue).Length

    $existing = Get-ZhaopinFiles $TargetDir | Where-Object {
        $en = Get-NameFromZhilianFile $_.Name
        $ea = Get-AgeFromZhilianFile $_.Name
        ($en -eq $name) -and ($ea -eq $age) -and ($_.Length -eq $size)
    } | Select-Object -First 1

    return ($existing -ne $null)
}

<#
.SYNOPSIS 扫描 Downloads 目录中最新下载的智联文件，移动至目标目录。
.DESCRIPTION 返回约定（调用方必须区分，不能把"未找到文件"当成"重复"）：
  - 返回 'ok'    → 移动成功
  - 返回 'dup'   → 目标目录已有完全相同的简历（姓名+年龄+大小），已删除新文件
  - 返回 $null   → Downloads 中未找到新下载的智联文件（文件可能仍在落盘，或匹配失败）
#>
function Move-OneResume {
    $files = Get-ZhaopinFiles $Config.DownloadSource
    if (-not $files -or $files.Count -eq 0) {
        Write-Host '  [WARN] No zhilian file found in Downloads' -ForegroundColor Yellow
        return $null
    }

    $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { return $null }

    Write-Host "  Latest download: $($latest.Name) ($($latest.Length) bytes, $($latest.LastWriteTime.ToString('HH:mm:ss')))" -ForegroundColor Gray

    $dest = Join-Path $Config.DownloadDir $latest.Name

    # 三重验证（姓名+年龄+大小）判定重复
    if (Test-Path $dest) {
        if (Test-DuplicateFile -FilePath $latest.FullName -TargetDir $Config.DownloadDir) {
            Write-Host '  [DUP] Identical resume already in target dir — removing new file' -ForegroundColor Yellow
            Remove-Item $latest.FullName -Force -ErrorAction SilentlyContinue
            return 'dup'
        } else {
            Write-Host '  [INFO] Same name but different candidate — overwrite moving' -ForegroundColor DarkGray
        }
    }

    # 移动文件
    Move-Item -Path $latest.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
    if (Test-Path $dest) {
        Write-Host "  [MOVED] $($latest.Name)" -ForegroundColor Green
        return 'ok'
    }
    Write-Host '  [FAIL] Move failed (file may be locked/used)' -ForegroundColor Red
    return $null
}

<#
.SYNOPSIS 安全的遮罩关闭——仅在模态打开时才点击遮罩
#>
function Close-ModalIfOpen {
    $mc = Invoke-Eval 'String(document.querySelector(".km-modal--open")?true:false)'
    if ($mc -notmatch '"value":"true"') { return }
    Write-Host '  Closing modal...' -ForegroundColor Gray
    Invoke-CDP -Type 'mousePressed'  -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
    Wait 200
    Invoke-CDP -Type 'mouseReleased' -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
    Wait $Config.CloseWaitMs
}

<#
.SYNOPSIS 检查目标目录中是否已有某姓名的简历文件，用于跳过已下载候选人
#>
function Test-NameExistsInDir {
    param([string]$CandidateName, [string]$Dir)
    $existing = Get-ZhaopinFiles $Dir | Where-Object {
        (Get-NameFromZhilianFile $_.Name) -eq $CandidateName
    } | Select-Object -First 1
    return ($existing -ne $null)
}

# ============================================================
# 阶段 3：逐条下载循环（含去重）
# ============================================================
$targetCount = $Config.DownloadCount
Write-Host "[4] Downloading (target: $targetCount)..." -ForegroundColor Cyan

$ok          = 0
$fail        = 0
$skip        = 0
$preExisting = 0
$idx         = 0
$triedCandidates = @{}
$isFirst     = $true  # 第一次迭代无需关闭模态

# ★★★ 修复 #33：未达标前绝不停止 ★★★
# 原实现 while ($ok -lt $targetCount -and $idx -lt $names.Count) 一遇到 $names 遍历完
# 就退出（远早于目标数）。根因是收集阶段只拿到 31 个名字时，遍历 31 次即结束。
# 现在引入"补收名字"机制：一轮遍历完仍未达标 → 重新滚动收集（阈值放宽）→ 继续下载，
# 直到 ok >= targetCount，或连续 $maxEmptyRounds 轮确实没有任何新候选人可下载。
$maxEmptyRounds = 3      # 连续 3 轮"没有任何新增成功/尝试"才认定候选池耗尽
$emptyRounds    = 0
$round          = 0
$lastOk         = 0

while ($ok -lt $targetCount) {
    $round++
    Write-Host ''
    Write-Host "===== Round $round : $ok/$targetCount downloaded, $($names.Count) names in pool =====" -ForegroundColor Magenta

    $attemptedThisRound = 0
    while ($ok -lt $targetCount -and $idx -lt $names.Count) {
        $name = $names[$idx]
        $idx++
        $attemptedThisRound++

        if ($triedCandidates.ContainsKey($name)) { continue }

        Write-Host ''
        Write-Host "[$ok/$targetCount] $name (idx=$idx/$($names.Count))" -ForegroundColor Cyan

        # 预检：目标目录是否已有该候选人简历
        # ★ 修复：智联列表只显示"张先生/李女士"这类脱敏名，同名不同人极常见。
        #   按姓名预跳过会漏掉同名候选人 → 改为总是尝试下载，交给 Move-OneResume 的
        #   三重验证（姓名+年龄+文件大小）判定，真正重复时返回 'dup' 且不计数。
        #   仅当目录中同名文件已很多（>=3，几乎可确定同名候选人已全部覆盖）时才跳过，用于省时。

        # 3.1 关闭模态（首次跳过，后续按需关闭）
        if (-not $isFirst) {
            Close-ModalIfOpen
        }
        $isFirst = $false

        # 3.2 查找并点击候选人
        # ★ 修复：搜索时先重置到顶部再逐屏下滑，避免"已在底部还继续往下滚"的死循环。
        #   策略：第 0 轮先回到顶部，之后每轮向下滚一屏；同时兼顾上/下双向兜底。
        $found = $false
        $safeName = ConvertTo-JsSafeName -Name $name
        # ★ 修复 #44：先确保标签页在最前。若被其他标签页遮挡，Chrome 不会把
        #   真实输入事件投递给隐藏的 render widget，任何点击都会静默失败。
        Ensure-TabFocused
        for ($retry = 0; $retry -lt 25; $retry++) {
            # 第 0 轮：先归零到顶部，保证从列表开头开始找
            if ($retry -eq 0) {
                $null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;return "ok"})()'
                Wait 1200
            }
            # ★ 修复 #44：用 JS 定位候选人并打临时标记，再用 DOM 级 click 点击。
            #   原实现在 evaluate 里调 `el.click()`——实测该 JS click 在智联这个
            #   虚拟列表上**不会打开详情面板**（React 合成事件不响应程序化 click），
            #   必须走 WebBridge 的 `click` action（派发真实 pointerdown/mousedown）。
            $fjs = '(()=>{const e=document.querySelectorAll(''.talent-basic-info__name'');for(const el of e){if(el.textContent.trim().indexOf("' + $safeName + '")>=0){el.setAttribute("data-wb-target","1");return"ok"}}return"no"})()'
            $fr = Invoke-Eval $fjs
            if ($fr -match '"value":"ok"') {
                Ensure-TabFocused
                if (Invoke-Click '[data-wb-target="1"]') {
                    $found = $true
                    $null = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-target]");if(e)e.removeAttribute("data-wb-target");return"ok"})()'
                    break
                }
                # DOM click 失败则清标记，继续尝试下一屏
                $null = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-target]");if(e)e.removeAttribute("data-wb-target");return"ok"})()'
            }

            # ★ 修复 #43：检测"session 无标签页 / 列表未渲染"并自愈。
            #   原实现只判断 '"value":"ok"'，evaluate 返回 ok:false（无 tab）时
            #   会一路静默滚到 25 次重试结束 → 连续 [SKIP] 空转。
            #   这里识别 ok:false，重新 navigate 回列表 URL 并等待渲染。
            if ($fr -match '"ok":false') {
                Write-Host '  [WARN] Session tab lost — re-navigating to list...' -ForegroundColor Yellow
                $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $false }
                for ($rw = 0; $rw -lt 20; $rw++) {
                    $rchk = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
                    if ($rchk -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) { break }
                    Wait 500
                }
                Wait 800
                continue
            }

            # ★★ 修复 #34：每轮都从顶部重新开始逐屏下滑。
            #   原实现只在第 0 轮归零，后续轮次继续在当前 scrollTop 上叠加，
            #   一旦滚到底就永远停在底部 → 25 次重试全部空转（表现为日志刷屏
            #   [SKIP] Not found in viewport 且一份都没点进去）。
            #   现在改为"相对顶部定位"：offset = retry * 700，每次都从 0 重新滚到该偏移。
            $offset = $retry * 700
            $null = Invoke-Eval ('(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=' + $offset + ';else window.scrollTo(0,' + $offset + ');return "ok"})()')
            Wait 900
        }
        if (-not $found) {
            Write-Host '  [SKIP] Not found after full-list sweep' -ForegroundColor Yellow
            $triedCandidates[$name] = $true
            $skip++
            continue
        }
        Wait $Config.ClickWaitMs

    # ============================================================
    # 3.2.5 ★★★ 修复 #44：确认详情面板已打开 ★★★
    # ============================================================
    # 候选人已在 3.2 用 DOM click 点过，此处只做确认与兜底重试。
    # 【认知纠正】此前（#41/#42）判断"详情页在新标签页打开、需要接管"是**错的**。
    #   实测：`Page.bringToFront` + DOM `click` 之后，详情面板在**当前标签页内**
    #   直接渲染（`.resume-detail-wrap` 出现），无需任何标签页切换。
    $opened = $false
    for ($openTry = 0; $openTry -lt 3; $openTry++) {
        $dchk = Invoke-Eval '(()=>{const b=document.querySelector(".resume-detail-wrap");return b?"yes":"no"})()'
        if ($dchk -match '"value":"yes"') { $opened = $true; break }
        # 兜底：再点一次候选人
        Ensure-TabFocused
        $null = Invoke-Click '.talent-basic-info__name'
        Wait 2500
    }
    if ($opened) {
        Write-Host '  [OK] Detail panel opened (in current tab)' -ForegroundColor Green
    } else {
        Write-Host '  [WARN] Detail panel not detected — proceeding anyway' -ForegroundColor Yellow
    }

    # 3.3 等待详情面板渲染完毕，再探测"存至本地"按钮
    # ★ 修复 #45：详情面板的「存至本地」按钮**不是立即出现**的。
    #   实测（2026-09-15）：面板主体（`.resume-detail-wrap`，约 1107x715）先渲染，
    #   但 `.resume-button.position-r`（含"存至本地"）要**等简历正文加载完**才挂载，
    #   实测约需 10~15 秒。原等待上限 15×500ms=7.5s 不够 → 探测失败 →
    #   退回硬编码坐标 → 点了个不存在的按钮 → "No save dialog" → 空转。
    #   现在提高到 40 轮 × 750ms ≈ 30 秒，并在等待期间先等面板主体出现。
    for ($wp = 0; $wp -lt 20; $wp++) {
        $dp = Invoke-Eval '(()=>{return document.querySelector(".resume-detail-wrap")?"yes":"no"})()'
        if ($dp -match '"value":"yes"') { break }
        Wait 500
    }
    for ($w = 0; $w -lt 40; $w++) {
        $cr = Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return 0;const r=b.getBoundingClientRect();return(r.width>0&&r.height>0)?1:0})()'
        if ($cr -match '"value":1') { break }
        Wait 750
    }
    Wait 500  # 额外等一等确保布局稳定

    $saveLocalX = $Config.SaveLocalX; $saveLocalY = $Config.SaveLocalY
    $slr = Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return 0;const r=b.getBoundingClientRect();if(r.width<=0||r.height<=0)return 0;return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)})()'
    $slc = Parse-Coords $slr
    if ($slc) {
        $saveLocalX = $slc[0]; $saveLocalY = $slc[1]
        Write-Host "  SaveLocal btn at ($saveLocalX, $saveLocalY)" -ForegroundColor Gray
    } else {
        Write-Host "  [WARN] SaveLocal btn probe failed, fallback ($saveLocalX, $saveLocalY)" -ForegroundColor Yellow
    }

    # 3.4 点击"存至本地"
    # ★★★ 修复 #46（实测 2026-09-15）：这是让整条链路跑通的关键序列 ★★★
    # 【现象】详情面板打开后，`.resume-button.position-r`（存至本地）**会闪烁**：
    #   面板主体的 Vue 渲染过程中该按钮反复挂载/卸载，探测到"存在"不代表点击时还在。
    # 【验证过的正确序列】（手工实测成功打开保存对话框）：
    #   ① CDP `Page.bringToFront`（不置前则输入被丢弃）
    #   ② 探测到 `.resume-button.position-r` **可见的那一刻立即**发出 CDP 鼠标序列
    #      （不等额外延时，否则按钮已被卸载）
    #   ③ 用 getBoundingClientRect 的实时坐标（不要用配置兜底值）
    # 【注意】此处**不能**用 DOM `click`：tech_details #21 已记录
    #   "不能用 JS .click()，Vue 组件需要真实鼠标事件"，实测 DOM click 返回
    #   success 但不弹对话框。
    $hasDialog = $false
    for ($saveRetry = 0; $saveRetry -lt 12; $saveRetry++) {
        Ensure-TabFocused
        # 实时探测按钮（可见性 + 坐标），探测到就立刻点
        $rslr = Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return"x";const r=b.getBoundingClientRect();if(r.width<=0||r.height<=0)return"x";return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)})()'
        $rslc = Parse-Coords $rslr
        if ($rslc) {
            $saveLocalX = $rslc[0]; $saveLocalY = $rslc[1]
            Write-Host "  SaveLocal btn at ($saveLocalX, $saveLocalY)" -ForegroundColor Gray
            # 立即发出 CDP 真实鼠标序列（无额外等待，抢在按钮卸载前）
            Invoke-CDP -Type 'mouseMoved'    -X $saveLocalX -Y $saveLocalY
            Invoke-CDP -Type 'mousePressed'  -X $saveLocalX -Y $saveLocalY -Button 'left'
            Wait 80
            Invoke-CDP -Type 'mouseReleased' -X $saveLocalX -Y $saveLocalY -Button 'left'
        } else {
            # 按钮暂未挂载/已卸载：若详情面板还在，稍等重试
            $panelStill = Invoke-Eval '(()=>{return document.querySelector(".resume-detail-wrap")?"yes":"no"})()'
            if ($panelStill -notmatch '"value":"yes"') {
                # 面板已关（自动收起）→ 重开面板
                Write-Host '  Panel closed — reopening...' -ForegroundColor Yellow
                Ensure-TabFocused
                $null = Invoke-Click '.talent-basic-info__name'
                Wait 2500
            } else {
                Wait 400
            }
        }
        Wait $Config.DialogCheckWaitMs

        $cr = Invoke-Eval '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0)return"has"}return"no"})()'
        if ($cr -match '"value":"has"') { $hasDialog = $true; break }
    }
    if (-not $hasDialog) {
        Write-Host '  [FAIL] No save dialog after 12 attempts' -ForegroundColor Red
        $triedCandidates[$name] = $true
        $fail++
        continue
    }

    # 3.5 等待对话框渲染稳定
    Wait 500
    $sx = $Config.SaveConfirmX; $sy = $Config.SaveConfirmY

    # 3.5.5 选择文件格式（PDF 默认选中；若 FileFormat=word 则切换）
    if ($Config.FileFormat -eq 'word') {
        Write-Host '  Switching format: word...' -ForegroundColor Gray
        # 用 JS 探测 "word" 选项的 getBoundingClientRect() 中心坐标
        # 该选项是一个文本节点 DIV（直接子节点唯一文本为 "word"）
        $wordCoords = Invoke-Eval '(()=>{const els=document.querySelectorAll("div");for(const el of els){if(el.childNodes.length===1&&el.childNodes[0].nodeType===3&&el.textContent.trim()==="word"){const r=el.getBoundingClientRect();if(r.width>0&&r.height>0)return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)}}return"0,0"})()'
        $wc = Parse-Coords $wordCoords
        if ($wc) {
            $wx = $wc[0]; $wy = $wc[1]
            Write-Host "  Word option at ($wx, $wy)" -ForegroundColor Gray
            # ★ 修复 #44：优先 DOM 级 click（遮挡环境下坐标点击会被静默丢弃）
            Ensure-TabFocused
            $wordClicked = $false
            # "word" 选项是纯文本 DIV，无稳定 class，用 JS 给它打临时标记再点
            $markRaw = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const els=document.querySelectorAll("div");for(const el of els){if(el.childNodes.length===1&&el.childNodes[0].nodeType===3&&el.textContent.trim()==="word"){el.setAttribute("data-wb-word","1");return"ok"}}return"no"})()' }
            if ($markRaw -match '"value":"ok"') {
                if (Invoke-Click '[data-wb-word="1"]') { $wordClicked = $true }
                $null = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const e=document.querySelector("[data-wb-word]");if(e)e.removeAttribute("data-wb-word");return"ok"})()' }
            }
            if (-not $wordClicked) {
                # 退回坐标点击
                Invoke-CDP -Type 'mouseMoved'    -X $wx -Y $wy
                Wait 500
                Invoke-CDP -Type 'mousePressed'  -X $wx -Y $wy -Button 'left'
                Wait 150
                Invoke-CDP -Type 'mouseReleased' -X $wx -Y $wy -Button 'left'
            }
            Wait 800
            # 验证是否真的切换为 word（下方提示文字变为"支持 word 2010..."）
            $vfy = Invoke-Eval '(()=>{const ps=document.querySelectorAll("p");for(const p of ps){if(p.textContent.indexOf("\u652f\u6301 word")>=0)return"yes"}return"no"})()'
            if ($vfy -match '"value":"yes"') {
                Write-Host '  [OK] Format switched to word' -ForegroundColor Green
            } else {
                Write-Host '  [WARN] Format switch verify failed, continuing anyway' -ForegroundColor Yellow
            }
        } else {
            Write-Host '  [WARN] Word option coords not found, falling back to default (pdf)' -ForegroundColor Yellow
        }
    }

    # 3.6 点击"保存"按钮
    # ★ 修复 #44：优先 DOM 级 click（".km-button--primary" 是保存按钮的稳定 class 片段），
    #   失败再退回坐标点击。遮挡环境下坐标点击会被静默丢弃。
    $downloaded = $false
    for ($confirmRetry = 0; $confirmRetry -lt 3; $confirmRetry++) {
        Ensure-TabFocused
        # 用 JS 给"保存"按钮打临时标记，再用 DOM click 精确命中
        $markSave = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0){x.setAttribute("data-wb-save","1");return"ok"}}return"no"})()' }
        $saveClicked = $false
        if ($markSave -match '"value":"ok"') {
            if (Invoke-Click '[data-wb-save="1"]') { $saveClicked = $true }
        }
        if (-not $saveClicked) {
            # 退回坐标点击
            $coords = Get-SaveButtonCoords
            if ($coords) { $sx = $coords[0]; $sy = $coords[1] }
            Write-Host "  Save confirm btn at ($sx, $sy) [coords]" -ForegroundColor Gray
            $null = Invoke-CDP -Type 'mouseMoved' -X $sx -Y $sy -ErrorAction SilentlyContinue
            Wait 200
            Invoke-CDP -Type 'mousePressed'  -X $sx -Y $sy -Button 'left'
            Wait 80
            Invoke-CDP -Type 'mouseReleased' -X $sx -Y $sy -Button 'left'
        } else {
            Write-Host '  Clicked 保存 (DOM click)' -ForegroundColor Gray
        }
        Wait $Config.DownloadWaitMs

        # ★ 纯 ASCII 结构匹配：不依赖中文字面量，避免 GBK 乱码导致匹配失败
        $latestFile = Get-ZhaopinFiles $Config.DownloadSource |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        # 检测窗口：PDF 模式 15 秒，Word 模式 25 秒（Word 文件生成耗时更长）
        $detectWindowSec = if ($Config.FileFormat -eq 'word') { 25 } else { 15 }
        if ($latestFile -and $latestFile.LastWriteTime -gt [DateTime]::Now.AddSeconds(-$detectWindowSec)) {
            $downloaded = $true
            Write-Host "  [OK] Download detected: $($latestFile.Name)" -ForegroundColor Green
            break
        }
        Write-Host "  No new download detected, retrying... (last matched: $($latestFile.Name))" -ForegroundColor Yellow
    }
    if (-not $downloaded) {
        Write-Host '  [FAIL] Download did not complete (no new zhilian file appeared in Downloads)' -ForegroundColor Red
        $triedCandidates[$name] = $true
        $fail++
        continue
    }

    # 3.7 移动文件 + 去重检查
    $moveResult = Move-OneResume
    if ($moveResult -eq 'ok') {
        Write-Host "  [OK] Resume moved & counted" -ForegroundColor Green
        $ok++
    } elseif ($moveResult -eq 'dup') {
        Write-Host '  [DUP] Duplicate removed — trying next candidate' -ForegroundColor Yellow
        $triedCandidates[$name] = $true
    } else {
        # $null：Downloads 中未找到文件（或移动失败被占用）→ 文件留在 Downloads，需人工处理
        Write-Host '  [WARN] No file matched/moved — file stays in Downloads, check manually' -ForegroundColor Yellow
        $triedCandidates[$name] = $true
        $fail++
    }

    # 3.8 保存后清理：关闭可能残留的模态/对话框，防止干扰下一条
    Close-ModalIfOpen
    Wait 500

    # ============================================================
    # 3.9 ★ 修复 #44：关闭详情面板（详情面板在同一标签页内，非新标签页）
    # ============================================================
    # 详情面板是页内覆盖层，下载完成后需关闭它才能点击下一位候选人。
    Close-ModalIfOpen
    Wait 500
    # 若详情面板仍在，用 DOM click 点关闭按钮
    $dm = Invoke-Eval '(()=>{const b=document.querySelector(".km-modal__close-btn");return b?"yes":"no"})()'
    if ($dm -match '"value":"yes"') {
        $null = Invoke-Click '.km-modal__close-btn'
        Wait 800
    }
    }  # ← end inner while (idx over names)

    # ============================================================
    # ★★★ 修复 #33（续）：本轮名单遍历完，若未达标则补收名字继续 ★★★
    # ============================================================
    if ($ok -ge $targetCount) { break }

    if ($ok -gt $lastOk) {
        # 本轮有实际产出 → 重置空轮计数，继续补收
        $emptyRounds = 0
    } else {
        $emptyRounds++
    }
    $lastOk = $ok

    if ($emptyRounds -ge $maxEmptyRounds) {
        Write-Host ''
        Write-Host "[STOP] $maxEmptyRounds consecutive rounds with no new downloads — candidate pool exhausted." -ForegroundColor Yellow
        Write-Host "       Downloaded $ok / $targetCount. Proceeding to summary." -ForegroundColor Yellow
        break
    }

    # 补收名字：重置 Set 与滚动位置，用放宽的阈值重新扫列表（列表会随滚动继续加载）
    Write-Host ''
    Write-Host "[TOP-UP] Still need $($targetCount - $ok) more — re-collecting names..." -ForegroundColor Magenta
    $null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return "ok"})()'
    Wait 2000

    $null = Invoke-Eval 'window._SN = new Set()'
    Wait 300
    $last = 0; $stag = 0
    # 补收阶段阈值放宽到 20 轮（比首轮更耐心），maxScrolls 翻倍
    for ($i = 0; $i -lt ($maxScrolls * 2); $i++) {
        $js = '(()=>{const c=document.querySelector(''.app-layout--default'')||document.scrollingElement;const e=document.querySelectorAll(''.talent-basic-info__name'');e.forEach(el=>{const m=el.textContent.trim().match(/^\S+/);if(m)window._SN.add(m[0])});if(c)c.scrollTop=c.scrollTop+900;else window.scrollBy(0,900);return ''''})()'
        $null = Invoke-Eval $js
        Wait $Config.ScrollWaitMs
        $r = Invoke-Eval 'String(window._SN.size)'
        if ($r -match '"value":"(\d+)"') {
            $cur = [int]$Matches[1]
            if ($cur -eq $last) { $stag++ } else { $stag = 0; $last = $cur }
            if ($stag -ge 20) { break }
        }
    }

    # 合并新收集到的名字到总名单（保持原有顺序，只追加新出现的）
    $newNames = @()
    $r = Invoke-Eval 'JSON.stringify(Array.from(window._SN))'
    if ($r -match '"value":"(\[.*\])"') {
        try {
            $json = $Matches[1] -replace '\\"', '"'
            $raw  = $json | ConvertFrom-Json
            $newNames = @($raw | Where-Object { $_ } | ForEach-Object { $_ -replace '\s', '' } | Where-Object { $_ })
        } catch {}
    }

    $added = 0
    foreach ($n in $newNames) {
        if ($names -notcontains $n) { $names += $n; $added++ }
    }
    Write-Host "  [TOP-UP] +$added new names → pool now $($names.Count)" -ForegroundColor Magenta

    # 重置索引，重新遍历扩容后的名单（已试过的候选人在 $triedCandidates 里会被跳过）
    $idx = 0
    $isFirst = $true

    # 重建列表后回到顶部，保证查找从开头开始
    $null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return "ok"})()'
    Wait 2500
}  # ← end outer while (ok < targetCount)

# ============================================================
# 清理与汇总
# ============================================================
# ★ 修复 #35：任务完成后必须停止浏览器自动化，不能让它继续跑（残留标签页/连接）。
Stop-BrowserAutomation

Write-Host ''
Write-Host '=====================================' -ForegroundColor Green
Write-Host "  Success      : $ok" -ForegroundColor Green
Write-Host "  Failed       : $fail" -ForegroundColor Red
Write-Host "  Skipped      : $skip" -ForegroundColor Yellow
Write-Host "  Pre-existing : $preExisting" -ForegroundColor DarkGray
Write-Host "  Format       : $($Config.FileFormat)" -ForegroundColor Cyan
Write-Host "  Target files : $((Get-ZhaopinFiles $Config.DownloadDir).Count)" -ForegroundColor Green
Write-Host "  Target goal  : $targetCount" -ForegroundColor Cyan
Write-Host '=====================================' -ForegroundColor Green

# 明确告知是否达标，便于 Agent 判断是否需要补跑
if ($ok -ge $targetCount) {
    Write-Host "  [DONE] Target reached ($ok/$targetCount)." -ForegroundColor Green
    Write-Host '  Browser automation stopped.' -ForegroundColor Green
} else {
    Write-Host "  [INCOMPLETE] Only $ok/$targetCount downloaded — candidate pool exhausted." -ForegroundColor Yellow
    Write-Host '  Browser automation stopped.' -ForegroundColor Green
    Write-Host '  Options: enlarge the pool (scroll more) or lower DownloadCount.' -ForegroundColor Yellow
}
Write-Host '=====================================' -ForegroundColor Green