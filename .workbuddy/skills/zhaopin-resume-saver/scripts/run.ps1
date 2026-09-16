# ============================================================
# 智联招聘简历批量下载 — 主编排（模块化重构版 2026-09-15）
# ============================================================
# 架构：run.ps1（本文件，薄编排）+ lib/wb-core.ps1（客户端/环境/原语）
#       + lib/zhaopin-page.ps1（页面交互函数库）
# 流程细节与踩坑原因见 lib 内注释与 SKILL.md 问题记录。
#
# 用法：
#   方式一：编辑 config.json 后直接运行 .\run.ps1
#   方式二：参数覆盖  .\run.ps1 -DownloadCount 20 -DownloadDir "C:\Resumes"
#
# 退出码：0=达标完成  2=未达标（候选池耗尽）  1=参数/岗位致命  3=环境致命
#         4=停滞（卡住看门狗触发，wrapper 检测到后自动重启续传，#62）
# 产物：<DownloadDir>\_summary.json（机器可读结果）+ _run_log.txt（结构化日志）
# ============================================================
param(
    [string]$Url,
    [string]$JobName,
    [string]$DownloadDir,
    [int]$DownloadCount = 0,
    [ValidateSet('pdf','word')]
    [string]$FileFormat = ''
)

# ★ 强制 UTF-8 输出编码（Agent 后台任务环境中文乱码防护）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = 'Continue'
$Utf8NoBom = New-Object System.Text.UTF8Encoding $false

# ============================================================
# 配置加载与校验
# ============================================================
$sd = Split-Path -Parent $MyInvocation.MyCommand.Path
$jsonPath = Join-Path $sd 'config.json'
if (-not (Test-Path $jsonPath)) {
    Write-Host 'ERROR: config.json not found!' -ForegroundColor Red
    exit 1
}
$jsonConfig = Get-Content -Raw -Encoding UTF8 $jsonPath | ConvertFrom-Json

$Config = @{
    Url           = $jsonConfig.Url
    JobName       = $jsonConfig.JobName
    DownloadDir   = $jsonConfig.DownloadDir
    DownloadCount = [int]$jsonConfig.DownloadCount
    FileFormat    = $jsonConfig.FileFormat

    WebBridgeUrl  = 'http://127.0.0.1:10086/command'
    Session       = 'resume-screening'
    DownloadSource = "$env:USERPROFILE\Downloads"
    DownloadFilter = '*智联简历*'

    # 坐标兜底值（1920x1080；实际以 getBoundingClientRect 实时探测为准）
    SaveLocalX   = 1079
    SaveLocalY   = 142
    SaveConfirmX = 996
    SaveConfirmY = 561
    MaskCloseX   = 30
    MaskCloseY   = 300

    # 时序参数（多轮踩坑调优值，勿随意压缩——原因见 SKILL.md）
    CloseWaitMs         = 700    # 遮罩点击释放后等模态关闭
    ClickWaitMs         = 800    # 点中卡片后等面板开始渲染
    DialogCheckWaitMs   = 2000   # 存至本地后查保存对话框
    ScrollWaitMs        = 800    # 列表下滚一屏等待（#53 用户指定）
    SaveCloseWaitMs     = 3200   # #57 用户指定：release 保存按钮后关面板前的等待
    ModalRetryWaitMs    = 1000   # #57 用户指定：模态未关闭时的关闭重试间隔
    ModalCloseRetryMax  = 10     # #57 关闭重试上限（防死循环，超限 WARN 后继续）
    FileDetectWordSec   = 25     # word 落盘轮询窗口
    FileDetectPdfSec    = 15     # pdf 落盘轮询窗口
    StallTimeoutSec     = 300    # #62：无任何卡片处理进展超过此秒数 → 判定卡住 exit 4
}
if ($Url)                 { $Config.Url          = $Url }
if ($JobName)             { $Config.JobName      = $JobName }
if ($DownloadDir)         { $Config.DownloadDir  = $DownloadDir }
if ($DownloadCount -gt 0) { $Config.DownloadCount = $DownloadCount }
if ($FileFormat)          { $Config.FileFormat   = $FileFormat.ToLower() }
$Config.FileExt = if ($Config.FileFormat -eq 'word') { 'docx' } else { 'pdf' }

# 强制参数校验
$missing = @()
if (-not $Config.Url)         { $missing += 'Url' }
if (-not $Config.JobName)     { $missing += 'JobName' }
if (-not $Config.DownloadDir) { $missing += 'DownloadDir' }
if ($missing.Count -gt 0) {
    Write-Host "MISSING REQUIRED PARAMETERS: $($missing -join ', ')" -ForegroundColor Red
    Write-Host 'Usage: .\run.ps1 -Url "https://rd6.zhaopin.com/app/recommend?jobNumber=XXX" -JobName "岗位名" -DownloadDir "C:\Resumes" -DownloadCount 30' -ForegroundColor Gray
    exit 1
}

# 确保目标目录存在 + 初始化结构化日志
if (-not (Test-Path $Config.DownloadDir)) {
    $null = New-Item -ItemType Directory -Path $Config.DownloadDir -Force
}

$Session = 'resume-screening'

# ============================================================
# 加载函数库（必须在 $Config/$Session/$Utf8NoBom 定义之后）
# ============================================================
. (Join-Path $sd 'lib\wb-core.ps1')
Initialize-LogFile -Directory $Config.DownloadDir
. (Join-Path $sd 'lib\zhaopin-page.ps1')

# #62 断点续传自动折算：config 提供 DownloadTarget（跨重启恒定的总目标）时，
# 本次实际目标 = 总目标 - 目录已有份数。中断/自动重启后无需手工改 DownloadCount。
if ($DownloadCount -eq 0 -and $jsonConfig.DownloadTarget) {
    $existing = (Get-ZhaopinFiles $Config.DownloadDir).Count
    $Config.DownloadCount = [Math]::Max(0, [int]$jsonConfig.DownloadTarget - $existing)
    Write-Log "DownloadTarget=$($jsonConfig.DownloadTarget) existing=$existing -> this-run target=$($Config.DownloadCount)" -Level INFO
}

# 兼容清理：历史遗留的固定名临时请求文件
Remove-Item (Join-Path $env:TEMP 'wb-download.json') -Force -ErrorAction SilentlyContinue

$Script:AutomationStopped = $false
$EffectiveJobNumber = ''   # #61：岗位卡选好后从 location.href 读出的实际 jobNumber
Write-Log "=== zhaopin resume download start ===" -Level STEP
Write-Log "job=[$($Config.JobName)] dir=[$($Config.DownloadDir)] target=$($Config.DownloadCount) format=$($Config.FileFormat)"

# ============================================================
# [0] 环境自检（daemon + 扩展真探测；失败 exit 3）
# ============================================================
if (-not (Initialize-WebBridgeEnv)) {
    Write-Log 'environment check failed - aborting' -Level FAIL
    exit 3
}

# ============================================================
# [1] 导航
# ============================================================
Write-Log 'navigating to recommend page...' -Level STEP
# newTab=$true 确保创建新标签页并激活为 current tab
$null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
Wait 3000
# bringToFront 强制浏览器 UI 切到 CDP 连接的标签页（不能用 find_tab/Target.activateTarget）
Ensure-TabFocused
Wait 3000

# ============================================================
# [2] 岗位选择 + 验证（#61 增强：点击必须"验证生效"才算选岗成功）
# ============================================================
Write-Log "selecting job: $($Config.JobName)" -Level STEP
if (-not (Select-JobTab -JobName $Config.JobName)) {
    Write-Log 'dumping all job tabs on page for diagnosis:' -Level WARN
    $tabs = Get-JobTabDump
    Write-Log ("job tabs on page: " + ($tabs -join ' | ')) -Level WARN
    Write-Log "FATAL: job tab not found OR click never took effect after 15 retries. Check JobName spelling / jobNumber correspondence / slow render." -Level FAIL
    Stop-BrowserAutomation -Quiet
    exit 1
}
Wait 3000

# #61：记录点击选岗后的实际 URL jobNumber（点击正确岗位卡会纠正错误的 config Url，
# 这里只记录用于溯源，不作为失败依据——岗位名验证才是判定标准）
$u = Invoke-Eval 'location.href'
if ($u -match 'jobNumber=([A-Za-z0-9]+)') {
    $EffectiveJobNumber = $Matches[1]
    Write-Log "effective jobNumber after selection: $EffectiveJobNumber" -Level INFO
}

# 候选人列表加载
Write-Log 'waiting for candidate list...' -Level INFO
$null = Wait-Until -TimeoutMs 10000 -PollMs 1000 -Description 'candidate list load' -Condition {
    (Get-VisibleCardCount) -gt 0
}

# 岗位验证（#58：探测失败=无法证明岗位正确=FATAL，除非 config 显式 AllowUnverifiedJob）
Write-Log 'verifying active job...' -Level STEP
$activeJob = Get-ActiveJobName
if ($activeJob) {
    Write-Log "active job=[$activeJob] expected=[$($Config.JobName)]" -Level INFO
    $activeClean   = $activeJob.Trim()
    $expectedClean = $Config.JobName.Trim()
    $activeNoParen   = $activeClean   -replace '\s*\(.*$', ''
    $expectedNoParen = $expectedClean -replace '\s*\(.*$', ''
    if ($activeClean -eq $expectedClean -or $activeNoParen -eq $expectedNoParen -or $activeClean -like "*$expectedClean*") {
        Write-Log 'active job matches expected - proceeding' -Level OK
    } else {
        Write-Log "FATAL: job mismatch! expected=[$expectedClean] active=[$activeClean]" -Level FAIL
        Stop-BrowserAutomation -Quiet
        exit 1
    }
} else {
    # #58 修复（问题1）：探测失败曾降级为 WARN 继续 → 曾导致用联想渠道经理岗位的
    # 推荐池按"AI产品经理"指令下载了错误简历。现在探测失败=无法证明岗位正确=FATAL，
    # 除非 config.json 显式设置 "AllowUnverifiedJob": true 豁免
    if ($Config.AllowUnverifiedJob -eq $true) {
        Write-Log 'could not detect active job - AllowUnverifiedJob=true, proceeding with caution' -Level WARN
    } else {
        Write-Log "FATAL: could not verify active job (probe failed) - refusing to risk wrong-job download. Fix page/login state or set config AllowUnverifiedJob=true to override." -Level FAIL
        Stop-BrowserAutomation -Quiet
        exit 1
    }
}

# ============================================================
# [3] 列表就绪 + 视口健康检查
# ============================================================
Write-Log 'waiting for list container to render...' -Level INFO
# #39：ok:false（session 无标签页）→ 重新导航自愈，不静默空转
$listReady = $false
for ($w = 0; $w -lt 20; $w++) {
    $chk = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
    if ($chk -match '"ok":false') {
        Write-Log 'session has no tab yet - re-navigating...' -Level WARN
        $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
        Wait 3000
        Ensure-TabFocused
        Wait 2000
        continue
    }
    if ($chk -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) {
        Write-Log "list ready ($($Matches[1]) nodes in viewport)" -Level OK
        $listReady = $true
        break
    }
    Wait 1000
}
if (-not $listReady) {
    Write-Log 'list container not confirmed - proceeding anyway (main loop will self-heal)' -Level WARN
}

$visNames = 0
for ($vi = 0; $vi -lt 3; $vi++) {
    $visNames = Get-VisibleCardCount
    if ($visNames -gt 0) { break }
    if ($visNames -lt 0) {
        Write-Log 'session tab lost during health check - re-navigating...' -Level WARN
        $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
        Wait 4000
        Ensure-TabFocused
        Wait 2000
    } else { Wait 1000 }
}
Write-Log "visible cards in viewport: $visNames" -Level INFO
if ($visNames -eq 0) {
    Write-Log 'FATAL: no candidate cards in viewport - abort' -Level FAIL
    Stop-BrowserAutomation -Quiet -KeepDaemon   # 保留 daemon 便于立即重试
    exit 2
}

# 收集阶段不再滚动（#53：下载循环顺序推进自取），但重置到顶部保证从头开始
$null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return "ok"})()'
Wait 2500

# 清理历史残留详情标签页（#42/#43：绝不关 active 标签页）
$staleClosed = Clear-StaleDetailTabs
if ($staleClosed -gt 0) {
    Write-Log "closed $staleClosed stale detail tab(s)" -Level OK
    Wait 800
}
# 兜底导航回列表页
$null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $false }
$null = Wait-Until -TimeoutMs 10000 -PollMs 500 -Description 'back-to-list render' -Condition { (Get-VisibleCardCount) -gt 0 }
Wait 1000

# ============================================================
# [4] 下载主循环（#53 卡片登记 + 顺序推进；#54 三元组 key；#55 快速关面板）
# ============================================================
$targetCount = $Config.DownloadCount
Write-Log "downloading (target: $targetCount)..." -Level STEP

$ok = 0; $fail = 0; $skip = 0; $dup = 0; $preExisting = 0
# 登记表：$processed=本次运行完整 key（姓名_年龄_工作经历摘要）
#          $prefilled =断点续传基础 key（姓名_年龄，文件名不含工作经历）
$processed = @{}
$prefilled = @{}
$resumeSkipped = 0
foreach ($f in (Get-ZhaopinFiles $Config.DownloadDir)) {
    $prevKey = (Get-NameFromZhilianFile $f.Name) + '_' + (Get-AgeFromZhilianFile $f.Name)
    if (-not $prefilled.ContainsKey($prevKey)) { $prefilled[$prevKey] = $true; $resumeSkipped++ }
}
if ($resumeSkipped -gt 0) {
    Write-Log "RESUME: $resumeSkipped resume(s) already in target dir - those cards will be skipped" -Level STEP
}
$isFirst = $true

$maxEmptyRounds  = 3     # 连续 3 轮"回顶重扫均无新卡片"才认定候选池耗尽
$emptyRounds     = 0
$round           = 0
$lastOk          = 0
$consecBottom    = 0     # 连续"已到底且视口无未处理卡片"轮数
$lastSh          = 0     # #58：上一轮读到的 scrollHeight（判断懒加载是否仍在增长）
$scrollRounds    = 0
$maxScrollRounds = 300   # 滚动硬上限防死循环
$consecErrors    = 0     # 单轮未预期异常连续计数（稳定性增强：防异常风暴）
$maxConsecErrors = 5
$detectedWindowSec = if ($Config.FileFormat -eq 'word') { $Config.FileDetectWordSec } else { $Config.FileDetectPdfSec }

# #62 停滞看门狗：ok+fail+skip+dup 任一变化 = 有卡片被处理 = 有进展；
# 全部纹丝不动超过 StallTimeoutSec(300s) → 判定卡住 → exit 4 交 wrapper 重启续传
$stallDetected  = $false
$lastActivity   = 0
$lastProgressAt = [DateTime]::Now

while ($ok -lt $targetCount) {
    $round++

    # #62 停滞看门狗检查（每轮循环顶部）
    $activity = $ok + $fail + $skip + $dup
    if ($activity -gt $lastActivity) {
        $lastActivity   = $activity
        $lastProgressAt = [DateTime]::Now
    } elseif ((([DateTime]::Now) - $lastProgressAt).TotalSeconds -gt $Config.StallTimeoutSec) {
        Write-Log "STALL: no card progress for $([int](([DateTime]::Now - $lastProgressAt).TotalSeconds))s (timeout $($Config.StallTimeoutSec)s) - exiting for auto-restart (exit 4)" -Level FAIL
        $stallDetected = $true
        break
    }

    try {

    # ============================================================
    # [A] 视觉识别：提取视口内所有卡片（姓名+年龄+工作经历摘要）
    # ============================================================
    $cardsRaw = Invoke-Eval $CardExtractJs
    if ($cardsRaw -match '"ok":false') {
        # 自愈（#43）：session 无标签页 → 重新导航回列表页
        Write-Log 'RECOVER: session tab lost - re-navigating to list' -Level WARN
        $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $false }
        $null = Wait-Until -TimeoutMs 10000 -PollMs 500 -Description 'recovered list render' -Condition { (Get-VisibleCardCount) -gt 0 }
        Wait 800
        continue
    }
    $cards = @()
    if ($cardsRaw -match '"value":"(\[.*\])"') {
        try {
            $json = $Matches[1] -replace '\\"', '"'
            $cards = @( ($json | ConvertFrom-Json) | Where-Object { $_ -and $_.n } )
        } catch {
            Write-Log 'card info parse failed - treating viewport as empty' -Level WARN
        }
    }

    # ============================================================
    # [B] 找第一条未登记卡片（双表过滤）
    # ============================================================
    $next = $null
    foreach ($c in $cards) {
        $cKey = ($c.n + '_' + $c.a + '_' + $c.w)
        if (-not $processed.ContainsKey($cKey) -and -not $prefilled.ContainsKey(($c.n + '_' + $c.a))) { $next = $c; break }
    }

    if ($next) {
        $consecBottom = 0
        $name    = $next.n
        $nameKey = $next.n + '_' + $next.a + '_' + $next.w
        $wShow   = "$($next.w)"; if ($wShow.Length -gt 16) { $wShow = $wShow.Substring(0, 16) + '…' }
        Write-Log "----- round $round : $ok/$targetCount -----" -Level STEP
        Write-Log "[$ok/$targetCount] $name ($($next.a)岁) work=$wShow" -Level INFO

        # 3.1 关闭残留模态（首次跳过）
        if (-not $isFirst) { Close-ModalIfOpen }
        $isFirst = $false

        # ============================================================
        # [C] 点击卡片：前台化 → JS 三重校验打标记 → DOM click（#53：不 sweep）
        # ============================================================
        Ensure-TabFocused
        $markJs = Get-CardMarkJs -Name $name -Age $next.a -Work $next.w
        $fr = Invoke-Eval $markJs
        $clicked = $false
        if ($fr -match '"value":"ok"') {
            Ensure-TabFocused
            if (Invoke-Click '[data-wb-target="1"]') { $clicked = $true }
            $null = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-target]");if(e)e.removeAttribute("data-wb-target");return"ok"})()'
        }
        if (-not $clicked) {
            # 点击失败 → 登记后跳过（不 sweep 重试，#52 教训：名单过期时逐人 40s 空转）
            Write-Log "SKIP: card click failed - marked processed" -Level WARN
            $processed[$nameKey] = $true
            $skip++
            continue
        }
        # 已进入处理流程即登记（成功/失败/重复统一登记，防反复点击同一卡片）
        $processed[$nameKey] = $true
        Wait $Config.ClickWaitMs

        # ------------------------------------------------------------
        # 3.2.5 确认详情面板打开（#44：面板在当前标签页内渲染，无需切 tab）
        # ------------------------------------------------------------
        $opened = Wait-Until -TimeoutMs 7500 -PollMs 2500 -Description 'detail panel open' -Condition {
            (Invoke-Eval '(()=>{const b=document.querySelector(".resume-detail-wrap");return b?"yes":"no"})()') -match '"value":"yes"'
        }
        if (-not $opened) {
            # 兜底：再点一次候选人（3 次）
            $reopened = Invoke-WithRetry -MaxRetries 3 -DelayMs 2500 -Description 'reopen detail panel' -Handler {
                Ensure-TabFocused
                $null = Invoke-Click '.talent-basic-info__name'
                (Invoke-Eval '(()=>{const b=document.querySelector(".resume-detail-wrap");return b?"yes":"no"})()') -match '"value":"yes"'
            }
            if ($reopened) { $opened = $true }
        }
        if ($opened) { Write-Log 'detail panel opened' -Level OK }
        else         { Write-Log 'detail panel not detected - proceeding anyway' -Level WARN }

        # ------------------------------------------------------------
        # 3.3 等面板渲染 + "存至本地"按钮可见（#45：按钮延迟 10~15s 才挂载）
        # ------------------------------------------------------------
        $null = Wait-Until -TimeoutMs 10000 -PollMs 500 -Description 'panel body render' -Condition {
            (Invoke-Eval '(()=>{return document.querySelector(".resume-detail-wrap")?"yes":"no"})()') -match '"value":"yes"'
        }
        $saveLocalReady = Wait-Until -TimeoutMs 30000 -PollMs 750 -Description 'save-local button visible' -Condition {
            (Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return 0;const r=b.getBoundingClientRect();return(r.width>0&&r.height>0)?1:0})()') -match '"value":1'
        }
        if (-not $saveLocalReady) {
            Write-Log 'FAIL: save-local button never appeared (30s)' -Level FAIL
            $fail++
            continue
        }
        Wait 500  # 布局稳定

        # ------------------------------------------------------------
        # 3.4 点击"存至本地"（#46：按钮闪烁——探测到可见立即 CDP 点击，无额外延时；
        #     不能用 DOM click，Vue 需真实鼠标事件）
        # ------------------------------------------------------------
        $hasDialog = $false
        for ($saveRetry = 0; $saveRetry -lt 12; $saveRetry++) {
            Ensure-TabFocused
            $rslr = Invoke-Eval '(()=>{const b=document.querySelector(".resume-button.position-r");if(!b)return"x";const r=b.getBoundingClientRect();if(r.width<=0||r.height<=0)return"x";return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)})()'
            $rslc = Parse-Coords $rslr
            if ($rslc) {
                $slx = $rslc[0]; $sly = $rslc[1]
                Write-Log "save-local btn at ($slx, $sly) - clicking now" -Level INFO
                Invoke-CDP -Type 'mouseMoved'    -X $slx -Y $sly
                Invoke-CDP -Type 'mousePressed'  -X $slx -Y $sly -Button 'left'
                Wait 80
                Invoke-CDP -Type 'mouseReleased' -X $slx -Y $sly -Button 'left'
            } else {
                # 按钮未挂载/已卸载：面板还在→稍等重试；面板关了→重开
                $panelStill = Invoke-Eval '(()=>{return document.querySelector(".resume-detail-wrap")?"yes":"no"})()'
                if ($panelStill -notmatch '"value":"yes"') {
                    Write-Log 'panel closed - reopening...' -Level WARN
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
            Write-Log 'FAIL: no save dialog after 12 attempts' -Level FAIL
            $fail++
            continue
        }

        # 3.5 等对话框渲染稳定
        Wait 500

        # ------------------------------------------------------------
        # 3.5.5 选择文件格式（PDF 默认选中；word 需切换）
        # ------------------------------------------------------------
        if ($Config.FileFormat -eq 'word') {
            Write-Log 'switching format: word...' -Level INFO
            $wordCoords = Invoke-Eval '(()=>{const els=document.querySelectorAll("div");for(const el of els){if(el.childNodes.length===1&&el.childNodes[0].nodeType===3&&el.textContent.trim()==="word"){const r=el.getBoundingClientRect();if(r.width>0&&r.height>0)return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)}}return"0,0"})()'
            $wc = Parse-Coords $wordCoords
            if ($wc) {
                $wx = $wc[0]; $wy = $wc[1]
                Write-Log "word option at ($wx, $wy)" -Level INFO
                Ensure-TabFocused
                $wordClicked = $false
                $markRaw = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const els=document.querySelectorAll("div");for(const el of els){if(el.childNodes.length===1&&el.childNodes[0].nodeType===3&&el.textContent.trim()==="word"){el.setAttribute("data-wb-word","1");return"ok"}}return"no"})()' }
                if ($markRaw -match '"value":"ok"') {
                    if (Invoke-Click '[data-wb-word="1"]') { $wordClicked = $true }
                    $null = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const e=document.querySelector("[data-wb-word]");if(e)e.removeAttribute("data-wb-word");return"ok"})()' }
                }
                if (-not $wordClicked) {
                    Invoke-CDP -Type 'mouseMoved'    -X $wx -Y $wy
                    Wait 500
                    Invoke-CDP -Type 'mousePressed'  -X $wx -Y $wy -Button 'left'
                    Wait 150
                    Invoke-CDP -Type 'mouseReleased' -X $wx -Y $wy -Button 'left'
                }
                Wait 800
                # 验证：下方提示文字变为"支持 word 2010..."
                $vfy = Invoke-Eval '(()=>{const ps=document.querySelectorAll("p");for(const p of ps){if(p.textContent.indexOf("\u652f\u6301 word")>=0)return"yes"}return"no"})()'
                if ($vfy -match '"value":"yes"') { Write-Log 'format switched to word' -Level OK }
                else { Write-Log 'format switch verify failed - continuing anyway' -Level WARN }
            } else {
                Write-Log 'word option coords not found - falling back to default (pdf)' -Level WARN
            }
        }

        # ------------------------------------------------------------
        # 3.6 点击"保存"（#55：release 后 1000ms 即关面板，文件落盘由轮询检测）
        # ------------------------------------------------------------
        $clickTime = [DateTime]::Now
        Ensure-TabFocused
        $markSave = Send-Web -Action 'evaluate' -Payload @{ code = '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0){x.setAttribute("data-wb-save","1");return"ok"}}return"no"})()' }
        $saveClicked = $false
        if ($markSave -match '"value":"ok"') {
            if (Invoke-Click '[data-wb-save="1"]') { $saveClicked = $true }
        }
        if (-not $saveClicked) {
            # 退回 CDP 坐标点击
            $coords = Get-SaveButtonCoords
            if ($coords) { $sx = $coords[0]; $sy = $coords[1] }
            Write-Log "save confirm btn at ($sx, $sy) [coords fallback]" -Level INFO
            $null = Invoke-CDP -Type 'mouseMoved' -X $sx -Y $sy -ErrorAction SilentlyContinue
            Wait 200
            Invoke-CDP -Type 'mousePressed'  -X $sx -Y $sy -Button 'left'
            Wait 80
            Invoke-CDP -Type 'mouseReleased' -X $sx -Y $sy -Button 'left'
        } else {
            Write-Log 'clicked save (DOM click)' -Level INFO
        }

        # #57 核心：release → 3200ms → 关详情面板并验证（下载已由浏览器接管）；
        # 未关闭则每 ModalRetryWaitMs(1000ms) 重试关闭，上限 ModalCloseRetryMax(10) 次，
        # 已验证关闭才继续查看下一位候选人（超限 WARN 不阻断，交由 3.8/3.9 双保险兜底）
        Wait $Config.SaveCloseWaitMs
        if (Close-ModalVerified) { Write-Log 'modal close verified - proceeding to file poll' -Level OK }
        else { Write-Log 'modal failed to close after retries - continuing anyway' -Level FAIL }

        # 轮询检测新文件落盘（面板已关，不重开重试点击）
        $dlDeadline = [DateTime]::Now.AddSeconds($detectedWindowSec)
        $downloaded = $false
        while ([DateTime]::Now -lt $dlDeadline) {
            $latestFile = Get-ZhaopinFiles $Config.DownloadSource |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latestFile -and $latestFile.LastWriteTime -gt $clickTime.AddSeconds(-5)) {
                $downloaded = $true
                Write-Log "download detected: $($latestFile.Name)" -Level OK
                break
            }
            Start-Sleep -Milliseconds 1000
        }
        if (-not $downloaded) {
            Write-Log 'FAIL: download did not complete (no new zhilian file)' -Level FAIL
            $fail++
            continue
        }

        # 3.7 移动文件 + 去重
        $moveResult = Move-OneResume
        if ($moveResult -eq 'ok') {
            $ok++
            Write-Log "progress: $ok/$targetCount" -Level OK
        } elseif ($moveResult -eq 'dup') {
            $dup++
            Write-Log 'duplicate removed - trying next candidate' -Level WARN
        } else {
            Write-Log 'no file matched/moved - file stays in Downloads, check manually' -Level WARN
            $fail++
        }

        # 3.8/3.9 关闭残留模态与详情面板（#57：带验证+重试，幂等——已关闭时仅 2 次 eval 即返回）
        if (-not (Close-ModalVerified)) { Write-Log 'residual modal failed to close - next card click may fail' -Level FAIL }

    } else {
        # ============================================================
        # [D] 视口无未处理卡片 → 从当前位置下滚一屏（#53：不回顶）
        #     #58 修复（问题2）：scrollTop 赋值不触发列表懒加载（实测 scrollHeight
        #     恒定 5350、21 份即误判"池耗尽"），必须用 CDP mouseWheel 真实滚轮事件；
        #     并以 scrollHeight 增长 >200px 作为"懒加载仍生效"判据——仍在增长时
        #     绝不计入到底空轮
        #     #62 增强：连续空转最终由循环顶部的停滞看门狗兜底——即使 [D] 判定
        #     逻辑被页面改版绕过，300s 无进展也会 exit 4 触发 wrapper 自动重启
        # ============================================================
        $scrollPos = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(!c)return"0,0,0";return Math.round(c.scrollTop)+","+Math.round(c.clientHeight)+","+Math.round(c.scrollHeight)})()'
        $atBottom = $false
        $shNow = 0
        if ($scrollPos -match '"value":"(\d+),(\d+),(\d+)"') {
            $st = [int]$Matches[1]; $ch = [int]$Matches[2]; $shNow = [int]$Matches[3]
            if (($st + $ch) -ge ($shNow - 50)) { $atBottom = $true }
        }
        # #58：CDP 真实滚轮下滚一屏（3 次小步各 300ms；遮挡标签页收不到输入事件，先置前）
        Ensure-TabFocused
        for ($wi = 0; $wi -lt 3; $wi++) {
            Invoke-CDP -Type 'mouseWheel' -X 660 -Y 400 -DeltaY 1200
            Wait 300
        }
        Wait $Config.ScrollWaitMs
        $scrollRounds++
        if ($scrollRounds -gt $maxScrollRounds) {
            Write-Log "scroll hard limit ($maxScrollRounds) reached without progress - treating as stall for auto-restart" -Level WARN
            $stallDetected = $true
            break
        }
        # #58：scrollHeight 增长 → 懒加载生效，重置到底计数；只有"到底且列表不再增长"才计入空轮
        $listGrowing = ($lastSh -gt 0 -and ($shNow - $lastSh) -gt 200)
        if ($listGrowing)      { $consecBottom = 0 }
        elseif ($atBottom)     { $consecBottom++ }
        else                   { $consecBottom = 0 }
        $lastSh = $shNow

        if ($consecBottom -ge 3) {
            # 到底且连续 3 轮无新未处理卡片 → 空轮判定
            if ($ok -gt $lastOk) { $emptyRounds = 0 } else { $emptyRounds++ }
            $lastOk = $ok
            if ($emptyRounds -ge $maxEmptyRounds) {
                Write-Log "STOP: $maxEmptyRounds consecutive re-scans with no new cards - pool exhausted ($ok/$targetCount)" -Level WARN
                break
            }
            # TOP-UP（#33 精神）：回顶重扫——推荐列表动态重排会产生新卡片
            Write-Log "TOP-UP: no new cards at bottom - back to top for re-scan ($emptyRounds/$maxEmptyRounds)" -Level STEP
            $null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return"ok"})()'
            Wait 2500
            $consecBottom = 0
            $lastSh = 0
        }
    }

    # 本轮无异常走完 → 复位连续异常计数
    $consecErrors = 0

    } catch {
        # 稳定性增强：单轮未预期异常不再杀死整个任务——关面板、计数、继续
        Write-Log "unexpected error in round $round : $($_.Exception.Message)" -Level FAIL
        try { Close-ModalIfOpen } catch {}
        $consecErrors++
        if ($consecErrors -ge $maxConsecErrors) {
            # #62：连续异常终止视为停滞——exit 4 让 wrapper 重启续传，比直接放弃更符合预期
            Write-Log "STOP: $maxConsecErrors consecutive unexpected errors - treating as stall for auto-restart" -Level FAIL
            $stallDetected = $true
            break
        }
    }
}

# ============================================================
# [5] 清理与汇总
# ============================================================
Stop-BrowserAutomation

$fileCount = (Get-ZhaopinFiles $Config.DownloadDir).Count
$status = if ($ok -ge $targetCount) { 'DONE' } elseif ($stallDetected) { 'STALL' } else { 'INCOMPLETE' }

Write-Log '=====================================' -Level STEP
Write-Log "Success      : $ok"
Write-Log "Failed       : $fail"
Write-Log "Skipped      : $skip"
Write-Log "Duplicates   : $dup"
Write-Log "Format       : $($Config.FileFormat)"
Write-Log "Target files : $fileCount"
Write-Log "Target goal  : $targetCount"
if ($status -eq 'DONE') {
    Write-Log "DONE: target reached ($ok/$targetCount)" -Level OK
} elseif ($status -eq 'STALL') {
    Write-Log "STALL: no progress detected (ok=$ok/$targetCount) - exit 4, wrapper should restart with resume" -Level WARN
} else {
    Write-Log "INCOMPLETE: only $ok/$targetCount downloaded - candidate pool exhausted" -Level WARN
    Write-Log 'options: enlarge the pool (scroll more) or lower DownloadCount' -Level WARN
}

# 机器可读结果（Agent 直接读此文件判断是否需要补跑，无需解析日志）
$summary = @{
    status      = $status
    ok          = $ok
    fail        = $fail
    skip        = $skip
    dup         = $dup
    target      = $targetCount
    fileCount   = $fileCount
    format      = $Config.FileFormat
    downloadDir = $Config.DownloadDir
    jobNumber   = $EffectiveJobNumber
    finishedAt  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Compress
[System.IO.File]::WriteAllText((Join-Path $Config.DownloadDir '_summary.json'), $summary, $Utf8NoBom)

$exitCode = if ($status -eq 'DONE') { 0 } elseif ($stallDetected) { 4 } else { 2 }
exit $exitCode
