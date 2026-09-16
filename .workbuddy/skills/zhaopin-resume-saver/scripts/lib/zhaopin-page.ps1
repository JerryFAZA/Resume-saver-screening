# ============================================================
# zhaopin-page.ps1 — 智联页面交互函数库
# ============================================================
# 依赖 wb-core.ps1（Send-Web/Invoke-Eval/Invoke-Click/Invoke-CDP/
# Ensure-TabFocused/Wait/Wait-Until/ConvertTo-JsSafeName）。
# 作用域契约：引用 run.ps1 的 $Config（含坐标兜底值/DownloadDir/
# DownloadSource/FileFormat/CloseWaitMs/DialogCheckWaitMs/ScrollWaitMs/
# ModalCloseRetryMax/ModalRetryWaitMs）。
# 所有函数均为实战踩坑后验证过的版本，改动前先读 SKILL.md 对应问题记录。
# ============================================================

# ------------------------------------------------------------
# 坐标探测（getBoundingClientRect，CSS 像素；勿用 DOM.getBoxModel）
# ------------------------------------------------------------
function Parse-Coords {
    param([string]$Raw)
    if ($Raw -match '"value":"(\d+),(\d+)"') {
        $cx = [int]$Matches[1]; $cy = [int]$Matches[2]
        if ($cx -gt 10 -and $cy -gt 10) { return @($cx, $cy) }
    }
    return $null
}

function Get-SaveButtonCoords {
    $r = Invoke-Eval '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0){const r=x.getBoundingClientRect();if(r.width<=0||r.height<=0)continue;return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)}}return "0,0"})()'
    return Parse-Coords $r
}

# ------------------------------------------------------------
# 阶段 1：导航 + 岗位选择 + 岗位验证
# ------------------------------------------------------------

# 激活岗位名是否匹配目标（#61：去括号后缀 + 精确/包含两种判定）
function Test-JobActive {
    param([string]$JobName)
    $active = Get-ActiveJobName
    if (-not $active) { return $false }
    $a = $active.Trim() -replace '\s*\(.*$', ''
    $e = $JobName.Trim() -replace '\s*\(.*$', ''
    $a = $a -replace '\s*·.*$', ''   # 页面岗位标签带"·协作未上线"等点号后缀
    $e = $e -replace '\s*·.*$', ''
    if ($a -ieq $e) { return $true }          # 大小写不敏感（页面标签可能为小写 ai产品经理）
    if ($a.ToLower().Contains($e.ToLower())) { return $true }
    return $false
}

# 清除岗位卡临时标记
function Remove-JobMark {
    $null = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-job]");if(e)e.removeAttribute("data-wb-job");return"ok"})()'
}

# #61 重构（问题1，2026-09-16）：岗位卡选择必须"点击生效 + 激活岗位名验证匹配"才算成功。
# 根因：旧版只发一次 JS click 就返回成功，Vue 组件可能忽略合成 click 事件，
# 激活岗位根本没切换 → 后续按错误岗位的推荐池下载（26091609 晨批事故根因之一）。
# 现行：优先精确匹配岗位名（includes 模糊仅兜底），三级点击升级——
#   ① JS click（合成事件）→ ② Invoke-Click DOM click → ③ CDP 真实鼠标事件；
# 每级点击后都用 Test-JobActive 复核激活岗位名确实切换，未切换才试下一级。
# 返回 $true=已点击且验证匹配；$false=重试耗尽（调用方必须 FATAL 退出）
function Select-JobTab {
    param([string]$JobName)
    $safeJobName = ConvertTo-JsSafeName -Name $JobName
    for ($retry = 0; $retry -lt 15; $retry++) {
        # 查找岗位卡：优先精确匹配，退而 includes 模糊；命中即打 data 标记
        $findJs = '(()=>{const items=document.querySelectorAll(''.job-pane__item'');const q="' + $safeJobName + '".toLowerCase();let exact=null,fuzzy=null;for(const l of items){const t=l.textContent.trim();const tl=t.toLowerCase();if(tl===q){exact=l;break}if(!fuzzy&&tl.includes(q)){fuzzy=l}}const el=exact||fuzzy;if(!el)return"no";el.setAttribute("data-wb-job","1");return(exact?"exact:":"fuzzy:")+el.textContent.trim().substring(0,40)})()'
        $fr = Invoke-Eval $findJs
        if ($fr -match '"value":"(exact|fuzzy):(.*?)"') {
            $kind = $Matches[1]
            Write-Log "job card found ($kind): $($Matches[2])" -Level INFO
            # 一级：JS click（合成事件，开销最小）
            $null = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-job]");if(e){e.click();return"ok"}return"no"})()'
            Wait 1500
            if (Test-JobActive -JobName $JobName) {
                Remove-JobMark
                Write-Log 'job switched (verified via active name) [js click]' -Level OK
                return $true
            }
            # 二级：DOM click（WebBridge click action，DOM 级派发）
            Ensure-TabFocused
            if (Invoke-Click '[data-wb-job="1"]') {
                Wait 1500
                if (Test-JobActive -JobName $JobName) {
                    Remove-JobMark
                    Write-Log 'job switched (verified via active name) [dom click]' -Level OK
                    return $true
                }
            }
            # 三级：CDP 真实鼠标事件点元素中心（Vue 对合成事件免疫时的最终手段）
            $cr = Invoke-Eval '(()=>{const e=document.querySelector("[data-wb-job]");if(!e)return"0,0";const r=e.getBoundingClientRect();return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)})()'
            $c = Parse-Coords $cr
            if ($c) {
                Ensure-TabFocused
                Invoke-CDP -Type 'mouseMoved'    -X $c[0] -Y $c[1]
                Invoke-CDP -Type 'mousePressed'  -X $c[0] -Y $c[1] -Button 'left'
                Wait 80
                Invoke-CDP -Type 'mouseReleased' -X $c[0] -Y $c[1] -Button 'left'
                Wait 2000
                if (Test-JobActive -JobName $JobName) {
                    Remove-JobMark
                    Write-Log 'job switched (verified via active name) [cdp mouse]' -Level OK
                    return $true
                }
            }
            Write-Log "job card click did not take effect (round $($retry + 1)/15) - retrying..." -Level WARN
        }
        if ($retry -eq 0) { Write-Log 'job card not found/verified yet, retry every 2s (max 15)...' -Level INFO }
        Wait 2000
    }
    return $false
}

# 失败诊断：dump 页面上所有岗位标签
function Get-JobTabDump {
    $dumpJs = '(()=>{const links=document.querySelectorAll(''.job-pane__item'');const names=[];for(const l of links){names.push(l.textContent.trim().substring(0,50))}return JSON.stringify(names)})()'
    $dumpR = Invoke-Eval $dumpJs
    if ($dumpR -match '"value":"(\[.*\])"') {
        try { return (($Matches[1] -replace '\\"','"') | ConvertFrom-Json) } catch { return @() }
    }
    return @()
}

# 岗位验证：返回当前激活岗位名（可能为空串=无法检测）
# #38：只接受"看起来像岗位名"的值（含中文、长度>=2、排除常见 JS 返回值）
function Get-ActiveJobName {
    $activeJobJs = @'
(()=>{
  let el = document.querySelector('.job-pane__item--active');
  if (el && el.textContent.trim()) return el.textContent.trim();
  const container = document.querySelector('.job-pane, [class*="job-pane"], [class*="job-list"]');
  if (container) {
    const activeEl = container.querySelector('.job-pane__item--active, .is-active, .active, [class*="is-active"], [class*="active"], [aria-selected="true"]');
    if (activeEl && activeEl.textContent.trim()) return activeEl.textContent.trim();
  }
  el = document.querySelector('[aria-selected="true"]');
  if (el && el.textContent.trim()) return el.textContent.trim();
  el = document.querySelector('[aria-current="page"], [aria-current="true"]');
  if (el && el.textContent.trim()) return el.textContent.trim();
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
    $activeR = Invoke-Eval $activeJobJs
    if ($activeR -match '"value":"([^"]*)"') {
        $cand = $Matches[1].Trim()
        if ($cand -and $cand.Length -ge 2 -and
            $cand -notin @('no', 'ok', 'has', 'yes', 'none', 'null', 'true', 'false', '0', '1') -and
            $cand -match '[\u4e00-\u9fa5]') {
            return $cand
        }
    }
    # JS 探测失败 → snapshot 降级（#38）
    $snap = Send-Web -Action 'snapshot' -Payload @{}
    if ($snap -match '"role":"link"[^}]*"name":"([^"]+)"[^}]*"active"')   { return $Matches[1] }
    if ($snap -match '"role":"link"[^}]*"name":"([^"]+)"[^}]*"selected"') { return $Matches[1] }
    return ''
}

# ------------------------------------------------------------
# 候选人列表就绪检查（#39：session 无标签页自愈由调用方 navigate 完成）
# ------------------------------------------------------------
function Test-CandidateListReady {
    param([int]$TimeoutMs = 20000)
    return (Wait-Until -TimeoutMs $TimeoutMs -PollMs 1000 -Description 'candidate list render' -Condition {
        $chk = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
        if ($chk -match '"ok":false') { return $false }
        if ($chk -match '"value":"(\d+)"' -and [int]$Matches[1] -gt 0) { return $true }
        return $false
    })
}

# 视口内候选人卡片数（健康检查用）
function Get-VisibleCardCount {
    $cr = Invoke-Eval 'String(document.querySelectorAll(".talent-basic-info__name").length)'
    if ($cr -match '"value":"(\d+)"') { return [int]$Matches[1] }
    return -1
}

# ------------------------------------------------------------
# #64 网页崩溃自愈（2026-09-16 用户指令：刷简历时网页崩溃 → 自动重新加载）
# ------------------------------------------------------------

# 页面存活真探测：evaluate 最小表达式。崩溃 tab / session 无 tab / daemon 死 → 均返回 $false
function Test-PageAlive {
    $r = Invoke-Eval '1'
    return ($r -match '"ok":true')
}

# 崩溃恢复：newTab 重建标签页（死 tab 上 navigate newTab=$false 无法恢复，必须 newTab=$true 新开）
# → 等列表渲染（30s 宽限，崩溃后重载更慢）→ 岗位复核（#61 精神：恢复后必须证明激活岗位正确，
#   探测不到或不匹配 → 重新走 Select-JobTab 三级点击选岗）→ 回顶。
# $processed 登记表由主循环持有不受影响，恢复后从头扫描、已下载卡片自动跳过。返回 $true=恢复成功。
function Repair-PageCrash {
    param([string]$Reason)
    Write-Log "RECOVER: page crash detected ($Reason) - rebuilding tab..." -Level FAIL
    Ensure-TabFocused
    $null = Send-Web -Action 'navigate' -Payload @{ url = $Config.Url; newTab = $true; group_title = 'Zhaopin Resume Screening' }
    Wait 1000
    Ensure-TabFocused
    $listOk = Wait-Until -TimeoutMs 30000 -PollMs 1000 -Description 'crash reload - list render' -Condition { (Get-VisibleCardCount) -gt 0 }
    if (-not $listOk) {
        Write-Log 'RECOVER: list still not rendered after reload' -Level WARN
        return $false
    }
    # 岗位复核：URL 带正确 jobNumber 时重载后通常自动激活正确岗位；不匹配则重新选岗
    $activeJob = Get-ActiveJobName
    $expectedClean = $Config.JobName.Trim()
    $expectedNoParen = $expectedClean -replace '\s*\(.*$', ''
    $jobOk = $false
    if ($activeJob) {
        $activeClean   = $activeJob.Trim()
        $activeNoParen = $activeClean -replace '\s*\(.*$', ''
        if ($activeClean -eq $expectedClean -or $activeNoParen -eq $expectedNoParen -or $activeClean -like "*$expectedClean*") { $jobOk = $true }
    }
    if (-not $jobOk) {
        Write-Log "RECOVER: active job undetected/mismatched after reload (active=[$activeJob]) - re-selecting job..." -Level WARN
        if (-not (Select-JobTab -JobName $Config.JobName)) {
            Write-Log 'RECOVER: job re-selection failed after crash reload' -Level FAIL
            return $false
        }
        Wait 1000
        $null = Wait-Until -TimeoutMs 15000 -PollMs 1000 -Description 'post-reselect list render' -Condition { (Get-VisibleCardCount) -gt 0 }
    } else {
        Write-Log "RECOVER: active job verified [$activeJob] after reload" -Level OK
    }
    $null = Invoke-Eval '(()=>{const c=document.querySelector(".app-layout--default")||document.scrollingElement;if(c)c.scrollTop=0;window.scrollTo(0,0);return "ok"})()'
    Wait 1000
    Write-Log 'RECOVER: page crash recovery complete - resuming download loop' -Level OK
    return $true
}

# ------------------------------------------------------------
# 卡片提取 / 标记（#53/#54：key = 姓名_年龄_工作经历摘要）
# ------------------------------------------------------------

# 提取视口内所有卡片：[{n:姓名, a:年龄, w:工作经历摘要}]
# 摘要归一化链：剔除姓名 / "N岁" / 易变时间词（刚刚/N秒钟前/N分钟前/N小时前/
# N天前/昨天/本周/本月/在线/活跃/看过）/空白/引号 → 前 30 字符。
# ★ 标记端 Get-CardMarkJs 必须用完全相同的归一化链，否则 indexOf 失配。
$CardExtractJs = '(()=>{const out=[];const els=document.querySelectorAll(".talent-basic-info__name");els.forEach(el=>{const m=el.textContent.trim().match(/^\S+/);if(!m)return;const nm=m[0];let p=el,card=null,age="",i=0;while(p&&i<6){p=p.parentElement;if(!p)break;const a=p.textContent.match(/(\d{1,2})\u5c81/);if(a){age=a[1];card=p;break}}let w="";if(card){w=card.textContent.split(nm).join("").replace(/\s+/g,"").replace(/["\\]/g,"").replace(/\d{1,2}\u5c81/g,"").replace(/(\u521a\u521a|\d+\u79d2\u949f\u524d|\d+\u5206\u949f\u524d|\d+\u5c0f\u65f6\u524d|\d+\u5929\u524d|\u6628\u5929|\u672c\u5468|\u672c\u6708|\u5728\u7ebf|\u6d3b\u8dc3|\u770b\u8fc7)/g,"").substring(0,30)}out.push({n:nm,a:age,w:w})});return JSON.stringify(out)})()'

function Get-CardMarkJs {
    param([string]$Name, [string]$Age, [string]$Work)
    $safeName = ConvertTo-JsSafeName -Name $Name
    $safeWork = ConvertTo-JsSafeName -Name $Work
    # 与 $CardExtractJs 完全一致的归一化链（少一个 substring(0,30)）
    $normTail = '.replace(/\s+/g,"").replace(/["\\]/g,"").replace(/\d{1,2}\u5c81/g,"").replace(/(\u521a\u521a|\d+\u79d2\u949f\u524d|\d+\u5206\u949f\u524d|\d+\u5c0f\u65f6\u524d|\d+\u5929\u524d|\u6628\u5929|\u672c\u5468|\u672c\u6708|\u5728\u7ebf|\u6d3b\u8dc3|\u770b\u8fc7)/g,"")'
    $checks = ''
    if ($Age) {
        $checks = $checks + 'let p=el,i=0,card=null;while(p&&i<6){p=p.parentElement;if(!p)break;if(/(\d{1,2})\u5c81/.test(p.textContent)){card=p;break}}if(!card)continue;if(card.textContent.indexOf("' + $Age + '\u5c81")<0)continue;'
        if ($Work) {
            $checks = $checks + 'const norm=card.textContent.split("' + $safeName + '").join("")' + $normTail + ';if(norm.indexOf("' + $safeWork + '")<0)continue;'
        }
    }
    return '(()=>{const els=document.querySelectorAll(".talent-basic-info__name");for(const el of els){const m=el.textContent.trim().match(/^\S+/);if(!m||m[0]!=="' + $safeName + '")continue;' + $checks + 'el.setAttribute("data-wb-target","1");return"ok"}return"no"})()'
}

# ------------------------------------------------------------
# 模态/对话框
# ------------------------------------------------------------

# 安全的遮罩关闭——仅在模态打开时才点击遮罩
function Close-ModalIfOpen {
    $mc = Invoke-Eval 'String(document.querySelector(".km-modal--open")?true:false)'
    if ($mc -notmatch '"value":"true"') { return }
    Write-Log 'closing modal (mask click)...' -Level INFO
    Invoke-CDP -Type 'mousePressed'  -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
    Wait 200
    Invoke-CDP -Type 'mouseReleased' -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
    Wait $Config.CloseWaitMs
}

# #57：带验证的模态关闭——先执行一轮关闭（遮罩 + 关闭按钮兜底），再探测 .km-modal--open；
# 仍未关闭则每 ModalRetryWaitMs(1000ms) 重试，上限 ModalCloseRetryMax(10) 次（防死循环）。
# 返回 $true=已验证关闭；$false=超限仍未关闭（调用方决定是否继续，不阻断主流程）。
function Close-ModalVerified {
    Close-ModalIfOpen
    Wait 500
    $dm = Invoke-Eval '(()=>{const b=document.querySelector(".km-modal__close-btn");return b?"yes":"no"})()'
    if ($dm -match '"value":"yes"') {
        $null = Invoke-Click '.km-modal__close-btn'
        Wait 800
    }
    for ($i = 1; $i -le $Config.ModalCloseRetryMax; $i++) {
        $stillOpen = Invoke-Eval 'String(document.querySelector(".km-modal--open")?true:false)'
        if ($stillOpen -notmatch '"value":"true"') { return $true }
        Write-Log "modal still open - retry close $i/$($Config.ModalCloseRetryMax) (every $($Config.ModalRetryWaitMs)ms)" -Level WARN
        Wait $Config.ModalRetryWaitMs
        Close-ModalIfOpen
        Wait 500
        $dm = Invoke-Eval '(()=>{const b=document.querySelector(".km-modal__close-btn");return b?"yes":"no"})()'
        if ($dm -match '"value":"yes"') {
            $null = Invoke-Click '.km-modal__close-btn'
            Wait 800
        }
    }
    return $false
}

# ------------------------------------------------------------
# 智联简历文件识别 / 移动 / 去重
# （文件名结构化识别，纯 ASCII 安全：中文 Windows 下 UTF-8 无 BOM 的
#   .ps1 会被 GBK 解析，脚本内中文字面量必乱码——见 SKILL.md）
# ------------------------------------------------------------

# 智联文件名格式: {姓名}_{年龄}岁_智联简历_{5位数字}.{pdf|docx}
function Get-ZhaopinFiles {
    param([string]$Dir)
    $all = Get-ChildItem -Path "$Dir\*" -Include "*.pdf", "*.docx" -ErrorAction SilentlyContinue
    if (-not $all) { return @() }
    $result = $all | Where-Object {
        $n = $_.Name
        $parts = $n -split '_'
        ($parts.Count -ge 4) -and ($n -match '_\d+\.(pdf|docx)$')
    }
    return @($result)
}

function Get-NameFromZhilianFile {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $idx = $base.IndexOf('_')
    if ($idx -gt 0) { return $base.Substring(0, $idx) }
    return ''
}

function Get-AgeFromZhilianFile {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $parts = $base -split '_'
    if ($parts.Count -ge 2) {
        if ($parts[1] -match '^(\d+)') { return $Matches[1] }
    }
    return ''
}

# 重复判定：姓名 + 年龄 + 文件大小 三重验证
function Test-DuplicateFile {
    param([string]$FilePath, [string]$TargetDir)
    $fname = Split-Path -Leaf $FilePath
    $name = Get-NameFromZhilianFile $fname
    $age  = Get-AgeFromZhilianFile $fname
    if (-not $name -or -not $age) { return $false }
    $size = (Get-Item $FilePath -ErrorAction SilentlyContinue).Length
    $existing = Get-ZhaopinFiles $TargetDir | Where-Object {
        (Get-NameFromZhilianFile $_.Name) -eq $name -and
        (Get-AgeFromZhilianFile $_.Name)  -eq $age -and
        ($_.Length -eq $size)
    } | Select-Object -First 1
    return ($null -ne $existing)
}

# 移动 Downloads 最新智联文件至目标目录。
# 返回约定（调用方必须区分）：'ok' 成功 / 'dup' 重复已删 / $null 未找到或移动失败
function Move-OneResume {
    $files = Get-ZhaopinFiles $Config.DownloadSource
    if (-not $files -or $files.Count -eq 0) {
        Write-Log 'no zhilian file found in Downloads' -Level WARN
        return $null
    }
    $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { return $null }
    Write-Log ("latest download: {0} ({1} bytes, {2})" -f $latest.Name, $latest.Length, $latest.LastWriteTime.ToString('HH:mm:ss')) -Level INFO

    $dest = Join-Path $Config.DownloadDir $latest.Name
    if (Test-Path $dest) {
        if (Test-DuplicateFile -FilePath $latest.FullName -TargetDir $Config.DownloadDir) {
            Write-Log 'identical resume already in target dir - removing new file' -Level WARN
            Remove-Item $latest.FullName -Force -ErrorAction SilentlyContinue
            return 'dup'
        }
        Write-Log 'same name but different candidate - overwrite moving' -Level INFO
    }
    # 移动重试：文件可能被杀软/索引服务短暂占用
    $moved = $false
    for ($mi = 0; $mi -lt 3 -and -not $moved; $mi++) {
        Move-Item -Path $latest.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
        if (Test-Path $dest) { $moved = $true } else { Wait 500 }
    }
    if ($moved) {
        Write-Log "moved: $($latest.Name)" -Level OK
        return 'ok'
    }
    Write-Log 'move failed (file may be locked/used)' -Level FAIL
    return $null
}
