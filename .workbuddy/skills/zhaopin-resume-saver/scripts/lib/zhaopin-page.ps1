# ============================================================
# zhaopin-page.ps1 — 智联页面交互函数库
# ============================================================
# 依赖 wb-core.ps1（Send-Web/Invoke-Eval/Invoke-Click/Invoke-CDP/
# Ensure-TabFocused/Wait/Wait-Until/ConvertTo-JsSafeName）。
# 作用域契约：引用 run.ps1 的 $Config（含坐标兜底值/DownloadDir/
# DownloadSource/FileFormat/CloseWaitMs/DialogCheckWaitMs/ScrollWaitMs）。
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

# 返回 $true=点击成功；$false=15 轮未命中（调用方负责 FATAL 退出）
function Select-JobTab {
    param([string]$JobName)
    $jobTabClicked = $false
    $safeJobName = ConvertTo-JsSafeName -Name $JobName
    # 双路径策略（#28）：snapshot 包含匹配（支持带后缀岗位名）+ JS 模糊匹配兜底
    for ($retrySnap = 0; $retrySnap -lt 15; $retrySnap++) {
        # 路径 A: snapshot 包含匹配
        $snap = Send-Web -Action 'snapshot' -Payload @{}
        if ($snap -match '"name":"[^"]*' + [regex]::Escape($JobName) + '[^"]*","ref":"(@e\d+)"') {
            $null = Send-Web -Action 'click' -Payload @{ selector = $Matches[1] }
            Write-Log "job tab matched via snapshot: $($Matches[1])" -Level OK
            $jobTabClicked = $true
            break
        }
        # 路径 B: JS evaluate 模糊匹配 .job-pane__item
        $jsFindTab = '(()=>{const links=document.querySelectorAll(''.job-pane__item'');for(const l of links){if(l.textContent.trim().includes("' + $safeJobName + '")){l.click();return"ok:"+l.textContent.trim().substring(0,40)}}return"no"})()'
        $ftResult = Invoke-Eval $jsFindTab
        if ($ftResult -match '"value":"ok:(.*?)"') {
            Write-Log "job tab matched via JS fuzzy: $($Matches[1])" -Level OK
            $jobTabClicked = $true
            break
        }
        if ($retrySnap -eq 0) { Write-Log 'job tab not found yet, retry every 2s (max 15)...' -Level INFO }
        Wait 2000
    }
    return $jobTabClicked
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
