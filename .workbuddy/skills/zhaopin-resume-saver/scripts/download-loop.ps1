# ============================================================
# 智联招聘简历批量下载 — 下载循环逻辑（可独立引入）
#
# 依赖：config.ps1 + webbridge-utils.ps1（或 run.ps1 内建函数）
# ============================================================

<#
.SYNOPSIS 纯 ASCII 结构匹配识别智联简历文件（不依赖中文字面量）
.DESCRIPTION 智联文件名格式: {姓名}_{年龄}岁_智联简历_{数字}.{pdf|docx}
  用纯 ASCII 特征: 至少 4 段 _ 分割、末段纯数字+扩展名。
  避免中文通配符在 GBK 代码页下的匹配失败问题。
#>
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
        if ($parts[1] -match '^(\d+)') {
            return $Matches[1]
        }
    }
    return ''
}

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
.SYNOPSIS 从 Downloads 移动最新智联 PDF 到目标目录。重复则删除并返回 $null。
#>
function Move-OneResume {
    param([string]$TargetDir)
    # 纯 ASCII 结构匹配：不依赖中文字面量，避免 GBK 代码页下的通配符匹配失败
    $files = Get-ZhaopinFiles "$env:USERPROFILE\Downloads"
    if (-not $files -or $files.Count -eq 0) { return $null }

    $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { return $null }

    $dest = Join-Path $TargetDir $latest.Name
    if ((Test-Path $dest) -and (Test-DuplicateFile -FilePath $latest.FullName -TargetDir $TargetDir)) {
        Write-Host '  [DUP] Identical resume already in target dir — removing new file' -ForegroundColor Yellow
        Remove-Item $latest.FullName -Force -ErrorAction SilentlyContinue
        return $null
    }
    Move-Item -Path $latest.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
    if (Test-Path $dest) {
        Write-Host "  [MOVED] $($latest.Name)" -ForegroundColor Green
        return $latest.Name
    }
    Write-Host '  [FAIL] Move failed (file may be locked/used)' -ForegroundColor Red
    return $null
}

<#
.SYNOPSIS 收集当前页面可见候选人姓名（累积 Set）
#>
function Collect-Names {
    param([int]$MaxScroll = 12)
    $null = Invoke-Evaluate 'window._SN = new Set()'
    Wait 300
    $last = 0; $stag = 0
    for ($i = 0; $i -lt $MaxScroll; $i++) {
        # 用 ^\S+ 正则提取纯姓名，避免换行符干扰
        $js = '(()=>{const e=document.querySelectorAll(''.talent-basic-info__name'');e.forEach(el=>{const m=el.textContent.trim().match(/^\S+/);if(m)window._SN.add(m[0])});return''''})()'
        $null = Invoke-Evaluate $js
        Invoke-CDP -Type 'mouseWheel' -X 500 -Y 600 -DeltaY 800
        Wait 1200
        $r = Invoke-Evaluate 'String(window._SN.size)'
        if ($r -match '"value":"(\d+)"') {
            $cur = [int]$Matches[1]
            if ($cur -eq $last) { $stag++ } else { $stag = 0; $last = $cur }
            if ($stag -ge 4) { break }
        }
    }
    $r = Invoke-Evaluate 'JSON.stringify(Array.from(window._SN))'
    if ($r -match '"value":"(\[.*\])"') {
        try {
            $json = $Matches[1] -replace '\\"', '"'
            $raw  = $json | ConvertFrom-Json
            return @($raw | Where-Object { $_ } | ForEach-Object { $_ -replace '\s', '' } | Where-Object { $_ })
        } catch {}
    }
    return @()
}

<#
.SYNOPSIS 按姓名查找并点击候选人 → 返回是否成功
#>
function Find-Click-Candidate {
    param([string]$Name, [int]$MaxRetry = 15)
    # 将中文姓名转为 JS Unicode 转义，避免 PS 双引号字符串乱码
    $safe = ''
    for ($ci = 0; $ci -lt $Name.Length; $ci++) {
        $c = $Name[$ci]
        if ([int]$c -gt 127) { $safe += '\u{0:X4}' -f [int]$c } else { $safe += $c }
    }
    for ($a = 0; $a -lt $MaxRetry; $a++) {
        $js = '(()=>{const e=document.querySelectorAll(''.talent-basic-info__name'');for(const el of e){if(el.textContent.trim().indexOf("' + $safe + '")>=0){el.click();return"ok"}}return"no"})()'
        $r = Invoke-Evaluate $js
        if ($r -match '"value":"ok"') { return $true }
        Invoke-CDP -Type 'mouseWheel' -X 500 -Y 600 -DeltaY 600
        Wait 1000
    }
    return $false
}

<#
.SYNOPSIS 下载单条简历（含去重检查）
#>
function Download-One {
    param([string]$Name, [string]$TargetDir, [int]$Idx, [int]$Total)
    Write-Host "=== [$( if ($Total -gt 0) { "$Idx/$Total" } else { $Idx } )] $Name ==="

    # 关闭模态（仅在打开时才点击遮罩）
    $mc = Invoke-Evaluate 'String(document.querySelector(".km-modal--open")?true:false)'
    if ($mc -match '"value":"true"') {
        Write-Host '  Closing modal...' -ForegroundColor Gray
        $null = Invoke-Evaluate 'window.scrollTo(0,0)'
        Wait 500
        Invoke-CDP -Type 'mousePressed'  -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
        Wait 200
        Invoke-CDP -Type 'mouseReleased' -X $Config.MaskCloseX -Y $Config.MaskCloseY -Button 'left'
        Wait $Config.CloseWaitMs
    }

    # 查找并点击候选人
    if (-not (Find-Click-Candidate $Name)) {
        Write-Host '  [SKIP] Not found' -ForegroundColor Yellow; return @{ Status = 'skip' }
    }
    Wait $Config.ClickWaitMs

    # CDP 点击"存至本地"
    Invoke-CDP -Type 'mouseMoved'    -X $Config.SaveLocalX -Y $Config.SaveLocalY
    Wait $Config.MouseMoveWaitMs
    Invoke-CDP -Type 'mousePressed'  -X $Config.SaveLocalX -Y $Config.SaveLocalY -Button 'left'
    Wait $Config.PressReleaseWaitMs
    Invoke-CDP -Type 'mouseReleased' -X $Config.SaveLocalX -Y $Config.SaveLocalY -Button 'left'
    Wait $Config.DialogCheckWaitMs

    # 检查保存对话框（含重试）
    # 使用 \uXXXX Unicode 转义避免 PS 编码乱码；单个 " 包裹 JS 字符串避免双引号陷阱
    $has = $false
    for ($retry = 0; $retry -lt 3; $retry++) {
        $r = Invoke-Evaluate '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0)return"has"}return"no"})()'
        if ($r -match '"value":"has"') { $has = $true; break }
        Invoke-CDP -Type 'mousePressed'  -X $Config.SaveLocalX -Y $Config.SaveLocalY -Button 'left'
        Wait $Config.PressReleaseWaitMs
        Invoke-CDP -Type 'mouseReleased' -X $Config.SaveLocalX -Y $Config.SaveLocalY -Button 'left'
        Wait $Config.DialogCheckWaitMs
    }
    if (-not $has) {
        Write-Host '  [FAIL] No save dialog' -ForegroundColor Red; return @{ Status = 'fail' }
    }

    # 点击保存
    Invoke-CDP -Type 'mousePressed'  -X $Config.SaveConfirmX -Y $Config.SaveConfirmY -Button 'left'
    Wait 200
    Invoke-CDP -Type 'mouseReleased' -X $Config.SaveConfirmX -Y $Config.SaveConfirmY -Button 'left'
    Wait $Config.DownloadWaitMs

    # 移动并去重
    $moved = Move-OneResume -TargetDir $TargetDir
    if ($moved) {
        Write-Host "  [OK] $moved" -ForegroundColor Green
        return @{ Status = 'ok'; File = $moved }
    }
    Write-Host '  [DUP] Duplicate removed' -ForegroundColor Yellow
    return @{ Status = 'dup' }
}