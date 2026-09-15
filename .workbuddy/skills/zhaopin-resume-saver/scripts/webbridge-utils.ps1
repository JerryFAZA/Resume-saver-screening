# ============================================================
# WebBridge 通信工具模块（保留作为高级函数库）
# 
# 注意：主流程现在使用 run.ps1 中内建的 Send-Web / Invoke-Eval /
#       Invoke-CDP 函数，它们使用 ConvertTo-Json 构建请求，
#       比本模块的手动拼接方式更可靠。
#
# 以下函数在需要单独调用 WebBridge 命令时使用。
# ============================================================

$Script:UTF8 = New-Object System.Text.UTF8Encoding $False
$Script:ReqFile = "$env:TEMP\wb-utils.json"

function Write-RequestFile {
    param([string]$Json)
    [System.IO.File]::WriteAllText($Script:ReqFile, $Json, $Script:UTF8)
}

function Send-Request {
    curl.exe -s -X POST $Config.WebBridgeUrl -H "Content-Type: application/json" --data-binary "@$Script:ReqFile"
}

<#
.SYNOPSIS 导航到指定 URL（推荐使用 run.ps1 内建函数）
#>
function Invoke-Navigate {
    param([string]$Url, [bool]$NewTab = $true, [string]$GroupTitle = "Zhaopin", [string]$Session = $Config.Session)
    $body = @{ action = 'navigate'; args = @{ url = $Url; newTab = $NewTab; group_title = $GroupTitle }; session = $Session } | ConvertTo-Json -Compress -Depth 5
    Write-RequestFile $body
    return Send-Request
}

<#
.SYNOPSIS 获取页面快照
#>
function Invoke-Snapshot {
    param([string]$Session = $Config.Session)
    $body = @{ action = 'snapshot'; args = @{}; session = $Session } | ConvertTo-Json -Compress
    Write-RequestFile $body
    return Send-Request
}

<#
.SYNOPSIS 在页面执行 JS → 返回原始 JSON 字符串
#>
function Invoke-Evaluate {
    param([string]$JsCode, [string]$Session = $Config.Session)
    $body = @{ action = 'evaluate'; args = @{ code = $JsCode }; session = $Session } | ConvertTo-Json -Compress -Depth 5
    Write-RequestFile $body
    return Send-Request
}

<#
.SYNOPSIS 确保守护进程运行
#>
function Start-Daemon {
    $null = & "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start 2>&1
}

<#
.SYNOPSIS 清理临时文件
#>
function Clear-Temp {
    Remove-Item $Script:ReqFile -Force -ErrorAction SilentlyContinue
}

function Wait { param([int]$Ms) Start-Sleep -Milliseconds $Ms }