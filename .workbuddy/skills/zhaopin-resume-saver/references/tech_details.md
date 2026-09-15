# 智联招聘页面技术参考

## 页面 URL 结构

```
https://rd6.zhaopin.com/app/recommend?tab=recommend&jobNumber={JOB_NUMBER}#sortType=recommend
```

- `jobNumber`: 职位编号，由创建职位时系统生成
- `tab=recommend`: 推荐 Tab（默认）
- `#sortType=recommend`: 排序方式
- ⚠️ URL 中 `#` 片段可能导致 PowerShell 字符串拼接异常，建议去掉，仅用查询参数

## 关键 CSS 选择器

| 用途 | 选择器 | 说明 |
|------|--------|------|
| 候选人卡片 | `.recommend-item` | 带 `data-index` 属性，从 0 开始 |
| 候选人姓名 | `.talent-basic-info__name` | 点击此元素打开简历详情 |
| 存至本地按钮 | `.resume-button.position-r` | 简历详情右侧动作区的 `<div>` |
| 简历详情模态 | `.km-modal--open` | Vue 模态组件 |
| 关闭按钮（备选） | `.new-shortcut-resume__close` | 仅在有模态时存在 |
| 关闭按钮（备选） | `.job-pane__item--close` | 旧版关闭方式 |
| **遮罩关闭（推荐）** | **左侧/右侧黑色区域** | **CDP 点击 (30, 300) 最可靠** |
| 保存按钮 | `button` (textContent = "保存") | 保存对话框内的确认按钮 |
| 取消按钮 | `button` (textContent = "取消") | 保存对话框内的取消按钮 |

## CDP 鼠标事件坐标

> 以下坐标适用于 1920×1080 分辨率的视口。页面缩放或窗口大小变化时需重新探测。**推荐在 config.ps1 中配置，优先使用动态探测。**

| 操作 | 坐标 (cx, cy) | 说明 |
|------|---------------|------|
| 存至本地 — mouseMoved | (1079, 142) | 简历详情右侧 |
| 存至本地 — mousePressed/Released | (1079, 142) | 同上 |
| 保存 — mousePressed/Released | (996, 561) | 弹出对话框中的保存按钮 |
| **遮罩关闭** | **(30, 300)** | **左侧黑色区域，最可靠的关闭方式** |
| 确认弹窗 — 取消按钮 | (846, 423) | 意外弹窗 |

## 虚拟滚动机制（重要）

智联推荐列表使用**虚拟滚动**，DOM 中始终仅保留约 20 个 `.recommend-item` 节点。

- `document.querySelectorAll('.talent-basic-info__name').length` ≈ 20
- `document.querySelectorAll('.recommend-item').length` ≈ 20
- 页面 `scrollHeight` 很大但 `clientHeight` 固定

**影响**：
1. 无法一次性获取全部候选人列表，需要滚动累积收集
2. 关闭模态后 DOM 可能变化，需要重新滚动定位目标候选人
3. 按索引 `[i]` 点击不可靠——索引对应的是当前视口位置，不是全局位置
4. **循环中使用的姓名与实际被点击的候选人可能不一致**——虚拟滚动导致 DOM 元素在关闭/打开模态间被替换

**解决方案**：
- 使用 `window._SN` Set 累积去重收集
- 每次循环用 `includes()` 按姓名匹配，找不到则 `mouseWheel` 滚动
- 关闭模态 + 打开新候选人之间需要充分等待（700-800ms）
- **收集姓名时用 `/^\S+/` 正则提取第一个非空白序列**，避免 `textContent` 中换行符的干扰

## 动态获取坐标（推荐）

> **统一使用 `getBoundingClientRect()`（CSS 像素）获取所有按钮坐标**。CDP `Input.dispatchMouseEvent` 使用 CSS 像素坐标系，两者一致。**不要用 `DOM.getBoxModel`**（返回设备像素，且 nodeId=0 时失败，详见踩坑记录 #18/#19）。

```javascript
// 获取存至本地按钮坐标（返回 "cx,cy" 字符串）
(() => {
  const btn = document.querySelector(".resume-button.position-r");
  if (!btn) return "0,0";
  const r = btn.getBoundingClientRect();
  if (r.width <= 0 || r.height <= 0) return "0,0";
  return Math.round(r.x + r.width/2) + "," + Math.round(r.y + r.height/2);
})();

// 获取保存按钮坐标（返回 "cx,cy" 字符串）
(() => {
  const btns = document.querySelectorAll("button");
  for (const b of btns) {
    if (b.textContent.trim() === "保存" && b.offsetWidth > 0) {
      const r = b.getBoundingClientRect();
      if (r.width <= 0 || r.height <= 0) continue;
      return Math.round(r.x + r.width/2) + "," + Math.round(r.y + r.height/2);
    }
  }
  return "0,0";
})();
```

PowerShell 侧用 `Parse-Coords` 解析 `"value":"(\d+),(\d+)"` 格式提取坐标。

## CDP 完整点击序列

存至本地按钮的可靠点击序列（不可省略任何一步）：

```
1. mouseMoved   → (SaveLocalX, SaveLocalY)
2. 等待         → ≥800ms
3. mousePressed → (SaveLocalX, SaveLocalY, button:"left", clickCount:1)
4. 等待         → 250-400ms
5. mouseReleased→ (SaveLocalX, SaveLocalY, button:"left", clickCount:1)
6. 等待         → ≥1500ms
7. 检查         → evaluate 查找"保存"按钮
8. 如未出现     → 重复步骤 3-5（省略 mouseMoved），再等待 1500ms，再检查
```

保存按钮的点击序列：

```
1. mousePressed → (SaveConfirmX, SaveConfirmY, button:"left", clickCount:1)
2. 等待         → 200ms
3. mouseReleased→ (SaveConfirmX, SaveConfirmY, button:"left", clickCount:1)
4. 等待         → 5-6 秒（确保文件完全写入磁盘）
```

遮罩关闭模态的点击序列（**当前最可靠的关闭方式**）：

```
1. mousePressed → (MaskCloseX, MaskCloseY) = (30, 300)
2. 等待         → 200ms
3. mouseReleased→ (MaskCloseX, MaskCloseY)
4. 等待         → 700ms
```

## 文件管理

### 下载路径

- Windows 默认下载目录：`$env:USERPROFILE\Downloads`
- 文件命名格式：`{姓名}_{年龄}岁_智联简历_{5位随机数}.pdf`
- 示例：`刘先生_24岁_智联简历_18535.pdf`、`戴先生_33岁_智联简历_99875.pdf`

### 下载后立即移动

每条简历下载完成后，必须**立即**将文件从 Downloads 移动到用户指定的 `downloadDir`：

```powershell
$latest = Get-ChildItem -Path "$env:USERPROFILE\Downloads" -Filter "*智联*" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latest) {
    Move-Item -Path $latest.FullName -Destination "$downloadDir\" -Force
}
```

### 文件去重（三重验证）

**重复判断条件（必须同时满足）：**

1. **姓名相同** — 文件名中 `_` 分隔的第一部分（如 `张先生`）
2. **年龄相同** — 文件名中 `_` 后、`岁` 前的数字（如 `35`）
3. **文件大小相同** — 字节数完全一致

**重复处理策略（不重置流程）：**
- 发现重复 → 删除新下载的重复文件
- 当前循环迭代不计入成功数
- 继续尝试下一位候选人
- 流程中的 `ok` 计数器保持不变，直到成功数达到 `downloadCount`

```powershell
function Test-DuplicateFile {
    param([string]$FilePath, [string]$TargetDir)
    $fname = Split-Path -Leaf $FilePath
    if ($fname -match '^(.+)_(\d+)岁_智联简历_\d+\.pdf$') {
        $name = $Matches[1]
        $age  = $Matches[2]
        $size = (Get-Item $FilePath -ErrorAction SilentlyContinue).Length
        $existing = Get-ChildItem $TargetDir -Filter "${name}_*岁_智联简历_*.pdf" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -match "^${name}_(\d+)岁_智联简历_\d+\.pdf$" -and
                ($Matches[1] -eq $age) -and ($_.Length -eq $size)
            } | Select-Object -First 1
        return ($existing -ne $null)
    }
    return $false
}
```

## PowerShell JSON 与 curl 实践

### 编码要求（CRITICAL）

> **所有 .ps1 脚本文件必须使用 UTF-8 BOM 编码。** 在中文 Windows 上，PowerShell 5.x 默认以系统 ANSI 代码页（GBK）解析无 BOM 的 .ps1 文件，导致中文字符乱码、箭头函数语法被错误解析等问题。

```powershell
# ✅ 正确：写入 UTF-8 BOM 文件
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)

# ✅ 请求 JSON 文件使用无 BOM UTF-8
$utf8nobom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText("$env:TEMP\request.json", $jsonContent, $utf8nobom)
```

> 禁止使用 `Out-File -Encoding utf8`（会产生 BOM，导致 curl 返回 `invalid JSON`）。

### 推荐：用 ConvertTo-Json 替代手动字符串拼接

```powershell
# ✅ 推荐：ConvertTo-Json 自动处理转义
$body = @{ action = 'evaluate'; args = @{ code = $jsCode }; session = 'resume-screening' } | ConvertTo-Json -Compress -Depth 5
[System.IO.File]::WriteAllText($file, $body, $utf8nobom)
curl.exe -s -X POST http://127.0.0.1:10086/command -H 'Content-Type: application/json' --data-binary "@$file"

# ❌ 避免：手动拼接 JSON（引号、URL中的#等容易出错）
'{"action":"navigate","args":{"url":"' + $Url + '",...}}'
```

### PowerShell 参数名陷阱

> **不要将函数参数命名为 `$Args`**。`$Args` 是 PowerShell 的自动变量（存储所有未绑定参数），使用同名参数会导致类型转换错误。

```powershell
# ❌ 错误：$Args 与自动变量冲突
function Web-Request { param([hashtable]$Args) ... }

# ✅ 正确：使用其他名称
function Web-Request { param($Payload) ... }
```

### JavaScript 中的中文处理

> **核心原则：避免在 PowerShell 字符串中直接出现中文字符**

```powershell
# ❌ 错误：中文会在 PS 中变成乱码
$js = "(()=>{el.click();return '张先生'})()"

# ✅ 正确：用 Unicode 转义
$js = "(()=>{el.click();return '\u5f20\u5148\u751f'})()"

# ✅ 正确：将 JS 写入独立文件再读取
[System.IO.File]::WriteAllText("$env:TEMP\script.js", $jsContent, $utf8nobom)
$code = [System.IO.File]::ReadAllText("$env:TEMP\script.js")

# ✅ 可行：.ps1 文件为 UTF-8 BOM 时，字符串内的中文可正常解析
```

### 姓名提取最佳实践

`textContent` 不以姓名开头，而是以换行和空格开头：
- 原始值：`"\n        张先生\n        \n    12小时前浏览过职位"`
- **推荐用正则 `/^\S+/` 提取**：`trim()` 后匹配第一个非空白连续序列
- 比 `indexOf(' ')` + `substring` 更可靠，因为内部的换行 `\n` 也是空白字符
- 查找时使用 `includes('张先生')` 而非 `startsWith('张先生')`

```javascript
// 推荐：正则提取纯姓名
const m = el.textContent.trim().match(/^\S+/);
if (m) window._SN.add(m[0]);

// 备选：取第一个空格前（不如正则可靠）
const t = el.textContent.trim();
const sp = t.indexOf(' ');
const n = sp > 0 ? t.substring(0, sp) : t;
```

### JSON 解析姓名数组

```powershell
# 获取姓名数组
$r = Invoke-Evaluate 'JSON.stringify(Array.from(window._SN))'

# 响应格式："value":"[\"张三\n\",\"李四\\n\",...]"
# 1. 用 .* 贪婪匹配捕获完整数组
if ($r -match '"value":"(\[.*\])"') {
    # 2. 转换 JSON 转义 \" → "
    $json = $Matches[1] -replace '\\"', '"'
    # 3. 交由 ConvertFrom-Json 正确解析
    $raw = $json | ConvertFrom-Json
    # 4. 清理白字符
    $names = @($raw | Where-Object { $_ } | ForEach-Object { $_ -replace '\s', '' } | Where-Object { $_ })
}
```

## AI Agent 执行经验

### A1. PowerShell exit code 1 但脚本实际成功（非 bug）

- **现象**：AI Agent 通过 `execute_command` 运行 `run.ps1`，脚本输出显示 `Success: 5` 且所有简历均已下载并移动到目标目录，但 `execute_command` 返回 `exit code 1`，导致 Agent 误判脚本失败。
- **根因**：PowerShell 进度条输出（`#< CLIXML` + `<Objs>` XML）被写入 stderr，`$?` 或 `$LASTEXITCODE` 被污染为 1。同时 `Write-Host` 输出进度对象（如模块首次加载的 `正在准备首次使用模块`）也会触发 CLIXML 进度流。
- **判断标准**：应检查脚本的标准输出内容（`Success : N`、`Target files : N`）而非依赖 exit code。脚本末尾的汇总输出 `===== Success : N =====` 是唯一可靠的成败判定依据。
- **解决**：Agent 解析 stdout 中的 `Success\s*:\s*(\d+)` 和 `Target files\s*:\s*(\d+)`，两者匹配且 > 0 即视为成功。

### A2. Python 路径含中文空格时 PowerShell 调用失败

- **现象**：`c:\Sinlein\timeline\260724\ 【人资】\简历筛选\.venv\Scripts\python.exe -c "..."` 报 `CommandNotFoundException`，错误指向路径中的第一个空格前的部分。
- **根因**：PowerShell 在解析含空格的路径时，如果没有用引号包裹或使用调用运算符 `&`，会将空格前的部分当作独立命令名，空格后的部分当作参数，导致路径被截断。
- **解决**：使用调用运算符 `&` 包裹完整路径：
  ```powershell
  # ❌ 错误：路径中空格导致解析截断
  c:\path\with spaces\python.exe -c "..."

  # ✅ 正确：用 & 运算符 + 双引号包裹路径
  & "c:\path\with spaces\python.exe" -c "..."
  ```

### A3. config.json 的 JobName 必须与用户指令保持一致

- **现象**：用户指令要求下载"渠道部经理"的简历，但 `config.json` 中 `JobName` 仍为上次使用的"销售部经理"，脚本会在阶段 2.1 岗位验证时检测到不匹配并终止。
- **解决**：每次执行前，Agent 必须从用户消息中提取岗位名称，更新 `config.json` 的 `JobName` 和 `DownloadDir` 字段。不要假设 config.json 中的值仍然正确。
- **更新字段清单**：`Url`（含 jobNumber）、`JobName`、`DownloadDir`、`DownloadCount`、`FileFormat`。

### A4. 页面岗位标签可能带后缀（如"·协作未上线"）

- **现象**：点击岗位标签后，阶段 2.1 岗位验证检测到的激活岗位名称为 `渠道部经理·协作未上线`，与参数 `渠道部经理` 不完全相等。
- **根因**：智联招聘页面在某些岗位标签后会追加状态后缀（如"·协作未上线"），这是页面正常行为，不影响候选人列表的正确性。
- **解决**：脚本的岗位验证逻辑已使用 `-like "*$expectedClean*"` 进行包含匹配，`渠道部经理·协作未上线` 包含 `渠道部经理` 子串，正确判定为匹配。这是预期行为，不是 bug。
- **经验**：页面 UI 标签可能包含动态后缀，验证逻辑应使用包含匹配而非严格相等。

### A5. WebBridge 版本过旧提示可忽略

- **现象**：启动守护进程时输出 `kimi-webbridge v1.11.5 available (current v1.11.3). Run: kimi-webbridge upgrade`。
- **影响**：v1.11.3 与 v1.11.5 之间不存在破坏性 API 变更，所有工具调用均正常工作。版本提示仅为建议升级，不影响功能。
- **处理**：Agent 可忽略此提示继续执行。不要自动执行 `kimi-webbridge upgrade`。

## 踩坑记录

### 1. 候选人卡片不响应 .click()

- **现象**：`document.querySelectorAll('[role="listitem"]')[0].click()` 无效
- **原因**：智联使用 Vue，卡片上的事件委托与 DOM 合成点击不兼容
- **解决**：点击卡片内的姓名 div（`.talent-basic-info__name`）

### 2. 存至本地按钮需要 CDP 真实事件

- **现象**：`.click()`、`.dispatchEvent(new MouseEvent('click', ...))`、webbridge `click` 工具均无法弹出对话框
- **原因**：Vue 组件检查 `event.isTrusted`（合成事件 isTrusted 为 false）
- **解决**：使用 CDP `Input.dispatchMouseEvent`，发送完整 mouseMoved → mousePressed → mouseReleased 序列

### 3. 必须先 mouseMoved 再 click

- **现象**：直接发送 `mousePressed` + `mouseReleased` 无法触发对话框（首次必定失败）
- **原因**：Vue 组件需要 hover 状态（鼠标必须在元素上方）
- **解决**：在 `mousePressed` 之前先发送 `mouseMoved` 到目标坐标，等待 ≥800ms

### 4. 存至本地对话框需两次点击

- **现象**：约 50% 情况下，首次 CDP click 序列（含 mouseMoved）不会弹出保存对话框
- **解决**：自动重试机制——检查"保存"按钮是否存在；不存在则再执行一次 press→release

### 5. 遮罩关闭模态（比 close 按钮更可靠）

- **现象**：`.new-shortcut-resume__close` 的 JS `.click()` 在虚拟滚动后可能失效（DOM 节点被替换）
- **解决**：CDP 点击模态外的左侧黑色遮罩区域 (30, 300)，物理遮罩始终存在且可点击

### 6. PowerShell 中文编码问题（CRITICAL）

- **现象**：PS 命令行中嵌入中文 → 乱码；`.ps1` 文件中中文导致 ParseError
- **本质**：PowerShell 5.x 在中文 Windows 上以 GBK 解析无 BOM 的 .ps1 文件
- **解决**：
  - **所有 .ps1 文件必须使用 UTF-8 BOM 编码**
  - JS 内中文优先用 `\uXXXX` Unicode 转义
  - 中文数据可存独立 `.txt` 文件（UTF-8），通过 `ReadAllText` 读取

### 7. 虚拟滚动 DOM 节点数固定

- **现象**：无论怎么滚动，`document.querySelectorAll('.talent-basic-info__name').length` 始终约为 20
- **解决**：用 `window._SN` Set 累积收集，每次滚动后把当前视口姓名加入 Set

### 8. PowerShell 参数名 `$Args` 冲突

- **现象**：`Cannot convert the "System.Object[]" value to type "System.Collections.Hashtable"`
- **原因**：`$Args` 是 PowerShell 自动变量，存储未绑定的参数值
- **解决**：使用其他参数名（如 `$Payload`）

### 9. 下载文件延迟落盘

- **现象**：点击保存后立即检查 Downloads 文件夹，找不到新文件
- **解决**：点击保存后至少等待 5-6 秒

### 10. 批量下载文件顺序

- **现象**：`Get-ChildItem | Sort LastWriteTime -Desc | Select -First 1` 可能取到旧文件
- **原因**：上一条文件未移走，Downloads 中积压
- **解决**：每条下载完成后**立即**移动文件到 downloadDir

### 11. scrollBy 不可靠，mouseWheel 更可靠

- **现象**：`window.scrollBy(0, 500)` 不一定触发虚拟列表加载
- **解决**：使用 CDP `mouseWheel` 事件模拟真实滚动

### 12. textContent 中的换行与时间戳干扰

- **现象**：`trim()` + `indexOf(' ')` 提取姓名时，内部换行符导致提取不完整
- **解决**：用正则 `/^\S+/` 匹配第一个非空白序列，更可靠

### 13. 虚拟滚动导致循环姓名与实际点击不匹配

- **现象**：循环中获取的姓名列表与实际下载的文件名不一致（错位 1-2 个候选人）
- **原因**：关闭模态 → 重新滚动 → DOM 节点被 Vue 虚拟列表替换
- **影响**：可能出现重复下载同一候选人的情况
- **解决**：通过去重检测（姓名+年龄+大小）自动处理

### 14. 候选人列表加载延迟

- **现象**：点击岗位标签后立即收集姓名，size=0
- **原因**：Vue 组件渲染需要时间，候选人卡片不会立即可见
- **解决**：在收集姓名前主动等待 `.talent-basic-info__name` 元素数量 > 0

### 15. URL 中 `#` 片段在 PowerShell 字符串拼接中的问题

- **现象**：含 `#sortType=recommend` 的 URL 在拼接 JSON 时可能被 PowerShell 解析器当作注释
- **解决**：URL 中去掉 `#` 片段（仅保留查询参数 `?tab=recommend&jobNumber=XXX`），页面功能不受影响

### 16. PowerShell 单引号字符串内嵌 JS 的双引号陷阱（CRITICAL）

- **现象**：脚本中 `Invoke-Eval` 调用的 JS 表达式用了 `""button""` 包裹字符串字面量，导致 JS 端收到的是 `""`（空字符串）+ `button` + `""`（空字符串）—— 三个独立 token 连在一起，JS 报 `SyntaxError: missing ) after argument list`。但 `Invoke-Eval` 函数不检查响应是否成功，错误响应不匹配 `"value":"has"` 等正则，脚本静默判定为"找不到元素"，一直重试。
- **本质**：PowerShell 单引号字符串 `'...'` 中，`""` 是**两个双引号字面量**（不是转义），`''` 是**两个单引号字面量**。要表示一个双引号只需写一个 `"`，要表示一个单引号写两个 `''`。
- **影响**：所有用 `""...""` 包裹 JS 字符串的代码都会产生 JS 语法错误，表现为"存至本地按钮点击后检测不到保存对话框"、"关闭模态检测失效"等隐性故障，**日志看不到任何错误**，只有 `[FAIL]` 或 `[WARN]`。
- **验证方法**：单独通过 `curl` 调用 `evaluate` 传这段 JS，会返回 `{"ok":false,"error":{"code":"extension_error","message":"evaluate: SyntaxError: ..."}}`。
- **解决**：
  ```powershell
  # ❌ 错误："" 是两个双引号，JS 端收到 ""button"" → SyntaxError
  $js = '(()=>{const b=document.querySelectorAll(""button"");...})()'

  # ✅ 正确：单个 " 即可，JS 端收到 "button"
  $js = '(()=>{const b=document.querySelectorAll("button");...})()'

  # ✅ 正确：也可以用单引号 ''（PS 单引号字符串中 '' 表示一个单引号）
  $js = '(()=>{const b=document.querySelectorAll(''button'');...})()'
  ```
- **修复清单**（run.ps1 中所有需要修正的位置）：
  - `Close-ModalIfOpen` 函数：`"".km-modal--open""` → `".km-modal--open"`
  - 3.4 阶段重试探测：`"".resume-button.position-r""` → `".resume-button.position-r"`
  - 3.4 阶段对话框检测：`""button"`、`""\u4fdd\u5b58""`、`""has""`、`""no""` → 单个双引号包裹

### 17. Invoke-Eval 不检查响应错误（隐性 bug 来源）

- **现象**：`Invoke-Eval` 只返回原始响应字符串，不检查 `ok` 字段。当 JS 报语法错误或运行时错误时，响应是 `{"ok":false,"error":{...}}`，不含 `"value"` 字段，正则匹配全部失败，调用方静默走 fallback/重试分支。
- **影响**：任何 JS 语法错误都会被掩盖为"找不到元素"，极难排查。
- **建议**：在 `Invoke-Eval` 中增加错误检查，或至少在调用方对 `"ok":false` 做日志输出：
  ```powershell
  function Invoke-Eval {
      param([string]$Code)
      $r = Send-Web -Action 'evaluate' -Payload @{ code = $Code }
      if ($r -match '"ok":false') { Write-Host "  [JS ERROR] $r" -ForegroundColor Red }
      return $r
  }
  ```

### 18. DOM.getBoxModel 在未注册节点上返回 nodeId=0

- **现象**：`Get-SaveButtonCoords` 函数用 `Runtime.evaluate`（returnByValue=false）→ `DOM.requestNode` → `DOM.getBoxModel` 三步获取按钮坐标。但 `DOM.requestNode` 对未通过 `DOM.getDocument`/`DOM.enable` 注册的节点返回 `nodeId: 0`，后续 `DOM.getBoxModel` 失败。
- **影响**：`Get-SaveButtonCoords` 总是返回 `$null`，3.6 阶段一直 fallback 到 config 中的静态坐标。在多数情况下静态坐标恰好正确，所以"看起来能用"，但任何按钮位置变化都会立即失败。
- **解决**：改用 `getBoundingClientRect()` 直接在 JS 中获取坐标（CSS 像素），与 `CDP Input.dispatchMouseEvent` 使用的坐标系一致：
  ```powershell
  function Get-SaveButtonCoords {
      $r = Invoke-Eval '(()=>{const b=document.querySelectorAll("button");for(const x of b){if(x.textContent.trim()==="\u4fdd\u5b58"&&x.offsetWidth>0){const r=x.getBoundingClientRect();if(r.width<=0||r.height<=0)continue;return Math.round(r.x+r.width/2)+","+Math.round(r.y+r.height/2)}}return "0,0"})()'
      return Parse-Coords $r
  }
  ```
- **注意**：`DOM.getBoxModel` 返回的可能是**设备像素**（受 `devicePixelRatio` 影响），而 `CDP Input.dispatchMouseEvent` 使用 **CSS 像素**。两者混用会导致坐标偏移（如 1.25 倍缩放下偏差 25%）。

### 19. devicePixelRatio 与 CDP 坐标系

- **关键事实**：`CDP Input.dispatchMouseEvent` 的 `x`/`y` 参数使用 **CSS 像素**（与 `getBoundingClientRect()` 一致），不是设备像素。
- **影响**：`DOM.getBoxModel` 返回设备像素，若直接用于 `Input.dispatchMouseEvent`，在 `devicePixelRatio != 1` 的环境中坐标会偏移。
- **验证**：`window.devicePixelRatio` 在 125% 缩放下为 1.25。按钮 CSS 坐标 (996, 561) 对应设备像素 (1245, 701)。
- **解决**：统一使用 `getBoundingClientRect()`（CSS 像素）获取所有点击坐标。

### 20. "存至本地"按钮等待不充分导致探测失败

- **现象**：点击候选人姓名后，详情面板的 Vue 渲染需要时间。原代码只等待 5 次 × 500ms = 2.5s，且检测条件只看元素存在不看可见性（`offsetWidth > 0`）。
- **解决**：延长到 15 次 × 500ms = 7.5s，检测条件改为 `getBoundingClientRect().width > 0 && height > 0`，并在等待过程中偶发滚动（每 3 轮滚动一次）触发虚拟列表重新渲染。

### 21. Word 格式下载："保存到本地"对话框中的文件格式选择

- **背景**：智联 HR 端的"保存到本地"弹窗（`.km-modal--open` 内的子对话框）提供两个文件格式选项：`pdf`（默认选中）和 `word`（生成 `.docx`）。
- **关键事实**：必须**在点击"保存"按钮之前**点击"word"图标，否则下载的就是 PDF。
- **定位 "word" 选项**：该选项是一个 `<div>`，其直接子节点唯一文本就是 `"word"`（trim 后严格相等）。可用 JS 直接定位：
  ```javascript
  const els = document.querySelectorAll('div');
  for (const el of els) {
    if (el.childNodes.length === 1 &&
        el.childNodes[0].nodeType === 3 &&
        el.textContent.trim() === 'word') {
      const r = el.getBoundingClientRect();
      // cx,cy = r.x + r.width/2, r.y + r.height/2
    }
  }
  ```
- **切换验证**：切换成功后，对话框下方会出现"支持 word 2010 及以上版本"提示文字。可用以下 JS 验证：
  ```javascript
  const ps = document.querySelectorAll('p');
  for (const p of ps) {
    if (p.textContent.indexOf('支持 word') >= 0) return 'yes';
  }
  return 'no';
  ```
- **CDP 点击序列**：与 PDF 切换相同——`mouseMoved` → 等待 500ms → `mousePressed` → 等待 150ms → `mouseReleased` → 等待 800ms。**不能用 JS `.click()`**，Vue 组件需要真实鼠标事件。
- **时长调整**：Word 文件通常比 PDF 大（实测 80KB+ vs 30KB），生成耗时更长：
  - `DownloadWaitMs` 从 6000 提升到 10000（Word 模式自动）
  - 下载检测窗口从 15s 提升到 25s（Word 模式自动）
- **去重要点**：Word 文件后缀是 `.docx`，脚本中的 `\.pdf$` 正则匹配和 `*.pdf` 过滤器都要参数化为 `$Config.FileExt`，避免混合 PDF/Word 时误判重复。

### 22. 中文通配符在 PowerShell `-Filter` 中的匹配问题（CRITICAL — 导致"下载后文件没转移"）

- **现象**：简历已成功下载到 `Downloads`，但脚本的检测/移动环节找不到文件，文件一直留在 Downloads，未被转移到目标目录。
- **本质**：
  - PowerShell 的 `Get-ChildItem -Filter "*智联*"` 中文通配符匹配依赖系统代码页，行为不稳定。
  - 尤其当脚本通过**命令行内嵌中文参数**执行时（`powershell -Command "... -Filter '*智联*' ..."`），CLI 传输层会把中文字符破坏成乱码，`-Filter` 条件彻底失效，静默返回空集，`Move-OneResume` 判定"找不到文件"。
  - 实测：`cmd /c dir "%USERPROFILE%\Downloads\*智联简历*.docx"` 能匹配到 26 个文件，而 PowerShell 命令行内嵌中文 `-Filter` 返回 0；但**在 UTF-8 BOM 的 .ps1 文件内**，`-Include` 和 `-Filter` 都能正确匹配中文文件名。
- **影响**：下载文件无法被识别和转移 → 全部积压在 Downloads。这是"下载了很多简历但没转移到指定文件夹"的**核心根因**。
- **解决**：
  1. **统一改用 `-Include` + `-Path "$dir\*"`**，不依赖 `-Filter` 的中文通配符：
     ```powershell
     # ✅ 正确（本 skill 已统一采用）
     Get-ChildItem -Path "$DownloadSource\*" -Include "*智联简历*.$FileExt"
     # ❌ 避免：中文通配符在部分环境匹配失败
     Get-ChildItem $DownloadSource -Filter "*智联简历*.$FileExt"
     ```
  2. **运行方式**：优先用 `-File scripts\run.ps1`（UTF-8 BOM 文件），避免命令行内嵌中文参数。命令行传中文参数（如 `-JobName`、`-DownloadDir`）可能被 CLI 破坏。
  3. 脚本内中文匹配全部放在 UTF-8 BOM 的 .ps1 文件中，不通过命令行传入。

### 23. `Move-OneResume` 返回值语义与"误判重复"陷阱

- **现象**：下载成功后脚本把 `Move-OneResume` 的 `$null` 返回当成"重复"，`ok` 不递增且标记候选人已尝试，文件留在 Downloads，但日志只显示 `[DUP]`，掩盖了真实问题。
- **根因**：原 `Move-OneResume` 在"Downloads 找不到文件"和"发现重复"两种情况下都返回 `$null`，调用方无法区分，且重复判定只比对"同名+文件大小"，未校验年龄。
- **修复**：
  1. 明确返回值语义：`'ok'`（移动成功）/ `'dup'`（完全相同的简历已存在，删除新文件）/ `$null`（未找到文件或移动失败）。
  2. 调用方区分处理：`$null` 时输出 `[WARN]` 并保留文件供人工处理，不再静默当作重复。
  3. 重复判定用 `Test-DuplicateFile` 做**姓名+年龄+文件大小**三重验证。
- **经验**：任何返回 `$null`/`$false` 的分支，若代表多种不同情况，必须显式区分并在日志中说明，否则会掩盖真实故障。

### 24. FileFormat 配置与网页实际下载格式不一致导致文件无法转移（CRITICAL）

- **现象**：简历已成功下载到 `Downloads`（如 `张先生_37岁_智联简历_86023.docx`），但脚本的 `Move-OneResume`、下载检测、去重检查等环节都找不到文件，文件留在 Downloads 未被转移。
- **根因**：智联的"保存到本地"对话框会**记住用户上次选择的文件格式**。如果用户上次选了"word"，即使 `config.ps1` 中 `FileFormat = "pdf"`，网页端实际下载的仍是 `.docx` 文件。而脚本中所有文件搜索逻辑（`Move-OneResume`、下载检测、`Test-DuplicateFile`、`Test-NameExistsInDir`、汇总输出）都硬编码使用 `$Config.FileExt`（来自 `FileFormat` 派生，如 `.pdf`），导致 `Get-ChildItem -Include "*智联简历*.pdf"` 完全匹配不到 `.docx` 文件 → 返回空集 → 被当作"找不到文件"处理。
- **影响**：文件搜索层与网页实际下载格式脱钩，导致整个文件转移管线失效。且这个问题对用户完全透明——脚本输出 `[FAIL] Download did not complete` 或 `[WARN] No file matched/moved`，但文件确实已下载成功。
- **解决**：
  1. **文件搜索层不再依赖 `$Config.FileExt`**：所有 `Get-ChildItem -Include` 改为同时匹配 `"*智联简历*.pdf", "*智联简历*.docx"` 两种后缀。
  2. **正则匹配也双格式**：`Test-DuplicateFile` 中的文件名正则从 `\.pdf$` 改为 `\.(pdf|docx)$`。
  3. **`FileFormat` 保留**用于其他逻辑（如 word 模式下的等待时间调整、对话框格式切换），但不影响文件搜索。
  4. 修改涉及 6 处（`Test-DuplicateFile`、`Move-OneResume`、`Test-NameExistsInDir`、3.6 下载检测、汇总输出、config 注释）。
- **经验**：文件搜索模式不应与用户偏好配置强绑定——下载产物的实际格式受外部系统状态（如网页 cookie/localStorage 记忆）影响，搜索层应使用宽松的格式匹配策略。

### 25. 岗位 Tab 未找到/误匹配导致下载错误岗位的候选人（CRITICAL）

- **现象**：页面 URL 进入后默认激活的是上次浏览遗留的选中 Tab（如"联想渠道经理-青岛"），而用户指定的岗位 Tab（如"销售部经理"）未被正确切换，脚本静默降级直接用当前激活 Tab 的候选人列表继续下载。
- **根因**：
  1. 原代码在快照中找不到目标岗位 Tab 时，只输出 `[WARN]` 但**不终止执行**，继续用当前激活的（错误）岗位的候选人列表
  2. **完全没有岗位验证步骤**——点击 Tab 后不检查实际激活的岗位是否与参数一致
  3. 页面 Vue 渲染需要时间，首次 snapshot 可能还未包含完整的岗位标签列表
- **修复**（`run.ps1` 阶段 1 改造）：
  1. **渲染等待 + 重试 snapshot**：首次 snapshot 找不到目标 Tab 时，等待 1s 后重试（最多 8 次），给 Vue 渲染留足时间
  2. **找不到即终止**：8 次重试后仍然找不到目标岗位 Tab → 输出 `FATAL` 错误并 `exit 1`，不再静默降级
  3. **⭐ 岗位验证步骤（新增阶段 2.1）**：点击岗位标签后，通过 JS 多种策略检测当前激活/高亮的岗位名称：
     - 策略1：`[aria-selected="true"]` 元素的文本
     - 策略2：岗位容器（`.job-pane` 等）内的 `.is-active`/`.active` 类元素
     - 策略3：`[aria-current="page"]` 或 `[aria-current="true"]`
     - 策略4：遍历 `[class*="tab"]`/`[role="tab"]` 等元素，检查高亮样式类
  4. **匹配校验**：将检测到的激活岗位名称与 `jobName` 参数比较（支持去除括号内数量信息如 `(30)` 后的模糊匹配）：
     - 匹配成功 → 继续下载流程
     - 匹配失败 → 输出 `FATAL: Job mismatch detected!` 并 `exit 1`
     - 无法检测 → 输出 `[WARN]` 但继续（兼容极端情况）
- **经验**：依赖页面快照查找 UI 元素时，必须考虑渲染时序，增加重试机制。任何"找不到"的 fallback 路径都不应该静默降级到错误数据源。

### 26. `run.ps1` 缺少 UTF-8 BOM 导致 PowerShell 解析错误（CRITICAL）

- **现象**：执行 `run.ps1` 时报 `Unexpected token '}' in expression or statement`，错误指向正常的花括号位置（如第 58 行、第 138 行、第 446 行等）。
- **根因**：`run.ps1` 文件缺少 UTF-8 BOM 头。在中文 Windows + PowerShell 5.x 环境下，无 BOM 的 `.ps1` 文件被以 GBK（系统 ANSI 代码页 936）解析。文件中包含的中文注释字符被误解析，导致语法结构被破坏——例如中文注释中的某个字节恰好是 `}` 的 GBK 编码片段，PowerShell 将其识别为多余的右花括号。
- **影响**：脚本完全无法运行，所有流程卡在第一步。
- **验证方法**：检查文件前 3 字节是否为 `0xEF 0xBB 0xBF`（UTF-8 BOM 标识）。
  ```powershell
  $bytes = [System.IO.File]::ReadAllBytes("run.ps1")
  if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
      Write-Output "Has BOM"
  } else {
      Write-Output "NO BOM - NEED FIX"
  }
  ```
- **修复方法**：用 `System.Text.UTF8Encoding($true)` 重新写入文件（`$true` 表示生成 BOM）：
  ```powershell
  $content = [System.IO.File]::ReadAllText("run.ps1")
  $utf8Bom = New-Object System.Text.UTF8Encoding $true
  [System.IO.File]::WriteAllText("run.ps1", $content, $utf8Bom)
  ```
- **AI Agent 执行准则**：在执行任何 `.ps1` 脚本前，应自动检查并修复 BOM。`run.ps1` 是主要入口，`webbridge-utils.ps1` 和 `download-loop.ps1` 通常已有 BOM，但 `run.ps1` 由于频繁修改可能丢失 BOM。

### 27. `navigate` 新标签页创建后浏览器 UI 未自动聚焦（用户需手动切换）

- **现象**：脚本调用 `navigate` 创建新标签页（`newTab:true`）后，WebBridge 层面已将新标签页设为 "current tab"，后续 `snapshot`/`click`/`evaluate` 等操作都正确作用于新标签页。但**浏览器 UI 层面没有自动切换到新标签页**——用户看到的仍然是之前的标签页，必须手动点击标签栏切换到新标签页才能看到脚本的操作过程。`list_tabs` 确认新标签页的 `active` 字段为 `false`。
- **根因**：WebBridge 的 `navigate` 内部没有调用 `chrome.tabs.update(tabId, {active: true})` 来聚焦浏览器 UI，导致 Chrome 窗口中的标签页切换未发生。
- **影响**：用户体验差——脚本在后台正常运行，但用户看不到进度，容易误以为脚本没有工作。
- **失败的方案及原因**：
  1. ❌ **`find_tab` + CDP `Target.activateTarget`** — `find_tab` 的 URL 匹配是域名/路径前缀匹配，不精确。当 session 中有多个同域名标签页时（如不同 jobNumber），会错误地切回旧标签页。且 `Target.activateTarget` 被 Chrome 安全策略阻止（`Not allowed`），`list_tabs` 返回的 `tabId` 是 `chrome.tabs` 的 id（数字），不是 CDP targetId。
  2. ❌ **`evaluate` 中调用 `chrome.tabs.update`** — `evaluate` 运行在 content script 上下文中，无法访问 `chrome.tabs` API。
  3. ❌ **CDP `Browser.getVersion` 等 Browser 级方法** — WebBridge 的 CDP 透传不支持 Browser 级 domain。
- **正确方案**（已应用到 `run.ps1` 阶段 1）：**CDP `Page.bringToFront`**。`navigate` 之后 WebBridge 的 CDP 已连接到新创建的标签页，此时调用 `Page.bringToFront` 会将当前 CDP 连接的标签页带到浏览器最前面：
  ```powershell
  # Step 1: navigate 创建新标签页
  Send-Web -Action 'navigate' -Payload @{
      url = $Config.Url
      newTab = $true
      group_title = 'Zhaopin Resume Screening'
  }
  Wait 3000

  # Step 2: CDP Page.bringToFront 强制浏览器切换到当前 CDP 连接的标签页
  Send-Web -Action 'cdp' -Payload @{
      method = 'Page.bringToFront'
      params = @{}
  }
  Wait 3000
  ```
- **原理**：`navigate(newTab=true)` 做了两件事：1) 创建 Chrome 标签页 2) 通过 `chrome.debugger.attach` 连接 CDP 并将该标签页设为 session 的 current tab。此时 WebBridge 的 CDP 通道已经指向新标签页，`Page.bringToFront` 直接作用于正确的标签页，无需额外的 tabId 匹配。
- **经验**：WebBridge 中 CDP 连接的生命周期与 session current tab 绑定，利用 `Page.bringToFront` 是当前最可靠的标签页聚焦方案。不要试图用 `find_tab`（URL 模糊匹配不可靠）或 `Target.activateTarget`（被安全策略阻止）。

### 28. 阶段 2 snapshot 精确匹配在岗位名带后缀时失败（CRITICAL — 导致聚焦后卡住）

- **现象**：`navigate` + `bringToFront` 成功聚焦新标签页后，脚本没有继续操作，用户只看到标签页打开但没有任何动作。实际是脚本在阶段 2 找不到岗位 Tab 后 `exit 1` 静默退出了。
- **根因**：阶段 2 的 snapshot 匹配使用精确正则：
  ```powershell
  $snap -match '"name":"' + [regex]::Escape($Config.JobName) + '","ref":"(@e\d+)"'
  ```
  这要求 snapshot 中 link 的 `name` **精确等于** `$Config.JobName`。但在不同电脑/不同账号上，同一个岗位可能在 name 中带后缀（如 `"渠道部经理·协作未上线"` 而非纯净的 `"渠道部经理"`）。精确匹配直接失败。
- **修复**：改为双路径策略（每 2s 一轮，最多 15 轮/30s）：
  - 路径 A：snapshot **包含匹配** `'"name":"[^"]*渠道部经理[^"]*","ref":"(@e\d+)"'` — 允许 name 前后有任意字符
  - 路径 B：JS evaluate 模糊匹配 — 用 `textContent.includes(jobName)` 查找 `.job-pane__item` 并 `.click()`
  - 两路径交叉进行，任一命中即成功
  - 失败时输出页面上所有岗位标签名称帮助诊断
- **相关**：阶段 2.1 的岗位验证已使用包含匹配（`-like "*$expectedClean*"`），所以点击后即使激活岗位带后缀也能通过验证。

### 29. 阶段 2.1 岗位验证策略 2 的 `.active` 选择器误匹配导航栏"推荐"（CRITICAL — 导致验证失败退出）

- **现象**：即使阶段 2 成功点击了正确的岗位 Tab，阶段 2.1 的岗位验证 JS 仍然返回 `S2:推荐`，与 `JobName` 不匹配，触发 `FATAL: Job mismatch detected!` → `exit 1`。用户看到标签页聚焦后短暂闪了一下就停了。
- **根因**：岗位验证 JS 的策略 2（`run.ps1` 原第 269 行）：
  ```javascript
  el = document.querySelector('.is-active, .active, [class*="is-active"], [class*="active"]');
  ```
  选择器 `.active` 太宽泛。智联页面左侧导航栏的"推荐" Tab 也有 `.active` 类名，且它在 DOM 中出现得更早，`document.querySelector` 总是先匹配到它。返回的文本是 `"推荐"` 而非岗位名称。
- **关键发现**：页面岗位标签使用 class `job-pane__item job-pane__item--active`，而非通用的 `.active` 或 `.is-active`。`document.querySelector('.job-pane__item--active')` 直接命中正确的激活岗位。
- **修复**：策略优先级重排：
  1. 策略 1（新增）：`document.querySelector('.job-pane__item--active')` — 精确选择器，直接命中岗位标签栏的激活 class
  2. 策略 2（修正）：在岗位容器 `.job-pane` 内查找激活元素（限定作用域）
  3. 策略 3-5：保持原有逻辑作为兜底
  4. 策略 5 的候选列表也优先遍历 `.job-pane__item`，并增加 `job-pane__item--active` class 检测
- **经验**：CSS class 选择器 `.active` 和 `.is-active` 在 SPA 页面中极易误匹配导航组件。岗位标签这类业务特定元素应使用其专属 class（如 `.job-pane__item--active`），并限定在业务容器内查找。

## 文件命名规则

`{姓名}_{年龄}岁_智联简历_{5位随机数}.{pdf|docx}`

示例：
- `左先生_37岁_智联简历_11618.pdf`
- `董先生_34岁_智联简历_29335.pdf`
- `梁女士_43岁_智联简历_28378.pdf`
- `王先生_25岁_智联简历_88629.docx`

## Agent 执行经验（2026-08-04 实战总结）

### 30. `ConvertTo-JsSafeName` 函数定义在调用之后导致 `CommandNotFoundException`（CRITICAL）

- **现象**：`run.ps1` 执行时报 `ConvertTo-JsSafeName : The term 'ConvertTo-JsSafeName' is not recognized as the name of a cmdlet`
- **根因**：PowerShell 是顺序解释执行的脚本语言，函数必须先定义后使用。`ConvertTo-JsSafeName` 原定义在第 551 行，但阶段 1 的第 188 行就调用了它（`$safeJobName = ConvertTo-JsSafeName -Name $Config.JobName`）。
- **修复**：将函数定义移至第 150 行（`Get-SaveButtonCoords` 之后、阶段 1 之前），并删除第 565 行的重复定义。同时在函数上方添加 `★★★ CRITICAL` 注释警告未来维护者不要将其移回后半部分。
- **经验**：在 PowerShell 中定义工具函数时，应集中放在脚本开头（参数解析之后、主流程之前），类似 C 语言的函数声明前置。

### 31. cmd 命令行传递 URL 参数时 `&tab=` 被误解析为命令分隔符

- **现象**：`cmd /c "powershell ... -Url "https://...?jobNumber=XXX&tab=recommend" ..."` 报错 `'tab' 不是内部或外部命令`
- **根因**：cmd 中 `&` 是命令分隔符。即使 URL 在双引号内，外层 `cmd /c` 仍可能在解析时将 `&tab=` 截断为两条命令。
- **修复**：Agent 应**优先通过 `config.json` 传参**，不依赖命令行参数传递 URL：
  ```powershell
  # 正确做法：先写 config.json，再直接运行（无命令行参数）
  powershell -ExecutionPolicy Bypass -File "$SKILL_DIR\scripts\run.ps1"
  ```
- **备选方案**：如果必须在命令行传 URL，在 cmd 中用 `^&` 转义（`^&tab=recommend`），或在 PowerShell 中直接执行而非通过 `cmd /c` 中转。

### 32. 全局 skill 目录下的 `_user_meta.json` 导致 skill 不被加载

- **现象**：两个 zhaopin skill 安装到 `~/.codebuddy/skills/` 后，CodeBuddy 的可用 skill 列表中找不到它们，`use_skill` 工具无法触发。
- **排查**：对比其他正常工作的 skill（`kimi-webbridge`、`ppt-master`），发现只有这两个 skill 目录下有 `_user_meta.json` 文件（内容 `{"name":"...","installedAt":...,"source":"userImport"}`）。
- **根因**：该文件是 user-import skill 的安装元数据。如果 CodeBuddy 的 skill 加载器在解析该文件时出现问题（JSON 格式不兼容、字段校验失败等），可能静默跳过该 skill。
- **修复**：删除 `_user_meta.json` 后重启 CodeBuddy 窗口，skill 被重新扫描并正常加载。
- **经验**：如果 skill 安装后不生效，检查 skill 目录下是否存在 `_user_meta.json`，尝试删除并重载。
