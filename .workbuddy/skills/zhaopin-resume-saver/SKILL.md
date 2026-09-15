---
name: zhaopin-resume-saver
description: |
  从智联招聘（rd6.zhaopin.com）推荐页面自动批量浏览并保存候选人简历（Word/DOCX）。
  This skill should be used when the user wants to:
  - 批量下载智联招聘指定岗位下的候选人简历
  - 自动浏览并保存某个岗位的候选人简历到本地
  - 从智联招聘推荐页面按岗位筛选后保存简历
  关键参数：岗位名称（jobName）、下载数量（downloadCount，默认 5）、目标保存目录（downloadDir）。
agent_created: true
---

# 智联招聘简历批量下载

通过 Kimi WebBridge 操控真实浏览器，自动浏览智联招聘推荐页面中指定岗位下的候选人简历，逐条点击"存至本地"并保存为 Word（DOCX），完成后自动移动至用户指定的目标文件夹。

## 🚀 执行原则（最高优先级）

- **整个下载流程完全自动化，不要在每步之间询问用户"是否继续"**。只在最终输出汇总结果。
- 所有 `execute_command` 调用一律使用 `requires_approval: false`。
- 参数缺失时，**一次性列出所有缺失项**让用户补充，不要逐项追问。
- **与 `zhaopin-resume-screening` 组合使用时**：下载完成后应无缝衔接评估流程，不得停顿等待用户催促。Agent 应在下载脚本运行期间预计算评估参数（SKILL_DIR、Python 路径、输出路径等）。

## 🚨 环境前置检查（2026-09-15 新增，Agent 必读）

### A. 代理劫持 curl（HTTP_PROXY 环境变量）
- 若环境设置了 `HTTP_PROXY/HTTPS_PROXY`（如 `http://127.0.0.1:12887`），`curl.exe` 访问 `http://127.0.0.1:10086` 会被代理拦截，返回极具迷惑性的 `{"ok":..., "error":{"message":"upstream connect failed ... os error 10061"}}` —— **这是代理的错误响应，不是 daemon 的**。
- **修复**：运行 `run.ps1` 前在同一会话设置 `$env:NO_PROXY = "127.0.0.1,localhost"`；Agent 手动 curl 调试时一律加 `--noproxy "*"`。
- 诊断特征：`netstat` 查不到 10086 监听但 curl 仍返回 JSON 错误 → 响应来自代理。

### B. daemon 随父进程被清理
- 由短生命周期的 Agent 命令（如一次性 `kimi-webbridge.exe start`）启动的 daemon，可能在命令退出后被连带清理（表现为：启动日志正常、扩展握手成功，但约 1 分钟后端口无监听、进程消失，日志无任何崩溃记录）。
- **修复**：用 `run_in_background` 的 PowerShell 常驻任务承载 daemon（`start` 后接 `Start-Sleep`），保证 daemon 父进程存活。
- 注意 `run.ps1` 末尾的 `Stop-BrowserAutomation` 会执行 `daemon stop`——这是预期行为，后续如需浏览器需重新启动 daemon。

### C. WebBridge 请求临时文件的编码
- PowerShell 5.1 的 `Out-File -Encoding utf8` 写出**带 BOM** 的文件，daemon 解析 JSON 会报 `invalid character '茂'`（BOM 被当数据）。
- **修复**：请求体一律用 `[System.IO.File]::WriteAllText($path, $json, [System.Text.UTF8Encoding]::new($false))` 写无 BOM UTF-8。
- curl 响应含中文时，**不要**用 `| Out-String` 捕获（PS 按 GBK 解码导致乱码双倍破坏）；用 `curl.exe --output 文件` 直写文件后用 Read 工具读取。

### D. daemon 版本与扩展版本必须匹配
- daemon v1.11.3 + 扩展 v2.0.9 握手成功（`hello from extension v2.0.9 (daemon v1.11.3)`）但请求异常。执行 `kimi-webbridge upgrade` 对齐后正常。

### ⭐ E. 标准启动序列（已代码化为 `Initialize-WebBridgeEnv`，此节保留作排障知识）

> **2026-09-15 #56 起，A-D 四项自检已内建于 run.ps1 的 [0] 阶段**（失败 exit 3），
> Agent 无需再手工执行本序列；但在**计划任务 wrapper** 中仍需按此顺序写代码（wrapper 先于 run.ps1 运行）。

以上 A-D 四个坑在 2026-09-15 实战中全部踩过、合计浪费约 40 分钟排查。**不要逐个踩完再修**，启动前一次性按下面顺序做完：

```powershell
# ① 版本对齐（防坑 D）：先 stop 再 upgrade，幂等
& "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" stop 2>$null
& "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" upgrade

# ② 常驻承载 daemon（防坑 B）：必须用独立进程树（计划任务）承载，
#    命令 = start + 长时间 Start-Sleep（如 7200 秒），保证父进程存活
& "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start 2>&1 | Out-String; Start-Sleep -Seconds 7200

# ③ 就绪探测（防坑 A）：等扩展重连（约 10-30 秒），循环探测必须 --noproxy
#    ★ 必须用 list_tabs 真探测——daemon 活着但扩展未连接时 snapshot 返回业务错误
#      "no tab"，会把未就绪误判为就绪（#55 排障实测，扩展掉线后 round 6 才重连）
$ok = $false
for ($i=1; $i -le 40; $i++) {
    Start-Sleep -Seconds 5
    $probe = '{"action":"list_tabs","args":{},"session":"probe"}'
    [System.IO.File]::WriteAllText("$env:TEMP\wb-probe.json", $probe, [System.Text.UTF8Encoding]::new($false))
    curl.exe -sS --noproxy "*" -X POST http://127.0.0.1:10086/command -H "Content-Type: application/json" --data-binary "@$env:TEMP\wb-probe.json" --output "$env:TEMP\wb-probe-res.json"
    if (Select-String -Path "$env:TEMP\wb-probe-res.json" -Pattern '"ok":true' -Quiet) { $ok = $true; break }
}
if (-not $ok) { throw "WebBridge extension not connected after 200s" }

# ④ 启动下载（防坑 A/B）：计划任务内设 NO_PROXY 后运行 run.ps1
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:NO_PROXY = "127.0.0.1,localhost"; $env:no_PROXY = "127.0.0.1,localhost"
& "$SKILL_DIR\scripts\run.ps1"
```

- ③ 探测通过 = daemon + 扩展 + 版本 + 代理四项全部就绪，此后再启动下载脚本。
- **运行中脚本热改无效**：修改 run.ps1/lib 后必须 stop 任务 → 重启才生效（PS 启动时已把脚本解析进内存）。

### ⭐ F. 中断续传（run.ps1 已内置，#52/#53/#56）

下载中断（手动停止/故障）后重启会自动续传，无需手工干预：

- 脚本启动时扫描 `DownloadDir` 已有简历，按 `姓名_年龄`（从文件名解析）预填 `$prefilled` 基础表，与主循环 `$processed` 完整表双表过滤，日志输出 `[RESUME] N resume(s) already in target dir`。
- **Agent 唯一要做的事**：重启前把 `config.json` 的 `DownloadCount` 改为 `目标总数 - 目录已有份数`（如 100-87=13），否则会超额下载。
- 同名不同人风险由 `Move-OneResume` 的 DUP 三重验证（姓名+年龄+文件大小）兜底；实测 45 份文件名无重名。
- 主循环中 `[SKIP]`（点击失败登记跳过）、`[TOP-UP]`（到底回顶重扫）、`RECOVER`（session 自愈）都是**预期自愈行为，不是故障**，不要手动干预。

## 参数

| 参数 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `url` | string | 是 | — | 智联招聘推荐页完整 URL（含 `jobNumber`），去掉 `#` 片段 |
| `jobName` | string | 是 | — | 岗位名称，如"AI产品销售"，需与页面上职位标签完全一致 |
| `downloadDir` | string | 是 | — | 下载完成后简历的最终保存目录（绝对路径） |
| `downloadCount` | number | 否 | 5 | 期望下载的唯一简历数量 |
| `fileFormat` | string | 否 | `"word"` | 下载文件格式，可选 `"word"`（默认，生成 .docx）或 `"pdf"` |

## ⚠️ 强制参数校验（最高优先级）

**本 Skill 启动时必须检查以下三个参数是否全部提供。若任一缺失，立即停止并提示用户补充。**

| 参数 | 必填 | 说明 |
|------|------|------|
| `url` | **是** | 智联招聘推荐页完整 URL（含 `jobNumber`），如 `https://rd6.zhaopin.com/app/recommend?tab=recommend&jobNumber=XXX` |
| `jobName` | **是** | 岗位名称，需与页面标签完全一致，如"销售部经理" |
| `downloadDir` | **是** | 简历保存目录（绝对路径），如 `C:\Resumes\销售部经理` |
| `downloadCount` | 否 | 下载数量，默认 5 |
| `fileFormat` | 否 | 下载格式：`"word"`（默认，生成 `.docx`）或 `"pdf"` |

### 参数校验提示模板

当用户未提供全部参数时，输出以下提示：

```
请提供以下必要信息后我才能开始下载简历：

1. **智联推荐页 URL**：含 jobNumber 的完整链接。
2. **岗位名称**：页面上显示的岗位标签文字。
3. **保存目录**：简历文件存放的文件夹绝对路径。
4. **下载数量**（可选）：默认 5 份。

示例：
- URL：https://rd6.zhaopin.com/app/recommend?tab=recommend&jobNumber=CC136786060J40995378602
- 岗位：销售部经理
- 保存目录：C:\Resumes\销售部经理
- 数量：30
```

## ⚠️ Skill 目录变量（最高优先级）

**本 Skill 可能安装在全局目录或项目本地目录。执行任何脚本前，必须先用本 SKILL.md 文件的实际路径推导 `SKILL_DIR`：**

```
SKILL_DIR = 本 SKILL.md 所在目录的绝对路径
```

- 脚本位于 `$SKILL_DIR/scripts/` 下
- 所有 `powershell -File` 和 Python 调用必须使用 `$SKILL_DIR/scripts/xxx` 的绝对路径
- **禁止使用相对于项目根目录的路径**（如 `.\workbuddy\skills\...`），因为 skill 不一定在项目目录下

## 前置条件

1. 用户的真实浏览器已安装 Kimi WebBridge 扩展并连接成功
2. 浏览器已登录智联招聘企业账号（HR 端），推荐页面可正常访问
3. 守护进程运行中（`~/.kimi-webbridge/bin/kimi-webbridge start`）
4. **所有 `.ps1` 脚本文件必须使用 UTF-8 BOM 编码**（中文 Windows 强制要求）


## 脚本架构（2026-09-15 模块化重构，#56）

**固定流程已全部沉淀为脚本；本文档只保留经验。改流程 = 改脚本，不是改文档。**

| 文件 | 职责 |
|---|---|
| `scripts/run.ps1` | 主编排（薄层）：配置校验 → 环境自检 → 导航/岗位选择/验证 → 下载主循环 → 汇总。所有流程步骤在此串联 |
| `scripts/lib/wb-core.ps1` | 通用层：WebBridge 客户端（Send-Web/Invoke-Eval/Invoke-Click/Invoke-CDP）、Write-Log 结构化日志（UTF-8 无 BOM 落盘）、Invoke-WithRetry / Wait-Until 稳定性原语、Initialize-WebBridgeEnv 环境自检、Stop-BrowserAutomation（幂等）、标签页管理（Switch-ToDetailTab/Clear-StaleDetailTabs）、ConvertTo-JsSafeName |
| `scripts/lib/zhaopin-page.ps1` | 智联页面层：Select-JobTab 双路径选岗（#28）、Get-ActiveJobName 验证（#38）、卡片提取 + 三重校验标记（#53/#54）、Close-ModalIfOpen、存至本地/word/保存序列（#44/#45/#46/#55）、Get-ZhaopinFiles 结构化文件识别 + Move-OneResume 移动去重 |
| `scripts/config.json` | 运行参数（数据，非逻辑） |
| `assets/*.txt` | 可复用 JS 模板（参考用，与 lib 内实现同源） |
| `references/tech_details.md` | 页面结构、选择器、CDP 坐标、完整踩坑记录 |

### 稳定性机制（内建于脚本，#56）

1. **环境自检** `Initialize-WebBridgeEnv`：设 NO_PROXY → daemon 端口检查/启动 → **`list_tabs` 真探测**（snapshot 在"daemon 活着但扩展未连接"时返回业务错误 "no tab"，会糊弄探测——必须用 list_tabs）等扩展重连（最多 200s）
2. **结构化日志** `Write-Log`：时间戳分级（INFO/OK/WARN/FAIL/STEP）写 `<DownloadDir>\_run_log.txt`（UTF-8 无 BOM）——替代 `*>` 重定向产生 UTF-16 文件的历史痛点
3. **deadline 制等待** `Wait-Until`：所有关键等待从"固定次数×固定间隔"改为超时制，单个等待永不过期也永不失控
4. **单轮异常保护**：主循环 try/catch——单轮未预期异常记日志+关面板+继续，连续 5 轮才终止（防异常风暴杀死整任务）
5. **断点续传**：启动扫描 DownloadDir，`$prefilled`（姓名_年龄）+ `$processed`（完整 key）双表过滤
6. **移动重试**：Move-OneResume 对被占用文件重试 3 次（每 500ms）

### 产物与退出码

- `<DownloadDir>\_summary.json`：机器可读结果 `{status: DONE|INCOMPLETE, ok, fail, skip, target, fileCount, format, finishedAt}` —— **Agent 判断是否补跑只看这个文件，不要解析日志**
- `<DownloadDir>\_run_log.txt`：结构化日志
- 退出码：**0**=达标完成 / **2**=未达标（候选池耗尽）/ **1**=参数或岗位致命 / **3**=环境致命

## 快速运行（Agent 执行准则）

1. 更新 `scripts/config.json`（**优先 config.json 传参，避免命令行 URL 的 `&tab=` 被 cmd 误解析**）
2. 环境自检已内建于 run.ps1（[0] 阶段，失败 exit 3）。**但 15 分钟级批量任务必须用计划任务承载**（见下节），Agent 会话内后台任务约 2 分钟被宿主强杀
3. 调用时前置 `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`；**不要嵌套 `powershell -File`**，用 PS 的 `&` 调用运算符
4. 中断重启：`DownloadCount` 改为 `目标总数 - 目录已有份数`（断点续传自动跳过，`[RESUME]` 日志确认）
5. 结束后读 `_summary.json` 判断结果；日志中 `[TOP-UP]`/`[SKIP]`/`RECOVER` 是**预期自愈行为**，不要手动干预

### ⭐ 长任务调度经验（CRITICAL，2026-09-15 实测）

- **会话内后台任务（含非沙箱 run_in_background）约 2 分钟被宿主强杀**，daemon 作为子进程连带死亡 → 批量下载在会话内根本跑不完。这就是"必须计划任务"的根因
- **唯一可靠方式：`Register-ScheduledTask` 计划任务**（独立进程树，宿主杀不到）：

```powershell
$act = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File <wrapper.ps1>"
Register-ScheduledTask -TaskName "WbResumeDL" -Action $act -Trigger (New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2)) -Force
Start-ScheduledTask -TaskName "WbResumeDL"
```

- wrapper.ps1 内容 = E 节标准启动序列（daemon start → list_tabs 轮询等扩展重连 → 设 NO_PROXY → 运行 run.ps1）。**wrapper 必须纯 ASCII 路径**（中文放 config.json/数据文件，trash/命令行对中文路径有 GBK mojibake 坑）
- `schtasks /create` 在本机报 exit -1，用 `Register-ScheduledTask` 替代

## 配置项说明（`scripts/config.json`）

| 键 | 说明 |
|---|---|
| `Url` | 智联推荐页完整 URL（含 jobNumber，去掉 `#` 片段） |
| `JobName` | 岗位名称，需与页面标签完全一致 |
| `DownloadDir` | 简历最终保存目录（绝对路径） |
| `DownloadCount` | 目标下载数量 |
| `FileFormat` | `"word"`（默认，生成 .docx）或 `"pdf"` |

> `DownloadWaitMs` 已废弃（#55 后不再阻塞等待），保留键仅为兼容旧 config。

## 下载主循环语义（已代码化，此处只述流程不述实现）

主循环 = **卡片登记 + 顺序推进（#53/#54）**：

1. **[A] 视觉识别**：DOM 提取视口所有卡片 → key = `姓名_年龄_工作经历摘要`（摘要=容器文本剔除易变时间词后前 30 字符）
2. **[B] 双表过滤**：`$processed`（本次运行完整 key）+ `$prefilled`（断点续传 姓名_年龄）命中任一即跳过
3. **[C] 点击与处理**：前台化 → 三重校验打标记 → DOM click → 确认面板 → 存至本地（探测到立即 CDP 点击，#46）→ word 切换 → 保存 → **release+1000ms 关面板（#55）** → 轮询检测落盘 → 移动去重 → 登记完毕
4. **[D] 顺序推进**：视口消化完才下滚一屏（800ms，不回顶）；到底+连续 3 轮无新卡片 → 回顶重扫（推荐列表动态重排产生新卡片）；连续 3 空轮才停

已知限制：同名同龄且摘要相同（理论上同一人）视为同一卡片；同名同龄不同经历可正确区分。

## 时序参数（多轮踩坑调优值，勿随意压缩——每项都有原因）

| 参数 | 值 | 原因 |
|---|---|---|
| CloseWaitMs | 700ms | 遮罩点击释放后等模态关闭 |
| ClickWaitMs | 800ms | 点中卡片后等面板开始渲染 |
| 存至本地按钮等待 | ≤30s（Wait-Until） | #45：按钮随简历正文加载延迟 10~15s 才挂载 |
| DialogCheckWaitMs | 2000ms | 保存对话框弹出延迟；按钮闪烁期 12 次重试窗口（#46） |
| ScrollWaitMs | 800ms | #53 用户指定（原 1200）；虚拟列表滞后由停滞轮数兜底 |
| SaveCloseWaitMs | 1000ms | #55 用户指定：release 保存按钮后关面板前 |
| 落盘轮询窗口 | word 25s / pdf 15s | word 文件生成更慢；每秒一查，命中基准 = 点击时刻-5s |

## 去重策略（关键规则）

### 重复判断条件（三重验证）

文件命名格式为 `{姓名}_{年龄}岁_智联简历_{5位随机数}.docx`（或 `.pdf`）。

**以下三个条件必须同时满足才认定为重复：**

1. **姓名相同** — 文件名中第一个 `_` 之前的部分（如 `张先生`）
2. **年龄相同** — 文件名中 `_` 之后、`岁` 之前的数字（如 `35`）
3. **文件大小相同** — 字节数完全一致

### 重复处理策略（不重置流程）

- **不删除所有已下载数据重新执行**
- 发现重复 → 仅删除新下载的重复文件
- 当前候选人不计入成功数（`ok` 不递增）
- 将该候选人标记为"已尝试"，继续尝试列表中的下一位
- 持续尝试直到 `ok` 达到 `downloadCount` 或列表耗尽

```powershell
$moved = Move-OneResume
if ($moved) {
    $ok++  # 成功，计数递增
} else {
    # 重复 → 计数器不变，继续下一个
    $triedCandidates[$name] = $true
}
```

## 关键注意事项

### 编码要求

- **所有 .ps1 脚本文件必须使用 UTF-8 BOM 编码**。在中文 Windows 上，PowerShell 5.x 默认以 GBK 解析无 BOM 文件，导致中文乱码和语法错误（典型症状：`Unexpected token '}'` 出现在正常花括号处，实为中文注释被 GBK 误解析破坏了语法结构）
- **AI Agent 执行前检查 BOM**：用 `[System.IO.File]::ReadAllBytes()` 读取前3字节，若不为 `0xEF 0xBB 0xBF` 则自动用 `[System.Text.UTF8Encoding]::new($true)` 重新写入
- 写给 curl 请求的临时 JSON 文件使用 **无 BOM UTF-8**
- 推荐使用 `ConvertTo-Json` 构建 WebBridge 请求 JSON，而非手动拼接字符串
- 函数参数名避免使用 `$Args`（PowerShell 保留的自动变量）

### PowerShell 单引号字符串内嵌 JS 的双引号陷阱（CRITICAL）

> **这是最隐蔽的 bug 来源**。PowerShell 单引号字符串 `'...'` 中，`""` 是**两个双引号字面量**（不是转义为单个双引号），`''` 才是转义为单个单引号。

在 PowerShell 单引号字符串中嵌入 JS 时：
- ✅ `document.querySelector(".x")` — 正确，单个 `"` 即可
- ❌ `document.querySelectorAll("".x"")` — 错误，JS 端收到 `""`（空串）+ `x` + `""`（空串）→ `SyntaxError`
- ✅ `document.querySelectorAll('.x')` — 正确，JS 单引号
- ✅ `document.querySelectorAll(''.x'')` — 正确，PS 中 `''` 转义为单个 `'`

**`Invoke-Eval` 不检查响应错误**，JS 语法错误会被静默掩盖为"找不到元素"，日志只显示 `[FAIL]`/`[WARN]`，看不到 JS 错误。排查时需单独用 `curl` 调用 `evaluate` 确认 JS 是否报 `SyntaxError`。

### 坐标探测统一用 getBoundingClientRect()

- **CDP `Input.dispatchMouseEvent` 使用 CSS 像素**，与 `getBoundingClientRect()` 一致
- **不要用 `DOM.getBoxModel`**：返回设备像素（受 `devicePixelRatio` 影响），且 `DOM.requestNode` 对未注册节点返回 `nodeId: 0` 导致失败
- `Get-SaveButtonCoords` 函数已改为基于 `getBoundingClientRect()` 的 JS 探测

### 姓名提取

- 用正则 `/^\S+/` 提取第一个非空白序列（比 `indexOf(' ')` + `substring` 更可靠）
- `textContent` 包含换行和时间信息（如 `\n  张先生\n  \n  12小时前看过`）
- 查找时使用 `includes('张先生')` 而非 `startsWith('张先生')`

### 虚拟滚动

- DOM 中始终仅 ~20 个 `.talent-basic-info__name` 节点
- 关闭模态后可能需要重新滚动让目标候选人进入视口
- **循环中使用的姓名与实际被点击的候选人可能不一致**（DOM 延迟替换）
- 使用 CDP `mouseWheel` 而非 `window.scrollBy` 触发加载

### CDP 事件时序

| 步骤 | 操作 | 间隔 |
|------|------|------|
| mouseMoved | 移动到按钮位置 | — |
| 等待 | 页面响应 | ≥ 800ms |
| mousePressed | 按下左键 | — |
| 等待 | 保持按压 | 250-400ms |
| mouseReleased | 释放左键 | — |
| 等待 | 检查对话框 | ≥ 1500ms |

### 文件命名规则

`{姓名}_{年龄}岁_智联简历_{5位随机数}.docx`（或 `.pdf`）

示例：`刘先生_24岁_智联简历_18535.docx`

### ⚠️ 下载后文件转移（CRITICAL — 已修复）

**"简历下载成功但没转移到目标目录"的核心根因与修复：**

1. **中文通配符匹配失败**：脚本查找 Downloads 中的智联文件时，原用 `Get-ChildItem -Filter "*智联*"`。中文通配符在 PowerShell 中依赖代码页，且命令行内嵌中文会被 CLI 传输破坏，导致**匹配返回空** → 无法识别已下载文件 → 文件留在 Downloads。
   - **修复**：统一改用 `Get-ChildItem -Path "$dir\*" -Include "*智联简历*.$FileExt"`（见 `run.ps1` 的 `Move-OneResume`、阶段 3.6 检测、`Test-NameExistsInDir`、`Test-DuplicateFile`）。
   - 优先用 `-File scripts\run.ps1` 运行，避免命令行内嵌中文参数。

2. **返回值语义混淆**：原 `Move-OneResume` 把"找不到文件"和"重复"都返回 `$null`，调用方静默当"重复"，掩盖真实故障。
   - **修复**：现在返回 `'ok'` / `'dup'` / `$null` 三种语义，`$null` 明确输出 `[WARN]` 并保留文件供人工处理。

3. **文件被占用无法移动**：个别文件被程序（如 Word/WPS 或遗留下载进程）占用，`Move-Item` 失败。此时脚本用 `copy` 复制到目标目录并提示用户手动删除原文件。

> 验证方法：下载后若发现文件没转移，先 `cmd /c dir /b "%USERPROFILE%\Downloads\*智联简历*"` 确认文件确实存在，再排查是匹配问题还是被占用问题。

## 已解决问题记录

1. ✅ 虚拟滚动 DOM 节点限制 → 滚动累积去重收集
2. ✅ JS `.click()` 关闭模态不可靠 → CDP 遮罩点击
3. ✅ `startsWith()` 姓名匹配失败 → 改用 `trim()` + `includes()`
4. ✅ PowerShell 中文乱码 → 强制 UTF-8 BOM + `\uXXXX` 转义
5. ✅ 约 50% CDP 首次点击失败 → 自动重试机制
6. ✅ 重复下载 → JS Set 去重跟踪 + 三重文件验证去重
7. ✅ 姓名提取包含换行 → 改用 `/^\S+/` 正则
8. ✅ `$Args` 参数名冲突 → 改用 `$Payload`
9. ✅ 手动拼接 JSON 出错 → 改用 `ConvertTo-Json`
10. ✅ 候选人列表加载延迟 → 轮询等待 `.talent-basic-info__name` 出现
11. ✅ URL 中 `#` 片段解析异常 → 去掉 `#` 片段
12. ✅ 虚拟滚动导致姓名与实际文件不一致 → 三重验证去重 + while 循环补下载
13. ✅ PS 单引号字符串内嵌 JS 用 `""` 包裹 → JS SyntaxError 静默失败 → 改为单个 `"`（详见 tech_details #16）
14. ✅ `Invoke-Eval` 不检查响应错误 → JS 语法错误被掩盖为"找不到元素" → 建议增加 `ok:false` 检查（详见 tech_details #17）
15. ✅ `DOM.getBoxModel` 获取坐标失败（nodeId=0）→ 改用 `getBoundingClientRect()`（详见 tech_details #18）
16. ✅ `DOM.getBoxModel` 返回设备像素与 CDP CSS 像素不一致 → 统一用 `getBoundingClientRect()`（详见 tech_details #19）
17. ✅ "存至本地"按钮等待不充分 → **两段式等待**：先等 `.resume-detail-wrap`（20×500ms），再等 `.resume-button.position-r` 可见（**40×750ms ≈ 30s**）；且探测到即刻点击（按钮会闪烁）。详见问题 #45/#46 与 tech_details #20
18. ✅ 仅支持 PDF 格式下载 → 增加 `-FileFormat "word"` 参数，在保存对话框中切换"word"图标后再点"保存"，自动产出 `.docx` 文件（详见 tech_details #21）
19. ✅ 下载成功但文件没转移到目标目录 → 中文通配符 `-Filter "*智联*"` 匹配失败（CLI/代码页），统一改用 `-Include + -Path "$dir\*"`；`Move-OneResume` 返回值区分 `'ok'/'dup'/$null`，不再把"找不到文件"误判为"重复"（详见 tech_details #22/#23）
20. ✅ FileFormat 配置与网页实际下载格式不一致导致文件无法转移 → 文件搜索层改为同时匹配 `.pdf` 和 `.docx`，不依赖 `$Config.FileExt`（详见 tech_details #24）
21. ✅ `run.ps1` 无 UTF-8 BOM 导致中文 Windows PowerShell 5.x 解析错误（`Unexpected token '}'`）→ **所有 .ps1 文件必须确保 UTF-8 BOM**，AI Agent 执行前应检查并自动添加 BOM 头（详见 tech_details #25）
22. ✅ `navigate` 新标签页创建后浏览器 UI 未自动聚焦 → 使用 CDP `Page.bringToFront` 强制切换到当前 CDP 连接的标签页。排查过程发现 `find_tab` URL 模糊匹配会错误切回旧标签页，`Target.activateTarget` 被 Chrome 安全策略阻止（详见 tech_details #27）
23. ✅ PowerShell exit code 1 但脚本实际成功（CLIXML 进度流污染）→ Agent 应解析 stdout 中的 `Success : N` 和 `Target files : N` 判定成败，不依赖 exit code（详见 tech_details A1）
24. ✅ Python 路径含中文空格时 PowerShell 调用失败（`CommandNotFoundException`）→ 统一使用 `& "path"` 调用运算符包裹路径（详见 tech_details A2）
25. ✅ config.json 的 JobName 与用户指令不一致导致岗位验证失败 → Agent 每次执行前必须更新 config.json 的所有字段（详见 tech_details A3）
26. ✅ 页面岗位标签带状态后缀（如"·协作未上线"）导致严格相等匹配失败 → 岗位验证已使用 `-like "*$expectedClean*"` 包含匹配（详见 tech_details A4）
27. ✅ snapshot 精确匹配不允许岗位名带后缀 + 阶段 2.1 策略 2 的 `.active` 选择器误匹配导航栏"推荐" → snapshot 改为包含匹配 + JS 模糊匹配双路径；岗位验证策略 1 改用 `.job-pane__item--active` 精确选择器（详见 tech_details #28/#29）
28. ✅ `fix_config.py` 硬编码用户/项目路径 → 改为基于脚本自身目录动态定位；`config.json` 清空运行时数据为占位符（详见"跨环境可移植性"章节）
29. ✅ `ConvertTo-JsSafeName` 函数在第551行定义但第188行调用（PowerShell 顺序执行，函数必须先定义后使用）→ 已将函数定义移至第150行之前（详见 tech_details #30）
30. ✅ 命令行通过 `cmd /c` 传递 URL 参数时 `&tab=recommend` 被 cmd 误解析为命令分隔符（`'tab' 不是内部或外部命令`）→ Agent 应优先使用 `config.json` 传参，避免命令行内嵌 URL 特殊字符（详见 tech_details #31）
31. ✅ 全局 skill 目录下 `_user_meta.json` 文件（含 `agent_created: true`）导致 CodeBuddy 不加载该 skill → 删除 `_user_meta.json` 并重载窗口后恢复正常（详见 tech_details #32）
32. ✅ 中文 JobName 在 Agent 后台任务环境中乱码（"销售部经理"→"销售部经í"）→ `run.ps1` 内置 `[Console]::OutputEncoding = UTF8` + Agent 调用时前置设置编码；禁止嵌套 `powershell -File`（详见 2026-08-04 对话记录）
33. ✅ **姓名收集阶段误判"推荐池已到底"（只收集到 31 个，实际 142 个）** → 根因：CDP `mouseWheel` 固定坐标滚动对虚拟列表不可靠 + 停滞阈值 `$stag -ge 3` 太激进（第 11 次滚动即误判）。**修复**：改用 JS 直接滚容器 `scrollTop += 900`；停滞阈值提高到 12 轮；加 `$maxScrolls = 80` 硬上限。用户质疑："明明有非常多简历，为什么你说才31份"
34. ✅ **下载阶段全程空转，从第 37 位起全部 `[SKIP] Not found in viewport`** → 根因：收集阶段把列表滚到底部后**下载阶段未重置回顶部**，且查找循环只往下滚，已到底无法再下滚。**修复**：① 进入下载前 `scrollTop = 0` 归零；② 每个候选人查找循环第 0 轮先归零再逐屏下滑；③ 重试上限 15 → 25 次。修复后 0 失败 0 跳过。用户质疑："你一直在重复循环翻候选人列表，但没点进去任何一个候选人详情"
35. ✅ **未满足目标份数就提前停止**（如目标 100 份，遍历完 31 个名字就结束）→ 根因：单层循环 `while ($ok -lt $target -and $idx -lt $names.Count)`，`$names` 遍历完即退出，远早于目标数。**修复**：改为「外层轮次 + 内层遍历 + 补收名字（TOP-UP）」双循环。只有 `$ok >= $targetCount`，或**连续 3 轮无任何新下载**（候选池确实耗尽）才停。补收阶段停滞阈值放宽到 20 轮、maxScrolls 翻倍，新名字追加到 `$names` 后 `$idx=0` 重新遍历（`$triedCandidates` 跨轮去重）。汇总区区分 `[DONE] Target reached` / `[INCOMPLETE] pool exhausted`。用户要求："修复未满足要求的简历份数，就停止"
36. ✅ **查找候选人时滚动累加导致卡死空转** → 根因：查找循环用 `c.scrollTop = c.scrollTop + 700`（**累加式**），只在第 0 轮归零。一旦滚到底部，后续所有重试都停在底部原地打转，25 次重试全部 `[SKIP]`。**修复**：改用**相对顶部定位** `c.scrollTop = $offset`（其中 `$offset = $retry * 700`），每轮都从 0 重新滚到目标偏移，保证能覆盖整份列表。用户要求："修复获取人员信息时卡死空转情况"
37. ✅ **任务完成后浏览器自动化仍在跑** → 根因：脚本汇总后直接结束，未清理为任务创建的标签页、CDP 连接与 daemon 常驻进程，用户看到浏览器"还在自动下载"。**修复**：新增 `Stop-BrowserAutomation` 函数（幂等，`$Script:AutomationStopped` 防重入），依次执行 `close_tab` → `cdp_disable` → `daemon stop` → 清理临时文件；并在汇总前调用。
    ⚠️ **补充纠正（2026-09-15 实测）**：**不要用 `trap` 做兜底清理**。实测 `trap` 会捕获作用域内任何 terminating error，在长流程脚本中极易被无关的瞬时错误触发，从而**误杀正常执行中的任务**（实测下载第 1 条时就被 `trap` 中断退出）。清理只放在正常流程末尾的 `Stop-BrowserAutomation` 中，并通过幂等标记 + 多处显式调用来覆盖。用户要求："完成任务后，停止浏览器自动下载任务运行"
38. ✅ **岗位验证误报 mismatch**（`Active job detected: [no]` → `FATAL: Job mismatch` → exit 1）→ 根因：正则 `'"value":"([^"]*)"'` 会贪婪匹配响应里任意 `"value":"..."` 字段，捕获到无关的 `"no"`。**修复**：加白名单过滤（排除 `no/ok/yes/none/null/true/false/0/1`）+ 必须含中文字符（`[\u4e00-\u9fa5]`）+ 长度 ≥ 2；探测失败时降级为警告继续，不阻断流程。
39. ✅ **名字收集返回 0**（`Found 0 unique names / [FAIL] No names — abort.`）→ 根因：上一轮 cleanup 关闭了标签页，重新 `navigate` 后页面尚未渲染完成即进入循环，首轮 evaluate 返回 `ok:false`（"session has no tab"），循环静默空转到 maxScrolls。**修复**：新增"等待列表容器渲染"预检（最多 20 轮）；循环内新增 `$evalErrors` 计数，`ok:false` 时计数，连续 3 次则重新 `navigate` + `Page.bringToFront` 恢复。
40. ✅ **`wb-download.json` 文件锁导致首次请求之后所有请求全部失败**（`IOException: 文件正由另一进程使用`）→ 根因：`Send-Web` 固定复用同一个临时文件名，`curl.exe` 可能仍持有该文件句柄（异步/未完全退出），下一次 `WriteAllText` 立刻抛异常。表现极具迷惑性：`SaveLocal 探测成功但点击无响应`、`No save dialog`、卡死。**修复**：每次请求使用 **GUID 独立临时文件**（`wb-req-{guid}.json`），写后立即交给 curl，用毕重试删除（最多 3 次，失败留给系统清理）。这是本轮最关键的基础设施修复。
41. ⚠️ **[已被 #44 推翻 — 误判]** **"智联详情页在「新标签页」打开，而 session 一直绑在「列表标签页」"**
    > **结论订正（2026-09-15 最终定位）**：此项判断**错误**。智联的详情面板是**当前标签页内的 `.resume-detail-wrap`**，并不在新标签页打开。当时之所以在"列表页"探测不到详情元素，真正原因是 **#44：标签页被遮挡，Chrome 不投递输入事件**，导致点击根本没生效。
    > 下面保留原始排查记录作为**方法论参考**（`list_tabs` 仍是有用的排查工具），但 `Switch-ToDetailTab` 方案**已废弃，不要在新实现中沿用**。
    > 原始记录：现象：`SaveLocal btn at (1079, 142)`（其实是硬编码兜底值）→ `No save dialog after 4 attempts` → 空转卡死；用户描述为"一直在翻候选人列表，但没点进去任何一份详情"。
    **根因定位过程**：
    - `list_tabs` 发现**两个标签页**：列表页 `...&tab=recommend#sortType=recommend` 与详情页 `...&tab=recommend&resumeNumber=<encoded>#sortType=...`
    - 在详情页里探测：`.resume-button.position-r` 存在且坐标 **(1079, 142)**、`button[保存]` 存在且坐标 **(996, 561)** —— 与 `Config.SaveLocalX/Y`、`SaveConfirmX/Y` **完全一致**（说明配置坐标本身没错，是事件打在了错误的 tab 上）
    - 在列表页里探测：这些元素**根本不存在**；且列表持续重渲染，每次探测看到的姓名集合都不同
    - 结论：CDP 事件与 `evaluate` 全部作用于 active tab（列表页）→ 点击落空 → 探测永远失败
    **（已废弃的）修复**：新增 `Switch-ToDetailTab` 函数 —— 点击候选人后，用 `list_tabs` 枚举标签页，正则提取 URL 含 `resumeNumber=` 的详情页，再用 `navigate`（`newTab:false`）把 session 的 active tab 切到该 URL，并轮询确认 `.resume-button.position-r` 已可见。调用点：**3.2.5 步骤**（点击候选人之后、探详情面板之前），失败时重试一次，仍失败才跳过该候选人。
    **（已废弃的）配套修复 3.9 步骤**：每条下载完成后（无论成败）检测 `Test-OnDetailTab`，若停在详情页则显式 `navigate` 回列表 URL 并等列表容器渲染，保证下一轮"点击 → 接管详情页"链路可用。
    **当时实测**：切到详情页后完整跑通 —— 点击"存至本地" → 弹窗出现（`bodyCls: km-modal__wrapper--locked`、`saveBtns: 1`）→ 点 "word" 选项（876, 384）→ 点"保存"（996, 561）→ 成功产出 `260915_test\何先生_24岁_智联简历_31791.docx`（100694 字节）。
    **方法教训（仍然有效）**：当按钮坐标"看起来对但点击无效"时，`list_tabs` 是排查"事件是否作用在了错误标签页"的第一手工具。用户要求："修复获取人员信息时卡死空转情况"
42. ⚠️ **[随 #41 一并废弃 — 建立在错误前提上]** **详情页取"第一个"导致 navigate 到上一位候选人**
    > 该问题只在 `Switch-ToDetailTab`（#41 的误判方案）存在时才成立，方案废弃后**不再适用**。
    > **但其中的两条经验仍然有效**：① `list_tabs` 返回的数组**顺序不可靠**，取"第一个"是危险写法；② **tabId 单调递增**，需要"最新"时按 tabId 降序取最大者。
    > 原始记录：根因：`list_tabs` 会返回**多个** `resumeNumber` 标签页（历史遗留的旧详情页排在数组最前），按首个匹配就会 navigate 到**旧简历**（表现为"能点开保存框但下载的是别人"）。`Switch-ToDetailTab` 改为按 **tabId 降序**取最大者；正则（已对真实输出验证）：`"tabId"\s*:\s*(\d+)\s*,\s*"url"\s*:\s*"([^"]*resumeNumber=[^"]*)"`
43. ✅ **★ `close_tab` 关掉 session 的 active 标签页 → session 变成"无标签页"状态** → 现象：`list_tabs` 返回 `tabs: []`，此后所有 `evaluate` 返回 `ok:false`，主循环连续 `[SKIP] Not found after full-list sweep`，且**浏览器里连页面都没了**。
    **根因**：`Clear-StaleDetailTabs` 按 URL 正则匹配所有 `resumeNumber` 标签页并全部 `close_tab`，**把 session 当前 active 的那个也关了**。
    > 注：`Clear-StaleDetailTabs` 本身随 #41/#42 的方案废弃；但**下面这条铁律必须保留**，它适用于任何需要在多标签页环境中清理标签的场景。
    **修复**：
    - 逐个 tab 解析 `tabId`/`url`/`active` **三个字段**，**只关 `active=false`** 的详情页（`if ($isActive) { continue }`）；正则：`'"tabId"\s*:\s*(\d+)\s*,\s*"url"\s*:\s*"([^"]*)"[^}]*?"active"\s*:\s*(true|false)'`
    - 清理后**兜底 navigate 回列表 URL** 并等列表容器渲染，确保主循环从列表页开始
    - 需要离开详情页时**用 `navigate` 复用当前标签页**，不要 `close_tab`（导航不会堆积也不会丢 session）
    - 候选人查找循环新增**自愈分支**：识别 `"ok":false`（session 无 tab）→ 重新 `navigate` 回列表 URL + 等渲染 + `continue`，不再静默滚完 25 次重试变 `[SKIP]`
    **核心教训**：**永远不要 `close_tab` session 当前 active 的标签页**。需要离开时用 `navigate` 复用标签页；只有确认 `active:false` 的残留页才可以关。
44. ✅ **★★★ 真正根因：标签页被遮挡 → Chrome 不投递输入事件；必须用 DOM 级 `click` 而非坐标点击**
    **这是本轮最终定位到的根因，#41/#42/#43 的"切标签页"方案均为误判。**
    **决定性证据**：调用 `mouse_click` 时扩展返回的原始报错——
    > `the click did not reach the page — no pointerdown/mousedown fired. The tab is likely backgrounded and occluded (another tab in its window is shown, so its render widget is hidden and real input isn't delivered). Bring the tab to the front of its window, or hand off to the user, then retry.`
    **两个独立失效原因**：
    1. **标签页遮挡**：同窗口存在多个标签页时，session 的标签页处于 background/occluded，Chrome **不会把真实输入事件投递给隐藏的 render widget** → 所有 `Input.dispatchMouseEvent` / `mouse_click` 静默失败（返回 ok:true 但页面无反应）。
    2. **虚拟列表视口外**：候选人节点可能在视口外（实测 y=1853 / 视口高 735），普通 `scrollIntoView` 对该虚拟滚动容器无效。
    **修复**：
    - 新增 `Ensure-TabFocused`：调 CDP `Page.bringToFront` 把标签页提到窗口最前。**每次交互前调用**（候选人查找前、点击存至本地前、点 word 前、点保存前）。
    - 新增 `Invoke-Click`：改用 WebBridge 的 **`click` action**（DOM 级派发，内部自动处理滚动与不可见元素），**不再依赖坐标**。对无稳定 class 的元素（如纯文本 "word" 选项），先 `setAttribute` 打临时标记（`data-wb-word` / `data-wb-save`）再用选择器点击。
    - 候选人详情改为 `Invoke-Click '.talent-basic-info__name'` → **详情面板在同一标签页内直接渲染**（`.resume-detail-wrap` + `.resume-button.position-r` 随即出现），**根本不需要切换标签页**。
    - 坐标点击保留为兜底路径（DOM click 失败时才走）。
    **实测验证**：`Page.bringToFront` + `click .talent-basic-info__name` → `saveLocalBtn: true`、`detailEl: resume-detail-wrap new-shortcut-resume__resume-view-wrap`、`bodyCls: km-modal__wrapper--locked`。
    **核心教训**：**WebBridge 中优先用 `click`（选择器）而不是坐标点击**；坐标点击只在 DOM click 不可用时兜底。多标签页场景下**务必先 `Page.bringToFront`**。
45. ✅ **详情面板的「存至本地」按钮延迟挂载（约 10~15 秒）** → 根因：面板主体 `.resume-detail-wrap`（约 1107×715）先渲染，但 `.resume-button.position-r`（含"存至本地"）要**等简历正文加载完**才挂载。原等待上限 15×500ms=7.5s 不够 → 探测失败 → 退回硬编码坐标 → 点了个不存在的按钮 → `No save dialog` → 空转。**修复**：① 先等 `.resume-detail-wrap` 出现（20 轮 × 500ms）；② 再等 `.resume-button.position-r` 可见（上限提到 **40 轮 × 750ms ≈ 30 秒**）。
46. ✅ **★★★ 最终跑通的关键序列：按钮会"闪烁"，必须"探测到即刻点击"**
    **现象**：详情面板打开后 `.resume-button.position-r` **反复挂载/卸载**（Vue 渲染过程）。探测到"存在"不代表点击那一刻还在。
    **验证过的正确序列**（手工实测成功打开保存对话框并完成下载）：
    1. CDP `Page.bringToFront` —— 不置前则输入被 Chrome 丢弃
    2. DOM `click`（`.talent-basic-info__name`）打开详情面板 → **面板在当前标签页内渲染**
    3. **探测到 `.resume-button.position-r` 可见的那一瞬间立即发出 CDP 鼠标序列**（`mouseMoved` → `mousePressed` → 80ms → `mouseReleased`），**中间不加额外等待**，否则按钮已被卸载
    4. 用 `getBoundingClientRect()` 的**实时坐标**，不要用配置兜底值
    5. 按钮暂未挂载时：判断面板是否还在；面板已自动收起则重新 DOM click 打开，再重试
    **重要**：此处**不能**用 DOM `click` 点"存至本地"——tech_details #21 已记录"不能用 JS `.click()`，Vue 组件需要真实鼠标事件"，实测 DOM click 返回 `success` 但不弹对话框。**候选人姓名用 DOM click，面板内按钮用 CDP 坐标点击**，二者不可混用。
    **实测结果**：`Success: 5 / Failed: 0 / Skipped: 1`，`[DONE] Target reached (5/5)`，并正确执行 `[9] Stopping browser automation...`（tab 关闭 + CDP 断开 + daemon 停止，端口 10086 拒绝连接）。产出 5 份 .docx（91568~276810 字节）。
    **重试次数**：3.4 阶段从 4 次提到 **12 次**，配合"面板关闭则重开"逻辑。
47. ✅ **daemon 请求全部失败但日志显示握手正常 → HTTP_PROXY 代理劫持** → 根因：环境变量 `HTTP_PROXY=http://127.0.0.1:12887` 使 `curl.exe` 把发往 `127.0.0.1:10086` 的请求全部交给本地代理，代理连不上目标时返回伪装成 daemon 错误的 `upstream connect failed (os error 10061)`。**修复**：运行前设 `$env:NO_PROXY="127.0.0.1,localhost"`；调试 curl 一律加 `--noproxy "*"`。详见"环境前置检查 A"。
48. ✅ **daemon 启动后约 1 分钟静默消失（端口无监听、日志无崩溃记录）** → 根因：daemon 由 Agent 的一次性命令启动，命令进程退出时子进程被连带清理。**修复**：用后台常驻任务（`start` + `Start-Sleep`）承载 daemon。另注意 run.ps1 结尾 `Stop-BrowserAutomation` 会主动 `daemon stop`（预期行为）。详见"环境前置检查 B"。
49. ✅ **WebBridge 请求 JSON 带 BOM → `invalid character '茂'`** → 根因：PS 5.1 `Out-File -Encoding utf8` 写出带 BOM 文件，daemon 把 BOM 字节当 JSON 数据。**修复**：用 `[System.IO.File]::WriteAllText(..., UTF8Encoding::new($false))`；curl 响应用 `--output` 直写文件避免 Out-String 的 GBK 双重破坏。详见"环境前置检查 C"。
50. ✅ **daemon v1.11.3 + 扩展 v2.0.9 版本错配** → 扩展升级后旧 daemon 无法正常转发（握手成功但请求失败）。**修复**：执行 `kimi-webbridge upgrade` 对齐版本。详见"环境前置检查 D"。
51. ✅ **下载中段连续 `[SKIP] Not found after full-list sweep` 空转（用户观察："拉到最下方然后重置，往复循环，并未下载新简历"）** → 根因：收集阶段的姓名快照会过期——智联推荐列表动态重排/刷新，部分旧名字已不在列表中；查找逻辑对每个找不到的名字做 25 次重试 × 900ms 全列表往返扫描（相对顶部定位 0→底），单人浪费约 40 秒，且 SKIP 风暴可能连续持续 7+ 人（实测 idx 43-50 连续 7 个 SKIP 后在 idx 51 自愈）。**修复**：新增 `$consecSkip` 连续 SKIP 计数——找到候选人即清零；**连续 ≥8 个 SKIP 视为名单过期，立即 break 内层循环进入 TOP-UP 重新收集新名单**，不再逐个空扫剩余名单。当前运行中的实例不受影响（PS 脚本启动时已解析进内存），修改对下次运行生效。
52. ✅ **查找候选人 sweep 重试 25 → 10**（配合 #51）：25 次 × 900ms 全列表往返在名单过期时空转成本过高，压缩到 10 次（覆盖顶部 ~7000px）。（注：本条后续被 #53 的"取消 sweep"整体取代）
53. ✅ **★★★ 主循环重构：卡片登记 + 顺序推进，下载后不再回滚列表顶部（2026-09-15 用户指定流程）** → 旧模式（姓名名单 + 每人从顶部 sweep）效率低且名单易过期。**新流程**：① 点开卡片前视觉识别（DOM 提取）卡片关键信息（姓名+年龄）；② 处理完毕（成功/失败/重复）立即登记 `$processed["姓名_年龄"]`；③ 关闭模态后从当前位置直接寻找下一条未登记卡片点击，**不回顶**；④ 视口消化完才下滚一屏（ScrollWaitMs **1200→800ms**，用户指定）；⑤ 到底且连续 3 轮无新卡片 → 回顶重扫，连续 3 轮空轮才停（保留 #33 精神）。配套：断点续传按 姓名+年龄 精确预填；删除 `$triedCandidates`/双循环/补收名单逻辑；阶段 2 全列表预收集简化为视口健康检查。**已知限制**：同名同龄不同人视为同一卡片（与文件三重去重口径一致）。文件改动：`scripts/run.ps1`（ScrollWaitMs、阶段 2 收集段、阶段 3 初始化与主循环、SKILL.md 工作流章节）。
54. ✅ **卡片标识升级：姓名+年龄 极易重复 → 加入工作经历摘要（2026-09-15 用户指定）** → 用户指出智联推荐池大量脱敏同名卡（"张先生"）同龄极常见，姓名+年龄做 key 会误并不同候选人。**升级**：① 提取 JS 增加 `w` 字段——卡片容器文本剔除姓名/"N岁"/易变活跃时间词（刚刚/N秒钟前/N分钟前/N小时前/N天前/昨天/本周/本月/在线/活跃/看过）后取**前 30 字符**作摘要；剔除时间词是关键：推荐列表"12分钟前看过"类文案随时间变化，混入 key 会导致同一卡片跨重扫被视为新卡片反复点击；② key 升级为 `姓名_年龄_工作经历摘要`；③ `Get-CardMarkJs` 三重校验（姓名+年龄+摘要），匹配端用与提取端**完全相同的归一化链**（否则摘要跨被剔除词拼接时 indexOf 失配）；④ 断点续传因文件名不含工作经历，改用独立 `$prefilled`（姓名_年龄）基础表 + `$processed` 完整表双表过滤。**验证**：Node 模拟 DOM 冒烟测试——两个"张先生/34岁"不同公司卡片 key 正确区分、时间词 0 泄漏、标记 JS 精确命中目标卡片。**测试中抓到并修复 1 个拼接 bug**：`join("").` + 以 `.replace` 开头的归一化片段产生 `..` 双点 JS 语法错误（Invoke-Eval 会静默掩盖，靠 mock 测试暴露）。文件改动：`scripts/run.ps1`（$CardExtractJs、Get-CardMarkJs、$prefilled、[B]/[C] key 逻辑）。
52. ✅ **#51 修复对运行中实例无效 + 重启后重复下载** → 用户再次观察到空转（idx 56-61 连续 6 个 SKIP）——**改 PS 脚本不影响已启动的实例**（脚本在启动时已解析进内存），必须停任务→改→重启才生效。重启又引入新问题：`$triedCandidates` 在内存里，重启后已下载的 45 人会被重新下载（每人 ~40s，45 份=30 分钟浪费）。**修复（双管齐下）**：① sweep 重试 25→10（单人 SKIP 代价 ~40s→~15s）；② **断点续传**——启动时扫描 `DownloadDir` 已有 `*.docx`，取文件名第一段（`_` 前的姓名）预填 `$triedCandidates`，日志输出 `[RESUME] N resume(s) already in target dir`；同名不同人风险由 `Move-OneResume` 的 DUP 三重验证（姓名+年龄+文件大小）兜底。③ 重启前把 config `DownloadCount` 减去目录已有份数（100-45=55），避免总数超 100。**运行中脚本热改无效**是 PS 脚本调度的通用陷阱，见问题 #51 教训的组合。
55. ✅ **保存流程提速：release 保存按钮后 1000ms 即关闭详情面板（2026-09-15 用户指定）** → 原流程 release 后阻塞等待 `DownloadWaitMs`（word 模式强制 10s/人）才检测文件、关面板，单人固定等待成本过高。**新流程**：release → **Wait 1000ms** → 立即 `Close-ModalIfOpen` + `.km-modal__close-btn` DOM click 关闭详情面板 → 轮询检测新文件落盘（word 25s / pdf 15s 窗口，每秒一查，`LastWriteTime > clickTime-5s` 判定命中）。**配套**：面板关闭后无法重开重点保存，原 3 次保存重试 for 循环整体移除（单次点击失败走 [FAIL] 分支按已登记跳过）；文件命中基准改用点击前的 `$clickTime`（替代原 `Now-detectWindow` 滑动窗口，长轮询下更精确）。文件改动：`scripts/run.ps1` 3.6 节。
56. ✅ **★★★ Skill 模块化重构：固定流程沉淀为脚本，文档只留经验（2026-09-15 用户指令）** → run.ps1 曾是 1337 行单体（客户端/页面交互/编排混杂），SKILL.md 用 400+ 行伪代码复述流程（与代码漂移、维护双份）。**新架构**：`lib/wb-core.ps1`（通用层：WebBridge 客户端、Write-Log 结构化日志、Invoke-WithRetry/Wait-Until 稳定性原语、Initialize-WebBridgeEnv 环境自检、Stop-BrowserAutomation、标签页管理）+ `lib/zhaopin-page.ps1`（页面层：选岗/验证、卡片提取与三重标记、模态与保存序列、文件移动去重）+ `run.ps1` 薄编排。**稳定性增强**：① 环境自检内建（list_tabs 真探测 + NO_PROXY + 等扩展重连 200s）；② deadline 制等待替代固定次数循环；③ 主循环 try/catch 单轮异常保护（连续 5 轮才终止）；④ 移动文件重试 3 次；⑤ 产物 `_summary.json`（机器可读 DONE/INCOMPLETE）+ `_run_log.txt`（UTF-8 结构化日志，替代 `*>` 重定向 UTF-16 痛点）；⑥ 明确退出码 0/1/2/3。**文档**：SKILL.md 893→约 460 行，删除流程伪代码与弃用章节，新增脚本架构/稳定性机制/长任务调度经验/时序参数表；删除弃用文件 config.ps1/webbridge-utils.ps1/download-loop.ps1/fix_config.py。**排障新知**：扩展掉线需用 list_tabs 真探测（snapshot "no tab" 会糊弄探测）；会话内后台任务约 2 分钟被宿主强杀→15 分钟级任务必须 Register-ScheduledTask 计划任务承载；safe-delete/trash 对中文路径报 GBK mojibake 错误（删除实际可能成功，需复核）；PS 工具与 Edit 工具均出现过"报成功未落盘"——所有关键写入必须 Grep/Read 复核。

## 绑定资源

- `scripts/run.ps1` — **主编排入口**（配置校验/环境自检/主循环/汇总）
- `scripts/lib/wb-core.ps1` — 通用层：WebBridge 客户端 + 稳定性原语 + 环境自检 + 清理
- `scripts/lib/zhaopin-page.ps1` — 智联页面层：选岗/卡片/保存/文件函数库
- `scripts/config.json` — 运行参数（Agent 执行前必须更新）
- `assets/` — 可复用 JavaScript 模板
- `references/tech_details.md` — 页面结构、选择器、CDP 坐标、编码规范、完整踩坑记录

> 弃用文件（config.ps1 / webbridge-utils.ps1 / download-loop.ps1 / fix_config.py）已于 #56 删除。

## 跨环境可移植性

本 skill 可安装于项目本地目录（`.workbuddy/skills/`）或全局目录（`~\.workbuddy\skills\`），所有路径引用均基于脚本自身目录动态解析，无硬编码绝对路径：

- `run.ps1` 使用 `$MyInvocation.MyCommand.Path` 定位同目录 `config.json`
- 所有 WebBridge 函数内建于 `run.ps1`，不依赖 dot-source 其他 .ps1 模块
- 下载源（`$env:USERPROFILE\Downloads`）和守护进程路径（`~\.kimi-webbridge\bin\`）使用环境变量
- Python 脚本全部通过 `sys.argv` 接收路径参数，不依赖固定位置
