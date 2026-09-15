# 项目长期记忆 — 260724 【人资】简历筛选

## zhaopin-resume-saver skill 关键约定

- **skill 修改纪律**：每次修改必须清晰列出根因分析、规避方案、文件改动明细，并写入项目记忆（用户要求）。
- **Edit 工具本机偶发"报成功但未落盘"**（已发生 3 次）：关键修改后必须 Grep/Read 复核；多处修改宁可拆分+逐一复核。
- **本机 PowerShell 工具 stdout 捕获异常**：命令实际执行但无输出，勿把"无输出"当失败；用文件落盘 + Read 复核（如语法检查结果写入临时文件再读）。
- **运行中的 PS 脚本热改无效**：run.ps1 启动时已解析进内存，修改只对下次运行生效；改完需 TaskStop 旧实例并重启。
- **下载主循环现行架构（2026-09-15 #53 重构，用户指定流程）**：卡片登记 + 顺序推进——点开前 DOM 提取卡片关键信息（key=**姓名_年龄_工作经历摘要**，#54 升级：摘要=容器文本剔除易变时间词后前 30 字符，区分同名同龄不同人）→ 处理完登记 `$processed` → 关模态后从当前位置直接点下一条未登记卡片，**不回顶**；视口消化完才下滚（ScrollWaitMs=800ms，用户指定 1200→800）；到底+连续 3 轮无新卡片→回顶重扫，连续 3 空轮才停。`$triedCandidates`/双循环/补收名单已删除；断点续传走 `$prefilled`（姓名_年龄）+`$processed`（完整 key）双表过滤。提取与标记 JS 必须用相同归一化链（时间词剔除表：刚刚/N秒钟前/N分钟前/N小时前/N天前/昨天/本周/本月/在线/活跃/看过）。
- 断点续传：启动扫描 DownloadDir，按 姓名+年龄 精确预填登记表；中断重启前 DownloadCount=目标-已有份数。
- **保存流程（#55，2026-09-15 用户指定）**：release 保存按钮 → **1000ms** → 立即关详情面板（Close-ModalIfOpen + `.km-modal__close-btn` 兜底）→ 轮询检测落盘（word 25s/pdf 15s 窗口每秒一查，`LastWriteTime > clickTime-5s` 命中）；3 次保存重试循环已移除（面板关闭后无法重点保存）。
- **长任务调度铁律（本会话实测）**：会话内后台任务（含非沙箱 run_in_background）约 2 分钟被宿主强杀，daemon 连带死亡——15 分钟级批量任务必须用 **Register-ScheduledTask 计划任务**（独立进程树）承载；就绪探测必须用 `list_tabs`（snapshot 的 "no tab" 业务错误会糊弄探测），6 次 stop/start 后扩展掉线需轮询等重连。
- 环境铁律：NO_PROXY=127.0.0.1,localhost（HTTP_PROXY 劫持）；daemon 用 run_in_background 常驻承载；WebBridge 请求体无 BOM UTF-8；daemon 与扩展版本用 `kimi-webbridge upgrade` 对齐；run.ps1 结尾 Stop-BrowserAutomation 会 daemon stop（预期）。
- WebBridge 交互铁律：每次交互前 `Page.bringToFront`（遮挡则输入静默丢失）；候选人姓名用 DOM click（打标记+选择器），面板内按钮（存至本地/word/保存）用 CDP 真实鼠标事件；存至本地按钮会闪烁，探测到即刻点击；绝不 `close_tab` active 标签页。

## 简历评估（zhaopin-resume-screening）

- 模版 EvaluationTemplate.xlsx：12 列（无建议行动）；列宽 A/C=36、K/L=50、其余 12；全表垂直居中+自动换行；冻结 A2；B 列色阶 高C6EFCE/中FFEB9C/低FFC7CE。
- `fill_excel_template(json, xlsx)` 是就地覆盖——先复制模版再填；extract_docx.py 字段名是 `filename`；按行索引写（row=idx+2），勿按姓名匹配（同名同龄会串行）。
- 智联推荐池存在大量同名同龄重复卡片（脱敏名"张先生"），文件级去重靠 姓名+年龄+大小 三重验证。

## 进行中任务

- 联想渠道经理-青岛（26091512，岗位 CC136786060J41007428302）：13:48 用户叫停下载时已落盘 87/100；评估表（26091512\联想青岛.xlsx）未开始，待用户指示。
