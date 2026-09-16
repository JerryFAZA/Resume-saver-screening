# 项目长期记忆 — 260724 【人资】简历筛选

## zhaopin-resume-saver skill 关键约定

- **skill 修改纪律**：每次修改必须清晰列出根因分析、规避方案、文件改动明细，并写入项目记忆（用户要求）。
- **Edit 工具"报成功但未落盘"（已 7+ 次）+ 精确机制（#57 实测确认）：同一消息内对同一文件的多个 Edit 只有最后一个生效**（各自基于原始内容写盘，最后者胜）——批量 run.ps1 3 处丢 2 处、SKILL.md 5 处丢 4 处。铁律：关键修改后必须 Grep/Read 复核；同一文件多处修改必须逐消息拆分+逐一复核（跨文件并行安全）。
- **写入 BOM 铁律**：.ps1 含中文必须 UTF-8 **带 BOM**（PS 5.1 GBK 解析会乱码）——Write 工具不保证 BOM，写完后用 `[System.IO.File]::WriteAllText($f, $t, UTF8Encoding($true))` 补 BOM；WebBridge 请求 JSON 恰恰相反必须**无 BOM**。
- **skill 架构（#56 重构）**：saver=run.ps1（薄编排）+ lib/wb-core.ps1（客户端/日志/重试原语/环境自检）+ lib/zhaopin-page.ps1（页面函数）；screening=pipeline.py（prepare/evaluate/verify）+ extract_docx.py + fill_excel.py；SKILL.md 只放经验，改流程=改脚本。产物 `_summary.json` 为机器可读结果，退出码 0/1/2/3。
- **本机 PowerShell 工具 stdout 捕获异常**：命令实际执行但无输出，勿把"无输出"当失败；用文件落盘 + Read 复核（如语法检查结果写入临时文件再读）。
- **运行中的 PS 脚本热改无效**：run.ps1 启动时已解析进内存，修改只对下次运行生效；改完需 TaskStop 旧实例并重启。
- **下载主循环现行架构（2026-09-15 #53 重构，用户指定流程）**：卡片登记 + 顺序推进——点开前 DOM 提取卡片关键信息（key=**姓名_年龄_工作经历摘要**，#54 升级：摘要=容器文本剔除易变时间词后前 30 字符，区分同名同龄不同人）→ 处理完登记 `$processed` → 关模态后从当前位置直接点下一条未登记卡片，**不回顶**；视口消化完才下滚（ScrollWaitMs=800ms，用户指定 1200→800）；到底+连续 3 轮无新卡片→回顶重扫，连续 3 空轮才停。`$triedCandidates`/双循环/补收名单已删除；断点续传走 `$prefilled`（姓名_年龄）+`$processed`（完整 key）双表过滤。提取与标记 JS 必须用相同归一化链（时间词剔除表：刚刚/N秒钟前/N分钟前/N小时前/N天前/昨天/本周/本月/在线/活跃/看过）。
- 断点续传：启动扫描 DownloadDir，按 姓名+年龄 精确预填登记表；config 写 **DownloadTarget**（总目标恒定）后 run.ps1 自动折算 本次目标=总目标-已有份数（#62），重启免手工改 DownloadCount。
- **#61 选岗验证闭环（2026-09-16）**：Select-JobTab 重构——job-pane 精确匹配优先（includes 兜底）+ data-wb-job 打标，**三级点击升级（JS click→DOM click→CDP 真实鼠标），每级后 Test-JobActive 复核激活岗位名确实切换**；成功后读 location.href jobNumber 写 _summary.json 溯源。config URL 带错 jobNumber 也会被选岗纠正。
- **#62 停滞自动重启（2026-09-16）**：run.ps1 看门狗（ok+fail+skip+dup 连续 300s 无变化→STALL→**exit 4**；滚动硬上限/连续5轮异常同归 stall）；wrapper 重写为最多 8 轮重启循环（exit 0/1/2 停、3/4 重建 daemon 续跑，轮次写 %TEMP%\wb_wrapper_status.txt）。退出码现为 0/1/2/3/**4**。wrapper 模板存 skill scripts/wb_run_wrapper.template.ps1。
- **保存流程（#55/#57，2026-09-15 用户指定）**：release 保存按钮 → **3200ms** → `Close-ModalVerified`（关闭一轮 → 探测 `.km-modal--open`，未关闭每 1000ms 重试，上限 10 次超限 WARN 继续；3.6 保存后与 3.8/3.9 双保险段共用）→ 轮询检测落盘（word 25s/pdf 15s 窗口每秒一查，`LastWriteTime > clickTime-5s` 命中）；3 次保存重试循环已移除（面板关闭后无法重点保存）。
- **长任务调度铁律（本会话实测）**：会话内后台任务（含非沙箱 run_in_background）约 2 分钟被宿主强杀，daemon 连带死亡——15 分钟级批量任务必须用 **Register-ScheduledTask 计划任务**（独立进程树）承载；就绪探测必须用 `list_tabs`（snapshot 的 "no tab" 业务错误会糊弄探测），6 次 stop/start 后扩展掉线需轮询等重连。
- 环境铁律：NO_PROXY=127.0.0.1,localhost（HTTP_PROXY 劫持）；daemon 用 run_in_background 常驻承载；WebBridge 请求体无 BOM UTF-8；daemon 与扩展版本用 `kimi-webbridge upgrade` 对齐；run.ps1 结尾 Stop-BrowserAutomation 会 daemon stop（预期）。
- WebBridge 交互铁律：每次交互前 `Page.bringToFront`（遮挡则输入静默丢失）；候选人姓名用 DOM click（打标记+选择器），面板内按钮（存至本地/word/保存）用 CDP 真实鼠标事件；存至本地按钮会闪烁，探测到即刻点击；绝不 `close_tab` active 标签页。

## 简历评估（zhaopin-resume-screening）

- 模版 EvaluationTemplate.xlsx：12 列（无建议行动）；列宽 A/C=36、K/L=50、其余 12；全表垂直居中+自动换行；冻结 A2；B 列色阶 高C6EFCE/中FFEB9C/低FFC7CE。
- `fill_excel_template(json, xlsx)` 是就地覆盖——先复制模版再填；extract_docx.py 字段名是 `filename`；按行索引写（row=idx+2），勿按姓名匹配（同名同龄会串行）。
- 智联推荐池存在大量同名同龄重复卡片（脱敏名"张先生"），文件级去重靠 姓名+年龄+大小 三重验证。

## 任务状态

- **AI产品经理（岗位 CC136786060J41031060702，薪资 12000-20000）已于 26091517 完成全链路**：output/26091517 ai产品经理 下 200 份 docx（DONE 0 失败 0 跳过，3 次断点续传重启）+ ai产品经理.xlsx（200 提取→去重 15→185 行评估：高13/中104/低68，verify 通过）。高匹配代表：戴洪远33（AI产品总监/Agent矩阵）、杜女士29（AI管理平台+单据识别智能体）、海先生30（海尔AIGC）、李先生26（海信AI客服平台）、张先生29（RAG/LLM+Agent 0-1）、董先生27（LangGraph+Qwen-VL 农资智能助手）。评估规则：年龄>35 封顶中；期望下限>20k 降级；暂不找工作→低；意向地非青岛→低。评估中间产物：_digest.jsonl（digest 提取，_make_digest.py v2）+ evaluations.json（185 条 index 对齐）。
- **saver skill 修复清单**：#58 岗位校验失败→FATAL exit 1；#59 [D] 滚动改 CDP mouseWheel+scrollHeight 增长检测（scrollTop 赋值不触发懒加载）；#60 Send-Web curl `-m 20 --noproxy '*'` 硬超时；#61 选岗三级点击+验证闭环（26091609 晨批错岗位事故沉淀）；#62 停滞看门狗 exit 4 + wrapper 自动重启循环 + DownloadTarget 自动折算。
- **联想渠道经理-青岛（岗位 CC136786060J41007428302）已于 26091516 完成全链路**：output/26091516 下 100 份 docx（DONE，0 失败 0 跳过，16:27-16:59 约 33 分钟）+ 联想青岛.xlsx（100 提取→去重 8 份→92 行评估：高2/中71/低19，verify 通过）。高匹配：张女士30（鼎信通讯经销商全周期管理）、杨先生24（云服务/AI工具企业销售）。
- **历史批次归档位置（"文件夹消失"之谜已解）**：260914/26091512(87份)/26091515(22份) 均在 `output/` 目录下，数据未丢失。
- 计划任务 WbResumeDL（wrapper=%TEMP%\wb_run_wrapper.ps1，路径表=%TEMP%\wb_wrapper_paths.txt）可复用：改 config.json 后直接 Start-ScheduledTask 即可，wrapper 自动 daemon stop→start→list_tabs 探测→跑 run.ps1。
