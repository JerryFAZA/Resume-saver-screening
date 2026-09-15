---
name: zhaopin-resume-screening
description: |
  智联招聘简历批量筛选与 AI 评估。
  从指定文件夹读取智联招聘简历 PDF/DOCX，提取关键信息填入 Excel，再根据岗位 JD 进行 AI 匹配度评估。
  This skill should be used when the user wants to:
  - 批量筛选智联招聘导出的候选人简历 PDF 或 DOCX
  - 自动提取简历信息（姓名、年龄、性别、现居、求职意向、工作经历）到 Excel
  - 根据岗位 JD 对候选人进行 AI 匹配度评估（高/中/低）和标签标注
  必填参数：岗位 JD、简历文件夹路径、输出 Excel 文件路径。
agent_created: true
---

# 智联招聘简历批量筛选与 AI 评估

从智联招聘下载的候选人简历（PDF/DOCX）中自动提取关键信息，填入 Excel 表格，并根据岗位 JD 进行 AI 匹配度评估与标签标注。

## 🚀 AI Agent 执行准则（最高优先级）

### 免确认原则
- **整个筛选+评估流程完全自动化，不需要用户逐步确认**。
- 所有 `execute_command` 调用一律使用 `requires_approval: false`。
- 步骤1-3（提取→填表→AI评估写入）应尽量合并为最少的命令调用。
- 只在最终输出汇总结果，中间步骤静默执行。

### ⚠️ Skill 目录变量（最高优先级）

**本 Skill 可能安装在全局目录或项目本地目录。执行任何脚本前，必须先用本 SKILL.md 文件的实际路径推导 `SKILL_DIR`：**

```
SKILL_DIR = 本 SKILL.md 所在目录的绝对路径
```

- 脚本位于 `$SKILL_DIR/scripts/` 下
- 所有 Python 脚本调用必须使用 `$SKILL_DIR/scripts/xxx.py` 的绝对路径
- **禁止使用相对于项目根目录的路径**（如 `.\workbuddy\skills\...`），因为 skill 不一定在项目目录下

### Python 环境定位策略

**Python 环境仍在用户的项目目录中（`.venv`），不在 skill 目录中。** 执行流程：
1. 定位 Python 解释器：优先使用用户当前工作目录（workspace）下的 `.venv/Scripts/python.exe`
2. 脚本路径：使用 `$SKILL_DIR/scripts/xxx.py`（绝对路径）
3. 数据路径（简历目录、输出 Excel）：使用用户指定的绝对路径（通常在工作目录下）

```powershell
# 示例：从任意 skill 目录执行
<workspace>\.venv\Scripts\python.exe <SKILL_DIR>\scripts\extract_docx.py <resume_dir> <output_json>
```

### 参数推断策略

| 参数 | 推断来源 |
|------|---------|
| `jd` | 用户消息中的岗位职责和任职要求文本 |
| `resume_dir` | 用户指定的简历文件夹，或从 `@folder` 引用获取；如果刚从 zhaopin-resume-saver 下载，使用同一个 downloadDir |
| `output_xlsx` | 默认在 `resume_dir` 下生成 `ai初筛.xlsx`，用户有明确命名时用指定名称 |

### 评估标准的 JD 自适应
匹配度标准不是固定的——**必须根据实际 JD 灵活调整**。例如：
- JD 要求"0-2年销售经验" → 销售经验2年以内为加分项，5年以上可能不匹配
- JD 要求"大专及以上学历" → 学历门槛低，不过滤学历
- JD 要求"熟悉企业微信/腾讯生态" → 有相关经验为加分项
- JD 是初级销售岗 → 管理经验不是必要条件，不应扣分
- JD 是管理岗 → 管理经验是核心条件

**核心原则**：先提取 JD 中的关键硬性条件（经验要求、技能要求、地点要求、**薪资区间**），再逐份对照评估，而非套用固定模板。

### 薪资对比规则（CRITICAL）

当用户提供的岗位 JD 中包含薪资区间时，**必须**将岗位薪资与每位候选人的简历期望工资进行对比：

1. **提取岗位薪资区间**：从 JD 中提取薪资上下限（如 `15k-25k`），统一转换为月薪（千/月或万/月）。
2. **提取候选人期望薪资**：从 `job_intentions` 字段中提取每条求职意向的薪资（如 `【产品经理】1.5万-2万/月-深圳-不限行业`），取其中与目标岗位最相关的那条意向的薪资。
3. **对比规则**：
   - 候选人期望薪资上限 < 岗位薪资下限 → **薪资偏低**（候选人低配，可能经验不足或求职意向不匹配）
   - 候选人期望薪资下限 > 岗位薪资上限 → **薪资不匹配**（候选人期望过高，超出岗位预算）
   - 候选人期望薪资区间与岗位薪资区间有重叠 → 薪资匹配
   - 候选人期望薪资在岗位区间内 → 薪资高度匹配
4. **薪资不匹配对评估的影响**：
   - 薪资严重不匹配（期望远超预算）→ 匹配度至少降一档（如"高"→"中"，"中"→"低"）
   - 薪资略超但候选人其他条件优秀 → 标注"薪资略超"但不强制降档
   - 薪资偏低 → 不影响匹配度（低薪求职者可能是潜力股），但可标注"薪资偏低"

## ⚠️ 强制参数校验（最高优先级）

**本 Skill 启动时必须检查以下三个参数是否全部提供。若任一缺失，立即停止并提示用户补充，直到所有信息满足后才开始任务。**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `jd` | string | **是** | 岗位 JD 完整文本（包含工作职责和任职要求） |
| `resume_dir` | string | **是** | 简历文件所在文件夹的绝对路径（支持 PDF/DOCX 混合） |
| `output_xlsx` | string | **是** | 输出 Excel 文件绝对路径（如已存在则更新 Sheet1） |

### 参数校验提示模板

当用户未提供全部参数时，输出以下提示：

```
请提供以下必要信息后我才能开始简历筛选：

1. **岗位 JD**：请提供完整的岗位描述，包括工作职责和任职要求。
2. **简历文件夹路径**：智联简历文件所在的文件夹绝对路径（支持 PDF 和 DOCX 混合）。
3. **输出 Excel 路径**：希望保存筛选结果的 Excel 文件路径（.xlsx 格式）。

示例：
- 岗位 JD：负责15-20人销售团队管理，制定销售策略与目标，开拓区域市场，维护大客户关系...
- 简历路径：C:\Resumes\销售部经理
- 输出路径：C:\Resumes\AI初筛.xlsx
```

## 工作流程

```
用户提供 JD + 简历文件夹 + 输出 Excel 路径
        │
        ▼
   校验三个参数是否齐全（缺失则提示用户）
        │
        ▼
   步骤1: 检测文件格式，选择提取脚本
        │  PDF → scripts/extract_resumes.py
        │  DOCX → scripts/extract_docx.py
        │  （两者输出格式一致，后续流程通用）
        │
        ▼
   步骤2: 自验证 + 自动修复（仅 PDF 需此步骤，scripts/auto_fix.py）
        │  DOCX 格式文本整洁，通常不需修复
        │
        ▼
   步骤3: 填入 Excel（scripts/fill_excel.py）
        │  写入基本信息，预留匹配度和评价列
        │
        ▼
   步骤4: AI 评估（由 AI 阅读数据逐份评估）
        │  根据 JD 判断匹配度（高/中/低）+ 标签评价
        │  行业判断依据：公司名 + 工作描述内容，双重交叉验证
        │  薪资判断：若 JD 含薪资区间，对比候选人期望薪资，标注薪资匹配/不匹配
        │  ⚠️ 注意识别「同名同年龄重复候选人」，逐条评估并标注"与上一条为同一候选人重复"
        │  ⚠️ 注意英文版简历（字段大量为空），需人工阅卷
        │
        ▼
   步骤5: 写入评估结果到 Excel
        │  首选「按行索引」直接写 B/C 列（row = JSON索引 + 2），避开重名串行 bug
        │  匹配度列（绿=高 C6EFCE / 黄=中 FFEB9C / 红=低 FFC7CE）、简要评价列
        │  写后必须读回校验 max_row / max_col / 样本行 / 各档位统计
        │
        ▼
   输出最终 Excel + 结果摘要
```

## 输出 Excel 格式

| 列 | 内容 | 说明 |
|---|---|---|
| A: 简历名称 | 文件名 | 自动提取 |
| B: 简历匹配度 | 高/中/低 | AI 评估，带颜色标注 |
| C: 简要评价 | 标签总结 | 如"超龄、行业不符、管理经验不足" |
| D: 年龄 | 数字 | 自动提取 |
| E: 性别 | 男/女 | 自动提取 |
| F: 现居 | 城市 | 自动提取 |
| G: 求职意向 | 格式化文本 | 格式：【岗位】薪资-地点-行业，多条换行 |
| H: 工作经历 | 格式化文本 | 格式：【岗位】公司-时间，多条换行 |

## 简要评价标签参考

| 标签 | 含义 | 使用场景 |
|------|------|---------|
| 超龄 | 年龄 > 35 岁 | 固定规则，始终判断 |
| 行业不符 | 工作经历与目标行业不匹配 | JD 有行业偏好时 |
| 行业匹配 | 工作经历与目标行业匹配（正向标签） | 正向标签 |
| 岗位经验不符 | 求职方向/经验与岗位要求不符 | 求职意向偏差时 |
| 无团队管理经验 | 没有带团队的经历 | JD 要求管理经验时 |
| 管理经验不足 | 有管理经验但年限偏短 | JD 要求管理经验时 |
| 销售经验不足 | 销售总年限不足 | JD 要求销售经验时 |
| 销售经验丰富 | 销售经验充分匹配 | 正向标签 |
| 地点不符 | 现居地与岗位所在地不一致 | 地点不匹配时 |
| 行业部分匹配 | 部分经历相关但不完全对口 | 边界情况 |
| AI认知强 | 有AI相关产品/项目经验 | JD 涉及AI产品时 |
| SaaS经验匹配 | 有SaaS/软件销售经验 | JD 偏好SaaS背景时 |
| 纯技术背景 | 只有技术经验无销售经验 | 技术转销售评估时 |
| 薪资不匹配 | 候选人期望薪资超出岗位预算 | 期望薪资下限 > JD 薪资上限时 |
| 薪资略超 | 期望薪资略超预算但其他条件优秀 | 略超预算但候选人优秀时不强制降档 |
| 薪资偏低 | 期望薪资明显低于岗位薪资区间 | 期望薪资上限 < JD 薪资下限时 |
| 薪资匹配 | 期望薪资在岗位区间内或与区间有重叠 | 正向标签，薪资层面无问题 |

## 简历匹配度判断标准

### 固定规则（始终适用，不受 JD 影响）

**年龄硬性约束**：年龄 > 35 岁 → 最高只能评为"中"，不能评为"高"。
- 这是固定规则，无论 JD 是否提及年龄要求都必须执行。
- 超龄候选人即使其他条件全部完美匹配，也不能进入"高"匹配档位。

**薪资硬性约束**：当 JD 提供薪资区间时，候选人期望薪资下限 > JD 薪资上限 → 标注"薪资不匹配"，匹配度至少降一档。
- 薪资严重超出预算（如 JD 15k-25k，候选人期望 30k-40k）→ 匹配度至少降一档
- 薪资略超（如 JD 15k-25k，候选人期望 20k-30k，重叠部分 > 50%）→ 标注"薪资略超"，不强制降档
- 薪资偏低 → 不影响匹配度，标注"薪资偏低"即可

### JD 自适应规则

| 级别 | 条件 |
|------|------|
| **高** | 年龄 ≤ 35 岁 + 行业匹配 + 核心硬性条件全部满足 + 地点匹配 + 薪资匹配 |
| **中** | 部分条件匹配，有培养潜力或可转型（包括年龄 > 35 但其他条件优秀者） |
| **低** | 行业完全不符 / 核心经验严重不足 / 求职方向错误 / 薪资严重不匹配 |

## 脚本说明

### 环境要求

**优先使用用户项目现有的 `.venv` 环境（不在 skill 目录下），不新建环境。** Python 解释器路径示例：

```powershell
<workspace>\.venv\Scripts\python.exe
```

运行前验证依赖是否就绪：
```powershell
<workspace>\.venv\Scripts\python.exe -c "import openpyxl, docx, pypdf; print('OK')"
```

脚本路径使用 `$SKILL_DIR/scripts/xxx.py`（绝对路径）：
```powershell
<workspace>\.venv\Scripts\python.exe <SKILL_DIR>\scripts\extract_docx.py <resume_dir> <output_json>
```

若确需从头搭建环境，在项目目录下用 `uv`：
```bash
cd <workspace>
uv init --no-readme --name resume-screening
uv add openpyxl pandas pypdf python-docx
```

### scripts/extract_resumes.py — PDF 提取模块

核心提取模块。提供 `extract_text_from_pdf()` 和 `parse_resume()` 两个核心函数。

```python
from extract_resumes import extract_text_from_pdf, parse_resume

text = extract_text_from_pdf("path/to/resume.pdf")
data = parse_resume(text)
# data: {name, gender, age, current_residence, job_intentions, work_experiences}
```

工作经历使用 v7 三元组切分法：找时间→反向搜岗位→反向搜公司。

### scripts/extract_docx.py — DOCX 提取模块

处理智联招聘 DOCX 格式简历。DOCX 文件结构清晰：
- **段落**包含：姓名、基本信息、求职意向、教育经历
- **表格**（单列）包含：工作经历（公司+岗位 → 薪资+时间 → 工作描述 → 空行，循环）

**提取策略**：
- `work_experiences`（写入 Excel H 列）：仅含【岗位】公司-时间，**不包含工作描述**
- `work_descriptions`（仅 AI 内部使用）：所有工作描述的合并文本，用于行业匹配度判断
- 行业评估同时基于**公司名语义推断**和**工作描述内容**双重交叉验证

```python
from extract_docx import batch_extract

results = batch_extract("/path/to/docx/folder")
# 每份简历包含: name, age, gender, current_residence,
#   job_intentions, work_experiences(H列), work_descriptions(仅AI评估用)
```

#### DOCX 表格结构详解

智联招聘 DOCX 简历的工作经历在**单列表格**（Table 1）中，每组 4 行：

| 行 | 内容 | 示例 |
|----|------|------|
| 0 | 公司 + 岗位（别名） | `浪潮通信信息系统有限公司   销售  (销售顾问)` |
| 1 | 薪资 + 时间 + 工期 | `1.1万/月   2022.01-至今  (4年6个月)` |
| 2 | 工作描述 | `工作描述：\n1. 负责...` |
| 3 | 空行分隔 | `` |

解析策略：用 `\s{2,}` 拆分公司名和岗位名，从下行提取时间范围，跳过描述和空行。

### scripts/fill_excel.py

将提取的 JSON 数据写入 Excel 并应用格式。同时提供 `write_evaluations()` 用于写入 AI 评估结果。

### scripts/auto_fix.py

PDF 提取结果的自验证+自动修复模块（DOCX 通常不需要）。

---

## PDF 文本提取已知问题（仅 PDF 格式）

智联招聘 PDF 通过 pypdf 提取文本时存在多种格式问题，`extract_resumes.py` 已做全面容错处理。

详细故障模式与修复策略见 PDF 提取脚本注释。关键问题包括：文本跨行拆分、公司名无后缀、教育经历混入、项目经历截断等 26 项容错处理。

> **DOCX 格式优势**：DOCX 简历文本结构清晰，无跨行拆分问题，工作经历在独立表格中，提取精度显著高于 PDF。

## AI Agent 执行经验

### B1. Python 路径含中文空格时需用 `&` 调用运算符

- **现象**：`c:\path\with 中文 spaces\python.exe` 直接执行报 `CommandNotFoundException`。
- **解决**：必须用 `& "完整路径"` 方式调用：
  ```powershell
  & "<workspace>\.venv\Scripts\python.exe" "<SKILL_DIR>\scripts\extract_docx.py" <args>
  ```

### B2. AI 评估写入脚本的构造方式

#### ⭐⭐ 首选：按行索引写入（2026-09-15，重名场景唯一可靠方案）

**`write_evaluations()` 存在重名串行 bug**：它遍历 evaluations 字典并用 `eval_name in name_cell` 做**子串匹配**，智联推荐页常出现**同名同年龄的重复候选人**（如 2 个 `樊先生_38岁`、2 个 `陈先生_28岁`），此时字典 key 相同，第二个会覆盖第一个，且 `break` 会让两行都写同一个评价。

**推荐做法：直接按行索引写入**。数据行 = JSON 索引 + 2（第 1 行是表头，`fill_excel()` 按 `enumerate(data, start=2)` 逐行写入，顺序与 JSON 完全一致）：

```python
# write_eval.py（一次性脚本，执行后删除）
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

XLSX = r"<输出.xlsx>"

# 索引 -> (匹配度, 简要评价)
E = {
    0:  ("低", "英文简历、博士在读、无渠道销售经验"),
    1:  ("中", "渠道分销专员、渠道销售意向明确、行业不符"),
    8:  ("高", "IT软硬件金牌代理商背景、技术销售总监、青岛"),
    # ...
}

FILL = {
    '高': PatternFill('solid', fgColor='C6EFCE'),
    '中': PatternFill('solid', fgColor='FFEB9C'),
    '低': PatternFill('solid', fgColor='FFC7CE'),
}

wb = load_workbook(XLSX)
ws = wb["Sheet1"]
for idx, (lv, tags) in E.items():
    row = idx + 2
    ws.cell(row=row, column=2, value=lv)
    ws.cell(row=row, column=3, value=tags)
    c = ws.cell(row=row, column=2)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.font = Font(name='微软雅黑', bold=True, size=11)
    c.fill = FILL[lv]
    t = ws.cell(row=row, column=3)
    t.font = Font(name='微软雅黑', size=10)
    t.alignment = Alignment(vertical='top', wrap_text=True)
wb.save(XLSX)
```

**校验要点**：写入后必须读回确认 `max_row`（= 简历数 + 1）、`max_col`、几行样本，并统计各档位数量。

#### 备选：`write_evaluations()`（仅适用于无重名的干净数据）

- 位于 `fill_excel.py`，通过 `from fill_excel import write_evaluations` 导入。
- skill 目录不在 `sys.path`，需 `sys.path.insert(0, r"<SKILL_DIR>\scripts")`。
- **key 格式**：`{姓名}_{年龄}岁`（如 `刘先生_30岁`）。文件名格式 `{姓名}_{年龄}岁_智联简历_{随机数}.docx`。
- ⚠️ 只要存在重名同年龄候选人，就**必须改用按行索引写入**。

### B3. AI 评估时需读取 extracted_data.json 的完整内容

- `extract_docx.py` 输出的 JSON 每项键为：`name` / `gender` / `age` / `current_residence` / `job_intentions` / `salary_info` / `work_experiences` / `work_descriptions` / **`filename`**。
- ⚠️ **字段名是 `filename`，不是 `file_name`**，读 JSON 时注意。
- `work_descriptions` 字段（工作描述合并文本）**不写入 Excel H 列**，仅供 AI 做行业匹配度判断时内部使用。
- AI 评估时需要同时参考 `work_experiences`（H 列内容）和 `work_descriptions`（仅 AI 内部使用）来做行业判断。
- 行业判断应基于**公司名语义推断**和**工作描述内容**双重交叉验证。例如：公司名是"联想（北京）有限公司" + 工作描述含"渠道销售、项目开发" → IT/硬件行业。

### B3.1 英文版简历与特殊格式（2026-09-15 新增）

智联推荐池中可能混入**英文版简历**（如 `M先生_24岁`）：段落为 `Desired Job：Foreign Trade Salesman`、`Basic Information`、`Male｜Age 24(2001 Years 12 Months)`，且**公司名与岗位位于文本框/图片中**，表格内只有时间行，导致 `parse_name` 之外的字段几乎全部为空（`age=0`、`residence` 为空）。

处理原则：
- 这类简历**是真实候选人，不能丢弃**；应打开原始 docx 人工判读后给出评估。
- 若 `age=0` 且 `gender`/`current_residence`/`job_intentions` 全空，即为该类型，需人工阅卷。

### B3.2 PowerShell 控制台中文乱码 ≠ 文件编码损坏（2026-09-15 新增）

用 PowerShell 打印中文常显示为 `M鍏堢敓`（GBK/UTF-8 误解），但**文件内容本身是正常 UTF-8**。**不要据此判定数据损坏或重跑提取**。验证方法：把结果写入临时文件后用 Read 工具读取，或直接用 Python `json.load` 打印到文件。

### B4. 临时文件清理

执行完成后必须清理的临时文件：
- `extracted_data.json` — 中间提取数据
- `write_eval.py`（或其他一次性评估脚本）— 评估写入脚本

### B5. 多步骤任务必须无缝衔接（CRITICAL）

当与 `zhaopin-resume-saver` 组合使用时（下载→评估），**下载脚本执行完毕应立即进入评估流程，不得停顿等待用户催促**：
- 下载脚本执行期间就预先准备好评估所需的路径参数
- 下载脚本返回后立刻执行 `extract_docx.py` → `fill_excel.py` → AI 评估写入
- 整个链路应连续执行，中间不输出"是否继续"等交互提示
- 如果下载脚本耗时较长（如2分钟），Agent 不应在等待期间"忘记"后续任务

### B6. cmd 路径含中文空格时的调用方式

- `cmd /c` 传递含中文空格的路径时，内层用双引号包裹：`cmd /c ""path\to\python.exe" "path\to\script.py" args""`
- 外层 `cmd /c` 用双引号包裹整个命令，内层路径用双引号包裹，Windows 下可正确解析
- PowerShell 直接用 `&` 运算符也可行但需注意 `cmd /c` 的嵌套引号规则

### B7. 薪资对比执行方法

- JD 中的薪资区间由 AI 从用户提供的岗位描述中提取，格式如 `15k-25k`、`1.5万-2.5万/月`。
- 候选人期望薪资从 `job_intentions` 字段中提取，格式如 `【AI产品经理】2万-2.5万/月-青岛-不限行业`。
- 对比时统一转换为数字（千元/月）进行比较：
  ```python
  # 示例：JD 15k-25k，候选人 20k-30k → 下限重叠，上限略超
  jd_low, jd_high = 15, 25
  cand_low, cand_high = 20, 30
  if cand_low > jd_high:      # 期望下限 > JD 上限 → 严重不匹配
  elif cand_high < jd_low:    # 期望上限 < JD 下限 → 薪资偏低
  elif cand_low <= jd_low and cand_high >= jd_high:  # 完全包含 → 匹配
  else:                       # 有重叠 → 部分匹配
  ```
- 评估时优先取与目标岗位最相关的那条求职意向的薪资（如岗位是"AI产品经理"，则优先匹配求职意向中"AI产品经理"或"产品经理"的薪资）。
- 如果候选人有多条求职意向但薪资差异大，取与岗位最匹配的那条。

## 注意事项

1. **文件格式**：支持智联招聘（zhaopin.com）导出的简历 PDF 和 DOCX，优先使用 DOCX（提取精度更高）
2. **AI 评估**：匹配度和简要评价由 AI 阅读提取数据后逐份分析，需结合具体 JD 做综合判断。**禁止杜撰数据**——所有评价必须基于实际提取内容。**若 JD 含薪资区间，必须对比候选人期望薪资，并标注薪资匹配/不匹配/略超/偏低等标签**。
3. **Python 环境**：**优先使用用户项目目录下的 `.venv` 环境（不在 skill 目录下），不新建环境或执行 `uv sync`/`uv add` 重建依赖**。脚本路径用 `$SKILL_DIR/scripts/xxx.py`（绝对路径），解释器路径用 `<workspace>\.venv\Scripts\python.exe`
4. **临时文件**：执行完成后需清理 `extracted_data.json` 和调试文件
5. **纠偏循环**：如果发现大量数据缺失，应先导出原始文本到文件，对比诊断问题再修复脚本，不靠猜测修改
6. **AI 评估写入**：使用 `fill_excel.write_evaluations()` 方法。evaluations 字典 key 为**姓名+年龄**（如 `刘先生_27岁`），value 为 (匹配度, 简要评价) 元组
7. **`write_evaluations()` 的 key 匹配陷阱**：该函数通过 `eval_name in name_cell` 子串匹配 A 列文件名。简历中姓名可能重复（如多个"李先生"），因此 **key 必须包含年龄**（如 `李先生_35岁`）才能精确匹配到对应文件行。只用姓名会导致错误匹配到同姓的其他候选人。文件名格式为 `{姓名}_{年龄}岁_智联简历_{随机数}.docx`，用 `姓名_年龄` 即可唯一匹配
   - ⚠️ **但智联推荐页常出现「同名 + 同年龄」的重复候选人**（同一人被重复推荐，如 2 个 `陈先生_28岁`），此时 key 无法区分，`write_evaluations()` 会把两行写成同一评价。**遇到这种情况必须改用按行索引写入**（见 B2 首选方案），并如实标注"与上一条为同一候选人重复"
8. **H 列只写岗位+公司+时间**：`work_experiences` 写入 Excel H 列的格式为 `【岗位】公司-时间`，**不含工作描述**。工作描述存于 `work_descriptions` 字段，仅供 AI 做行业匹配度判断时内部使用。行业判断同时基于公司名和工作描述内容双重交叉验证
9. **执行流程尽量自动化**：合并提取+填表步骤，用一条命令完成。评估写入也尽量一次完成。减少中间文件读写和用户授权确认步骤
10. **与 zhaopin-resume-saver 组合使用时**：下载脚本耗时可能较长，Agent 应在等待期间预计算评估参数，下载完成后立即进入提取→填表→AI评估写入流程，**不得停顿**。整个"下载+评估"链路应视为一个原子任务
