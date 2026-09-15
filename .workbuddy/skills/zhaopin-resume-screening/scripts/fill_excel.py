"""
将提取的简历数据写入 Excel 并应用格式。

两种输出模式:
1. fill_excel(): 旧 8 列格式（自建表头）——向后兼容
2. fill_excel_template(): 按 EvaluationTemplate.xlsx 模版填充（13 列，覆盖模版数据行）

Usage:
    python fill_excel.py <json_path> <xlsx_path>            # 旧格式
    python fill_excel.py <json_path> <template.xlsx> template  # 模版格式
"""

import json
import re
import sys
import copy
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# 模版 12 列表头（与 EvaluationTemplate.xlsx 一致；2026-09-15 起删除「建议行动」列）
TEMPLATE_HEADERS = ['简历名称', '胜任力评级', '简要评价', '年龄', '性别',
                    '工作年限', '现居住地', '学历', '求职状态', '期望薪资', '求职意向', '工作经历']
# 胜任力评级配色（取自模版实测值）
RATING_FILLS = {'高': 'C6EFCE', '中': 'FFFFB1', '低': 'FFD9D9'}
ZEBRA_FILL = 'F2F7FB'
CENTER_COLS = [2, 4, 5, 6, 7, 8, 9, 10]  # B/D/E/F/G/H/I/J 居中
# 列宽（字符）：简历名称/简要评价 36，求职意向/工作经历 50，其余 12
COLUMN_WIDTHS = {'A': 36, 'B': 12, 'C': 36, 'D': 12, 'E': 12, 'F': 12, 'G': 12,
                 'H': 12, 'I': 12, 'J': 12, 'K': 50, 'L': 50}


def dedup_resumes(data):
    """按「姓名+年龄」去重：重复投递的简历只保留首份。

    从 filename 解析去重键，如 '吴先生_37岁_智联简历_21020.docx' -> ('吴先生', '37')。
    Returns: (unique_data, skipped_count)
    """
    seen = set()
    unique_data = []
    for r in data:
        fname = str(r.get('filename', ''))
        m = re.match(r'^(.+?)_(\d+)岁_', fname)
        key = (m.group(1), m.group(2)) if m else (fname, "")
        if key in seen:
            continue
        seen.add(key)
        unique_data.append(r)
    return unique_data, len(data) - len(unique_data)


def extract_expected_salary(job_intentions_text):
    """从求职意向文本提取第一条期望薪资，如 '8千-1.3万/月'。"""
    if not job_intentions_text:
        return ""
    m = re.search(r'】\s*([\d,.]+\s*[万千百元]?(?:\s*-\s*[\d,.]+\s*[万千百元]?)?\s*/\s*[年月])',
                  job_intentions_text)
    if m:
        return re.sub(r'\s+', '', m.group(1))
    m = re.search(r'([\d,.]+\s*[万千百元]?(?:\s*-\s*[\d,.]+\s*[万千百元]?)?\s*/\s*[年月])',
                  job_intentions_text)
    return re.sub(r'\s+', '', m.group(1)) if m else ""


def fill_excel_template(json_path, xlsx_path):
    """按 EvaluationTemplate.xlsx 模版填充（覆盖内容）。

    - 加载模版文件，保留第 1 行表头及其样式，删除第 2 行起全部数据行
    - 样式从模版残留首行数据复制（字体/边框/对齐/斑马纹），保证与模版一致
    - B 胜任力评级 / C 简要评价 留空，由 AI 评估后按行索引写入
    - 期望薪资列从求职意向文本自动提取第一条
    - 内置按「姓名+年龄」去重：重复投递简历只保留首份

    Args:
        json_path: extract_docx.py / extract_resumes.py 输出的 JSON 路径
        xlsx_path: 模版文件路径（就地覆盖内容）
    Returns:
        写入的数据行数（去重后）
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    data, skipped = dedup_resumes(data)

    wb = load_workbook(str(xlsx_path))
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active

    # 采样模版样式（删除前）
    ref_font = {}
    ref_border = {}
    ref_align = {}
    zebra_fill = PatternFill('solid', fgColor=ZEBRA_FILL)
    if ws.max_row >= 2:
        for c in range(1, 13):
            cell = ws.cell(row=2, column=c)
            ref_font[c] = copy.copy(cell.font)
            ref_border[c] = copy.copy(cell.border)
            ref_align[c] = copy.copy(cell.alignment)
        zf = ws.cell(row=4 if ws.max_row >= 4 else 2, column=1).fill
        if zf and zf.patternType:
            zebra_fill = PatternFill('solid', fgColor=str(zf.fgColor.rgb))

    # 清空模版中的旧数据行（覆盖内容）
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    data_font = Font(name='微软雅黑', size=10)
    bold_center = Font(name='微软雅黑', bold=True, size=11)
    # 所有单元格一律垂直居中 + 自动换行（wrap_text=True），无一例外
    top_align = Alignment(vertical='center', wrap_text=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    for i, r in enumerate(data, start=2):
        expected_salary = extract_expected_salary(r.get('job_intentions', ''))
        row_data = [
            r['filename'],
            '',  # B 胜任力评级 (AI)
            '',  # C 简要评价 (AI)
            r['age'],
            r['gender'],
            r.get('work_years', ''),
            r.get('current_residence', ''),
            r.get('education', ''),
            r.get('job_status', ''),
            expected_salary,
            r.get('job_intentions', ''),
            r.get('work_experiences', ''),
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.font = data_font
            cell.alignment = top_align
            cell.border = ref_border.get(col, thin_border)
            if col in CENTER_COLS:
                cell.alignment = center_align
            if i % 2 == 0:
                cell.fill = zebra_fill

        # 行高按最长文本行数自适应（模版基准 36）
        max_lines = 1
        for col in (1, 11, 12):
            v = row_data[col - 1]
            if v:
                max_lines = max(max_lines, str(v).count('\n') + 1)
        ws.row_dimensions[i].height = max(36, max_lines * 14 + 8)

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:L{len(data) + 1}"
    # 列宽固定规则（覆盖模版原列宽）：A/C=36，K/L=50，其余=12
    for col, w in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = w
    wb.save(str(xlsx_path))
    return len(data)


def write_evaluation_row(xlsx_path, row_index, rating, comment):
    """按行索引写入单条 AI 评估结果（row_index 从 0 计，Excel 行 = index + 2）。

    首选方案：按行索引写入可避免 write_evaluations() 在「同名同年龄重复候选人」时串行写错。
    （2026-09-15 起模版无「建议行动」列，仅写 B/C 两列；重复简历已在填充阶段去重。）
    """
    wb = load_workbook(str(xlsx_path))
    ws = wb["Sheet1"]
    row = row_index + 2

    def _apply(col, value, font, align, fill_color=None):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = font
        cell.alignment = align
        if fill_color:
            cell.fill = PatternFill('solid', fgColor=fill_color)

    _apply(2, rating, bold_center_font(), Alignment(horizontal='center', vertical='center', wrap_text=True),
           RATING_FILLS.get(rating))
    _apply(3, comment, Font(name='微软雅黑', size=10), Alignment(vertical='center', wrap_text=True))
    wb.save(str(xlsx_path))


def bold_center_font():
    return Font(name='微软雅黑', bold=True, size=11)


def fill_excel(json_path, xlsx_path):
    """Fill extracted resume data into Excel.

    Args:
        json_path: Path to extracted_data.json from extract_resumes.py.
        xlsx_path: Path to output Excel file (creates if not exists).
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    xlsx_path = Path(xlsx_path)
    if xlsx_path.exists():
        wb = load_workbook(str(xlsx_path))
        if "Sheet1" in wb.sheetnames:
            ws = wb["Sheet1"]
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
        else:
            ws = wb.active
            ws.title = "Sheet1"
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

    # Headers
    headers = ['简历名称', '简历匹配度', '简要评价', '年龄', '性别', '现居', '求职意向', '工作经历']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    # Column widths
    col_widths = {'A': 30, 'B': 14, 'C': 35, 'D': 8, 'E': 6, 'F': 10, 'G': 50, 'H': 65}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Styles
    header_fill = PatternFill('solid', fgColor='4472C4')
    header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_font = Font(name='微软雅黑', size=10)
    data_align = Alignment(vertical='top', wrap_text=True)
    center_align = Alignment(horizontal='center', vertical='top')
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )
    even_fill = PatternFill('solid', fgColor='F2F7FB')

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Data
    for i, r in enumerate(data, start=2):
        row_data = [
            r['filename'],
            '',  # 简历匹配度 (to be filled by AI)
            '',  # 简要评价 (to be filled by AI)
            r['age'],
            r['gender'],
            r['current_residence'],
            r['job_intentions'],
            r['work_experiences'],
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border

        for col in [2, 4, 5, 6]:
            ws.cell(row=i, column=col).alignment = center_align

        ws.row_dimensions[i].height = 80

        if i % 2 == 0:
            for col in range(1, 9):
                ws.cell(row=i, column=col).fill = even_fill

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:H{len(data) + 1}"

    wb.save(str(xlsx_path))
    return len(data)


def write_evaluations(xlsx_path, evaluations):
    """Write AI evaluation results (match level + tags) into Excel.

    Args:
        xlsx_path: Path to the Excel file.
        evaluations: Dict mapping name -> (match_level, tags_string).
    """
    from openpyxl import load_workbook

    wb = load_workbook(str(xlsx_path))
    ws = wb["Sheet1"]

    fill_colors = {
        '高': PatternFill('solid', fgColor='C6EFCE'),
        '中': PatternFill('solid', fgColor='FFEB9C'),
        '低': PatternFill('solid', fgColor='FFC7CE'),
    }

    for row in range(2, ws.max_row + 1):
        name_cell = ws.cell(row=row, column=1).value or ''
        # Extract name from filename (去掉姓氏+先生/女士后缀)
        for eval_name, (match_level, tags) in evaluations.items():
            if eval_name in name_cell:
                ws.cell(row=row, column=2, value=match_level)
                ws.cell(row=row, column=3, value=tags)
                cell = ws.cell(row=row, column=2)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name='微软雅黑', bold=True, size=11)
                if match_level in fill_colors:
                    cell.fill = fill_colors[match_level]
                ws.cell(row=row, column=3).font = Font(name='微软雅黑', size=10)
                ws.cell(row=row, column=3).alignment = Alignment(vertical='top', wrap_text=True)
                break

    wb.save(str(xlsx_path))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fill_excel.py <extracted_data.json> <output.xlsx> [template]")
        sys.exit(1)

    if len(sys.argv) > 3 and sys.argv[3] == "template":
        n = fill_excel_template(sys.argv[1], sys.argv[2])
        print(f"Filled {n} rows into template -> {sys.argv[2]}")
    else:
        n = fill_excel(sys.argv[1], sys.argv[2])
        print(f"Filled {n} rows -> {sys.argv[2]}")
