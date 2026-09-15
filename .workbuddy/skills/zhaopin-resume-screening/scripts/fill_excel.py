"""
将提取的简历数据写入 Excel 并应用格式。

Usage:
    python fill_excel.py <json_path> <xlsx_path>
"""

import json
import sys
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


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
        print("Usage: python fill_excel.py <extracted_data.json> <output.xlsx>")
        sys.exit(1)

    n = fill_excel(sys.argv[1], sys.argv[2])
    print(f"Filled {n} rows → {sys.argv[2]}")
