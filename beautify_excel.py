"""
美化 周黑鸭毛利测算_最终.xlsx，生成可分享的专业报表
"""
import shutil
from openpyxl import load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule, ColorScaleRule
from copy import copy

SRC = '周黑鸭毛利测算_最终.xlsx'
OUT = '周黑鸭毛利测算_报表.xlsx'

# Backup
shutil.copy(SRC, SRC.replace('.xlsx', '_美化前备份.xlsx'))

wb = load_workbook(SRC)

# ============================================================
# Color scheme
# ============================================================
DARK_BLUE = '1F3864'
MED_BLUE = '2F5597'
LIGHT_BLUE = 'D6E4F0'
ACCENT_GOLD = 'FFC000'
WHITE = 'FFFFFF'
LIGHT_GRAY = 'F2F2F2'
MED_GRAY = 'D9D9D9'
DARK_GRAY = '404040'
RED_ACCENT = 'C00000'
GREEN_ACCENT = '375623'

header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type='solid')
header_font = Font(name='微软雅黑', bold=True, color=WHITE, size=11)
title_font = Font(name='微软雅黑', bold=True, color=DARK_BLUE, size=14)
subtitle_font = Font(name='微软雅黑', bold=True, color=MED_BLUE, size=12)
data_font = Font(name='微软雅黑', color=DARK_GRAY, size=10)
bold_font = Font(name='微软雅黑', bold=True, color=DARK_GRAY, size=10)
number_font = Font(name='微软雅黑', color=DARK_GRAY, size=11)

thin_border = Border(
    left=Side(style='thin', color=MED_GRAY),
    right=Side(style='thin', color=MED_GRAY),
    top=Side(style='thin', color=MED_GRAY),
    bottom=Side(style='thin', color=MED_GRAY),
)
bottom_border = Border(bottom=Side(style='medium', color=DARK_BLUE))
header_border = Border(
    left=Side(style='thin', color=DARK_BLUE),
    right=Side(style='thin', color=DARK_BLUE),
    top=Side(style='thin', color=DARK_BLUE),
    bottom=Side(style='medium', color=ACCENT_GOLD),
)
alt_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type='solid')
blue_tint_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
gold_fill = PatternFill(start_color=ACCENT_GOLD, end_color=ACCENT_GOLD, fill_type='solid')

center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
right_align = Alignment(horizontal='right', vertical='center')

def apply_cell(ws, row, col, value=None, font=None, fill=None, alignment=None, border=None, number_format=None):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = border
    if number_format: cell.number_format = number_format
    return cell

# ============================================================
# Sheet2 — 门店汇总仪表盘（设为首页）
# ============================================================
ws2 = wb['Sheet2']

# Set tab color
ws2.sheet_properties.tabColor = DARK_BLUE

# Clear stray data columns (B-D, G) — keep only the dashboard area
for row in range(1, 22):
    for col in [2, 3, 4, 7]:  # B, C, D, G
        ws2.cell(row=row, column=col).value = None

# Column widths
ws2.column_dimensions['A'].width = 3
ws2.column_dimensions['B'].width = 3
ws2.column_dimensions['C'].width = 3
ws2.column_dimensions['D'].width = 3
ws2.column_dimensions['E'].width = 3
ws2.column_dimensions['F'].width = 3
ws2.column_dimensions['G'].width = 3
ws2.column_dimensions['H'].width = 3
ws2.column_dimensions['I'].width = 16
ws2.column_dimensions['J'].width = 18
ws2.column_dimensions['K'].width = 16
ws2.column_dimensions['L'].width = 20
ws2.column_dimensions['M'].width = 16
ws2.column_dimensions['N'].width = 16

# Title row
ws2.merge_cells('I2:M2')
apply_cell(ws2, 2, 9, '周黑鸭门店毛利分析报表', font=title_font, alignment=center_align)
ws2.row_dimensions[2].height = 32

ws2.merge_cells('I3:M3')
apply_cell(ws2, 3, 9, '数据日期：2026年4月28日', font=Font(name='微软雅黑', color=MED_BLUE, size=10), alignment=center_align)

# Headers (row 4)
headers = ['门店', '销售额（元）', '整体毛利率', '高毛利产品占比', '平均折扣']
for i, h in enumerate(headers):
    apply_cell(ws2, 4, 9 + i, h, font=header_font, fill=header_fill, alignment=center_align, border=header_border)
ws2.row_dimensions[4].height = 28

# Data rows 5-10
stores_order = ['尚峰', '宣化', '民心', '长安', '定州', '未来石']
store_row_map = {ws2.cell(row=r, column=9).value: r for r in range(5, 11)}

for idx, store in enumerate(stores_order):
    r = store_row_map.get(store)
    if not r:
        continue
    row_fill = alt_fill if idx % 2 == 0 else None
    is_last = (idx == len(stores_order) - 1)

    for col_offset in range(5):  # I-M
        col = 9 + col_offset
        cell = ws2.cell(row=r, column=col)
        cell.font = bold_font if col_offset == 0 else number_font
        cell.alignment = center_align if col_offset == 0 else right_align
        cell.border = thin_border
        if is_last:
            cell.border = bottom_border
        if row_fill:
            cell.fill = row_fill

    # Number formats
    ws2.cell(row=r, column=10).number_format = '#,##0.00'  # 销售
    ws2.cell(row=r, column=11).number_format = '0.00%'     # 毛利率
    ws2.cell(row=r, column=12).number_format = '0.00%'     # 高毛利占比
    ws2.cell(row=r, column=13).number_format = '0.00%'     # 折扣

    # Color-code毛利率 cell
    margin_val = ws2.cell(row=r, column=11).value
    if margin_val and isinstance(margin_val, (int, float)):
        if margin_val >= 0.42:
            ws2.cell(row=r, column=11).font = Font(name='微软雅黑', bold=True, color=GREEN_ACCENT, size=11)
        elif margin_val < 0.30:
            ws2.cell(row=r, column=11).font = Font(name='微软雅黑', bold=True, color=RED_ACCENT, size=11)

ws2.row_dimensions[5].height = 24
ws2.row_dimensions[6].height = 24
ws2.row_dimensions[7].height = 24
ws2.row_dimensions[8].height = 24
ws2.row_dimensions[9].height = 24
ws2.row_dimensions[10].height = 24

# Notes at bottom
ws2.merge_cells('I12:M12')
apply_cell(ws2, 12, 9, '说明：整体毛利率 = 实际毛利 / 销售额；高毛利产品占比 = 单品毛利≥0.60的产品销售额占比；平均折扣 = 销售额 / 应收金额',
          font=Font(name='微软雅黑', color=MED_GRAY, size=9, italic=True), alignment=left_align)

# Freeze pane
ws2.freeze_panes = 'I5'

# ============================================================
# 毛利测算 sheet — 产品明细
# ============================================================
ws = wb['毛利测算']
ws.sheet_properties.tabColor = MED_BLUE

# Column widths
col_widths = {1: 22, 2: 26, 3: 12, 4: 14, 5: 22, 6: 12, 7: 12, 8: 12,
              9: 14, 10: 12, 11: 14, 12: 14, 13: 12, 14: 3, 15: 28, 16: 12,
              17: 12, 18: 12, 19: 12, 20: 14}
for c, w in col_widths.items():
    ws.column_dimensions[get_column_letter(c)].width = w

# Hide helper columns (O-R, T)
for c in ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W']:
    ws.column_dimensions[c].hidden = True

# --- Header row (row 1) ---
headers_main = {1: '日期', 2: '产品名称', 3: '单品毛利', 4: '销售数量', 5: '销售金额',
                6: '销售占比', 7: '成交单价', 8: '应收单价', 9: '应收金额', 10: '折扣',
                11: '成本', 12: '实际毛利额', 13: '实际毛利率'}
for c, h in headers_main.items():
    apply_cell(ws, 1, c, h, font=header_font, fill=header_fill, alignment=center_align, border=header_border)
ws.row_dimensions[1].height = 30

# Freeze header
ws.freeze_panes = 'A2'

# --- Product rows (2-18) ---
for row in range(2, 19):
    row_fill = alt_fill if row % 2 == 0 else None
    for col in range(1, 14):
        cell = ws.cell(row=row, column=col)
        cell.font = data_font
        cell.border = thin_border
        if row_fill:
            cell.fill = row_fill
        if col == 2:
            cell.alignment = left_align
        elif col == 1:
            cell.alignment = center_align
        else:
            cell.alignment = right_align

    # Number formats
    ws.cell(row=row, column=3).number_format = '0.00'      # 单品毛利
    ws.cell(row=row, column=4).number_format = '#,##0'      # 数量
    ws.cell(row=row, column=5).number_format = '#,##0.00'   # 销售金额
    ws.cell(row=row, column=6).number_format = '0.00%'      # 销售占比
    ws.cell(row=row, column=7).number_format = '#,##0.00'   # 成交单价
    ws.cell(row=row, column=8).number_format = '#,##0.00'   # 应收单价
    ws.cell(row=row, column=9).number_format = '#,##0.00'   # 应收金额
    ws.cell(row=row, column=10).number_format = '0.00%'     # 折扣
    ws.cell(row=row, column=11).number_format = '#,##0.00'  # 成本
    ws.cell(row=row, column=12).number_format = '#,##0.00'  # 毛利额
    ws.cell(row=row, column=13).number_format = '0.00%'     # 毛利率

    ws.row_dimensions[row].height = 22

# --- Summary row (19) ---
for col in range(1, 14):
    apply_cell(ws, 19, col, font=bold_font, fill=blue_tint_fill, border=bottom_border)
    if col in [4, 5, 9, 11, 12]:
        ws.cell(row=19, column=col).number_format = '#,##0.00'
    elif col in [6, 10, 13]:
        ws.cell(row=19, column=col).number_format = '0.00%'
    ws.cell(row=19, column=col).alignment = right_align if col > 1 else center_align
ws.cell(row=19, column=2).value = '合计'
ws.cell(row=19, column=2).font = Font(name='微软雅黑', bold=True, color=DARK_BLUE, size=11)
ws.row_dimensions[19].height = 26

# --- Key metrics rows (20-23) ---
# Row 20: labels
ws.merge_cells('F20:H20')
apply_cell(ws, 20, 6, '高毛利占比', font=Font(name='微软雅黑', bold=True, color=MED_BLUE, size=10), alignment=center_align)
ws.merge_cells('I20:K20')
apply_cell(ws, 20, 9, '整体折扣', font=Font(name='微软雅黑', bold=True, color=MED_BLUE, size=10), alignment=center_align)

# Row 21: values
apply_cell(ws, 21, 6, font=bold_font, alignment=center_align, number_format='0.00%',
           fill=PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid'))
apply_cell(ws, 21, 9, font=bold_font, alignment=center_align, number_format='0.00%',
           fill=PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid'))

# Row 22: labels for the three key metrics
apply_cell(ws, 22, 9, '整体毛利率', font=Font(name='微软雅黑', bold=True, color=MED_BLUE, size=10), alignment=center_align)
apply_cell(ws, 22, 10, '高毛利占比', font=Font(name='微软雅黑', bold=True, color=MED_BLUE, size=10), alignment=center_align)
apply_cell(ws, 22, 11, '整体折扣', font=Font(name='微软雅黑', bold=True, color=MED_BLUE, size=10), alignment=center_align)

# Row 23: I23/J23/K23 (final metrics)
for col in [9, 10, 11]:
    apply_cell(ws, 23, col, font=Font(name='微软雅黑', bold=True, color=DARK_BLUE, size=12),
               alignment=center_align, number_format='0.00%',
               fill=PatternFill(start_color=ACCENT_GOLD, end_color=ACCENT_GOLD, fill_type='solid'))

# --- Data area headers (row 24) ---
data_headers = {1: '日期', 2: '商品名称', 3: '品牌', 4: '销售数量', 5: '销售金额'}
for c, h in data_headers.items():
    apply_cell(ws, 24, c, h, font=Font(name='微软雅黑', bold=True, color=WHITE, size=10),
               fill=PatternFill(start_color=MED_BLUE, end_color=MED_BLUE, fill_type='solid'),
               alignment=center_align, border=thin_border)
ws.row_dimensions[24].height = 24

# Data area rows (25+)
for row in range(25, 80):
    val_in_row = any(ws.cell(row=row, column=c).value is not None for c in [1, 2, 4, 5])
    if val_in_row:
        for col in [1, 2, 3, 4, 5]:
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            if col == 2:
                cell.alignment = left_align
            elif col == 3:
                cell.alignment = center_align
            else:
                cell.alignment = right_align
        ws.cell(row=row, column=4).number_format = '#,##0'
        ws.cell(row=row, column=5).number_format = '#,##0.00'

# ============================================================
# Sheet1 — 简化汇总
# ============================================================
ws1 = wb['Sheet1']
ws1.sheet_properties.tabColor = MED_GRAY

for c, w in {1: 18, 2: 26, 3: 14, 4: 24, 5: 12, 6: 12, 7: 12, 8: 14}.items():
    ws1.column_dimensions[get_column_letter(c)].width = w

# Headers
for c, h in {1: '日期', 2: '产品名称', 3: '销售数量', 4: '销售金额', 5: '销售占比',
             6: '成交单价', 7: '折扣', 8: '实际毛利率'}.items():
    apply_cell(ws1, 1, c, h, font=header_font, fill=header_fill, alignment=center_align, border=header_border)
ws1.row_dimensions[1].height = 30
ws1.freeze_panes = 'A2'

for row in range(2, 20):
    row_fill = alt_fill if row % 2 == 0 else None
    for col in range(1, 9):
        cell = ws1.cell(row=row, column=col)
        if cell.value is not None:
            cell.font = data_font
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill
            cell.alignment = right_align if col > 2 else (left_align if col == 2 else center_align)
    ws1.cell(row=row, column=3).number_format = '#,##0'
    ws1.cell(row=row, column=4).number_format = '#,##0.00'
    for c in [5, 7, 8]:
        ws1.cell(row=row, column=c).number_format = '0.00%'
    ws1.cell(row=row, column=6).number_format = '#,##0.00'

# Rename sheets for clarity
ws2.title = '门店汇总'
ws1.title = '0401-0408汇总'

# Reorder sheets: 门店汇总 first, then 毛利测算, then 0401-0408汇总
wb.move_sheet('门店汇总', offset=-1)

# ============================================================
# Print settings
# ============================================================
for sn in wb.sheetnames:
    ws_ = wb[sn]
    ws_.sheet_properties.pageSetUpPr = None
    ws_.page_setup.orientation = 'landscape'
    ws_.page_setup.fitToWidth = 1
    ws_.page_setup.fitToHeight = 0

# Save
wb.save(OUT)
print(f'Saved to {OUT}')
print('Sheets: 门店汇总 | 毛利测算 | 0401-0408汇总')
