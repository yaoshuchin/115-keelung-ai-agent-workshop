# -*- coding: utf-8 -*-
"""
處理 1151031基隆百福科技團線上研習 Excel 檔案：
1. 建立「有身分證字號」工作表，篩選挑出提供身分證之學員資料並加上格式檢核。
2. 建立「研習統計圖表」工作表，彙整縣市、學程教育階段、時數需求與填寫時段之統計數據並繪製 Excel 原生圖表。
3. 優化原「工作表1」表頭與版面樣式。
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
import re
from collections import Counter

excel_file = "1151031基隆百福科技團線上研習.xlsx"
wb = openpyxl.load_workbook(excel_file)

ws_orig = wb["工作表1"]

# 樣式定義
font_title = Font(name="微軟正黑體", size=14, bold=True, color="1E3A8A")
font_section = Font(name="微軟正黑體", size=12, bold=True, color="1E293B")
font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
font_data = Font(name="微軟正黑體", size=10, color="000000")
font_kpi_num = Font(name="Arial", size=18, bold=True, color="1E3A8A")
font_kpi_label = Font(name="微軟正黑體", size=9, color="475569")

fill_header_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
fill_header_blue = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
fill_header_slate = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_kpi = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
fill_highlight = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")

# 1. 整理「工作表1」
headers_orig = ["填寫時間", "服務單位", "姓名", "身分證字號", "研習心得與收穫"]
for c_idx, val in enumerate(headers_orig, start=1):
    cell = ws_orig.cell(1, c_idx, val)
    cell.font = font_header
    cell.fill = fill_header_navy
    cell.alignment = align_center
    cell.border = thin_border
ws_orig.row_dimensions[1].height = 28

for r in range(2, ws_orig.max_row + 1):
    ws_orig.row_dimensions[r].height = 36
    for c in range(1, 6):
        cell = ws_orig.cell(r, c)
        cell.font = font_data
        cell.border = thin_border
        if c in [1, 3, 4]:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

ws_orig.column_dimensions['A'].width = 24
ws_orig.column_dimensions['B'].width = 24
ws_orig.column_dimensions['C'].width = 14
ws_orig.column_dimensions['D'].width = 16
ws_orig.column_dimensions['E'].width = 65

# 2. 判斷是否有身分證字號
negatives = ['不需', '不用', '不需要', '無', 'no', '.', 'none']
def is_valid_or_intended_id(val):
    if not val:
        return False
    s = str(val).strip()
    if any(nw in s.lower() for nw in negatives):
        return False
    return True

def check_id_status(val):
    s = str(val).strip()
    if re.match(r'^[A-Za-z][1289]\d{8}$', s):
        return "正常 (標準10碼)"
    if re.match(r'^[A-Za-z]\d{7,8}$', s):
        return f"格式異常 ({len(s)}碼，疑漏碼)"
    if '-' in s:
        return "部分遮罩"
    return "非標準字號"

id_rows = []
for r in range(2, ws_orig.max_row + 1):
    vals = [ws_orig.cell(r, c).value for c in range(1, 6)]
    if not any(v is not None for v in vals):
        continue
    id_val = vals[3]
    if is_valid_or_intended_id(id_val):
        status = check_id_status(id_val)
        id_rows.append((vals, status))

# 建立「有身分證字號」工作表
sheet_name_id = "有身分證字號"
if sheet_name_id in wb.sheetnames:
    del wb[sheet_name_id]
ws_id = wb.create_sheet(title=sheet_name_id)

headers_id = ["序號", "填寫時間", "服務單位", "姓名", "身分證字號", "身分證檢核", "研習心得與收穫"]
ws_id.row_dimensions[1].height = 28
for c_idx, val in enumerate(headers_id, start=1):
    cell = ws_id.cell(1, c_idx, val)
    cell.font = font_header
    cell.fill = fill_header_blue
    cell.alignment = align_center
    cell.border = thin_border

for idx, (vals, status) in enumerate(id_rows, start=1):
    r_idx = idx + 1
    ws_id.row_dimensions[r_idx].height = 36
    row_data = [idx, vals[0], vals[1], vals[2], vals[3], status, vals[4]]
    for c_idx, val in enumerate(row_data, start=1):
        cell = ws_id.cell(r_idx, c_idx, val)
        cell.font = font_data
        cell.border = thin_border
        if idx % 2 == 0:
            cell.fill = fill_zebra
        if c_idx in [1, 2, 4, 5, 6]:
            cell.alignment = align_center
        else:
            cell.alignment = align_left
        
        # 標註異常狀態
        if c_idx == 6 and status != "正常 (標準10碼)":
            cell.fill = fill_highlight

ws_id.column_dimensions['A'].width = 8
ws_id.column_dimensions['B'].width = 24
ws_id.column_dimensions['C'].width = 24
ws_id.column_dimensions['D'].width = 14
ws_id.column_dimensions['E'].width = 16
ws_id.column_dimensions['F'].width = 20
ws_id.column_dimensions['G'].width = 65

# 3. 數據統計與圖表工作表
sheet_name_stats = "研習統計圖表"
if sheet_name_stats in wb.sheetnames:
    del wb[sheet_name_stats]
ws_stats = wb.create_sheet(title=sheet_name_stats)
ws_stats.views.sheetView[0].showGridLines = True

# 統計分析
def get_region(s):
    s = str(s).strip()
    if '基隆' in s: return '基隆市'
    if '台北' in s or '臺北' in s: return '台北市'
    if '新北' in s: return '新北市'
    if '桃園' in s: return '桃園市'
    if '新竹' in s: return '新竹縣市'
    if '苗栗' in s: return '苗栗縣'
    if '台中' in s or '臺中' in s: return '台中市'
    if '彰化' in s: return '彰化縣'
    if '雲林' in s: return '雲林縣'
    if '嘉義' in s: return '嘉義市'
    if '台南' in s or '臺南' in s: return '台南市'
    if '高雄' in s: return '高雄市'
    if '屏東' in s: return '屏東縣'
    if '宜蘭' in s: return '宜蘭縣'
    if '花蓮' in s: return '花蓮縣'
    if '台東' in s or '臺東' in s: return '台東縣'
    return '其他'

def get_level(s):
    s = str(s).strip()
    if '國小' in s: return '國小'
    if '國中' in s: return '國中'
    if any(k in s for k in ['高中', '高職', '商工', '附中', '陽明']) and '國中' not in s: return '高中職'
    if any(k in s for k in ['大學', '科技大學', '科大']): return '大專院校'
    return '機關/機構/其他'

def get_time_period(t_str):
    s = str(t_str)
    if '上午 9:' in s or '上午 10:' in s:
        return '09:00 - 11:00 (課程進行中)'
    elif '上午 11:' in s:
        return '11:00 - 12:00 (課程後半/結尾)'
    elif '下午 12:' in s:
        return '12:00 - 13:00 (中午簽退高潮)'
    else:
        return '13:00 以後 (課後陸續回饋)'

region_counts = Counter()
level_counts = Counter()
time_counts = Counter()
total_records = 0
has_id_count = len(id_rows)

for r in range(2, ws_orig.max_row + 1):
    vals = [ws_orig.cell(r, c).value for c in range(1, 6)]
    if not any(v is not None for v in vals):
        continue
    total_records += 1
    region_counts[get_region(vals[1])] += 1
    level_counts[get_level(vals[1])] += 1
    time_counts[get_time_period(vals[0])] += 1

no_id_count = total_records - has_id_count

# 標題
ws_stats.cell(1, 1, "基隆百福科技團線上研習 - 數據分析與視覺化統計").font = font_title
ws_stats.row_dimensions[1].height = 32

# KPI 卡片區
kpis = [
    ("總填寫人次", total_records, "B3", "C3", "B4", "C4"),
    ("需研習時數(有身分證)", has_id_count, "E3", "F3", "E4", "F4"),
    ("免時數需求", no_id_count, "H3", "I3", "H4", "I4"),
    ("涵蓋縣市數", len(region_counts), "K3", "L3", "K4", "L4"),
]

for label, val, c1, c2, c3, c4 in kpis:
    ws_stats.merge_cells(f"{c1}:{c2}")
    ws_stats.merge_cells(f"{c3}:{c4}")
    top_cell = ws_stats[c1]
    bottom_cell = ws_stats[c3]
    top_cell.value = label
    top_cell.font = font_kpi_label
    top_cell.alignment = align_center
    top_cell.fill = fill_kpi
    bottom_cell.value = val
    bottom_cell.font = font_kpi_num
    bottom_cell.alignment = align_center
    bottom_cell.fill = fill_kpi
    
    for r in range(ws_stats[c1].row, ws_stats[c3].row + 1):
        for c in range(ws_stats[c1].column, ws_stats[c2].column + 1):
            ws_stats.cell(r, c).border = thin_border

# 表格 1：縣市分布統計 (Row 6)
ws_stats.cell(6, 2, "縣市分布統計表").font = font_section
ws_stats.cell(7, 2, "縣市").font = font_header
ws_stats.cell(7, 2).fill = fill_header_slate
ws_stats.cell(7, 2).alignment = align_center
ws_stats.cell(7, 2).border = thin_border

ws_stats.cell(7, 3, "人數").font = font_header
ws_stats.cell(7, 3).fill = fill_header_slate
ws_stats.cell(7, 3).alignment = align_center
ws_stats.cell(7, 3).border = thin_border

sorted_regions = sorted(region_counts.items(), key=lambda x: x[1], reverse=True)
for idx, (reg, cnt) in enumerate(sorted_regions, start=8):
    c_reg = ws_stats.cell(idx, 2, reg)
    c_cnt = ws_stats.cell(idx, 3, cnt)
    c_reg.font = font_data
    c_reg.alignment = align_center
    c_reg.border = thin_border
    c_cnt.font = font_data
    c_cnt.alignment = align_right
    c_cnt.border = thin_border
    if idx % 2 == 1:
        c_reg.fill = fill_zebra
        c_cnt.fill = fill_zebra

reg_end_row = 7 + len(sorted_regions)

# 表格 2：教育階段分布 (Row 6, Col E & F)
ws_stats.cell(6, 5, "教育階段分布表").font = font_section
ws_stats.cell(7, 5, "教育階段").font = font_header
ws_stats.cell(7, 5).fill = fill_header_slate
ws_stats.cell(7, 5).alignment = align_center
ws_stats.cell(7, 5).border = thin_border

ws_stats.cell(7, 6, "人數").font = font_header
ws_stats.cell(7, 6).fill = fill_header_slate
ws_stats.cell(7, 6).alignment = align_center
ws_stats.cell(7, 6).border = thin_border

level_order = ["國小", "高中職", "國中", "大專院校", "機關/機構/其他"]
for idx, lev in enumerate(level_order, start=8):
    cnt = level_counts.get(lev, 0)
    c_lev = ws_stats.cell(idx, 5, lev)
    c_cnt = ws_stats.cell(idx, 6, cnt)
    c_lev.font = font_data
    c_lev.alignment = align_center
    c_lev.border = thin_border
    c_cnt.font = font_data
    c_cnt.alignment = align_right
    c_cnt.border = thin_border
    if idx % 2 == 1:
        c_lev.fill = fill_zebra
        c_cnt.fill = fill_zebra

lev_end_row = 7 + len(level_order)

# 表格 3：研習時數需求統計 (Row 15, Col E & F)
ws_stats.cell(15, 5, "研習時數核發需求表").font = font_section
ws_stats.cell(16, 5, "時數需求類別").font = font_header
ws_stats.cell(16, 5).fill = fill_header_slate
ws_stats.cell(16, 5).alignment = align_center
ws_stats.cell(16, 5).border = thin_border

ws_stats.cell(16, 6, "人數").font = font_header
ws_stats.cell(16, 6).fill = fill_header_slate
ws_stats.cell(16, 6).alignment = align_center
ws_stats.cell(16, 6).border = thin_border

id_stats = [("需核發時數 (已填身分證)", has_id_count), ("免核發時數 (不需/無)", no_id_count)]
for idx, (cat, cnt) in enumerate(id_stats, start=17):
    c_cat = ws_stats.cell(idx, 5, cat)
    c_cnt = ws_stats.cell(idx, 6, cnt)
    c_cat.font = font_data
    c_cat.alignment = align_center
    c_cat.border = thin_border
    c_cnt.font = font_data
    c_cnt.alignment = align_right
    c_cnt.border = thin_border

id_end_row = 18

# 表格 4：填報時段分布 (Row 21, Col E & F)
ws_stats.cell(21, 5, "回饋填報時間分布表").font = font_section
ws_stats.cell(22, 5, "時段區間").font = font_header
ws_stats.cell(22, 5).fill = fill_header_slate
ws_stats.cell(22, 5).alignment = align_center
ws_stats.cell(22, 5).border = thin_border

ws_stats.cell(22, 6, "人次").font = font_header
ws_stats.cell(22, 6).fill = fill_header_slate
ws_stats.cell(22, 6).alignment = align_center
ws_stats.cell(22, 6).border = thin_border

time_order = [
    '09:00 - 11:00 (課程進行中)',
    '11:00 - 12:00 (課程後半/結尾)',
    '12:00 - 13:00 (中午簽退高潮)',
    '13:00 以後 (課後陸續回饋)'
]
for idx, t_label in enumerate(time_order, start=23):
    cnt = time_counts.get(t_label, 0)
    c_t = ws_stats.cell(idx, 5, t_label)
    c_cnt = ws_stats.cell(idx, 6, cnt)
    c_t.font = font_data
    c_t.alignment = align_center
    c_t.border = thin_border
    c_cnt.font = font_data
    c_cnt.alignment = align_right
    c_cnt.border = thin_border

time_end_row = 26

# 設定欄寬
ws_stats.column_dimensions['A'].width = 4
ws_stats.column_dimensions['B'].width = 16
ws_stats.column_dimensions['C'].width = 12
ws_stats.column_dimensions['D'].width = 4
ws_stats.column_dimensions['E'].width = 30
ws_stats.column_dimensions['F'].width = 12
ws_stats.column_dimensions['G'].width = 4
ws_stats.column_dimensions['H'].width = 18
ws_stats.column_dimensions['I'].width = 14
ws_stats.column_dimensions['J'].width = 4
ws_stats.column_dimensions['K'].width = 16
ws_stats.column_dimensions['L'].width = 14

# 建立圖表 1：教育階段圓餅圖 (放置於 H6)
pie_level = PieChart()
pie_level.title = "教育階段分布佔比"
labels_level = Reference(ws_stats, min_col=5, min_row=8, max_row=lev_end_row)
data_level = Reference(ws_stats, min_col=6, min_row=7, max_row=lev_end_row)
pie_level.add_data(data_level, titles_from_data=True)
pie_level.set_categories(labels_level)
pie_level.dataLabels = DataLabelList()
pie_level.dataLabels.showVal = True
pie_level.dataLabels.showPercent = True
pie_level.width = 16
pie_level.height = 10
ws_stats.add_chart(pie_level, "H6")

# 建立圖表 2：時數需求圓餅圖 (放置於 H17)
pie_id = PieChart()
pie_id.title = "研習時數核發需求比例"
labels_id = Reference(ws_stats, min_col=5, min_row=17, max_row=id_end_row)
data_id = Reference(ws_stats, min_col=6, min_row=16, max_row=id_end_row)
pie_id.add_data(data_id, titles_from_data=True)
pie_id.set_categories(labels_id)
pie_id.dataLabels = DataLabelList()
pie_id.dataLabels.showVal = True
pie_id.dataLabels.showPercent = True
pie_id.width = 16
pie_id.height = 8
ws_stats.add_chart(pie_id, "H17")

# 建立圖表 3：縣市分布長條圖 (放置於 B26)
bar_region = BarChart()
bar_region.type = "col"
bar_region.style = 10
bar_region.title = "各縣市參與學員人數分布"
bar_region.y_axis.title = "人數"
bar_region.x_axis.title = "縣市"
data_reg = Reference(ws_stats, min_col=3, min_row=7, max_row=reg_end_row)
labels_reg = Reference(ws_stats, min_col=2, min_row=8, max_row=reg_end_row)
bar_region.add_data(data_reg, titles_from_data=True)
bar_region.set_categories(labels_reg)
bar_region.legend = None
bar_region.width = 24
bar_region.height = 12
ws_stats.add_chart(bar_region, "B26")

# 建立圖表 4：時間分布長條圖 (放置於 H26)
bar_time = BarChart()
bar_time.type = "col"
bar_time.style = 11
bar_time.title = "各時段回饋填寫人次"
bar_time.y_axis.title = "人次"
bar_time.x_axis.title = "時段"
data_time = Reference(ws_stats, min_col=6, min_row=22, max_row=time_end_row)
labels_time = Reference(ws_stats, min_col=5, min_row=23, max_row=time_end_row)
bar_time.add_data(data_time, titles_from_data=True)
bar_time.set_categories(labels_time)
bar_time.legend = None
bar_time.width = 16
bar_time.height = 12
ws_stats.add_chart(bar_time, "H26")

wb.save(excel_file)
print("處理完成！成功更新 Excel 檔案：", excel_file)
print(f"原表筆數: {total_records}, 有身分證筆數: {has_id_count}, 免時數筆數: {no_id_count}")
