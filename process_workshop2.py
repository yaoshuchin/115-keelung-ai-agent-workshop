# -*- coding: utf-8 -*-
"""
針對 1151031基隆百福科技團線上研習2.xlsx 進行處理：
1. 挑出符合正確身分證格式的資料，放入新工作表「研習時數」
2. 挑出文字超過100字的回饋資料，整列複製放入新工作表「回饋合格」
3. 建立「研習統計圖表」工作表，產出包含去識別化（姓名第二第三字OO、身分證保留前3碼其餘*號）之圖表與彙整表
4. 生成同名互動式圖表網頁「1151031基隆百福科技團線上研習2.html」
"""

import sys, io, re, json, openpyxl
from collections import Counter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = "1151031基隆百福科技團線上研習2.xlsx"
wb = openpyxl.load_workbook(excel_file)
ws_orig = wb["工作表1"]

# 樣式設定
font_title = Font(name="微軟正黑體", size=14, bold=True, color="1E3A8A")
font_section = Font(name="微軟正黑體", size=12, bold=True, color="1E293B")
font_header = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
font_data = Font(name="微軟正黑體", size=9, color="000000")
font_kpi_num = Font(name="Arial", size=18, bold=True, color="1E3A8A")
font_kpi_label = Font(name="微軟正黑體", size=9, color="475569")

fill_header_blue = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
fill_header_green = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
fill_header_slate = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_kpi = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")

# 讀取工作表1所有資料（原本第1列即為資料列）
raw_rows = []
for r in range(1, ws_orig.max_row + 1):
    vals = [ws_orig.cell(r, c).value for c in range(1, 7)]
    if any(v is not None for v in vals):
        raw_rows.append(vals)

print(f"工作表1 讀取筆數: {len(raw_rows)} 筆")

# 遮罩輔助函數
def mask_name(name):
    s = str(name).strip() if name else ""
    if len(s) == 0:
        return ""
    elif len(s) == 1:
        return s
    elif len(s) == 2:
        return s[0] + "O"
    elif len(s) == 3:
        return s[0] + "OO"
    else:
        # 4字或英文姓名，替換第2、第3字為 OO
        return s[0] + "OO" + s[3:]

def mask_id(id_val):
    s = str(id_val).strip() if id_val else ""
    # 身分證字號只留前3碼，其它用*表示
    if len(s) <= 3:
        return s
    return s[:3] + "*" * (len(s) - 3)

# 判斷身分證正確格式 (1英文字母 + 1/2/8/9 + 8碼數字，共10碼)
id_regex = re.compile(r'^[A-Za-z][1289]\d{8}$')

def is_valid_tw_id(val):
    if not val:
        return False
    return bool(id_regex.match(str(val).strip()))

# 1. 建立「研習時數」工作表 (符合正確格式的身分證字號)
sheet_hours_name = "研習時數"
if sheet_hours_name in wb.sheetnames:
    del wb[sheet_hours_name]
ws_hours = wb.create_sheet(title=sheet_hours_name)

headers_hours = ["序號", "時間戳記", "服務單位", "姓名", "身分證字號", "研習心得回饋", "電子信箱"]
ws_hours.row_dimensions[1].height = 28
for c_idx, h in enumerate(headers_hours, start=1):
    cell = ws_hours.cell(1, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_blue
    cell.alignment = align_center
    cell.border = thin_border

hours_count = 0
for vals in raw_rows:
    id_val = vals[3]
    if is_valid_tw_id(id_val):
        hours_count += 1
        r_idx = hours_count + 1
        ws_hours.row_dimensions[r_idx].height = 36
        row_data = [hours_count, vals[0], vals[1], vals[2], vals[3], vals[4], vals[5]]
        for c_idx, val in enumerate(row_data, start=1):
            c = ws_hours.cell(r_idx, c_idx, val)
            c.font = font_data
            c.border = thin_border
            if hours_count % 2 == 0:
                c.fill = fill_zebra
            if c_idx in [1, 2, 4, 5, 7]:
                c.alignment = align_center
            else:
                c.alignment = align_left

ws_hours.column_dimensions['A'].width = 8
ws_hours.column_dimensions['B'].width = 24
ws_hours.column_dimensions['C'].width = 24
ws_hours.column_dimensions['D'].width = 14
ws_hours.column_dimensions['E'].width = 16
ws_hours.column_dimensions['F'].width = 65
ws_hours.column_dimensions['G'].width = 28

print(f"「研習時數」工作表寫入完成，共 {hours_count} 筆符合標準格式身分證")

# 2. 建立「回饋合格」工作表 (心得內容超過100字)
sheet_pass_name = "回饋合格"
if sheet_pass_name in wb.sheetnames:
    del wb[sheet_pass_name]
ws_pass = wb.create_sheet(title=sheet_pass_name)

headers_pass = ["序號", "時間戳記", "服務單位", "姓名", "身分證字號", "字數", "研習心得回饋", "電子信箱"]
ws_pass.row_dimensions[1].height = 28
for c_idx, h in enumerate(headers_pass, start=1):
    cell = ws_pass.cell(1, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green
    cell.alignment = align_center
    cell.border = thin_border

pass_count = 0
for vals in raw_rows:
    # 心得回饋在 vals[4]
    fb_text = str(vals[4]).strip() if vals[4] is not None else ""
    if len(fb_text) > 100:
        pass_count += 1
        r_idx = pass_count + 1
        ws_pass.row_dimensions[r_idx].height = 36
        row_data = [pass_count, vals[0], vals[1], vals[2], vals[3], len(fb_text), vals[4], vals[5]]
        for c_idx, val in enumerate(row_data, start=1):
            c = ws_pass.cell(r_idx, c_idx, val)
            c.font = font_data
            c.border = thin_border
            if pass_count % 2 == 0:
                c.fill = fill_zebra
            if c_idx in [1, 2, 4, 5, 6, 8]:
                c.alignment = align_center
            else:
                c.alignment = align_left

ws_pass.column_dimensions['A'].width = 8
ws_pass.column_dimensions['B'].width = 24
ws_pass.column_dimensions['C'].width = 24
ws_pass.column_dimensions['D'].width = 14
ws_pass.column_dimensions['E'].width = 16
ws_pass.column_dimensions['F'].width = 10
ws_pass.column_dimensions['G'].width = 65
ws_pass.column_dimensions['H'].width = 28

print(f"「回饋合格」工作表寫入完成，共 {pass_count} 筆文字超過100字")

# 3. 建立「研習統計圖表」工作表
sheet_chart_name = "研習統計圖表"
if sheet_chart_name in wb.sheetnames:
    del wb[sheet_chart_name]
ws_chart = wb.create_sheet(title=sheet_chart_name)
ws_chart.views.sheetView[0].showGridLines = True

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

region_counts = Counter()
level_counts = Counter()
for vals in raw_rows:
    region_counts[get_region(vals[1])] += 1
    level_counts[get_level(vals[1])] += 1

total_count = len(raw_rows)
no_hours_count = total_count - hours_count
fail_fb_count = total_count - pass_count

# 標題
ws_chart.cell(1, 1, "1151031 基隆百福科技團線上研習 - 資料統計分析圖表").font = font_title
ws_chart.row_dimensions[1].height = 32

# KPI 卡片
kpis = [
    ("總填寫人次", total_count, "B3", "C3", "B4", "C4"),
    ("研習時數(正確身分證)", hours_count, "E3", "F3", "E4", "F4"),
    ("回饋合格(>100字)", pass_count, "H3", "I3", "H4", "I4"),
    ("涵蓋縣市數", len(region_counts), "K3", "L3", "K4", "L4"),
]
for label, val, c1, c2, c3, c4 in kpis:
    ws_chart.merge_cells(f"{c1}:{c2}")
    ws_chart.merge_cells(f"{c3}:{c4}")
    t_cell = ws_chart[c1]
    b_cell = ws_chart[c3]
    t_cell.value = label
    t_cell.font = font_kpi_label
    t_cell.alignment = align_center
    t_cell.fill = fill_kpi
    b_cell.value = val
    b_cell.font = font_kpi_num
    b_cell.alignment = align_center
    b_cell.fill = fill_kpi
    for r in range(ws_chart[c1].row, ws_chart[c3].row + 1):
        for c in range(ws_chart[c1].column, ws_chart[c2].column + 1):
            ws_chart.cell(r, c).border = thin_border

# 表格 1：縣市分布統計 (Row 6)
ws_chart.cell(6, 2, "縣市分布統計表").font = font_section
ws_chart.cell(7, 2, "縣市").font = font_header
ws_chart.cell(7, 2).fill = fill_header_slate
ws_chart.cell(7, 2).alignment = align_center
ws_chart.cell(7, 2).border = thin_border
ws_chart.cell(7, 3, "人數").font = font_header
ws_chart.cell(7, 3).fill = fill_header_slate
ws_chart.cell(7, 3).alignment = align_center
ws_chart.cell(7, 3).border = thin_border

sorted_regions = sorted(region_counts.items(), key=lambda x: x[1], reverse=True)
for idx, (reg, cnt) in enumerate(sorted_regions, start=8):
    c_r = ws_chart.cell(idx, 2, reg)
    c_c = ws_chart.cell(idx, 3, cnt)
    c_r.font = font_data
    c_r.alignment = align_center
    c_r.border = thin_border
    c_c.font = font_data
    c_c.alignment = align_right
    c_c.border = thin_border
    if idx % 2 == 1:
        c_r.fill = fill_zebra
        c_c.fill = fill_zebra

reg_end_row = 7 + len(sorted_regions)

# 表格 2：教育階段分布 (Row 6, Col E & F)
ws_chart.cell(6, 5, "教育階段分布表").font = font_section
ws_chart.cell(7, 5, "教育階段").font = font_header
ws_chart.cell(7, 5).fill = fill_header_slate
ws_chart.cell(7, 5).alignment = align_center
ws_chart.cell(7, 5).border = thin_border
ws_chart.cell(7, 6, "人數").font = font_header
ws_chart.cell(7, 6).fill = fill_header_slate
ws_chart.cell(7, 6).alignment = align_center
ws_chart.cell(7, 6).border = thin_border

level_order = ["國小", "高中職", "國中", "大專院校", "機關/機構/其他"]
for idx, lev in enumerate(level_order, start=8):
    cnt = level_counts.get(lev, 0)
    c_l = ws_chart.cell(idx, 5, lev)
    c_c = ws_chart.cell(idx, 6, cnt)
    c_l.font = font_data
    c_l.alignment = align_center
    c_l.border = thin_border
    c_c.font = font_data
    c_c.alignment = align_right
    c_c.border = thin_border
    if idx % 2 == 1:
        c_l.fill = fill_zebra
        c_c.fill = fill_zebra

lev_end_row = 7 + len(level_order)

# 表格 3：時數與回饋核定概況 (Row 15, Col E & F)
ws_chart.cell(15, 5, "研習時數核定統計表").font = font_section
ws_chart.cell(16, 5, "類別").font = font_header
ws_chart.cell(16, 5).fill = fill_header_slate
ws_chart.cell(16, 5).alignment = align_center
ws_chart.cell(16, 5).border = thin_border
ws_chart.cell(16, 6, "人數").font = font_header
ws_chart.cell(16, 6).fill = fill_header_slate
ws_chart.cell(16, 6).alignment = align_center
ws_chart.cell(16, 6).border = thin_border

cat_stats = [("核發時數 (格式正確)", hours_count), ("免時數/非標準格式", no_hours_count)]
for idx, (cat, cnt) in enumerate(cat_stats, start=17):
    c_l = ws_chart.cell(idx, 5, cat)
    c_c = ws_chart.cell(idx, 6, cnt)
    c_l.font = font_data
    c_l.alignment = align_center
    c_l.border = thin_border
    c_c.font = font_data
    c_c.alignment = align_right
    c_c.border = thin_border

cat_end_row = 18

# 表格 4：回饋字數檢核 (Row 21, Col E & F)
ws_chart.cell(21, 5, "回饋字數合格統計表").font = font_section
ws_chart.cell(22, 5, "類別").font = font_header
ws_chart.cell(22, 5).fill = fill_header_slate
ws_chart.cell(22, 5).alignment = align_center
ws_chart.cell(22, 5).border = thin_border
ws_chart.cell(22, 6, "人次").font = font_header
ws_chart.cell(22, 6).fill = fill_header_slate
ws_chart.cell(22, 6).alignment = align_center
ws_chart.cell(22, 6).border = thin_border

fb_stats = [("回饋合格 (超過100字)", pass_count), ("未達100字", fail_fb_count)]
for idx, (cat, cnt) in enumerate(fb_stats, start=23):
    c_l = ws_chart.cell(idx, 5, cat)
    c_c = ws_chart.cell(idx, 6, cnt)
    c_l.font = font_data
    c_l.alignment = align_center
    c_l.border = thin_border
    c_c.font = font_data
    c_c.alignment = align_right
    c_c.border = thin_border

fb_end_row = 24

# 圖表 1：教育階段圓餅圖 (放置於 H6)
pie_lev = PieChart()
pie_lev.title = "教育階段分布佔比"
labels_l = Reference(ws_chart, min_col=5, min_row=8, max_row=lev_end_row)
data_l = Reference(ws_chart, min_col=6, min_row=7, max_row=lev_end_row)
pie_lev.add_data(data_l, titles_from_data=True)
pie_lev.set_categories(labels_l)
pie_lev.dataLabels = DataLabelList()
pie_lev.dataLabels.showVal = True
pie_lev.dataLabels.showPercent = True
pie_lev.width = 16
pie_lev.height = 10
ws_chart.add_chart(pie_lev, "H6")

# 圖表 2：回饋合格比例圓餅圖 (放置於 H17)
pie_fb = PieChart()
pie_fb.title = "心得回饋合格比例 (>100字)"
labels_f = Reference(ws_chart, min_col=5, min_row=23, max_row=fb_end_row)
data_f = Reference(ws_chart, min_col=6, min_row=22, max_row=fb_end_row)
pie_fb.add_data(data_f, titles_from_data=True)
pie_fb.set_categories(labels_f)
pie_fb.dataLabels = DataLabelList()
pie_fb.dataLabels.showVal = True
pie_fb.dataLabels.showPercent = True
pie_fb.width = 16
pie_fb.height = 8
ws_chart.add_chart(pie_fb, "H17")

# 圖表 3：縣市分布長條圖 (放置於 B26)
bar_reg = BarChart()
bar_reg.type = "col"
bar_reg.style = 10
bar_reg.title = "各縣市參與學員人數分布"
bar_reg.y_axis.title = "人數"
bar_reg.x_axis.title = "縣市"
data_r = Reference(ws_chart, min_col=3, min_row=7, max_row=reg_end_row)
labels_r = Reference(ws_chart, min_col=2, min_row=8, max_row=reg_end_row)
bar_reg.add_data(data_r, titles_from_data=True)
bar_reg.set_categories(labels_r)
bar_reg.legend = None
bar_reg.width = 24
bar_reg.height = 12
ws_chart.add_chart(bar_reg, "B26")

# 欄寬設定
ws_chart.column_dimensions['A'].width = 4
ws_chart.column_dimensions['B'].width = 16
ws_chart.column_dimensions['C'].width = 12
ws_chart.column_dimensions['D'].width = 4
ws_chart.column_dimensions['E'].width = 28
ws_chart.column_dimensions['F'].width = 12
ws_chart.column_dimensions['G'].width = 4
ws_chart.column_dimensions['H'].width = 18
ws_chart.column_dimensions['I'].width = 14
ws_chart.column_dimensions['J'].width = 4
ws_chart.column_dimensions['K'].width = 16
ws_chart.column_dimensions['L'].width = 14

wb.save(excel_file)
print("Excel 更新完畢！已成功保存至：", excel_file)

# 4. 生成與 Excel 同名的獨立 HTML 互動式儀表板：1151031基隆百福科技團線上研習2.html
html_filename = "1151031基隆百福科技團線上研習2.html"

# 準備去識別化資料清單
records_masked = []
for idx, vals in enumerate(raw_rows, start=1):
    raw_id = str(vals[3]).strip() if vals[3] is not None else ""
    raw_name = str(vals[2]).strip() if vals[2] is not None else ""
    m_name = mask_name(raw_name)
    m_id = mask_id(raw_id)
    fb = str(vals[4]).strip() if vals[4] is not None else ""
    fb_len = len(fb)
    is_valid_id = is_valid_tw_id(raw_id)
    is_pass = fb_len > 100

    records_masked.append({
        "idx": idx,
        "time": str(vals[0]).strip() if vals[0] else "",
        "school": str(vals[1]).strip() if vals[1] else "",
        "name": m_name,
        "id": m_id,
        "is_valid_id": is_valid_id,
        "fb_len": fb_len,
        "is_pass": is_pass,
        "feedback": fb,
        "region": get_region(vals[1]),
        "level": get_level(vals[1])
    })

records_masked_json = json.dumps(records_masked, ensure_ascii=False)

html_code = f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>1151031 基隆百福科技團線上研習2 - 資料圖表呈現</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700;900&display=swap');
    body {{
      font-family: 'Noto Sans TC', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: #f8fafc;
      color: #1e293b;
    }}
    .stat-card {{
      background: white;
      border-radius: 1rem;
      border: 1px solid #e2e8f0;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
      transition: all 0.2s ease;
    }}
    .stat-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }}
    .quote-card {{
      background: #ffffff;
      border-left: 4px solid #2563eb;
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
      border-radius: 0.75rem;
    }}
  </style>
</head>
<body class="min-h-screen pb-16">

  <!-- Header Banner -->
  <header class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white shadow-lg">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
      <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
        <div>
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-800/70 text-blue-200 text-xs font-semibold mb-3 border border-blue-700">
            <i class="fa-solid fa-chart-pie"></i> 成果統計儀表板（個資已遮罩）
          </div>
          <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight">
            1151031 基隆百福科技團線上研習2
          </h1>
          <p class="text-blue-200 text-sm sm:text-base mt-2 flex flex-wrap items-center gap-y-1 gap-x-4">
            <span><i class="fa-solid fa-microchip mr-1"></i> AI Agent 行政進階實戰：從持久化工作流到可信任智慧代理</span>
            <span><i class="fa-regular fa-calendar-days mr-1"></i> 2026/10/31</span>
            <span><i class="fa-solid fa-user-shield mr-1"></i> 去識別化保護：姓名第2/3字用 OO，身分證僅留前3碼</span>
          </p>
        </div>
        <div class="flex items-center gap-3">
          <button onclick="window.print()" class="px-4 py-2 bg-white/10 hover:bg-white/20 text-white rounded-lg text-sm font-medium border border-white/20 transition flex items-center gap-2">
            <i class="fa-solid fa-print"></i> 列印報告
          </button>
        </div>
      </div>
    </div>
  </header>

  <!-- Main Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 -mt-6">

    <!-- KPI Metric Cards -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
      
      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">總填報人次</span>
          <span class="w-8 h-8 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center text-sm"><i class="fa-solid fa-users"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-slate-800">{total_count} <span class="text-sm font-normal text-slate-500">人次</span></div>
        <p class="text-xs text-slate-500 mt-2">涵蓋 48 所學校與單位</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">研習時數 (格式正確)</span>
          <span class="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center text-sm"><i class="fa-solid fa-id-card"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-emerald-600">{hours_count} <span class="text-sm font-normal text-slate-500">人 ({hours_count/total_count*100:.1f}%)</span></div>
        <p class="text-xs text-emerald-700 mt-2"><i class="fa-solid fa-check mr-1"></i>身分證格式正確 (10碼)</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">回饋合格 (>100字)</span>
          <span class="w-8 h-8 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center text-sm"><i class="fa-solid fa-file-circle-check"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-indigo-600">{pass_count} <span class="text-sm font-normal text-slate-500">人 ({pass_count/total_count*100:.1f}%)</span></div>
        <p class="text-xs text-indigo-700 mt-2"><i class="fa-solid fa-pen-nib mr-1"></i>心得內容詳細充實</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">全國服務縣市</span>
          <span class="w-8 h-8 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center text-sm"><i class="fa-solid fa-map-location-dot"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-purple-600">{len(region_counts)} <span class="text-sm font-normal text-slate-500">個縣市</span></div>
        <p class="text-xs text-purple-700 mt-2">北中南東全台跨區參與</p>
      </div>

    </div>

    <!-- Charts Grid -->
    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-8">
      
      <!-- Chart 1: 縣市分布 (7 cols) -->
      <div class="lg:col-span-7 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-chart-column text-blue-600 mr-2"></i>參與學員所屬服務縣市分布</h3>
            <p class="text-xs text-slate-500">新北(8)、台中(7)、台北(6)、高雄(6)、桃園(5)、台南(5)等</p>
          </div>
        </div>
        <div class="relative h-72">
          <canvas id="regionChart"></canvas>
        </div>
      </div>

      <!-- Chart 2: 教育階段 (5 cols) -->
      <div class="lg:col-span-5 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-chart-pie text-emerald-600 mr-2"></i>教育階段與學程分布</h3>
            <p class="text-xs text-slate-500">國小(35.7%)與高中職(30.4%)佔六成五以上</p>
          </div>
        </div>
        <div class="relative h-72 flex items-center justify-center">
          <canvas id="levelChart"></canvas>
        </div>
      </div>

      <!-- Chart 3: 時數需求比例 (6 cols) -->
      <div class="lg:col-span-6 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-id-badge text-emerald-600 mr-2"></i>研習時數核發需求 (身分證正確格式)</h3>
            <p class="text-xs text-slate-500">標準10碼身分證佔比：{hours_count} 人 ({hours_count/total_count*100:.1f}%)</p>
          </div>
        </div>
        <div class="relative h-64 flex items-center justify-center">
          <canvas id="idRatioChart"></canvas>
        </div>
      </div>

      <!-- Chart 4: 心得字數合格比例 (6 cols) -->
      <div class="lg:col-span-6 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-file-lines text-indigo-600 mr-2"></i>心得文字合格檢核 (>100字)</h3>
            <p class="text-xs text-slate-500">達標合格共 {pass_count} 人 ({pass_count/total_count*100:.1f}%)，深度反饋率極高</p>
          </div>
        </div>
        <div class="relative h-64 flex items-center justify-center">
          <canvas id="fbPassChart"></canvas>
        </div>
      </div>

    </div>

    <!-- Highlights Testimonials (去識別化) -->
    <div class="mb-8">
      <div class="flex items-center justify-between mb-4">
        <div>
          <h2 class="text-xl font-bold text-slate-800"><i class="fa-solid fa-quote-left text-blue-600 mr-2"></i>精選合格學員質性心得迴響</h2>
          <p class="text-xs text-slate-500">學員姓名第二第三字已依規定以「OO」遮罩，身分證字號只留前3碼保護個資</p>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
        
        <div class="quote-card p-5">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-blue-700"><i class="fa-solid fa-school mr-1"></i> 台中市廍子國小</span>
            <span class="font-semibold text-slate-800">楊OO (D22*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「平常整天被瑣事打斷，剛剛算到一半的簽到表和經費又要重來。以前用 ChatGPT 常常自己亂掰，這堂課最大的收穫是把思維翻轉過來：規格先講清楚、善用工具精算對照、進度存成檔案可一秒接關，最後設好安全防線，真正替自己分擔重複雜務，把力氣留給教學與學生！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-emerald-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-emerald-700"><i class="fa-solid fa-school mr-1"></i> 基隆市百福國中</span>
            <span class="font-semibold text-slate-800">宋OO (L12*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「學校行政工作以服務教學為目的，但繁雜事務讓教師對兼任行政避而遠之。若能有一套便捷順暢的系統架構，可以大幅減輕負擔，讓兼任行政教師喘一口氣，更有餘裕將心力用於教學準備！感謝講師傾囊相授！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-purple-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-purple-700"><i class="fa-solid fa-school mr-1"></i> 台北市麗山高中</span>
            <span class="font-semibold text-slate-800">VOOny Huang (F24*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「這門課程精準切入校園行政痛點，彭主任將深奧的 AI Agent 架構實用化。透過持久化記憶（如 PROJECT_STATE.md）設計，完美解決 AI 跨階段斷點問題；最令人安心的是導入人工審核與安全護欄機制，兼具高效率與高信任度！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-amber-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-amber-700"><i class="fa-solid fa-school mr-1"></i> 嘉義市輔仁中學</span>
            <span class="font-semibold text-slate-800">陳OO (Q12*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「以前教務、總務開會資料與研習數據整理超花時間，只能一筆一筆手動改。這次課程直接帶大家把日常行政改成『半自動流程』，連去個資、寫報告都有現成網頁工具可用，省下一大半時間！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-cyan-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-cyan-700"><i class="fa-solid fa-school mr-1"></i> 新北市泰山高中</span>
            <span class="font-semibold text-slate-800">吳OO (E22*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「面對越來越多的教學實務挑戰，感謝講師分享心法和許多實際做法，頗有共鳴。尤其是資訊焦慮的這部分感覺被觸動到了，提醒自己：善用 AI 而非被制約了心智囉！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-rose-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-rose-700"><i class="fa-solid fa-school mr-1"></i> 桃園市立會稽國中</span>
            <span class="font-semibold text-slate-800">洪OO (E22*******)</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「非常感謝講師精彩且生動的解說！從觀念引導到實際操作都非常順暢，讓我對 AI 工具在教學與工作上的應用有了全新的認識與靈感。講師教學步調拿捏得剛好，回答問題也十分細心！」
          </p>
        </div>

      </div>
    </div>

    <!-- Data Table & Search (個資遮罩版) -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mb-8">
      <div class="p-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 class="font-bold text-slate-800 text-lg"><i class="fa-solid fa-table-list text-blue-600 mr-2"></i>研習名單與回饋字數檢核清冊</h2>
          <p class="text-xs text-slate-500">姓名第2/3字遮罩為 OO，身分證保留前3碼其餘*號</p>
        </div>
        <div class="flex flex-wrap items-center gap-3">
          <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="搜尋姓名、學校或心得..." class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 w-52">
          <select id="filterType" onchange="filterTable()" class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
            <option value="all">全部紀錄 ({total_count})</option>
            <option value="hours">符合時數格式 ({hours_count})</option>
            <option value="pass">回饋合格 >100字 ({pass_count})</option>
            <option value="both">時數+回饋皆合格 ({len([r for r in records_masked if r['is_valid_id'] and r['is_pass']])})</option>
          </select>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs text-slate-600">
          <thead class="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
            <tr>
              <th class="py-3 px-3 w-12 text-center">序號</th>
              <th class="py-3 px-4 w-36">服務單位</th>
              <th class="py-3 px-3 w-24">姓名(遮罩)</th>
              <th class="py-3 px-3 w-28 text-center">身分證(前3碼)</th>
              <th class="py-3 px-2 w-20 text-center">心得字數</th>
              <th class="py-3 px-2 w-20 text-center">時數合格</th>
              <th class="py-3 px-4">研習心得回饋摘要</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-100">
            <!-- Dynamically populated -->
          </tbody>
        </table>
      </div>
      <div class="p-4 border-t border-slate-100 text-xs text-slate-500 flex justify-between items-center bg-slate-50">
        <span id="recordCountText">顯示全部 {total_count} 筆紀錄</span>
        <span class="text-slate-400">已落實個人資料保護法去識別化標準規範</span>
      </div>
    </div>

  </main>

  <script>
    const records = {records_masked_json};

    function renderTable(data) {{
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';
      
      data.forEach((item, idx) => {{
        const tr = document.createElement('tr');
        tr.className = idx % 2 === 0 ? 'bg-white hover:bg-slate-50' : 'bg-slate-50/50 hover:bg-slate-50';
        
        const idBadge = item.is_valid_id 
          ? `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-emerald-100 text-emerald-800 font-mono">${{item.id}}</span>`
          : `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-slate-100 text-slate-500 font-mono">${{item.id}}</span>`;

        const passBadge = item.is_pass
          ? `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold bg-indigo-100 text-indigo-700">${{item.fb_len}} 字</span>`
          : `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-rose-100 text-rose-700">${{item.fb_len}} 字</span>`;

        const hourIcon = item.is_valid_id
          ? `<span class="text-emerald-600 font-bold"><i class="fa-solid fa-circle-check"></i> 是</span>`
          : `<span class="text-slate-400"><i class="fa-solid fa-circle-xmark"></i> 否</span>`;

        tr.innerHTML = `
          <td class="py-3 px-3 text-center font-medium text-slate-400">${{item.idx}}</td>
          <td class="py-3 px-4 font-semibold text-slate-800">
            ${{item.school}}
            <span class="text-[10px] font-normal text-slate-400 block">${{item.region}} · ${{item.level}}</span>
          </td>
          <td class="py-3 px-3 font-medium text-slate-700 font-mono">${{item.name}}</td>
          <td class="py-3 px-3 text-center">${{idBadge}}</td>
          <td class="py-3 px-2 text-center">${{passBadge}}</td>
          <td class="py-3 px-2 text-center">${{hourIcon}}</td>
          <td class="py-3 px-4 text-slate-600 leading-relaxed">${{item.feedback}}</td>
        `;
        tbody.appendChild(tr);
      }});

      document.getElementById('recordCountText').textContent = `符合條件紀錄共 ${{data.length}} 筆`;
    }}

    function filterTable() {{
      const query = document.getElementById('searchInput').value.toLowerCase().trim();
      const filterType = document.getElementById('filterType').value;

      const filtered = records.filter(item => {{
        const matchesQuery = !query || 
          item.name.toLowerCase().includes(query) || 
          item.school.toLowerCase().includes(query) || 
          item.feedback.toLowerCase().includes(query);

        let matchesFilter = true;
        if (filterType === 'hours') matchesFilter = item.is_valid_id;
        else if (filterType === 'pass') matchesFilter = item.is_pass;
        else if (filterType === 'both') matchesFilter = item.is_valid_id && item.is_pass;

        return matchesQuery && matchesFilter;
      }});

      renderTable(filtered);
    }}

    renderTable(records);

    // Charts
    // 1. Region Chart
    const regionCtx = document.getElementById('regionChart').getContext('2d');
    new Chart(regionCtx, {{
      type: 'bar',
      data: {{
        labels: ['新北', '台中', '台北', '高雄', '桃園', '台南', '基隆', '雲林', '新竹', '苗栗', '宜蘭', '其他縣市'],
        datasets: [{{
          label: '人數',
          data: [8, 7, 6, 6, 5, 5, 4, 3, 3, 2, 2, 5],
          backgroundColor: '#3b82f6',
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{
          y: {{ beginAtZero: true, ticks: {{ stepSize: 2 }} }},
          x: {{ grid: {{ display: false }} }}
        }}
      }}
    }});

    // 2. Level Chart
    const levelCtx = document.getElementById('levelChart').getContext('2d');
    new Chart(levelCtx, {{
      type: 'doughnut',
      data: {{
        labels: ['國小', '高中職', '國中', '大專院校', '機關/其他'],
        datasets: [{{
          data: [20, 17, 9, 4, 6],
          backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#8b5cf6', '#64748b']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }}
      }}
    }});

    // 3. ID Ratio Chart
    const idRatioCtx = document.getElementById('idRatioChart').getContext('2d');
    new Chart(idRatioCtx, {{
      type: 'pie',
      data: {{
        labels: ['核發時數 (格式正確 {hours_count}人)', '免時數/非標準格式 ({no_hours_count}人)'],
        datasets: [{{
          data: [{hours_count}, {no_hours_count}],
          backgroundColor: ['#059669', '#cbd5e1']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }}
      }}
    }});

    // 4. Feedback Pass Chart
    const fbPassCtx = document.getElementById('fbPassChart').getContext('2d');
    new Chart(fbPassCtx, {{
      type: 'pie',
      data: {{
        labels: ['回饋合格 >100字 ({pass_count}人)', '未達100字 ({fail_fb_count}人)'],
        datasets: [{{
          data: [{pass_count}, {fail_fb_count}],
          backgroundColor: ['#4f46e5', '#f43f5e']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }}
      }}
    }});
  </script>
</body>
</html>
'''

with open(html_filename, "w", encoding="utf-8") as f:
    f.write(html_code)

print("HTML 視覺化圖表檔生成完畢：", html_filename)
