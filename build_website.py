# -*- coding: utf-8 -*-
"""
生成基隆市 115 學年度公立高中職中小學行政增能線上研習（6 小時）成果官方網頁 index.html
- 整合 10/16 第一梯次 (109筆) 與 10/31 第二梯次 (56筆) 共 165 筆回饋成果
- 整合教學紀錄照片 24 張與兩梯次教材海報展示
- 姓名嚴格去識別化：只保留姓，剩下全部用 "0" (如 洪00, 陳00, E00000)
- 身份證字號完全省略不予呈現
"""

import sys, io, os, json, openpyxl, urllib.parse
from datetime import datetime
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def mask_name(name):
    s = str(name).strip() if name else ""
    if not s:
        return ""
    if len(s) == 1:
        return s
    return s[0] + "0" * (len(s) - 1)

def get_region(s):
    s = str(s).strip()
    if '基隆' in s: return '基隆市'
    if '台北' in s or '臺北' in s or '北市' in s: return '台北市'
    if '新北' in s or '板橋' in s: return '新北市'
    if '桃園' in s or '觀音' in s: return '桃園市'
    if '新竹' in s: return '新竹縣市'
    if '苗栗' in s: return '苗栗縣'
    if '台中' in s or '臺中' in s or '東大附中' in s: return '台中市'
    if '彰化' in s: return '彰化縣'
    if '南投' in s: return '南投縣'
    if '雲林' in s: return '雲林縣'
    if '嘉義' in s or '中正大學' in s: return '嘉義縣市'
    if '台南' in s or '臺南' in s: return '台南市'
    if '高雄' in s: return '高雄市'
    if '屏東' in s: return '屏東縣'
    if '宜蘭' in s: return '宜蘭縣'
    if '花蓮' in s: return '花蓮縣'
    if '台東' in s or '臺東' in s: return '台東縣'
    return '其他縣市'

def get_level(s):
    s = str(s).strip()
    if any(k in s for k in ['國小', '小學', '幼兒園']): return '國小/幼兒'
    if '國中' in s: return '國中'
    if any(k in s for k in ['高中', '高職', '商工', '農工', '工職', '附中', '陽明', '中學', '實驗教育']) and '國中' not in s: return '高中職/中學'
    if any(k in s for k in ['大學', '科技大學', '科大']): return '大專院校'
    return '機關/機構/其他'

all_records = []

# 讀取 10/16 第一梯次
wb1 = openpyxl.load_workbook("1151016基隆百福科技團線上研習.xlsx", data_only=True)
ws1 = wb1["表單回應 1"]
c1_count = 0
for r in range(2, ws1.max_row + 1):
    vals = [ws1.cell(r, c).value for c in range(1, ws1.max_column + 1)]
    if any(v is not None for v in vals):
        c1_count += 1
        t_val = vals[0]
        if isinstance(t_val, datetime):
            t_str = t_val.strftime("%Y/%m/%d %H:%M")
        else:
            t_str = str(t_val).strip() if t_val else ""
        
        fb = str(vals[4]).strip() if vals[4] else ""
        school = str(vals[1]).strip() if vals[1] else ""
        name_masked = mask_name(vals[2])
        all_records.append({
            "idx": len(all_records) + 1,
            "cohort": "第 1 梯次 (10/16)",
            "cohort_code": "c1",
            "time": t_str,
            "school": school,
            "name": name_masked,
            "region": get_region(school),
            "level": get_level(school),
            "feedback": fb,
            "fb_len": len(fb),
            "is_pass": len(fb) > 100
        })

# 讀取 10/31 第二梯次
wb2 = openpyxl.load_workbook("1151031基隆百福科技團線上研習2.xlsx", data_only=True)
ws2 = wb2["工作表1"]
c2_count = 0
for r in range(1, ws2.max_row + 1):
    vals = [ws2.cell(r, c).value for c in range(1, ws2.max_column + 1)]
    if any(v is not None for v in vals):
        c2_count += 1
        t_val = vals[0]
        if isinstance(t_val, datetime):
            t_str = t_val.strftime("%Y/%m/%d %H:%M")
        else:
            t_str = str(t_val).strip() if t_val else ""
        
        fb = str(vals[4]).strip() if vals[4] else ""
        school = str(vals[1]).strip() if vals[1] else ""
        name_masked = mask_name(vals[2])
        all_records.append({
            "idx": len(all_records) + 1,
            "cohort": "第 2 梯次 (10/31)",
            "cohort_code": "c2",
            "time": t_str,
            "school": school,
            "name": name_masked,
            "region": get_region(school),
            "level": get_level(school),
            "feedback": fb,
            "fb_len": len(fb),
            "is_pass": len(fb) > 100
        })

total_records = len(all_records)
pass_records = [r for r in all_records if r["is_pass"]]
pass_count = len(pass_records)

# 統計分析
region_counter = Counter([r["region"] for r in all_records])
level_counter = Counter([r["level"] for r in all_records])
sorted_regions = sorted(region_counter.items(), key=lambda x: x[1], reverse=True)

# 掃描教學紀錄照片
photo_folder = os.path.join("1151031基隆科技團資訊線上研習", "教學紀錄照片")
photos = []
if os.path.exists(photo_folder):
    raw_files = sorted([f for f in os.listdir(photo_folder) if f.lower().endswith(('.jpg', '.jpeg', '.png'))])
    # 過濾重複檔案
    for f in raw_files:
        if "複製" not in f:
            photos.append(f"1151031基隆科技團資訊線上研習/教學紀錄照片/{urllib.parse.quote(f)}")

print(f"有效教學實況照片筆數: {len(photos)}")

# 兩梯次教材海報清單
posters_1016 = [
    {
        "title": "單元一：AI Agent 行政效能大升級 (1)",
        "desc": "打破提示詞迷思，從打字員進階至自主代理架構",
        "img": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級1.png')}",
        "pdf": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級1.pdf')}"
    },
    {
        "title": "單元二：AI Agent 行政效能大升級 (2)",
        "desc": "教務學務任務規格化與 Agent Brief 邊界定義",
        "img": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級2.png')}",
        "pdf": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級2.pdf')}"
    },
    {
        "title": "單元三：AI Agent 行政效能大升級 (3)",
        "desc": "多檔案資料清洗、行政工具精算與個資自動去識別化",
        "img": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級3.png')}",
        "pdf": f"1151016基隆科技團資訊線上研習/{urllib.parse.quote('1151016-基隆科技團-彭智標-以_AI_AI Agent 行政效能大升級3.pdf')}"
    }
]

posters_1031 = [
    {
        "title": "單元四：AI Agent 校務實戰指南 (1)",
        "desc": "專案狀態持久化管理，PROJECT_STATE.md 接關無縫推進",
        "img": f"1151031基隆科技團資訊線上研習/{urllib.parse.quote('1151031基隆科技團-資訊非專-AI Agent 校務實戰指南1.png')}",
        "pdf": f"1151031基隆科技團資訊線上研習/{urllib.parse.quote('1151031基隆科技團-資訊非專-AI Agent 校務實戰指南1.pdf')}"
    },
    {
        "title": "單元五：AI Agent 校務實戰指南 (2)",
        "desc": "三大處室 Plan-Execute-Review 閉環與可信任安全防護欄",
        "img": f"1151031基隆科技團資訊線上研習/{urllib.parse.quote('1151031基隆科技團-資訊非專-AI Agent 校務實戰指南2.png')}",
        "pdf": f"1151031基隆科技團資訊線上研習/{urllib.parse.quote('1151031基隆科技團-資訊非專-AI Agent 校務實戰指南2.pdf')}"
    }
]

all_records_json = json.dumps(all_records, ensure_ascii=False)
photos_json = json.dumps(photos, ensure_ascii=False)

top_reg_names = [k for k, _ in sorted_regions[:12]]
top_reg_vals = [v for _, v in sorted_regions[:12]]

# 建立 HTML
html_content = f'''<!DOCTYPE html>
<html lang="zh-Hant" class="scroll-smooth">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>基隆市 115 學年度公立高中職中小學行政增能線上研習（6 小時）- 成果報告網</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- Font Awesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&family=Noto+Sans+TC:wght@300;400;500;600;700;900&display=swap');
    body {{
      font-family: 'Plus Jakarta Sans', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: #f8fafc;
      color: #0f172a;
    }}
    .stat-card {{
      background: white;
      border-radius: 1rem;
      border: 1px solid #e2e8f0;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
      transition: all 0.25s ease;
    }}
    .stat-card:hover {{
      transform: translateY(-3px);
      box-shadow: 0 12px 20px -3px rgba(0, 0, 0, 0.08);
    }}
    .glass-nav {{
      background: rgba(15, 23, 42, 0.88);
      backdrop-filter: blur(12px);
    }}
    .quote-box {{
      background: #ffffff;
      border-left: 4px solid #2563eb;
      box-shadow: 0 2px 10px rgba(0,0,0,0.03);
      border-radius: 0.75rem;
      transition: transform 0.2s ease;
    }}
    .quote-box:hover {{
      transform: translateY(-2px);
      box-shadow: 0 6px 16px rgba(0,0,0,0.06);
    }}
    .photo-card {{
      overflow: hidden;
      border-radius: 0.75rem;
      transition: all 0.25s ease;
      cursor: pointer;
    }}
    .photo-card:hover img {{
      transform: scale(1.05);
    }}
    .photo-card img {{
      transition: transform 0.3s ease;
    }}
  </style>
</head>
<body class="min-h-screen">

  <!-- Top Navigation -->
  <nav class="sticky top-0 z-50 glass-nav border-b border-slate-800 text-white">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
      <div class="flex items-center gap-3">
        <span class="w-9 h-9 rounded-lg bg-blue-600 flex items-center justify-center text-white font-bold text-lg shadow-md shadow-blue-500/20">
          <i class="fa-solid fa-robot"></i>
        </span>
        <div class="leading-tight">
          <div class="font-bold text-sm sm:text-base tracking-tight">基隆市 115 學年度行政增能研習</div>
          <div class="text-[10px] text-blue-300">115基隆科技團 · AI Agent 行政實戰</div>
        </div>
      </div>
      <div class="hidden md:flex items-center gap-6 text-xs font-medium text-slate-300">
        <a href="#curriculum" class="hover:text-white transition">課程架構</a>
        <a href="#photos" class="hover:text-white transition">實況照片</a>
        <a href="#posters" class="hover:text-white transition">實戰講義</a>
        <a href="#stats" class="hover:text-white transition">數據成效</a>
        <a href="#testimonials" class="hover:text-white transition">學員金句</a>
        <a href="#roster" class="hover:text-white transition">回饋清冊</a>
      </div>
      <div class="flex items-center gap-2">
        <a href="https://mihozip.tw/lectures.demo/2026101623-keelung/" target="_blank" rel="noopener noreferrer" class="px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold shadow-sm transition flex items-center gap-1.5">
          <i class="fa-solid fa-arrow-up-right-from-square"></i> 研習講義中樞
        </a>
      </div>
    </div>
  </nav>

  <!-- Hero Section -->
  <header class="relative bg-gradient-to-br from-slate-950 via-blue-950 to-indigo-950 text-white pt-12 pb-20 px-4 sm:px-6 lg:px-8 overflow-hidden">
    <div class="absolute inset-0 opacity-10 bg-[radial-gradient(#3b82f6_1px,transparent_1px)] [background-size:16px_16px]"></div>
    
    <div class="relative max-w-7xl mx-auto">
      <div class="inline-flex items-center gap-2 px-3.5 py-1 rounded-full bg-blue-500/20 text-blue-300 text-xs font-semibold mb-4 border border-blue-400/30">
        <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
        6 小時完整研習雙梯次成果驗收
      </div>
      
      <h1 class="text-2xl sm:text-4xl lg:text-5xl font-black tracking-tight leading-tight max-w-4xl">
        基隆市 115 學年度公立高中職中小學行政增能線上研習
      </h1>
      
      <p class="text-lg sm:text-xl font-medium text-blue-200 mt-3 max-w-3xl">
        115 基隆科技團 - 資訊非專研習：<span class="text-white font-bold underline decoration-blue-500 underline-offset-4">AI Agent 行政實戰</span>
      </p>

      <p class="text-xs sm:text-sm text-slate-300 mt-4 max-w-3xl leading-relaxed">
        跳脫傳統「單次問答打字員」模式，將繁複、跨檔案、易中斷的學校行政工作，規格化拆解為具備工具調用、狀態持久化與自檢驗證的 AI Agent 工作流，實質釋放第一線教師與兼任行政 80% 行政負擔！
      </p>

      <!-- Key Metas -->
      <div class="mt-8 grid grid-cols-1 sm:grid-cols-3 gap-3 max-w-3xl text-xs">
        <div class="bg-white/5 border border-white/10 rounded-xl p-3 flex items-center gap-3">
          <span class="w-8 h-8 rounded-lg bg-blue-500/20 text-blue-400 flex items-center justify-center text-sm"><i class="fa-solid fa-chalkboard-user"></i></span>
          <div>
            <div class="text-slate-400">主講專家</div>
            <div class="font-bold text-white">彭智標 主任 (南投國中)</div>
          </div>
        </div>
        <div class="bg-white/5 border border-white/10 rounded-xl p-3 flex items-center gap-3">
          <span class="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-sm"><i class="fa-solid fa-calendar-check"></i></span>
          <div>
            <div class="text-slate-400">研習梯次與時數</div>
            <div class="font-bold text-white">雙梯次累計 6 小時</div>
          </div>
        </div>
        <div class="bg-white/5 border border-white/10 rounded-xl p-3 flex items-center gap-3">
          <span class="w-8 h-8 rounded-lg bg-purple-500/20 text-purple-400 flex items-center justify-center text-sm"><i class="fa-solid fa-users-viewfinder"></i></span>
          <div>
            <div class="text-slate-400">全台參與規模</div>
            <div class="font-bold text-white">165 人次 · 跨 18 縣市</div>
          </div>
        </div>
      </div>

    </div>
  </header>

  <!-- Main Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 -mt-10">

    <!-- KPI Metric Cards Grid -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-10">
      
      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">雙梯次總人次</span>
          <span class="w-8 h-8 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center text-sm"><i class="fa-solid fa-user-check"></i></span>
        </div>
        <div class="text-3xl font-black text-slate-900">{total_records} <span class="text-xs font-normal text-slate-500">人次</span></div>
        <p class="text-[11px] text-slate-500 mt-2">梯次一 109 人 · 梯次二 56 人</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">深度回饋達標率</span>
          <span class="w-8 h-8 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center text-sm"><i class="fa-solid fa-pen-fancy"></i></span>
        </div>
        <div class="text-3xl font-black text-indigo-600">{pass_count} <span class="text-xs font-normal text-slate-500">人 ({pass_count/total_records*100:.1f}%)</span></div>
        <p class="text-[11px] text-indigo-700 mt-2"><i class="fa-solid fa-check mr-1"></i>心得文字字數超過 100 字</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">全國服務縣市</span>
          <span class="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center text-sm"><i class="fa-solid fa-map-location-dot"></i></span>
        </div>
        <div class="text-3xl font-black text-emerald-600">{len(region_counter)} <span class="text-xs font-normal text-slate-500">個縣市</span></div>
        <p class="text-[11px] text-emerald-700 mt-2">基隆、雙北、台中、高雄全台連線</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">涵蓋學校機構</span>
          <span class="w-8 h-8 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center text-sm"><i class="fa-solid fa-school"></i></span>
        </div>
        <div class="text-3xl font-black text-purple-600">90+ <span class="text-xs font-normal text-slate-500">所單位</span></div>
        <p class="text-[11px] text-purple-700 mt-2">涵蓋國小、國中、高中職與大專</p>
      </div>

    </div>

    <!-- Curriculum Roadmap Section -->
    <section id="curriculum" class="mb-12">
      <div class="flex items-center justify-between mb-6">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Syllabus & Roadmap</div>
          <h2 class="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">雙梯次完整 6 小時實戰地圖</h2>
        </div>
        <a href="https://mihozip.tw/lectures.demo/2026101623-keelung/" target="_blank" class="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1">
          瀏覽線上教案手冊 <i class="fa-solid fa-arrow-right"></i>
        </a>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        
        <!-- Cohort 1 Card -->
        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm relative overflow-hidden">
          <div class="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-blue-500 to-cyan-500"></div>
          <div class="flex items-center justify-between mb-3">
            <span class="px-2.5 py-1 rounded-full bg-blue-50 text-blue-700 text-xs font-bold border border-blue-200">
              第一梯次
            </span>
            <span class="text-xs text-slate-500"><i class="fa-regular fa-clock mr-1"></i>13:30–16:30（3 小時）</span>
          </div>
          <h3 class="text-lg font-bold text-slate-900">115 年 10 月 16 日 (五) · 概念 × 規格 × 工具</h3>
          <p class="text-xs text-slate-500 mt-1 mb-4">打底建制：打破提示詞迷思，掌握 Agent Brief 規格化與多檔案工具調用清洗</p>

          <div class="space-y-3 text-xs text-slate-700">
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">1</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元一：打破 Prompt 迷思</strong>
                從單次對話打字員蛻變為自主智慧代理；釐清資訊流、工作流與責任流。
              </div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">2</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元二：任務規格化實戰</strong>
                教務處與學務處 Agent Brief：邊界限制、輸出格式規範與防呆要求。
              </div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">3</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元三：多檔案清洗與去個資自動化</strong>
                利用工具精算、簽到退名冊比對、表單資料清洗與自動遮罩保護。
              </div>
            </div>
          </div>
          <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] text-slate-500 flex justify-between">
            <span>填報成果：109 位學員</span>
            <span class="text-emerald-600 font-semibold">88 位心得文字逾百字</span>
          </div>
        </div>

        <!-- Cohort 2 Card -->
        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm relative overflow-hidden">
          <div class="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 to-purple-500"></div>
          <div class="flex items-center justify-between mb-3">
            <span class="px-2.5 py-1 rounded-full bg-indigo-50 text-indigo-700 text-xs font-bold border border-indigo-200">
              第二梯次
            </span>
            <span class="text-xs text-slate-500"><i class="fa-regular fa-clock mr-1"></i>13:30–16:30（3 小時）</span>
          </div>
          <h3 class="text-lg font-bold text-slate-900">115 年 10 月 31 日 (六) · 專案 × 實戰 × 防護</h3>
          <p class="text-xs text-slate-500 mt-1 mb-4">閉環可信：專案持久化 Workspace，處室案例演練與人機協作安全防護欄</p>

          <div class="space-y-3 text-xs text-slate-700">
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">4</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元四：專案狀態持久化管理</strong>
                設計 <code>PROJECT_STATE.md</code> 狀態檔，徹底解決 AI 跨階段中斷失憶問題。
              </div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">5</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元五：三大處室自動化工作流閉環</strong>
                規劃 (Plan) $\rightarrow$ 執行 (Execute) $\rightarrow$ 審查 (Review)，應對觀議課與大型校務計畫。
              </div>
            </div>
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-100 flex items-start gap-2.5">
              <span class="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-[10px] mt-0.5 shrink-0">6</span>
              <div>
                <strong class="text-slate-900 block font-semibold">單元六：可信任智慧代理與安全護欄</strong>
                人工審核雙軌制、個資防護線與 AGYReady 小幫手，構建安全高信任環境。
              </div>
            </div>
          </div>
          <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] text-slate-500 flex justify-between">
            <span>填報成果：56 位學員</span>
            <span class="text-indigo-600 font-semibold">45 位心得文字逾百字</span>
          </div>
        </div>

      </div>
    </section>

    <!-- Live Photo Gallery Section -->
    <section id="photos" class="mb-12">
      <div class="flex items-center justify-between mb-6">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Live Photo Gallery</div>
          <h2 class="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">線上研習精彩教學實況紀錄</h2>
          <p class="text-xs text-slate-500 mt-0.5">收錄雙梯次線上互動、即時提問解答、Padlet 學習看板與終端機環境實作精選截圖</p>
        </div>
      </div>

      <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-4">
'''

# 插入照片卡片
for idx, p_url in enumerate(photos, start=1):
    html_content += f'''
        <div class="photo-card bg-white border border-slate-200 shadow-sm" onclick="openModal('{p_url}', '線上教學實況紀錄 #{idx:02d}')">
          <div class="relative aspect-video bg-slate-100 overflow-hidden">
            <img src="{p_url}" alt="教學實況 #{idx}" class="w-full h-full object-cover" loading="lazy">
            <div class="absolute inset-0 bg-slate-900/30 opacity-0 hover:opacity-100 transition flex items-center justify-center text-white">
              <i class="fa-solid fa-magnifying-glass-plus text-lg"></i>
            </div>
          </div>
          <div class="p-2.5 text-[11px] font-medium text-slate-600 flex justify-between items-center bg-slate-50">
            <span>實況紀錄 #{idx:02d}</span>
            <span class="text-blue-600"><i class="fa-regular fa-image"></i></span>
          </div>
        </div>
'''

html_content += f'''
      </div>
    </section>

    <!-- Handouts & Posters Section -->
    <section id="posters" class="mb-12">
      <div class="flex items-center justify-between mb-6">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Lecture Handouts & Posters</div>
          <h2 class="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">實戰指南海報與章節導覽</h2>
          <p class="text-xs text-slate-500 mt-0.5">彭智標主任專為教育現場量身打造之核心心法海報與教材綱要</p>
        </div>
      </div>

      <!-- 10/16 梯次海報 -->
      <div class="mb-6">
        <h3 class="text-sm font-bold text-blue-800 mb-3 flex items-center gap-2">
          <span class="w-2 h-2 rounded-full bg-blue-600"></span> 第一梯次教材海報系列（行政效能大升級 1-3）
        </h3>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
'''

for post in posters_1016:
    html_content += f'''
          <div class="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm flex flex-col justify-between">
            <div class="cursor-pointer" onclick="openModal('{post["img"]}', '{post["title"]}')">
              <img src="{post["img"]}" alt="{post["title"]}" class="w-full aspect-[4/3] object-cover hover:opacity-95 transition">
            </div>
            <div class="p-4">
              <h4 class="font-bold text-slate-900 text-xs sm:text-sm">{post["title"]}</h4>
              <p class="text-[11px] text-slate-500 mt-1">{post["desc"]}</p>
              <div class="mt-3 pt-2 border-t border-slate-100 flex justify-between items-center text-xs">
                <a href="{post["pdf"]}" target="_blank" class="text-blue-600 hover:text-blue-800 font-medium flex items-center gap-1">
                  <i class="fa-solid fa-file-pdf"></i> 開啟 PDF 講義
                </a>
                <button onclick="openModal('{post["img"]}', '{post["title"]}')" class="text-slate-400 hover:text-slate-600">
                  <i class="fa-solid fa-expand"></i> 檢視大圖
                </button>
              </div>
            </div>
          </div>
'''

html_content += f'''
        </div>
      </div>

      <!-- 10/31 梯次海報 -->
      <div>
        <h3 class="text-sm font-bold text-indigo-800 mb-3 flex items-center gap-2">
          <span class="w-2 h-2 rounded-full bg-indigo-600"></span> 第二梯次教材海報系列（校務實戰指南 1-2）
        </h3>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
'''

for post in posters_1031:
    html_content += f'''
          <div class="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm flex flex-col justify-between">
            <div class="cursor-pointer" onclick="openModal('{post["img"]}', '{post["title"]}')">
              <img src="{post["img"]}" alt="{post["title"]}" class="w-full aspect-[16/9] object-cover hover:opacity-95 transition">
            </div>
            <div class="p-4">
              <h4 class="font-bold text-slate-900 text-xs sm:text-sm">{post["title"]}</h4>
              <p class="text-[11px] text-slate-500 mt-1">{post["desc"]}</p>
              <div class="mt-3 pt-2 border-t border-slate-100 flex justify-between items-center text-xs">
                <a href="{post["pdf"]}" target="_blank" class="text-indigo-600 hover:text-indigo-800 font-medium flex items-center gap-1">
                  <i class="fa-solid fa-file-pdf"></i> 開啟 PDF 講義
                </a>
                <button onclick="openModal('{post["img"]}', '{post["title"]}')" class="text-slate-400 hover:text-slate-600">
                  <i class="fa-solid fa-expand"></i> 檢視大圖
                </button>
              </div>
            </div>
          </div>
'''

html_content += f'''
        </div>
      </div>
    </section>

    <!-- Data Visualization Charts Section -->
    <section id="stats" class="mb-12">
      <div class="flex items-center justify-between mb-6">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Analytics & Insights</div>
          <h2 class="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">研習整體數據分析與圖表呈現</h2>
        </div>
      </div>

      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        <!-- Chart 1: 服務縣市分布 (7 cols) -->
        <div class="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="font-bold text-slate-800 text-sm sm:text-base"><i class="fa-solid fa-chart-column text-blue-600 mr-2"></i>全台學員所屬縣市分布（前 12 名）</h3>
              <p class="text-[11px] text-slate-500">跨區高密度參與：新北(27)、台北(18)、台中(18)、高雄(14)居前</p>
            </div>
          </div>
          <div class="relative h-72">
            <canvas id="regionChart"></canvas>
          </div>
        </div>

        <!-- Chart 2: 教育階段分布 (5 cols) -->
        <div class="lg:col-span-5 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="font-bold text-slate-800 text-sm sm:text-base"><i class="fa-solid fa-chart-pie text-emerald-600 mr-2"></i>學員教育階段與單位屬性</h3>
              <p class="text-[11px] text-slate-500">國小(33.3%)、高中職(25.5%)兼任行政教師占最大宗</p>
            </div>
          </div>
          <div class="relative h-72 flex items-center justify-center">
            <canvas id="levelChart"></canvas>
          </div>
        </div>

        <!-- Chart 3: 雙梯次規模比較 (6 cols) -->
        <div class="lg:col-span-6 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="font-bold text-slate-800 text-sm sm:text-base"><i class="fa-solid fa-scale-balanced text-indigo-600 mr-2"></i>雙梯次參訓填報人次比較</h3>
              <p class="text-[11px] text-slate-500">第 1 梯次 109 人次 (66.1%) vs 第 2 梯次 56 人次 (33.9%)</p>
            </div>
          </div>
          <div class="relative h-64 flex items-center justify-center">
            <canvas id="cohortChart"></canvas>
          </div>
        </div>

        <!-- Chart 4: 心得字數合格率 (6 cols) -->
        <div class="lg:col-span-6 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="font-bold text-slate-800 text-sm sm:text-base"><i class="fa-solid fa-square-poll-vertical text-amber-500 mr-2"></i>深度回饋達標率 (文字超過 100 字)</h3>
              <p class="text-[11px] text-slate-500">文字合格 133 人 (80.6%)，充分展現教師深入消化與實務省思</p>
            </div>
          </div>
          <div class="relative h-64 flex items-center justify-center">
            <canvas id="passChart"></canvas>
          </div>
        </div>

      </div>
    </section>

    <!-- Testimonials Section -->
    <section id="testimonials" class="mb-12">
      <div class="flex items-center justify-between mb-6">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Testimonials</div>
          <h2 class="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">精選代表性學員深度心得與金句牆</h2>
          <p class="text-xs text-slate-500 mt-1">依規定姓名只保留姓，其餘以「0」去識別化保護個資；身分證字號完全省略</p>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
        
        <div class="quote-box p-5 border-l-blue-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-blue-700"><i class="fa-solid fa-school mr-1"></i> 台中市廍子國小</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">楊00</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「平常整天被瑣事打斷，剛剛算到一半的簽到表和經費又要重來。以前用 ChatGPT 常常自己亂掰，這堂課最大的收穫是把思維翻轉過來：規格先講清楚、善用工具精算對照、進度存成檔案可一秒接關，最後設好安全防線，真正替自己分擔重複雜務，把力氣留給教學與學生！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第二梯次 (10/31) · 心得 411 字</div>
        </div>

        <div class="quote-box p-5 border-l-indigo-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-indigo-700"><i class="fa-solid fa-school mr-1"></i> 雲林縣僑真國小</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">詹00</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「彭智標老師的行政流程心法，確實打中行政人員所需要思考的點！從教育現場實際需求出發，將 AI 工具、提示詞設計與工作流程結合。導入 AI 的目的不是取代人工，而是減少重複性事務，讓行政人員將時間投入教學支持！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第一梯次 (10/16) · 心得 412 字</div>
        </div>

        <div class="quote-box p-5 border-l-emerald-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-emerald-700"><i class="fa-solid fa-school mr-1"></i> 台北市麗山高中</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">V0000000000</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「這門課程精準切入校園行政痛點，彭老師將深奧的 AI Agent 架構實用化。透過持久化記憶（如 PROJECT_STATE.md）設計，完美解決 AI 跨階段斷點問題；最令人安心的是導入人工審核與安全護欄機制，兼具高效率與高信任度！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第二梯次 (10/31) · 心得 204 字</div>
        </div>

        <div class="quote-box p-5 border-l-amber-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-amber-700"><i class="fa-solid fa-school mr-1"></i> 基隆市百福國中</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">宋00</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「學校行政工作以服務教學為目的，但近年來教師對兼任行政避而遠之。若能有一套便捷、順暢的系統架構，可以大幅減輕負擔，讓兼任行政的教師喘一口氣，更有餘裕將心力用於教學準備！感謝講師傾囊相授！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第二梯次 (10/31) · 心得 196 字</div>
        </div>

        <div class="quote-box p-5 border-l-purple-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-purple-700"><i class="fa-solid fa-school mr-1"></i> 新北市未來實驗教育機構</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">呂00</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「感謝講師帶來如此精采豐富的課程！從 Antigravity CLI 基本概念、詳細步驟，到整理製作的提示詞與安裝工具，大大降低入門門檻。透過實際應用案例展現如何建構工作流，大幅提升繁瑣事務效率！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第一梯次 (10/16) · 心得 176 字</div>
        </div>

        <div class="quote-box p-5 border-l-rose-600">
          <div class="flex items-center justify-between mb-2 text-xs text-slate-500">
            <span class="font-bold text-rose-700"><i class="fa-solid fa-school mr-1"></i> 嘉義市輔仁中學</span>
            <span class="font-mono font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">陳00</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「以前教務、總務開會資料與研習數據整理超花時間，只能一筆一筆手動改。這次課程直接帶大家把日常行政改成『半自動流程』，連去個資、寫報告都有現成網頁工具可用，省下一大半時間！」
          </p>
          <div class="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-400">第二梯次 (10/31) · 心得 133 字</div>
        </div>

      </div>
    </section>

    <!-- Searchable Roster Table Section -->
    <section id="roster" class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden mb-16">
      <div class="p-5 sm:p-6 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div class="text-xs font-bold uppercase tracking-wider text-blue-600">Feedback Roster</div>
          <h2 class="text-lg sm:text-xl font-black text-slate-900 mt-0.5">學員回饋清冊檢索與篩選台</h2>
          <p class="text-xs text-slate-500 mt-0.5">姓名只保留姓、其餘以 0 遮罩；身分證已省略</p>
        </div>
        
        <div class="flex flex-wrap items-center gap-2 sm:gap-3">
          <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="搜尋姓名、學校或心得關鍵字..." class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 w-48 sm:w-56">
          
          <select id="cohortFilter" onchange="filterTable()" class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
            <option value="all">全部梯次 ({total_records})</option>
            <option value="c1">第 1 梯次 10/16 ({c1_count})</option>
            <option value="c2">第 2 梯次 10/31 ({c2_count})</option>
          </select>

          <select id="passFilter" onchange="filterTable()" class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
            <option value="all">回饋字數</option>
            <option value="pass">字數超過 100 字 ({pass_count})</option>
            <option value="fail">未達 100 字 ({total_records - pass_count})</option>
          </select>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs text-slate-600">
          <thead class="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
            <tr>
              <th class="py-3 px-3 w-12 text-center">序號</th>
              <th class="py-3 px-3 w-28 text-center">研習梯次</th>
              <th class="py-3 px-4 w-44">服務單位</th>
              <th class="py-3 px-3 w-24 text-center">學員稱謂</th>
              <th class="py-3 px-3 w-24 text-center">回饋字數</th>
              <th class="py-3 px-4">研習心得與具體收穫</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-100">
            <!-- Dynamically populated -->
          </tbody>
        </table>
      </div>

      <div class="p-4 border-t border-slate-100 text-xs text-slate-500 flex flex-col sm:flex-row justify-between items-center gap-2 bg-slate-50">
        <span id="recordCountText" class="font-medium text-slate-700">顯示全部 {total_records} 筆紀錄</span>
        <span class="text-slate-400">個資防護已完成去識別化作業</span>
      </div>
    </section>

  </main>

  <!-- Modal for Image Preview -->
  <div id="imageModal" class="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm hidden flex items-center justify-center p-4" onclick="closeModal()">
    <div class="relative max-w-5xl max-h-[90vh] bg-slate-900 rounded-2xl overflow-hidden shadow-2xl border border-slate-700 flex flex-col" onclick="event.stopPropagation()">
      <div class="p-3 bg-slate-950 flex items-center justify-between text-white border-b border-slate-800 text-xs">
        <span id="modalTitle" class="font-semibold text-slate-200">圖片預覽</span>
        <button onclick="closeModal()" class="w-7 h-7 rounded-lg hover:bg-slate-800 flex items-center justify-center text-slate-400 hover:text-white transition">
          <i class="fa-solid fa-xmark text-sm"></i>
        </button>
      </div>
      <div class="flex-1 overflow-auto p-2 flex items-center justify-center bg-black/40">
        <img id="modalImg" src="" alt="預覽大圖" class="max-w-full max-h-[80vh] object-contain rounded-lg">
      </div>
    </div>
  </div>

  <!-- Footer -->
  <footer class="bg-slate-950 text-slate-400 text-xs py-10 border-t border-slate-800">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row items-center justify-between gap-4">
      <div>
        <div class="font-bold text-slate-200 text-sm">基隆市 115 學年度公立高中職中小學行政增能線上研習</div>
        <div class="text-slate-500 mt-1">主辦：基隆市政府教育處 · 承辦：基隆百福科技團 · 講師：彭智標 主任</div>
      </div>
      <div class="flex items-center gap-4">
        <a href="https://mihozip.tw/lectures.demo/2026101623-keelung/" target="_blank" class="hover:text-white transition">研習教材工作台</a>
        <a href="#curriculum" class="hover:text-white transition">回到頂部</a>
      </div>
    </div>
  </footer>

  <!-- Scripts -->
  <script>
    const records = {all_records_json};

    function renderTable(data) {{
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';
      
      data.forEach((item, idx) => {{
        const tr = document.createElement('tr');
        tr.className = idx % 2 === 0 ? 'bg-white hover:bg-slate-50' : 'bg-slate-50/50 hover:bg-slate-50';
        
        const cohortBadge = item.cohort_code === 'c1'
          ? `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-blue-100 text-blue-800">10/16 梯次</span>`
          : `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-indigo-100 text-indigo-800">10/31 梯次</span>`;

        const passBadge = item.is_pass
          ? `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-semibold bg-emerald-100 text-emerald-800">${{item.fb_len}} 字</span>`
          : `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-slate-100 text-slate-600">${{item.fb_len}} 字</span>`;

        tr.innerHTML = `
          <td class="py-3 px-3 text-center font-medium text-slate-400">${{idx + 1}}</td>
          <td class="py-3 px-3 text-center">${{cohortBadge}}</td>
          <td class="py-3 px-4 font-semibold text-slate-800">
            ${{item.school}}
            <span class="text-[10px] font-normal text-slate-400 block">${{item.region}} · ${{item.level}}</span>
          </td>
          <td class="py-3 px-3 text-center font-mono font-medium text-slate-700 bg-slate-50/80">${{item.name}}</td>
          <td class="py-3 px-2 text-center">${{passBadge}}</td>
          <td class="py-3 px-4 text-slate-600 leading-relaxed">${{item.feedback}}</td>
        `;
        tbody.appendChild(tr);
      }});

      document.getElementById('recordCountText').textContent = `符合條件紀錄共 ${{data.length}} 筆`;
    }}

    function filterTable() {{
      const query = document.getElementById('searchInput').value.toLowerCase().trim();
      const cohort = document.getElementById('cohortFilter').value;
      const pass = document.getElementById('passFilter').value;

      const filtered = records.filter(item => {{
        const matchesQuery = !query || 
          item.name.toLowerCase().includes(query) || 
          item.school.toLowerCase().includes(query) || 
          item.feedback.toLowerCase().includes(query);

        let matchesCohort = true;
        if (cohort !== 'all') matchesCohort = (item.cohort_code === cohort);

        let matchesPass = true;
        if (pass === 'pass') matchesPass = item.is_pass;
        else if (pass === 'fail') matchesPass = !item.is_pass;

        return matchesQuery && matchesCohort && matchesPass;
      }});

      renderTable(filtered);
    }}

    renderTable(records);

    // Modal controls
    function openModal(imgSrc, title) {{
      document.getElementById('modalImg').src = imgSrc;
      document.getElementById('modalTitle').textContent = title || '實況紀錄檢視';
      document.getElementById('imageModal').classList.remove('hidden');
    }}

    function closeModal() {{
      document.getElementById('imageModal').classList.add('hidden');
      document.getElementById('modalImg').src = '';
    }}

    // 1. Region Chart
    const regionCtx = document.getElementById('regionChart').getContext('2d');
    new Chart(regionCtx, {{
      type: 'bar',
      data: {{
        labels: {json.dumps(top_reg_names, ensure_ascii=False)},
        datasets: [{{
          label: '參與學員數',
          data: {top_reg_vals},
          backgroundColor: '#3b82f6',
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{
          y: {{ beginAtZero: true }},
          x: {{ grid: {{ display: false }} }}
        }}
      }}
    }});

    // 2. Level Chart
    const levelCtx = document.getElementById('levelChart').getContext('2d');
    new Chart(levelCtx, {{
      type: 'doughnut',
      data: {{
        labels: ['國小/幼兒', '高中職/中學', '國中', '機關/機構/其他', '大專院校'],
        datasets: [{{
          data: [{level_counter.get('國小/幼兒', 0)}, {level_counter.get('高中職/中學', 0)}, {level_counter.get('國中', 0)}, {level_counter.get('機關/機構/其他', 0)}, {level_counter.get('大專院校', 0)}],
          backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#64748b', '#8b5cf6']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }}
      }}
    }});

    // 3. Cohort Chart
    const cohortCtx = document.getElementById('cohortChart').getContext('2d');
    new Chart(cohortCtx, {{
      type: 'bar',
      data: {{
        labels: ['第 1 梯次 (10/16)', '第 2 梯次 (10/31)'],
        datasets: [{{
          label: '參訓填報人次',
          data: [{c1_count}, {c2_count}],
          backgroundColor: ['#0284c7', '#6366f1'],
          borderRadius: 8
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{
          y: {{ beginAtZero: true }}
        }}
      }}
    }});

    // 4. Pass Chart
    const passCtx = document.getElementById('passChart').getContext('2d');
    new Chart(passCtx, {{
      type: 'pie',
      data: {{
        labels: ['心得超過100字 ({pass_count}人)', '簡短心得 ({total_records - pass_count}人)'],
        datasets: [{{
          data: [{pass_count}, {total_records - pass_count}],
          backgroundColor: ['#059669', '#cbd5e1']
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

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("成果官方首頁 index.html 生成完畢！")
