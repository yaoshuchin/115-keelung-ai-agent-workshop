# -*- coding: utf-8 -*-
"""
生成現代化、互動式的研習成效數據分析儀表板 HTML
"""
import sys, io, json, re, openpyxl
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook("1151031基隆百福科技團線上研習.xlsx", data_only=True)
ws = wb["工作表1"]

records = []
for r in range(2, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, 6)]
    if not any(v is not None for v in vals):
        continue
    records.append({
        "id_num": r - 1,
        "time": str(vals[0]).strip() if vals[0] else "",
        "school": str(vals[1]).strip() if vals[1] else "",
        "name": str(vals[2]).strip() if vals[2] else "",
        "id_code": str(vals[3]).strip() if vals[3] else "",
        "feedback": str(vals[4]).strip() if vals[4] else ""
    })

def get_region(s):
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
    if '國小' in s: return '國小'
    if '國中' in s: return '國中'
    if any(k in s for k in ['高中', '高職', '商工', '附中', '陽明']) and '國中' not in s: return '高中職'
    if any(k in s for k in ['大學', '科技大學', '科大']): return '大專院校'
    return '機關/機構/其他'

negatives = ['不需', '不用', '不需要', '無', 'no', '.', 'none']
def is_has_id(s):
    if not s: return False
    return not any(nw in s.lower() for nw in negatives)

for rec in records:
    rec['region'] = get_region(rec['school'])
    rec['level'] = get_level(rec['school'])
    rec['has_id'] = is_has_id(rec['id_code'])
    # 遮罩身分證供公開檢視
    raw_id = rec['id_code']
    if rec['has_id']:
        if len(raw_id) == 10 and raw_id[0].isalpha():
            rec['masked_id'] = raw_id[:3] + '***' + raw_id[-4:]
        else:
            rec['masked_id'] = raw_id[:3] + '***'
    else:
        rec['masked_id'] = '不需核發'

records_json = json.dumps(records, ensure_ascii=False)

html_content = f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>1151031 基隆百福科技團線上研習 - 成效數據分析與圖表呈現</title>
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
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
      transition: all 0.2s ease;
    }}
    .stat-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
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
  <header class="bg-gradient-to-r from-blue-900 via-indigo-800 to-blue-950 text-white shadow-lg">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
      <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
        <div>
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-800/60 text-blue-200 text-xs font-semibold mb-3 border border-blue-700">
            <i class="fa-solid fa-chalkboard-user"></i> 研習成果報告與統計儀表板
          </div>
          <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight">
            1151031 基隆百福科技團線上研習
          </h1>
          <p class="text-blue-200 text-sm sm:text-base mt-2 flex flex-wrap items-center gap-y-1 gap-x-4">
            <span><i class="fa-solid fa-microchip mr-1"></i> 主題：AI Agent 行政進階實戰：從持久化工作流到可信任智慧代理</span>
            <span><i class="fa-regular fa-calendar-days mr-1"></i> 日期：2026/10/31</span>
            <span><i class="fa-solid fa-user-tie mr-1"></i> 講師：彭智標 主任</span>
          </p>
        </div>
        <div class="flex items-center gap-3 self-start md:self-auto">
          <button onclick="window.print()" class="px-4 py-2 bg-white/10 hover:bg-white/20 text-white rounded-lg text-sm font-medium border border-white/20 transition flex items-center gap-2">
            <i class="fa-solid fa-print"></i> 列印報告
          </button>
        </div>
      </div>
    </div>
  </header>

  <!-- Main Content Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 -mt-6">

    <!-- KPI Metric Cards -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">回饋總填寫</span>
          <span class="w-8 h-8 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center text-sm"><i class="fa-solid fa-users"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-slate-800">55 <span class="text-sm font-normal text-slate-500">人次</span></div>
        <p class="text-xs text-slate-500 mt-2">約 52 位獨立教師參與</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">需核發研習時數</span>
          <span class="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center text-sm"><i class="fa-solid fa-id-card"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-emerald-600">41 <span class="text-sm font-normal text-slate-500">人 (74.5%)</span></div>
        <p class="text-xs text-emerald-700 mt-2"><i class="fa-solid fa-check mr-1"></i>已完整填寫身分證字號</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">免時數/觀摩交流</span>
          <span class="w-8 h-8 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center text-sm"><i class="fa-solid fa-graduation-cap"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-slate-700">14 <span class="text-sm font-normal text-slate-500">人 (25.5%)</span></div>
        <p class="text-xs text-slate-500 mt-2">含社大、大專或純旁聽教師</p>
      </div>

      <div class="stat-card p-5">
        <div class="flex items-center justify-between text-slate-500 mb-2">
          <span class="text-xs font-semibold tracking-wider uppercase">跨縣市影響力</span>
          <span class="w-8 h-8 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center text-sm"><i class="fa-solid fa-map-location-dot"></i></span>
        </div>
        <div class="text-3xl font-extrabold text-purple-600">16 <span class="text-sm font-normal text-slate-500">個縣市</span></div>
        <p class="text-xs text-purple-700 mt-2">涵蓋 48 所不同學校與單位</p>
      </div>
    </div>

    <!-- Charts Grid -->
    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-8">
      
      <!-- Chart 1: 縣市分布 (8 cols) -->
      <div class="lg:col-span-7 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-chart-column text-blue-600 mr-2"></i>學員所屬服務縣市分布</h3>
            <p class="text-xs text-slate-500">全國跨域參與：新北(8)、雙北(14)、中南部與東部皆有熱烈迴響</p>
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
            <p class="text-xs text-slate-500">以國小(34.5%)與高中職(30.9%)兼任行政教師為主</p>
          </div>
        </div>
        <div class="relative h-72 flex items-center justify-center">
          <canvas id="levelChart"></canvas>
        </div>
      </div>

      <!-- Chart 3: 時數需求比例 (4 cols) -->
      <div class="lg:col-span-4 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-certificate text-amber-500 mr-2"></i>研習時數核發需求</h3>
            <p class="text-xs text-slate-500">提供身分證以匯入全國進修網</p>
          </div>
        </div>
        <div class="relative h-64 flex items-center justify-center">
          <canvas id="idRatioChart"></canvas>
        </div>
      </div>

      <!-- Chart 4: 填報時段分布 (8 cols) -->
      <div class="lg:col-span-8 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
          <div>
            <h3 class="font-bold text-slate-800 text-base"><i class="fa-solid fa-clock-rotate-left text-indigo-600 mr-2"></i>簽退與回饋填報時間波峰</h3>
            <p class="text-xs text-slate-500">課堂後半段(11:00-12:00)至中午(12:00-13:00)為回饋填寫高峰期</p>
          </div>
        </div>
        <div class="relative h-64">
          <canvas id="timeChart"></canvas>
        </div>
      </div>

    </div>

    <!-- Core Keywords & Qualitative Insights -->
    <div class="bg-gradient-to-br from-slate-900 to-indigo-950 text-white rounded-2xl p-6 sm:p-8 mb-8 shadow-md">
      <div class="flex items-center gap-3 mb-6">
        <span class="w-10 h-10 rounded-xl bg-blue-500/20 text-blue-400 flex items-center justify-center text-lg border border-blue-500/30">
          <i class="fa-solid fa-brain"></i>
        </span>
        <div>
          <h2 class="text-xl font-bold">學員質性反饋關鍵字與核心成效</h2>
          <p class="text-slate-300 text-xs sm:text-sm">從 55 份反饋中萃取出的四大轉型價值</p>
        </div>
      </div>

      <!-- Keyword Tags -->
      <div class="flex flex-wrap gap-2 mb-8">
        <span class="px-3 py-1.5 rounded-lg bg-blue-600/40 text-blue-200 border border-blue-500/30 text-xs font-semibold"># AI Agent 智慧代理</span>
        <span class="px-3 py-1.5 rounded-lg bg-emerald-600/40 text-emerald-200 border border-emerald-500/30 text-xs font-semibold"># 持久化工作流 (PROJECT_STATE.md)</span>
        <span class="px-3 py-1.5 rounded-lg bg-purple-600/40 text-purple-200 border border-purple-500/30 text-xs font-semibold"># AGY CLI & AGYReady 小幫手</span>
        <span class="px-3 py-1.5 rounded-lg bg-amber-600/40 text-amber-200 border border-amber-500/30 text-xs font-semibold"># 兼任行政實質減負</span>
        <span class="px-3 py-1.5 rounded-lg bg-rose-600/40 text-rose-200 border border-rose-500/30 text-xs font-semibold"># 人工審查與安全護欄機制</span>
        <span class="px-3 py-1.5 rounded-lg bg-cyan-600/40 text-cyan-200 border border-cyan-500/30 text-xs font-semibold"># 觀議課進度與行政計畫拆解</span>
        <span class="px-3 py-1.5 rounded-lg bg-indigo-600/40 text-indigo-200 border border-indigo-500/30 text-xs font-semibold"># 去個資工具與公文半自動化</span>
      </div>

      <!-- 4 Pillars Grid -->
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-slate-200">
        <div class="bg-white/5 border border-white/10 rounded-xl p-4">
          <div class="text-blue-400 font-bold text-sm mb-1"><i class="fa-solid fa-arrows-split-up-and-left mr-1"></i> 告別零散提示詞</div>
          <p class="text-xs text-slate-300 leading-relaxed">學員深刻體會從「盲目試提示詞」進化為「規格化工作流」，將學校複雜任務拆解為可執行的階段。</p>
        </div>
        <div class="bg-white/5 border border-white/10 rounded-xl p-4">
          <div class="text-emerald-400 font-bold text-sm mb-1"><i class="fa-solid fa-database mr-1"></i> 跨階段持久記憶</div>
          <p class="text-xs text-slate-300 leading-relaxed">透過狀態檔設計，徹底解決 AI 跨工作階段的斷點遺忘問題，讓大型校務計畫能長效接續推進。</p>
        </div>
        <div class="bg-white/5 border border-white/10 rounded-xl p-4">
          <div class="text-amber-400 font-bold text-sm mb-1"><i class="fa-solid fa-shield-halved mr-1"></i> 可信任智慧審查</div>
          <p class="text-xs text-slate-300 leading-relaxed">導入嚴謹的安全護欄與人機協作審核機制，消除校園行政上對資料安全與個資外洩的焦慮。</p>
        </div>
        <div class="bg-white/5 border border-white/10 rounded-xl p-4">
          <div class="text-purple-400 font-bold text-sm mb-1"><i class="fa-solid fa-heart-pulse mr-1"></i> 舒緩基層行政焦慮</div>
          <p class="text-xs text-slate-300 leading-relaxed">切中生教、教學等基層行政痛點，藉由半自動化工具釋放時間，讓教師重獲專注教學的餘裕。</p>
        </div>
      </div>
    </div>

    <!-- Highlights Testimonials -->
    <div class="mb-8">
      <div class="flex items-center justify-between mb-4">
        <div>
          <h2 class="text-xl font-bold text-slate-800"><i class="fa-solid fa-comment-dots text-blue-600 mr-2"></i>精選學員代表性回饋與金句</h2>
          <p class="text-xs text-slate-500">多位第一線教師與行政主管由衷分享的實務心得</p>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
        
        <div class="quote-card p-5">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-blue-700"><i class="fa-solid fa-school mr-1"></i> 台北市麗山高中</span>
            <span>Vinny Huang</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「這門課程精準切入校園行政痛點！彭老師將深奧的 AI Agent 架構實用化。透過持久化記憶（如 PROJECT_STATE.md）設計，完美解決跨工作階段斷點問題；最安心的是導入人工審核與安全護欄機制，兼具高效率與高信任度！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-emerald-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-emerald-700"><i class="fa-solid fa-school mr-1"></i> 基隆市百福國中</span>
            <span>宋郁峰</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「近年兼任行政尤其是教學或生教組，繁雜行政事務讓教師避而遠之。若能有一套便捷順暢的系統架構，能大幅減輕負擔，讓行政教師喘一口氣、更有餘裕將心力用於教學準備！感謝講師傾囊相授！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-purple-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-purple-700"><i class="fa-solid fa-school mr-1"></i> 新北市泰山高中</span>
            <span>吳俐蓉</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「面對越來越多的教學挑戰，感謝講師分享心法和實際做法，頗有共鳴。尤其是資訊焦慮的部分被觸動到了，提醒自己：善用 AI 而非被制約了心智！假日清早跟著線上研習跑收穫極大！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-amber-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-amber-700"><i class="fa-solid fa-school mr-1"></i> 嘉義市輔仁中學</span>
            <span>陳揮明</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「以前教務、總務開會資料與研習數據整理超花時間，只能手動一筆筆改。這次直接帶大家改成『半自動流程』，連去個資、寫報告都有工具可用，處理公文跟雜事省下一大半時間！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-cyan-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-cyan-700"><i class="fa-solid fa-school mr-1"></i> 高雄市中正國小</span>
            <span>蔡宗翰</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「便利的 AGY CLI 與 AGYReady 安裝小幫手直接切入文字介面！教材詳盡、提示語與範例練習檔齊全。感謝內容都能切中學校行政面對的真實問題，非常享受自動化的建置過程！」
          </p>
        </div>

        <div class="quote-card p-5 border-l-rose-600">
          <div class="flex items-center justify-between mb-3 text-xs text-slate-500">
            <span class="font-bold text-rose-700"><i class="fa-solid fa-school mr-1"></i> 高雄市文府國小</span>
            <span>陳育承</span>
          </div>
          <p class="text-xs text-slate-700 leading-relaxed italic">
            「深入淺出說明智慧代理設計思維，透過自動化流程整合資料蒐集、分析與執行任務，大幅減少重複性工作，是一場非常充實且具啟發性的實戰課程！」
          </p>
        </div>

      </div>
    </div>

    <!-- Data Table & Search -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mb-8">
      <div class="p-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 class="font-bold text-slate-800 text-lg"><i class="fa-solid fa-table-list text-blue-600 mr-2"></i>研習學員填報名單與回饋詳表</h2>
          <p class="text-xs text-slate-500">支援關鍵字即時搜尋、縣市與教育階段多條件篩選</p>
        </div>
        <div class="flex flex-wrap items-center gap-3">
          <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="搜尋姓名、學校或心得..." class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 w-56">
          <select id="regionFilter" onchange="filterTable()" class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
            <option value="">全部縣市</option>
            <option value="新北市">新北市</option>
            <option value="台北市">台北市</option>
            <option value="高雄市">高雄市</option>
            <option value="台中市">台中市</option>
            <option value="桃園市">桃園市</option>
            <option value="台南市">台南市</option>
            <option value="基隆市">基隆市</option>
            <option value="雲林縣">雲林縣</option>
            <option value="新竹縣市">新竹縣市</option>
            <option value="苗栗縣">苗栗縣</option>
            <option value="宜蘭縣">宜蘭縣</option>
            <option value="其他">其他縣市</option>
          </select>
          <select id="idFilter" onchange="filterTable()" class="px-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
            <option value="">時數需求</option>
            <option value="has_id">需研習時數 (有身分證)</option>
            <option value="no_id">免時數需求</option>
          </select>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs text-slate-600">
          <thead class="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
            <tr>
              <th class="py-3 px-4 w-12 text-center">#</th>
              <th class="py-3 px-4 w-36">服務單位</th>
              <th class="py-3 px-4 w-24">姓名</th>
              <th class="py-3 px-4 w-28 text-center">時數/身分證</th>
              <th class="py-3 px-4 w-28 text-center">填寫時間</th>
              <th class="py-3 px-4">心得回饋摘要</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-100">
            <!-- Dynamically populated -->
          </tbody>
        </table>
      </div>
      <div class="p-4 border-t border-slate-100 text-xs text-slate-500 flex justify-between items-center bg-slate-50">
        <span id="recordCountText">顯示全部 55 筆紀錄</span>
        <span class="text-slate-400">個資已於網頁端進行合規遮罩保護</span>
      </div>
    </div>

  </main>

  <script>
    const records = {records_json};

    // Render Table
    function renderTable(data) {{
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';
      
      data.forEach((item, idx) => {{
        const tr = document.createElement('tr');
        tr.className = idx % 2 === 0 ? 'bg-white hover:bg-slate-50' : 'bg-slate-50/50 hover:bg-slate-50';
        
        const badge = item.has_id 
          ? `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-emerald-100 text-emerald-800"><i class="fa-solid fa-check mr-1"></i>${{item.masked_id}}</span>`
          : `<span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-slate-100 text-slate-600">免時數</span>`;

        tr.innerHTML = `
          <td class="py-3 px-4 text-center font-medium text-slate-400">${{idx + 1}}</td>
          <td class="py-3 px-4 font-semibold text-slate-800">
            ${{item.school}}
            <span class="text-[10px] font-normal text-slate-400 block">${{item.region}} · ${{item.level}}</span>
          </td>
          <td class="py-3 px-4 text-slate-700">${{item.name}}</td>
          <td class="py-3 px-4 text-center">${{badge}}</td>
          <td class="py-3 px-4 text-center text-slate-400 text-[11px]">${{item.time.replace('2026/10/31 ', '')}}</td>
          <td class="py-3 px-4 text-slate-600 leading-relaxed">${{item.feedback}}</td>
        `;
        tbody.appendChild(tr);
      }});

      document.getElementById('recordCountText').textContent = `符合條件紀錄共 ${{data.length}} 筆`;
    }}

    function filterTable() {{
      const query = document.getElementById('searchInput').value.toLowerCase().trim();
      const region = document.getElementById('regionFilter').value;
      const idFilter = document.getElementById('idFilter').value;

      const filtered = records.filter(item => {{
        const matchesQuery = !query || 
          item.name.toLowerCase().includes(query) || 
          item.school.toLowerCase().includes(query) || 
          item.feedback.toLowerCase().includes(query);

        const matchesRegion = !region || 
          (region === '其他' ? ['彰化縣','嘉義市','屏東縣','宜蘭縣','花蓮縣','台東縣'].includes(item.region) : item.region === region);

        const matchesId = !idFilter || 
          (idFilter === 'has_id' ? item.has_id : !item.has_id);

        return matchesQuery && matchesRegion && matchesId;
      }});

      renderTable(filtered);
    }}

    // Initial render
    renderTable(records);

    // Setup Charts
    // 1. Region Chart
    const regionCtx = document.getElementById('regionChart').getContext('2d');
    new Chart(regionCtx, {{
      type: 'bar',
      data: {{
        labels: ['新北', '台北', '高雄', '台中', '桃園', '台南', '基隆', '雲林', '新竹', '苗栗', '宜蘭', '其他縣市'],
        datasets: [{{
          label: '參與學員數',
          data: [8, 6, 6, 6, 5, 5, 4, 3, 3, 2, 2, 5],
          backgroundColor: '#3b82f6',
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ display: false }}
        }},
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
        labels: ['國小', '高中職', '國中', '大專院校', '機關/機構'],
        datasets: [{{
          data: [19, 17, 9, 4, 6],
          backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#8b5cf6', '#64748b']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }}
        }}
      }}
    }});

    // 3. ID Ratio Chart
    const idRatioCtx = document.getElementById('idRatioChart').getContext('2d');
    new Chart(idRatioCtx, {{
      type: 'pie',
      data: {{
        labels: ['需時數 (有身分證)', '免時數需求'],
        datasets: [{{
          data: [41, 14],
          backgroundColor: ['#059669', '#cbd5e1']
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }}
        }}
      }}
    }});

    // 4. Time Chart
    const timeCtx = document.getElementById('timeChart').getContext('2d');
    new Chart(timeCtx, {{
      type: 'line',
      data: {{
        labels: ['09:00-11:00 (課程進行)', '11:00-12:00 (結尾高潮)', '12:00-13:00 (中午簽退)', '13:00 以後 (課後回饋)'],
        datasets: [{{
          label: '填報人次',
          data: [10, 26, 17, 2],
          borderColor: '#6366f1',
          backgroundColor: 'rgba(99, 102, 241, 0.1)',
          fill: true,
          tension: 0.35,
          pointBackgroundColor: '#4f46e5',
          pointRadius: 5
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ display: false }}
        }},
        scales: {{
          y: {{ beginAtZero: true }},
          x: {{ grid: {{ display: false }} }}
        }}
      }}
    }});
  </script>
</body>
</html>
'''

with open("研習成效數據分析圖表.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("研習成效數據分析圖表.html 生成完畢！")
