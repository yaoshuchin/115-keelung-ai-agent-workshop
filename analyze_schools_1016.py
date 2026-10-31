# -*- coding: utf-8 -*-
import sys, io, openpyxl
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('1151016基隆百福科技團線上研習.xlsx', data_only=True)
ws = wb['表單回應 1']

schools = []
for r in range(3, ws.max_row+1):
    vals = [ws.cell(r, c).value for c in range(1, 8)]
    if any(v is not None for v in vals):
        schools.append(str(vals[1]).strip() if vals[1] else '')

print(f'Total schools: {len(schools)}')

def get_region(s):
    if '基隆' in s: return '基隆市'
    if '台北' in s or '臺北' in s: return '台北市'
    if '新北' in s: return '新北市'
    if '桃園' in s: return '桃園市'
    if '新竹' in s: return '新竹縣市'
    if '苗栗' in s: return '苗栗縣'
    if '台中' in s or '臺中' in s: return '台中市'
    if '彰化' in s: return '彰化縣'
    if '南投' in s: return '南投縣'
    if '雲林' in s: return '雲林縣'
    if '嘉義' in s: return '嘉義縣市'
    if '台南' in s or '臺南' in s: return '台南市'
    if '高雄' in s: return '高雄市'
    if '屏東' in s: return '屏東縣'
    if '宜蘭' in s: return '宜蘭縣'
    if '花蓮' in s: return '花蓮縣'
    if '台東' in s or '臺東' in s: return '台東縣'
    if '金門' in s: return '金門縣'
    if '澎湖' in s: return '澎湖縣'
    return '其他'

def get_level(s):
    if any(k in s for k in ['國小', '小學', '幼兒園']): return '國小/幼兒'
    if '國中' in s: return '國中'
    if any(k in s for k in ['高中', '高職', '商工', '農工', '工職', '附中', '陽明', '中學', '實驗教育']) and '國中' not in s: return '高中職/中學'
    if any(k in s for k in ['大學', '科技大學', '科大']): return '大專院校'
    return '機關/機構/其他'

reg_cnt = Counter([get_region(s) for s in schools])
lev_cnt = Counter([get_level(s) for s in schools])

print('Regions:', reg_cnt.most_common())
print('Levels:', lev_cnt.most_common())

other_schools = [s for s in schools if get_region(s) == '其他']
print('Other schools:', other_schools)
