# -*- coding: utf-8 -*-
import sys, io, openpyxl, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('1151016基隆百福科技團線上研習.xlsx', data_only=True)
ws = wb['表單回應 1']

id_regex = re.compile(r'^[A-Za-z][1289]\d{8}$')

valid_ids = []
invalid_ids = []
feedback_over_100 = []

for r in range(2, ws.max_row+1):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column+1)]
    if not any(v is not None for v in vals):
        continue
    id_val = str(vals[3]).strip() if vals[3] is not None else ''
    fb_val = str(vals[4]).strip() if vals[4] is not None else ''
    
    if id_regex.match(id_val):
        valid_ids.append((r, vals, id_val))
    else:
        invalid_ids.append((r, vals, id_val))
        
    if len(fb_val) > 100:
        feedback_over_100.append((r, vals, len(fb_val)))

print(f'Total non-empty data rows: {len(valid_ids) + len(invalid_ids)}')
print(f'Valid IDs (符合正確格式): {len(valid_ids)}')
print(f'Invalid IDs: {len(invalid_ids)}')
print(f'Feedback > 100 chars: {len(feedback_over_100)}')

print('\nInvalid ID list:')
for r, vals, id_str in invalid_ids:
    print(f'Row {r:3d} | School: {vals[1]} | Name: {vals[2]} | ID: "{id_str}"')
