# -*- coding: utf-8 -*-
import sys, io, openpyxl, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('1151031基隆百福科技團線上研習2.xlsx', data_only=True)
ws = wb['工作表1']
print(f'Total rows in sheet: {ws.max_row}')

id_regex = re.compile(r'^[A-Za-z][1289]\d{8}$')

valid_ids = []
invalid_ids = []
feedback_over_100 = []

for r in range(1, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    if not any(v is not None for v in vals):
        continue
    time_val = vals[0]
    school_val = vals[1]
    name_val = vals[2]
    id_val = vals[3]
    fb_val = vals[4]
    email_val = vals[5] if len(vals) > 5 else None

    id_str = str(id_val).strip() if id_val is not None else ''
    fb_str = str(fb_val).strip() if fb_val is not None else ''
    
    if id_regex.match(id_str):
        valid_ids.append((r, vals, id_str))
    else:
        invalid_ids.append((r, vals, id_str))
        
    if len(fb_str) > 100:
        feedback_over_100.append((r, vals, len(fb_str)))

print(f'Valid IDs (符合正確格式): {len(valid_ids)}')
print(f'Invalid IDs: {len(invalid_ids)}')
print(f'Feedback > 100 chars: {len(feedback_over_100)}')

print('\nInvalid ID list:')
for r, vals, id_str in invalid_ids:
    print(f'Row {r:2d} | School: {vals[1]} | Name: {vals[2]} | ID: "{id_str}"')

print(f'\nFeedback <= 100 chars (未達 100 字):')
for r in range(1, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    if not any(v is not None for v in vals):
        continue
    fb_str = str(vals[4]).strip() if vals[4] is not None else ''
    if len(fb_str) <= 100:
        print(f'Row {r:2d} | Name: {vals[2]} | Len: {len(fb_str):2d} | Feedback: {fb_str}')
