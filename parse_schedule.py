#!/usr/bin/env python3
"""Parse weekly schedule Excel and update sidur_live.html."""
import sys, json, re, openpyxl
from datetime import datetime, timedelta

excel_path = sys.argv[1]
week_start = sys.argv[2]  # YYYY-MM-DD

wb = openpyxl.load_workbook(excel_path, data_only=True)
ws = wb.worksheets[0]

SKIP = {'-', '--', ' ', '', None}

# Detect week columns by matching dates in row 2
target_date = datetime.strptime(week_start, '%Y-%m-%d')
day_cols = {}
for col in range(1, ws.max_column + 1):
    v = ws.cell(2, col).value
    if isinstance(v, datetime):
        for d in range(7):
            if v.date() == (target_date + timedelta(days=d)).date():
                day_cols[d] = col

if not day_cols:
    print(f'ERROR: week {week_start} not found in excel row 2')
    sys.exit(1)

print(f'Found days: {sorted(day_cols.keys())} => cols {[day_cols[d] for d in sorted(day_cols)]}')

# Detect name columns (col A = שם, col B = משפחה)
shifts = []
for row in range(3, ws.max_row + 1):
    a = ws.cell(row, 1).value
    b = ws.cell(row, 2).value
    if a is None and b is None:
        continue
    parts = [str(x).strip() for x in [a, b] if x is not None]
    name = ' '.join(parts).strip()
    if not name:
        continue
    for day, col in day_cols.items():
        v = ws.cell(row, col).value
        if v is None:
            continue
        v = str(v).strip()
        if v in SKIP:
            continue
        shifts.append({'day': day, 'role': v, 'name': name, 'note': ''})

print(f'Parsed {len(shifts)} shifts')

# Update sidur_live.html
schedule_data = {'weekStart': week_start, 'shifts': shifts}
data_str = json.dumps(schedule_data, ensure_ascii=False, indent=2)
new_block = f'let scheduleData = {data_str};'

with open('sidur_live.html', encoding='utf-8') as f:
    html = f.read()

updated = re.sub(
    r'let scheduleData = \{[\s\S]*?\n\};',
    new_block,
    html
)

if updated == html:
    print('ERROR: scheduleData pattern not found in HTML')
    sys.exit(1)

with open('sidur_live.html', 'w', encoding='utf-8') as f:
    f.write(updated)

print(f'Updated sidur_live.html: weekStart={week_start}, {len(shifts)} shifts')
