# สร้าง js/data.js จากไฟล์ Excel "Part_Locations_Labels_Upper.xlsx"
# ใช้เมื่อมีการย้าย/เพิ่ม/ลบ Part:  python tools/build-data.py [path/to/file.xlsx]
# - ชีต "รายการ Part": ตำแหน่ง | ชั้น | คอลัมน์ | Line | Child Item Code | Specification | Name | กลุ่ม | Status | หมายเหตุ
# - ชีต "แผนผัง": ใช้ดูว่าแต่ละชั้นมีกี่ช่อง (ช่องที่เขียนว่า "ไม่มีช่อง" = ไม่มีจริง)
# ตำแหน่งที่ไม่มี H9- นำหน้า (เช่น L8-01) จะถูกเติมให้เป็น H9-L8-01 ให้เหมือนกันทั้งหมด
import json, os, re, sys, openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'Part_Locations_Labels_Upper.xlsx')
LOC = re.compile(r'^(?:H9-)?L(\d+)-(\d+)$', re.I)

def s(v):
    return '' if v is None else str(v).strip()

def norm_loc(v):
    m = LOC.match(s(v))
    return 'H9-L%d-%02d' % (int(m.group(1)), int(m.group(2))) if m else ''

wb = openpyxl.load_workbook(SRC, data_only=True)

parts, seen = [], set()
for r in wb['รายการ Part'].iter_rows(min_row=2, values_only=True):
    if not r or not s(r[4]):
        continue
    loc = norm_loc(r[0])
    lv = col = 0
    if loc:
        m = LOC.match(loc); lv, col = int(m.group(1)), int(m.group(2))
    parts.append({
        'loc': loc,                 # '' = รอจัดพื้นที่
        'lv': lv, 'col': col,
        'line': int(r[3]) if s(r[3]).isdigit() else s(r[3]),
        'code': s(r[4]).upper(),
        'spec': s(r[5]).upper(),
        'name': s(r[6]),
        'grp': s(r[7]),
        'st': s(r[8]),
        'note': s(r[9]),
    })
    if loc in seen:
        print('!! ตำแหน่งซ้ำ', loc)
    if loc:
        seen.add(loc)

# จำนวนช่องต่อชั้นจากชีตแผนผัง (แถวที่ช่องแรกเป็นตำแหน่ง Lx-01)
levels = {}
for row in wb['แผนผัง'].iter_rows(values_only=True):
    first = norm_loc(s(row[0]).split('\n')[0]) if row and row[0] else ''
    if not first:
        continue
    lv = int(LOC.match(first).group(1))
    levels[lv] = sum(1 for c in row if c is not None and s(c) != 'ไม่มีช่อง')
for p in parts:  # กันกรณีแผนผังไม่ครบ
    if p['loc']:
        levels[p['lv']] = max(levels.get(p['lv'], 0), p['col'])

out = os.path.join(ROOT, 'js', 'data.js')
with open(out, 'w', encoding='utf-8', newline='\n') as f:
    f.write('// สร้างอัตโนมัติจาก %s ด้วย tools/build-data.py — ห้ามแก้มือ\n' % os.path.basename(SRC))
    f.write('window.LEVELS = ' + json.dumps({str(k): levels[k] for k in sorted(levels)}) + ';\n')
    f.write('window.PARTS = [\n' + ',\n'.join(json.dumps(p, ensure_ascii=False) for p in parts) + '\n];\n')

waiting = sum(1 for p in parts if not p['loc'])
print(len(parts), 'parts ·', len(parts) - waiting, 'มีตำแหน่ง ·', waiting, 'รอจัดพื้นที่ · ชั้น', levels)
