"""성동구 탭 - 5마장 생활권 입력"""
import sys, io, re, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE'
FILE = BASE + r'\시세시트_링크추가.xlsx'

def norm(s):
    return re.sub(r'[()\[\]{}·.\-_,/&+ ]', '', str(s).lower()) if s else ''

IMAGE_DATA = [
    # 최상/상 없음
    # 중
    ('왕십리금호어울림', 2006, 367,  '중'),
    ('왕십리삼성',       1996, 430,  '중'),
    ('왕십리대성유니드', 2004, 248,  '중'),
    ('청계현대',         1998, 1017, '중'),
    ('청계미소지움',     2004, 286,  '중'),
    # 하
    ('마장세림',         1986, 811,  '하'),   # 탭: 세림
    ('왕십리두산',       1997, 251,  '하'),   # 탭에 없을 수 있음
    ('하이츠',           1989, 270,  '하'),
]

wb = openpyxl.load_workbook(FILE)
ws = wb['성동구']

# 탭 인덱스 (B열 없는 행만)
name_to_rows = {}
yhh_to_rows  = {}
for r in range(5, ws.max_row + 1):
    name  = ws.cell(r, 5).value
    b_val = ws.cell(r, 2).value
    if not name or b_val: continue
    f_val = ws.cell(r, 6).value
    g_val = ws.cell(r, 7).value
    try: year = int(str(f_val)[:4]); hh = int(g_val)
    except: year = hh = None
    name_to_rows.setdefault(norm(name), []).append(r)
    if year and hh:
        yhh_to_rows.setdefault((year, hh), []).append(r)

print('=== 5마장 생활권 입력 ===')
updated = 0
no_match = []

for img_name, img_year, img_hh, grade in IMAGE_DATA:
    img_norm = norm(img_name)

    # 방법1: 이름 완전일치
    matched_rows = name_to_rows.get(img_norm, [])
    method = '이름'

    # 방법2: 이름 포함
    if not matched_rows:
        for tab_norm, rows in name_to_rows.items():
            if img_norm and (img_norm in tab_norm or tab_norm in img_norm):
                matched_rows = rows; method = '포함'; break

    # 방법3: 연식+세대수 완전일치만
    if not matched_rows:
        matched_rows = yhh_to_rows.get((img_year, img_hh), [])
        method = '연식+세대수'

    if matched_rows:
        for r in matched_rows:
            ws.cell(r, 2, value='5마장')
            ws.cell(r, 3, value=grade)
            updated += 1
        print(f'  [{grade}] {img_name:18} → {ws.cell(matched_rows[0],5).value} ({method}, {len(matched_rows)}행)')
    else:
        no_match.append(img_name)
        print(f'  [미매칭] {img_name} ({img_year}년, {img_hh}세대)')

wb.save(FILE)
print(f'\n업데이트: {updated}행')
if no_match: print(f'매칭 실패: {no_match}')
else: print('모든 항목 매칭 성공!')

# 결과 확인
print('\n=== 결과 ===')
wb2 = openpyxl.load_workbook(FILE, data_only=True)
ws2 = wb2['성동구']
from collections import defaultdict
result = defaultdict(list); seen = set()
for r in range(5, ws2.max_row+1):
    if ws2.cell(r,2).value == '5마장':
        name = ws2.cell(r,5).value; grade = ws2.cell(r,3).value
        if name not in seen: seen.add(name); result[grade].append(name)
for g in ['최상','상','중','하']:
    if result[g]: print(f'  [{g}] {", ".join(result[g])}')
