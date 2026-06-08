"""성동구 탭 - 3왕십리 생활권 입력
1. 금호삼성래미안/벽산 B,C열 삭제 (잘못 매핑된 것 제거)
2. 3왕십리 생활권 입력 (최상/상/중/하)
   - 이름 매칭 우선, 안 되면 연식+세대수 완전 일치 시만 매핑
"""
import sys, io, re, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE'
FILE = BASE + r'\시세시트_링크추가.xlsx'

def norm(s):
    return re.sub(r'[()\[\]{}·.\-_,/&+ ]', '', str(s).lower()) if s else ''

# ── 이미지에서 읽은 3왕십리 단지 목록 (이름, 연식, 세대수, 등급) ──
IMAGE_DATA = [
    ('센트라스',          2016, 2529, '최상'),
    ('텐즈힐1',           2015, 1702, '상'),
    ('텐즈힐2',           2014, 1148, '상'),
    ('왕십리자이',         2017,  713, '상'),
    ('서울숲삼복',         1998,  498, '중'),   # 탭 이름: 서울숲삼부
    ('왕십리kcc스위첸',    2016,  272, '중'),
    ('왕십리풍림아이원',   2004,  758, '중'),
    ('성동삼성쉐르빌',     2006,  342, '중'),   # 탭: 성동삼성쉐르빌(주상복합)
    ('청계벽산',          1996, 1332, '중'),
    ('하왕십리극동미라주', 2001,  414, '중'),   # 탭: 극동미라주
    ('한진해모로',         2001,  362, '중'),
    ('왕십리금호베스트빌', 2001,  458, '중'),   # 탭: 금호베스트빌
    ('무학현대',          1996,  277, '하'),
    ('한신무학',          1989,  480, '하'),
]

# ── 1단계: 잘못 매핑된 금호삼성래미안/벽산 삭제 ──
print('=== 1단계: 금호삼성래미안/벽산 매핑 삭제 ===')
wb = openpyxl.load_workbook(FILE)
ws = wb['성동구']

delete_names = {'금호삼성래미안', '벽산'}
deleted = 0
for r in range(5, ws.max_row + 1):
    name = ws.cell(r, 5).value
    dong = str(ws.cell(r, 1).value or '')
    if name in delete_names and '금호' in dong:
        ws.cell(r, 2, value=None)
        ws.cell(r, 3, value=None)
        deleted += 1

wb.save(FILE)
print(f'  삭제 완료: {deleted}행')

# ── 2단계: 성동구 탭에서 전체 단지 정보 수집 ──
print('\n=== 2단계: 3왕십리 생활권 입력 ===')
wb = openpyxl.load_workbook(FILE)
ws = wb['성동구']

# 탭 단지 인덱스 구축
name_to_rows  = {}   # norm(name) → [row, ...]
yhh_to_rows   = {}   # (year, hh) → [row, ...]

for r in range(5, ws.max_row + 1):
    name = ws.cell(r, 5).value
    if not name: continue
    f_val = ws.cell(r, 6).value
    g_val = ws.cell(r, 7).value
    try:
        year = int(str(f_val)[:4])
        hh   = int(g_val)
    except:
        year = hh = None

    key_n = norm(name)
    if key_n not in name_to_rows:
        name_to_rows[key_n] = []
    name_to_rows[key_n].append(r)

    if year and hh:
        key_yh = (year, hh)
        if key_yh not in yhh_to_rows:
            yhh_to_rows[key_yh] = []
        yhh_to_rows[key_yh].append(r)

# ── 3단계: 이미지 항목 매칭 & 입력 ──
updated = 0
no_match = []

for img_name, img_year, img_hh, grade in IMAGE_DATA:
    img_norm = norm(img_name)

    # 방법1: 이름 정규화 완전일치
    matched_rows = name_to_rows.get(img_norm, [])

    # 방법2: 이름 포함관계
    if not matched_rows:
        for tab_norm, rows in name_to_rows.items():
            if img_norm and (img_norm in tab_norm or tab_norm in img_norm):
                matched_rows = rows
                break

    # 방법3: 연식+세대수 완전 일치 (이름 매칭 실패 시에만)
    if not matched_rows:
        matched_rows = yhh_to_rows.get((img_year, img_hh), [])

    if matched_rows:
        for r in matched_rows:
            tab_name = ws.cell(r, 5).value
            ws.cell(r, 2, value='3왕십리')
            ws.cell(r, 3, value=grade)
            updated += 1
        method = '이름' if name_to_rows.get(img_norm) else ('포함' if len(matched_rows) > 0 and not yhh_to_rows.get((img_year, img_hh)) else '연식+세대수')
        first_tab = ws.cell(matched_rows[0], 5).value
        print(f'  [{grade}] {img_name:20} → {first_tab} ({method}매칭, {len(matched_rows)}행)')
    else:
        no_match.append(img_name)
        print(f'  [미매칭] {img_name} ({img_year}년, {img_hh}세대)')

wb.save(FILE)
print(f'\n업데이트: {updated}행')

if no_match:
    print(f'매칭 실패: {no_match}')
else:
    print('모든 항목 매칭 성공!')

# ── 4단계: 결과 확인 ──
print('\n=== 결과 확인 ===')
wb2 = openpyxl.load_workbook(FILE, data_only=True)
ws2 = wb2['성동구']
from collections import defaultdict
result = defaultdict(list)
seen = set()
for r in range(5, ws2.max_row + 1):
    if ws2.cell(r, 2).value == '3왕십리':
        name = ws2.cell(r, 5).value
        grade = ws2.cell(r, 3).value
        if name not in seen:
            seen.add(name)
            result[grade].append(name)
for g in ['최상', '상', '중', '하']:
    if result[g]:
        print(f'  [{g}] {", ".join(result[g])}')
