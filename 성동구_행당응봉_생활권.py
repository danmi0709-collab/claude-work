"""성동구 탭 - 4행당응봉 생활권 입력
이름 매칭 우선, 안 되면 연식+세대수 완전 일치 시만 매핑 (유사 불가)
"""
import sys, io, re, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE'
FILE = BASE + r'\시세시트_링크추가.xlsx'

def norm(s):
    return re.sub(r'[()\[\]{}·.\-_,/&+ ]', '', str(s).lower()) if s else ''

# 이미지에서 읽은 4행당응봉 단지 목록 (이름, 연식, 세대수, 등급)
IMAGE_DATA = [
    # 최상
    ('라체르보푸르지오써밋', 2025, 958,  '최상'),
    # 상
    ('서울숲리버뷰자이',     2018, 1034, '상'),
    ('서울숲더샵',           2014,  495, '상'),   # 탭: 서울숲더샵(주상복합)
    # 중
    ('응봉대림1차',          1986,  855, '중'),   # 탭: 대림1차
    ('대림강변타운',         2001, 1150, '중'),
    ('행당두산위브',         2009,  465, '중'),   # 탭: 두산위브
    ('행당한진타운',         2000, 2123, '중'),
    ('서울숲한신더휴',       2003, 1410, '중'),
    ('행당대림',             2000, 3404, '중'),
    ('서울숲리버그린동아',   2003,  375, '중'),
    ('행당브라운스톤',       2005,  208, '중'),
    ('서울숲행당푸르지오',   2011,  457, '중'),
    ('금호삼성래미안',       2001,  582, '중'),
    ('행당동신동아',         1995,  636, '중'),   # 탭에 없을 수 있음
    ('금호동벽산',           2001, 1707, '중'),   # 탭: 벽산
    # 하
    ('응봉대림2차',          1989,  410, '하'),   # 탭: 대림2차
    ('금호현대',             1990,  644, '하'),
    ('응봉신동아',           1996,  434, '하'),   # 탭: 신동아
]

wb = openpyxl.load_workbook(FILE)
ws = wb['성동구']

# 탭 단지 인덱스 구축 (B열 없는 행만)
name_to_rows = {}   # norm(name) → [row, ...]
yhh_to_rows  = {}   # (year, hh) → [row, ...]

for r in range(5, ws.max_row + 1):
    name  = ws.cell(r, 5).value
    b_val = ws.cell(r, 2).value
    if not name or b_val:  # 이미 매핑된 건 건드리지 않음
        continue
    f_val = ws.cell(r, 6).value
    g_val = ws.cell(r, 7).value
    try:
        year = int(str(f_val)[:4])
        hh   = int(g_val)
    except:
        year = hh = None

    key_n = norm(name)
    name_to_rows.setdefault(key_n, []).append(r)

    if year and hh:
        yhh_to_rows.setdefault((year, hh), []).append(r)

# 매칭 & 입력
print('=== 4행당응봉 생활권 입력 ===')
updated  = 0
no_match = []

for img_name, img_year, img_hh, grade in IMAGE_DATA:
    img_norm = norm(img_name)

    # 방법1: 이름 정규화 완전일치
    matched_rows = name_to_rows.get(img_norm, [])
    method = '이름'

    # 방법2: 이름 포함관계
    if not matched_rows:
        for tab_norm, rows in name_to_rows.items():
            if img_norm and (img_norm in tab_norm or tab_norm in img_norm):
                matched_rows = rows
                method = '포함'
                break

    # 방법3: 연식+세대수 완전 일치만 (유사 불가)
    if not matched_rows:
        matched_rows = yhh_to_rows.get((img_year, img_hh), [])
        method = '연식+세대수'

    if matched_rows:
        for r in matched_rows:
            ws.cell(r, 2, value='4행당응봉')
            ws.cell(r, 3, value=grade)
            updated += 1
        first_name = ws.cell(matched_rows[0], 5).value
        print(f'  [{grade}] {img_name:20} → {first_name} ({method}, {len(matched_rows)}행)')
    else:
        no_match.append(img_name)
        print(f'  [미매칭] {img_name} ({img_year}년, {img_hh}세대)')

wb.save(FILE)
print(f'\n업데이트: {updated}행')

if no_match:
    print(f'매칭 실패: {no_match}')
else:
    print('모든 항목 매칭 성공!')

# 결과 확인
print('\n=== 결과 확인 ===')
wb2 = openpyxl.load_workbook(FILE, data_only=True)
ws2 = wb2['성동구']
from collections import defaultdict
result = defaultdict(list)
seen = set()
for r in range(5, ws2.max_row + 1):
    if ws2.cell(r, 2).value == '4행당응봉':
        name = ws2.cell(r, 5).value
        grade = ws2.cell(r, 3).value
        if name not in seen:
            seen.add(name)
            result[grade].append(name)
for g in ['최상', '상', '중', '하']:
    if result[g]:
        print(f'  [{g}] {", ".join(result[g])}')
