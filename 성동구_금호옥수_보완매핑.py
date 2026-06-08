"""성동구 탭 - 2금호옥수 보완 매핑
이름 매칭이 안 된 단지를 세대수 + 연식으로 가장 유사한 매핑 결과 찾아 등급 배정
"""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r'C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE'
FILE = BASE + r'\시세시트_링크추가.xlsx'

# 이미 매핑된 단지의 (세대수, 연식, 등급) 목록 (grade_map에 있던 것들)
grade_map = {
    'e편한세상옥수파크힐스': '최상',
    '신금호파크자이':        '최상',
    '래미안옥수리버젠':      '상',
    '래미안하이리버':        '상',
    'e편한세상금호파크힐스': '중',
    '한남하이츠':            '중',
    '옥수어울림':            '중',
    '옥수강변풍림아이원':    '중',
    '서울숲푸르지오1차':     '중',
    '서울숲푸르지오2차':     '중',
    '힐스테이트서울숲리버':  '중',
    '금호센트럴자이':        '중',
    '금호브라운스톤1차':     '중',
    '옥수삼성':              '중',
    '옥수하이츠':            '중',
    '금호대우':              '중',
    '옥수현대':              '중',
    '금호자이2차':           '중',
    '금호1차푸르지오':       '중',
    '두산':                  '중',
    '금호한신휴플러스':      '중',
    '롯데':                  '중',
    '극동그린':              '하',
    '옥수극동':              '하',
}

wb = openpyxl.load_workbook(FILE, data_only=True)
ws = wb['성동구']

# 1단계: 이미 매핑된 단지의 (세대수, 연식) 수집
matched_info = []  # [(세대수, 연식, 등급, 단지명)]
seen = set()
for r in range(5, ws.max_row + 1):
    dong = str(ws.cell(r, 1).value or '')
    if not any(x in dong for x in ['금호', '옥수']): continue
    name = ws.cell(r, 5).value
    b = ws.cell(r, 2).value
    if not name or name in seen: continue
    if b == '2금호옥수':  # 이미 매핑된 건물
        seen.add(name)
        f_val = ws.cell(r, 6).value  # 사용승인일
        g_val = ws.cell(r, 7).value  # 세대수
        try:
            year = int(str(f_val)[:4])
            hh = int(g_val)
        except:
            continue
        grade = ws.cell(r, 3).value
        matched_info.append((hh, year, grade, name))

print('=== 이미 매핑된 단지 ===')
for hh, yr, gr, nm in sorted(matched_info, key=lambda x: (x[1], x[0])):
    print(f'  {nm:30} | {yr}년 | {hh:5}세대 | {gr}')

# 2단계: 미매핑 단지 찾기
print('\n=== 세대수+연식 보완 매핑 ===')
unmatched = []
seen2 = set()
for r in range(5, ws.max_row + 1):
    dong = str(ws.cell(r, 1).value or '')
    if not any(x in dong for x in ['금호', '옥수']): continue
    name = ws.cell(r, 5).value
    b = ws.cell(r, 2).value
    if not name or name in seen2: continue
    seen2.add(name)
    if not b:  # 미매핑
        f_val = ws.cell(r, 6).value
        g_val = ws.cell(r, 7).value
        try:
            year = int(str(f_val)[:4])
            hh = int(g_val)
        except:
            continue
        unmatched.append((r, dong, name, year, hh))

# 3단계: 가장 유사한 매핑 단지 찾기 (세대수 차이 + 연식 차이 가중합)
def find_nearest_grade(target_hh, target_year, matched):
    best = None
    best_dist = float('inf')
    for hh, yr, grade, nm in matched:
        # 세대수 차이(정규화) + 연식 차이(정규화) 가중합
        hh_dist = abs(hh - target_hh) / 2000  # 세대수 최대 ~2000
        yr_dist = abs(yr - target_year) / 30   # 연식 최대 ~30년 차이
        dist = hh_dist + yr_dist * 2  # 연식에 가중치 더 줌
        if dist < best_dist:
            best_dist = dist
            best = (nm, hh, yr, grade, dist)
    return best

results = []
for r, dong, name, year, hh in unmatched:
    nearest = find_nearest_grade(hh, year, matched_info)
    if nearest:
        ref_nm, ref_hh, ref_yr, grade, dist = nearest
        print(f'  [{name}] {year}년 {hh}세대')
        print(f'    → 가장 유사: [{ref_nm}] {ref_yr}년 {ref_hh}세대 → 등급={grade} (거리={dist:.3f})')
        results.append((name, year, hh, grade, ref_nm))

wb.close()

# 4단계: 엑셀에 반영
print('\n=== 엑셀 업데이트 ===')
wb2 = openpyxl.load_workbook(FILE)
ws2 = wb2['성동구']

# 등급 결과를 단지명 기준으로 매핑
name_grade = {r[0]: r[3] for r in results}

updated = 0
for r in range(5, ws2.max_row + 1):
    dong = str(ws2.cell(r, 1).value or '')
    if not any(x in dong for x in ['금호', '옥수']): continue
    name = ws2.cell(r, 5).value
    b = ws2.cell(r, 2).value
    if not name or b: continue  # 이미 매핑된 건 건드리지 않음
    if name in name_grade:
        ws2.cell(r, 2, value='2금호옥수')
        ws2.cell(r, 3, value=name_grade[name])
        updated += 1

wb2.save(FILE)
print(f'완료: {updated}행 업데이트')
print()
for r in results:
    print(f'  {r[0]} → 2금호옥수 / {r[3]} (기준: {r[4]})')
