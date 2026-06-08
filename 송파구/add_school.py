# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

# 학교 데이터 로드 (헤더 12행)
df = pd.read_excel('2025년 하반기 교육통계 학교별 일람표(2025.10.1.기준).xlsx',
                    sheet_name='학교별 주요 통계', header=12)

# 송파구 초등학교 필터
sp = df[(df['행정구'] == '송파구') & (df['학교급'] == '초등학교')].copy()

# 학교명 → 학생수 매핑
school_map = {}
for _, r in sp.iterrows():
    name = str(r['학교명']).strip()
    students = int(r['학생수_총계_계']) if pd.notna(r['학생수_총계_계']) else 0
    school_map[name] = students
    # "서울토성초등학교" → "토성초"
    short = re.sub(r'^서울', '', name)
    short = re.sub(r'등학교$', '', short)
    if short:
        school_map[short] = students

# HTML의 schools 배열 업데이트
with open('송파구_한판정리(그늘)_4_용적률.html', encoding='utf-8') as f:
    html = f.read()

# schools 배열 찾기
m = re.search(r'(const\s+schoolAreas\s*=\s*)(\[.*?\]);', html, re.DOTALL)
schools_json = m.group(2)
schools = json.loads(schools_json)

matched = 0
results = []
for s in schools:
    nm = s.get('name', '')
    # "토성초/18" → ('토성초', '18')
    parts = nm.split('/')
    base = parts[0]  # "토성초"
    rest = '/'.join(parts[1:]) if len(parts) > 1 else ''

    # 매칭: "토성초" → "토성초등학교"
    full = base + '등학교'
    students = school_map.get(full) or school_map.get(base)

    if students:
        # "토성초/1201/18" 형식
        new_name = f"{base}/{students}/{rest}" if rest else f"{base}/{students}"
        s['name'] = new_name
        matched += 1
        results.append(f"OK [{nm}] → {new_name}")
    else:
        results.append(f"-- [{nm}] 미매칭")

# HTML 업데이트
new_json = json.dumps(schools, ensure_ascii=False)
new_html = html[:m.start(2)] + new_json + html[m.end(2):]

with open('송파구_한판정리(그늘)_5_학생수.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

with open('송파구/match_school.txt', 'w', encoding='utf-8') as f:
    f.write(f'매칭: {matched} / {len(schools)}\n\n')
    for r in results:
        f.write(r + '\n')

print(f'완료: {matched}/{len(schools)}')
