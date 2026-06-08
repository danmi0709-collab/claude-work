# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

df = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')
apt_all = df[df['주용도코드명'] == '공동주택'].copy()

# 사용승인일 → 연도 (2자리)
apt_all['연도'] = apt_all['사용승인일'].astype(str).str[:4]
apt_all['연도2'] = apt_all['연도'].str[-2:]

# 동 번호를 제거한 단지명 만들기 ("잠실주공아파트 518동" → "잠실주공아파트")
def strip_dong(name):
    if pd.isna(name):
        return ''
    n = re.sub(r'\s*\d+동$', '', str(name))
    n = re.sub(r'\s*제?\d+동$', '', n)
    n = re.sub(r'\s*상가동$', '', n)
    return n.strip()

apt_all['단지명'] = apt_all['건물명'].apply(strip_dong)

# 단지별 집계: 총세대수, 연도, 평균 용적률
grouped = apt_all.groupby('단지명').agg(
    총세대수=('세대수(세대)', 'sum'),
    연도=('연도2', 'first'),
    용적률=('용적률(%)', lambda x: x[x>0].mean() if (x>0).any() else 0)
).reset_index()

grouped['용적률'] = grouped['용적률'].fillna(0).round(0).astype(int)

# (연도, 세대수) → 용적률 매핑
year_house_map = {}
for _, row in grouped.iterrows():
    if row['용적률'] > 0:
        key = (row['연도'], int(row['총세대수']))
        year_house_map[key] = (row['단지명'], row['용적률'])

# 세대수만으로도 매핑 (오차 ±5)
house_only = {}
for _, row in grouped.iterrows():
    if row['용적률'] > 0:
        h = int(row['총세대수'])
        if h > 50:  # 너무 작은 건 제외
            house_only[h] = (row['단지명'], row['용적률'], row['연도'])

# HTML 읽기
with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

match = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(match.group(1))

def parse_apt(name):
    """이름에서 연도2자리, 세대수 추출"""
    # "잠실엘스 08' 5,678^" → ('08', 5678)
    m = re.search(r"(\d{2})['`'″]\s*([\d,]+)\^", name)
    if m:
        year = m.group(1)
        house = int(m.group(2).replace(',', ''))
        return year, house
    return None, None

matched = 0
unmatched = []
results = []

for apt in apartments:
    raw_name = apt.get('name', '')
    year, house = parse_apt(raw_name)

    yong_val = None
    matched_name = None

    if year and house:
        # 1) (연도, 세대수) 정확 매칭
        if (year, house) in year_house_map:
            matched_name, yong_val = year_house_map[(year, house)]
        else:
            # 2) 세대수 ±10 허용 + 연도 일치
            for (y, h), (nm, v) in year_house_map.items():
                if y == year and abs(h - house) <= 10:
                    matched_name, yong_val = nm, v
                    break

            # 3) 세대수만 매칭 (연도 ±2 허용)
            if not yong_val:
                for h, (nm, v, y) in house_only.items():
                    if abs(h - house) <= 5:
                        try:
                            if abs(int(y) - int(year)) <= 2:
                                matched_name, yong_val = nm, v
                                break
                        except:
                            pass

    if yong_val and yong_val > 0:
        if '용적률' not in apt['text']:
            apt['text'] = apt['text'].rstrip() + f'\n용적률 {yong_val}%'
        matched += 1
        results.append(f"OK [{raw_name}] → {matched_name} ({yong_val}%)")
    else:
        unmatched.append(f"{raw_name} (연도={year}, 세대={house})")
        results.append(f"-- [{raw_name}] 매칭실패 (연도={year}, 세대={house})")

with open('송파구/match_result3.txt', 'w', encoding='utf-8') as f:
    f.write(f'매칭 성공: {matched} / {len(apartments)}\n\n')
    f.write('=== 매칭 결과 ===\n')
    for r in results:
        f.write(r + '\n')
    f.write('\n=== 미매칭 ===\n')
    for u in unmatched:
        f.write(f'  - {u}\n')

new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:match.start(1)] + new_arr + html[match.end(1):]

out = '송파구_한판정리(그늘)_4_용적률.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'완료: {matched}/{len(apartments)} 매칭')
