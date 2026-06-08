# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

# 총괄표제부 + 표제부 모두 사용
df1 = pd.read_excel('C:/Users/한나/Downloads/02. 총괄표제부_2026-05-05 09_20_17.xlsx')
df2 = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')

# 둘 다 용적률>0인 행만
def prep(df):
    d = df[df['용적률(%)'] > 0].copy()
    d['연도2'] = d['사용승인일'].astype(str).str[2:4]
    d['세대수'] = d['세대수(세대)'].fillna(0).astype(int)
    return d

d1 = prep(df1)  # 총괄표제부 (큰 단지)
d2 = prep(df2)  # 표제부 (개별 동)

# 표제부는 동별이라 단지명+세대수합으로 집계
def strip_dong(name):
    if pd.isna(name): return ''
    n = re.sub(r'\s*\d+동$', '', str(name))
    n = re.sub(r'\s*제?\d+동$', '', n)
    n = re.sub(r'\s*상가동$', '', n)
    return n.strip()

d2['단지명'] = d2['건물명'].apply(strip_dong)
d2_grp = d2.groupby('단지명').agg(
    세대수=('세대수', 'sum'),
    연도2=('연도2', 'first'),
    용적률=('용적률(%)', 'mean'),
    건물명=('단지명', 'first')
).reset_index(drop=True)
d2_grp['용적률'] = d2_grp['용적률'].round(0).astype(int)

# 통합 매핑: (연도2, 세대수) → (이름, 용적률)
mapping = {}
name_map = {}  # 이름 → 용적률
for _, r in d1.iterrows():
    key = (r['연도2'], int(r['세대수']))
    mapping[key] = (r['건물명'], round(r['용적률(%)']))
    name_map[r['건물명']] = round(r['용적률(%)'])

for _, r in d2_grp.iterrows():
    key = (r['연도2'], int(r['세대수']))
    if key not in mapping:
        mapping[key] = (r['건물명'], r['용적률'])
    name_map[r['건물명']] = r['용적률']

# HTML 읽기
with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(m.group(1))

def parse_apt(name):
    mm = re.search(r"(\d{2})['`'″’]\s*([\d,]+)\^", name)
    if mm:
        return mm.group(1), int(mm.group(2).replace(',', ''))
    return None, None

def normalize(s):
    if not isinstance(s, str):
        return ''
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'아파트$', '', s)
    s = re.sub(r'\s+', '', s)
    return s.strip()

def extract_key(name):
    n = re.sub(r"\s+\d{2,4}['`'″’].*$", '', name)
    n = re.sub(r"\s+[\d,]+\^.*$", '', n)
    return n.strip()

matched = 0
results = []
unmatched = []

for apt in apartments:
    raw = apt.get('name', '')
    year, house = parse_apt(raw)
    yong = None
    found = None

    # 1) 연도+세대수 정확 매칭
    if year and house and (year, house) in mapping:
        found, yong = mapping[(year, house)]
    # 2) 연도일치 + 세대수 ±10
    if not yong and year and house:
        for (y, h), (nm, v) in mapping.items():
            if y == year and abs(h - house) <= 10:
                found, yong = nm, v
                break
    # 3) 세대수만 ±5 (연도 없는 경우 등)
    if not yong and house:
        for (y, h), (nm, v) in mapping.items():
            if abs(h - house) <= 5:
                found, yong = nm, v
                break
    # 4) 이름 매칭 (정규화 포함관계)
    if not yong:
        key = extract_key(raw)
        nkey = normalize(key)
        if key in name_map:
            found, yong = key, name_map[key]
        else:
            best_len = 0
            for nm, v in name_map.items():
                nnm = normalize(nm)
                if nkey and (nkey in nnm or nnm in nkey):
                    if len(nnm) > best_len:
                        found, yong, best_len = nm, v, len(nnm)

    if yong and yong > 0:
        if '용적률' not in apt['text']:
            apt['text'] = apt['text'].rstrip() + f'\n용적률 {yong}%'
        matched += 1
        results.append(f"OK [{raw}] → {found} ({yong}%)")
    else:
        unmatched.append(raw)
        results.append(f"-- [{raw}] 실패 (연도={year}, 세대={house})")

with open('송파구/match_final.txt', 'w', encoding='utf-8') as f:
    f.write(f'매칭 성공: {matched} / {len(apartments)}\n\n')
    for r in results:
        f.write(r + '\n')

new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:m.start(1)] + new_arr + html[m.end(1):]

with open('송파구_한판정리(그늘)_4_용적률.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'완료: {matched}/{len(apartments)} 매칭')
