# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

df1 = pd.read_excel('C:/Users/한나/Downloads/02. 총괄표제부_2026-05-05 09_20_17.xlsx')
df2 = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')

def prep(df):
    d = df[df['용적률(%)'] > 0].copy()
    d = d[d['건물명'].notna()]
    d = d[d['건물명'].astype(str).str.strip() != '']
    d = d[d['건물명'].astype(str) != 'nan']
    d['연도'] = pd.to_numeric(d['사용승인일'].astype(str).str[:4], errors='coerce')
    d['세대수'] = pd.to_numeric(d['세대수(세대)'], errors='coerce').fillna(0).astype(int)
    return d

d1 = prep(df1)
d2 = prep(df2)

def strip_dong(name):
    if pd.isna(name): return ''
    n = re.sub(r'\s*\d+동$', '', str(name))
    n = re.sub(r'\s*제?\d+동$', '', n)
    n = re.sub(r'\s*상가동$', '', n)
    return n.strip()

d2['단지명'] = d2['건물명'].apply(strip_dong)
d2_grp = d2.groupby('단지명').agg(
    세대수=('세대수', 'sum'),
    연도=('연도', 'first'),
    용적률=('용적률(%)', 'mean')
).reset_index()
d2_grp = d2_grp[d2_grp['단지명'].str.len() >= 2]
d2_grp['용적률'] = d2_grp['용적률'].round(0).astype(int)

# 통합: [(이름, 연도, 세대수, 용적률)]
records = []
for _, r in d1.iterrows():
    if pd.notna(r['연도']):
        records.append((str(r['건물명']).strip(), int(r['연도']), int(r['세대수']), round(r['용적률(%)'])))
for _, r in d2_grp.iterrows():
    if pd.notna(r['연도']):
        records.append((str(r['단지명']).strip(), int(r['연도']), int(r['세대수']), int(r['용적률'])))

with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(m.group(1))

def parse_apt(name):
    mm = re.search(r"(\d{2})['`'″’]\s*([\d,]+)\^", name)
    if mm:
        y2 = int(mm.group(1))
        year = 2000 + y2 if y2 < 50 else 1900 + y2
        return year, int(mm.group(2).replace(',', ''))
    return None, None

def normalize(s):
    if not isinstance(s, str): return ''
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'아파트$', '', s)
    s = re.sub(r'\s+', '', s)
    return s.strip()

def extract_key(name):
    n = re.sub(r"\s+\d{2,4}['`'″’].*$", '', name)
    n = re.sub(r"\s+[\d,]+\^.*$", '', n)
    return n.strip()

def find_match(raw):
    year, house = parse_apt(raw)
    key = extract_key(raw)
    nkey = normalize(key)

    candidates = []  # (점수, 이름, 용적률)

    for nm, y, h, v in records:
        if v <= 0: continue
        nnm = normalize(nm)
        if not nnm or len(nnm) < 2: continue

        score = 0
        # 연도 일치
        if year and y:
            diff = abs(y - year)
            if diff == 0: score += 100
            elif diff <= 1: score += 50
            elif diff <= 2: score += 20
            else: score -= 50
        # 세대수 일치
        if house and h:
            hdiff = abs(h - house)
            if hdiff == 0: score += 100
            elif hdiff <= 5: score += 60
            elif hdiff <= 20: score += 20
            else: score -= 50
        # 이름 유사도
        if nkey and nnm:
            if nkey == nnm: score += 200
            elif nkey in nnm or nnm in nkey:
                # 길이 비율로 가중
                ratio = min(len(nkey), len(nnm)) / max(len(nkey), len(nnm))
                score += int(80 * ratio)

        if score > 0:
            candidates.append((score, nm, v))

    if not candidates:
        return None, None
    candidates.sort(reverse=True)
    best_score, best_name, best_v = candidates[0]
    # 최소 점수 임계: 130 (이름 유사도 또는 연도+세대수 중 둘 다 어느정도 맞아야)
    if best_score >= 130:
        return best_name, best_v
    return None, None

matched = 0
results = []

for apt in apartments:
    raw = apt.get('name', '')
    # name에 연도/세대수가 없으면 text 첫줄에서 찾기
    if not re.search(r"\d{2}['`'″’]\s*[\d,]+\^", raw):
        text_first2 = '\n'.join(apt.get('text', '').split('\n')[:2])
        raw = text_first2.replace('\n', ' ')
    found, yong = find_match(raw)

    if yong and yong > 0:
        if '용적률' not in apt['text']:
            apt['text'] = apt['text'].rstrip() + f'\n용적률 {yong}%'
        matched += 1
        results.append(f"OK [{raw}] → {found} ({yong}%)")
    else:
        results.append(f"-- [{raw}] 실패")

with open('송파구/match_v2.txt', 'w', encoding='utf-8') as f:
    f.write(f'매칭: {matched} / {len(apartments)}\n\n')
    for r in results:
        f.write(r + '\n')

new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:m.start(1)] + new_arr + html[m.end(1):]

with open('송파구_한판정리(그늘)_4_용적률.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'완료: {matched}/{len(apartments)}')
