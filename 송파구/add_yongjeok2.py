# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

# 전체 공동주택 데이터 로드 (아파트 필터 제거)
df = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')
apt_all = df[df['주용도코드명'] == '공동주택']
df_nonzero = apt_all[apt_all['용적률(%)'] > 0]

# 건물명 기준 평균 용적률
yong = df_nonzero.groupby('건물명')['용적률(%)'].mean().round(0).astype(int).to_dict()

def normalize(s):
    s = re.sub(r'\(.*?\)', '', s)  # 괄호 내용 제거
    s = re.sub(r'아파트$', '', s)
    s = re.sub(r'\s+', '', s)
    return s.strip()

# 정규화된 excel 이름 맵
yong_norm = {}
for k, v in yong.items():
    nk = normalize(k)
    if nk not in yong_norm or v > 0:
        yong_norm[nk] = v

# HTML 읽기
with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

match = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(match.group(1))

def extract_apt_key(name):
    n = name.strip()
    # 연도 + 이후 제거: "잠실엘스 08' 5,678^" → "잠실엘스"
    n = re.sub(r"\s+\d{2,4}['`']\s.*$", '', n)
    n = re.sub(r"\s+[\d,]+\^.*$", '', n)
    return n.strip()

def find_yong(apt_name):
    key = extract_apt_key(apt_name)
    nkey = normalize(key)

    # 1) 정규화 정확 매칭
    if nkey in yong_norm and yong_norm[nkey] > 0:
        return yong_norm[nkey]
    # 2) 원본 정확 매칭
    if key in yong and yong[key] > 0:
        return yong[key]
    # 3) 포함 관계 매칭 (정규화)
    best = None
    best_len = 0
    for ename, val in yong_norm.items():
        if val <= 0:
            continue
        if nkey in ename or ename in nkey:
            if len(ename) > best_len:
                best = val
                best_len = len(ename)
    return best

matched = 0
unmatched = []

for apt in apartments:
    raw_name = apt.get('name', '')
    v = find_yong(raw_name)
    if v and v > 0:
        if '용적률' not in apt['text']:
            apt['text'] = apt['text'].rstrip() + f'\n용적률 {v}%'
        matched += 1
    else:
        unmatched.append(extract_apt_key(raw_name))

with open('송파구/match_result.txt', 'w', encoding='utf-8') as f:
    f.write(f'매칭 성공: {matched} / {len(apartments)}\n\n')
    f.write('미매칭 단지:\n')
    for u in unmatched:
        f.write(f'  - {u}\n')

new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:match.start(1)] + new_arr + html[match.end(1):]

out = '송파구_한판정리(그늘)_4_용적률.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'완료: {matched}/{len(apartments)} 매칭, 파일 저장됨')
