# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

# 용적률 데이터 로드
df = pd.read_excel('송파구/송파구_아파트_용적률.xlsx')
# 단지별 평균 용적률 (0 제외)
df_nonzero = df[df['용적률'] > 0]
yong = df_nonzero.groupby('아파트명')['용적률'].mean().round(0).astype(int).to_dict()

# HTML 읽기
with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

# apartments 배열 추출
match = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(match.group(1))

# 아파트명에서 핵심 이름 추출 (연도, 세대수 등 제거)
def extract_name(name):
    # "잠실엘스 08' 5,678^" → "잠실엘스"
    # "(풍납)극동 87' 415^" → 그대로 쓰거나 "극동"
    n = name.strip()
    # 연도 패턴 제거: 숫자2자리 + 따옴표 이후
    n = re.sub(r"\s+\d{2}['`']\s.*$", '', n)
    n = re.sub(r"\s+\d{4}['`']\s.*$", '', n)
    # 세대수 패턴 제거
    n = re.sub(r'\s+[\d,]+\^.*$', '', n)
    return n.strip()

# 유사도 매칭 함수
def find_yong(apt_name):
    key = extract_name(apt_name)
    # 1) 정확히 일치
    if key in yong:
        return yong[key]
    # 2) 엑셀 이름이 key를 포함하거나 key가 엑셀 이름을 포함
    for ename, val in yong.items():
        if key in ename or ename in key:
            return val
    # 3) 괄호 제거 후 재시도
    key2 = re.sub(r'\(.*?\)', '', key).strip()
    if key2 and key2 in yong:
        return yong[key2]
    for ename, val in yong.items():
        if key2 and (key2 in ename or ename in key2):
            return val
    return None

matched = 0
unmatched = []

for apt in apartments:
    raw_name = apt.get('name', '')
    v = find_yong(raw_name)
    if v and v > 0:
        # text 마지막에 용적률 추가 (이미 추가된 경우 스킵)
        if '용적률' not in apt['text']:
            apt['text'] = apt['text'].rstrip() + f'\n용적률 {v}%'
        matched += 1
    else:
        unmatched.append(extract_name(raw_name))

print(f'매칭 성공: {matched} / {len(apartments)}')
print(f'미매칭 단지:')
for u in unmatched:
    print(f'  - {u}')

# HTML에 수정된 apartments 삽입
new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:match.start(1)] + new_arr + html[match.end(1):]

# 새 파일로 저장
out = '송파구_한판정리(그늘)_4_용적률.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'\n저장 완료: {out}')
