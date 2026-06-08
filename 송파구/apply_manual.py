# -*- coding: utf-8 -*-
"""
match_v2.txt에서 수동 입력된 용적률을 읽어 HTML에 반영.
형식: '-- [단지명] 138%' (실패 → 숫자%로 바꿔주면 인식)
"""
import re
import json

# match_v2.txt 파싱
manual = {}  # apt_name → 표시 문자열
with open('송파구/match_v2.txt', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        # '-- [name] 실패' 는 스킵
        if re.search(r'\]\s*실패\s*$', line):
            continue

        # '-- [name] ...값...' 추출
        m = re.match(r"--\s*\[([^\]]+)\]\s*(.+)$", line)
        if not m:
            continue
        apt_name = m.group(1)
        value = m.group(2).strip()

        # 숫자%(메모) / 숫자% (메모) / 숫자% / -(메모) 등 처리
        # case 1: 숫자%로 시작
        m1 = re.match(r"(\d+)%\s*(\([^)]+\))?", value)
        if m1:
            pct = m1.group(1)
            note = m1.group(2) or ''
            display = f"용적률 {pct}% {note}".strip()
            manual[apt_name] = display
            continue
        # case 2: -(메모) 형태 (용적률 없음, 상태만)
        m2 = re.match(r"-\s*(\([^)]+\))", value)
        if m2:
            manual[apt_name] = m2.group(1)
            continue

print(f'수동 입력된 단지: {len(manual)}개')
for k, v in manual.items():
    print(f'  {k}: {v}%')

# HTML 업데이트
with open('송파구_한판정리(그늘)_5_학생수.html', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(m.group(1))

added = 0
for apt in apartments:
    name = apt.get('name', '')
    if name in manual:
        # 기존 '용적률 ...' 줄 또는 '(...)' 마지막 줄 제거 후 재추가
        lines = apt['text'].split('\n')
        # 마지막에 추가됐던 용적률 / 상태 라인만 제거
        while lines and (lines[-1].startswith('용적률') or re.match(r'^\(.+\)$', lines[-1].strip())):
            lines.pop()
        new_text = '\n'.join(lines).rstrip()
        apt['text'] = new_text + '\n' + manual[name]
        added += 1

new_arr = json.dumps(apartments, ensure_ascii=False)
new_html = html[:m.start(1)] + new_arr + html[m.end(1):]

with open('송파구_한판정리(그늘)_5_학생수.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f'\n완료: {added}건 추가')
