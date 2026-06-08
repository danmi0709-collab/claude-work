# -*- coding: utf-8 -*-
import json, re

with open('송파구_한판정리(그늘)_5_학생수.html', encoding='utf-8') as f:
    html = f.read()
m = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apts = json.loads(m.group(1))

checks = ['잠실주공5단지', '장미1차', '잠실르엘', '(풍납)극동', '풍납현대', '우성1,2,3차', '아시아선수촌']

with open('송파구/verify.txt', 'w', encoding='utf-8') as f:
    for apt in apts:
        for c in checks:
            if c in apt['name']:
                f.write(f'[{apt["name"]}]\n')
                f.write(f'  text: {apt["text"]}\n\n')
                break
