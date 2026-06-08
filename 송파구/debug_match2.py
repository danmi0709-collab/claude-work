# -*- coding: utf-8 -*-
import pandas as pd
import re
import json
import sys

df = pd.read_excel('송파구/송파구_아파트_용적률.xlsx')
df_nonzero = df[df['용적률'] > 0]
yong = df_nonzero.groupby('아파트명')['용적률'].mean().round(0).astype(int).to_dict()

with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

match = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(match.group(1))

html_names = [a['name'] for a in apartments]
excel_names = list(yong.keys())

with open('송파구/debug_names.txt', 'w', encoding='utf-8') as f:
    f.write("=== HTML 아파트명 목록 ===\n")
    for n in html_names:
        f.write(repr(n) + '\n')
    f.write("\n=== Excel 아파트명 목록 ===\n")
    for n in excel_names:
        f.write(repr(n) + '\n')
