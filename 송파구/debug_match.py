# -*- coding: utf-8 -*-
import pandas as pd
import re
import json

df = pd.read_excel('송파구/송파구_아파트_용적률.xlsx')
df_nonzero = df[df['용적률'] > 0]
yong = df_nonzero.groupby('아파트명')['용적률'].mean().round(0).astype(int).to_dict()

with open('송파구_한판정리(그늘)_3_시세완.html', encoding='utf-8') as f:
    html = f.read()

match = re.search(r'const apartments=(\[.*?\]);', html, re.DOTALL)
apartments = json.loads(match.group(1))

html_names = [a['name'] for a in apartments]
excel_names = list(yong.keys())

print("=== HTML 아파트명 목록 ===")
for n in html_names:
    print(repr(n))

print("\n=== Excel 아파트명 목록 (샘플 50개) ===")
for n in excel_names[:50]:
    print(repr(n))
