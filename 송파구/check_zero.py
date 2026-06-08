# -*- coding: utf-8 -*-
import pandas as pd

df = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')

# 잠실엘스, 헬리오시티 등 큰 단지의 용적률 확인
big = ['잠실엘스', '헬리오시티', '파크리오', '트리지움', '레이크팰리스', '잠실주공', '올림픽선수']

with open('송파구/check_zero.txt', 'w', encoding='utf-8') as f:
    for kw in big:
        sub = df[df['건물명'].str.contains(kw, na=False)]
        f.write(f"\n=== {kw} ({len(sub)}건) ===\n")
        for _, row in sub.head(5).iterrows():
            f.write(f"건물명={row['건물명']}, 용적률={row['용적률(%)']}, 연면적={row['연면적(㎡)']}, 대지면적={row['대지면적(㎡)']}\n")
