# -*- coding: utf-8 -*-
import pandas as pd

df = pd.read_excel('C:/Users/한나/Downloads/02. 총괄표제부_2026-05-05 09_20_17.xlsx')

with open('송파구/check_total.txt', 'w', encoding='utf-8') as f:
    f.write(f"총 행 수: {len(df)}\n")
    f.write(f"컬럼 수: {len(df.columns)}\n\n")
    f.write("컬럼 목록:\n")
    for i, c in enumerate(df.columns):
        f.write(f"  {i}: {c}\n")

    # 용적률 확인
    if '용적률(%)' in df.columns:
        f.write(f"\n용적률>0 건수: {(df['용적률(%)']>0).sum()}\n")
        # 잠실엘스, 헬리오시티 확인
        for kw in ['잠실엘스', '헬리오시티', '파크리오', '트리지움', '리센츠', '잠실주공']:
            sub = df[df['건물명'].astype(str).str.contains(kw, na=False)]
            f.write(f"\n=== {kw} ({len(sub)}건) ===\n")
            for _, row in sub.head(3).iterrows():
                f.write(f"건물명={row['건물명']}, 용적률={row['용적률(%)']}, 세대수={row.get('세대수(세대)', 'N/A')}, 사용승인일={row.get('사용승인일','N/A')}\n")
