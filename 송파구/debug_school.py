# -*- coding: utf-8 -*-
import pandas as pd

df = pd.read_excel('2025년 하반기 교육통계 학교별 일람표(2025.10.1.기준).xlsx',
                    sheet_name='학교별 주요 통계', header=12)

with open('송파구/debug_school.txt', 'w', encoding='utf-8') as f:
    f.write(f'전체 행수: {len(df)}\n')
    f.write(f'컬럼명: {list(df.columns)[:15]}\n\n')
    f.write(f'행정구 unique 샘플: {df["행정구"].dropna().unique()[:30]}\n\n')
    f.write(f'학교급 unique: {df["학교급"].dropna().unique()}\n\n')

    sp = df[(df['행정구'] == '송파구') & (df['학교급'] == '초등학교')]
    f.write(f'송파구 초등학교 수: {len(sp)}\n\n')
    f.write('학교명 + 학생수:\n')
    for _, r in sp.iterrows():
        f.write(f'  {r["학교명"]} - {r["학생수_총계_계"]}\n')
