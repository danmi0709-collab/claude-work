# -*- coding: utf-8 -*-
import pandas as pd

df = pd.read_excel('2025년 하반기 교육통계 학교별 일람표(2025.10.1.기준).xlsx',
                    sheet_name='학교별 주요 통계', header=None)

with open('송파구/check_school2.txt', 'w', encoding='utf-8') as f:
    # 헤더 후보 row 9-12
    for r in range(9, 13):
        f.write(f'=== Row {r} ===\n')
        for c, v in enumerate(df.iloc[r]):
            if pd.notna(v):
                f.write(f'  col{c}: {v}\n')
        f.write('\n')

    # 송파구 초등학교 샘플 행 1개
    # 학교명 컬럼 찾기 - 일단 13번째 행부터 데이터
    f.write('=== 데이터 행 13 ===\n')
    for c, v in enumerate(df.iloc[13]):
        if pd.notna(v):
            f.write(f'  col{c}: {v}\n')
