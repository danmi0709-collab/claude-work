# -*- coding: utf-8 -*-
import pandas as pd

xl = pd.ExcelFile('2025년 하반기 교육통계 학교별 일람표(2025.10.1.기준).xlsx')

with open('송파구/check_school.txt', 'w', encoding='utf-8') as f:
    f.write(f'시트 목록: {xl.sheet_names}\n\n')
    for sn in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sn, header=None, nrows=10)
        f.write(f'=== 시트: {sn} ===\n')
        f.write(f'행수(전체): {pd.read_excel(xl, sheet_name=sn, header=None).shape[0]}\n')
        f.write(f'열수: {df.shape[1]}\n')
        f.write('첫 10행:\n')
        f.write(df.to_string() + '\n\n')
