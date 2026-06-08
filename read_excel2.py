# -*- coding: utf-8 -*-
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from pathlib import Path
import openpyxl

path = Path(r'C:\Users\한나\OneDrive\월부\202605 실전 송파구\(송파구) 매임 체크리스트.xlsx')
wb = openpyxl.load_workbook(path)
ws = wb.active

# 모든 데이터가 있는 셀 출력 (행 1~50, 열 A~Z)
print("=== 데이터 있는 셀 목록 ===")
for row in ws.iter_rows(min_row=1, max_row=50):
    for cell in row:
        if cell.value is not None and str(cell.value).strip():
            print(f"  {cell.coordinate}: {repr(cell.value)}")
