"""
ORT 차트 이미지에서 책 표지를 개별 크롭하는 스크립트
이미지: ortchart_page1.png (5120×3640px)
출력: ort-covers/ort/{levelId}_{catIdx}_{bookIdx}.jpg
"""

from PIL import Image
import os

# ─── 경로 설정 ───────────────────────────────────────────────
IMG_PATH   = r"C:\Users\한나\Downloads\ortchart_page1.png"
OUTPUT_DIR = r"C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE\ort-covers\ort"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── 차트 좌표 (픽셀 분석으로 확인된 값) ─────────────────────

# 14개 행의 y 범위 [y_start, y_end]
ROW_Y = [
    (507,  668),   # row 0
    (723,  884),   # row 1
    (940,  1101),  # row 2
    (1158, 1319),  # row 3
    (1372, 1534),  # row 4
    (1590, 1751),  # row 5
    (1807, 1968),  # row 6
    (2024, 2185),  # row 7
    (2240, 2401),  # row 8
    (2460, 2621),  # row 9
    (2675, 2836),  # row 10
    (2893, 3054),  # row 11
    (3112, 3273),  # row 12
    (3328, 3489),  # row 13 (예비)
]

# 4개 그룹 × 6권 x 시작 좌표
GROUP_X = [
    [209,  375,  541,  708,  874,  1040],  # group 0
    [1451, 1617, 1783, 1949, 2116, 2282],  # group 1
    [2693, 2859, 3025, 3192, 3358, 3524],  # group 2
    [3935, 4103, 4271, 4439, 4607, 4776],  # group 3
]

# 각 그룹 내 책 1권 너비 (다음 x - 현재 x, 마지막 책은 동일 너비)
BOOK_W = 155   # px (안전 마진 포함)
BOOK_H_PAD = 5  # 상하 여백

# ─── ORT 레벨 & 카테고리 정의 ────────────────────────────────
ORT_LEVELS = [
    ('1',   5),   # Level 1: 5 cats
    ('1p',  7),   # Level 1+: 7 cats
    ('2',   6),   # Level 2: 6 cats
    ('3',   6),   # Level 3: 6 cats
    ('4',   6),   # Level 4: 6 cats
    ('5',   6),   # Level 5: 6 cats
    ('6',   3),   # Level 6: 3 cats
    ('7',   3),   # Level 7: 3 cats
    ('8',   2),   # Level 8: 2 cats
    ('9',   2),   # Level 9: 2 cats
    ('10p', 1),   # Level 10+: 1 cat
    ('11p', 1),   # Level 11+: 1 cat
    ('12p', 1),   # Level 12+: 1 cat
]
# 총 49 카테고리

# ─── 메인 크롭 로직 ──────────────────────────────────────────
print(f"Loading: {IMG_PATH}")
img = Image.open(IMG_PATH)
w, h = img.size
print(f"Image size: {w} x {h}")

global_cat_idx = 0  # 전체 카테고리 순번 (0부터)
total_saved = 0
total_skipped = 0

for (level_id, num_cats) in ORT_LEVELS:
    for cat_idx in range(num_cats):
        # 전체 순번 → 차트 위치
        row   = global_cat_idx // 4
        group = global_cat_idx % 4

        if row >= len(ROW_Y):
            print(f"  WARN row {row} out of range, skip (level={level_id}, cat={cat_idx})")
            global_cat_idx += 1
            continue

        y1, y2 = ROW_Y[row]
        y1 += BOOK_H_PAD
        y2 -= BOOK_H_PAD

        for book_idx in range(6):
            x1 = GROUP_X[group][book_idx]
            x2 = x1 + BOOK_W

            # 이미지 경계 클리핑
            x1c = max(0, x1)
            x2c = min(w, x2)
            y1c = max(0, y1)
            y2c = min(h, y2)

            crop = img.crop((x1c, y1c, x2c, y2c))

            fname = f"{level_id}_{cat_idx}_{book_idx}.jpg"
            fpath = os.path.join(OUTPUT_DIR, fname)
            crop.save(fpath, "JPEG", quality=90)
            total_saved += 1

        print(f"  OK Level {level_id} cat[{cat_idx}] -> chart row={row} group={group}  (6 saved)")
        global_cat_idx += 1

print(f"\nDONE! saved={total_saved} / skipped={total_skipped}")
print(f"Output: {OUTPUT_DIR}")
