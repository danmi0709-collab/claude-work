"""
Downloads 폴더 자동 분류기
- 새 파일 감지 시 규칙에 따라 폴더로 이동
- 해당 규칙 없으면 기타PDF/ 폴더로 이동
"""

import os
import shutil
import re

DOWNLOADS = r"C:\Users\한나\Downloads"

# 규칙: (폴더명, 파일명 패턴 리스트)
RULES = [
    ("월부강의교재", ["월부", "열반스쿨", "열중반", "열반실전", "지방투자", "서울투자",
                      "실전준비반", "실준반", "실준_", "재테크기초반", r"\[월부\]",
                      r"★월부", "wb0826"]),
    ("임장보고서",   ["임장보고서", "임보", "단지분석"]),
    ("영어교육",     ["Fun-English", "CK_Mini", "내셔널.*스터디"]),
    ("학교행정서류", ["교육청", "방과후학교", "교육과정", "보험금청구"]),
    ("오행활동지",   ["활동지", "불단어", "화요일_", "수요일_", "목요일_", r"화요일\d"]),
    ("연말정산서류", ["연말정산", r"download-\d{17}"]),
    ("라벨인쇄",     ["라벨_Formtec"]),
]


def get_target_folder(filename: str) -> str:
    lower = filename.lower()
    for folder, patterns in RULES:
        for pat in patterns:
            if re.search(pat, filename, re.IGNORECASE):
                return folder
    return "기타PDF"


def organize():
    moved = []
    for fname in os.listdir(DOWNLOADS):
        src = os.path.join(DOWNLOADS, fname)
        # 폴더 자체는 건너뜀
        if os.path.isdir(src):
            continue
        # PDF만 처리
        if not fname.lower().endswith(".pdf"):
            continue

        target_folder = get_target_folder(fname)
        target_dir = os.path.join(DOWNLOADS, target_folder)
        os.makedirs(target_dir, exist_ok=True)
        dst = os.path.join(target_dir, fname)
        shutil.move(src, dst)
        moved.append(f"{fname} → {target_folder}/")

    if moved:
        log_path = os.path.join(DOWNLOADS, "분류기록.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            from datetime import datetime
            f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}]\n")
            for line in moved:
                f.write(f"  {line}\n")
        print(f"분류 완료: {len(moved)}개")
    else:
        print("새 PDF 없음")


if __name__ == "__main__":
    import time
    while True:
        organize()
        time.sleep(300)  # 5분마다 실행
