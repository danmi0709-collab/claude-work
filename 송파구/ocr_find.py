import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import easyocr
import pandas as pd
from PIL import Image
import numpy as np

# 아파트 단지명 목록 로드
df = pd.read_excel('송파구/apartments_all (1).xlsx', engine='openpyxl')
songpa = df[df['시군구'] == '송파구']
units = songpa.groupby(['단지코드','단지명','읍면동','사용승인일','총세대수']).size().reset_index(name='cnt')
units = units[['단지명','읍면동','사용승인일','총세대수']].drop_duplicates().reset_index(drop=True)
apt_names = units['단지명'].tolist()

# 이미지 로드 (너무 크면 축소)
img = Image.open('송파구/송파구이미지2.png').convert('RGB')
W, H = img.size
print(f"원본 크기: {W}x{H}", flush=True)

# OCR용으로 2000px 너비로 축소
SCALE = 2000 / W
small = img.resize((int(W*SCALE), int(H*SCALE)), Image.LANCZOS)
sw, sh = small.size
print(f"OCR용 크기: {sw}x{sh}", flush=True)

# EasyOCR 실행
reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)
print("OCR 시작...", flush=True)
results = reader.readtext(np.array(small), detail=1, paragraph=False)
print(f"OCR 완료: {len(results)}개 텍스트 감지", flush=True)

# 아파트 이름 매칭
found = {}
for bbox, text, conf in results:
    text_clean = text.strip().replace(' ', '')
    for apt in apt_names:
        apt_clean = apt.replace(' ', '')
        # 부분 일치 (3글자 이상)
        if len(apt_clean) >= 3 and apt_clean in text_clean:
            if conf > 0.3:
                # bbox 중심점 계산 (원본 좌표로 변환)
                xs = [p[0] for p in bbox]
                ys = [p[1] for p in bbox]
                cx = (min(xs) + max(xs)) / 2 / SCALE
                cy = (min(ys) + max(ys)) / 2 / SCALE
                if apt not in found or conf > found[apt]['conf']:
                    found[apt] = {'x': cx, 'y': cy, 'conf': conf, 'raw': text}
                    print(f"  매칭: {apt} <- '{text}' (신뢰도:{conf:.2f}) at ({cx:.0f},{cy:.0f})", flush=True)

print(f"\n총 {len(found)}/{len(apt_names)}개 매칭됨")
with open('송파구/ocr_positions.json', 'w', encoding='utf-8') as f:
    json.dump(found, f, ensure_ascii=False, indent=2)
print("저장: 송파구/ocr_positions.json")
