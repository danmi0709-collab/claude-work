import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import easyocr
import pandas as pd
from PIL import Image
import numpy as np

df = pd.read_excel('송파구/apartments_all (1).xlsx', engine='openpyxl')
songpa = df[df['시군구'] == '송파구']
units = songpa.groupby(['단지코드','단지명','읍면동','사용승인일','총세대수']).size().reset_index(name='cnt')
units = units[['단지명','읍면동','사용승인일','총세대수']].drop_duplicates().reset_index(drop=True)
apt_names = units['단지명'].tolist()

# 이미지 중앙 부분 잘라서 테스트 (잠실 지역)
img = Image.open('송파구/송파구이미지2.png').convert('RGB')
W, H = img.size

# 잠실 지역 (이미지 왼쪽 중간 부분) 크롭
# 잠실은 대략 왼쪽 25~50%, 위쪽 40~65%
crop = img.crop((int(W*0.1), int(H*0.35), int(W*0.55), int(H*0.65)))
crop_w, crop_h = crop.size
print(f"크롭 크기: {crop_w}x{crop_h}")

# 1500px로 축소
scale = 1500 / crop_w
small = crop.resize((int(crop_w*scale), int(crop_h*scale)), Image.LANCZOS)
print(f"OCR 크기: {small.size}")

reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)
results = reader.readtext(np.array(small), detail=1, paragraph=False)
print(f"\n감지된 텍스트 ({len(results)}개):")
for bbox, text, conf in sorted(results, key=lambda x: -x[2])[:80]:
    print(f"  [{conf:.2f}] {text}")
