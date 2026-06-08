from PIL import Image, ImageDraw, ImageFont
import os

W, H = 1200, 630

# ── 색상 ──
BG        = (250, 248, 243)   # 아이보리
WHITE     = (255, 255, 255)
ACCENT    = (192, 122, 58)    # 따뜻한 갈색
ACCENT2   = (226, 180, 90)
TEXT      = (45, 42, 36)
TEXT2     = (122, 116, 104)
TEXT3     = (158, 149, 133)
GREEN     = (46, 125, 79)
RED       = (192, 57, 43)
BLUE      = (36, 113, 163)
CARD_BG   = (244, 241, 235)
BORDER    = (232, 227, 216)

img  = Image.new('RGB', (W, H), BG)
draw = ImageDraw.Draw(img)

# ── 폰트 (시스템 폰트 탐색) ──
def find_font(size):
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",        # 맑은 고딕
        "C:/Windows/Fonts/malgunbd.ttf",
        "C:/Windows/Fonts/NanumGothic.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()

f_xl   = find_font(56)
f_lg   = find_font(36)
f_md   = find_font(26)
f_sm   = find_font(20)
f_xs   = find_font(17)

def rounded_rect(draw, xy, r, fill, outline=None, width=1):
    x0, y0, x1, y1 = xy
    draw.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=fill,
                           outline=outline, width=width)

# ── 상단 그라데이션 바 ──
for i in range(12):
    t = i / 11
    r = int(192 + t * (226 - 192))
    g = int(122 + t * (180 - 122))
    b = int(58  + t * (90  - 58))
    draw.rectangle([i * (W // 12), 0, (i+1) * (W // 12), 10], fill=(r, g, b))

# ── 배경 원형 장식 ──
draw.ellipse([(-120, 360), (340, 780)], fill=(232, 227, 216))
draw.ellipse([(960, -80), (1300, 260)], fill=(232, 227, 216))

# ── 왼쪽 메인 카드 ──
rounded_rect(draw, [60, 50, 760, 560], r=24, fill=WHITE)

# 아이콘 배경
rounded_rect(draw, [92, 86, 164, 158], r=16, fill=(255, 248, 238))
draw.text((128, 122), "💰", font=find_font(44), fill=TEXT, anchor="mm")

# 타이틀
draw.text((188, 122), "나만의 가계부", font=f_xl, fill=TEXT, anchor="lm")

# 구분선
draw.rectangle([92, 170, 728, 172], fill=BORDER)

# 설명
draw.text((92, 204), "브라우저에서 바로 쓰는 스마트 가계부", font=f_md, fill=TEXT2)
draw.text((92, 244), "설치 없이 · 무료 · 내 데이터는 내 기기에만 저장", font=f_md, fill=TEXT2)

# ── 기능 칩 ──
chips = [
    ("📋 내역 기록", (240, 247, 244), GREEN, 92),
    ("📊 분석",      (255, 248, 238), ACCENT, 284),
    ("🏦 자산",      (237, 244, 251), BLUE,   424),
    ("🔄 구독",      (253, 240, 239), RED,    92),
    ("📅 연예산",    (244, 241, 235), TEXT2,  252),
]
row1 = [(chips[0]), (chips[1]), (chips[2])]
row2 = [(chips[3]), (chips[4])]

def chip(draw, x, y, label, bg, fg, font):
    tw = font.getlength(label)
    pw = 24
    w = int(tw) + pw * 2
    h = 44
    rounded_rect(draw, [x, y, x + w, y + h], r=22, fill=bg)
    draw.text((x + pw + tw // 2, y + h // 2), label, font=font, fill=fg, anchor="mm")
    return w

cx = 92
for label, bg, fg, _ in row1:
    w = chip(draw, cx, 292, label, bg, fg, f_sm)
    cx += w + 14

cx = 92
for label, bg, fg, _ in row2:
    w = chip(draw, cx, 352, label, bg, fg, f_sm)
    cx += w + 14

# 구분선
draw.rectangle([92, 414, 728, 415], fill=BORDER)

# 태그라인
draw.text((92, 446), "무료 · 광고 없음 · 개인정보 수집 없음", font=f_xs, fill=TEXT3)

# CTA 버튼
rounded_rect(draw, [92, 474, 316, 534], r=14, fill=ACCENT)
draw.text((204, 504), "지금 시작하기  →", font=find_font(22), fill=WHITE, anchor="mm")

# ── 오른쪽 미니 앱 화면 ──
rounded_rect(draw, [800, 70, 1140, 560], r=22, fill=(255, 255, 255, 180))
# 상단 바
draw.rectangle([800, 70, 1140, 80], fill=ACCENT)
rounded_rect(draw, [800, 70, 1140, 120], r=22, fill=(255, 248, 238))
draw.text((970, 95), "나만의 가계부", font=find_font(20), fill=TEXT, anchor="mm")

# 잔액 카드
rounded_rect(draw, [820, 130, 1120, 226], r=14, fill=BG)
draw.text((970, 162), "이번 달 잔액", font=f_xs, fill=TEXT3, anchor="mm")
draw.text((970, 200), "+1,240,000원", font=find_font(30), fill=ACCENT, anchor="mm")

# 내역 목록
items = [
    ("점심 식사",  "−12,000",    RED),
    ("급여",       "+3,500,000", GREEN),
    ("적금",       "−400,000",   BLUE),
]
y = 240
for name, amt, color in items:
    draw.rectangle([820, y, 1120, y + 1], fill=BORDER)
    draw.text((830, y + 20), name, font=f_xs, fill=TEXT, anchor="lm")
    draw.text((1110, y + 20), amt, font=f_xs, fill=color, anchor="rm")
    y += 44

# 목표 바
draw.rectangle([820, y + 14, 1120, y + 15], fill=BORDER)
draw.text((830, y + 30), "🎯 순자산 목표  67%", font=f_xs, fill=TEXT3, anchor="lm")
y += 50
rounded_rect(draw, [830, y, 1110, y + 10], r=5, fill=CARD_BG)
rounded_rect(draw, [830, y, 830 + int(280 * 0.67), y + 10], r=5, fill=ACCENT)

# 하단 탭바 미니
rounded_rect(draw, [800, 510, 1140, 560], r=0, fill=CARD_BG)
tabs = ["내역", "분석", "카테고리", "연예산", "자산"]
for i, t in enumerate(tabs):
    x = 820 + i * 64
    col = ACCENT if t == "카테고리" else TEXT3
    draw.text((x + 20, 535), t, font=find_font(13), fill=col, anchor="mm")

# ── 하단 URL ──
draw.text((W // 2, 600), "danmi0709-collab.github.io/gagyebu/share",
          font=f_xs, fill=TEXT3, anchor="mm")

# 저장
out = r"C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE\gagyebu\og-image.png"
img.save(out, "PNG", optimize=True)
print(f"저장 완료: {out} ({os.path.getsize(out)//1024}KB)")
