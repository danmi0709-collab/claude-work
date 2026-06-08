from PIL import Image, ImageDraw, ImageFont
import os

def find_font(size):
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/malgunbd.ttf",
        "C:/Windows/Fonts/NanumGothic.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()

def make_icon(size):
    img = Image.new('RGB', (size, size), (192, 122, 58))  # 갈색 배경
    draw = ImageDraw.Draw(img)

    # 배경 원형 장식
    margin = size * 0.08
    draw.ellipse([margin, margin, size - margin, size - margin],
                 fill=(226, 180, 90))

    # 안쪽 원
    m2 = size * 0.18
    draw.ellipse([m2, m2, size - m2, size - m2],
                 fill=(192, 122, 58))

    # 💰 텍스트
    emoji_size = int(size * 0.42)
    try:
        font = find_font(emoji_size)
        draw.text((size // 2, size // 2), "💰",
                  font=font, fill=(255, 255, 255), anchor="mm")
    except Exception:
        # 폴백: 한글
        font = find_font(int(size * 0.3))
        draw.text((size // 2, size // 2 - size * 0.06), "가계부",
                  font=font, fill=(255, 255, 255), anchor="mm")

    return img

out_dir = r"C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE\gagyebu\share"

for sz in [192, 512]:
    icon = make_icon(sz)
    path = os.path.join(out_dir, f"icon-{sz}.png")
    icon.save(path, "PNG")
    print(f"OK icon-{sz}.png saved ({os.path.getsize(path)} bytes)")

print("icons done")
