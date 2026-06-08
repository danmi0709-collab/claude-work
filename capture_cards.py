import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

BASE = Path(r"C:\Users\한나\OneDrive\강성업무용\바탕 화면\문서\CLAUDE\OUTPUTS")

TARGETS = [
    {
        "html": BASE / "20260608_가계부앱_설치안내_카드뉴스.html",
        "png":  BASE / "20260608_가계부앱_설치안내_카드뉴스.png",
        "width": 660,
        "selector": ".card",  # 카드 요소만 캡처
    },
    {
        "html": BASE / "20260608_가계부앱_단톡공지카드.html",
        "png":  BASE / "20260608_가계부앱_단톡공지카드.png",
        "width": 660,
        "selector": ".card",
    },
]

async def capture():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        for t in TARGETS:
            page = await browser.new_page(viewport={"width": t["width"], "height": 1200})
            await page.goto(t["html"].as_uri())
            # 폰트·QR 이미지 완전 로딩 대기
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(1.5)

            # 카드 요소만 정확히 캡처
            el = page.locator(t["selector"])
            await el.screenshot(path=str(t["png"]))
            size = t["png"].stat().st_size // 1024
            print(f"OK  {t['png'].name}  ({size}KB)")
            await page.close()
        await browser.close()

asyncio.run(capture())
print("완료!")
