# -*- coding: utf-8 -*-
"""
네이버 블로그 자동 포스팅 스크립트 — 위정11 전용
버그 수정 버전 v2:
  [FIX1] 카테고리 선택 — JS click → ActionChains native click (React 합성이벤트 트리거)
  [FIX2] 배경 이미지 업로드 — SE3 배경 버튼 탐색 + hidden input 직접 send_keys
"""
import sys
import time
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# ───────── 위정11 설정 ─────────
TITLE   = "논어 위정11편 — [제목 입력 필요] | 아고산 아침고전산책"
CONTENT = """안녕하세요, 아고산 아침고전산책입니다.

위정편 11번째 에피소드를 올렸습니다.

[위정11 본문 내용 입력]

▶ https://www.youtube.com/watch?v=[YouTube_ID]

#논어 #고전 #아침고전산책 #아고산 #인문학 #논어위정편 #동양고전"""

SCHEDULE_DAYS_AHEAD = 1   # 내일 오전 5시 KST  ← 실행 날짜 확인 후 수정
SCHEDULE_HOUR   = 5
SCHEDULE_MINUTE = 0
PUBLISH_NOW     = False
CATEGORY        = '아침고전산책'
BANNER_FILE     = r'C:\Users\한나\OneDrive\아고산\배너.png'
# ─────────────────────────────

AUTO_PROFILE_DIR = r'C:\Users\한나\AppData\Local\agoSan\chrome-profile2'
DEBUG_DIR = r'C:\Users\한나\OneDrive\아고산'


def get_driver():
    options = ChromeOptions()
    options.add_argument(f'--user-data-dir={AUTO_PROFILE_DIR}')
    options.add_argument('--profile-directory=Default')
    options.add_argument('--start-maximized')
    options.add_argument('--no-first-run')
    options.add_argument('--no-default-browser-check')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--disable-session-crashed-bubble')
    options.add_argument('--hide-crash-restore-bubble')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    return driver


def upload_banner(driver, banner_file, debug_dir):
    """
    [FIX2] SE3 배경 이미지 업로드
    1. SE3 배경 이미지 버튼 탐색 → 클릭 (hidden input 노출)
    2. 모든 file input 목록 출력 (디버그)
    3. image/* accept input 또는 첫 번째 file input에 send_keys
    """
    print("배경 이미지 업로드...")
    banner_ok = False
    time.sleep(0.5)

    try:
        # 현재 file input 현황 파악
        inputs_before = driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]')
        print(f"  [배경 디버그] 초기 file input 수: {len(inputs_before)}")
        for i, fi in enumerate(inputs_before[:5]):
            try:
                accept = fi.get_attribute('accept') or ''
                id_    = fi.get_attribute('id') or ''
                name_  = fi.get_attribute('name') or ''
                cls_   = (fi.get_attribute('class') or '')[:40]
                print(f"    [before[{i}]] accept='{accept}' id='{id_}' name='{name_}' cls='{cls_}'")
            except Exception:
                pass

        # 제목 영역 호버 → SE3 배경 버튼 노출 시도
        try:
            title_el = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
            ActionChains(driver).move_to_element(title_el).perform()
            time.sleep(0.4)
        except Exception as e:
            print(f"  [배경 디버그] 제목 호버 실패: {e}")

        # SE3 배경 이미지 버튼 후보 셀렉터 (다양하게 시도)
        bg_selectors = [
            '[class*="document_bg"] button',
            '[class*="se-module-document_bg"] button',
            'button[class*="bg_btn"]',
            'button[class*="bg_image"]',
            'button[class*="cover_btn"]',
            '[data-module-type*="bg"] button',
            'button[aria-label*="배경"]',
            'button[aria-label*="커버"]',
            'button[aria-label*="이미지"]',
            '.se-module-document_bg',
            '[class*="bgArea"] button',
        ]

        bg_btn_clicked = False
        for sel in bg_selectors:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                for el in els:
                    cls_ = (el.get_attribute('class') or '')[:50]
                    aria_ = el.get_attribute('aria-label') or ''
                    visible = el.is_displayed()
                    print(f"  [배경 디버그] 후보({sel}): cls='{cls_}' aria='{aria_}' visible={visible}")
                    # 표시되거나 숨겨진 버튼 클릭 시도
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", el)
                        ActionChains(driver).move_to_element(el).click().perform()
                        time.sleep(0.8)
                        bg_btn_clicked = True
                        print(f"  [배경] 버튼 클릭 성공: {sel}")
                        break
                    except Exception as click_e:
                        # 안 보이면 JS click 시도
                        try:
                            driver.execute_script("arguments[0].click();", el)
                            time.sleep(0.8)
                            bg_btn_clicked = True
                            print(f"  [배경] 버튼 JS 클릭: {sel}")
                            break
                        except Exception:
                            pass
            except Exception:
                pass
            if bg_btn_clicked:
                break

        if not bg_btn_clicked:
            print("  [배경 디버그] 배경 이미지 버튼 못 찾음 — file input 직접 시도")

        # 버튼 클릭 후 file input 재탐색
        inputs_after = driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]')
        print(f"  [배경 디버그] 버튼 클릭 후 file input 수: {len(inputs_after)}")
        for i, fi in enumerate(inputs_after[:8]):
            try:
                accept = fi.get_attribute('accept') or ''
                id_    = fi.get_attribute('id') or ''
                name_  = fi.get_attribute('name') or ''
                cls_   = (fi.get_attribute('class') or '')[:40]
                print(f"    [after[{i}]] accept='{accept}' id='{id_}' name='{name_}' cls='{cls_}'")
            except Exception:
                pass

        # send_keys 시도 (hidden input도 강제 표시 후 업로드)
        for i, fi in enumerate(inputs_after):
            try:
                accept = fi.get_attribute('accept') or ''
                # hidden input을 보이게 만들기
                driver.execute_script("""
                    var el = arguments[0];
                    el.style.display   = 'block';
                    el.style.visibility = 'visible';
                    el.style.opacity   = '1';
                    el.style.width     = '1px';
                    el.style.height    = '1px';
                """, fi)
                fi.send_keys(banner_file)
                time.sleep(2)
                banner_ok = True
                print(f"  [완료] 배경 이미지 업로드 성공 (file_input[{i}], accept='{accept}')")
                break
            except Exception as e:
                print(f"  [배경 디버그] file_input[{i}] send_keys 실패: {e}")

    except Exception as e:
        print(f"  [경고] 배경 이미지 업로드 전체 오류: {e}")
        import traceback
        traceback.print_exc()

    if not banner_ok:
        print("  [경고] 배경 이미지 업로드 실패 → 수동으로 배너.png 추가 필요")
        try:
            driver.save_screenshot(f'{debug_dir}\\debug_banner_fail.png')
        except Exception:
            pass
    return banner_ok


def select_category(driver, category):
    """
    [FIX1] 카테고리 선택 — ActionChains native click (JS click 대신)
    React 컴포넌트는 JS .click() 이 아닌 실제 마우스 이벤트가 필요함
    """
    if not category:
        return
    try:
        print(f"카테고리 선택: '{category}'...")
        # 셀렉트박스 열기 — Selenium native click
        cat_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[class*="selectbox_button"]')))
        ActionChains(driver).move_to_element(cat_btn).click().perform()
        time.sleep(1.5)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_cat_dropdown.png')

        # 드롭다운 항목 탐색 — native click (JS click 금지)
        cat_ok = False
        item_selectors = [
            '[class*="selectbox_item"]',
            '[class*="category_item"]',
            '[role="option"]',
            '[role="listitem"]',
            'li',
        ]
        for sel in item_selectors:
            try:
                items = driver.find_elements(By.CSS_SELECTOR, sel)
                for item in items:
                    try:
                        txt = item.text.strip()
                        if txt == category and item.is_displayed():
                            ActionChains(driver).move_to_element(item).click().perform()
                            cat_ok = True
                            print(f"  [완료] 카테고리 선택: {category} (셀렉터: {sel})")
                            break
                    except Exception:
                        pass
            except Exception:
                pass
            if cat_ok:
                break

        # fallback: XPATH 텍스트 일치 → native click
        if not cat_ok:
            els = driver.find_elements(
                By.XPATH, f"//*[normalize-space(text())='{category}']")
            for el in els:
                try:
                    if el.is_displayed():
                        ActionChains(driver).move_to_element(el).click().perform()
                        cat_ok = True
                        print(f"  [완료] 카테고리 (XPath fallback): {category}")
                        break
                except Exception:
                    pass

        if not cat_ok:
            print(f"  [경고] 카테고리 '{category}' 드롭다운에서 찾지 못함 — 드롭다운 스크린샷 확인")
        time.sleep(0.5)
    except Exception as e:
        print(f"  [경고] 카테고리 선택 오류: {e}")


def post_blog():
    print("Chrome 시작 중...")
    driver = get_driver()
    wait = WebDriverWait(driver, 20)

    try:
        def do_login():
            print("=" * 50)
            print("로그인 필요 — 브라우저에서 네이버 로그인해주세요!")
            print("(5분 안에 로그인하면 이후 자동으로 진행됩니다)")
            print("=" * 50)
            driver.get("https://nid.naver.com/nidlogin.login")
            for i in range(60):
                time.sleep(5)
                cur = driver.current_url
                if "nidlogin" not in cur:
                    print("  로그인 확인됨!")
                    return True
                if i % 6 == 0:
                    remaining = (60 - i - 1) * 5
                    print(f"  대기 중... {remaining}초 남음")
            return False

        print("블로그 글쓰기 이동...")
        original_handles = driver.window_handles
        driver.get("https://blog.naver.com/PostWriteForm.naver?blogId=danmi0709")
        time.sleep(5)

        cur = driver.current_url
        if "nidlogin" in cur:
            if not do_login():
                driver.quit()
                raise RuntimeError("로그인 실패")
            time.sleep(2)
            driver.get("https://blog.naver.com/PostWriteForm.naver?blogId=danmi0709")
            time.sleep(5)

        all_handles = driver.window_handles
        if len(all_handles) > 1:
            new_window = [h for h in all_handles if h not in original_handles]
            if new_window:
                driver.switch_to.window(new_window[-1])
                time.sleep(2)

        print(f"  현재 URL: {driver.current_url}")

        # 팝업 닫기 — '새 글' 우선 (취소는 에디터를 닫을 수 있음)
        time.sleep(2)
        dismissed_by_js = driver.execute_script("""
            var btns = document.querySelectorAll('button');
            var priority = ['새 글', '새글', '취소'];
            for (var p = 0; p < priority.length; p++) {
                for (var i = 0; i < btns.length; i++) {
                    var t = btns[i].textContent.trim();
                    if (t === priority[p] && btns[i].offsetParent !== null) {
                        btns[i].click();
                        return t;
                    }
                }
            }
            return null;
        """)
        if dismissed_by_js:
            print(f"  [완료] 팝업 닫기: {dismissed_by_js}")
            time.sleep(3)

        time.sleep(3)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_editor_loaded.png')
        print(f"  [디버그] 에디터 URL: {driver.current_url}")

        # ─── 제목 입력 ───────────────────────────────────────
        print("제목 입력...")
        title_ok = False
        PLACEHOLDER_TEXTS = {'제목', '', 'Title'}

        # contenteditable 요소 디버그
        try:
            all_ce = driver.find_elements(By.CSS_SELECTOR, '[contenteditable]')
            print(f"  [디버그] contenteditable 요소 수: {len(all_ce)}")
            for i, ce in enumerate(all_ce[:5]):
                print(f"    [{i}] cls={ce.get_attribute('class') or '(없음)':30s} "
                      f"ce={ce.get_attribute('contenteditable')} "
                      f"text='{(ce.text or '').strip()[:20]}'")
        except Exception as e:
            print(f"  [디버그] contenteditable 확인 실패: {e}")

        # 방법 1: triple_click + send_keys
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.se-title-text')))
            print(f"  [디버그] .se-title-text: ce={el.get_attribute('contenteditable')} "
                  f"text='{(el.text or '').strip()[:30]}'")
            ActionChains(driver)\
                .triple_click(el)\
                .pause(0.4)\
                .send_keys(TITLE)\
                .perform()
            time.sleep(0.8)
            el2 = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
            actual = driver.execute_script("return arguments[0].textContent || '';", el2)
            actual_stripped = (actual or '').strip()
            print(f"  [디버그] 방법1(triple_click) 후: '{actual_stripped[:55]}'")
            if actual_stripped and actual_stripped not in PLACEHOLDER_TEXTS:
                title_ok = True
                print(f"  [완료] 제목 입력 (triple_click)")
        except Exception as e:
            print(f"  [디버그] 방법1 실패: {e}")

        # 방법 2: Ctrl+A + send_keys
        if not title_ok:
            try:
                el = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
                ActionChains(driver)\
                    .click(el).pause(0.3)\
                    .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
                    .pause(0.2)\
                    .send_keys(TITLE)\
                    .perform()
                time.sleep(0.8)
                el2 = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
                actual = driver.execute_script("return arguments[0].textContent || '';", el2)
                actual_stripped = (actual or '').strip()
                print(f"  [디버그] 방법2(Ctrl+A) 후: '{actual_stripped[:55]}'")
                if actual_stripped and actual_stripped not in PLACEHOLDER_TEXTS:
                    title_ok = True
                    print(f"  [완료] 제목 입력 (Ctrl+A)")
            except Exception as e:
                print(f"  [디버그] 방법2 실패: {e}")

        # 방법 3: execCommand (마지막 수단)
        if not title_ok:
            try:
                el = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
                driver.execute_script("""
                    var el = arguments[0]; var text = arguments[1];
                    el.focus();
                    document.execCommand('selectAll', false, null);
                    document.execCommand('delete', false, null);
                    document.execCommand('insertText', false, text);
                """, el, TITLE)
                time.sleep(0.5)
                el2 = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
                actual = driver.execute_script("return arguments[0].textContent || '';", el2)
                actual_stripped = (actual or '').strip()
                if actual_stripped and actual_stripped not in PLACEHOLDER_TEXTS:
                    title_ok = True
                    print(f"  [완료] 제목 입력 (execCommand)")
            except Exception as e:
                print(f"  [디버그] 방법3 실패: {e}")

        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_title.png')
        if not title_ok:
            raise RuntimeError("제목 입력 실패")

        time.sleep(2)

        # ─── [FIX2] 배경 이미지 업로드 ──────────────────────
        upload_banner(driver, BANNER_FILE, DEBUG_DIR)
        time.sleep(1)

        # ─── 본문 입력 ───────────────────────────────────────
        print("본문 입력...")
        body_input_ok = False
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_before_body.png')

        # iframe 구조 진단
        try:
            all_iframes = driver.find_elements(By.TAG_NAME, 'iframe')
            print(f"  [진단] 총 iframe 수: {len(all_iframes)}")
            for fi, iframe in enumerate(all_iframes[:6]):
                try:
                    driver.switch_to.frame(iframe)
                    info = driver.execute_script("""
                        return {
                            ce:   document.querySelectorAll('[contenteditable]').length,
                            p:    document.querySelectorAll('p').length,
                            text: ((document.body||{}).innerText||'').trim().substring(0,25)
                        };
                    """)
                    print(f"  [iframe[{fi}]] ce:{info['ce']} p:{info['p']} text:'{info['text']}'")
                    driver.switch_to.default_content()
                except Exception as e:
                    print(f"  [iframe[{fi}]] 오류: {e}")
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
        except Exception as e:
            print(f"  [진단 실패] {e}")

        # 방법 A: 제목→Ctrl+End→Tab→Ctrl+End→Enter→CONTENT 입력 (핵심 수정: Tab 전 Ctrl+End)
        try:
            title_el = driver.find_element(By.CSS_SELECTOR, '.se-title-text')
            ActionChains(driver)\
                .click(title_el)\
                .pause(0.3)\
                .key_down(Keys.CONTROL).send_keys(Keys.END).key_up(Keys.CONTROL)\
                .pause(0.3)\
                .send_keys(Keys.TAB)\
                .pause(0.6)\
                .key_down(Keys.CONTROL).send_keys(Keys.END).key_up(Keys.CONTROL)\
                .pause(0.3)\
                .send_keys(Keys.ENTER)\
                .pause(0.5)\
                .send_keys(CONTENT)\
                .perform()
            time.sleep(0.8)
            active = driver.execute_script(
                "var el=document.activeElement; "
                "return {tag:el.tagName, cls:(el.className||'').substring(0,30), "
                "ce:el.getAttribute('contenteditable')};")
            print(f"  [진단] Enter 후 활성요소: {active}")
            body_input_ok = True
            print("  [완료] 본문 입력 (Tab→Ctrl+End→Enter)")
        except Exception as e:
            print(f"  [방법A 실패] {e}")

        # 방법 B: iframe 탐색 (iframe[0]=제목 건너뜀)
        if not body_input_ok:
            try:
                all_iframes = driver.find_elements(By.TAG_NAME, 'iframe')
                for fi, iframe in enumerate(all_iframes[:6]):
                    if fi == 0:
                        print(f"  [방법B] iframe[0] 건너뜀 (제목 영역)")
                        continue
                    try:
                        driver.switch_to.frame(iframe)
                        info = driver.execute_script("""
                            var ce = document.querySelectorAll('[contenteditable]').length;
                            var p  = document.querySelectorAll('p').length;
                            return {ce:ce, p:p};
                        """)
                        if info['ce'] > 0:
                            driver.execute_script("""
                                var ce = document.querySelector('[contenteditable]');
                                ce.focus();
                                var r=document.createRange();
                                r.selectNodeContents(ce); r.collapse(false);
                                window.getSelection().removeAllRanges();
                                window.getSelection().addRange(r);
                            """)
                            driver.execute_script(
                                "document.execCommand('insertText', false, arguments[0]);", CONTENT)
                            body_input_ok = True
                            print(f"  [완료] 본문 입력 (방법B iframe[{fi}])")
                            break
                        driver.switch_to.default_content()
                    except Exception as e:
                        print(f"  [방법B] iframe[{fi}] 오류: {e}")
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
            except Exception as e:
                print(f"  [방법B 실패] {e}")

        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_body.png')
        if not body_input_ok:
            print("  [경고] 본문 입력 요소 미발견")
        time.sleep(1.5)

        # ─── 발행 버튼 ───────────────────────────────────────
        print("발행 버튼 클릭...")
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        pub_ok = False
        try:
            btns = driver.find_elements(By.TAG_NAME, 'button')
            for btn in btns:
                if btn.text.strip() == '발행':
                    driver.execute_script("arguments[0].click();", btn)
                    pub_ok = True
                    print("  [완료] 발행 버튼 클릭")
                    break
        except Exception:
            pass

        if not pub_ok:
            raise RuntimeError("발행 버튼을 찾을 수 없습니다.")

        time.sleep(2)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_panel_open.png')

        # ─── [FIX1] 카테고리 선택 ────────────────────────────
        select_category(driver, CATEGORY)

        # ─── 예약 발행 설정 ───────────────────────────────────
        print("예약 발행 설정...")
        tomorrow = datetime.datetime.now() + datetime.timedelta(days=SCHEDULE_DAYS_AHEAD)
        sched_date     = tomorrow.strftime('%Y.%m.%d')
        target_month   = f"{tomorrow.month:02d}"
        target_day_str = str(tomorrow.day)
        target_day_zp  = f"{tomorrow.day:02d}"
        print(f"  목표 날짜: {sched_date} 오전 {SCHEDULE_HOUR}시")

        # 예약 클릭
        reserve_clicked = False
        for xpath in ["//*[normalize-space(text())='예약']", "//label[contains(text(),'예약')]"]:
            try:
                el = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath)))
                driver.execute_script("arguments[0].click();", el)
                reserve_clicked = True
                print("  [완료] 예약 선택")
                break
            except Exception:
                pass

        if not reserve_clicked:
            print("  [경고] 예약 선택 실패")

        time.sleep(4)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_reserve.png')

        # ─── 날짜 설정 (jQuery UI datepicker) ────────────────
        date_ok = False
        sched_date_display = f"{tomorrow.year}. {tomorrow.month:02d}. {tomorrow.day:02d}"

        try:
            date_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[class*="input_date"]')))
            driver.execute_script("arguments[0].click();", date_input)
            time.sleep(1.5)

            datepicker = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, '.ui-datepicker')))
            print("  [달력] jQuery UI datepicker 열림")

            target_year_str  = str(tomorrow.year)
            target_month_ko  = f"{tomorrow.month}월"
            for nav_attempt in range(4):
                try:
                    title_text = datepicker.find_element(
                        By.CSS_SELECTOR, '.ui-datepicker-title').text
                    print(f"  [달력] 표시 월: {title_text}")
                    if target_year_str in title_text and target_month_ko in title_text:
                        break
                    next_btn = datepicker.find_element(By.CSS_SELECTOR, '.ui-datepicker-next')
                    driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(0.5)
                except Exception as nav_e:
                    print(f"  [달력] 월 이동 실패: {nav_e}")
                    break

            clicked = driver.execute_script(f"""
                var dp = document.querySelector('.ui-datepicker, #ui-datepicker-div');
                if (!dp) return 'datepicker not found';
                // 1) tbody button
                var btns = dp.querySelectorAll('tbody button, table button');
                for (var i = 0; i < btns.length; i++) {{
                    if (btns[i].textContent.trim() === '{target_day_str}') {{
                        btns[i].click();
                        return 'clicked tbody button';
                    }}
                }}
                // 2) tbody td a
                var links = dp.querySelectorAll('tbody td a, table td a');
                for (var i = 0; i < links.length; i++) {{
                    if (links[i].textContent.trim() === '{target_day_str}') {{
                        links[i].click();
                        return 'clicked td a';
                    }}
                }}
                // 3) td 직접
                var tds = dp.querySelectorAll('tbody td, table td');
                for (var i = 0; i < tds.length; i++) {{
                    if (tds[i].textContent.trim() === '{target_day_str}') {{
                        tds[i].click();
                        return 'clicked td';
                    }}
                }}
                return 'day not found';
            """)
            time.sleep(0.7)
            actual = date_input.get_attribute('value')
            print(f"  [달력] {clicked} → 날짜값: '{actual}'")
            date_ok = (target_month in actual and target_day_zp in actual)

        except Exception as e:
            print(f"  [경고] 날짜 입력 실패: {e}")

        if not date_ok:
            print(f"  [경고] 날짜 변경 실패 — 목표: '{sched_date_display}'")

        # ─── 시간 설정 ───────────────────────────────────────
        try:
            from selenium.webdriver.support.ui import Select as SeSelect
            hour_sel = driver.find_element(By.CSS_SELECTOR, 'select[class*="hour_option"]')
            opts = [o.get_attribute('value') for o in hour_sel.find_elements(By.TAG_NAME, 'option')]
            val_h = str(SCHEDULE_HOUR)
            if val_h not in opts:
                val_h = f'{SCHEDULE_HOUR:02d}'
            SeSelect(hour_sel).select_by_value(val_h)
            print(f"  [완료] 시간: {SCHEDULE_HOUR}시")
        except Exception as e:
            print(f"  [경고] 시간 select 실패: {e}")

        try:
            min_sel = driver.find_element(By.CSS_SELECTOR, 'select[class*="minute_option"]')
            opts = [o.get_attribute('value') for o in min_sel.find_elements(By.TAG_NAME, 'option')]
            val_m = str(SCHEDULE_MINUTE)
            if val_m not in opts:
                val_m = f'{SCHEDULE_MINUTE:02d}'
            SeSelect(min_sel).select_by_value(val_m)
            print(f"  [완료] 분: {SCHEDULE_MINUTE:02d}분")
        except Exception as e:
            print(f"  [경고] 분 select 실패: {e}")

        time.sleep(0.5)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_datetime_set.png')

        # ─── 예약 확정 ───────────────────────────────────────
        confirmed = False
        url_before = driver.current_url
        print(f"  [확정 전 URL] {url_before}")

        # 보이는 버튼 목록 디버그
        all_btn_info = driver.execute_script("""
            var btns = document.querySelectorAll('button');
            var info = [];
            btns.forEach(function(b) {
                if (b.offsetParent !== null) {
                    info.push(b.textContent.trim().substring(0,15)
                              + '|cls:' + (b.className||'').substring(0,30));
                }
            });
            return info;
        """)
        print(f"  [보이는 버튼] {all_btn_info}")

        # confirm_btn 클래스 (패널 안 최종 발행 확정)
        try:
            confirm_btns = driver.find_elements(
                By.XPATH, "//button[contains(@class,'confirm_btn')]")
            for btn in confirm_btns:
                if btn.is_displayed():
                    print(f"  [확정] confirm_btn: '{btn.text.strip()}'")
                    driver.execute_script("arguments[0].click();", btn)
                    confirmed = True
                    break
        except Exception as e:
            print(f"  [경고] confirm_btn 실패: {e}")

        if not confirmed:
            try:
                btns = driver.find_elements(By.TAG_NAME, 'button')
                candidates = []
                for btn in btns:
                    cls = btn.get_attribute('class') or ''
                    txt = btn.text.strip()
                    if btn.is_displayed() and txt == '발행' and 'publish_btn' not in cls:
                        candidates.append(btn)
                if candidates:
                    driver.execute_script("arguments[0].click();", candidates[-1])
                    confirmed = True
                    print("  [완료] 예약 확정 (publish_btn 제외 발행 버튼)")
            except Exception as e:
                print(f"  [경고] 발행 버튼(publish_btn 제외) 실패: {e}")

        if not confirmed:
            try:
                btns = driver.find_elements(By.TAG_NAME, 'button')
                for btn in btns:
                    txt = btn.text.strip()
                    if btn.is_displayed() and txt in ('예약발행', '확인'):
                        print(f"  [확정 fallback] '{txt}'")
                        driver.execute_script("arguments[0].click();", btn)
                        confirmed = True
                        break
            except Exception as e:
                print(f"  [경고] fallback 버튼 실패: {e}")

        time.sleep(3)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_final_result.png')
        url_after = driver.current_url
        print(f"  [확정 후 URL] {url_after}")

        url_changed   = (url_after != url_before)
        post_success  = url_changed or 'view' in url_after.lower() or 'PostView' in url_after

        if date_ok and confirmed and post_success:
            print(f"\n✅ [완료] 예약 발행 성공! {sched_date} 오전 {SCHEDULE_HOUR}시")
            print(f"   URL: {url_after}")
        elif date_ok and confirmed and not post_success:
            print(f"\n⚠️ [주의] 버튼 클릭됐지만 URL 변경 없음 → debug_final_result.png 확인")
        else:
            print(f"\n⚠️ [경고] 예약 미완료 — date_ok={date_ok}, confirmed={confirmed}")
            print("   → debug_final_result.png 확인해주세요")

    except Exception as e:
        print(f"\n[오류] {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        time.sleep(3)
        driver.quit()
        print("브라우저 종료")


if __name__ == '__main__':
    post_blog()
