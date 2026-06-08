# -*- coding: utf-8 -*-
"""
네이버 블로그 자동 포스팅 스크립트 — 위정10 전용
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

# ───────── 위정10 설정 ─────────
TITLE   = "논어 위정10편 — 당신을 꿰뚫어보는 3가지 질문 | 아고산 아침고전산책"
CONTENT = """안녕하세요, 아고산 아침고전산책입니다.

위정편 10번째 에피소드를 올렸습니다.

공자는 말했습니다. 어떻게 행동하는지, 왜 행동하는지, 무엇에서 안정감을 찾는지. 이 세 가지를 보면 어떤 사람이든 본심을 숨길 수 없습니다.

▶ https://www.youtube.com/watch?v=F61l__2BqqU

#논어 #고전 #아침고전산책 #아고산 #인문학 #논어위정편 #동양고전"""

SCHEDULE_DAYS_AHEAD = 2   # 모레(5/30) 오전 5시 KST
SCHEDULE_HOUR   = 5
SCHEDULE_MINUTE = 0
PUBLISH_NOW     = False
CATEGORY        = '아침고전산책'
BANNER_FILE     = r'C:\Users\한나\OneDrive\아고산\배너.png'
# ─────────────────────────────

AUTO_PROFILE_DIR = r'C:\Users\한나\AppData\Local\agoSan\chrome-profile2'
DEBUG_DIR = r'C:\Users\한나\OneDrive\아고산'


def type_text_to_element(driver, element, text):
    """요소에 텍스트 직접 입력 — selectAll+delete+insertText"""
    driver.execute_script("""
        var el = arguments[0];
        var text = arguments[1];
        el.focus();
        document.execCommand('selectAll', false, null);
        document.execCommand('delete', false, null);
        document.execCommand('insertText', false, text);
        if (!el.textContent || el.textContent.trim() === '') {
            el.innerText = text;
        }
    """, element, text)


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
            // 우선순위: '새 글' → '새글' → '취소'
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
            print(f"  [완료] 팝업 닫기 (JS: {dismissed_by_js})")
            time.sleep(3)  # 에디터 재초기화 대기

        # 에디터 완전 로딩 대기
        time.sleep(3)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_editor_loaded.png')
        print(f"  [디버그] 에디터 로딩 후 URL: {driver.current_url}")

        # 제목 입력 — el.send_keys() 직접 호출 (React onChange 확실히 트리거)
        print("제목 입력...")
        title_ok = False
        PLACEHOLDER_TEXTS = {'제목', '', 'Title'}

        # 제목 요소 디버그 (모든 contenteditable 확인)
        try:
            all_ce = driver.find_elements(By.CSS_SELECTOR, '[contenteditable]')
            print(f"  [디버그] contenteditable 요소 수: {len(all_ce)}")
            for i, ce in enumerate(all_ce[:5]):
                print(f"    [{i}] cls={ce.get_attribute('class') or '(없음)':30s} ce={ce.get_attribute('contenteditable')} text='{(ce.text or '').strip()[:20]}'")
        except Exception as e:
            print(f"  [디버그] contenteditable 확인 실패: {e}")

        # 방법 1: triple_click으로 기존 내용 전체선택 → TITLE 입력 (JS DOM 수정 없음 — stale 방지)
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.se-title-text')))
            print(f"  [디버그] .se-title-text: ce={el.get_attribute('contenteditable')} text='{(el.text or '').strip()[:30]}'")
            ActionChains(driver)\
                .triple_click(el)\
                .pause(0.4)\
                .send_keys(TITLE)\
                .perform()
            time.sleep(0.8)
            el2 = driver.find_element(By.CSS_SELECTOR, '.se-title-text')  # 재탐색
            actual = driver.execute_script("return arguments[0].textContent || '';", el2)
            actual_stripped = (actual or '').strip()
            print(f"  [디버그] 방법1(triple_click) 후: '{actual_stripped[:55]}'")
            if actual_stripped and actual_stripped not in PLACEHOLDER_TEXTS:
                title_ok = True
                print(f"  [완료] 제목 입력 (triple_click): {actual_stripped[:55]}")
        except Exception as e:
            print(f"  [디버그] 방법1 실패: {e}")

        # 방법 2: ActionChains key_down(CTRL)+a+key_up → TITLE (JS 수정 없음)
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
                    print(f"  [완료] 제목 입력 (Ctrl+A): {actual_stripped[:55]}")
            except Exception as e:
                print(f"  [디버그] 방법2 실패: {e}")

        # 방법 3: execCommand — 마지막 수단 (DOM 표시용, React state 별개)
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
                print(f"  [디버그] 방법3(execCommand) 후: '{actual_stripped[:55]}'")
                if actual_stripped and actual_stripped not in PLACEHOLDER_TEXTS:
                    title_ok = True
                    print(f"  [완료] 제목 입력 (execCommand): {actual_stripped[:55]}")
            except Exception as e:
                print(f"  [디버그] 방법3 실패: {e}")

        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_title.png')
        if not title_ok:
            raise RuntimeError("제목 입력 실패")

        time.sleep(2)

        # 배경 이미지 업로드
        print("배경 이미지 업로드...")
        fast_wait = WebDriverWait(driver, 3)
        try:
            file_input = fast_wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="file"]')))
            file_input.send_keys(BANNER_FILE)
            time.sleep(2)
            print("  [완료] 배경 이미지 업로드")
        except:
            print("  [경고] 배경 이미지 업로드 실패 (건너뜀)")

        time.sleep(1)

        # 본문 입력 — SE3 본문 영역 탐색 후 입력
        # ※ iframe[0]은 제목 영역이므로 반드시 건너뜀
        print("본문 입력...")
        body_input_ok = False
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_before_body.png')

        # 진단: 전체 iframe 구조 파악
        try:
            all_iframes = driver.find_elements(By.TAG_NAME, 'iframe')
            print(f"  [진단] 총 iframe 수: {len(all_iframes)}")
            for fi, iframe in enumerate(all_iframes[:6]):
                try:
                    driver.switch_to.frame(iframe)
                    info = driver.execute_script("""
                        return {
                            ce: document.querySelectorAll('[contenteditable]').length,
                            p:  document.querySelectorAll('p').length,
                            div: document.querySelectorAll('div').length,
                            text: ((document.body||{}).innerText||'').trim().substring(0,25),
                            ph: (function(){
                                var e=document.querySelector('[data-placeholder]');
                                return e ? e.getAttribute('data-placeholder').substring(0,30) : '';
                            })()
                        };
                    """)
                    print(f"  [iframe[{fi}]] ce:{info['ce']} p:{info['p']} div:{info['div']} text:'{info['text']}' ph:'{info['ph']}'")
                    driver.switch_to.default_content()
                except Exception as e:
                    print(f"  [iframe[{fi}]] 오류: {e}")
                    try: driver.switch_to.default_content()
                    except: pass
        except Exception as e:
            print(f"  [진단 실패] {e}")

        # 방법 A: 제목→Tab→Ctrl+End(제목 끝)→Enter(본문 이동)→CONTENT 입력
        # ※ Tab 후 커서가 제목 중간에 올 수 있으므로 Ctrl+End로 끝으로 이동 후 Enter로 본문 진입
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
                "var el=document.activeElement; return {tag:el.tagName, cls:(el.className||'').substring(0,30), ce:el.getAttribute('contenteditable')};")
            print(f"  [진단] Enter 후 활성요소: {active}")
            body_input_ok = True
            print("  [완료] 본문 입력 (Tab→Ctrl+End→Enter→send_keys)")
        except Exception as e:
            print(f"  [방법A(Tab+Enter)] 실패: {e}")

        # 방법 B: iframe[0] 제외, p>0 이거나 내용 없는 iframe에서 insertText
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
                            var tx = ((document.body||{}).innerText||'').trim();
                            return {ce:ce, p:p, empty:(tx.length===0)};
                        """)
                        print(f"  [방법B] iframe[{fi}] ce:{info['ce']} p:{info['p']} empty:{info['empty']}")
                        if info['ce'] > 0:
                            body_result = driver.execute_script("""
                                var ce = document.querySelector('[contenteditable]');
                                ce.focus();
                                var r=document.createRange(); r.selectNodeContents(ce); r.collapse(false);
                                window.getSelection().removeAllRanges(); window.getSelection().addRange(r);
                                return 'focused: '+ce.tagName+'.'+(ce.className||'').substring(0,20);
                            """)
                            print(f"  [방법B] iframe[{fi}] 포커스: {body_result}")
                            driver.execute_script(
                                "document.execCommand('insertText', false, arguments[0]);", CONTENT)
                            body_input_ok = True
                            print(f"  [완료] 본문 입력 (방법B iframe[{fi}] insertText)")
                            break
                        driver.switch_to.default_content()
                    except Exception as e:
                        print(f"  [방법B] iframe[{fi}] 오류: {e}")
                        try: driver.switch_to.default_content()
                        except: pass
                try: driver.switch_to.default_content()
                except: pass
            except Exception as e:
                print(f"  [방법B 실패] {e}")

        # 방법 C: 메인문서 본문 플레이스홀더 요소 클릭 + send_keys
        if not body_input_ok:
            try:
                body_clicked = driver.execute_script("""
                    var searches = [
                        '[data-placeholder*="기록"]',
                        '[data-placeholder*="글감"]',
                        '.se-placeholder',
                        '.se-text .se-text-paragraph',
                        '.se-component.se-text'
                    ];
                    for (var i=0; i<searches.length; i++) {
                        var el = document.querySelector(searches[i]);
                        if (el) {
                            var ce = el.closest('[contenteditable]') || el;
                            ce.click(); ce.focus();
                            var r=document.createRange(); r.selectNodeContents(ce); r.collapse(false);
                            window.getSelection().removeAllRanges(); window.getSelection().addRange(r);
                            return 'found('+searches[i]+'): '+ce.tagName+'.'+(ce.className||'').substring(0,25);
                        }
                    }
                    // 제목 아래 여러 y 지점 클릭 시도
                    var t = document.querySelector('.se-title-text');
                    if (t) {
                        var rect = t.getBoundingClientRect();
                        for (var dy of [120, 160, 200, 260, 320]) {
                            var el = document.elementFromPoint(rect.left+100, rect.bottom+dy);
                            if (el && el !== t && !t.contains(el)) {
                                el.click(); el.focus();
                                return 'below_title+'+dy+': '+el.tagName+'.'+(el.className||'').substring(0,25);
                            }
                        }
                    }
                    return 'not_found';
                """)
                print(f"  [방법C] 클릭: {body_clicked}")
                time.sleep(0.3)
                ActionChains(driver).send_keys(CONTENT).perform()
                body_input_ok = True
                print("  [완료] 본문 입력 (방법C)")
            except Exception as e:
                print(f"  [방법C 실패] {e}")

        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_body.png')
        if not body_input_ok:
            print("  [경고] 본문 입력 요소 미발견")
        time.sleep(1.5)

        # 발행 버튼
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
        except:
            pass

        if not pub_ok:
            raise RuntimeError("발행 버튼을 찾을 수 없습니다.")

        time.sleep(2)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_panel_open.png')

        # 카테고리 선택
        if CATEGORY:
            try:
                cat_btn = driver.find_element(By.CSS_SELECTOR, 'button[class*="selectbox_button"]')
                driver.execute_script("arguments[0].click();", cat_btn)
                time.sleep(1)
                # 드롭다운에서 카테고리 찾아 클릭
                cat_clicked = driver.execute_script(f"""
                    var items = document.querySelectorAll('[class*="selectbox_item"], [class*="category_item"], li');
                    for (var i=0; i<items.length; i++) {{
                        if (items[i].textContent.trim() === '{CATEGORY}') {{
                            items[i].click();
                            return 'clicked: ' + items[i].textContent.trim();
                        }}
                    }}
                    // 모든 클릭 가능한 요소에서 찾기
                    var btns = document.querySelectorAll('button, a, span[role="option"]');
                    for (var i=0; i<btns.length; i++) {{
                        if (btns[i].textContent.trim() === '{CATEGORY}' && btns[i].offsetParent !== null) {{
                            btns[i].click();
                            return 'clicked btn: ' + btns[i].textContent.trim();
                        }}
                    }}
                    return 'not found: ' + Array.from(items).slice(0,5).map(e=>e.textContent.trim()).join(',');
                """)
                print(f"  [카테고리] {cat_clicked}")
                time.sleep(0.5)
            except Exception as e:
                print(f"  [경고] 카테고리 선택 실패: {e}")

        # 예약 발행 설정
        print("예약 발행 설정...")
        tomorrow = datetime.datetime.now() + datetime.timedelta(days=SCHEDULE_DAYS_AHEAD)
        sched_date = tomorrow.strftime('%Y.%m.%d')
        target_month = f"{tomorrow.month:02d}"
        target_day_str = str(tomorrow.day)
        target_day_zp  = f"{tomorrow.day:02d}"
        print(f"  목표 날짜: {sched_date} (오전 {SCHEDULE_HOUR}시)")

        # 예약 클릭
        reserve_clicked = False
        for xpath in ["//*[normalize-space(text())='예약']", "//label[contains(text(),'예약')]"]:
            try:
                el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                driver.execute_script("arguments[0].click();", el)
                reserve_clicked = True
                print(f"  [완료] 예약 선택")
                break
            except:
                pass

        if not reserve_clicked:
            print("  [경고] 예약 선택 실패")

        time.sleep(4)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_reserve.png')

        # 날짜 입력 — jQuery UI datepicker
        date_ok = False
        sched_date_display = f"{tomorrow.year}. {tomorrow.month:02d}. {tomorrow.day:02d}"

        try:
            # input_date (readonly) 클릭 → jQuery UI 달력 팝업 열기
            date_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[class*="input_date"]')))
            driver.execute_script("arguments[0].click();", date_input)
            time.sleep(1.5)
            driver.save_screenshot(f'{DEBUG_DIR}\\debug_after_date_click.png')

            # .ui-datepicker 팝업 대기
            try:
                datepicker = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, '.ui-datepicker')))
                print(f"  [달력] jQuery UI datepicker 팝업 열림")

                # 현재 표시된 연월 확인 → 필요시 다음 달 이동
                target_year_str = str(tomorrow.year)
                target_month_ko = f"{tomorrow.month}월"
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

                # 달력 HTML tbody 부분 확인 (디버그 — 날짜 셀 구조 파악)
                tbody_html = driver.execute_script("""
                    var tb = document.querySelector('.ui-datepicker tbody, .ui-datepicker-calendar tbody');
                    return tb ? tb.innerHTML.substring(0, 800) : 'tbody not found';
                """)
                print(f"  [달력 tbody HTML]: {tbody_html[:500]}")

                # 목표 날짜 클릭
                found_day = False

                # JS로 datepicker tbody 내 정확한 날짜 클릭
                # 날짜 셀: td 또는 td 안의 button/a 중 텍스트가 정확히 target_day_str인 것
                clicked = driver.execute_script(f"""
                    var dp = document.querySelector('.ui-datepicker, #ui-datepicker-div');
                    if (!dp) return 'datepicker not found';

                    // 1) tbody 내 button 요소 (Naver 커스텀 datepicker)
                    var btns = dp.querySelectorAll('tbody button, table button');
                    for (var i = 0; i < btns.length; i++) {{
                        var txt = btns[i].textContent.trim();
                        if (txt === '{target_day_str}') {{
                            btns[i].click();
                            return 'clicked tbody button: ' + (btns[i].className||'').substring(0,40);
                        }}
                    }}

                    // 2) td 내 a 태그
                    var links = dp.querySelectorAll('tbody td a, table td a');
                    for (var i = 0; i < links.length; i++) {{
                        if (links[i].textContent.trim() === '{target_day_str}') {{
                            links[i].click();
                            return 'clicked tbody td a: ' + (links[i].className||'').substring(0,40);
                        }}
                    }}

                    // 3) td 직접 클릭
                    var tds = dp.querySelectorAll('tbody td, table td');
                    for (var i = 0; i < tds.length; i++) {{
                        if (tds[i].textContent.trim() === '{target_day_str}') {{
                            tds[i].click();
                            return 'clicked tbody td (direct): ' + (tds[i].className||'').substring(0,40);
                        }}
                    }}

                    // 디버그: 모든 button 텍스트 출력
                    var allBtns = dp.querySelectorAll('button');
                    var btnTexts = [];
                    allBtns.forEach(function(b) {{ btnTexts.push(b.textContent.trim().substring(0,5)); }});
                    return 'not found. all buttons: [' + btnTexts.join(',') + ']';
                """)
                time.sleep(0.7)
                actual = date_input.get_attribute('value')
                print(f"  [달력 JS 결과] {clicked}")
                print(f"  [달력 JS 후 날짜값] '{actual}'")
                date_ok = (target_month in actual and target_day_zp in actual)
                if 'clicked' in str(clicked):
                    found_day = True

                if not found_day:
                    print(f"  [경고] 달력에서 day={target_day_str} 못 찾음")

            except Exception as dp_e:
                print(f"  [경고] jQuery UI datepicker 처리 실패: {dp_e}")
                # fallback: input_date 현재값 확인
                try:
                    actual = date_input.get_attribute('value')
                    print(f"  [디버그] input_date 현재값: '{actual}'")
                except:
                    pass

        except Exception as e:
            print(f"  [경고] 날짜 입력 실패: {e}")
            try:
                driver.save_screenshot(f'{DEBUG_DIR}\\debug_date_error.png')
            except:
                pass

        if not date_ok:
            actual_val = ''
            try:
                actual_val = driver.find_element(
                    By.CSS_SELECTOR, 'input[class*="input_date"]').get_attribute('value')
            except:
                pass
            print(f"  [경고] 날짜 변경 실패 — 현재값: '{actual_val}', 목표: '{sched_date_display}'")

        # 시간 설정
        hour_ok = False
        try:
            from selenium.webdriver.support.ui import Select as SeSelect
            hour_sel = driver.find_element(By.CSS_SELECTOR, 'select[class*="hour_option"]')
            opts = [o.get_attribute('value') for o in hour_sel.find_elements(By.TAG_NAME, 'option')]
            val_h = str(SCHEDULE_HOUR)
            if val_h not in opts:
                val_h = f'{SCHEDULE_HOUR:02d}'
            SeSelect(hour_sel).select_by_value(val_h)
            hour_ok = True
            print(f"  [완료] 시간: {SCHEDULE_HOUR}시")
        except Exception as e:
            print(f"  [경고] 시간 select 실패: {e}")

        # 분 설정
        min_ok = False
        try:
            min_sel = driver.find_element(By.CSS_SELECTOR, 'select[class*="minute_option"]')
            opts = [o.get_attribute('value') for o in min_sel.find_elements(By.TAG_NAME, 'option')]
            val_m = str(SCHEDULE_MINUTE)
            if val_m not in opts:
                val_m = f'{SCHEDULE_MINUTE:02d}'
            SeSelect(min_sel).select_by_value(val_m)
            min_ok = True
            print(f"  [완료] 분: {SCHEDULE_MINUTE:02d}분")
        except Exception as e:
            print(f"  [경고] 분 select 실패: {e}")

        time.sleep(0.5)
        driver.save_screenshot(f'{DEBUG_DIR}\\debug_datetime_set.png')

        # 예약 확정: 패널 안의 발행 버튼 클릭
        confirmed = False
        url_before = driver.current_url
        print(f"  [확정 전 URL] {url_before}")

        # 현재 표시되는 모든 버튼 목록 출력 (디버그)
        all_btn_info = driver.execute_script("""
            var btns = document.querySelectorAll('button');
            var info = [];
            btns.forEach(function(b) {
                if (b.offsetParent !== null) {  // 보이는 버튼만
                    info.push(b.textContent.trim().substring(0,15) + '|cls:' + (b.className||'').substring(0,30));
                }
            });
            return info;
        """)
        print(f"  [보이는 버튼 목록] {all_btn_info}")

        try:
            # confirm_btn 클래스 — 패널 안의 최종 '발행' 확정 버튼
            confirm_btns = driver.find_elements(By.XPATH,
                "//button[contains(@class,'confirm_btn')]")
            for btn in confirm_btns:
                if btn.is_displayed():
                    print(f"  [확정] confirm_btn 발견: '{btn.text.strip()}' cls={btn.get_attribute('class')}")
                    driver.execute_script("arguments[0].click();", btn)
                    confirmed = True
                    break
        except Exception as e:
            print(f"  [경고] confirm_btn XPATH 실패: {e}")

        if not confirmed:
            try:
                # publish_btn 제외하고 텍스트가 '발행'인 버튼 찾기
                btns = driver.find_elements(By.TAG_NAME, 'button')
                candidates = []
                for btn in btns:
                    cls = btn.get_attribute('class') or ''
                    txt = btn.text.strip()
                    if btn.is_displayed() and txt == '발행' and 'publish_btn' not in cls:
                        candidates.append(btn)
                        print(f"  [확정 후보] '{txt}' cls={cls[:40]}")
                if candidates:
                    btn = candidates[-1]
                    driver.execute_script("arguments[0].click();", btn)
                    confirmed = True
                    print(f"  [완료] 예약 확정 (publish_btn 제외 '발행' 버튼)")
            except Exception as e:
                print(f"  [경고] 발행 버튼(publish_btn 제외) 실패: {e}")

        if not confirmed:
            try:
                # 최후 fallback: 예약발행, 확인 버튼
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

        url_changed = (url_after != url_before)
        post_success = url_changed or 'view' in url_after.lower() or 'PostView' in url_after

        if date_ok and confirmed and post_success:
            print(f"\n✅ [완료] 예약 발행 성공! {sched_date} 오전 {SCHEDULE_HOUR}시")
            print(f"   새 URL: {url_after}")
        elif date_ok and confirmed and not post_success:
            print(f"\n⚠️ [주의] 버튼은 클릭됐지만 URL 변경 없음 — 제목/내용 검증 실패 가능")
            print(f"   → debug_final_result.png 확인 필요")
        else:
            print(f"\n⚠️ [경고] 예약 미완료 — date_ok={date_ok}, confirmed={confirmed}, url변경={url_changed}")
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
