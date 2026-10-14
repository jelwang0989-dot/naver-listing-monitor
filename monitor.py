# -*- coding: utf-8 -*-
"""
네이버 부동산 신규 매물 감지 프로그램 (monitor.py)
- undetected-chromedriver를 통한 봇 탐지 우회
- 다산 지금지구 3개 단지 순회 크롤링 (동일 매물 묶음 펼치기 포함)
- DB(JSON) 대조를 통한 신규 매물 자동 감지 및 갱신
"""

import os
import sys
import json
import time
import argparse
import traceback
import requests
if sys.platform == "win32":
    import winreg
else:
    winreg = None
from datetime import datetime

# Google Apps Script 웹앱 URL
GOOGLE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxAjBEnfiDRS7L9Gvp1aypbTd6wQOsM-VhgjfVQ2rGxfUXVc_6196GIs4zGWdJrzk1u/exec"

# Windows 콘솔 인코딩 UTF-8 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# uc.Chrome 종료 시 Windows 환경의 OSError [WinError 6] 핸들 오류 방지 패치
def _safe_chrome_del(self):
    try:
        self.quit()
    except (OSError, Exception):
        pass

uc.Chrome.__del__ = _safe_chrome_del


# 대상 단지 정보 (총 6개 단지)
TARGET_COMPLEXES = [
    {
        "id": "115538",
        "name": "다산팬테리움리버테라스I",
        "naver_name": "다산펜테리움리버테라스Ⅰ",
        "url": "https://new.land.naver.com/complexes/115538"
    },
    {
        "id": "117630",
        "name": "다산신안인스빌퍼스트포레",
        "naver_name": "다산신안인스빌퍼스트포레",
        "url": "https://new.land.naver.com/complexes/117630"
    },
    {
        "id": "115545",
        "name": "다산팬테리움리버테라스II",
        "naver_name": "다산펜테리움리버테라스Ⅱ",
        "url": "https://new.land.naver.com/complexes/115545"
    },
    {
        "id": "115493",
        "name": "다산지금센트럴에일린의뜰",
        "naver_name": "다산지금센트럴에일린의뜰",
        "url": "https://new.land.naver.com/complexes/115493"
    },
    {
        "id": "115539",
        "name": "다산한강반도유보라",
        "naver_name": "다산반도유보라메이플타운2.0",
        "url": "https://new.land.naver.com/complexes/115539"
    },
    {
        "id": "115601",
        "name": "다산신안인스빌퍼스트리버",
        "naver_name": "다산신안인스빌퍼스트리버",
        "url": "https://new.land.naver.com/complexes/115601"
    }
]

CURRENT_FILE = "current_listings.json"
KNOWN_FILE = "known_listings.json"
NEW_FILE = "new_listings.json"


def get_installed_chrome_version():
    """운영체제(Windows/Linux)에 설치된 Chrome 브라우저의 메이저 버전을 조회합니다."""
    if sys.platform == "win32" and winreg:
        reg_paths = [
            (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon", "version"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\Chrome\BLBeacon", "version"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Update\Clients\{8A69D345-D564-463c-AFF1-A69D9E530F96}", "pv")
        ]
        for root, path, key_name in reg_paths:
            try:
                with winreg.OpenKey(root, path) as key:
                    val, _ = winreg.QueryValueEx(key, key_name)
                    if val:
                        major = int(str(val).split(".")[0])
                        return major
            except Exception:
                continue
    elif sys.platform.startswith("linux"):
        import subprocess
        for cmd in ["google-chrome", "google-chrome-stable", "chromium-browser", "chromium"]:
            try:
                out = subprocess.check_output([cmd, "--version"], stderr=subprocess.DEVNULL).decode("utf-8")
                parts = out.strip().split()
                for p in parts:
                    clean_p = p.replace(".", "")
                    if clean_p.isdigit():
                        return int(p.split(".")[0])
            except Exception:
                continue
    return None


def create_driver(headless=False):
    """undetected-chromedriver 인스턴스를 초기화하여 반환합니다."""
    chrome_version = get_installed_chrome_version()
    if chrome_version:
        print(f"[*] 설치된 Chrome 메이저 버전 감지: {chrome_version}")
    else:
        print("[*] Chrome 메이저 버전 자동 감지 모드로 구동합니다.")

    options = uc.ChromeOptions()
    # 1. 렌더러 충돌 및 타임아웃 방지 필수 플래그 설정
    options.page_load_strategy = "eager"
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-browser-side-navigation")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--lang=ko-KR,ko")

    # 2. 최신 일반 브라우저 User-Agent 헤더 적용
    modern_ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    options.add_argument(f"--user-agent={modern_ua}")

    # Cloud / Linux 환경 대응: headless 요청 시 옵션 추가
    is_headless = headless or os.environ.get("HEADLESS", "").lower() in ["1", "true", "yes"]
    if is_headless:
        print("[*] Headless(창 숨김) 모드로 브라우저를 실행합니다.")
        options.add_argument("--headless=new")

    driver_kwargs = {"options": options}
    if chrome_version:
        driver_kwargs["version_main"] = chrome_version

    driver = uc.Chrome(**driver_kwargs)
    try:
        driver.set_window_size(1920, 1080)
    except Exception:
        pass
    # 3. 페이지 로드 및 스크립트 타임아웃 설정 (30초)
    driver.set_page_load_timeout(30)
    driver.set_script_timeout(30)
    return driver


def safe_quit_driver(driver):
    """
    크롬 브라우저를 안전하게 종료합니다.
    이미 프로세스가 종료되었거나 핸들이 소멸하여 발생하는 OSError [WinError 6] 및
    각종 프로세스 종료 예외를 조용히 무시(pass)합니다.
    """
    if not driver:
        return
    try:
        driver.quit()
    except (OSError, Exception):
        pass
    finally:
        try:
            driver.__class__.__del__ = lambda self: None
        except Exception:
            pass


def close_extra_windows(driver, main_handle=None):
    """
    아실, 이실장 등 매물 클릭 시 새로 열린 외부 사이트 탭(driver.window_handles >= 2)을 감지하여
    즉시 닫고 네이버 부동산 메인 탭으로 포커스를 복귀합니다.
    """
    try:
        handles = driver.window_handles
        if len(handles) <= 1:
            return

        # 메인 핸들 결정
        target_main = main_handle if (main_handle and main_handle in handles) else None
        if not target_main:
            for h in handles:
                try:
                    driver.switch_to.window(h)
                    if "land.naver.com" in driver.current_url:
                        target_main = h
                        break
                except Exception:
                    continue
            if not target_main:
                target_main = handles[0]

        # 메인 창 이외의 모든 팝업 및 외부 사이트 새 창 닫기
        closed_count = 0
        for h in handles:
            if h != target_main:
                try:
                    driver.switch_to.window(h)
                    driver.close()
                    closed_count += 1
                except Exception:
                    pass

        # 네이버 부동산 메인 창으로 복귀
        driver.switch_to.window(target_main)
        if closed_count > 0:
            print(f"    [-] 외부 사이트 팝업/새 창 {closed_count}개를 닫고 메인 탭으로 복귀했습니다.")
    except Exception:
        pass


def dismiss_overlays(driver):
    """화면에 나타나는 팝업 레이어, 툴팁, 딤드(dimmed) 안내창을 닫습니다."""
    try:
        dismiss_js = """
        const popups = document.querySelectorAll('div.dimmed button, .layer_tooltip button, .layer_guide button, [class*="btn_close"]');
        popups.forEach(btn => {
            try { btn.click(); } catch(e) {}
        });
        """
        driver.execute_script(dismiss_js)
    except Exception:
        pass


def scroll_listing_panel(driver, max_scrolls=100, pause_sec=0.35, max_same=5):
    """
    매물 목록 좌측 패널을 끝까지 스크롤하여 모든 매물이 DOM에 100% 로드되도록 합니다.
    더 이상 스크롤 높이가 변하지 않을 때까지 충분히 반복하여 누락을 방지합니다.
    """
    scroll_js = """
    const container = document.querySelector('.item_list--article') || document.querySelector('#articleListArea');
    if (!container) return { done: true, height: 0 };
    const prev = container.scrollHeight;
    container.scrollTop = container.scrollHeight;
    return { done: false, prevHeight: prev, newHeight: container.scrollHeight };
    """
    last_height = 0
    same_count = 0
    for i in range(max_scrolls):
        try:
            res = driver.execute_script(scroll_js)
            time.sleep(pause_sec)
            if not res or res.get("done"):
                break
            curr_height = res.get("newHeight", 0)
            if curr_height == last_height:
                same_count += 1
                if same_count >= max_same:
                    break
            else:
                same_count = 0
                last_height = curr_height
        except Exception:
            break


# React Fiber 및 DOM 보조 파싱 JavaScript 함수
EXTRACT_REACT_ARTICLE_JS = """
function extractArticleFromNode(domNode) {
    if (!domNode) return null;
    
    // 1. React Fiber props 우선 탐색
    for (let key in domNode) {
        if (key.startsWith('__reactFiber') || key.startsWith('__reactInternalInstance')) {
            let fiber = domNode[key];
            let curr = fiber;
            let depth = 0;
            while (curr && depth < 25) {
                if (curr.memoizedProps) {
                    for (let p in curr.memoizedProps) {
                        let val = curr.memoizedProps[p];
                        if (val && typeof val === 'object' && val.articleNo) {
                            return {
                                articleNo: String(val.articleNo || '').trim(),
                                articleName: val.articleName || '',
                                buildingName: val.buildingName || '',
                                floorInfo: val.floorInfo || '',
                                tradeTypeName: val.tradeTypeName || '',
                                dealOrWarrantPrc: val.dealOrWarrantPrc || '',
                                rentPrc: val.rentPrc ? String(val.rentPrc).trim() : '',
                                realtorName: val.realtorName || (val.realtor && val.realtor.realtorName) || '',
                                sameAddrCnt: Number(val.sameAddrCnt || 1),
                                area1: val.area1 ? String(val.area1) : '',
                                area2: val.area2 ? String(val.area2) : '',
                                areaName: val.areaName || ''
                            };
                        }
                    }
                }
                curr = curr.return;
                depth++;
            }
        }
    }
    
    // 2. DOM 텍스트 보조 파싱
    try {
        const titleEl = domNode.querySelector('.item_title .text');
        const priceTypeEl = domNode.querySelector('.price_line .type');
        const priceValEl = domNode.querySelector('.price_line .price');
        const specEl = domNode.querySelector('.info_area .spec');
        const agentEl = domNode.querySelector('.cp_area .agent_name, [data-nclk="TAA.realtor"]');
        
        let titleText = titleEl ? titleEl.innerText.trim() : '';
        let dongMatch = titleText.match(/(\\d+동)/);
        let dong = dongMatch ? dongMatch[1] : '';
        
        let specText = specEl ? specEl.innerText.trim() : '';
        let floorMatch = specText.match(/(\\d+\\/\\d+층|\\d+층|[저중고]\\/\\d+층|[저중고]층)/);
        let floor = floorMatch ? floorMatch[1] : '';
        let areaMatch = specText.match(/(\\d+[A-Za-z가-힣]*\\/\\d+m²|\\d+m²|\\d+[A-Za-z])/);
        let areaStr = areaMatch ? areaMatch[1] : '';
        
        let tradeType = priceTypeEl ? priceTypeEl.innerText.trim() : '';
        let price = priceValEl ? priceValEl.innerText.trim() : '';
        let realtor = agentEl ? agentEl.innerText.trim() : '';
        
        // 매물번호 탐색 (이미지 썸네일 url 내 번호 등)
        let artNo = '';
        const thumb = domNode.querySelector('.thumbnail, [style*="landthumb"]');
        if (thumb) {
            let style = thumb.getAttribute('style') || '';
            let m = style.match(/\\/(\\d{10})_/);
            if (m) artNo = m[1];
        }

        if (price || tradeType || dong) {
            return {
                articleNo: artNo,
                articleName: '',
                buildingName: dong,
                floorInfo: floor,
                tradeTypeName: tradeType,
                dealOrWarrantPrc: price,
                rentPrc: '',
                realtorName: realtor,
                sameAddrCnt: 1,
                area1: '',
                area2: '',
                areaName: areaStr
            };
        }
    } catch(e) {}
    
    return null;
}
"""


def get_url_with_retry(driver, url, max_retries=3, delay=4):
    """
    driver.get(url) 실행 시 net::ERR_CONNECTION_RESET이나 TimeoutException
    (특히 'Timed out receiving message from renderer') 발생 시,
    window.stop()을 걸어 비동기 매물 호출을 차단하지 않고,
    매물 목록 영역이 화면에 나타날 때까지 최대 15초간 명시적 대기(WebDriverWait)를 수행합니다.
    """
    for attempt in range(1, max_retries + 1):
        try:
            print(f"  [-] 네이버 부동산 접속 중... ({attempt}/{max_retries}): {url}")
            driver.get(url)
        except TimeoutException as e:
            err_line = str(e).splitlines()[0] if str(e) else repr(e)
            print(f"  [!] 페이지 로딩 타임아웃 감지 ({attempt}/{max_retries}): {err_line}")
        except WebDriverException as e:
            err_line = str(e).splitlines()[0] if str(e) else repr(e)
            print(f"  [!] 웹드라이버 접속 오류 감지 ({attempt}/{max_retries}): {err_line}")

        # 매물 목록 영역(div.item_list, infinite_scroll, item) 로딩 대기 (최대 15초)
        try:
            print("  [-] 매물 목록 요소 로딩 대기 중 (최대 15초)...")
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    "div.item_list, div[class*='item_list'], div.infinite_scroll, div.item"
                ))
            )
            print("  [v] 매물 목록 영역 감지 완료!")
            return True
        except TimeoutException:
            print("  [*] 15초 내에 매물 목록 요소를 감지하지 못했습니다. 페이지 내부 요소 점검 중...")
            # 1개라도 매물 관련 요소가 이미 렌더링되었는지 자바스크립트로 확인
            try:
                has_items = driver.execute_script("""
                    return document.querySelectorAll('div.item_list, div[class*="item_list"], div.infinite_scroll, .item').length > 0;
                """)
                if has_items:
                    print("  [v] 매물 관련 요소가 1개 이상 감지되어 다음 단계로 진행합니다.")
                    return True
            except Exception:
                pass

        if attempt < max_retries:
            print(f"  [*] {delay}초 대기 후 페이지 재접속을 시도합니다...")
            time.sleep(delay)
        else:
            print(f"  [!] 최대 재시도 횟수({max_retries}회)에 도달했습니다. 현재 로드된 상태에서 수집을 시도합니다.")
            return True


def scrape_complex(driver, complex_info):
    """
    단일 단지의 매물 페이지를 방문하여 단일 매물 및 펼쳐진 동일 매물 묶음의 개별 매물들을 전부 수집합니다.
    """
    complex_name = complex_info["name"]
    url = complex_info["url"]
    print(f"\n[{complex_name}] 매물 수집 시작: {url}")

    get_url_with_retry(driver, url, max_retries=3, delay=4)
    time.sleep(2.5)
    main_window_handle = driver.current_window_handle
    close_extra_windows(driver, main_window_handle)
    dismiss_overlays(driver)

    # 0. 매물 탭 활성화 확인 (필요한 경우 클릭)
    ensure_article_tab_js = """
    const items = document.querySelectorAll('.item_list .item, .item_list--article');
    if (items.length === 0) {
        const tabs = document.querySelectorAll('button, a, [role="tab"]');
        for (let el of tabs) {
            let t = (el.innerText || '').trim();
            if (t === '매물' || t.startsWith('매물 ') || t.startsWith('매물(') || t.startsWith('매물\n') || t === '전체매물') {
                el.click();
                return true;
            }
        }
    }
    return false;
    """
    try:
        if driver.execute_script(ensure_article_tab_js):
            print("  [-] '매물' 탭 클릭 활성화. 목록 로딩 대기...")
            time.sleep(2)
    except Exception:
        pass

    # 1. '동일매물 묶기' 체크박스 활성화 확인 (동일 매물 묶음 구조 파악)
    ensure_grouped_js = """
    const chk = document.querySelector('input#address_group2, input[name="type"]');
    if (chk && !chk.checked) {
        const lbl = document.querySelector('label[for="address_group2"]');
        if (lbl) {
            lbl.click();
            return true;
        }
    }
    return false;
    """
    try:
        close_extra_windows(driver, main_window_handle)
        toggled = driver.execute_script(ensure_grouped_js)
        if toggled:
            time.sleep(2)
        close_extra_windows(driver, main_window_handle)
    except Exception:
        pass

    dismiss_overlays(driver)

    # 2. 전체 매물 로드를 위한 스크롤 수행 (누락 방지를 위해 끝까지 로드)
    print(f"  [-] 전체 매물 목록 로딩(스크롤) 진행 중...")
    scroll_listing_panel(driver, max_scrolls=80)
    time.sleep(1)
    close_extra_windows(driver, main_window_handle)

    # 3. DOM에서 매물 그룹 목록 분석
    find_items_js = EXTRACT_REACT_ARTICLE_JS + """
    const items = document.querySelectorAll('.item_list .item:not(.item--child)');
    let results = [];
    items.forEach((item, idx) => {
        let link = item.querySelector('a.item_link') || item;
        let data = extractArticleFromNode(link) || extractArticleFromNode(item);
        let text = item.innerText || '';
        let isGroup = false;
        
        if (data && data.sameAddrCnt > 1) {
            isGroup = true;
        } else if (text.includes('중개사') || text.includes('동일매물') || item.querySelector('.btn_more, [class*="group"]')) {
            isGroup = true;
        }

        results.push({
            index: idx,
            isGroup: isGroup,
            data: data
        });
    });
    return results;
    """
    items_meta = driver.execute_script(find_items_js)
    if not items_meta:
        print(f"  [!] 매물 항목을 찾지 못했습니다. 재시도 중...")
        time.sleep(2)
        items_meta = driver.execute_script(find_items_js) or []

    print(f"  [-] 로드된 카드 수: {len(items_meta)}개 (그룹 묶음 및 단일 매물 포함)")

    collected_dict = {}  # key: articleNo (또는 복합키) -> listing dict

    # 4. 단일 매물 선별 및 동일 매물 묶음(그룹) 클릭 펼치기
    group_indices = []
    for item in items_meta:
        idx = item["index"]
        is_group = item["isGroup"]
        data = item.get("data")

        if is_group:
            group_indices.append(idx)
        else:
            # 단일 매물 즉시 수집
            if data and data.get("articleNo"):
                listing = format_listing_dict(complex_name, data)
                collected_dict[listing["매물번호"]] = listing

    print(f"  [-] 동일 매물 묶음(그룹) {len(group_indices)}개 감지. 개별 매물 펼치기 진행...")

    # 5. 각 그룹 묶음을 클릭하여 펼쳐진 개별 중개업소 매물(.item--child) 수집
    extract_children_js = EXTRACT_REACT_ARTICLE_JS + """
    const children = document.querySelectorAll('.item_list .item--child');
    let res = [];
    children.forEach(c => {
        let link = c.querySelector('a.item_link') || c;
        let d = extractArticleFromNode(link) || extractArticleFromNode(c);
        if (d) res.push(d);
    });
    return res;
    """

    for i, g_idx in enumerate(group_indices):
        click_group_js = f"""
        const items = document.querySelectorAll('.item_list .item:not(.item--child)');
        if (items[{g_idx}]) {{
            const link = items[{g_idx}].querySelector('a.item_link') || items[{g_idx}];
            if (link) {{
                // 외부 새 창 팝업 방지를 위해 target 속성 제거
                link.removeAttribute('target');
                link.scrollIntoView({{block: 'center', inline: 'nearest'}});
                link.click();
                return true;
            }}
        }}
        return false;
        """
        try:
            # 클릭 전 새 창 여부 점검 및 닫기
            close_extra_windows(driver, main_window_handle)

            driver.execute_script(click_group_js)
            time.sleep(0.4)  # 펼쳐지는 애니메이션 및 자식 렌더링 대기

            # 클릭 후 아실/이실장 등 외부 사이트 새 창 생성 시 즉시 닫기
            close_extra_windows(driver, main_window_handle)

            children_data = driver.execute_script(extract_children_js)
            if children_data:
                for c_data in children_data:
                    listing = format_listing_dict(complex_name, c_data)
                    collected_dict[listing["매물번호"]] = listing
        except Exception:
            continue

        if (i + 1) % 15 == 0 or (i + 1) == len(group_indices):
            print(f"    -> 묶음 {i + 1}/{len(group_indices)}개 펼치기 완료 (누적 {len(collected_dict)}건)")

    # 6. 추가 보완: 동일매물 묶기를 해제하여 단독 등록된 잔여 매물 전수 교차 수집
    try:
        close_extra_windows(driver, main_window_handle)
        untoggle_js = """
        const chk = document.querySelector('input#address_group2, input[name="type"]');
        if (chk && chk.checked) {
            const lbl = document.querySelector('label[for="address_group2"]');
            if (lbl) {
                lbl.click();
                return true;
            }
        }
        return false;
        """
        untoggled = driver.execute_script(untoggle_js)
        if untoggled:
            time.sleep(1.5)
            close_extra_windows(driver, main_window_handle)
            scroll_listing_panel(driver, max_scrolls=60)
            extract_all_flat_js = EXTRACT_REACT_ARTICLE_JS + """
            const items = document.querySelectorAll('.item_list .item');
            let res = [];
            items.forEach(it => {
                let link = it.querySelector('a.item_link') || it;
                let d = extractArticleFromNode(link) || extractArticleFromNode(it);
                if (d) res.push(d);
            });
            return res;
            """
            flat_items = driver.execute_script(extract_all_flat_js) or []
            for f_data in flat_items:
                listing = format_listing_dict(complex_name, f_data)
                collected_dict[listing["매물번호"]] = listing
            close_extra_windows(driver, main_window_handle)
    except Exception:
        pass

    results_list = list(collected_dict.values())
    print(f"  [v] [{complex_name}] 수집 완료: 총 {len(results_list)}개 고유 매물 확보")
    return results_list


def format_listing_dict(complex_name, raw_data):
    """원시 매물 데이터에서 규격화된 매물 딕셔너리를 생성합니다."""
    # 층수 정제
    floor_raw = str(raw_data.get("floorInfo", "")).strip()
    if floor_raw and not floor_raw.endswith("층") and not floor_raw.endswith("F"):
        floor_str = f"{floor_raw}층"
    else:
        floor_str = floor_raw

    # 동 정보 정제
    dong_raw = str(raw_data.get("buildingName", "")).strip()
    if dong_raw and not dong_raw.endswith("동"):
        dong_str = f"{dong_raw}동"
    else:
        dong_str = dong_raw

    # 면적 정제
    area_name = str(raw_data.get("areaName", "")).strip()
    area2 = str(raw_data.get("area2", "")).strip()
    area1 = str(raw_data.get("area1", "")).strip()
    if area_name and area2:
        area_str = f"{area_name}/{area2}㎡"
    elif area1 and area2:
        area_str = f"{area1}/{area2}㎡"
    elif area2:
        area_str = f"{area2}㎡"
    elif area_name:
        area_str = area_name
    else:
        area_str = str(raw_data.get("area", "")).strip()

    # 가격 정제 (월세의 경우 보증금/월세 결합)
    trade_type = str(raw_data.get("tradeTypeName", "")).strip()
    deal_prc = str(raw_data.get("dealOrWarrantPrc", "")).strip()
    rent_prc = str(raw_data.get("rentPrc", "")).strip()

    if trade_type == "월세" and rent_prc and rent_prc != "0":
        price_str = f"{deal_prc}/{rent_prc}"
    else:
        price_str = deal_prc

    art_no = str(raw_data.get("articleNo", "")).strip()
    # 고유 ID가 없는 예외 상황 발생 시 복합 식별자 생성
    if not art_no:
        art_no = f"{complex_name}_{dong_str}_{floor_str}_{trade_type}_{price_str}"

    return {
        "단지명": complex_name,
        "매물번호": art_no,
        "동": dong_str,
        "층": floor_str,
        "면적": area_str,
        "거래유형": trade_type,
        "가격": price_str,
        "등록공인중개사명": str(raw_data.get("realtorName", "")).strip()
    }


def send_to_google_sheet(new_listings):
    """
    감지된 신규 매물 리스트를 구글 시트 웹앱(Google Apps Script)으로 POST 전송합니다.
    """
    if not GOOGLE_WEBAPP_URL:
        print("[!] 구글 시트 웹앱 URL이 설정되지 않았습니다.")
        return False

    if not new_listings:
        print("[*] 구글 시트 전송: 새로 감지된 신규 매물이 없어 전송을 건너뜁니다.")
        return True

    payload_items = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for nl in new_listings:
        item = {
            "detected_at": nl.get("감지일시", now_str),
            "complex_name": nl.get("단지명", ""),
            "trade_type": nl.get("거래유형", ""),
            "price": nl.get("가격", ""),
            "building": nl.get("동", ""),
            "floor": nl.get("층", ""),
            "area": nl.get("면적", ""),
            "agency": nl.get("등록공인중개사명", ""),
            "article_no": nl.get("매물번호", "")
        }
        payload_items.append(item)

    payload = {"listings": payload_items}
    print(f"\n[+] 구글 시트 웹앱으로 신규 매물 {len(payload_items)}건 전송 시작...")

    try:
        response = requests.post(
            GOOGLE_WEBAPP_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        if response.status_code == 200:
            print(f"[v] 구글 시트 전송 성공! (HTTP {response.status_code}) 응답: {response.text.strip()}")
            return True
        else:
            print(f"[!] 구글 시트 전송 경고: 상태 코드 {response.status_code}, 응답: {response.text.strip()}")
            return False
    except requests.exceptions.Timeout:
        print("[!] 구글 시트 전송 실패: 요청 시간 초과 (30초 timeout)")
        return False
    except Exception as e:
        print(f"[!] 구글 시트 전송 중 오류 발생: {e}")
        return False


def compare_and_record_listings(current_list):
    """
    수집된 전체 목록을 저장하고 기존 데이터(known_listings.json)와 대조하여
    신규 매물을 감지, new_listings.json에 저장 및 콘솔에 출력합니다.
    대조가 끝나면 known_listings.json을 갱신합니다.
    """
    print("\n" + "=" * 70)
    print("                [ DB 대조 및 신규 매물 감지 ]")
    print("=" * 70)

    # 1. 수집한 전체 목록을 current_listings.json에 저장
    with open(CURRENT_FILE, "w", encoding="utf-8") as f:
        json.dump(current_list, f, ensure_ascii=False, indent=2)
    print(f"[+] 전체 수집 매물 저장: {CURRENT_FILE} (총 {len(current_list)}건)")

    # 2. 기존 known_listings.json 확인 및 비교
    new_listings = []
    is_initial_run = not os.path.exists(KNOWN_FILE)

    if not is_initial_run:
        try:
            with open(KNOWN_FILE, "r", encoding="utf-8") as f:
                known_list = json.load(f)
        except Exception as e:
            print(f"[!] 기존 파일({KNOWN_FILE}) 읽기 실패: {e}")
            known_list = []

        # 기존 매물 식별 세트 구축
        known_ids = {str(item.get("매물번호", "")).strip() for item in known_list if item.get("매물번호")}
        known_dong_floor_set = {
            (
                str(item.get("단지명", "")).strip(),
                str(item.get("동", "")).strip(),
                str(item.get("층", "")).strip()
            )
            for item in known_list
        }

        for item in current_list:
            art_id = str(item.get("매물번호", "")).strip()
            dong_floor_key = (
                str(item.get("단지명", "")).strip(),
                str(item.get("동", "")).strip(),
                str(item.get("층", "")).strip()
            )

            is_new = False
            reasons = []

            # 조건 1: 기존에 없던 새로운 매물번호
            if art_id and art_id not in known_ids:
                is_new = True
                reasons.append("신규 매물번호")

            # 조건 2: 신규 동/층 조합
            if dong_floor_key not in known_dong_floor_set:
                is_new = True
                reasons.append("신규 동/층")

            if is_new:
                record = dict(item)
                record["감지일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                record["감지이유"] = ", ".join(reasons)
                new_listings.append(record)

        # 신규 매물 출력
        print("\n" + "-" * 70)
        if new_listings:
            print(f"★ [신규 등록 매물 감지 알림] 총 {len(new_listings)}건의 신규 매물이 발견되었습니다!\n")
            header = f"{'번호':<4} {'단지명':<20} {'동/층':<14} {'유형':<6} {'가격':<12} {'공인중개사명':<20} {'매물번호'}"
            print(header)
            print("-" * 95)
            for i, nl in enumerate(new_listings, 1):
                dong_floor = f"{nl['동']} {nl['층']}"
                print(f"[{i:02d}] {nl['단지명']:<20} {dong_floor:<14} {nl['거래유형']:<6} {nl['가격']:<12} {nl['등록공인중개사명']:<20} {nl['매물번호']} ({nl['감지이유']})")
            print("-" * 95)
        else:
            print("✓ 감지 결과: 새로 등록된 신규 매물이 없습니다. (기존 데이터와 일치)")
        print("-" * 70 + "\n")

    else:
        print(f"[*] 최초 실행 안내: 기존 매물 파일({KNOWN_FILE})이 없어, 현재 수집된 {len(current_list)}건을 기준 데이터로 최초 저장합니다.")
        new_listings = []

    # 3. new_listings.json 저장
    with open(NEW_FILE, "w", encoding="utf-8") as f:
        json.dump(new_listings, f, ensure_ascii=False, indent=2)
    print(f"[+] 신규 매물 목록 저장: {NEW_FILE} ({len(new_listings)}건)")

    # 4. 구글 시트 웹앱 전송
    if new_listings:
        send_to_google_sheet(new_listings)
    else:
        print("[*] 구글 시트 전송: 새로 감지된 신규 매물이 없어 전송을 건너뜁니다.")

    # 5. known_listings.json 갱신
    with open(KNOWN_FILE, "w", encoding="utf-8") as f:
        json.dump(current_list, f, ensure_ascii=False, indent=2)
    print(f"[+] 기준 매물 데이터 갱신 완료: {KNOWN_FILE} ({len(current_list)}건)")
    print("=" * 70 + "\n")


def run_crawler(headless=False):
    """크롤러 메인 실행 루틴"""
    start_time = time.time()
    print("=" * 70)
    print(f" 네이버 부동산 신규 매물 감지 크롤러 가동 [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]")
    print("=" * 70)

    driver = None
    all_collected = []

    try:
        driver = create_driver(headless=headless)

        for idx, target in enumerate(TARGET_COMPLEXES, 1):
            try:
                print(f"\n[{idx}/{len(TARGET_COMPLEXES)}] {target['name']} 처리 시작...")
                complex_listings = scrape_complex(driver, target)
                all_collected.extend(complex_listings)
            except Exception as e:
                print(f"[!] {target['name']} 수집 중 예외 발생: {e}")
                traceback.print_exc()

        # 전체 매물 중복 제거 (매물번호 기준)
        dedup_dict = {}
        for item in all_collected:
            dedup_dict[item["매물번호"]] = item
        final_list = list(dedup_dict.values())

        print(f"\n[+] {len(TARGET_COMPLEXES)}개 단지 크롤링 완료: 총 {len(final_list)}건 수집됨 (소요시간: {time.time() - start_time:.1f}초)")

        # 신규 매물 대조 및 저장
        compare_and_record_listings(final_list)

    finally:
        if driver:
            safe_quit_driver(driver)


def main():
    parser = argparse.ArgumentParser(description="네이버 부동산 타 부동산 신규 매물 감지 프로그램")
    parser.add_argument("--loop", action="store_true", help="지정된 주기마다 반복 감시 모드로 실행")
    parser.add_argument("--interval", type=int, default=300, help="반복 감시 주기 (초 단위, 기본 300초=5분)")
    parser.add_argument("--headless", action="store_true", help="브라우저 창을 띄우지 않고 백그라운드(Headless)로 실행")
    parser.add_argument("--test-sheet", action="store_true", help="구글 시트 웹앱 연동 테스트 (샘플 1건 전송)")
    args = parser.parse_args()

    if args.test_sheet:
        print("[*] 구글 시트 웹앱 연동 테스트를 진행합니다...")
        mock_sample = [{
            "단지명": "다산팬테리움리버테라스I",
            "거래유형": "매매",
            "가격": "8억 5,000",
            "동": "101동",
            "층": "15/25층",
            "면적": "111B/84㎡",
            "등록공인중개사명": "테스트공인중개사사무소",
            "매물번호": "TEST_" + datetime.now().strftime("%Y%m%d%H%M%S"),
            "감지일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "감지이유": "테스트 전송"
        }]
        success = send_to_google_sheet(mock_sample)
        if success:
            print("[v] 구글 시트 웹앱 연동 테스트 완료!")
        else:
            print("[x] 구글 시트 웹앱 연동 테스트 실패.")
        return

    if args.loop:
        print(f"[*] 주기적 반복 감시 모드 시작 (감시 주기: {args.interval}초)")
        while True:
            try:
                run_crawler(headless=args.headless)
            except KeyboardInterrupt:
                print("\n[*] 사용자에 의해 프로그램이 중단되었습니다.")
                break
            except Exception as e:
                print(f"[!] 실행 중 오류 발생: {e}")
            print(f"[*] 다음 감시까지 {args.interval}초 대기합니다... (Ctrl+C 로 종료)")
            time.sleep(args.interval)
    else:
        run_crawler(headless=args.headless)


if __name__ == "__main__":
    main()
