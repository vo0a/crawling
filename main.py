from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from datetime import datetime
import openpyxl
import xlrd
from bs4 import BeautifulSoup, NavigableString, Tag
import os
import time
import re
from typing import List, Dict
import logging
from dotenv import load_dotenv

load_dotenv()

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Rental Schedule API")

# 환경변수 및 경로 설정
BASE_DIR = os.getcwd()
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
SCREENSHOT_DIR = os.path.join(BASE_DIR, "screenshots")
LOGIN_URL = os.environ.get("LOGIN_URL")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# -----------------------------------------------------------
# 1. Selenium 및 기본 설정 함수
# -----------------------------------------------------------


def get_chrome_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument(
        '--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option(
        "excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(10)
    return driver


def save_screenshot(driver, name: str):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join(SCREENSHOT_DIR, f"{name}_{timestamp}.png")
        driver.save_screenshot(filepath)
        return filepath
    except Exception:
        return None


def login(driver):
    try:
        driver.get(LOGIN_URL)
        time.sleep(1)
        driver.find_element(By.ID, "Login_id").send_keys(
            os.environ.get("USERNAME"))
        driver.find_element(By.ID, "Login_pw").send_keys(
            os.environ.get("PASSWORD"))
        driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
        time.sleep(2)
        if "login" in driver.current_url.lower():
            raise Exception("로그인 실패")
        logger.info("로그인 성공")
        return True
    except Exception as e:
        save_screenshot(driver, "error_login")
        raise e


def navigate_to_daily_schedule(driver):
    try:
        logger.info("메뉴 이동 중...")
        driver.switch_to.default_content()
        time.sleep(0.5)
        try:
            driver.switch_to.frame("topFrame")
        except:
            driver.switch_to.frame(0)

        try:
            menu = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(text(), '대여일정')]")))
            driver.execute_script("arguments[0].click();", menu)
        except:
            menu = driver.find_element(By.XPATH, "/html/body/div/div/div[3]")
            driver.execute_script("arguments[0].click();", menu)

        time.sleep(2)
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(1)
        except:
            for f in driver.find_elements(By.TAG_NAME, "frame"):
                if f.get_attribute("name") != "topFrame":
                    driver.switch_to.frame(f)
                    break

        daily_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, "//a[contains(text(), '일간')] | //a[contains(@href, 'rent_day')]")))
        driver.execute_script("arguments[0].click();", daily_btn)
        time.sleep(2)
    except Exception as e:
        save_screenshot(driver, "error_nav")
        raise e


def navigate_to_date(driver, target_date: datetime):
    try:
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(1)
        except:
            driver.switch_to.frame("mainFrame")

        t_str = target_date.strftime("%Y-%m-%d")
        t_day = str(target_date.day)

        for _ in range(24):
            try:
                selects = driver.find_elements(
                    By.CSS_SELECTOR, "#sidebar select")
                if len(selects) >= 2:
                    c_yr = int(
                        Select(selects[0]).first_selected_option.text.replace('년', '').strip())
                    c_mo = int(
                        Select(selects[1]).first_selected_option.text.replace('월', '').strip())
                else:
                    h_txt = driver.find_element(
                        By.CSS_SELECTOR, "#sidebar .lnb-cal tr:first-child").text
                    c_yr = int(re.search(r'(\d{4})', h_txt).group(1))
                    c_mo = int(re.search(r'(\d{1,2})', h_txt).group(1))

                if c_yr == target_date.year and c_mo == target_date.month:
                    break

                if (c_yr * 12 + c_mo) < (target_date.year * 12 + target_date.month):
                    btn = driver.find_element(
                        By.CSS_SELECTOR, "#sidebar .next, #sidebar a:contains('>')")
                else:
                    btn = driver.find_element(
                        By.CSS_SELECTOR, "#sidebar .prev, #sidebar a:contains('<')")
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.5)
            except:
                break

        try:
            xpath = f"//div[contains(@class,'lnb-cal')]//td[not(contains(@class,'other'))]//a[normalize-space(text())='{t_day}']"
            date_link = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].click();", date_link)
        except:
            driver.execute_script(
                f"goPlanToday('{t_str}', '{target_date.year}', '{target_date.month}');")

        time.sleep(1)
        return True
    except Exception as e:
        save_screenshot(driver, f"err_date_{t_str}")
        raise e


def download_excel_for_date(driver, target_date: datetime):
    date_str = target_date.strftime("%Y-%m-%d")
    try:
        navigate_to_date(driver, target_date)
        time.sleep(2)
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(1)
        except:
            driver.switch_to.frame("mainFrame")

        existing_files = set(os.listdir(DOWNLOAD_DIR))

        try:
            excel_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(text(), '엑셀')] | //a[contains(@href, 'excel')] | //img[contains(@src, 'excel')]/parent::a")))
            driver.execute_script("arguments[0].click();", excel_btn)
        except Exception as e:
            logger.error(f"엑셀 버튼 못찾음: {e}")
            return None

        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
            return None
        except:
            pass

        for i in range(15):
            current_files = set(os.listdir(DOWNLOAD_DIR))
            new_files = current_files - existing_files
            valid_files = [f for f in new_files if f.endswith(
                '.xls') or f.endswith('.xlsx')]
            if valid_files:
                f_path = os.path.join(DOWNLOAD_DIR, valid_files[0])
                time.sleep(1)
                return f_path
            time.sleep(1)
        return None
    except Exception:
        return None


# -----------------------------------------------------------
# 2. [핵심] HTML(Fake Excel) 파싱 및 문자열 정제 로직
# -----------------------------------------------------------
import re
from bs4 import NavigableString, Tag


def clean_text_list(text_list):
    """
    [문자열 세탁기]
    리스트를 합친 뒤, 지저분한 콤마(,,,)와 공백을 깔끔하게 정리합니다.
    """
    # 1. 빈 값 제거
    valid_texts = [t.strip()
                   for t in text_list if t.strip() and t.strip() != ',']

    # 2. 합치기 (일단 콤마로 연결)
    full_text = ", ".join(valid_texts)

    # 3. 정규식으로 청소
    # ", ," 또는 ",," 처럼 콤마가 반복되는 것을 하나로 통일
    full_text = re.sub(r'\s*,\s*', ', ', full_text)  # 공백 정리
    full_text = re.sub(r'(,\s*){2,}', ', ', full_text)  # 중복 콤마 제거

    # 4. 앞뒤 콤마 제거
    return full_text.strip(', ').strip()


def extract_red_text_html(td_tag):
    """
    HTML 태그 내 텍스트 색상 판별 (CSS 상속 우선순위 적용)
    """
    normal_parts = []
    red_parts = []

    # <br> 태그는 콤마로 치환 (줄바꿈 = 상품 구분)
    for br in td_tag.find_all("br"):
        br.replace_with(",")

    # 모든 하위 텍스트 노드를 순회
    for element in td_tag.descendants:
        # 텍스트 노드인 경우만 처리
        if isinstance(element, NavigableString):
            text = element.strip()
            # 의미 없는 기호 무시
            if not text or text == ",":
                continue

            # [핵심 로직] 부모를 거슬러 올라가며 "가장 가까운 색상" 찾기
            is_red = False      # 기본값
            color_found = False  # 색상 정의를 찾았는지 여부

            parent = element.parent
            while parent:
                # 더 이상 검사할 태그가 없거나 테이블 셀을 벗어나면 중단
                if not isinstance(parent, Tag) or parent.name == '[document]':
                    break

                attrs = parent.attrs
                style = attrs.get('style', '').lower().replace(" ", "")
                color = attrs.get('color', '').lower()

                # 1. 빨간색 정의 확인
                is_explicit_red = ('color:red' in style) or \
                                  ('color:#ff' in style and len(style.split('color:#ff')[1].split(';')[0]) <= 4) or \
                                  (color == 'red') or (color.startswith('#ff'))

                # 2. 빨간색이 아닌 다른 색상 정의 확인 (blue, black 등)
                # color: 속성이 있는데 red가 아니면 다른 색임
                is_explicit_other = ('color:' in style and not is_explicit_red) or \
                                    (color and not is_explicit_red)

                if is_explicit_other:
                    # 다른 색(파란색 등)이 먼저 감지됨 -> 일반 상품
                    is_red = False
                    color_found = True
                    break

                if is_explicit_red:
                    # 빨간색이 먼저 감지됨 -> 추가 상품
                    is_red = True
                    color_found = True
                    break

                # td 태그까지 왔는데 별다른 색상 정의가 없었다면?
                # (td 자체의 색상을 확인하고 루프 종료)
                if parent.name == 'td':
                    break

                parent = parent.parent

            # 찾은 결과에 따라 분류
            if is_red:
                red_parts.append(text)
            else:
                normal_parts.append(text)

    return clean_text_list(normal_parts), clean_text_list(red_parts)


def parse_html_xls(path, date):
    logger.info(f"▶ HTML(Fake Excel) 파서 실행: {path}")
    res = []

    # 1. 인코딩 감지 (제공해주신 파일이 euc-kr이므로 이것부터 시도)
    encodings = ['euc-kr', 'cp949', 'utf-8']
    content = ""

    for enc in encodings:
        try:
            with open(path, 'r', encoding=enc, errors='ignore') as f:
                temp = f.read()
                # 헤더 키워드로 올바른 인코딩인지 검증
                if '고객' in temp or '지점' in temp or '대여' in temp:
                    content = temp
                    logger.info(f"✔ 인코딩 확정: {enc}")
                    break
        except:
            continue

    if not content:
        logger.error("❌ 파일 내용을 읽을 수 없습니다.")
        return []

    try:
        soup = BeautifulSoup(content, 'html.parser')
        table = soup.find('table')
        if not table:
            return []

        rows = table.find_all('tr')
        headers = []
        start_row = 0

        # 2. 헤더 찾기
        for idx, row in enumerate(rows[:10]):
            cells = [c.get_text(strip=True).replace('\xa0', '')
                     for c in row.find_all(['td', 'th'])]
            if any(h in cells for h in ['지점', '고객명']):
                headers = cells
                start_row = idx + 1
                logger.info(f"✅ 헤더 발견: {headers}")
                break

        if not headers and rows:
            headers = [c.get_text(strip=True).replace('\xa0', '')
                       for c in rows[0].find_all(['td', 'th'])]
            start_row = 1

        # 3. 데이터 파싱
        for row in rows[start_row:]:
            cells = row.find_all(['td', 'th'])
            row_data = {}

            for c_idx, c in enumerate(cells):
                if c_idx >= len(headers):
                    break
                h = headers[c_idx]
                if not h:
                    h = f"Col_{c_idx}"

                # [수정됨] 대여상품 컬럼 파싱
                if '대여' in h and '상품' in h:  # '대여상품', '대여 상품' 등 유연하게
                    normal_txt, red_txt = extract_red_text_html(c)
                    row_data['대여상품'] = normal_txt
                    row_data['추가상품'] = red_txt
                else:
                    # 일반 컬럼도 텍스트 클리닝 적용
                    raw_texts = [t.strip() for t in c.stripped_strings]
                    row_data[h] = clean_text_list(raw_texts)

            # 유효 데이터 필터링
            cust = row_data.get('고객명', '').strip()
            # 고객명에 '담당자' 같은 헤더성 데이터가 섞여 들어오면 제외
            if cust and cust != 'None' and '담당자' not in cust:
                row_data['대여일자'] = date
                res.append(row_data)

        logger.info(f"✔ 파싱 완료: {len(res)}건")
        return res

    except Exception as e:
        logger.error(f"HTML 파싱 오류: {e}")
        return []


def parse_excel(path, date):
    try:
        # Fake Excel Check
        with open(path, 'rb') as f:
            if b'<!DOCTYPE' in f.read(20) or b'<html' in f.read(20):
                return parse_html_xls(path, date)

        # 실제 XLS/XLSX는 기존 로직 유지 (생략)
        # 하지만 보통 Fake Excel이므로 여기로 넘어오는 경우는 드뭄
        return parse_html_xls(path, date)  # 일단 HTML 파서로 보냄 (안전장치)

    except Exception:
        return parse_html_xls(path, date)


def clean_dirs():
    # (폴더경로, 삭제할 확장자 튜플) 리스트
    targets = [
        (DOWNLOAD_DIR, ('.xls', '.xlsx')),
        (SCREENSHOT_DIR, ('.png', '.html'))
    ]

    for folder, extensions in targets:
        if not os.path.exists(folder):
            continue

        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) and filename.lower().endswith(extensions):
                    os.unlink(file_path)
            except Exception:
                pass

# -----------------------------------------------------------
# 3. API 엔드포인트
# -----------------------------------------------------------


@app.get("/rentals")
async def get_rentals(
    dates: List[str] = Query(...,
                             description="조회할 날짜 리스트 (예: ?dates=2025-12-10&dates=2025-12-15)")
):
    driver = None
    try:
        clean_dirs()

        # FastAPI 기본 스펙: &dates=... 사용
        target_dates = []
        for d in dates:
            # 콤마로 들어오는 경우 방어 코드 (혹시나 해서)
            for split_d in d.split(','):
                try:
                    target_dates.append(datetime.strptime(
                        split_d.strip(), "%Y-%m-%d"))
                except:
                    pass

        target_dates = sorted(list(set(target_dates)))
        logger.info(f"수집 대상: {[d.strftime('%Y-%m-%d') for d in target_dates]}")

        driver = get_chrome_driver()
        login(driver)
        navigate_to_daily_schedule(driver)

        all_data = []
        for curr in target_dates:
            d_str = curr.strftime("%Y-%m-%d")
            f = download_excel_for_date(driver, curr)
            if f:
                data = parse_excel(f, d_str)
                all_data.extend(data)
                # os.remove(f) # 디버깅용 (삭제 안함)
            else:
                logger.warning(f"{d_str}: 데이터 없음")

        return JSONResponse(content={
            "success": True,
            "total_count": len(all_data),
            "data": all_data
        })

    except Exception as e:
        logger.error(f"API 오류: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
