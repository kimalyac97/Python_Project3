# -*- coding: utf-8 -*-
"""
pp.kepco.co.kr 배치 자동화 – Streamlit Cloud 대응판 (Chromium 설치용)
- 원본 기능(로그인/세션 리셋/고객정보 진입/탭 열기/필드 수집/엑셀 저장) 동일
- Streamlit Cloud, Docker 등 컨테이너 환경 호환 (chromium + chromedriver 설치 기반)

필수 파일:
  requirements.txt:
    streamlit>=1.37
    selenium>=4.24
    pandas
    openpyxl
    python-dotenv

  packages.txt:
    chromium
    chromium-driver

실행:
  $ streamlit run streamlit_app.py
"""

import os, re, sys, time, zipfile
from io import BytesIO
from pathlib import Path
from typing import Tuple, Optional
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv

import streamlit as st

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    UnexpectedAlertPresentException,
)

from openpyxl import Workbook, load_workbook

URL_BASE = "https://pp.kepco.co.kr"
URL_INTRO = f"{URL_BASE}/intro.do"
CUSTOMER_INFO_PATH = "/mb/mb0101.do?menu_id=O010601"

HEADERS = ["자원명", "고객사명", "ID(고객번호)", "PW", "계기번호", "한전 계약전력(kW)", "계약종별"]

ENABLE_DEBUG_DUMP = True
HEADLESS = True
SCREEN_BASE = Path.cwd() / "screenshots"
DEBUG_BASE = Path.cwd() / "debug"

# ===== 유틸 =====
def sanitize_sheet(name: str) -> str:
    return re.sub(r"[\\/*?:\[\]]", "_", str(name).strip() or "NONAME")[:31]

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", str(name).strip() or "NONAME")

# ===== 드라이버 (Cloud 호환) =====
def build_driver() -> webdriver.Chrome:
    candidate_bins = [
        os.environ.get("GOOGLE_CHROME_BIN"),
        os.environ.get("CHROME_BIN"),
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
    ]
    chrome_bin = next((p for p in candidate_bins if p and os.path.exists(p)), None)

    opts = Options()
    if chrome_bin:
        opts.binary_location = chrome_bin

    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-software-rasterizer")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-features=VizDisplayCompositor")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    opts.add_experimental_option('useAutomationExtension', False)
    opts.add_argument("--log-level=3")

    candidate_drivers = [
        os.environ.get("CHROMEDRIVER"),
        "/usr/bin/chromedriver",
        "/usr/lib/chromium/chromedriver",
        "/usr/local/bin/chromedriver",
    ]
    driver_path = next((p for p in candidate_drivers if p and os.path.exists(p)), None)

    if driver_path:
        service = Service(executable_path=driver_path)
    else:
        service = Service()

    try:
        driver = webdriver.Chrome(service=service, options=opts)
    except Exception as e:
        raise RuntimeError(
            (
                "Chrome/Chromedriver 실행 실패. 다음을 확인하세요:\n"
                "- 서버에 Chrome/Chromium가 설치되어 있는지 (또는 GOOGLE_CHROME_BIN 환경변수 설정)\n"
                "- --no-sandbox / --disable-dev-shm-usage 플래그 적용 여부\n"
                "- 시스템 chromedriver와 chrome 버전 호환 여부\n"
                f"원본 오류: {e}"
            )
        )

    driver.set_page_load_timeout(60)
    return driver

# ===== 공통 유틸 =====
def reset_session(driver: Optional[webdriver.Chrome]) -> webdriver.Chrome:
    try:
        if driver:
            driver.quit()
    except Exception:
        pass
    time.sleep(0.5)
    return build_driver()

def wait_ready(driver, timeout=20):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

def wait_click(driver, by, value, timeout=15):
    elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    try:
        elem.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", elem)
    return elem

def wait_sendkeys(driver, by, value, text, timeout=15):
    elem = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    try:
        elem.clear()
    except Exception:
        pass
    elem.send_keys(text)
    return elem

# ===== 로그인 =====
def _is_logged_in(driver, timeout=20) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        if "intro.do" not in (driver.current_url or "").lower():
            return True
        time.sleep(0.5)
    return False

def run_once_with_credentials(driver, user_id: str, user_pw: str) -> Tuple[bool, Optional[str]]:
    try:
        driver.get(URL_INTRO)
        wait_ready(driver, 20)
        wait_sendkeys(driver, By.ID, "RSA_USER_ID", user_id)
        wait_sendkeys(driver, By.ID, "RSA_USER_PWD", user_pw)
        wait_click(driver, By.CSS_SELECTOR, 'input.intro_btn[type="button"][value="로그인"]')
        if _is_logged_in(driver, timeout=25):
            return True, None
        return False, "LOGIN_FAILED"
    except UnexpectedAlertPresentException:
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        return False, "UNEXPECTED_ALERT"
    except Exception as e:
        return False, repr(e)

# ===== 고객정보 =====
def goto_customer_info(driver):
    driver.get(URL_BASE + CUSTOMER_INFO_PATH)
    wait_ready(driver, 20)
    time.sleep(0.6)
    return True

def open_meter_tab(driver, timeout=10) -> bool:
    try:
        link = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, \"tabTable('3')\")]"))
        )
        driver.execute_script("arguments[0].click();", link)
    except TimeoutException:
        try:
            driver.execute_script("tabTable('3')")
        except Exception:
            return False

    end = time.time() + timeout
    while time.time() < end:
        try:
            if driver.find_elements(By.CSS_SELECTOR, "#table3"):
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return True

def fetch_three_fields(driver) -> Tuple[str, str, str, Optional[str]]:
    try:
        meter = driver.find_element(By.CSS_SELECTOR, "#table3 tbody tr:nth-child(1) td:nth-child(2)").text.strip()
        kw    = driver.find_element(By.CSS_SELECTOR, "div.table_info table tbody tr:nth-child(2) td:nth-child(4)").text.strip()
        ctype = driver.find_element(By.CSS_SELECTOR, "div.table_info table tbody tr:nth-child(2) td:nth-child(2)").text.strip()
        return meter, kw, ctype, None
    except Exception:
        return "", "", "", "FIELD_NOT_FOUND"

def center_mouse_and_screenshot(driver, sheet_name: str, cust_name: str):
    try:
        fname = sanitize_filename(cust_name) + ".png"
        save_path = SCREEN_BASE / sanitize_sheet(sheet_name) / fname
        save_path.parent.mkdir(parents=True, exist_ok=True)
        driver.save_screenshot(str(save_path))
        return True, str(save_path)
    except Exception as e:
        return False, repr(e)

def dump_debug_html(driver, sheet_name: str, cust_name: str):
    if not ENABLE_DEBUG_DUMP:
        return
    try:
        fname = sanitize_filename(cust_name) + "_page.html"
        path = DEBUG_BASE / sanitize_sheet(sheet_name) / fname
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(driver.page_source, encoding="utf-8", errors="ignore")
    except Exception:
        pass

# ===== 엑셀 =====
def ensure_workbook(path: Path):
    if not path.exists():
        Workbook().save(path)

def ensure_sheet_with_header(path: Path, sheet_name: str):
    sheet_name = sanitize_sheet(sheet_name)
    wb = load_workbook(path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(HEADERS)
    wb.save(path)
    wb.close()
    return sheet_name

def append_row(path: Path, sheet_name: str, row_values: list):
    sheet_name = sanitize_sheet(sheet_name)
    wb = load_workbook(path)
    ws = wb[sheet_name]
    ws.append(row_values)
    wb.save(path)
    wb.close()

# ===== 계정 처리 =====
def process_account(sheet_name: str, cust_name: str, user_id: str, user_pw: str) -> Tuple[str, str, str, bool]:
    driver = reset_session(None)
    ok, reason = run_once_with_credentials(driver, user_id, user_pw)
    if not ok:
        driver.quit()
        return "", "", "", False

    try:
        goto_customer_info(driver)
        center_mouse_and_screenshot(driver, sheet_name, cust_name)

        if not open_meter_tab(driver, timeout=10):
            dump_debug_html(driver, sheet_name, cust_name)
            driver.quit()
            return "", "", "", False

        meter, kw, ctype, err = fetch_three_fields(driver)
        driver.quit()
        if err:
            dump_debug_html(driver, sheet_name, cust_name)
            return "", "", "", False
        return meter, kw, ctype, True
    except Exception:
        dump_debug_html(driver, sheet_name, cust_name)
        driver.quit()
        return "", "", "", False

# ===== 메인 배치 =====
def read_excel_all_sheets(xlsx_path: Path) -> dict:
    xls = pd.ExcelFile(xlsx_path)
    return {s: pd.read_excel(xlsx_path, sheet_name=s, dtype=str).fillna("").astype(str) for s in xls.sheet_names}

def run_batch(excel_path: Path, out_xlsx_name: Path, progress_cb=None, log_cb=None):
    load_dotenv()
    if not excel_path.exists():
        raise FileNotFoundError(f"엑셀 없음: {excel_path.resolve()}")

    SCREEN_BASE.mkdir(parents=True, exist_ok=True)
    DEBUG_BASE.mkdir(parents=True, exist_ok=True)

    sheets = read_excel_all_sheets(excel_path)
    total_jobs = sum(
        len(df[(df.get("ID", "") != "") & (df.get("PW", "") != "")])
        for df in sheets.values()
        if "ID" in df and "PW" in df and "고객명" in df
    )

    ensure_workbook(out_xlsx_name)
    for s in sheets.keys():
        ensure_sheet_with_header(out_xlsx_name, s)

    success_cnt, fail_cnt, done = 0, 0, 0

    for sheet_name, df in sheets.items():
        if not all(c in df.columns for c in ("ID", "PW", "고객명")):
            continue
        safe_sheet = ensure_sheet_with_header(out_xlsx_name, sheet_name)
        for _, row in df.iterrows():
            user_id = str(row["ID"]).strip()
            user_pw = str(row["PW"]).strip()
            cust_name = str(row["고객명"]).strip()
            if not user_id or not user_pw:
                continue

            if log_cb:
                log_cb(f"처리 중: [{sheet_name}] {cust_name} ({user_id})")

            meter, kw, ctype, ok = process_account(sheet_name, cust_name, user_id, user_pw)
            if ok:
                success_cnt += 1
            else:
                fail_cnt += 1
            append_row(out_xlsx_name, safe_sheet, [sheet_name, cust_name, user_id, user_pw, meter, kw, ctype])

            done += 1
            if progress_cb and total_jobs:
                progress_cb(done / total_jobs)

    summary = f"총 작업 {total_jobs}건\n성공 {success_cnt}건, 실패 {fail_cnt}건\n결과: {out_xlsx_name.resolve()}"
    return summary, success_cnt, fail_cnt, total_jobs

# ===== Streamlit UI =====
st.set_page_config(page_title="파워플래너 정보 취합 (KEPCO)", layout="wide")
st.title("파워플래너 정보 취합 – Streamlit Cloud")

with st.sidebar:
    st.header("실행 옵션")
    HEADLESS = st.checkbox("Headless", value=True)
    ENABLE_DEBUG_DUMP = st.checkbox("디버그 HTML 저장", value=True)

    up = st.file_uploader("입력 엑셀 업로드", type=["xlsx"])
    excel_path = None
    if up is not None:
        tmp = Path(f"uploaded_{int(time.time())}.xlsx")
        with open(tmp, "wb") as f:
            f.write(up.getbuffer())
        excel_path = tmp
        st.success(f"업로드됨: {tmp}")
    else:
        manual = st.text_input("또는 로컬 경로 입력", value="20251010고객사id,pw2.xlsx")
        if manual:
            excel_path = Path(manual)

out_name = f"(자원등록용)파워플래너 정보 취합 초안_{datetime.now().strftime('%Y%m%d')}.xlsx"
out_path = Path(out_name)

col1, col2 = st.columns([1, 2])
run_btn = col1.button("실행")
log_box = st.empty()
progress = st.progress(0.0)
result_box = st.empty()

def _log(msg):
    prev = st.session_state.get("_logs", "")
    new = prev + msg + "\n"
    st.session_state["_logs"] = new
    log_box.text(new)

if run_btn:
    if not excel_path:
        st.error("엑셀 파일을 업로드하거나 경로를 입력하세요.")
        st.stop()

    st.session_state["_logs"] = ""
    progress.progress(0.0)

    SCREEN_BASE = Path.cwd() / "screenshots"
    DEBUG_BASE = Path.cwd() / "debug"

    try:
        summary, ok_cnt, fail_cnt, total = run_batch(
            excel_path=excel_path,
            out_xlsx_name=out_path,
            progress_cb=lambda v: progress.progress(min(max(v, 0.0), 1.0)),
            log_cb=_log,
        )
        st.success("작업 완료")
        result_box.code(summary)
        if out_path.exists():
            with open(out_path, "rb") as f:
                st.download_button("결과 엑셀 다운로드", data=f.read(), file_name=out_path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
