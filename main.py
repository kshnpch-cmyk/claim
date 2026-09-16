import os
import re
import glob
import time
import json
import warnings
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# openpyxl CellStyle count 속성 오류 방지 패치
import openpyxl.styles.cell_style
_original_cell_style_init = openpyxl.styles.cell_style.CellStyle.__init__

def _patched_cell_style_init(self, *args, **kwargs):
    kwargs.pop('count', None)
    _original_cell_style_init(self, *args, **kwargs)

openpyxl.styles.cell_style.CellStyle.__init__ = _patched_cell_style_init

warnings.filterwarnings('ignore')

# 💡 환경 변수 설정
LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"
WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxFmxSGgRypmiO-bi4Xrs52r-mqWdV-JGHVRU762_txBxR3d3LXn7XDiJo6KbtMbfe1/exec"

OMS_COMPANY_CODE = os.environ.get("COMPANY_CD", "1000").strip()
OMS_ID = os.environ.get("USER_ID", "").strip()
OMS_PW = os.environ.get("USER_PW", "").strip()

output_dir = "./output"
os.makedirs(output_dir, exist_ok=True)

# 1. KST 날짜 계산 (과거 3주간)
KST = timezone(timedelta(hours=9))
now_kst = datetime.now(KST)

start_date_obj = now_kst - timedelta(days=20)
end_date_obj = now_kst

target_start_date = start_date_obj.strftime("%Y/%m/%d")
target_end_date = end_date_obj.strftime("%Y/%m/%d")
daterange_str = f"{target_start_date} - {target_end_date}"

# 2. 크롬 브라우저 다운로드 설정 (Headless)
download_dir = os.getcwd()
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
driver.execute_script("return null;")
driver.execute("send_command", params)


def send_data_to_google_sheet(excel_file_path):
    if not excel_file_path or not os.path.exists(excel_file_path):
        print("⚠️ 전송할 엑셀 파일이 존재하지 않습니다.", flush=True)
        return

    print("🚀 Google Apps Script 웹앱으로 데이터 전송 시작...", flush=True)

    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except Exception:
        df = pd.read_csv(excel_file_path)

    df = df.fillna("")
    rows_data = df.values.tolist()

    if not rows_data:
        print("⚠️ 엑셀 파일 내 데이터가 없습니다.", flush=True)
        return

    try:
        response = requests.post(
            WEBAPP_URL,
            data=json.dumps(rows_data),
            headers={"Content-Type": "application/json"}
        )
        
        res_json = response.json()
        if res_json.get("result") == "success":
            print(f"✅ 구글 시트 업로드 성공! (신규 추가: {res_json.get('added')}건)", flush=True)
        else:
            print(f"❌ Apps Script 오류: {res_json.get('error')}", flush=True)

    except Exception as e:
        print(f"❌ HTTP 요청 실패: {e}", flush=True)


try:
    print(f"[{now_kst.strftime('%Y-%m-%d %H:%M:%S')}] 품질 클레임 관리 자동 수집 시작", flush=True)
    print(f"조회 지정 기간 (과거 3주): {target_start_date} ~ {target_end_date}", flush=True)

    # 3. 로그인 진행
    print("[1/6] OMS 로그인 진행 중...", flush=True)
    driver.get(LOGIN_URL)
    time.sleep(2)
    driver.find_element(By.ID, 'companyCd').send_keys(OMS_COMPANY_CODE)
    driver.find_element(By.ID, 'userId').send_keys(OMS_ID)
    driver.find_element(By.ID, 'userPw').send_keys(OMS_PW + Keys.ENTER)
    time.sleep(4)

    # 4. 메뉴 순차적 클릭 이동
    print("[2/6] 메뉴 계층 탐색 중...", flush=True)
    
    try:
        bor_menu = driver.find_element(By.CSS_SELECTOR, "a[data-menu-id='BOR']")
        driver.execute_script("arguments[0].click();", bor_menu)
        time.sleep(2)
    except Exception:
        pass

    try:
        return_folder = driver.find_element(By.XPATH, "//*[contains(text(), '반품관리')]")
        driver.execute_script("arguments[0].click();", return_folder)
        time.sleep(1)
    except Exception:
        pass

    claim_menu = driver.find_element(By.XPATH, "//*[contains(text(), '품질 클레임 관리')]")
    driver.execute_script("arguments[0].click();", claim_menu)
    time.sleep(4)

    # 💡 5. iframe 전환 (품질 클레임 관리 탭 화면 내부로 스위치)
    driver.switch_to.default_content() # 기본 프레임 초기화
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if iframes:
        print(f"👉 {len(iframes)}개의 iframe 감지됨. 콘텐츠 iframe으로 전환합니다.", flush=True)
        # 작업 탭에 해당하는 iframe(보통 마지막 또는 활성화된 iframe)으로 전환
        driver.switch_to.frame(iframes[-1])

    driver.save_screenshot(os.path.join(output_dir, "step2_page_loaded.png"))

    # 6. 날짜 범위 입력 (BOR210)
    print("[3/6] 3주간 날짜 범위 설정 중...", flush=True)
    js_script = f"""
        var rangeInput = document.getElementsByName('BOR210_daterange')[0];
        if(rangeInput) {{
            rangeInput.value = '{daterange_str}';
            rangeInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }}
        var startInput = document.getElementById('BOR210_startDt');
        var endInput = document.getElementById('BOR210_endDt');
        if(startInput) startInput.value = '{target_start_date}';
        if(endInput) endInput.value = '{target_end_date}';
    """
    driver.execute_script(js_script)
    time.sleep(2)

    # 7. [조회] 실행 (F2 키 또는 버튼)
    print("[4/6] [조회] 실행...", flush=True)
    try:
        driver.find_element(By.CSS_SELECTOR, "button.form_btn_search[data-shortcut='F2']").click()
    except Exception:
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
    time.sleep(6)

    driver.save_screenshot(os.path.join(output_dir, "step4_search_clicked.png"))

    # 8. 우클릭 대상 탐색 및 우클릭 (RealGrid 및 TD 요소 타겟팅)
    print("[5/6] 그리드 영역 우클릭 메뉴 호출...", flush=True)
    header_element = None
    
    # 우클릭 타겟 셀렉터 순서 (데이터 셀 ➔ 그리드 canvas ➔ 테이블 캔버스)
    selectors = [
        "//*[contains(text(), 'RTN')]", 
        "//td[contains(@class, 'rg')]", 
        "//div[contains(@class, 'realgrid')]",
        "//canvas",
        "//td",
        "//th"
    ]
    
    for sel in selectors:
        try:
            elements = driver.find_elements(By.XPATH, sel)
            if elements:
                header_element = elements[0]
                print(f"👉 우클릭 타겟 발견: XPath('{sel}')", flush=True)
                break
        except Exception:
            pass

    if not header_element:
        raise Exception("그리드 우클릭 대상 요소(테이블/셀)를 화면에서 찾을 수 없습니다.")

    # 마우스 우클릭 수행
    ActionChains(driver).context_click(header_element).perform()
    time.sleep(2)

    driver.save_screenshot(os.path.join(output_dir, "step5_context_menu.png"))

    # 9. '엑셀다운로드' 메뉴 클릭 및 SweetAlert 팝업 제어
    print("[6/6] 엑셀다운로드 실행 중...", flush=True)
    
    # 컨텍스트 메뉴 내 '엑셀다운로드' 찾기 (iframe 내부 또는 최상위 메인 프레임 탐색)
    excel_btn = None
    try:
        excel_btn = driver.find_element(By.XPATH, "//*[contains(text(), '엑셀다운로드')]")
    except Exception:
        # iframe 바깥 메인 DOM에 우클릭 메뉴가 뜬 경우 상위로 이동해서 검색
        driver.switch_to.default_content()
        excel_btn = driver.find_element(By.XPATH, "//*[contains(text(), '엑셀다운로드')]")

    driver.execute_script("arguments[0].click();", excel_btn)
    time.sleep(2)

    # SweetAlert 파일명 입력 팝업 처리
    try:
        swal_input = driver.find_element(By.CSS_SELECTOR, "input.swal2-input")
        swal_input.clear()
        swal_input.send_keys("claim_download")
    except Exception:
        pass
    time.sleep(1)

    download_btn = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
    driver.execute_script("arguments[0].click();", download_btn)
    time.sleep(5)

    try:
        ok_btn = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
        driver.execute_script("arguments[0].click();", ok_btn)
    except Exception:
        pass
    time.sleep(5)

    # 10. 파일 수신 및 데이터 전송
    list_of_files = glob.glob(os.path.join(download_dir, '*.xlsx')) or glob.glob(os.path.join(download_dir, '*.xls'))
    if not list_of_files:
        raise Exception("다운로드 파일 수신 실패")

    latest_file = max(list_of_files, key=os.path.getctime)
    print(f"📥 수신된 파일: {latest_file}", flush=True)

    send_data_to_google_sheet(latest_file)

    if os.path.exists(latest_file):
        os.remove(latest_file)

except Exception as e:
    print(f"❌ 오류 발생: {e}", flush=True)
    try:
        driver.save_screenshot(os.path.join(output_dir, "error_screenshot.png"))
        print(f"📸 에러 화면 캡처 완료: {os.path.join(output_dir, 'error_screenshot.png')}", flush=True)
    except Exception:
        pass
finally:
    driver.quit()
