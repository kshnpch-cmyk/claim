import os
import json
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 환경 변수 및 상수 설정
# ==========================================
LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"
SPREADSHEET_KEY = "16lfX5WG4cPnRLZe2k8quGLWA_oF1xpI79ppAE4VMzGk" # 제공해주신 시트 Key

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")


def calculate_dates():
    today = datetime.now()
    start_date = today - timedelta(days=20) # 오늘 포함 3주(21일)
    
    end_dt_str = today.strftime("%Y/%m/%d")
    start_dt_str = start_date.strftime("%Y/%m/%d")
    daterange_str = f"{start_dt_str} - {end_dt_str}"
    
    return start_dt_str, end_dt_str, daterange_str


def download_excel_file():
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)
    
    start_dt, end_dt, daterange = calculate_dates()
    print(f"[날짜 설정] 조회 기간: {start_dt} ~ {end_dt}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("[1/6] 로그인 페이지 접속 중...")
        page.goto(LOGIN_URL)
        page.wait_for_load_state("networkidle")

        print("[2/6] 로그인 정보 입력 중...")
        page.fill('#companyCd', COMPANY_CD)
        page.fill('#userId', USER_ID)
        page.fill('#userPw', USER_PW)
        
        page.click('#btnLogin')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[3/6] '품질 클레임 관리' 메뉴 이동...")
        page.click('#side-menu-BOR210')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[4/6] 3주간 날짜 범위 설정 중...")
        page.fill('input[name="BOR210_daterange"]', daterange)
        
        page.evaluate(f'''() => {{
            const startInput = document.getElementById("BOR210_startDt");
            const endInput = document.getElementById("BOR210_endDt");
            if (startInput) startInput.value = "{start_dt}";
            if (endInput) endInput.value = "{end_dt}";
        }}''')

        print("[5/6] 우상단 [조회] 버튼 클릭...")
        page.click('button:has-text("조회"), a:has-text("조회"), input[value="조회"]')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[6/6] 그리드 우클릭 ➔ 팝업창 파일명 입력 ➔ 다운로드...")
        target_cell = page.locator('td:has-text("RTN"), th:has-text("반품번호")').first
        target_cell.click(button="right")
        page.wait_for_timeout(500)

        page.click('text="엑셀다운로드"')
        page.wait_for_timeout(1000)

        filename_input = page.locator('input[placeholder="파일명"]').first
        filename_input.fill('claim_data')

        with page.expect_download() as download_info:
            page.click('button:has-text("다운로드")')
            
        download = download_info.value
        file_path = os.path.join(output_dir, download.suggested_filename)
        download.save_as(file_path)
        
        print(f"🎉 엑셀 다운로드 완료: {file_path}")
        browser.close()
        
        return file_path


def update_google_sheet(excel_file_path):
    if not excel_file_path or not os.path.exists(excel_file_path):
        print("업데이트할 엑셀 파일이 존재하지 않습니다.")
        return

    print("📊 [구글 시트] 데이터 변환 및 업로드 시작...")

    # 1. 다운로드받은 엑셀 파일 읽기
    try:
        df = pd.read_excel(excel_file_path)
    except Exception:
        df = pd.read_csv(excel_file_path)

    # 데이터 내 빈값(NaN)을 빈 문자열("")로 대체
    df = df.fillna("")

    # 2. Google Sheets API 인증 연결
    sa_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(credentials)

    # 지정하신 시트의 첫 번째 워크시트(gid=0) 열기
    sheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

    # 3. 중복 저장 방지 (기존 시트 1열에 존재하는 반품번호/클레임ID 가져오기)
    existing_ids = set(sheet.col_values(1)[1:])

    # 4. 업로드할 행 추출 (중복 제외)
    rows_to_append = []
    headers = list(df.columns)

    # 시트가 아예 비어있는 경우 헤더(컬럼명)를 먼저 집어넣음
    if len(sheet.get_all_values()) == 0:
        sheet.append_row(headers)

    for record in df.to_dict(orient="records"):
        # 엑셀의 첫 번째 컬럼 값을 고유 식별키로 사용
        first_col_val = str(record[headers[0]])
        
        if first_col_val and first_col_val not in existing_ids:
            # 엑셀 행 전체 값을 리스트 형태로 추출
            row_data = [str(record[col]) for col in headers]
            rows_to_append.append(row_data)

    # 5. 구글 시트에 일괄 추가 (append)
    if rows_to_append:
        sheet.append_rows(rows_to_append)
        print(f"✅ 구글 스프레드시트에 신규 클레임 {len(rows_to_append)}건 업로드 완료!")
    else:
        print("ℹ️ 신규로 추가할 클레임 데이터가 없습니다. (모두 기존 등록건)")


if __name__ == "__main__":
    downloaded_path = download_excel_file()
    if downloaded_path and GOOGLE_SERVICE_ACCOUNT_JSON:
        update_google_sheet(downloaded_path)
