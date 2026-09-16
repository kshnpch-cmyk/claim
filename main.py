import os
import json
import pandas as pd
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# GitHub Secrets 환경 변수 로드
# ==========================================
LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")

def download_and_process():
    download_dir = "./downloads"
    os.makedirs(download_dir, exist_ok=True)

    with sync_playwright() as p:
        # 1. Chromium 브라우저 실행 (Headless)
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("[1/3] 더본코리아 OMS 로그인 페이지 접속 중...")
        page.goto(LOGIN_URL)
        page.wait_for_load_state("networkidle")

        # 2. 알려주신 폼 엘리먼트에 데이터 입력
        print("[2/3] 로그인 정보 입력 중...")
        page.fill('#companyCd', COMPANY_CD)
        page.fill('#userId', USER_ID)
        page.fill('#userPw', USER_PW)
        
        # 로그인 버튼 클릭 및 완료 대기
        page.click('#btnLogin')
        page.wait_for_load_state("networkidle")

        # 3. 클레임 다운로드 페이지 이동 및 파일 저장
        # ⚠️ 로그인 성공 후 이동할 클레임 페이지 주소 및 다운로드 버튼 selector는 추후 수정 필요합니다.
        print("[3/3] 클레임 파일 다운로드 시도...")
        # page.goto("https://admin.theborn.co.kr/oms-manager/claim/list")  # 예시 주소
        
        # 다운로드 실행 예시
        # with page.expect_download() as download_info:
        #     page.click('#btnExcelDownload') # 다운로드 버튼 클릭
        # download = download_info.value
        # file_path = os.path.join(download_dir, download.suggested_filename)
        # download.save_as(file_path)

        browser.close()
        return None

def upload_to_sheets(file_path):
    if not file_path or not os.path.exists(file_path):
        print("업로드할 파일이 없습니다.")
        return

    # GCP 서비스 계정 인증
    sa_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(credentials)
    
    sheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

    # 파일 읽기 및 업로드 로직 (CSV/XLSX)
    df = pd.read_excel(file_path) if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
    df = df.fillna("")

    existing_ids = set(sheet.col_values(1)[1:])
    rows_to_append = []

    for row in df.to_dict(orient="records"):
        claim_id = str(row.get("클레임ID", ""))
        if claim_id and claim_id not in existing_ids:
            rows_to_append.append([
                claim_id,
                str(row.get("접수일", "")),
                str(row.get("고객명", "")),
                str(row.get("내용", ""))
            ])

    if rows_to_append:
        sheet.append_rows(rows_to_append)
        print(f"구글 시트에 {len(rows_to_append)}건의 신규 클레임 추가 완료.")

if __name__ == "__main__":
    file_path = download_and_process()
    if file_path:
        upload_to_sheets(file_path)
