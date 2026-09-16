import os
import json
import requests
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

# ==========================================
# 환경 변수 및 설정
# ==========================================
LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"
WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxFmxSGgRypmiO-bi4Xrs52r-mqWdV-JGHVRU762_txBxR3d3LXn7XDiJo6KbtMbfe1/exec"

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")


def calculate_dates():
    today = datetime.now()
    start_date = today - timedelta(days=20)  # 오늘 포함 3주(21일)
    
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

        print("[1/6] OMS 로그인 페이지 접속...")
        page.goto(LOGIN_URL)
        page.wait_for_load_state("networkidle")

        print("[2/6] 로그인 정보 입력...")
        page.fill('#companyCd', COMPANY_CD)
        page.fill('#userId', USER_ID)
        page.fill('#userPw', USER_PW)
        
        page.click('#btnLogin')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)

        print("[3/6] '품질 클레임 관리' 페이지 이동...")
        page.goto("https://admin.theborn.co.kr/fms-manager/rtn-approval-manage")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[4/6] 3주간 날짜 범위 설정...")
        page.wait_for_selector('input[name="BOR210_daterange"]', timeout=10000)
        page.fill('input[name="BOR210_daterange"]', daterange)
        
        page.evaluate(f'''() => {{
            const startInput = document.getElementById("BOR210_startDt");
            const endInput = document.getElementById("BOR210_endDt");
            if (startInput) startInput.value = "{start_dt}";
            if (endInput) endInput.value = "{end_dt}";
        }}''')

        print("[5/6] [조회] 버튼 클릭...")
        page.click('button:has-text("조회"), a:has-text("조회"), input[value="조회"]')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[6/6] 그리드 우클릭 ➔ 엑셀다운로드 팝업 ➔ 다운로드...")
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


def send_data_to_google_sheet(excel_file_path):
    if not excel_file_path or not os.path.exists(excel_file_path):
        print("전송할 엑셀 파일이 존재하지 않습니다.")
        return

    print("🚀 Google Apps Script 웹앱으로 데이터 전송 시작...")

    try:
        df = pd.read_excel(excel_file_path)
    except Exception:
        df = pd.read_csv(excel_file_path)

    df = df.fillna("")
    rows_data = df.values.tolist()

    if not rows_data:
        print("엑셀 파일 내 데이터가 없습니다.")
        return

    try:
        response = requests.post(
            WEBAPP_URL,
            data=json.dumps(rows_data),
            headers={"Content-Type": "application/json"}
        )
        
        res_json = response.json()
        if res_json.get("result") == "success":
            print(f"✅ 구글 시트 업로드 성공! (신규 추가: {res_json.get('added')}건)")
        else:
            print(f"❌ Apps Script 오류: {res_json.get('error')}")

    except Exception as e:
        print(f"❌ HTTP 요청 실패: {e}")


if __name__ == "__main__":
    print("=== 자동화 스크립트 실행 시작 ===")
    
    # Secrets 환경변수 체크
    if not COMPANY_CD or not USER_ID or not USER_PW:
        print("❌ 오류: GitHub Secrets (COMPANY_CD, USER_ID, USER_PW) 환경변수가 설정되지 않았습니다.")
    else:
        try:
            downloaded_path = download_excel_file()
            send_data_to_google_sheet(downloaded_path)
        except Exception as e:
            print(f"❌ 실행 중 오류 발생: {e}")
            
    print("=== 자동화 스크립트 실행 종료 ===")
