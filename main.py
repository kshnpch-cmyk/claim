import os
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")

def calculate_dates():
    today = datetime.now()
    start_date = today - timedelta(days=20) # 오늘 포함 3주(21일)
    
    end_dt_str = today.strftime("%Y/%m/%d")
    start_dt_str = start_date.strftime("%Y/%m/%d")
    daterange_str = f"{start_dt_str} - {end_dt_str}"
    
    return start_dt_str, end_dt_str, daterange_str

def run_automation():
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

        print("[6/6] 그리드 우클릭 ➔ 팝업창 파일명 입력 ➔ 다운로드 실행 중...")
        
        # 1. 셀 우클릭
        target_cell = page.locator('td:has-text("RTN"), th:has-text("반품번호")').first
        target_cell.click(button="right")
        page.wait_for_timeout(500)

        # 2. 컨텍스트 메뉴의 '엑셀다운로드' 클릭 (팝업이 뜸)
        page.click('text="엑셀다운로드"')
        page.wait_for_timeout(1000)

        # 3. 팝업창 폼 제어 (placeholder="파일명" 또는 type="text" 입력창 찾기)
        filename_input = page.locator('input[placeholder="파일명"]').first
        filename_input.fill('claim_data')

        # 📸 [확인용] 팝업창에 파일명이 잘 채워졌는지 캡처
        popup_screenshot = os.path.join(output_dir, "popup_filled_result.png")
        page.screenshot(path=popup_screenshot, full_page=True)
        print(f"팝업 입력 캡처 완료: {popup_screenshot}")

        # 4. 팝업창 내 [다운로드] 버튼 클릭 및 파일 수신
        with page.expect_download() as download_info:
            page.click('button:has-text("다운로드")')
            
        download = download_info.value
        file_path = os.path.join(output_dir, download.suggested_filename)
        download.save_as(file_path)
        
        print(f"🎉 엑셀 파일 다운로드 성공: {file_path}")

        browser.close()

if __name__ == "__main__":
    run_automation()
