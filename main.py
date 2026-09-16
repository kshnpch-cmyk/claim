import os
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")

def calculate_dates():
    # 오늘 날짜 (종료일)
    today = datetime.now()
    # 3주 전 날짜 (시작일: 오늘 포함 21일 전)
    start_date = today - timedelta(days=20)
    
    # YYYY/MM/DD 포맷팅
    end_dt_str = today.strftime("%Y/%m/%d")
    start_dt_str = start_date.strftime("%Y/%m/%d")
    
    # daterange 표시용 포맷 (예: 2026/08/27 - 2026/09/16)
    daterange_str = f"{start_dt_str} - {end_dt_str}"
    
    return start_dt_str, end_dt_str, daterange_str

def test_login_and_download():
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)
    
    # 3주간의 날짜 범위 계산
    start_dt, end_dt, daterange = calculate_dates()
    print(f"[날짜 설정] 조회 기간: {start_dt} ~ {end_dt} ({daterange})")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("[1/5] 로그인 페이지 접속 중...")
        page.goto(LOGIN_URL)
        page.wait_for_load_state("networkidle")

        print("[2/5] 로그인 정보 입력 중...")
        page.fill('#companyCd', COMPANY_CD)
        page.fill('#userId', USER_ID)
        page.fill('#userPw', USER_PW)
        
        page.click('#btnLogin')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[3/5] '품질 클레임 관리' 메뉴 이동...")
        page.click('#side-menu-BOR210')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[4/5] 3주간 조회 기간 설정 입력 중...")
        # 1. daterange input 채우기
        page.fill('input[name="BOR210_daterange"]', daterange)
        
        # 2. hidden input 값 강제 변경 (JavaScript 바인딩 처리)
        page.evaluate(f'''() => {{
            const startInput = document.getElementById("BOR210_startDt");
            const endInput = document.getElementById("BOR210_endDt");
            if (startInput) startInput.value = "{start_dt}";
            if (endInput) endInput.value = "{end_dt}";
        }}''')

        # 📸 [확인용] 날짜 설정 반영 및 페이지 화면 캡처
        screenshot_path = os.path.join(output_dir, "claim_date_set_result.png")
        page.screenshot(path=screenshot_path, full_page=True)
        print(f"날짜 적용 화면 캡처 완료: {screenshot_path}")

        print("[5/5] 조회 버튼 클릭 및 엑셀 다운로드 시도 대기...")
        # ⚠️ 조회 버튼 클릭이나 엑셀 다운로드 버튼 element를 확인하여 아래에 추가하면 됩니다.
        # page.click('#btnSearch') # 예시 조회버튼

        browser.close()

if __name__ == "__main__":
    test_login_and_download()
