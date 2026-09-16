import os
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://admin.theborn.co.kr/oms-manager/login"

COMPANY_CD = os.environ.get("COMPANY_CD")
USER_ID = os.environ.get("USER_ID")
USER_PW = os.environ.get("USER_PW")

def test_login_and_download():
    # 캡처 및 다운로드 저장용 디렉토리 생성
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)

    with sync_playwright() as p:
        # 1. Chromium 브라우저 실행
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("[1/4] 로그인 페이지 접속 중...")
        page.goto(LOGIN_URL)
        page.wait_for_load_state("networkidle")

        print("[2/4] 로그인 정보 입력 중...")
        page.fill('#companyCd', COMPANY_CD)
        page.fill('#userId', USER_ID)
        page.fill('#userPw', USER_PW)
        
        # 로그인 버튼 클릭
        page.click('#btnLogin')
        page.wait_for_load_state("networkidle")
        
        # 페이지 이동 후 2초 추가 대기 (메인 화면 로딩 확보)
        page.wait_for_timeout(2000)

        # 📸 [확인용] 로그인 성공 화면 캡처 저장
        screenshot_path = os.path.join(output_dir, "login_result.png")
        page.screenshot(path=screenshot_path, full_page=True)
        print(f"로그인 성공 여부 화면 캡처 완료: {screenshot_path}")

        print("[3/4] 클레임 관리/조회 페이지 이동...")
        # ⚠️ 로그인 후 실제로 이동해야 하는 클레임 목록 페이지 URL 입력 필요
        # page.goto("https://admin.theborn.co.kr/oms-manager/claim/list")
        # page.wait_for_load_state("networkidle")

        print("[4/4] 엑셀 다운로드 버튼 클릭 시도...")
        # ⚠️ 실제 페이지의 엑셀 다운로드 버튼 ID 또는 Selector 지정 필요
        # with page.expect_download() as download_info:
        #     page.click('#btnExcelDownload') # 예: 엑셀다운로드 버튼 클릭
        # 
        # download = download_info.value
        # download_path = os.path.join(output_dir, download.suggested_filename)
        # download.save_as(download_path)
        # print(f"엑셀 파일 다운로드 성공: {download_path}")

        browser.close()

if __name__ == "__main__":
    test_login_and_download()
