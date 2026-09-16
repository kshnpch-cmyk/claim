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
        page.wait_for_timeout(3000) # 로그인 후 리다이렉트 완전 대기

        print("[3/6] '품질 클레임 관리' 페이지로 직접 이동...")
        # 💡 [핵심 수정] 메뉴 클릭 대신 URL 직접 접속으로 Timeout 에러 원천 차단
        page.goto("https://admin.theborn.co.kr/fms-manager/rtn-approval-manage")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        print("[4/6] 3주간 날짜 범위 설정...")
        # daterange input 존재 여부 확인 및 입력
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
