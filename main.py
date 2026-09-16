import os
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. 환경 변수 불러오기 (GitHub Secrets 관리)
# ==========================================
ADMIN_ID = os.environ.get("ADMIN_ID")
ADMIN_PW = os.environ.get("ADMIN_PW")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")

GMAIL_USER = os.environ.get("GMAIL_USER")          # 발신용 Gmail 주소
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")  # Gmail 앱 비밀번호
ALERT_RECEIVER = os.environ.get("ALERT_RECEIVER")  # 알림 받을 이메일 주소


# ==========================================
# 2. 웹 크롤링 (Playwright)
# ==========================================
def fetch_claims():
    claims_data = []
    
    with sync_playwright() as p:
        # GitHub Actions(헤드리스 서버) 환경에 맞춰 headless=True 설정
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        print("[1/4] admin.theborn.co.kr 로그인 시도...")
        page.goto("https://admin.theborn.co.kr")
        
        # ⚠️ 실제 로그인 폼 셀렉터로 수정 필요
        page.fill('input[name="username"]', ADMIN_ID)
        page.fill('input[name="password"]', ADMIN_PW)
        page.click('button[type="submit"]')
        page.wait_for_load_state("networkidle")

        print("[2/4] 클레임 조회 페이지 이동...")
        # ⚠️ 실제 클레임 관리 URL로 변경 필요
        page.goto("https://admin.theborn.co.kr/claim/list")
        page.wait_for_selector("table.claim-list-table")  # 데이터 테이블 로딩 대기

        # 테이블에서 클레임 데이터 추출
        rows = page.query_selector_all("table.claim-list-table tbody tr")
        for row in rows:
            cols = row.query_selector_all("td")
            if len(cols) >= 4:
                claim_id = cols[0].inner_text().strip()
                date = cols[1].inner_text().strip()
                customer = cols[2].inner_text().strip()
                content = cols[3].inner_text().strip()

                claims_data.append({
                    "id": claim_id,
                    "date": date,
                    "customer": customer,
                    "content": content
                })

        browser.close()
        
    print(f"총 {len(claims_data)}건의 클레임을 수집했습니다.")
    return claims_data


# ==========================================
# 3. Google Sheets 저장 (gspread)
# ==========================================
def update_google_sheet(claims_data):
    if not claims_data:
        print("업데이트할 데이터가 없습니다.")
        return []

    print("[3/4] 구글 스프레드시트 기록 중...")
    
    # GCP Service Account 인증
    sa_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(credentials)

    # 스프레드시트 열기
    sheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

    # 기존 ID 가져와서 중복 수집 방지
    existing_ids = set(sheet.col_values(1)[1:])  # 첫 번째 열(클레임 ID)의 기존 데이터
    
    new_claims = []
    rows_to_append = []

    for claim in claims_data:
        if claim["id"] not in existing_ids:
            new_claims.append(claim)
            rows_to_append.append([
                claim["id"],
                claim["date"],
                claim["customer"],
                claim["content"]
            ])

    if rows_to_append:
        sheet.append_rows(rows_to_append)
        print(f"신규 클레임 {len(rows_to_append)}건 등록 완료.")
    else:
        print("신규 클레임이 없습니다.")

    return new_claims


# ==========================================
# 4. Gmail 자동 발송 (SMTP)
# ==========================================
def send_email_notification(new_claims):
    if not new_claims:
        print("신규 클레임이 없어 메일을 발송하지 않습니다.")
        return

    print("[4/4] 신규 클레임 알림 메일 발송 중...")

    # 메일 본문 구성 (HTML 타블 구조)
    table_rows = ""
    for c in new_claims:
        table_rows += f"""
        <tr>
            <td style="border: 1px solid #ddd; padding: 8px;">{c['id']}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">{c['date']}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">{c['customer']}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">{c['content']}</td>
        </tr>
        """

    html_content = f"""
    <html>
      <body>
        <h2>🚨 [신규 클레임 알림] {len(new_claims)}건의 새로운 클레임이 접수되었습니다.</h2>
        <table style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px;">클레임 ID</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">접수일</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">고객명</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">내용</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
      </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[더본코리아] 신규 클레임 알림 ({len(new_claims)}건)"
    msg["From"] = GMAIL_USER
    msg["To"] = ALERT_RECEIVER
    msg.attach(MIMEText(html_content, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_USER, ALERT_RECEIVER, msg.as_string())
        print("메일 발송 완료!")
    except Exception as e:
        print(f"메일 발송 실패: {e}")


# ==========================================
# 메인 실행
# ==========================================
if __name__ == "__main__":
    claims = fetch_claims()
    new_added_claims = update_google_sheet(claims)
    send_email_notification(new_added_claims)
