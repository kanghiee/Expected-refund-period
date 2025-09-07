"""
반품 현황 자동화 + 환불 안내 알림톡 발송 시스템
--------------------------------------------------
1. Selenium으로 택배사(CJ Logistics) 사이트 자동 로그인 및 엑셀 다운로드
2. 다운로드한 반품 데이터 Excel → Google Sheets 업데이트
3. 배송완료 건을 탐색하여 자동으로 환불 예정일 계산
4. 환불 예정일 안내 알림톡(Lunasoft API) 발송 + 발송 여부 기록
"""

import os
import time
import glob
import imaplib
import email
from email import policy
from bs4 import BeautifulSoup
import pandas as pd
import gspread
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# ======================== 환경 변수 로드 ========================
load_dotenv()
CNPLUS_ID = os.getenv("CNPLUS_ID")      # CJ Logistics 로그인 ID
CNPLUS_PW = os.getenv("CNPLUS_PW")      # CJ Logistics 로그인 PW
NAVER_EMAIL = os.getenv("NAVER_EMAIL")  # 네이버 메일 ID
NAVER_PW = os.getenv("NAVER_PW")        # 네이버 메일 앱 비밀번호
GOOGLE_KEY_PATH = os.getenv("GOOGLE_KEY_PATH")  # 구글 서비스 계정 키 파일
SPREADSHEET_URL = os.getenv("SPREADSHEET_URL")  # 구글 시트 URL
LUNASOFT_API_KEY = os.getenv("LUNASOFT_API_KEY")

print("🚀 실행 시작:", datetime.now())

# ======================== 1. Selenium 로그인 & OTP 처리 ========================
options = webdriver.ChromeOptions()
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920x1080")
driver = webdriver.Chrome(options=options)

driver.get("https://loisparcelp.cjlogistics.com/index.do")
time.sleep(2)

# ID / PW 입력 후 로그인
driver.find_element(By.CSS_SELECTOR, "input[type='text']").send_keys(CNPLUS_ID)
driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(CNPLUS_PW)
driver.find_element(By.ID, "uuid-3r").click()

# 이메일 OTP 가져오기 (네이버 IMAP)
imap = imaplib.IMAP4_SSL("imap.naver.com")
imap.login(NAVER_EMAIL, NAVER_PW)
imap.select("INBOX")
resp, data = imap.uid("search", None, "ALL")
latest_email = data[0].split()[-1]
result, data = imap.uid("fetch", latest_email, "(RFC822)")
raw_email = data[0][1]
email_message = email.message_from_bytes(raw_email, policy=policy.default)
soup = BeautifulSoup(email_message.get_payload()[0].get_payload(decode=True), "html.parser")
auth_code = soup.find("font", {"size": "7"}).get_text()

# OTP 입력
driver.find_element(By.CSS_SELECTOR, "input.cl-text").send_keys(auth_code)
driver.find_element(By.CSS_SELECTOR, ".btn-login a").click()

print("✅ 로그인 & OTP 완료")

# ======================== 2. 반품 Excel 다운로드 ========================
# (달력 선택, 조회, 엑셀 다운로드 → 생략 가능, 실제 동작 코드는 Private repo에 유지)

# ======================== 3. Excel 데이터 → Google Sheets 업데이트 ========================
def get_latest_excel_file(download_path="./downloads"):
    excel_files = glob.glob(os.path.join(download_path, "*.xlsx"))
    return max(excel_files, key=os.path.getctime)

def connect_google_sheet_by_url(spreadsheet_url, worksheet_index=0):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_KEY_PATH, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(spreadsheet_url).get_worksheet(worksheet_index)

# ✅ 최신 Excel 불러오기 → 시트 업데이트
def update_google_sheet_from_excel():
    latest_file = get_latest_excel_file()
    df = pd.read_excel(latest_file, skiprows=2, usecols=range(1, 16), engine="openpyxl")
    df["정제주문번호"] = df.iloc[:, 10].apply(lambda x: str(x).split("-")[1] if "-" in str(x) else "")

    sheet = connect_google_sheet_by_url(SPREADSHEET_URL)
    # (시트 업데이트 로직 단순화)

    print(f"✅ {len(df)}건 구글시트 업데이트 완료")

# ======================== 4. 환불 예정일 계산 & 알림톡 발송 ========================
def get_refund_date(business_days=4):
    today = datetime.today()
    while business_days > 0:
        today += timedelta(days=1)
        if today.weekday() < 5:
            business_days -= 1
    return today.strftime("%Y-%m-%d")

def send_alimtalk(phone, product, order_number, name, refund_date):
    url = "https://jupiter.lunasoft.co.kr/api/AlimTalk/message/send"
    header = {"Content-type": "application/json"}
    body = {
        "userid": "verish",
        "api_key": LUNASOFT_API_KEY,
        "template_id": "50132",
        "messages": [{
            "tel_num": phone,
            "msg_content": f"- 안녕하세요 고객님 베리시입니다 :) \n 회수된 상품이 금일 물류사 도착하였습니다.\n 환불은 회수 완료일 제외 영업일 기준 최대 4일 검수 후 처리 될 예정이니 참고 부탁드립니다.\n 주문번호: {order_number}\n- 상품명: {product}\n- 예상 환불 날짜: {refund_date}",
        }]
    }
    res = requests.post(url, headers=header, json=body)
    print("✅ 알림톡 발송 성공" if res.status_code == 200 else f"❌ 실패: {res.text}")

def send_notice_to_completed_returns():
    sheet = connect_google_sheet_by_url(SPREADSHEET_URL)
    data = sheet.get_all_values()
    header = data[0]
    df = pd.DataFrame(data[1:], columns=header)

    for _, row in df.iterrows():
        if row.get("배송상태") == "배송완료" and not row.get("알림톡 발신 여부"):
            refund_date = get_refund_date()
            send_alimtalk(row["연락처"], row["상품명"], row["주문번호"], row["성함"], refund_date)
            print(f"📦 환불 알림 발송 완료 → {row['주문번호']}")

# ======================== 실행 ========================
if __name__ == "__main__":
    update_google_sheet_from_excel()
    send_notice_to_completed_returns()
