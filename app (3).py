import imaplib
import email
from email import policy
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
from dotenv import load_dotenv
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta



print('실행시간',datetime.now())
# 🔐 환경변수 로드
load_dotenv()
CNPLUS_ID = os.getenv("CNPLUS_ID")
CNPLUS_PW = os.getenv("CNPLUS_PW")

# 이메일 헤더에서 제목 가져오기 함수
def find_encoding_info(txt):
    info = email.header.decode_header(txt)
    subject, encode = info[0]
    return subject, encode

# 🌐 Selenium 웹드라이버 실행 (크롬)
options = webdriver.ChromeOptions()
#options.add_argument('--headless')  # 창 안띄우고 실행
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')  # GPU 비활성화 (mac에서도 안정성↑)
options.add_argument('--window-size=1920x1080')

options.add_argument('--start-maximized')  # 창 최대화
driver = webdriver.Chrome(options=options)

# 1. CNPLUS 로그인 페이지 접속
driver.get("https://loisparcelp.cjlogistics.com/index.do")
time.sleep(3)

# 📌 1. 아이디 입력 필드 찾기
id_input = driver.find_element(By.CSS_SELECTOR, "input.cl-text[data-ndid='4e'][type='text']")

# 📌 2. 아이디 입력
id_input.clear()
id_input.send_keys(CNPLUS_ID)

# 3️⃣ 비밀번호 입력 필드 찾기
pw_input = driver.find_element(By.CSS_SELECTOR, "input.cl-text[data-ndid='4h'][type='password']")
pw_input.clear()
pw_input.send_keys(CNPLUS_PW)
print("✅ 비밀번호 입력 완료")

# 6️⃣ 로그인 버튼 클릭
login_button = driver.find_element(By.ID, "uuid-3r")
login_button.click()
print("🚪 로그인 버튼 클릭 완료")

time.sleep(3)

# 로그인 후 인증번호 받기
# 네이버 imap 접속주소로 접속하기
imap = imaplib.IMAP4_SSL('imap.naver.com')
id = '2kh616@naver.com'  # 네이버 이메일 주소
pw = 'NMPRHKH6S3TC'  # 네이버 이메일 비밀번호
imap.login(id, pw)
time.sleep(1.5)

# 최근 1개의 이메일 가져오기
imap.select('INBOX')
resp, data = imap.uid('search', None, 'ALL')
all_email = data[0].split()
latest_email = all_email[-1]
time.sleep(1.5)

# 1개의 이메일에서 정보 가져오기
result, data = imap.uid('fetch', latest_email, '(RFC822)')
raw_email = data[0][1]
email_message = email.message_from_bytes(raw_email, policy=policy.default)
time.sleep(1.5)

# 본문 내용 출력하기
message = ''
if email_message.is_multipart():
    for part in email_message.get_payload():
        if part.get_content_type() == 'text/html':  # HTML 형식의 본문을 찾습니다.
            bytes = part.get_payload(decode=True)
            encode = part.get_content_charset()
            message = str(bytes, encode)
else:
    bytes = email_message.get_payload(decode=True)
    encode = email_message.get_content_charset()
    message = str(bytes, encode)

# BeautifulSoup을 사용하여 HTML 파싱
soup = BeautifulSoup(message, 'html.parser')

# 인증번호 추출: <font size="7">036523</font> 형태에서 숫자만 추출
auth_code = soup.find_all('font', {'size': '7'})
auth_code_list = [code.get_text() for code in auth_code if code.get_text().isdigit()]

# 인증번호 출력
if auth_code_list:
    auth_code = auth_code_list[0]  # 첫 번째 인증번호
    print(f"인증번호: {auth_code}")  
else:
    print("인증번호를 찾을 수 없습니다.")

imap.close()
imap.logout()

# 인증번호를 입력 필드에 자동으로 입력하기 위해서
# 📌 인증번호 입력 필드 찾기
input_field = driver.find_element(By.CSS_SELECTOR, 'input.cl-text')  # <input class="cl-text"> 찾기
input_field.send_keys(auth_code)  # 인증번호 입력

# 📌 로그인 버튼 클릭
login_button = driver.find_element(By.CSS_SELECTOR, '.btn-login a')  # 로그인 버튼의 a 태그 찾기
login_button.click()

# 로그인 후 일정 시간 대기
time.sleep(3)  # 페이지 로딩을 기다립니다.



try:
    while True:
        close_buttons = driver.find_elements(By.XPATH, '//div[contains(@class, "cl-text") and text()="닫기"]')
        
        # 닫기 버튼이 있으면 클릭
        if close_buttons:
            for btn in close_buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", btn)
                    print("✅ 닫기 버튼 클릭 완료")
                    break  # 하나 클릭하고 다시 체크로 돌아감
            time.sleep(0.5)  # 팝업 사라지는 시간 기다리기
        else:
            print("✅ 더 이상 닫기 버튼이 없습니다.")
            break
except Exception as e:
    print(f"❌ 닫기 버튼 처리 중 오류 발생: {e}")


try:
    # 1️⃣ 나의메뉴 클릭
    my_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@title="나의메뉴"]'))
    )
    
    # 이미 펼쳐져 있는지 확인 (aria-expanded="false"일 경우에만 클릭)
    if my_menu.get_attribute("aria-expanded") == "false":
        driver.execute_script("arguments[0].scrollIntoView(true);", my_menu)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", my_menu)
        print("✅ '나의메뉴' 클릭 완료 (펼쳐짐 처리)")

    # 2️⃣ 기업고객반품상세 클릭
    sub_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@title="기업고객반품상세"]'))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", sub_menu)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", sub_menu)
    print("✅ '기업고객반품상세' 클릭 완료")

except Exception as e:
    print(f"❌ 메뉴 클릭 흐름 중 오류 발생: {e}")

time.sleep(2)

try:
    # 1️⃣ 달력 버튼 클릭
    calendar_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@aria-label="달력으로 날짜 선택"]'))
    )
    calendar_button.click()
    print("✅ 달력 버튼 클릭 완료")
    time.sleep(1)  # 달력이 열릴 시간 확보

    # 2️⃣ 날짜 input 필드 다시 포커스
    date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//input[@aria-label="기준일자"]'))
    )
    date_input.click()
    time.sleep(0.2)

    # 3️⃣ 왼쪽 방향키 7번 → 날짜 이동
    for _ in range(8):
        date_input.send_keys(Keys.ARROW_LEFT)
        time.sleep(0.4)

    # 4️⃣ Enter로 선택 완료
    date_input.send_keys(Keys.ENTER)
    print("✅ 7일 전 날짜 선택 완료 (← 방향키 방식)")

except Exception as e:
    print(f"❌ 달력 선택 실패, 보스: {e}")
    

try:
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.btn-i-search a div.cl-text'))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
    time.sleep(0.2)
    search_button.click()
    print("✅ 조회 버튼 클릭 성공, 강히야 축하하고 이제 다음꺼 해보자")

except Exception as e:
    print(f"❌ 조회 버튼 클릭 실패, 강히야 다시 햅ㅗ자 하: {e}")

time.sleep(2)

try:
    excel_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.btn-i-excel a div.cl-text'))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", excel_button)
    time.sleep(0.2)
    excel_button.click()
    print("✅ 엑셀 다운로드 버튼 클릭 성공")

except Exception as e:
    print(f"❌ 엑셀 클릭 실패 : {e}")



time.sleep(20)




import imaplib
import email
from email import policy
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
from dotenv import load_dotenv
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
######################################아래는 데이터 구글드라이브 업로드 ##############################################
import os
import glob
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# ✅ 구글시트 URL
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1FX9Lh1fEDNszkItL-z4uj6rX7QVlDjwvP9evvjXRxSM/edit#gid=193342666"

# ✅ 최신 엑셀 파일 경로 가져오기
def get_latest_excel_file(download_path="/Users/deepdive/Downloads"):
    excel_files = glob.glob(os.path.join(download_path, "*.xlsx"))
    if not excel_files:
        raise FileNotFoundError("❌ 다운로드 폴더에 엑셀 파일이 없소, 보스.")
    return max(excel_files, key=os.path.getctime)

# ✅ 주문번호 정제 함수
def extract_order_code(raw):
    try:
        parts = raw.split('-')
        if len(parts) >= 3:
            date = parts[1]
            code = parts[2].split('_')[0]
            return f"{date.strip()}-{code.strip()}"
    except:
        return None

# ✅ NaN, 공백 등 정리
def clean_value(val):
    if pd.isna(val) or val in [None, float('inf'), float('-inf'), 'nan']:
        return ''
    return str(val).strip()

# ✅ 구글시트 연결
def connect_google_sheet_by_url(spreadsheet_url, worksheet_index=0):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "/Users/deepdive/Documents/강희/Github/customer-alimtalk/예상반품환불일 안내/new_google_API_KEY/mystical-button-438607-h5-ecc3ead3147d.json",
        scope
    )
    client = gspread.authorize(creds)
    return client.open_by_url(spreadsheet_url).get_worksheet(worksheet_index)

# ✅ 통합 실행 함수
def update_google_sheet_from_excel():
    latest_file = get_latest_excel_file()
    df = pd.read_excel(latest_file, skiprows=2, usecols=range(1, 16), engine="openpyxl")
    df['정제주문번호'] = df.iloc[:, 10].apply(lambda x: extract_order_code(str(x)))

    sheet = connect_google_sheet_by_url(spreadsheet_url)
    data = sheet.get_all_values()
    header = data[0]
    gsheet_df = pd.DataFrame(data[1:], columns=header)

    updates_dict = {}

    for idx, row in df.iterrows():
        주문번호 = row['정제주문번호']
        if not 주문번호:
            continue

        match_index = gsheet_df[gsheet_df['주문번호'].apply(lambda x: clean_value(x)) == 주문번호].index
        if not match_index.empty:
            i = match_index[0] + 2  # 구글시트는 1-index + 헤더 고려

            반송장 = clean_value(row.iloc[2])  # D열
            원송장 = clean_value(row.iloc[3])  # E열
            배송상태 = clean_value(row.iloc[4])  # F열

            기존_배송상태 = clean_value(gsheet_df.iloc[match_index[0]].get('배송상태', ''))

            # 🚫 배송완료면 skip
            if 기존_배송상태 == "배송완료":
                print(f"🚫 {주문번호}: 시트에 이미 배송완료 → 생략")
                continue
            if 기존_배송상태 == 배송상태:
                print(f"🚫 {주문번호}: 배송상태 동일 → 생략")
                continue

            updates_dict[i] = [원송장, 반송장, 배송상태]
        else:
            print(f"⚠️ {주문번호}: 구글시트에 없음")

    # ✅ 촤자작 일괄 업데이트
    if updates_dict:
        min_row = min(updates_dict.keys())
        max_row = max(updates_dict.keys())
        update_range = f"I{min_row}:K{max_row}"

        full_update_list = []
        for r in range(min_row, max_row + 1):
            full_update_list.append(updates_dict.get(r, ['', '', '']))

        sheet.update(values=full_update_list, range_name=update_range)
        print(f"✅ 총 {len(updates_dict)}건 촤자작 업데이트 완료! 🎯")
    else:
        print("📭 업데이트할 항목 없음.")

# ✅ 실행
if __name__ == "__main__":
    update_google_sheet_from_excel()

time.sleep(30)


import os
import gspread
import requests
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from dotenv import load_dotenv

# ✅ 환경 변수 불러오기 (.env)
load_dotenv()

# ✅ 구글 시트 URL
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1FX9Lh1fEDNszkItL-z4uj6rX7QVlDjwvP9evvjXRxSM/edit#gid=193342666"

# ✅ 구글 시트 연결
def connect_google_sheet_by_url(spreadsheet_url, worksheet_index=0):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "/Users/deepdive/Documents/강희/step1/반품환불 안내 자동화/new_google_API_KEY/mystical-button-438607-h5-ecc3ead3147d.json", 
        scope
    )
    client = gspread.authorize(creds)
    return client.open_by_url(spreadsheet_url).get_worksheet(worksheet_index)

# ✅ 주말 제외한 3영업일 후 날짜 계산
def get_refund_date(business_days=4):
    today = datetime.today()
    while business_days > 0:
        today += timedelta(days=1)
        if today.weekday() < 5:  # 월~금 (0~4)만 카운트
            business_days -= 1
    return today.strftime('%Y-%m-%d')



def send_alimtalk(phone, product, order_number, name, except_date):
    url = "https://jupiter.lunasoft.co.kr/api/AlimTalk/message/send"
    header = {"Content-type": "application/json"}
    body = {
        "userid": "verish",
        "api_key": "rxnqi0z69j5te3d3duhgvxi12dfxio7jdl58a8la",
        "template_id": "50132", 
        "messages": [{
            "no": "0",
            "tel_num": phone,
            "use_sms": "1",  
            "msg_content": f"- 주문번호: {order_number}\n- 상품명: {product}\n\n- 예상 환불 날짜: {except_date}",
            "sms_content": f"- 주문번호: {order_number}\n- 상품명: {product}\n\n- 예상 환불 날짜: {except_date}",
            "title": "보내주신 상품이 회수지에 도착했습니다.",
            "template_data": {
                "ORDER_NUMBER": order_number,
                "PRODUCT": product,
                "EXCEPT_DATE": except_date
            },
            "btn_url": [{
                "url_pc": "https://www.verish.me",
                "url_mobile": "https://m.verish.me"
            }]
        }]
    }
    res = requests.post(url, headers=header, json=body)
    if res.status_code == 200:
        print(f"✅ 알림톡 발송 성공: {name}  ({phone})")
    else:
        print(f"❌ 알림톡 실패: {name} ({phone}) | {res.status_code} - {res.text}")
    






# ✅ 전체 실행 함수 (촤자작 + 개별 셀 업데이트 적용)
def send_notice_to_completed_returns():
    sheet = connect_google_sheet_by_url(spreadsheet_url)
    data = sheet.get_all_values()
    header = data[0]
    df = pd.DataFrame(data[1:], columns=header)

    updates = []  # 일괄 업데이트 리스트

    for idx, row in df.iterrows():
        배송상태 = row.get('배송상태', '').strip()
        알림톡발송여부 = row.get('알림톡 발신 여부', '').strip()
        연락처 = row.get('연락처', '').strip()
        주문번호 = row.get('주문번호', '').strip()
        상품명 = row.get('상품명', '').strip()
        성함 = row.get('성함', '').strip()

        if 배송상태 == "배송완료" and 알림톡발송여부 == '':
            except_date = get_refund_date()
            send_alimtalk(연락처, 상품명, 주문번호, 성함, except_date)
            
            # 개별 셀 업데이트 형식으로 리스트 추가
            updates.append({
                "range": f"L{idx+2}",
                "values": [[datetime.today().strftime('%Y-%m-%d')]]
            })

    # 촤자작 (batch_update 활용)
    if updates:
        sheet.batch_update(updates)
        print(f"\n📦 총 {len(updates)}건 알림톡 발송 + L열 'O' 입력 완료 🎉")
    else:
        print("\n📭 발송할 항목이 없어요!")
print('실행시간',datetime.now())
# 실행
if __name__ == "__main__":
    send_notice_to_completed_returns()
