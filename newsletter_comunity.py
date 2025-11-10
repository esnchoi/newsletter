import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Google Sheets API 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ✅ GitHub / 리눅스 환경에서도 인식되도록 절대경로 대신 상대경로로 수정
CLIENT_SECRET_FILE = './client_secret.json'
TOKEN_FILE = './token.json'

# 스프레드시트 ID와 범위 설정
SPREADSHEET_ID = '1M5kUEQJtwGtrCqQbJHtCGvd5CFZagms07SgiZ4YZ3ZI'
RANGE_NAME = 'A:G'  # A~G열 전체

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # GitHub Actions에서는 브라우저 인증 불가 → 반드시 로컬에서 token.json 미리 만들어 올려야 함
            if os.getenv("GITHUB_ACTIONS") == "true":
                raise RuntimeError("token.json 파일이 필요합니다. 로컬에서 한 번 로그인해 만든 후 리포에 올리세요.")
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def read_sheet_data(sheets_service, range_name):
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])
    # A열(번호)만 반환
    return set(row[0] for row in values if row)

def write_sheet_data(sheets_service, range_name, values):
    body = {'values': values}
    result = sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()
    print(f"{result.get('updates').get('updatedRows')} rows appended.")

def main():
    creds = get_credentials()
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Selenium 설정
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    # WebDriver 초기화
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # 사이트 접속
        driver.get('https://newsletter.cafe24.com/community/')
        time.sleep(3)

        # 스크롤 다운으로 모든 게시글 로드
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # BeautifulSoup 파싱
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rows = soup.select('#kboard-default-list > div.kboard-list > table > tbody > tr')

        data = []
        base_url = 'https://newsletter.cafe24.com/community/'

        for row in rows:
            number = row.select_one('td.kboard-list-uid').get_text(strip=True)
            if number == '공지사항':
                continue
            title = row.select_one('td.kboard-list-title a div').get_text(strip=True)
            author = row.select_one('td.kboard-list-user').get_text(strip=True)
            date = row.select_one('td.kboard-list-date').get_text(strip=True)
            url = base_url + row.select_one('td.kboard-list-title a')['href']

            # 게시글 본문 및 댓글 추출
            driver.get(url)
            time.sleep(2)
            post_soup = BeautifulSoup(driver.page_source, 'html.parser')
            content = post_soup.select_one(
                '#kboard-default-document > div.kboard-document-wrap > div.kboard-content > div'
