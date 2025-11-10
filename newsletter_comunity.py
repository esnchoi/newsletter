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
CLIENT_SECRET_FILE = './client_secret.json'  # 리눅스/Actions 호환
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
    print(f"{result.get('updates', {}).get('updatedRows', 0)} rows appended.")

def main():
    creds = get_credentials()
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Selenium 설정
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

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
            num_el = row.select_one('td.kboard-list-uid')
            title_el = row.select_one('td.kboard-list-title a div')
            link_el = row.select_one('td.kboard-list-title a')
            author_el = row.select_one('td.kboard-list-user')
            date_el = row.select_one('td.kboard-list-date')

            if not (num_el and title_el and link_el):
                continue

            number = num_el.get_text(strip=True)
            if number == '공지사항':
                continue

            title = title_el.get_text(strip=True)
            author = author_el.get_text(strip=True) if author_el else ''
            date = date_el.get_text(strip=True) if date_el else ''
            url = base_url + link_el.get('href', '')

            # 게시글 본문 및 댓글 추출
            driver.get(url)
            time.sleep(2)
            post_soup = BeautifulSoup(driver.page_source, 'html.parser')
            content_el = post_soup.select_one(
                '#kboard-default-document > div.kboard-document-wrap > div.kboard-content > div'
            )
            content = content_el.get_text(strip=True) if content_el else ''
            comments = post_soup.select('div.comments-list-content')
            comments_text = ' '.join(c.get_text(strip=True) for c in comments) if comments else ''

            data.append([number, title, author, date, url, content, comments_text])

        # 콘솔 출력
        for r in data:
            print(r)

        if data:
            existing_numbers = read_sheet_data(sheets_service, 'A:A')
            new_data = [r for r in data if r[0] not in existing_numbers]

            if new_data:
                write_sheet_data(sheets_service, RANGE_NAME, new_data)
                print(f"{len(new_data)}개의 새로운 데이터가 추가되었습니다.")
            else:
                print("중복된 데이터가 있어 새로 추가된 데이터가 없습니다.")
        else:
            print("기록할 데이터가 없습니다.")
    finally:
        driver.quit()

if __name__ == '__main__':
    main()
