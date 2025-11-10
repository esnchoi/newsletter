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
CLIENT_SECRET_FILE = r'C:\Users\7040_64bit\Documents\코드 테스트\사내뉴스레터 커뮤니티\client_secret.json'
TOKEN_FILE = 'token.json'

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

    # Selenium 설정 (정상 코드와 동일한 방식 적용)
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    # WebDriver 초기화 (불필요한 service 제거, 정상 코드와 동일하게 설정)
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # 사이트 접속
        driver.get('https://newsletter.cafe24.com/community/')
        time.sleep(3)  # 페이지 로딩 대기

        # 스크롤을 아래로 이동하며 데이터 수집
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # 페이지 소스 가져오기
        page_source = driver.page_source

        # BeautifulSoup으로 파싱
        soup = BeautifulSoup(page_source, 'html.parser')

        # 게시글 추출 (공지사항 제외)
        rows = soup.select('#kboard-default-list > div.kboard-list > table > tbody > tr')
        data = []
        base_url = 'https://newsletter.cafe24.com/community/'  # base URL 설정
        for row in rows:
            number = row.select_one('td.kboard-list-uid').get_text(strip=True)
            if number == '공지사항':  # 번호 항목에 "공지사항"이 있으면 제외
                continue
            title = row.select_one('td.kboard-list-title a div').get_text(strip=True)
            author = row.select_one('td.kboard-list-user').get_text(strip=True)
            date = row.select_one('td.kboard-list-date').get_text(strip=True)
            url = base_url + row.select_one('td.kboard-list-title a')['href']  # URL 추출

            # 게시글 본문 및 댓글 추출
            driver.get(url)
            time.sleep(3)  # 페이지 로딩 대기
            post_soup = BeautifulSoup(driver.page_source, 'html.parser')
            content = post_soup.select_one('#kboard-default-document > div.kboard-document-wrap > div.kboard-content > div').get_text(strip=True)
            comments = post_soup.select('div.comments-list-content')
            comments_text = ' '.join([comment.get_text(strip=True) for comment in comments])

            data.append([number, title, author, date, url, content, comments_text])

        # 추출된 데이터 콘솔에 출력
        for row in data:
            print(row)

        if data:
            # 기존 데이터의 번호 가져오기
            existing_numbers = read_sheet_data(sheets_service, 'A:A')

            # 중복 체크 및 필터링
            new_data = [row for row in data if row[0] not in existing_numbers]

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
