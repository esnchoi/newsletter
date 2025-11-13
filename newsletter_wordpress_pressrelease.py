import requests
from requests.auth import HTTPBasicAuth
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os

# Google Sheets API 설정 (읽기 및 쓰기 권한)
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRET_FILE = r'C:\Users\7040_64bit\Documents\코드 테스트\인사 정보 관리\client_secret.json'
TOKEN_FILE = 'token.json'

# 스프레드시트 ID와 범위 설정 (URL 변경)
SPREADSHEET_ID = '17P2B1JPjiLIX59aSdrrgrhxHKaOboZeAExqnJGXYvHQ'
RANGE_NAME = 'A1:G'  # G열에 이미지 URL 추가

# WordPress 정보
wp_url = "https://newsletterdev.hanpda.com/wp-json/wp/v2/posts"
wp_media_url = "https://newsletterdev.hanpda.com/wp-json/wp/v2/media"
wp_username = "marketing"
wp_password = "TkJW 2IT7 b9zX gxzx qSQN frpE"  # 응용 프로그램 비밀번호

def get_credentials():
    print("Google 인증 시작")
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        print("기존 토큰 파일에서 인증 완료")
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("토큰 갱신 완료")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            print("새 인증 완료")
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def read_sheet_data(sheets_service, range_name):
    print("Google Sheets 데이터 읽기 시작")
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    print(f"읽은 데이터: {result.get('values', [])}")
    return result.get('values', [])

def update_sheet_cell(sheets_service, row, col, value):
    range_to_update = f'{col}{row}'
    body = {'values': [[value]]}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_to_update,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    print(f"구글 시트 {range_to_update} 셀 업데이트 완료: '{value}'")

def upload_image_to_wordpress(image_url):
    try:
        # 이미지 다운로드
        image_response = requests.get(image_url)
        image_data = image_response.content
        image_name = image_url.split("/")[-1]
        
        headers = {
            'Content-Disposition': f'attachment; filename={image_name}',
            'Content-Type': 'image/jpeg'  # 이미지 유형에 따라 수정 필요
        }

        # 이미지 워드프레스에 업로드
        media_response = requests.post(wp_media_url, auth=HTTPBasicAuth(wp_username, wp_password), headers=headers, data=image_data)
        if media_response.status_code == 201:
            media_id = media_response.json()['id']
            print(f"이미지 업로드 성공: {media_id}")
            return media_id
        else:
            print(f"이미지 업로드 실패: {media_response.status_code}, {media_response.text}")
            return None
    except Exception as e:
        print(f"이미지 다운로드 또는 업로드 오류: {e}")
        return None

def add_spacer():
    return '<!-- wp:spacer {"height":"50px"} -->\n<div style="height:50px" aria-hidden="true" class="wp-block-spacer"></div>\n<!-- /wp:spacer -->'

def post_to_wordpress(title, content, subtitle, media_id=None):
    try:
        print(f"WordPress에 임시글 작성 시도: 제목={title}")
        
        # 부제목은 h3 블록으로만 처리
        subtitle_block = f'{add_spacer()}\n<!-- wp:heading {{"level":3}} -->\n<h3>{subtitle}</h3>\n<!-- /wp:heading -->'
        
        # 본문 내용을 클래식 블록으로 변환
        content_blocks = f'<!-- wp:freeform -->\n{content}\n<!-- /wp:freeform -->'
        
        # 마지막에 공백 블록 추가
        full_content = f'{subtitle_block}\n{content_blocks}\n{add_spacer()}'

        data = {
            "title": title,  # 제목만 단독으로 사용
            "content": full_content,
            "status": "draft"
        }
        
        if media_id:
            data["featured_media"] = media_id  # 특성 이미지 추가
        
        # WordPress API로 게시물 업로드
        response = requests.post(wp_url, auth=HTTPBasicAuth(wp_username, wp_password), json=data)
        
        if response.status_code == 201:
            print("WordPress에 임시글이 성공적으로 업로드되었습니다.")
            return True
        else:
            print(f"WordPress 업로드 실패: {response.status_code}, {response.text}")
            return False
            
    except Exception as e:
        print(f"WordPress 게시 중 오류 발생: {e}")
        return False

def main():
    try:
        print("프로그램 시작")
        creds = get_credentials()
        print("Google Sheets 인증 완료")
        
        sheets_service = build('sheets', 'v4', credentials=creds)
        print("Google Sheets API 서비스 빌드 완료")

        values = read_sheet_data(sheets_service, RANGE_NAME)
        
        if not values:
            print('데이터가 없습니다.')
            return

        # 워드프레스에 글 작성 로직
        for index, row in enumerate(values[1:], start=2):
            print(f"행 {index} 처리 시작")
            
            # 필수 필드 확인
            if len(row) < 4:
                print(f"행 {index}: 필수 필드 누락으로 건너뜁니다")
                continue

            # 이미 완료된 항목 확인
            if len(row) >= 5 and row[4].strip() == '완료':
                print(f"행 {index}: 이미 완료된 항목이므로 건너뜁니다.")
                continue

            # 데이터 추출
            release_date = row[0]
            title = row[1]
            subtitle = row[2] if len(row) > 2 else ''
            content = row[3]
            image_url = row[6] if len(row) > 6 else None

            print(f"처리중: {title}")

            # 이미지 처리
            media_id = None
            if image_url:
                media_id = upload_image_to_wordpress(image_url)

            # WordPress에 글 게시
            if post_to_wordpress(title, content, subtitle, media_id):
                update_sheet_cell(sheets_service, index, 'E', '완료')
            else:
                print(f"행 {index}: WordPress 게시 실패")

        print("프로그램 정상 종료")
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")

if __name__ == '__main__':
    main()