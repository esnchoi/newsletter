import os
import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from datetime import datetime

# === OAuth 스코프 (원본 유지) ===
SCOPES = [
    'https://www.googleapis.com/auth/analytics.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# === 여기에 지난번에 쓰던 ga_token.json "내용 전체(JSON 원문)"를 붙여 넣으세요 ===
# 예시 형태: {"token":"ya29....","refresh_token":"1//0...","client_id":"...","client_secret":"...","scopes":["..."],"expiry":"2025-11-10T09:33:00Z"}
TOKEN_JSON_STR = r'''
{ "PUT_YOUR_OLD_TOKEN_JSON_HERE": true }
'''.strip()

PROPERTY_ID = "464149233"
SEARCH_TERMS_SHEET_ID = "1vxP7tVII0oWaGtro8puSXy7lDvYDrnppaRPv2qFACm0"


def get_credentials():
    """
    파일 없이 동작: 코드에 하드코딩한 token.json 내용으로 인증.
    refresh_token이 있으면 자동 갱신됨.
    """
    if not TOKEN_JSON_STR or TOKEN_JSON_STR == '{ "PUT_YOUR_OLD_TOKEN_JSON_HERE": true }':
        raise SystemExit("TOKEN_JSON_STR에 이전 ga_token.json의 JSON 원문을 그대로 붙여 넣으세요.")

    info = json.loads(TOKEN_JSON_STR)
    creds = Credentials.from_authorized_user_info(info, SCOPES)

    # 만료 시 자동 갱신
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds


def get_search_terms_from_sheet():
    creds = get_credentials()
    sheets_service = build('sheets', 'v4', credentials=creds)

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SEARCH_TERMS_SHEET_ID,
        range="B:B"
    ).execute()

    values = result.get('values', [])
    search_terms = []

    for i, row in enumerate(values):
        if i == 0:
            continue
        if row and row[0].strip():
            search_terms.append(row[0].strip())

    print(f"구글 시트에서 {len(search_terms)}개의 검색어를 가져왔습니다.")
    return search_terms


def get_analytics_data_for_search_term(search_term, start_date, end_date):
    creds = get_credentials()
    analytics = build('analyticsdata', 'v1beta', credentials=creds)

    request_body = {
        "dateRanges": [{"startDate": start_date, "endDate": end_date}],
        "metrics": [{"name": "eventCount"}],
        "dimensions": [{"name": "sessionSource"}, {"name": "eventName"}],
        "dimensionFilter": {
            "andGroup": {
                "expressions": [
                    {
                        "filter": {
                            "fieldName": "eventName",
                            "stringFilter": {"matchType": "EXACT", "value": "click"}
                        }
                    },
                    {
                        "filter": {
                            "fieldName": "sessionSource",
                            "stringFilter": {"matchType": "EXACT", "value": search_term}
                        }
                    }
                ]
            }
        },
        "limit": 1000
    }

    try:
        response = analytics.properties().runReport(
            property=f"properties/{PROPERTY_ID}",
            body=request_body
        ).execute()
        return response
    except Exception as e:
        print(f"검색어 '{search_term}'에 대한 데이터 조회 중 오류: {e}")
        return None


def find_today_column(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SEARCH_TERMS_SHEET_ID,
        range="1:1"
    ).execute()

    values = result.get('values', [[]])
    if not values or not values[0]:
        print("헤더 행을 찾을 수 없습니다.")
        return None

    today = datetime.now().strftime("%Y-%m-%d")
    today_short = datetime.now().strftime("%m/%d")
    today_dot = datetime.now().strftime("%Y.%m.%d")
    date_formats = [today, today_short, today_dot]

    for i, cell in enumerate(values[0]):
        if cell:
            cell_str = str(cell).strip()
            for date_format in date_formats:
                if date_format in cell_str:
                    col_letter = chr(65 + i) if i < 26 else chr(64 + i // 26) + chr(65 + i % 26)
                    print(f"오늘 날짜 열 발견: {col_letter}1 ({cell_str})")
                    return col_letter

    print(f"오늘 날짜({today})에 해당하는 열을 찾을 수 없습니다.")
    return None


def update_single_cell(sheets_service, search_term, click_count, today_column, row_number):
    try:
        cell_range = f"{today_column}{row_number}"
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SEARCH_TERMS_SHEET_ID,
            range=cell_range,
            valueInputOption='USER_ENTERED',
            body={'values': [[click_count]]}
        ).execute()
        print(f"✓ '{search_term}': {click_count} → {cell_range}")
        return True
    except Exception as e:
        print(f"✗ '{search_term}' 기록 실패: {e}")
        return False


def main():
    # 2025년 2월 1일부터 오늘까지 누적 데이터 수집
    start_date = "2025-02-01"
    end_date = datetime.now().strftime("%Y-%m-%d")

    print(f"Google Analytics 데이터 수집: {start_date} ~ {end_date}")

    creds = get_credentials()
    sheets_service = build('sheets', 'v4', credentials=creds)

    today_column = find_today_column(sheets_service)
    if not today_column:
        print("오늘 날짜 열을 찾을 수 없어 업데이트를 중단합니다.")
        return

    search_terms = get_search_terms_from_sheet()
    if not search_terms:
        print("검색어를 가져올 수 없습니다.")
        return

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SEARCH_TERMS_SHEET_ID,
        range="B:B"
    ).execute()
    values = result.get('values', [])

    additional_values = {
        'sellerocean': 6,
        'sba': 22,
        'd2c': 1,
        'etc': 2,
        'closet': 11,
        'salecafe': 6
    }

    print("\n=== 검색어별 클릭 이벤트 수 ===")
    print("검색어\t\t클릭 이벤트 수")
    print("-" * 40)

    success_count = 0
    fail_count = 0

    for search_term in search_terms:
        print(f"'{search_term}' 조회 중...")
        response = get_analytics_data_for_search_term(search_term, start_date, end_date)

        total_clicks = 0
        if response and 'rows' in response:
            for row in response['rows']:
                clicks = int(row['metricValues'][0]['value'])
                total_clicks += clicks

        if search_term in additional_values:
            total_clicks += additional_values[search_term]

        print(f"{search_term}\t\t{total_clicks}")

        actual_row_number = None
        for row_idx, row in enumerate(values):
            if row and row[0].strip() == search_term:
                actual_row_number = row_idx + 1
                break

        if actual_row_number:
            if update_single_cell(sheets_service, search_term, total_clicks, today_column, actual_row_number):
                success_count += 1
            else:
                fail_count += 1

    print(f"\n모든 데이터 기록 완료! (성공: {success_count}, 실패: {fail_count})")
    if fail_count > 0:
        exit(1)


if __name__ == "__main__":
    main()
