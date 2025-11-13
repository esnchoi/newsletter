import os, json, mimetypes, sys
import requests
from requests.auth import HTTPBasicAuth
from googleapiclient.discovery import build
from google.oauth2 import service_account

# ===== 환경변수 =====
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
RANGE_NAME     = os.environ.get("RANGE_NAME", "A1:G")
WP_API_BASE    = os.environ["WP_API_BASE"].rstrip("/")  # e.g. https://newsletterdev.hanpda.com/wp-json/wp/v2
WP_USER        = os.environ["WP_USER"]
WP_APP_PASSWORD= os.environ["WP_APP_PASSWORD"]
# 서비스 계정 JSON(Base64 또는 원문 JSON 중 하나)
GOOGLE_SA_JSON = os.environ["GOOGLE_SA_JSON"]

TIMEOUT = 20

def build_sheets_service():
    # GOOGLE_SA_JSON이 Base64 인코딩돼 있든, JSON 원문이든 처리
    try:
        if GOOGLE_SA_JSON.strip().startswith("{"):
            info = json.loads(GOOGLE_SA_JSON)
        else:
            import base64
            info = json.loads(base64.b64decode(GOOGLE_SA_JSON))
    except Exception as e:
        print(f"[ERR] GOOGLE_SA_JSON 파싱 실패: {e}", file=sys.stderr)
        raise

    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)

def read_sheet_data(svc, rng):
    sheet = svc.spreadsheets()
    res = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=rng).execute()
    return res.get("values", [])

def update_sheet_cell(svc, row, col_letter, value):
    rng = f"{col_letter}{row}"
    body = {"values": [[value]]}
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    print(f"[OK] 시트 업데이트: {rng} -> {value}")

def upload_image_to_wordpress(image_url):
    try:
        r = requests.get(image_url, timeout=TIMEOUT)
        r.raise_for_status()
        image_data = r.content

        # 파일명/콘텐츠타입 추정
        name = image_url.split("?")[0].split("/")[-1] or "image.jpg"
        ctype = r.headers.get("Content-Type")
        if not ctype:
            ctype, _ = mimetypes.guess_type(name)
        if not ctype:
            ctype = "application/octet-stream"

        headers = {
            "Content-Disposition": f'attachment; filename="{name}"',
            "Content-Type": ctype
        }

        media_url = f"{WP_API_BASE}/media"
        mr = requests.post(
            media_url,
            auth=HTTPBasicAuth(WP_USER, WP_APP_PASSWORD),
            headers=headers,
            data=image_data,
            timeout=TIMEOUT,
        )
        if mr.status_code == 201:
            media_id = mr.json().get("id")
            print(f"[OK] 이미지 업로드: id={media_id}")
            return media_id
        else:
            print(f"[ERR] 이미지 업로드 실패 {mr.status_code} {mr.text}", file=sys.stderr)
            return None
    except Exception as e:
        print(f"[ERR] 이미지 처리 오류: {e}", file=sys.stderr)
        return None

def spacer_block(px=50):
    return f'<!-- wp:spacer {{"height":"{px}px"}} -->\n<div style="height:{px}px" aria-hidden="true" class="wp-block-spacer"></div>\n<!-- /wp:spacer -->'

def post_to_wordpress(title, content, subtitle, media_id=None):
    try:
        subtitle_block = f'{spacer_block()}\n<!-- wp:heading {{"level":3}} -->\n<h3>{subtitle}</h3>\n<!-- /wp:heading -->' if subtitle else ""
        content_blocks = f'<!-- wp:freeform -->\n{content}\n<!-- /wp:freeform -->'
        full_content = f'{subtitle_block}\n{content_blocks}\n{spacer_block()}'

        data = {
            "title": title,
            "content": full_content,
            "status": "draft",
        }
        if media_id:
            data["featured_media"] = media_id

        posts_url = f"{WP_API_BASE}/posts"
        resp = requests.post(
            posts_url,
            auth=HTTPBasicAuth(WP_USER, WP_APP_PASSWORD),
            json=data,
            timeout=TIMEOUT,
        )
        if resp.status_code == 201:
            pid = resp.json().get("id")
            print(f"[OK] 임시글 업로드: post_id={pid}")
            return True
        print(f"[ERR] WP 업로드 실패 {resp.status_code} {resp.text}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[ERR] WP 게시 오류: {e}", file=sys.stderr)
        return False

def main():
    print("[RUN] 시작")
    svc = build_sheets_service()
    values = read_sheet_data(svc, RANGE_NAME)
    if not values:
        print("[WARN] 시트 데이터 없음")
        return

    for idx, row in enumerate(values[1:], start=2):
        try:
            if len(row) < 4:
                print(f"[SKIP] {idx}행: 필수 필드 부족")
                continue
            if len(row) >= 5 and (row[4] or "").strip() == "완료":
                print(f"[SKIP] {idx}행: 이미 완료")
                continue

            release_date = row[0] if len(row) > 0 else ""
            title        = row[1] if len(row) > 1 else ""
            subtitle     = row[2] if len(row) > 2 else ""
            content      = row[3] if len(row) > 3 else ""
            image_url    = row[6] if len(row) > 6 else ""

            print(f"[PROC] {idx}행: {title}")

            media_id = upload_image_to_wordpress(image_url) if image_url else None
            ok = post_to_wordpress(title, content, subtitle, media_id)
            if ok:
                update_sheet_cell(svc, idx, "E", "완료")
            else:
                print(f"[FAIL] {idx}행: 게시 실패")
        except Exception as e:
            print(f"[ERR] {idx}행 처리 오류: {e}", file=sys.stderr)

    print("[DONE] 종료")

if __name__ == "__main__":
    main()
