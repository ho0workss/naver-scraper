"""
notify.py — 스크래핑 결과를 Google Sheets에 자동 저장합니다.

GitHub Secrets에 등록할 항목 (2개):
    GOOGLE_CREDENTIALS  — 서비스 계정 JSON 전체 내용
    SPREADSHEET_ID      — 시트 URL에서 /d/ 뒤의 ID 문자열

사용법:
    python src/notify.py competitor_latest.xlsx
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


def _clean_rows(df: pd.DataFrame) -> list[list]:
    """DataFrame → Sheets에 올릴 2D 리스트 (NaN 제거, float 정수화)"""
    header = [df.columns.tolist()]
    body = []
    for row in df.values.tolist():
        cleaned = []
        for v in row:
            if isinstance(v, float):
                cleaned.append(int(v) if v == int(v) else round(v, 4))
            elif v is None or str(v) == "nan":
                cleaned.append("")
            else:
                cleaned.append(v)
        body.append(cleaned)
    return header + body


def _write_sheet(sh, name: str, rows: list[list]) -> None:
    """시트를 찾거나 생성한 뒤 rows로 덮어씁니다."""
    import gspread
    try:
        ws = sh.worksheet(name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=len(rows) + 10, cols=len(rows[0]) + 2)
    ws.update("A1", rows, value_input_option="USER_ENTERED")


def upload_to_sheets(df: pd.DataFrame) -> None:
    creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
    sheet_id   = os.environ.get("SPREADSHEET_ID", "")

    if not creds_json:
        print("오류: GOOGLE_CREDENTIALS 환경변수가 없습니다.", file=sys.stderr)
        sys.exit(1)
    if not sheet_id:
        print("오류: SPREADSHEET_ID 환경변수가 없습니다.", file=sys.stderr)
        sys.exit(1)

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("오류: pip install gspread google-auth 필요", file=sys.stderr)
        sys.exit(1)

    print(f"Google Sheets 업로드 중 ({len(df)}개 상품)...")

    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    rows  = _clean_rows(df)
    today = datetime.now().strftime("%Y-%m-%d")

    _write_sheet(sh, today, rows)     # 날짜별 이력 시트
    _write_sheet(sh, "최신", rows)    # 항상 최신 결과 시트

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"완료 → {url}")


def main() -> None:
    if len(sys.argv) < 2:
        print("사용법: python src/notify.py <excel_path>", file=sys.stderr)
        sys.exit(1)

    excel_path = Path(sys.argv[1])
    if not excel_path.exists():
        print(f"파일 없음: {excel_path}", file=sys.stderr)
        sys.exit(1)

    upload_to_sheets(pd.read_excel(excel_path))


if __name__ == "__main__":
    main()
