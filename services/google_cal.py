"""
Google Calendar 客戶端

設定步驟：
1. 前往 https://console.cloud.google.com → 建立專案
2. 啟用 Google Calendar API
3. 建立 OAuth 2.0 憑證（桌面應用程式類型）
4. 下載 credentials.json 放到專案根目錄
5. 執行 python setup_google_cal.py 完成授權（會產生 token.json）
6. 在 .env 填入 GOOGLE_CALENDAR_ID

若 credentials.json 不存在，自動跳過 Calendar 功能。
"""

from datetime import datetime, timedelta
from pathlib import Path
import pytz

from config import settings

SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]


class GoogleCalendarClient:
    def __init__(self):
        self._service = None

    def get_upcoming_deliveries(self, days_ahead: int = 14) -> list[str]:
        """
        取得 Google Calendar 中未來的送貨行程。
        搜尋包含「送貨」、「配送」的事件。

        Returns: 行程摘要列表（最多回傳 5 筆）
        """
        service = self._get_service()
        if not service or not settings.GOOGLE_CALENDAR_ID:
            return []

        tz = pytz.timezone(settings.BUSINESS_TZ)
        now = datetime.now(tz)
        end = now + timedelta(days=days_ahead)

        try:
            # 先搜尋「送貨」
            events = self._query_events(service, now, end, keyword="送貨")
            # 再搜尋「配送」（避免重複以 id 去重）
            seen_ids = {e["id"] for e in events}
            for ev in self._query_events(service, now, end, keyword="配送"):
                if ev["id"] not in seen_ids:
                    events.append(ev)

            # 依時間排序，格式化回傳
            result = []
            for ev in sorted(events, key=lambda e: e["start"].get("dateTime", e["start"].get("date", "")))[:5]:
                start_raw = ev["start"].get("dateTime", ev["start"].get("date", ""))
                date_str = start_raw[:10]  # 只取 YYYY-MM-DD
                summary = ev.get("summary", "送貨行程")
                result.append(f"{date_str} {summary}")

            return result
        except Exception as e:
            print(f"[GoogleCal] 查詢失敗: {e}")
            return []

    def _query_events(
        self,
        service,
        time_min: datetime,
        time_max: datetime,
        keyword: str,
    ) -> list[dict]:
        result = (
            service.events()
            .list(
                calendarId=settings.GOOGLE_CALENDAR_ID,
                timeMin=time_min.isoformat(),
                timeMax=time_max.isoformat(),
                singleEvents=True,
                orderBy="startTime",
                q=keyword,
            )
            .execute()
        )
        return result.get("items", [])

    def _get_service(self):
        if self._service:
            return self._service

        creds_path = Path(settings.GOOGLE_CREDENTIALS_FILE)
        if not creds_path.exists():
            return None

        try:
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from google.auth.transport.requests import Request
            from googleapiclient.discovery import build

            creds = None
            token_path = Path(settings.GOOGLE_TOKEN_FILE)

            if token_path.exists():
                creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        str(creds_path), SCOPES
                    )
                    creds = flow.run_local_server(port=0)
                token_path.write_text(creds.to_json(), encoding="utf-8")

            self._service = build("calendar", "v3", credentials=creds)
        except Exception as e:
            print(f"[GoogleCal] 初始化失敗: {e}")
            return None

        return self._service


calendar_client = GoogleCalendarClient()
