"""
Google Calendar OAuth 授權腳本

第一次使用前執行：
    python setup_google_cal.py

會開啟瀏覽器讓你登入 Google 並授權。
授權成功後產生 token.json，之後 Bot 自動使用。
"""

from services.google_cal import calendar_client

if __name__ == "__main__":
    print("開始 Google Calendar 授權流程...")
    service = calendar_client._get_service()
    if service:
        print("✅ 授權成功！token.json 已儲存。")
        print("請確認 .env 中的 GOOGLE_CALENDAR_ID 已填入正確的日曆 ID。")
    else:
        print("❌ 授權失敗。")
        print("請確認 credentials.json 已下載並放在專案根目錄。")
        print("下載方式：Google Cloud Console → API & Services → Credentials → 下載 OAuth 2.0 Client ID")
