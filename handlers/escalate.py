"""
真人客服升級（Escalation）模組

當 Bot 無法處理客戶訊息時：
1. 寫入 issue_store（type='unknown'），納入每小時待處理清單
2. 回覆客戶「已轉真人客服」
3. 記錄 log
"""

from linebot.v3.messaging import MessagingApi

from handlers import tone
from storage.customers import customer_store
from storage.issues import issue_store


def handle_unknown(user_id: str, text: str, line_api: MessagingApi) -> str:
    """
    處理 Bot 無法識別的訊息：
    - 寫入待處理清單（每小時彙整推送）
    - 回覆客戶等待訊息
    """
    # 取得客戶顯示名稱
    cust_info = customer_store.get_by_line_id(user_id)
    display_name = cust_info.get("display_name", "") if cust_info else ""
    cust_label = display_name or user_id

    # 寫入待處理清單（每小時推送給真人看）
    issue_store.add(user_id, "unknown", text)

    try:
        print(f"[escalate] 未知訊息 | {cust_label}: {text!r}")
    except UnicodeEncodeError:
        _safe = text.encode("utf-8", errors="replace").decode("utf-8", errors="replace")
        print(f"[escalate] 未知訊息 | {cust_label}: {_safe!r}")

    # 休息時間 → 回覆下次上班時間
    from handlers.hours import _is_open_now, next_open_reply
    from datetime import datetime
    import pytz
    from config import settings
    now = datetime.now(pytz.timezone(settings.BUSINESS_TZ))
    if not _is_open_now(now):
        return next_open_reply()

    return ""  # 上班時間靜默記錄，等真人處理
