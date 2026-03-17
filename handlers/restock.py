"""
HQ 群組回覆處理模組

監聽 HQ 群組 (LINE_GROUP_ID_HQ) 的訊息，根據回覆內容：
- 有貨可調 → 建立 Ecount 訂單，通知客戶
- 需要叫貨 X 週 → 推送客戶詢問是否能等
"""

import re

from linebot.v3.messaging import MessagingApi, PushMessageRequest, TextMessage

from handlers import tone
from services.ecount import ecount_client
from storage.customers import customer_store
from storage.restock import restock_store
from storage.state import state_manager
from config import settings


def handle_hq_reply(text: str, line_api: MessagingApi) -> str | None:
    """
    處理 HQ 群組訊息。

    回傳值：
    - str：要在 HQ 群組回覆的 ack 訊息
    - None：非調貨相關，不在群組回覆
    """
    intent = _detect_intent(text)
    if intent is None:
        return None

    request = _find_matching_request(text)
    if not request:
        return None

    if intent == "available":
        return _handle_available(request, line_api)
    elif intent == "ordering":
        wait_time = _extract_wait_time(text)
        return _handle_ordering(request, wait_time, line_api)

    return None


def _detect_intent(text: str) -> str | None:
    """偵測 HQ 回覆意圖"""
    available_kw = ["有貨", "可以調", "可調", "可以出", "有的", "有庫存", "可出貨", "可以出貨", "有存貨"]
    ordering_kw = ["叫貨", "訂貨", "沒有", "缺貨", "需要等", "要等", "需叫", "沒貨", "無貨"]

    if any(kw in text for kw in available_kw):
        return "available"
    if any(kw in text for kw in ordering_kw):
        return "ordering"
    # 有時間詞（週/天/個月）也視為需要等待
    if re.search(r"\d+[-~～到]\d*\s*[週周天個月]|\d+\s*[週周天個月]", text):
        return "ordering"
    return None


def _extract_wait_time(text: str) -> str:
    """從 HQ 回覆中提取等待時間"""
    m = re.search(r"(\d+[-~～到]\d*\s*[週周天個月]|\d+\s*[週周天個月][以內左右]?)", text)
    if m:
        return m.group(1)
    if "一個月" in text:
        return "一個月"
    if "兩個月" in text or "2個月" in text:
        return "兩個月"
    return "幾週"


def _find_matching_request(text: str) -> dict | None:
    """先嘗試從文字中抓產品碼比對，找不到就取最新的 pending"""
    m = re.search(r"[A-Za-z][A-Za-z0-9\-_]{2,}", text)
    if m:
        req = restock_store.find_pending_by_product(m.group(0))
        if req:
            return req
    return restock_store.get_latest_pending()


def _handle_available(request: dict, line_api: MessagingApi) -> str | None:
    """有貨可調 → 建立 Ecount 訂單，通知客戶"""
    user_id = request["user_id"]
    prod_name = request["prod_name"]
    prod_cd = request["prod_cd"]
    qty = request["qty"]

    restock_store.update_status(request["id"], "available")

    cust_code = customer_store.get_ecount_cust_code(
        user_id, default=settings.ECOUNT_DEFAULT_CUST_CD
    )

    _phone = (customer_store.get_by_line_id(user_id) or {}).get("phone", "") or ""
    slip_no = ecount_client.save_order(
        cust_code=cust_code,
        items=[{"prod_cd": prod_cd, "qty": qty}],
        phone=_phone,
    )

    if slip_no:
        restock_store.update_status(request["id"], "confirmed")
        print(f"[restock] 訂單建立成功: {slip_no} | {cust_code} | {prod_name} x{qty}")
        _push_to_customer(user_id, tone.restock_order_confirmed(prod_name, qty, slip_no), line_api)
    else:
        print(f"[restock] 訂單建立失敗: {cust_code} | {prod_name} x{qty}")
        from storage.issues import issue_store
        issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個（調貨）")

    return "好的"


def _handle_ordering(request: dict, wait_time: str, line_api: MessagingApi) -> str | None:
    """需要叫貨 → 推送客戶詢問是否能等"""
    user_id = request["user_id"]
    prod_name = request["prod_name"]
    qty = request["qty"]

    restock_store.update_status(request["id"], "ordering", wait_time)

    state_manager.set(user_id, {
        "action": "awaiting_wait_confirm",
        "restock_id": request["id"],
        "prod_name": prod_name,
        "prod_cd": request["prod_cd"],
        "qty": qty,
        "wait_time": wait_time,
    })

    _push_to_customer(user_id, tone.restock_wait_ask(prod_name, qty, wait_time), line_api)
    return "好的"


def _push_to_customer(user_id: str, msg: str, line_api: MessagingApi):
    try:
        line_api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=msg)],
            )
        )
    except Exception as e:
        print(f"[restock] 推送客戶失敗: {e}")
