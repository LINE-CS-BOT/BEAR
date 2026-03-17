"""
對話跟進處理

- 每小時執行一次
- 24小時沒回應 → 發提醒訊息
- 48小時沒回應 → 清除狀態
"""

from storage.persistent_state import persistent_state_store, EXPIRE_AFTER_HOURS, REMIND_AFTER_HOURS
from storage.state import state_manager

# 各狀態的提醒文字
_REMIND_TEXT = {
    "awaiting_quantity":              "您好～請問您剛才詢問的商品，需要幾個呢？😊",
    "awaiting_restock_qty":           "您好～請問您需要調貨的數量是多少呢？",
    "awaiting_wait_confirm":          "您好～想請問您是否願意等待補貨呢？",
    "awaiting_order_confirm":         "您好～請問您剛才的訂單要確認送出嗎？",
    "awaiting_group_address_confirm": "您好～請問您的收件地址確認了嗎？",
    "awaiting_order_id":              "您好～請問您的訂單編號是？",
}

_DEFAULT_REMIND = "您好～請問還需要幫忙嗎？😊"


def check_and_followup(line_api) -> dict:
    """
    檢查過期狀態，發提醒或清除。
    回傳: {"reminded": N, "expired": N}
    """
    reminded = 0
    expired  = 0

    # ── 48小時過期 → 直接清除 ──────────────────────────────
    for user_id in persistent_state_store.get_expired():
        persistent_state_store.delete(user_id)
        state_manager.clear(user_id)
        expired += 1
        print(f"[followup] 清除過期狀態: {user_id}")

    # ── 24小時未回 → 提醒 ─────────────────────────────────
    for entry in persistent_state_store.get_need_remind():
        user_id = entry["user_id"]
        action  = entry["action"]
        text    = _REMIND_TEXT.get(action, _DEFAULT_REMIND)
        try:
            line_api.push_message(
                user_id,
                {"type": "text", "content": text},
            )
            persistent_state_store.mark_reminded(user_id)
            reminded += 1
            print(f"[followup] 提醒送出: {user_id} ({action})")
        except Exception as e:
            print(f"[followup] 提醒失敗: {user_id} → {e}")

    return {"reminded": reminded, "expired": expired}
