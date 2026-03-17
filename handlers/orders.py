"""
訂單查詢處理

Ecount OAPI 目前無訂單查詢端點。
接到客戶詢問時：
  1. 記錄到 issue_store（type='order_query'）
  2. 回覆「我幫您查看看 請稍等」語氣
  3. 納入每小時待處理清單，由人工跟進回覆
"""

from handlers import tone
from storage.issues import issue_store


def handle_order_tracking(user_id: str, text: str) -> str:
    """記錄訂單查詢，回覆稍等語氣"""
    issue_store.add(user_id, "order_query", text)
    return tone.order_tracking_ack()
