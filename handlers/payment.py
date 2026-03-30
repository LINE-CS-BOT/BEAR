"""
轉帳確認處理

偵測客戶傳來的轉帳通知訊息，
記錄到 payment_confirmations DB，
並回覆「等等確認喔」類型的語氣。
"""

from storage.payments import payment_store
from handlers import tone

# 轉帳相關關鍵字
_PAYMENT_KW = [
    "已轉", "轉帳", "匯款", "付款", "已付", "打款",
    "轉過去", "已匯", "匯過去", "匯給", "轉給",
    "ATM", "atm", "網銀", "網路銀行", "網路轉帳",
    "收到款", "確認款項", "查收", "匯款細項",
]


def is_payment_message(text: str) -> bool:
    """判斷是否為轉帳確認訊息"""
    return any(kw in text for kw in _PAYMENT_KW)


def handle_payment(user_id: str, text: str) -> str:
    """記錄轉帳，回覆等等確認"""
    payment_store.add(user_id, text)
    return tone.payment_ack()
