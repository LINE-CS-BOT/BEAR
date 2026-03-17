"""
配送詢問處理

客戶詢問送貨時間 → bot 回覆「跟司機確認後回覆」
並記錄到 delivery_store，納入每小時待處理清單。
內部人員跟司機確認後手動回覆客戶，再用 ✅ D{id} 標記完成。
"""

import random

from storage.delivery import delivery_store
from handlers import tone


def handle_delivery(user_id: str, text: str) -> str:
    """
    記錄配送詢問，回覆等待確認訊息。

    若此客戶已有未處理的詢問，不重複新增記錄，
    只回覆「已在確認中」。
    """
    if delivery_store.has_pending(user_id):
        # 已有未處理的詢問，避免重複記錄
        b = tone.boss()
        return random.choice([
            f"{b}，送貨時間還在跟司機確認中，確認好馬上通知您哦",
            f"還在確認中{tone.suffix_light()} 確認好再通知{b}哦",
            f"稍等一下{tone.suffix_light()} 還在跟司機確認，有消息馬上告訴{b}",
        ])

    # 新增記錄
    delivery_store.add(user_id, text)

    b = tone.boss()
    return random.choice([
        f"跟司機確認時間後再回覆您{tone.suffix_light()}",
        f"幫{b}跟司機確認一下，確認好馬上通知您哦",
        f"送貨時間讓我跟司機確認一下{tone.suffix_light()} 確認好再通知{b}哦",
        f"等等幫{b}確認司機的時間，有消息馬上回覆您",
    ])
