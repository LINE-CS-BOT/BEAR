# -*- coding: utf-8 -*-
"""
本地模擬 Bot 回覆（不需要 LINE / webhook）
用法：python simulate.py
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

import random
from handlers.intent import detect_intent, Intent
from handlers.inventory import _extract_product
from handlers import tone


def simulate(text: str):
    intent = detect_intent(text)
    label = {
        Intent.INVENTORY:      "庫存查詢",
        Intent.ORDER_TRACKING: "訂單查詢",
        Intent.DELIVERY:       "送貨查詢",
        Intent.BUSINESS_HOURS: "營業時間",
        Intent.GREETING:       "打招呼",
        Intent.CONFIRMATION:   "確認",
        Intent.UNKNOWN:        "不明（升級真人）",
    }[intent]

    print("─" * 40)
    print(f"客戶: {text}")
    print(f"意圖: {label}")

    if intent == Intent.INVENTORY:
        product = _extract_product(text)
        print(f"識別產品: [{product}]")
        if product:
            print(f"Bot 回覆: {tone.in_stock(product)}")
        else:
            print(f"Bot 回覆: {tone.ask_product()}")

    elif intent == Intent.GREETING:
        print(f"Bot 回覆: {tone.greeting_reply()}")

    elif intent == Intent.CONFIRMATION:
        print(f"Bot 回覆: {tone.confirmation_ack()}")

    elif intent == Intent.BUSINESS_HOURS:
        from handlers.hours import handle_business_hours
        print(f"Bot 回覆: {handle_business_hours()}")

    elif intent == Intent.ORDER_TRACKING:
        print(f"Bot 回覆: (查訂單功能)")

    elif intent == Intent.DELIVERY:
        from handlers.delivery import handle_delivery
        print(f"Bot 回覆: {handle_delivery()}")

    elif intent == Intent.UNKNOWN:
        print(f"Bot 回覆: {tone.escalating()}  <-- 同時推通知給真人")

    print()


# ── 測試案例 ───────────────────────────────────────
tests = [
    "歌林電風扇還有嗎",
    "金豬報喜存錢筒 還有嗎",
    "你們還有金豬報喜存錢筒嗎",
    "有沒有歌林電風扇",
    "請問ABC-001有貨嗎",
    "你好",
    "謝謝",
    "今天有營業嗎",
    "到了嗎",
    "10個",
    "算了不要了",
]

for t in tests:
    simulate(t)
