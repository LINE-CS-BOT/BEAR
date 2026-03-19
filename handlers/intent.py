from enum import Enum


class Intent(Enum):
    INVENTORY = "inventory"
    PRICE = "price"
    ORDER_TRACKING = "order_tracking"
    DELIVERY = "delivery"
    BUSINESS_HOURS = "business_hours"
    GREETING = "greeting"
    CONFIRMATION = "confirmation"
    BARGAINING = "bargaining"
    SPEC = "spec"
    RETURN = "return"
    MULTI_PRODUCT = "multi_product"
    ADDRESS_CHANGE = "address_change"
    COMPLAINT = "complaint"
    URGENT_ORDER = "urgent_order"
    NOTIFY_REQUEST = "notify_request"
    CHECKOUT = "checkout"
    MACHINE_SIZE = "machine_size"
    VISIT_STORE = "visit_store"
    UNKNOWN = "unknown"


_MACHINE_SIZE_KEYWORDS = [
    "標準", "中巨", "巨無霸", "小k", "k霸", "小K", "K霸", "迷你機",
]

# ── 新場景關鍵字 ──────────────────────────────────────

_COMPLAINT_KEYWORDS = [
    "壞了", "壞掉", "破損", "有問題", "不對", "爛了", "有缺陷",
    "品質有問題", "貨有問題", "不滿意", "收到問題", "少了", "缺了",
    "少一個", "數量不對", "出錯了", "投訴", "客訴", "要投訴", "太差了",
    "有壞", "壞的", "有破", "破掉", "漏了", "錯了一個",
]

_RETURN_KEYWORDS = [
    "退貨", "退換", "換貨", "換一個", "要退", "想退",
    "退掉", "不要了", "退錢", "退款", "要換", "申請退",
    "可以退", "可以換", "能退嗎", "能換嗎", "辦退貨",
]

_ADDRESS_CHANGE_KEYWORDS = [
    "改地址", "換地址", "地址改", "地址錯了", "送錯地址",
    "改收件", "換收件", "改配送", "地址要改", "地址變更",
    "收件人", "改名字", "地址打錯", "地址填錯",
]

_CHECKOUT_KEYWORDS = [
    "好了", "沒了", "沒有了", "就這樣", "就這些", "以上",
    "結帳", "下單", "確認訂單", "送出訂單", "訂這些", "就訂這些",
    "不用了謝謝", "這樣就好", "這樣就夠了",
]

_BARGAINING_KEYWORDS = [
    "便宜一點", "可以優惠", "打折", "折扣", "算便宜", "算我便宜",
    "有沒有優惠", "有優惠嗎", "可以便宜", "殺價", "讓一點",
    "算便宜點", "再便宜", "優惠一下", "有折扣嗎",
]

_SPEC_KEYWORDS = [
    "有什麼顏色", "顏色", "尺寸", "規格", "幾公分", "幾公斤",
    "重量", "材質", "多大", "多重", "幾號", "介紹一下",
    "這款怎麼樣", "好不好", "推薦", "品質", "哪款好",
    "有幾種", "容量", "幾ml", "幾ML", "幾升", "長寬高",
]

# 正式查單（有訂單號可查，或明確問查詢）
_ORDER_KEYWORDS = [
    "到了嗎", "到了没", "我的貨", "我的訂單", "訂單狀態",
    "出貨了嗎", "出貨了没", "寄出了嗎", "到货了吗",
    "查訂單", "查單",
]

# 非制式催貨（沒有單號、純粹催問人工）→ 通知 staff + 記錄
_URGENT_ORDER_KEYWORDS = [
    "什麼時候出", "幾時出", "怎麼還沒出", "還沒出嗎",
    "催一下", "到哪裡了", "貨呢", "出了沒",
    "何時出", "幾時到", "什麼時候到", "多久會到",
    "怎麼還沒來", "還沒到嗎", "等很久了", "等好久",
]

_PRICE_KEYWORDS = [
    "多少錢", "多少钱", "幾錢", "幾塊", "價格", "价格",
    "單價", "售價", "多少", "報價", "價位", "價錢",
    "多少一", "一個多少", "一箱多少", "一盒多少",
]

_NOTIFY_REQUEST_KEYWORDS = [
    "有貨通知我", "到貨通知我", "有貨了通知", "到了通知我",
    "有貨通知一下", "有貨再通知", "有貨時通知", "有貨叫我",
    "到貨告訴我", "有貨告訴我", "來貨通知我", "到貨叫我",
    "到貨再告訴", "有貨再告訴", "有貨提醒", "到貨提醒",
]

_INVENTORY_KEYWORDS = [
    "有貨嗎", "有货吗", "有沒有貨", "有没有货", "庫存",
    "可以訂嗎", "能訂嗎", "還有嗎", "還有", "缺貨", "有得訂",
    "可以下單", "訂購", "有沒有",
    "都有嗎", "都有貨嗎", "各有嗎",   # 複合詢問（在 inventory.py 再細分）
]

_DELIVERY_KEYWORDS = [
    "送貨", "配送", "送達", "什麼時候送", "幾時送",
    "可以送嗎", "有送貨嗎", "送貨時間", "送到", "宅配",
]

_HOURS_KEYWORDS = [
    # ── 明確問「幾點」 ────────────────────────────
    "營業時間", "上班時間",
    "幾點開", "幾點關", "幾點到幾點",
    "幾點營業", "營業到", "幾時開", "幾時營業",
    "開到幾點", "幾點打烊", "幾點收",
    "幾點上班", "幾點下班", "幾點有人",

    # ── 明確問「有沒有開/在」 ─────────────────────
    "有開嗎", "有上班嗎", "有在開", "有開門嗎",
    "有在上班", "有在營業", "有營業嗎",
    "今天有開", "今天有沒有開", "今天有營業", "今天有上班",
    "今天上班嗎", "今天營業嗎",
    "你們有開", "你們有營業",
    "還有開", "還有上班", "還有營業",
    "有沒有開", "有沒有上班", "有沒有營業",

    # ── 否定問法 ──────────────────────────────────
    "沒開嗎", "今天沒開", "沒在開", "沒上班嗎", "沒有開嗎",

    # ── 公休相關 ──────────────────────────────────
    "公休", "幾號公休",

    # ⚠️ 移除：「休息」「假日」「開門」「週一」「禮拜一」「今天開」
    #    「幾點到」「幾點可以」「上班嗎」「到幾點」「今天上班」
    #    → 這些太泛，容易在非營業時間情境誤觸發
]


_GREETING_KEYWORDS = [
    "你好", "您好", "嗨嗨", "哈囉", "安安", "嗨",
    "早安", "午安", "晚安", "早哦", "早唷", "hi", "hello",
    # 有人在嗎類（availability check）
    "在不在", "有人嗎", "有在嗎", "在嗎", "老闆在嗎",
    "請問有人嗎", "在線嗎", "在線上嗎", "人在嗎",
]

_CONFIRMATION_KEYWORDS = [
    "好的", "好喔", "好哦", "好👌", "好！", "謝謝", "感謝", "感恩",
    "了解", "收到", "辛苦了", "沒問題", "可以", "要了",
    "OK", "ok", "Ok", "對", "是的", "是哦", "嗯嗯",
    "👌", "哈哈", "哦哦", "好",
]


def detect_intent(text: str) -> Intent:
    # 忽略 LINE 收回訊息
    if "此內容已收回" in text:
        return Intent.CONFIRMATION  # 靜默處理

    # 娃娃機尺寸詢問（靜默記錄）
    for kw in _MACHINE_SIZE_KEYWORDS:
        if kw in text:
            return Intent.MACHINE_SIZE

    # 投訴/問題（最具體，優先判）
    for kw in _COMPLAINT_KEYWORDS:
        if kw in text:
            return Intent.COMPLAINT

    # 退換貨
    for kw in _RETURN_KEYWORDS:
        if kw in text:
            return Intent.RETURN

    # 地址更改
    for kw in _ADDRESS_CHANGE_KEYWORDS:
        if kw in text:
            return Intent.ADDRESS_CHANGE

    # 非制式催貨（比正式查單先判，避免「到了嗎」干擾）
    for kw in _URGENT_ORDER_KEYWORDS:
        if kw in text:
            return Intent.URGENT_ORDER

    # 正式查單
    for kw in _ORDER_KEYWORDS:
        if kw in text:
            return Intent.ORDER_TRACKING

    # 砍價（在價格之前判）
    for kw in _BARGAINING_KEYWORDS:
        if kw in text:
            return Intent.BARGAINING

    # 運費相關問題優先排除，不走 PRICE（→ UNKNOWN 轉真人）
    _SHIPPING_WORDS = ["運費", "含運", "郵寄", "宅配費", "快遞費", "物流費", "運送費"]
    if not any(w in text for w in _SHIPPING_WORDS):
        for kw in _PRICE_KEYWORDS:
            if kw in text:
                return Intent.PRICE

    # 規格/介紹（在庫存之前，避免「有什麼顏色」被判成庫存）
    for kw in _SPEC_KEYWORDS:
        if kw in text:
            return Intent.SPEC

    # 營業時間比庫存先判（避免「還有開嗎」「有沒有開」被誤判成庫存）
    for kw in _HOURS_KEYWORDS:
        if kw in text:
            return Intent.BUSINESS_HOURS

    # 到貨通知登記（含「有貨通知我」→ 比庫存查詢先判）
    for kw in _NOTIFY_REQUEST_KEYWORDS:
        if kw in text:
            return Intent.NOTIFY_REQUEST

    # 排除「有沒有空/時間/人/機會」等非庫存語境
    _INV_EXCLUDE = ["有沒有空", "有沒有時間", "有沒有人", "有沒有機會", "有沒有辦法", "有沒有問題"]
    if not any(w in text for w in _INV_EXCLUDE):
        for kw in _INVENTORY_KEYWORDS:
            if kw in text:
                return Intent.INVENTORY

    for kw in _DELIVERY_KEYWORDS:
        if kw in text:
            return Intent.DELIVERY

    for kw in _GREETING_KEYWORDS:
        if kw in text.lower():
            return Intent.GREETING

    # 結帳（在 CONFIRMATION 之前，避免「好了」被誤判成確認）
    # 問句排除：「好了嗎」「有沒有結帳」之類問法不是結帳指令
    _CHECKOUT_QUESTION_EXCLUDE = ["好了嗎", "好了沒", "好了嗎？", "有沒有結帳"]
    if not any(e in text for e in _CHECKOUT_QUESTION_EXCLUDE):
        for kw in _CHECKOUT_KEYWORDS:
            if kw in text:
                return Intent.CHECKOUT

    # 確認語放最後（避免「好的有貨嗎」被誤判）
    # 若同時含有運費相關詞（如「好，謝謝，含運多少？」）→ 轉真人，不走確認
    _SHIPPING_WORDS = ["運費", "含運", "郵寄", "宅配費", "快遞費", "物流費", "運送費"]
    if not any(w in text for w in _SHIPPING_WORDS):
        for kw in _CONFIRMATION_KEYWORDS:
            if text.strip() == kw or text.strip().startswith(kw):
                return Intent.CONFIRMATION

    # 到店預告（放 UNKNOWN 前，關鍵字較具體）
    from handlers.visit import is_visit_message
    if is_visit_message(text):
        return Intent.VISIT_STORE

    return Intent.UNKNOWN
