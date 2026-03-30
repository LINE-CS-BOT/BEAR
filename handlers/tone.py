"""
小蠻牛-新北旗艦店 真人客服語氣模組
分析自 269 則真實 LINE OA 對話記錄（2025/09 ~ 2026/03）

語氣特徵：
- 稱客戶為「老闆」
- 語尾常用：唷、喔、嘿、囉、哦、~~
- 確認中：「我先查一下嘿」「等等我」
- 有貨：「有喔」「是有的！」「老闆有唷」
- 無貨：「目前沒貨唷」「這款沒有了，到貨再通知您哦」
- 道歉：「不好意思」「抱歉嘿」
- 感謝：「謝謝老闆」「感謝您☺️」
- 問候：「哈囉~~」「老闆您好」
"""

import random

from config import settings as _settings

# ── 營業日字串（從 settings 動態產生）─────────────────
_WEEKDAY_NAMES = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '日'}

def _biz_days_label() -> str:
    """產生如 '週二～週日（週一公休）' 的營業日標示"""
    biz = _settings.business_days_list()
    all_days = set(range(1, 8))
    off = sorted(all_days - set(biz))
    if not off:
        return "每日營業"
    off_label = "、".join(f"週{_WEEKDAY_NAMES[d]}" for d in off)
    biz_names = [f"週{_WEEKDAY_NAMES[d]}" for d in sorted(biz)]
    if len(biz) >= 2:
        biz_label = f"{biz_names[0]}～{biz_names[-1]}"
    else:
        biz_label = "、".join(biz_names)
    return f"{biz_label}（{off_label}公休）"


# ── 稱呼 ─────────────────────────────────────────
def boss() -> str:
    """隨機選一個稱呼，大部分都是「老闆」"""
    return random.choice(["老闆", "老闆", "老闆", "老闆", "您"])


# ── 語尾詞 ────────────────────────────────────────
def suffix_light() -> str:
    """輕鬆語尾（較短回覆用）"""
    return random.choice(["唷", "喔", "嘿", "哦", "囉"])


def suffix_warm() -> str:
    """溫暖語尾（~~系列）"""
    return random.choice(["~~", "～～", "哦~~", "唷~~"])


# ── 問候語 ────────────────────────────────────────
def greeting() -> str:
    return random.choice([
        "哈囉~~",
        "哈囉~~~~",
        f"{boss()}您好",
        f"{boss()}好",
        "您好！",
    ])


# ── 確認中 ────────────────────────────────────────
def checking() -> str:
    return random.choice([
        f"我先確認一下等等我{suffix_light()}",
        f"等我查一下{suffix_light()}",
        f"我問一下，你等我一下",
        f"幫{boss()}查一下哦",
        f"我查一下，等等回您{suffix_light()}",
        f"上班馬上查詢{suffix_light()}",
    ])


# ── 有貨回覆 ──────────────────────────────────────
def in_stock(name: str) -> str:
    templates = [
        f"{boss()}有喔，「{name}」有貨，請問要幾個呢？",
        f"有唷！「{name}」有貨，{boss()}需要幾個呢？",
        f"是有的！「{name}」有貨，{boss()}要下單的話告訴我嘿",
        f"老闆有喔，「{name}」有貨，請問要幾個？",
        f"「{name}」有{suffix_light()} {boss()}需要幾個？",
    ]
    return random.choice(templates)


def in_stock_low(name: str) -> str:
    """數量偏少時用（剩不多，但不顯示具體數量）"""
    templates = [
        f"「{name}」還有少量現貨，{boss()}要的話趕快唷，請問要幾個？",
        f"{boss()}，「{name}」數量不多了，要的話快點確認嘿，請問要幾個？",
        f"有喔但剩不多，「{name}」{boss()}要幾個？",
    ]
    return random.choice(templates)


# ── 無貨回覆 ──────────────────────────────────────
def out_of_stock(name: str) -> str:
    templates = [
        f"不好意思，「{name}」目前沒貨了，到貨再通知{boss()}哦",
        f"「{name}」目前沒有{suffix_light()} 有調貨中，到貨馬上通知您",
        f"這款沒有了，我們有和總部調貨，到貨再和{boss()}通知哦",
        f"目前沒現貨{suffix_light()} 「{name}」到貨後馬上通知{boss()}",
        f"「{name}」現在缺貨，等到貨我們會通知您哦",
    ]
    return random.choice(templates)


def out_of_stock_reserved(name: str, stock: int) -> str:
    """有庫存但全被訂單佔用"""
    templates = [
        f"不好意思{suffix_light()} 「{name}」庫存 {stock} 件都已被預留了，有補貨馬上通知{boss()}",
        f"「{name}」目前 {stock} 件都已有訂單，{boss()}稍等一下，補貨後通知您哦",
    ]
    return random.choice(templates)


# ── 查無此產品 ────────────────────────────────────
def product_not_found(product: str) -> str:
    templates = [
        f"目前沒找到「{product}」的資料{suffix_light()} 讓我幫{boss()}問問看，確認後回覆您哦",
        f"這個編號查不到{suffix_light()} 幫{boss()}向總部確認一下，稍後回覆您",
        f"不好意思，「{product}」目前沒有資料，我幫您查一下嘿",
    ]
    return random.choice(templates)


# ── 詢問哪個產品 ──────────────────────────────────
def ask_product() -> str:
    return random.choice([
        f"哈囉~~ 請問{boss()}要查哪個產品的庫存呢？（輸入產品編號或名稱）",
        f"{boss()}好，請問要查哪款的庫存呢？",
        f"請問要查哪個品項的庫存{suffix_light()}（輸入編號或品名）",
    ])


# ── 通用確認/感謝 ─────────────────────────────────
def ok() -> str:
    return random.choice(["好的", "收到", "好唷", "好哦", "好"])


def thanks() -> str:
    return random.choice([
        "謝謝老闆☺️",
        "謝謝您的支持",
        "感謝您",
        "感謝老闆",
    ])


# ── 道歉 ──────────────────────────────────────────
def sorry() -> str:
    return random.choice([
        "不好意思",
        "抱歉嘿",
        "非常抱歉",
        "不好意思造成您的困擾",
    ])


# ── 預設/不懂意思 ─────────────────────────────────
def default_menu() -> str:
    b = boss()
    return f"哈囉~~ 請問有什麼可以幫{b}的嗎？"


# ── 營業時間 ──────────────────────────────────────
def business_hours_open(hours_start: str, hours_end: str, address: str) -> str:
    return (
        f"有開哦！歡迎老闆來\n"
        f"🕐 {hours_start} 開門，{hours_end} 休息\n"
        f"📅 {_biz_days_label()}\n"
        f"📍 {address}"
    )


def business_hours_closed(hours_start: str, hours_end: str, address: str) -> str:
    """保留備用（目前 hours.py 不呼叫此函式）"""
    return business_hours_open(hours_start, hours_end, address)


def business_hours_specific_open(date_label: str, hours_start: str, hours_end: str, address: str) -> str:
    return random.choice([
        f"{date_label} 有開唷！\n🕐 {hours_start} 開門，{hours_end} 休息\n📍 {address}",
        f"{date_label} 有營業喔，歡迎老闆來\n🕐 {hours_start}～{hours_end}\n📍 {address}",
    ])


def business_hours_specific_closed(date_label: str, hours_start: str, hours_end: str) -> str:
    return random.choice([
        f"不好意思，{date_label} 我們休息唷\n📅 {_biz_days_label()} {hours_start}～{hours_end} 營業",
        f"{date_label} 是公休日嘿，歡迎其他時間來\n📅 {_biz_days_label()} {hours_start}～{hours_end}",
    ])


def business_hours_holiday(hours_start: str, hours_end: str, address: str) -> str:
    return random.choice([
        f"不好意思，今天休息哦\n📅 我們{_biz_days_label()} {hours_start} ～ {hours_end} 營業",
        f"今天休息唷，{boss()}明天再來找我們哦\n🕐 {hours_start} ～ {hours_end}（{_biz_days_label()}）",
    ])


def business_hours_after_close(hours_start: str, hours_end: str) -> str:
    return random.choice([
        f"今天已打烊囉 🌙\n明天 {hours_start} 再開門，歡迎老闆再來～",
        f"不好意思，今天 {hours_end} 已經打烊了\n明天 {hours_start} 開門，{boss()}明天見 😊",
    ])


def business_hours_not_open_yet(hours_start: str, hours_end: str, address: str) -> str:
    return random.choice([
        f"還沒開門唷！今天 {hours_start} 開始營業\n📍 {address}",
        f"{hours_start} 才開門哦，{boss()}等一下再來 😊\n📍 {address}",
    ])


# ── 問候語回覆 ────────────────────────────────────
def greeting_reply() -> str:
    b = boss()
    return random.choice([
        f"哈囉~~ {b}您好！請問有什麼可以幫您的嗎？",
        f"{b}好！請問需要什麼幫助{suffix_light()}",
        f"哈囉~~ 請問有什麼可以幫{b}的嗎？",
        f"您好{suffix_light()} 請問有什麼需要嗎？",
    ])


# ── 確認語回覆 ────────────────────────────────────
def confirmation_ack() -> str:
    b = boss()
    return random.choice([
        f"不客氣唷，有問題隨時找我哦",
        f"不客氣{suffix_light()} {b}有問題隨時找我哦",
        f"不會不會~~ 有需要再找我哦",
        f"哈哈不客氣~~ 有問題隨時說{suffix_light()}",
    ])


# ── 轉真人客服 ────────────────────────────────────
def escalating() -> str:
    return random.choice([
        "不好意思稍等一下喔~",
        "不好意思，稍等一下唷~",
        "稍等一下喔~~",
        "不好意思，等一下嘿~",
    ])


# ── 詢問數量 ──────────────────────────────────────
def ask_quantity(name: str) -> str:
    b = boss()
    return random.choice([
        f"好的！請問{b}需要幾個「{name}」呢？",
        f"請問{b}要幾個「{name}」{suffix_light()}",
        f"「{name}」請問要幾個呢？",
    ])


def ask_product_clarify(keyword: str, candidates: list) -> str:
    """多款商品符合關鍵字，請客戶選擇"""
    lines = [f"請問「{keyword}」是哪一款呢？"]
    for i, (code, name) in enumerate(candidates, 1):
        lines.append(f"{i}. {name}")
    lines.append("\n回覆數字或品名就好嘿～")
    return "\n".join(lines)


def preorder_ask_qty(name: str) -> str:
    """預購商品，問客戶要幾個"""
    b = boss()
    return random.choice([
        f"「{name}」是預購商品唷！請問{b}需要幾個呢？",
        f"這款是預購商品{suffix_light()} 請問{b}要訂幾個呢？",
        f"「{name}」目前是預購的哦～ 請問{b}需要幾個？",
    ])


def out_of_stock_ask_qty(name: str) -> str:
    """缺貨，但先問數量以便詢問總公司調貨"""
    b = boss()
    return random.choice([
        f"不好意思{suffix_light()} 「{name}」目前沒貨，請問{b}需要幾個？我幫您和總部詢問調貨",
        f"「{name}」目前缺貨{suffix_light()} 請問{b}需要幾個？我向總部確認調貨時間哦",
        f"不好意思，「{name}」暫時缺貨，請問需要幾個？我幫您問總部看看唷",
    ])


def preorder_ask_qty(name: str, po_info: str = "") -> str:
    """預購品，問客戶要幾個"""
    b = boss()
    # 從 PO文提取到貨時間
    eta = ""
    if po_info:
        import re
        m = re.search(r'(預計|預估|大約|約)?\s*(\d+月[中下旬底]*|\d+/\d+)', po_info)
        if m:
            eta = f"，{m.group(0).strip()}到貨"
        else:
            m2 = re.search(r'(\d+月\S*到貨)', po_info)
            if m2:
                eta = f"，{m2.group(1)}"
    return random.choice([
        f"「{name}」目前是預購品{eta}{suffix_light()} 請問{b}要預訂幾個呢？",
        f"這款「{name}」是預購款{eta}哦～ 請問{b}需要幾個？",
    ])


def restock_inquiry_sent(name: str, qty: int) -> str:
    """已向總公司詢問，回覆客戶（說明已記錄需求，等回覆）"""
    b = boss()
    return random.choice([
        f"好的！已幫{b}向總部詢問「{name}」× {qty} 個的調貨情況，\n確認後馬上通知您哦",
        f"收到{suffix_light()} 已記錄{b}需要「{name}」× {qty} 個，\n正在向總部確認，有消息馬上告訴您",
        f"「{name}」× {qty} 個已向總部詢問囉，\n確認調貨情況後會通知{b}哦",
    ])


# ── 調貨成功建立訂單 ──────────────────────────────
def restock_order_confirmed(name: str, qty: int, slip_no: str) -> str:
    """HQ 有貨可調，訂單已建立，通知客戶（含前文脈絡）"""
    b = boss()
    return random.choice([
        f"{b}您好！關於您之前詢問的「{name}」× {qty} 個調貨，\n"
        f"已確認總公司有貨可調，訂單已幫您建立 📋\n"
        f"感謝您的等待☺️",

        f"{b}好消息！您之前問的「{name}」× {qty} 個，\n"
        f"總公司確認有貨，已幫您建立訂單{suffix_light()}\n"
        f"感謝您的耐心等候",
    ])


# ── 詢問客戶是否能等叫貨 ──────────────────────────
def restock_wait_ask(name: str, qty: int, wait_time: str) -> str:
    """HQ 需要叫貨，問客戶是否能等（含前文脈絡）"""
    b = boss()
    return random.choice([
        f"{b}您好！關於您詢問的「{name}」× {qty} 個，\n"
        f"向總公司確認後，目前需要跟廠商叫貨，\n"
        f"大概需要 {wait_time}，請問{b}方便等待嗎？",

        f"不好意思{suffix_light()} 您之前問的「{name}」× {qty} 個，\n"
        f"總公司表示需要叫貨，預計等待時間約 {wait_time}，\n"
        f"請問可以接受嗎？",
    ])


# ── 客戶確認等待，叫貨訂單建立 ────────────────────
def restock_wait_confirmed(name: str, qty: int, slip_no: str) -> str:
    """客戶同意等，訂單建立（確認客戶知道在等什麼）"""
    b = boss()
    return random.choice([
        f"謝謝{b}！「{name}」× {qty} 個叫貨訂單已建立 📋\n"
        f"到貨後馬上通知您哦，感謝您的耐心等待☺️",

        f"好的{suffix_light()} 已幫{b}建立「{name}」× {qty} 個的訂單！\n"
        f"到貨後第一時間通知您哦",
    ])


# ── 客戶不願等待 ──────────────────────────────────
def restock_wait_declined(name: str) -> str:
    """客戶不願等，感謝並告知到貨仍會通知"""
    b = boss()
    return random.choice([
        f"好的，了解{suffix_light()} 感謝{b}的詢問！\n"
        f"「{name}」若日後有現貨，我們會再通知您哦",

        f"沒關係{suffix_light()} 謝謝{b}的詢問，\n"
        f"「{name}」到貨後我們會主動通知您，歡迎再來找我哦",
    ])


# ── 到貨通知（push 給等待中的客戶）──────────────────
def restock_back_in_stock(name: str, code: str) -> str:
    """庫存自動到貨，主動 push 通知客戶"""
    b = boss()
    return random.choice([
        f"🎉 {b}好！您之前詢問的「{name}」（{code}）現在已有現貨囉！\n"
        f"有需要的話歡迎來訊，我幫您安排哦",

        f"老闆好消息！您之前等的「{name}」到貨了 🎉\n"
        f"有需要的話跟我說一聲{suffix_light()} 幫您安排",

        f"🎉 通知一下{b}，「{name}」（{code}）有貨了！\n"
        f"有要的話歡迎再來找我哦",
    ])


# ── 訂單建立成功 ──────────────────────────────────
def order_confirmed(name: str, qty: int, slip_no: str) -> str:
    b = boss()
    return random.choice([
        f"好的！已幫{b}建立訂單 📋\n「{name}」× {qty} 個\n謝謝{b}的訂購☺️",
        f"收到！「{name}」{qty} 個已登記{suffix_light()}\n感謝{b}☺️",
    ])


# ── 購物車 ────────────────────────────────────────
def cart_item_added(cart: list[dict]) -> str:
    """加入品項後顯示目前購物車，並詢問是否繼續"""
    lines = ["好的！目前清單："]
    for item in cart:
        lines.append(f"  • {item['prod_name']} × {item['qty']}")
    lines.append("")
    lines.append("還有其他要訂的嗎？如果好了就跟我說幫你送出唷✉️")
    return "\n".join(lines)


def checkout_confirmed(cart: list[dict]) -> str:
    """結帳成功回覆（不顯示單號）"""
    b = boss()
    lines = [f"收到！幫{b}送出訂單了 📋"]
    for item in cart:
        lines.append(f"  • {item['prod_name']} × {item['qty']}")
    lines.append(f"感謝{b}的訂購☺️")
    return "\n".join(lines)


def cart_empty_checkout() -> str:
    """結帳時購物車是空的"""
    return "目前沒有訂購任何商品喔，有需要再告訴我～"


# ── 訂單查詢 ──────────────────────────────────────
def order_tracking_ack() -> str:
    b = boss()
    return random.choice([
        f"我幫{b}查看看{suffix_light()} 請稍等",
        f"好的，幫{b}查一下，請稍等嘿",
        f"我查查看{suffix_light()} 稍等一下哦",
        f"等等幫{b}查一下訂單狀態，有消息馬上回覆您唷",
    ])


# ── 轉帳確認 ──────────────────────────────────────
def payment_ack() -> str:
    b = boss()
    return random.choice([
        f"等等幫{b}確認一下嘿",
        f"好的，等等幫{b}確認一下嘿",
        f"收到唷，等等幫{b}確認一下嘿",
        f"等等幫{b}核對一下嘿",
    ])


# ── 價格查詢回覆 ──────────────────────────────────
def price_reply(name: str, price: float, unit: str) -> str:
    b = boss()
    price_str = f"{int(price)}" if price == int(price) else f"{price:.1f}"
    return random.choice([
        f"「{name}」售價是 ${price_str} / {unit}{suffix_light()}",
        f"{b}，「{name}」一{unit} ${price_str} 哦",
        f"「{name}」$  {price_str} 一{unit}{suffix_light()} {b}需要的話告訴我哦",
    ])


# ── 地址選擇（多地址客戶下單） ────────────────────
def ask_address_selection(codes: list[dict]) -> str:
    """客戶有多個送貨地址時，詢問這次要送到哪裡"""
    b = boss()
    lines = [f"請問{b}這次要送到哪個地址呢？"]
    for i, c in enumerate(codes, 1):
        name = (c.get("cust_name") or "").strip()
        addr = (c.get("address_label") or "").strip()
        if name and addr:
            label = f"{name}　{addr}"
        elif name:
            label = name
        elif addr:
            label = addr
        else:
            label = c["ecount_cust_cd"]
        lines.append(f"{i}. {label}")
    lines.append("（請回覆數字選擇哦~~）")
    return "\n".join(lines)


# ── 詢問客戶聯絡資料（代碼空白且缺資料時） ──────────
def ask_contact_info() -> str:
    """客戶代碼空白且缺少姓名或手機時，詢問客戶提供資料"""
    b = boss()
    return random.choice([
        f"不好意思{suffix_light()} 幫{b}建立資料需要一點資訊\n"
        f"請問{b}的姓名和手機號碼是？\n（例如：王小明 0912345678）",

        f"為了幫{b}建立訂單，可以告訴我您的姓名和手機嗎？\n"
        f"（格式：姓名 手機，例如：王小明 0912345678）",

        f"請問{b}的大名和手機號碼{suffix_light()} 讓我幫您建立資料哦\n"
        f"（例如：王小明 0912345678）",
    ])


# ── 訂單建立失敗 ──────────────────────────────────
def order_failed(name: str) -> str:
    return (
        f"不好意思{suffix_light()} 「{name}」訂單建立時遇到問題，"
        f"我通知真人客服幫您處理，請稍候哦"
    )


# ── 砍價婉拒 ──────────────────────────────────────────
def bargaining_reply() -> str:
    b = boss()
    return random.choice([
        f"不好意思{suffix_light()} 我們的售價都是固定的，沒有辦法再優惠囉\n感謝{b}的理解☺️",
        f"感謝{b}！我們都是統一定價，沒有折扣的空間嘿\n有需要的話歡迎下單喔",
        f"不好意思，售價都是固定的，無法再優惠了{suffix_light()} 謝謝{b}的諒解哦",
    ])


# ── 規格/介紹詢問（轉真人） ────────────────────────────
def spec_escalate() -> str:
    b = boss()
    return random.choice([
        f"我幫{b}問一下，稍等一下{suffix_light()}",
        f"等我查一下資料，等等回{b}哦",
        f"這個讓我問看看，確認後馬上回覆{b}{suffix_light()}",
    ])


# ── 顏色/款式變體詢問（轉真人） ─────────────────────────
def spec_color_escalate() -> str:
    b = boss()
    return random.choice([
        f"{b}，顏色的部分我幫您問一下唷，稍等嘿",
        f"顏色款式這邊我幫{b}確認一下，稍等一下唷",
        f"我去幫{b}查有沒有這個顏色，稍等嘿 🙏",
        f"顏色的部分讓我確認看看{suffix_light()} 稍等一下哦",
    ])


# ── 退換貨已記錄 ──────────────────────────────────────
def return_ack() -> str:
    return random.choice([
        "收到！現在忙碌中，等等回覆您唷",
        "好的！現在忙碌中，等等回覆您哦",
        "收到唷，現在忙碌中，等等回覆您嘿",
    ])


# ── 地址更改已記錄 ────────────────────────────────────
def address_query() -> str:
    """回覆店家地址"""
    b = boss()
    addr = _settings.STORE_ADDRESS
    return random.choice([
        f"{b}您好！我們的地址是：\n📍 {addr}\n歡迎來店面逛逛哦",
        f"我們的店在這邊唷～\n📍 {addr}\n{b}有需要可以直接過來哦",
        f"地址給您：\n📍 {addr}\n歡迎{b}來店裡逛逛",
    ])


def address_change_ack() -> str:
    b = boss()
    return random.choice([
        f"收到{suffix_light()} 已記錄地址更改需求，客服人員會盡快確認並處理哦",
        f"好的，已幫{b}記錄地址更改申請{suffix_light()} 我們會確認後盡快回覆您",
        f"已記錄{suffix_light()} 客服會盡快幫{b}確認地址更改，請稍候哦",
    ])


# ── 投訴已記錄（安撫語氣） ────────────────────────────
def complaint_ack() -> str:
    b = boss()
    return random.choice([
        f"非常抱歉造成{b}的困擾{suffix_light()} 已記錄您的問題，\n客服人員會盡快和您聯繫處理，非常感謝您的反饋",
        f"不好意思{suffix_light()} 您的問題已記錄，我們客服會盡快聯繫{b}處理，\n造成不便深感抱歉",
        f"抱歉嘿，問題已記錄{suffix_light()} 客服人員會馬上跟進處理，\n感謝{b}告知我們，非常抱歉",
    ])


# ── 催貨安撫（非制式催問，人工跟進） ─────────────────────
def urgent_order_ack() -> str:
    b = boss()
    return random.choice([
        f"不好意思讓{b}久等了，我確認一下等等回您{suffix_light()}",
        f"抱歉嘿，讓{b}等了，我問一下幫您確認，等等回覆您哦",
        f"不好意思{suffix_light()} {b}稍等我一下，我去確認看看再跟您說哦",
    ])


# ── 複合詢問引導（找不到編號時才用） ──────────────────────
def multi_product_guide() -> str:
    b = boss()
    return random.choice([
        f"請問{b}要查多款的話，可以一款一款問我哦{suffix_light()} 請先告訴我第一款的品名？",
        f"{b}好，方便的話一次查一款嘿，請先告訴我第一款要查哪個？",
        f"我一次幫{b}查一款，請先說第一款是哪個{suffix_light()}",
    ])


# ── 複合庫存查詢結果 ──────────────────────────────────────
def multi_stock_reply(results: list[dict]) -> str:
    """
    同時回覆多款產品庫存狀態。
    results 每筆：{"name": str, "code": str, "in_stock": bool|None, "low": bool}
    """
    lines = []
    for r in results:
        if r["in_stock"] is None:
            lines.append(f"・「{r['name']}」找不到這款唷")
        elif r["in_stock"]:
            if r["low"]:
                lines.append(f"・「{r['name']}」有嘿，不過庫存不多囉")
            else:
                lines.append(f"・「{r['name']}」有唷！")
        else:
            lines.append(f"・「{r['name']}」目前沒貨唷")
    body = "\n".join(lines)
    return f"{body}\n\n有需要的跟我說一聲嘿，幫您安排！"


# ── 規格資訊回覆（有本地規格庫時直接回答） ───────────────
def spec_info_reply(
    name: str, code: str, size: str, weight: str,
    machine: list, price: str,
) -> str:
    """格式化規格資訊回覆"""
    machines = "、".join(machine) if machine else ""
    lines = [f"📦 {name}（{code}）"]
    if size:
        lines.append(f"📐 尺寸：{size}")
    if weight:
        lines.append(f"⚖️ 重量：{weight}")
    if machines:
        lines.append(f"🎮 適用台型：{machines}")
    if price:
        lines.append(f"💰 售價：{price}")
    lines.append(f"\n{boss()}還有需要的話隨時說{suffix_light()}")
    return "\n".join(lines)


# ── 圖片識別成功（有貨）：回覆產品資訊，不顯示庫存數量 ─────────
def image_product_found(
    code: str, name: str,
    spec: dict | None,
) -> str:
    """
    客戶傳圖、比對成功且有貨時的回覆。
    - 不顯示庫存數量（只說「有喔」）
    - 缺貨情形由 service.py 改走 out_of_stock_ask_qty 流程
    """
    b = boss()
    lines = [f"這款是「{name}」{suffix_light()}"]
    lines.append(random.choice(["有貨嘿", "有唷", "有的嘿", "有喔"]))

    # 規格資訊（若有）
    if spec:
        if spec.get("size"):
            lines.append(f"📐 尺寸：{spec['size']}")
        if spec.get("weight"):
            lines.append(f"⚖️ 重量：{spec['weight']}")
        if spec.get("machine"):
            lines.append(f"🎮 適用台型：{'、'.join(spec['machine'])}")
        if spec.get("price"):
            lines.append(f"💰 售價：{spec['price']}")

    lines.append(f"\n{b}需要幾個呢？")
    return "\n".join(lines)


# ── 圖片識別失敗 ──────────────────────────────────────
def image_not_recognized() -> str:
    return "稍等一下我幫你問問看喔！"


# ── 圖片下載失敗 ──────────────────────────────────────
def image_download_failed() -> str:
    return f"不好意思，圖片讀取失敗{suffix_light()} 可以重新傳一次嗎？"


# ── 到貨通知登記 ────────────────────────────────────────
def notify_request_ack(prod_name: str) -> str:
    """缺貨時登記到貨通知"""
    b = boss()
    return random.choice([
        f"好的！「{prod_name}」到貨後馬上通知{b}唷",
        f"收到！「{prod_name}」有貨了第一時間通知{b}哦",
        f"好的，「{prod_name}」到貨會通知{b}的，請稍候唷",
    ])


def notify_request_in_stock(prod_name: str) -> str:
    """客戶要登記通知但查到現在有貨"""
    b = boss()
    return random.choice([
        f"「{prod_name}」現在有貨唷！{b}要下單的話告訴我幾個哦",
        f"好消息！「{prod_name}」目前有現貨，{b}要幾個呢？",
        f"欸！「{prod_name}」現在有貨嘿，{b}需要幾個呢？",
    ])


def notify_ask_product() -> str:
    """詢問要登記哪個產品"""
    b = boss()
    return random.choice([
        f"請問{b}要登記哪個產品的到貨通知呢？（輸入品名或編號）",
        f"哦哦，請告訴我產品編號或品名{suffix_light()} 到貨馬上通知您",
        f"好的！請問是哪款產品{suffix_light()} 告訴我編號或品名哦",
    ])


# ── 群組預設地址確認 ─────────────────────────────────────
def ask_group_address_confirm(addr_label: str) -> str:
    """群組訂單：詢問是否送到預設地址（而非列出全部選項）"""
    b = boss()
    return random.choice([
        f"請問{b}這次是送到「{addr_label}」嗎？（是/否）",
        f"{b}這次送「{addr_label}」對嗎？",
        f"確認一下，送到「{addr_label}」嗎？（是/否）",
    ])


# ── 離峰時段自動回覆 ───────────────────────────────────
def quiet_hours_ack() -> str:
    return "上班時幫您處理唷"
