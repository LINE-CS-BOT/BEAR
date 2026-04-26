"""
新場景處理模組

統一處理以下場景：
  - 非制式催貨（urgent_order）  → 記錄 DB（納入每小時待處理清單）
  - 砍價（bargaining）          → 禮貌婉拒固定售價
  - 規格詢問（spec）            → 先查本地規格庫，有資料直接回；沒有才記錄 DB 等真人
  - 退換貨（return）            → 記錄 DB（納入每小時待處理清單）
  - 複合詢問（multi）           → 引導一次問一款
  - 地址更改（address）         → 記錄 DB（納入每小時待處理清單）
  - 投訴（complaint）           → 記錄 DB（納入每小時待處理清單）
  - 圖片詢問（image）           → 感知雜湊比對圖片庫 → 識別產品 → 查庫存+規格
  - 到貨通知登記（notify）      → 有貨時告知可下單；無貨時登記 notify_store
"""

import re

from linebot.v3.messaging import MessagingApi

from handlers import tone
from storage.customers import customer_store
from storage.issues import issue_store
import storage.specs as spec_store


# ── 非制式催貨 ───────────────────────────────────────
def handle_urgent_order(user_id: str, text: str, line_api: MessagingApi) -> str:
    """記錄催貨請求，納入每小時待處理清單"""
    issue_id = issue_store.add(user_id, "urgent_order", text)
    cust_info = customer_store.get_by_line_id(user_id)
    cust_label = (cust_info.get("display_name") or user_id) if cust_info else user_id

    print(f"[service] 催貨 #I{issue_id} | {cust_label}: {text!r}")
    return tone.urgent_order_ack()


# ── 砍價 ─────────────────────────────────────────────
def handle_bargaining(user_id: str, text: str) -> str:
    """固定售價，禮貌婉拒，不轉真人"""
    return tone.bargaining_reply()


# ── 規格/介紹 ─────────────────────────────────────────
def handle_spec(user_id: str, text: str, line_api: MessagingApi) -> str:
    """
    規格詢問處理：
    1. 先從文字擷取產品編號（如 T1202、Z3300）
    2. 再嘗試 Ecount 品名模糊比對
    3. 查本地 specs.json：有資料 → 直接回覆；無資料 → 轉真人
    """
    from services.ecount import ecount_client

    # 1. 先嘗試從文字找產品編號（字母+4位數字）
    code = None
    m = re.search(r"[A-Z]\d{4,}", text.upper())
    if m:
        code = m.group(0)

    # 2. 若無明確編號，用 Ecount 品名模糊比對
    if not code:
        code = ecount_client._resolve_product_code(text)

    # 2.5 還是沒有 → 從對話上下文找最近聊過的產品
    if not code:
        try:
            from services.claude_ai import _get_chat_context
            _ctx = _get_chat_context(user_id)
            if _ctx:
                _ctx_codes = re.findall(r'[A-Z]\d{4,}', _ctx.upper())
                if _ctx_codes:
                    code = _ctx_codes[-1]  # 最近的產品
                    print(f"[spec] 從對話上下文找到產品: {code}", flush=True)
        except Exception:
            pass

    # 3. 查本地規格庫
    spec = None
    if code:
        spec = spec_store.get_by_code(code)

    # 4. 有規格資料 → 直接回覆，不轉真人
    if spec:
        print(f"[spec] 規格庫命中 {code}：{spec.get('name')}")
        return tone.spec_info_reply(
            name    = spec.get("name", code),
            code    = spec["code"],
            size    = spec.get("size", ""),
            weight  = spec.get("weight", ""),
            machine = spec.get("machine", []),
            price   = spec.get("price", ""),
        )

    # 5. 規格庫無資料 → 記錄 DB，納入每小時待處理清單
    issue_store.add(user_id, "spec_query", text)
    return tone.spec_escalate()


# ── 台型查詢：「巨無霸有什麼」→ 有庫存的產品清單 + PO文 + 圖片 ──────
# 支援的台型關鍵字 → 正規化名稱
_MACHINE_KEYWORDS = {
    "巨無霸": "巨無霸台",
    "中巨":   "中巨台",
    "K霸":    "K霸台",
    "k霸":    "K霸台",
    "小K霸":  "小K霸台",
    "小k霸":  "小K霸台",
    "標準":   "標準台",
    "迷你":   "迷你台",
    "超K":    "超K台",
    "超k":    "超K台",
}

# 主動詢問語：必須含才算「想看該台型有什麼」，否則是順便提到台型不該推送
_MACHINE_QUERY_KEYWORDS = [
    "有什麼", "有哪些", "推薦", "可以選", "想看", "看看",
    "介紹", "哪些", "什麼好", "找一下", "看一下",
]

# 推薦觸發語：「需要 X 貨」「我要 X 的」「給我 X」這類請求語
_MACHINE_RECOMMEND_PREFIXES = [
    "需要", "想要", "我要", "給我", "找", "幫我找", "有沒有",
]


def extract_machine_type_loose(text: str) -> str | None:
    """從訊息抓台型（不要求「主動詢問語」）。用於主動推薦場景。
    例：「需要中巨貨」「我要 K 霸的」「巨無霸 300 內」都能命中。
    """
    if not text:
        return None
    # 按關鍵字長度降冪，避免「小K霸」被「K霸」攔截
    for kw in sorted(_MACHINE_KEYWORDS.keys(), key=len, reverse=True):
        if kw in text:
            return _MACHINE_KEYWORDS[kw]
    return None


_BUDGET_RE = re.compile(
    r"(\d{2,4})\s*(?:元|塊|圓)?\s*(?:內|以內|以下|左右|底|以)",
)


def extract_budget(text: str) -> int | None:
    """從訊息抓預算數字。支援：「300內」「200元以下」「500以內」。"""
    if not text:
        return None
    m = _BUDGET_RE.search(text)
    if not m:
        return None
    try:
        n = int(m.group(1))
    except ValueError:
        return None
    # 排除過小（可能是時間/數量）或過大（不合理）
    if n < 30 or n > 5000:
        return None
    return n


def _spec_price_int(spec: dict) -> int | None:
    """從 spec 的 price 欄位抓出整數價格"""
    p = spec.get("price", "")
    if not p:
        return None
    m = re.search(r"\d+", str(p))
    return int(m.group()) if m else None


# 記錄每個 user 上次的 machine_recommend：用於追問預算的 follow-up
# {user_id: (machine_type, ts_epoch)}；TTL 5 分鐘
_RECOMMEND_FOLLOWUP_TTL_SEC = 300
_recommend_followup: dict[str, tuple[str, float]] = {}


def remember_machine_followup(user_id: str, machine_type: str) -> None:
    import time as _t
    _recommend_followup[user_id] = (machine_type, _t.time())


def try_machine_budget_followup(user_id: str, text: str, line_api: MessagingApi) -> str | None:
    """客戶剛問過台型（5 分內），這次只傳預算（如「300內」）→ 用同台型+新預算重列"""
    import time as _t
    if not text:
        return None
    last = _recommend_followup.get(user_id)
    if not last:
        return None
    machine_type, ts = last
    if _t.time() - ts > _RECOMMEND_FOLLOWUP_TTL_SEC:
        _recommend_followup.pop(user_id, None)
        return None
    # 必須 text 只是預算（不含台型也不含其他內容）
    if extract_machine_type_loose(text):
        return None  # 已含台型，正常 dispatch 處理
    budget = extract_budget(text)
    if not budget:
        return None
    # 額外保險：text 不能太長（避免完整句子被誤認）
    if len(text.strip()) > 12:
        return None
    return handle_machine_recommend(user_id, machine_type, budget, line_api)


def handle_machine_recommend(
    user_id: str, machine_type: str, budget: int | None, line_api: MessagingApi
) -> str | None:
    """
    主動推薦：給定台型 + 可選預算 → 列出有現貨的產品（top 10）。
    格式：貨號 + 品名 + 價格（不顯示庫存數，符合 feedback_no_stock_qty）。
    """
    from services.ecount import ecount_client

    all_specs = spec_store.get_by_machine(machine_type)
    if not all_specs:
        return f"目前沒有{machine_type}的產品資料唷～"

    # 過濾：有庫存 + 預算內
    candidates: list[tuple[dict, int]] = []
    for sp in all_specs:
        code = sp.get("code", "")
        item = ecount_client.lookup(code)
        qty = item.get("qty", 0) if item else 0
        if qty <= 0:
            continue
        price = _spec_price_int(sp)
        if budget is not None:
            if price is None or price > budget:
                continue
        candidates.append((sp, qty))

    if not candidates:
        if budget:
            remember_machine_followup(user_id, machine_type)  # 仍記住，方便客戶換預算追問
            return f"目前{machine_type} {budget}元以內的現貨剛好都沒有耶～有別的預算嗎？"
        return f"目前{machine_type}剛好都缺貨耶～有別的台型嗎？"

    # 排序：價格低 → 高（讓便宜的優先），同價按 qty 降冪當 tiebreak
    def _sort_key(t):
        sp, qty = t
        return (_spec_price_int(sp) or 99999, -qty)
    candidates.sort(key=_sort_key)

    total = len(candidates)
    show = candidates[:10]

    title = f"🎯 {machine_type}" + (f" / {budget}元內" if budget else "") + " 現貨："
    lines = [title]
    for sp, _qty in show:
        code = sp.get("code", "")
        name = sp.get("name", "")[:20]
        price = _spec_price_int(sp)
        price_s = f"  ${price}" if price else ""
        lines.append(f"• {code} {name}{price_s}")
    if total > 10:
        lines.append(f"（共 {total} 件，先列前 10 件）")
    lines.append("")
    lines.append("需要哪款跟小編說貨號就好🙏")
    remember_machine_followup(user_id, machine_type)
    return "\n".join(lines)


def detect_machine_query(text: str) -> str | None:
    """從訊息中偵測台型查詢，回傳正規化台型名稱（或 None）。
    要同時含「主動詢問語」+「台型關鍵字」才算，避免「K霸先取消」這種被誤觸推送。
    iteration 時按 key 長度降冪排，避免「小K霸」被「K霸」攔截。"""
    if not any(qkw in text for qkw in _MACHINE_QUERY_KEYWORDS):
        return None
    for kw in sorted(_MACHINE_KEYWORDS.keys(), key=len, reverse=True):
        if kw in text:
            return _MACHINE_KEYWORDS[kw]
    return None


def handle_machine_query(
    user_id: str, machine_type: str, line_api: MessagingApi
) -> str:
    """
    查詢指定台型的有庫存產品。
    - 回傳標題文字（讓 caller 用 reply_message 送出，不佔額度）
    - 每個產品的 PO文 + 圖片 在背景執行緒用 push_message 依序送出
    """
    import threading
    from services.ecount import ecount_client
    from config import settings as _cfg
    from handlers.internal import _format_po, _match_product_media_files, _get_media_dir, _build_media_messages, _push_messages_chunked
    from linebot.v3.messaging import PushMessageRequest, TextMessage

    # 查規格庫
    all_specs = spec_store.get_by_machine(machine_type)
    if not all_specs:
        return f"目前沒有{machine_type}的產品資料唷，有需要可以問一下喔～"

    # 過濾有庫存的產品
    in_stock = []
    for sp in all_specs:
        code = sp.get("code", "")
        item = ecount_client.lookup(code)
        qty = item.get("qty", 0) if item else 0
        if qty and qty > 0:
            in_stock.append((sp, qty))

    if not in_stock:
        return f"目前{machine_type}的產品都暫時缺貨唷，到貨會再通知您嘿～"

    # 按庫存量排序，取前 10 項
    in_stock.sort(key=lambda x: -x[1])
    total_count = len(in_stock)
    in_stock = in_stock[:2]

    # 背景推送每個產品的 PO文 + 圖片
    def _push_products():
        base_url = "https://xmnline.duckdns.org/product-photo"
        for sp, qty in in_stock:
            code = sp.get("code", "")
            po_text = _format_po(code) or (
                f"{sp.get('name', code)}\n"
                f"編號：{code}\n"
                f"尺寸：{sp.get('size','')}\n"
                f"重量：{sp.get('weight','')}\n"
                f"適用：{machine_type}\n"
                f"價格：{sp.get('price','')}\n"
                f"庫存：{qty} 個"
            )
            media_dir = _get_media_dir()
            media_files = _match_product_media_files(code, media_dir) if base_url and media_dir else []
            media_msgs  = _build_media_messages(code, media_files, base_url) if media_files else []
            text_msg    = TextMessage(text=po_text)
            try:
                _push_messages_chunked(line_api, user_id, text_msg, media_msgs)
            except Exception as e:
                print(f"[machine-query] 推送 {code} 失敗: {e}", flush=True)

    threading.Thread(target=_push_products, daemon=True).start()

    return None  # 不回覆標題文字，直接推 PO文+圖


# ── 圖片詢問 ─────────────────────────────────────────
def handle_image_product(user_id: str, message_id: str, line_api: MessagingApi) -> str:
    """
    處理客戶傳來的產品圖片：
    1. 下載圖片
    2. 感知雜湊比對圖片庫
    3. 識別成功 → 查庫存 + 規格 → 回覆
    4. 識別失敗 → 通知真人 + 安撫客戶
    """
    from services.vision import download_image, identify_product, is_transfer_screenshot
    from services.ecount import ecount_client
    from handlers.hours import _is_open_now
    from datetime import datetime
    import pytz
    from config import settings as _settings

    # 下載圖片
    image_bytes = download_image(message_id)
    if not image_bytes:
        return tone.image_download_failed()

    # ── 轉帳截圖偵測（優先判斷）──────────────────────────
    if is_transfer_screenshot(image_bytes):
        print(f"[image] 偵測到轉帳截圖 → user={user_id}")
        issue_store.add(user_id, "payment_screenshot", "客戶傳了匯款截圖，待確認")
        now = datetime.now(pytz.timezone(_settings.BUSINESS_TZ))
        if _is_open_now(now):
            return "等等有空查詢確認喔～"
        else:
            return "上班時確認嘿"

    # ── 識別順序：① pHash 高可信 → ② OCR → ③ pHash 弱命中備援 ──
    from services.vision import ocr_extract_candidates, identify_product_weak
    prod_code = identify_product(image_bytes)   # pHash diff ≤ 10

    if not prod_code:
        # ② OCR 萃取貨號/品名候選詞，逐一比對 Ecount（優先於弱 pHash）
        # 只嘗試「貨號格式」(字母+數字) 或中文詞，跳過純英文短詞（IN/LL 等 OCR 雜訊）
        import re as _re
        _CODE_OR_ZH = _re.compile(r'(?:[A-Za-z]\d{2,}|[\u4e00-\u9fff]{2,})')
        for candidate in ocr_extract_candidates(image_bytes):
            if not _CODE_OR_ZH.search(candidate):
                continue  # 跳過純英文短詞
            matched = ecount_client._resolve_product_code(candidate)
            if matched:
                prod_code = matched
                print(f"[image] OCR 比對成功 → {prod_code}（候選詞：{candidate!r}）")
                break

    if not prod_code:
        # ③ pHash 弱命中備援（diff 10-15，僅在 OCR 也無結果時使用）
        prod_code = identify_product_weak(image_bytes)
        if prod_code:
            print(f"[image] pHash 弱命中備援 → {prod_code}")

    if not prod_code:
        # pHash + OCR 都失敗 → 嘗試 Claude 辨識
        from services.claude_ai import ask_claude_image
        _claude_reply = ask_claude_image(image_bytes, user_id=user_id)
        if _claude_reply:
            # 如果回覆裡有產品代碼，設 state + 回覆
            import re as _re_svc
            _svc_codes = _re_svc.findall(r'[A-Za-z]{1,3}-?\d{3,6}', _claude_reply)
            if _svc_codes:
                _svc_cd = _svc_codes[0].upper()
                from services.ecount import ecount_client as _ec_svc
                _svc_item = _ec_svc.get_product_cache_item(_svc_cd)
                _svc_name = (_svc_item.get("name") if _svc_item else None) or _svc_cd
                # 從 Claude 回覆抓「× N 台/個/組/箱/隻/盒」— 客戶手寫在圖上的數量
                _qty_m = _re_svc.search(
                    r'[×xX]\s*(\d+)\s*(?:個|台|組|箱|隻|盒|顆|件)', _claude_reply,
                )
                if _qty_m:
                    _svc_qty = int(_qty_m.group(1))
                    from storage import cart as _cart_svc
                    from storage.state import state_manager as _sm_svc
                    _sm_svc.clear(user_id)  # 不設 awaiting_quantity，直接建單
                    _new_cart = _cart_svc.add_item(user_id, _svc_cd, _svc_name, _svc_qty)
                    print(f"[claude-ai] 圖片辨識含數量 → 直接加購物車: {_svc_cd} × {_svc_qty}", flush=True)
                    return tone.cart_item_added(_new_cart)
                # 沒抓到數量 → 維持原本 awaiting_quantity 等客戶補
                from storage.state import state_manager as _sm_svc
                _sm_svc.set(user_id, {
                    "action":    "awaiting_quantity",
                    "prod_cd":   _svc_cd,
                    "prod_name": _svc_name,
                })
                print(f"[claude-ai] 純圖片辨識後設 awaiting_quantity: {_svc_cd}", flush=True)
                # chat_history 由 caller 的 _send_reply 在送出成功後自動記
                return _claude_reply
            # Claude 回覆沒貨號 → 靜默，記待處理
            print(f"[claude-ai] 純圖片 Claude 回覆無貨號，靜默", flush=True)
        # Claude 也失敗 → 靜默記錄，納入待處理清單
        issue_store.add(user_id, "image_query", "（傳來一張圖片，無法辨識）")
        print(f"[image] 辨識失敗 → 靜默進待處理 user={user_id[:10]}...")
        return None

    # 識別成功 → 查庫存
    from storage.state import state_manager
    result = ecount_client.lookup(prod_code)
    spec   = spec_store.get_by_code(prod_code)

    # 取得產品名稱（Ecount 快取 > 規格庫 > 用編號）
    name = (
        (result.get("name") if result else None)
        or (spec.get("name") if spec else None)
    ) or prod_code

    qty = result.get("qty") if result else None

    print(f"[image] 識別成功 → {prod_code} 庫存={qty}")

    if qty and qty > 0:
        # ── 有貨：等客戶回覆數量，走購物車流程 ──────────────
        state_manager.set(user_id, {
            "action":     "awaiting_quantity",
            "prod_cd":    prod_code,
            "prod_name":  name,
            "from_image": True,   # 圖片識別下單，回數量後需確認
        })
        return tone.image_product_found(code=prod_code, name=name, spec=spec)
    else:
        # ── 沒貨：判斷預購 or 一般缺貨，統一走購物車流程 ────
        from handlers.inventory import _check_preorder
        state_manager.set(user_id, {
            "action":     "awaiting_quantity",
            "prod_cd":    prod_code,
            "prod_name":  name,
            "from_image": True,
        })
        if _check_preorder(prod_code):
            return tone.preorder_ask_qty(name)
        else:
            return tone.out_of_stock_ask_qty(name)


# ── 退換貨 ───────────────────────────────────────────
def handle_return(user_id: str, text: str, line_api: MessagingApi) -> str:
    """記錄退換貨請求，納入每小時待處理清單"""
    issue_id = issue_store.add(user_id, "return", text)
    cust_info = customer_store.get_by_line_id(user_id)
    cust_label = (cust_info.get("display_name") or user_id) if cust_info else user_id

    print(f"[service] 退換貨 #I{issue_id} | {cust_label}: {text!r}")
    return tone.return_ack()


# ── 地址更改 ─────────────────────────────────────────
def handle_address_change(user_id: str, text: str, line_api: MessagingApi) -> str:
    """記錄地址更改請求，納入每小時待處理清單"""
    issue_id = issue_store.add(user_id, "address_change", text)
    cust_info = customer_store.get_by_line_id(user_id)
    cust_label = (cust_info.get("display_name") or user_id) if cust_info else user_id

    print(f"[service] 地址更改 #I{issue_id} | {cust_label}: {text!r}")
    return tone.address_change_ack()


# ── 投訴 ─────────────────────────────────────────────
def handle_complaint(user_id: str, text: str, line_api: MessagingApi) -> str:
    """記錄投訴，納入每小時待處理清單"""
    issue_id = issue_store.add(user_id, "complaint", text)
    cust_info = customer_store.get_by_line_id(user_id)
    cust_label = (cust_info.get("display_name") or user_id) if cust_info else user_id

    print(f"[service] 投訴 #I{issue_id} | {cust_label}: {text!r}")
    return tone.complaint_ack()


# ── 複合詢問 ─────────────────────────────────────────
def handle_multi_product(user_id: str, text: str) -> str:
    """引導客戶一次問一款"""
    return tone.multi_product_guide()


# ── 到貨通知登記 ─────────────────────────────────────
def handle_notify_request(user_id: str, text: str, line_api: MessagingApi) -> str:
    """
    客戶主動要求「有貨通知我」：
    1. 從訊息提取產品 → 查庫存
    2. 有貨 → 告知現在有貨，進入 awaiting_quantity
    3. 無貨 → 登記 notify_store
    4. 找不到產品 → 詢問品名（awaiting_notify_product）
    """
    from services.ecount import ecount_client
    from storage.notify import notify_store
    from storage.state import state_manager

    # Step 1：嘗試從文字提取產品碼（字母+數字）
    prod_code = None
    prod_name = None
    qty = None

    m = re.search(r"[A-Za-z]\d{3,}", text.upper())
    if m:
        result = ecount_client.lookup(m.group(0))
        if result:
            prod_code = result["code"]
            prod_name = result["name"] or prod_code
            qty = result.get("qty", 0)

    # Step 2：若無明確編號，用品名模糊比對
    if not prod_code:
        resolved = ecount_client._resolve_product_code(text)
        if resolved:
            result = ecount_client.lookup(resolved)
            if result:
                prod_code = result["code"]
                prod_name = result["name"] or prod_code
                qty = result.get("qty", 0)

    # Step 3：找不到產品 → 詢問
    if not prod_code:
        state_manager.set(user_id, {"action": "awaiting_notify_product"})
        return tone.notify_ask_product()

    # Step 4：有貨 → 告知並進入下單流程
    if qty and qty > 0:
        state_manager.set(user_id, {
            "action":    "awaiting_quantity",
            "prod_cd":   prod_code,
            "prod_name": prod_name,
        })
        return tone.notify_request_in_stock(prod_name)

    # Step 5：無貨 → 登記到貨通知
    notify_store.add(user_id, prod_code, prod_name, 1)
    print(f"[notify] 登記到貨通知：{prod_name}（{prod_code}）for {user_id[:10]}...")
    return tone.notify_request_ack(prod_name)


