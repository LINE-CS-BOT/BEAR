import re
from linebot.v3.messaging import MessagingApi, PushMessageRequest, TextMessage

from config import settings
from services.ecount import ecount_client
from storage.pending import pending_store
from storage.state import state_manager
from handlers import tone


def _check_preorder(prod_cd: str) -> bool:
    """檢查 PO文 或品名是否含「預購」，判斷是否為預購商品"""
    from storage import specs as spec_store
    spec = spec_store.get_by_code(prod_cd)
    if spec:
        name = spec.get("name", "")
        if "預購" in name:
            return True
    # 也檢查 Ecount 品名
    cache = ecount_client.get_product_cache_item(prod_cd)
    if cache and "預購" in (cache.get("name") or ""):
        return True
    # 檢查原始 PO文
    try:
        from handlers.internal import _get_raw_po_block
        raw = _get_raw_po_block(prod_cd)
        if raw and "預購" in raw:
            return True
    except Exception:
        pass
    return False


def handle_inventory(user_id: str, text: str, line_api: MessagingApi) -> str:
    """處理庫存查詢入口"""
    # 複合詢問（AA 和 BB 都有嗎）→ 嘗試同時查詢
    if _is_multi_product(text):
        codes = _extract_all_codes(text)
        if codes:
            return _query_multi_products(codes)
        return tone.multi_product_guide()

    # 顏色/款式詢問（產品編號 + 顏色詞）→ 轉真人確認
    if _has_color_query(text):
        from storage.issues import issue_store
        issue_store.add(user_id, "spec_query", text)
        return tone.spec_color_escalate()

    product = _extract_product(text)

    if not product:
        # 沒有提到產品名稱，進入多輪對話等待輸入
        state_manager.set(user_id, {"action": "awaiting_product"})
        return tone.ask_product()

    return query_product(user_id, product, line_api)


def _is_multi_product(text: str) -> bool:
    """偵測複合詢問（含兩款商品的庫存問法）"""
    has_connector = any(kw in text for kw in ["和", "跟", "還有", "以及"])
    has_both = any(kw in text for kw in ["都", "各", "分別"])
    return has_connector and has_both


_COLOR_WORDS = [
    "紅色", "藍色", "黑色", "白色", "綠色", "黃色",
    "粉色", "粉紅", "灰色", "橘色", "紫色", "咖啡色",
    "透明", "銀色", "金色", "深藍", "淺藍", "深綠", "淺綠",
]


def _has_color_query(text: str) -> bool:
    """偵測「產品編號 + 顏色詞」的組合詢問（顏色變體）"""
    has_code = bool(re.search(r"[A-Za-z]\d{3,}(?:-\d+)?", text))
    has_color = any(c in text for c in _COLOR_WORDS)
    return has_code and has_color


def _extract_all_codes(text: str) -> list[str]:
    """從複合詢問中提取所有產品編號（最多 3 款）"""
    found = re.findall(r"[A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?", text.upper())
    seen, result = set(), []
    for code in found:
        if code not in seen:
            seen.add(code)
            result.append(code)
    return result[:3]


def _query_multi_products(codes: list[str]) -> str:
    """同時查詢多款產品庫存，回傳彙整結果（不進入下單狀態）"""
    results = []
    for code in codes:
        item = ecount_client.lookup(code)
        if item:
            results.append({
                "name":     item["name"] or code,
                "code":     item["code"],
                "in_stock": item["qty"] > 0,
                "low":      0 < item["qty"] <= 5,
            })
        else:
            results.append({
                "name":     code,
                "code":     code,
                "in_stock": None,   # 找不到此編號
                "low":      False,
            })
    return tone.multi_stock_reply(results)


def query_product(user_id: str, product: str, line_api: MessagingApi = None) -> str:
    """查詢特定產品庫存並處理結果（多筆匹配只列有貨款式）"""

    all_codes = ecount_client.search_products_by_name(product)

    if not all_codes:
        pending_store.add(user_id, product)
        return tone.product_not_found(product)

    if len(all_codes) == 1:
        return _query_single_product(user_id, all_codes[0], line_api)

    # 多筆匹配 → 先篩有貨（qty > 0），最多 5 筆
    in_stock: list[tuple[str, str]] = []
    for code in all_codes[:10]:
        item = ecount_client.lookup(code)
        if item and (item.get("qty") or 0) > 0:
            in_stock.append((code, item.get("name") or code))
        if len(in_stock) >= 5:
            break

    if not in_stock:
        # 全部沒貨 → 第一筆走缺貨調貨流程
        return _query_single_product(user_id, all_codes[0], line_api)

    if len(in_stock) == 1:
        # 剛好只有一款有貨 → 直接查
        return _query_single_product(user_id, in_stock[0][0], line_api)

    # 多款有貨 → 讓客戶選
    state_manager.set(user_id, {
        "action":     "awaiting_product_clarify",
        "keyword":    product,
        "candidates": in_stock,
    })
    return tone.ask_product_clarify(product, in_stock)


def _find_case_variant(prod_cd: str) -> dict | None:
    """查找箱裝版本：U0192-1(個) → U0192(箱)，或 Z3432 → Z3432-1(箱)
    注意：箱裝和個裝共用同一個庫存代表號，箱裝本身可能沒有獨立庫存紀錄。
    優先用 product cache（本地，不打 API），避免 session 過期導致找不到。
    """
    code = prod_cd.upper()
    candidates = []
    if "-" in code:
        candidates.append(code.rsplit("-", 1)[0])  # 去後綴
    candidates.append(code + "-1")                   # 加 -1

    for c in candidates:
        # 先查本地快取（不需 API）
        cache = ecount_client.get_product_cache_item(c)
        if cache and ("箱" in cache.get("name", "") or "條" in cache.get("name", "")):
            return {"code": cache["code"], "name": cache["name"], "price": cache.get("price")}
    return None


def _query_single_product(user_id: str, prod_cd: str, line_api: MessagingApi = None) -> str:
    """以確定的 PROD_CD 查庫存並回覆"""
    item = ecount_client.lookup(prod_cd)

    if item is None:
        return tone.product_not_found(prod_cd)

    name = item["name"] or prod_cd
    qty  = item["qty"]

    # 箱裝推薦（箱裝和個裝共用庫存，從品項快取拿價格）
    case_tip = ""
    case_item = _find_case_variant(prod_cd)
    if case_item and case_item["code"].upper() != prod_cd.upper():
        # 價格優先從 product cache 拿（出庫單價），lookup 的 available.json 可能沒有箱裝
        case_cache = ecount_client.get_product_cache_item(case_item["code"])
        case_price = (case_cache or {}).get("price") or case_item.get("price")
        case_unit = (case_cache or {}).get("unit") or "箱"
        if case_price:
            case_tip = f"\n\n💡 整{case_unit}購買更划算唷！\n📦 {case_item['name']} ${int(float(case_price))}/{case_unit}"

    if qty > 0:
        state_manager.set(user_id, {
            "action":    "awaiting_quantity",
            "prod_cd":   item["code"],
            "prod_name": name,
        })
        if qty <= 5:
            return tone.in_stock_low(name) + case_tip
        return tone.in_stock(name) + case_tip
    else:
        # 預購判斷：PO文含「預購」→ 走預購流程（直接問數量下單）
        _is_preorder = _check_preorder(item["code"])
        if _is_preorder:
            state_manager.set(user_id, {
                "action":    "awaiting_quantity",
                "prod_cd":   item["code"],
                "prod_name": name,
            })
            return tone.preorder_ask_qty(name)
        state_manager.set(user_id, {
            "action":    "awaiting_restock_qty",
            "prod_name": name,
            "prod_cd":   item["code"],
        })
        return tone.out_of_stock_ask_qty(name)


def notify_hq_restock(prod_name: str, qty: int, line_api: MessagingApi | None) -> None:
    """通知總公司群組詢問調貨及到貨時間（公開函式，供 main.py 呼叫）"""
    if not line_api or not settings.LINE_GROUP_ID_HQ:
        print(f"[總公司通知] 未設定 LINE_GROUP_ID_HQ，跳過（{prod_name} × {qty}個）")
        return

    msg = f"請問一下\n📦 {prod_name} × {qty} 個\n是否有數量可以調貨? 如果叫貨需要多久時間?"
    try:
        line_api.push_message(
            PushMessageRequest(
                to=settings.LINE_GROUP_ID_HQ,
                messages=[TextMessage(text=msg)],
            )
        )
    except Exception as e:
        print(f"[總公司通知] 推送失敗: {e}")


def _extract_product(text: str) -> str:
    """從訊息中嘗試提取產品編號或名稱（支援中英文）"""

    # Step 1：剝離前綴（問候/代稱/助詞）
    t = re.sub(
        r"^(?:請問|想問|問一下|查一下|你們|妳們|你們的|我想問|老闆|嗨|哈囉)\s*",
        "", text,
    )
    # 剝離動詞前綴（還有/有沒有 → 出現在產品名之前，可能無空格）
    t = re.sub(r"^(?:還有|有沒有)\s*", "", t)

    # Step 2：剝離後綴問句（含「還有貨嗎」「還有嗎」「還有」等「還」開頭後綴）
    t = re.sub(
        r"\s*(?:還有貨嗎|還有沒有貨|還有嗎|還有貨|有貨嗎|有沒有貨|可以訂嗎|能訂嗎|有得訂|訂購|缺貨|有嗎|有貨|能訂|可訂|有沒有|還有|庫存)\s*$",
        "", t,
    )
    t = re.sub(r"\s*嗎\s*$", "", t)
    t = t.strip()

    # 如果有剝離到東西，就用結果
    if t and t != text:
        return t

    # Step 3：純英數編號 + 問句
    m = re.search(r"([A-Za-z0-9\-_]+)\s*(?:有貨|庫存|訂購|可以訂)", text)
    if m:
        return m.group(1).strip()

    # Step 4：最後手段 — 暴力刪除所有問句關鍵字
    cleaned = re.sub(
        r"(還有貨嗎|還有沒有貨|還有嗎|還有貨|有貨嗎|有沒有貨|可以訂嗎|能訂嗎|有得訂|訂購|缺貨|有嗎|有貨|能訂|可訂|請問|想問|問一下|查一下|有沒有|還有|庫存|你們|妳們|嗎)",
        "", text,
    ).strip()
    return cleaned
