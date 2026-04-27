import re
from pathlib import Path
from linebot.v3.messaging import MessagingApi, PushMessageRequest, TextMessage

from config import settings
from services.ecount import ecount_client
from storage.pending import pending_store
from storage.state import state_manager
from handlers import tone


_PREORDER_PATH = Path(__file__).parent.parent / "data" / "preorder.json"
_preorder_cache: dict[str, dict] = {}   # code → {"name": "...", "eta": "..."}
_preorder_loaded: float = 0.0


def _extract_eta(po_block: str) -> str:
    """從 PO文 block 提取預計到貨時間"""
    # 優先：具體日期（預計4/10到貨、預計4月底到貨、預計第三季到貨）
    m = re.search(r"預計\s*(.{2,20}?到貨)", po_block)
    if m:
        return m.group(1).strip()
    # 次之：X月底前到貨
    m = re.search(r"(\d+月[中下旬底]*前?到貨)", po_block)
    if m:
        return m.group(1)
    # 再次：預購N天到貨
    m = re.search(r"預購\s*(\d+)\s*天到貨", po_block)
    if m:
        return f"{m.group(1)}天到貨"
    # 最後：14-21天
    m = re.search(r"(\d+-\d+)\s*(?:工作)?天", po_block)
    if m:
        return f"{m.group(1)}工作天"
    return ""


def _load_preorder_cache() -> dict[str, dict]:
    """從 preorder.json 載入快取"""
    global _preorder_cache, _preorder_loaded
    import time
    if time.time() - _preorder_loaded < 60 and _preorder_cache:
        return _preorder_cache
    if _PREORDER_PATH.exists():
        try:
            import json
            data = json.loads(_PREORDER_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                _preorder_cache = data
            elif isinstance(data, list):
                _preorder_cache = {c: {} for c in data}
            _preorder_loaded = time.time()
        except Exception:
            pass
    return _preorder_cache


def _check_preorder(prod_cd: str) -> bool:
    """預購 = preorder.json 快取中有此貨號（快取已排除有庫存的）"""
    return prod_cd.upper() in _load_preorder_cache()


def refresh_preorder_list() -> int:
    """
    重新掃描：PO文/品名含「預購」且庫存 <= 0 → 寫入 data/preorder.json。
    到貨（有庫存）自動篩除。保留預購日期記錄。
    回傳預購品數量。在庫存同步後 / 上架完成後呼叫。
    """
    import json
    from datetime import datetime

    # 讀取現有快取（保留歷史日期）
    existing: dict[str, dict] = {}
    if _PREORDER_PATH.exists():
        try:
            data = json.loads(_PREORDER_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                existing = data
        except Exception:
            pass

    # 收集所有 PO文含「預購」的貨號 + 品名
    candidates: dict[str, str] = {}  # code → name

    # 1. specs 品名
    specs_path = Path(__file__).parent.parent / "data" / "specs.json"
    if specs_path.exists():
        try:
            specs = json.loads(specs_path.read_text(encoding="utf-8"))
            for code, s in specs.items():
                if "預購" in s.get("name", ""):
                    candidates[code.upper()] = s.get("name", "")
        except Exception:
            pass

    # 2. Ecount 品名快取（優先用 Ecount 品名，較正式）
    ecount_names: dict[str, str] = {}
    ecount_client._ensure_product_cache()
    for item in (ecount_client._product_cache or []):
        name = item.get("name") or ""
        code = (item.get("code") or "").upper()
        if code:
            ecount_names[code] = name
            if "預購" in name:
                candidates.setdefault(code, name)

    # 3. PO文
    po_etas: dict[str, str] = {}  # code → 到貨時間描述
    po_path = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt")
    if po_path.exists():
        try:
            text = po_path.read_text(encoding="utf-8")
            blocks = re.split(r"\n{2,}", text)
            for block in blocks:
                if "預購" in block:
                    m = re.search(r"([A-Za-z]{1,3}-?\d{3,6})", block)
                    if m:
                        code = m.group(1).upper()
                        if code not in candidates:
                            first_line = block.strip().split("\n")[0][:40]
                            candidates[code] = first_line
                        # 提取到貨時間
                        eta = _extract_eta(block)
                        if eta:
                            po_etas[code] = eta
        except Exception:
            pass

    # 4. 過濾掉有庫存的（有貨 = 已到貨，不算預購）
    avail_path = Path(__file__).parent.parent / "data" / "available.json"
    avail = {}
    if avail_path.exists():
        try:
            avail = json.loads(avail_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    result: dict[str, dict] = {}
    for code in sorted(candidates):
        qty = 0
        if code in avail:
            d = avail[code]
            qty = d.get("available", 0) if isinstance(d, dict) else d
        if qty <= 0:
            eta = po_etas.get(code, "")
            # 「14-21天」「15天到貨」等純天數是通用缺貨說明，不算真正預購品
            if not eta or re.match(r"^\d+(-\d+)?(?:工作)?天(?:到貨)?$", eta):
                continue
            best_name = ecount_names.get(code) or candidates[code]
            result[code] = {
                "name": best_name,
                "eta": eta,
            }

    # 5. 寫入
    _PREORDER_PATH.parent.mkdir(parents=True, exist_ok=True)
    _PREORDER_PATH.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 更新記憶體快取
    global _preorder_cache, _preorder_loaded
    import time
    _preorder_cache = result
    _preorder_loaded = time.time()

    print(f"[preorder] 更新預購清單：{len(result)} 筆")
    return len(result)


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
    # 先抽貨號（避免客戶 quote 整段 PO 文時找不到）
    _code_m = re.search(r"[A-Za-z]\d{3,}(?:-\d+)?", product)
    if _code_m:
        _direct_code = _code_m.group(0).upper()
        _direct_item = ecount_client.lookup(_direct_code)
        if _direct_item:
            return _query_single_product(user_id, _direct_code, line_api)

    all_codes = ecount_client.search_products_by_name(product)

    # 整段搜不到 → 拆 keyword 各別搜（如「野獸國 魯斯佛 跟 洪金寶」拆成三段各搜）
    if not all_codes:
        _STOP = {"請問", "還有", "有貨", "有沒", "有嗎", "有貨嗎", "嗎", "呢",
                 "請", "問", "可以", "可不可以", "可訂", "能訂",
                 "新北", "土城", "店", "謝謝", "感謝"}
        _STRIP_PREFIXES = ("請問", "我要", "想要", "要訂", "想訂", "請")
        _STRIP_SUFFIXES = ("嗎", "呢", "沒", "唷", "喔", "哦")
        raw_tokens = [t.strip() for t in re.split(r'[\s、，,。和跟與同或還有]+', product) if t.strip()]
        tokens = []
        for tok in raw_tokens:
            for pfx in _STRIP_PREFIXES:
                if tok.startswith(pfx):
                    tok = tok[len(pfx):]
                    break
            for sfx in _STRIP_SUFFIXES:
                if tok.endswith(sfx):
                    tok = tok[:-len(sfx)]
                    break
            if len(tok) >= 2 and tok not in _STOP:
                tokens.append(tok)
        # 排序：narrower (字數多) 在前，避免「野獸國」這類太廣 token 擠壓具體名稱
        tokens.sort(key=lambda t: -len(t))
        merged: list[str] = []
        # 先預跑找 narrow tokens（hits ≤15）
        token_hits = [(tok, ecount_client.search_products_by_name(tok)) for tok in tokens]
        narrow = [(tok, hits) for tok, hits in token_hits if 0 < len(hits) <= 15]
        # 有 narrow 命中就只用 narrow 結果（如「野獸國 魯斯佛」會選「魯斯佛」1 筆）
        # 否則 fallback 到所有結果
        chosen = narrow if narrow else token_hits
        for tok, hits in chosen:
            for c in hits:
                if c not in merged:
                    merged.append(c)
        all_codes = merged

    if not all_codes:
        pending_store.add(user_id, product)
        return tone.product_not_found(product)

    if len(all_codes) == 1:
        return _query_single_product(user_id, all_codes[0], line_api)

    # 多筆匹配 → 篩有貨（qty > 0）
    in_stock: list[tuple[str, str]] = []
    for code in all_codes:
        item = ecount_client.lookup(code)
        if item and (item.get("qty") or 0) > 0:
            in_stock.append((code, item.get("name") or code))

    if not in_stock:
        # 全部沒貨 → 第一筆走缺貨調貨流程
        return _query_single_product(user_id, all_codes[0], line_api)

    if len(in_stock) == 1:
        # 剛好只有一款有貨 → 直接查
        return _query_single_product(user_id, in_stock[0][0], line_api)

    # 多款有貨 → 讓客戶選（最多顯示 20 筆）
    display = in_stock[:20]
    state_manager.set(user_id, {
        "action":     "awaiting_product_clarify",
        "keyword":    product,
        "candidates": display,
    })
    extra = f"\n\n（共 {len(in_stock)} 款有貨，顯示前 20 筆）" if len(in_stock) > 20 else ""
    return tone.ask_product_clarify(product, display) + extra


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
            "action":    "awaiting_quantity",
            "prod_name": name,
            "prod_cd":   item["code"],
        })
        return tone.out_of_stock_ask_qty(name)



def _extract_product(text: str) -> str:
    """從訊息中嘗試提取產品編號或名稱（支援中英文）"""

    # Step 0：清除標點符號
    t = re.sub(r'[，,。.！!？?、；;：:～~\s]+', ' ', text).strip()

    # Step 1：剝離前綴（問候/代稱/助詞/動詞）
    t = re.sub(
        r"^(?:我要查詢|我要查|我想查詢|我想查|我要問|我想問|請問|想問|問一下|查一下|你們|妳們|你們的|老闆|嗨|哈囉|嘿|喂)\s*",
        "", t,
    )
    # 剝離動詞前綴（還有/有沒有 → 出現在產品名之前，可能無空格）
    t = re.sub(r"^(?:還有|有沒有)\s*", "", t)

    # Step 2：剝離後綴問句（含「還有貨嗎」「還有嗎」「還有」等「還」開頭後綴）
    t = re.sub(
        r"\s*(?:還有貨嗎|還有沒有貨|還有嗎|還有貨|有貨嗎|有沒有貨|可以訂嗎|能訂嗎|有得訂|訂購|缺貨|有嗎|有貨|能訂|可訂|有沒有|還有|庫存)\s*$",
        "", t,
    )
    t = re.sub(r"\s*嗎\s*$", "", t)
    t = re.sub(r"的$", "", t)
    t = t.strip()

    # 如果有剝離到東西，就用結果（排除太通用的字）
    _TOO_GENERIC = {"貨", "東西", "商品", "產品", "物", "品", "款", "這", "那", "它",
                    "照片", "圖片", "圖", "照", "相片"}
    if t and t != text and t not in _TOO_GENERIC:
        return t

    # Step 3：純英數編號 + 問句
    m = re.search(r"([A-Za-z0-9\-_]+)\s*(?:有貨|庫存|訂購|可以訂)", text)
    if m:
        return m.group(1).strip()

    # Step 4：最後手段 — 暴力刪除所有問句關鍵字
    cleaned = re.sub(
        r"(我要查詢|我要查|我想查詢|我想查|我要問|我想問|還有貨嗎|還有沒有貨|還有嗎|還有貨|有貨嗎|有沒有貨|可以訂嗎|能訂嗎|有得訂|訂購|缺貨|有嗎|有貨|能訂|可訂|請問|想問|問一下|查一下|有沒有|還有|庫存|你們|妳們|嗎)",
        "", text,
    ).strip()
    cleaned = re.sub(r"的$", "", cleaned).strip()
    return cleaned
