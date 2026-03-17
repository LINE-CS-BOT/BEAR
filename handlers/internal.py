"""
內部群組指令處理

1. 到貨通知：「T1202 到貨」→ push 等待通知的客戶
2. 幫訂單：  「幫 張三 訂 BB-232 3個」→ 建立 Ecount 訂單
3. 圖片識別：（在 on_image_message 呼叫）→ 識別產品 + 回 PO文
4. 通知登記：「T1202 通知 張三 3個」→ 手動幫客戶登記到貨通知
5. 庫存查詢：「K0236 庫存」「K0236 有多少」→ 查 Ecount 回覆
6. 上架：傳圖/影片 + PO文 → 儲存照片 + 更新產品PO文.txt
"""

import re
from pathlib import Path
from linebot.v3.messaging import (
    MessagingApi, PushMessageRequest, TextMessage, ImageMessage, VideoMessage,
)

from storage.notify import notify_store
from storage.customers import customer_store
from storage import specs as spec_store
from handlers.ordering import extract_quantity, _resolve_cust_code
from services.ecount import ecount_client, _sync_and_wait
from config import settings

# ── 正則 ──────────────────────────────────────────────────────────────
# 商品編號：英文1~3碼（可含 -）+ 數字3~6碼 + 可選後綴（-1、-2 等），例：T1202、BB-232、Q0312-1
_PROD_CODE_RE = re.compile(r'\b([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\b')

# 到貨觸發詞
_ARRIVAL_KW = ["到貨", "到了", "到齊", "收到了", "進來了", "到貨了", "貨到了", "貨到"]

# 格式A（每行一筆）：「張三 訂 T1202 3」
# group(1)=姓名  group(2)=產品代碼  group(3)=數量
_STAFF_ORDER_LINE_RE = re.compile(
    r'(.+?)\s+(?:訂|下單)\s+([A-Za-z]{1,3}-?\d{3,6})\s+([零一二三四五六七八九十百千\d]+)\s*(?:個|件|盒|套|箱|組)?\s*(.*)'
)
# group(4) = 尾段備註（如「不要黑色」），可為空
# 格式B 第一行：「張三訂」或「張三 訂」（後面沒有產品代碼）
# group(1)=姓名
_STAFF_ORDER_HEADER_RE = re.compile(
    r'^(.+?)\s*(?:訂|下單)\s*$'
)
# 格式B 後續每行：「T1202 3」
# group(1)=產品代碼  group(2)=數量
_STAFF_ORDER_ITEM_RE = re.compile(
    r'([A-Za-z]{1,3}-?\d{3,6})\s+([零一二三四五六七八九十百千\d]+)'
)
# 格式C（無需「訂」關鍵字）：「姓名 產品代碼 數量個 [備註]」，例：方力緯 Z3562 5個 不要黑色
# group(1)=姓名  group(2)=產品代碼  group(3)=數量  group(4)=尾段備註（可為空）
_STAFF_ORDER_DIRECT_RE = re.compile(
    r'^(.+?)\s+([A-Za-z]{1,3}-?\d{3,6})\s+([零一二三四五六七八九十百千\d]+)\s*(?:個|件|盒|套|箱|組)?\s*(.*?)$'
)

# 通知登記觸發詞：句首「通知登記」OR 句中/句尾含以下關鍵字
_NOTIFY_REG_START_RE = re.compile(r'^通知登記')
_NOTIFY_REG_INLINE_KW = ["需要到貨通知", "要到貨通知", "通知登記", "需要通知", "要通知", "登記通知", "到貨通知", "要登記", "需要登記"]
# 格式：「通知/登記 [姓名] 產品代碼」
# group(1)=可選姓名  group(2)=產品代碼
# 例：「通知 T1202」、「通知 張三 T1202」、「登記 T1202」、「登記 張三 T1202」
_NOTIFY_REG_SHORTHAND_RE = re.compile(
    r'^(?:通知|登記)\s+(?:(.+?)\s+)?([A-Za-z]{1,3}-?\d{3,6})', re.IGNORECASE
)
# 每一行：「姓名  產品代碼  [數量]」
_NOTIFY_REG_LINE_RE  = re.compile(
    r'(.+?)\s+([A-Za-z]{1,3}-?\d{3,6})(?:\s+([零一二三四五六七八九十百千\d]+)\s*(?:個|件|盒|套|箱|組)?)?'
)

# 品名下單 token（合體格式）：「衛生紙30箱」「泡澡球10件」
_ITEM_TOKEN_RE = re.compile(
    r'^([\u4e00-\u9fff\w]+?)(\d+)\s*(個|件|盒|套|箱|組)?$'
)
# 純數量 token（分離格式）：「30箱」「10件」「5個」
_QTY_ONLY_RE = re.compile(
    r'^(\d+)\s*(個|件|盒|套|箱|組)$'
)

# 中文數字解析
_CN_DIGITS = {'零':0,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9}
_CN_UNITS  = {'十':10,'百':100,'千':1000}

def _parse_qty(s: str) -> int:
    """支援阿拉伯數字和中文數字（三、十二、二十三）"""
    s = s.strip()
    if s.isdigit():
        return max(1, int(s))
    result, current = 0, 0
    for ch in s:
        if ch in _CN_DIGITS:
            current = _CN_DIGITS[ch]
        elif ch in _CN_UNITS:
            result += (current or 1) * _CN_UNITS[ch]
            current = 0
    result += current
    return result if result > 0 else 1


# ── 1. 到貨通知 ────────────────────────────────────────────────────────

def handle_internal_arrival(text: str, line_api: MessagingApi) -> str | None:
    """
    偵測到貨訊息，push 等待該產品通知的所有客戶。
    回傳回覆給群組的文字，若非到貨訊息則回傳 None。
    """
    # 含「到貨通知」「需要通知」等字眼 → 是登記指令而非到貨通知，不處理
    if any(kw in text for kw in _NOTIFY_REG_INLINE_KW):
        return None

    has_arrival_kw = any(kw in text for kw in _ARRIVAL_KW)
    if not has_arrival_kw:
        return None

    # 找出訊息中的產品編號
    codes = _PROD_CODE_RE.findall(text)
    if not codes:
        return None

    results = []
    for raw_code in codes:
        prod_code = raw_code.upper()
        pending   = [r for r in notify_store.get_pending()
                     if r["prod_code"].upper() == prod_code]
        if not pending:
            results.append(f"📦 {prod_code}：目前沒有人在等候通知")
            continue

        # push 每位等待的客戶
        notified = 0
        for entry in pending:
            uid       = entry["user_id"]
            prod_name = entry["prod_name"] or prod_code
            qty       = entry["qty_wanted"]
            cust      = customer_store.get_by_line_id(uid)
            boss_name = ""
            if cust:
                boss_name = (cust.get("real_name") or cust.get("display_name") or "").strip()

            msg = (
                f"老闆{'好！' if not boss_name else f' {boss_name}好！'}"
                f"您之前等待的【{prod_name}】已經到貨囉～\n"
                f"有需要請告訴我，幫您安排！😊"
            )
            try:
                line_api.push_message(
                    PushMessageRequest(to=uid, messages=[TextMessage(text=msg)])
                )
                notify_store.mark_notified(entry["id"])
                notified += 1
                print(f"[internal] 到貨通知: {prod_code} → {uid}")
            except Exception as e:
                print(f"[internal] 推播失敗 {uid}: {e}")

        results.append(f"📦 {prod_code}：已通知 {notified} 位客戶")

    return "\n".join(results)


# ── 2. 幫訂單 ─────────────────────────────────────────────────────────

def _do_order(
    cust_name_query: str,
    items_raw: list[tuple[str, int]],
    units: dict[str, str] | None = None,   # prod_cd → 單位（箱/件/個…）
    note: str = "",                         # 備註，放每個品項的 REMARK
) -> str:
    """
    共用下單邏輯。items_raw = [(prod_query, qty), ...]
    優先從 ecount_customers.json 查客戶，找不到再查 LINE 本地 DB。
    units 選填：{prod_cd: "箱"} 可讓訂單訊息顯示正確單位。
    回傳結果文字。
    """
    import json as _json
    from pathlib import Path as _Path

    _EC_PATH = _Path(__file__).parent.parent / "data" / "ecount_customers.json"
    cust_code  = ""
    cust_label = cust_name_query
    _phone     = ""

    # 1. 先查 Ecount 客戶清單
    try:
        ec_list  = _json.loads(_EC_PATH.read_text(encoding="utf-8"))
        ec_match = next((x for x in ec_list if x.get("name", "") == cust_name_query), None)
        if not ec_match:
            ec_match = next((x for x in ec_list if cust_name_query in x.get("name", "")), None)
        if ec_match:
            cust_code  = ec_match.get("code", "")
            cust_label = ec_match.get("name", cust_name_query)
            _phone     = ec_match.get("phone", "") or ec_match.get("tel", "") or ""
            print(f"[internal] Ecount 客戶: {cust_label} → {cust_code}", flush=True)
    except Exception as e:
        print(f"[internal] ecount_customers.json 讀取失敗: {e}", flush=True)

    # 2. 找不到 → fallback 查 LINE 本地 DB
    if not cust_code:
        cust_matches = customer_store.search_by_name(cust_name_query)
        if not cust_matches:
            return f"❌ 找不到客戶「{cust_name_query}」"
        if len(cust_matches) > 1:
            names = "、".join(c.get("real_name") or c.get("display_name", "?") for c in cust_matches[:5])
            return f"⚠️ 「{cust_name_query}」有多位：{names}"
        cust       = cust_matches[0]
        user_id    = cust["line_user_id"]
        cust_label = cust.get("real_name") or cust.get("display_name") or cust_name_query
        codes      = customer_store.get_ecount_codes_by_line_id(user_id)
        if codes:
            cust_code = codes[0]["ecount_cust_cd"]
        else:
            existing  = customer_store.get_ecount_cust_code(user_id, default="")
            cust_code = existing or _resolve_cust_code(user_id) or settings.ECOUNT_DEFAULT_CUST_CD
        _phone = (customer_store.get_by_line_id(user_id) or {}).get("phone", "") or ""

    # 3. 查詢產品，組成 items
    order_items = []
    for prod_query, qty in items_raw:
        item = ecount_client.lookup(prod_query)
        if not item:
            return f"❌ 找不到產品「{prod_query}」"
        order_items.append({
            "prod_cd":   item["code"],
            "prod_name": item["name"] or item["code"],
            "qty":       qty,
            "note":      note,
        })

    # 4. 建立訂單
    slip_no = ecount_client.save_order(
        cust_code=cust_code,
        items=[{"prod_cd": i["prod_cd"], "qty": i["qty"], "note": i.get("note", "")} for i in order_items],
        phone=_phone,
    )

    if slip_no:
        detail = "、".join(f"{i['prod_name']}×{i['qty']}" for i in order_items)
        print(f"[internal] 代訂成功: {slip_no} | {cust_label} | {detail}")
        lines_out = [f"✅ {cust_label}｜{slip_no}"]
        for i in order_items:
            unit     = (units or {}).get(i["prod_cd"], "個")
            note_str = f"（{i['note']}）" if i.get("note") else ""
            lines_out.append(f"  {i['prod_name']} × {i['qty']} {unit}{note_str}")
        return "\n".join(lines_out)
    else:
        detail = "、".join(f"{i['prod_name']}×{i['qty']}" for i in order_items)
        print(f"[internal] 代訂失敗: {cust_code} | {detail}")
        return f"❌ {cust_label}｜訂單建立失敗"


def handle_internal_order(
    text: str,
    line_api: MessagingApi,
    group_id: str | None = None,
) -> str | None:
    """
    代訂單，支援多種格式：

    格式A（每行獨立）：  張三 訂 T1202 3
    格式B（多行）：      張三訂 / T1202 3 / T1808 5
    格式C（直接）：      方力緯 Z3562 5個
    格式D（品名單品）：  曹竣智 要 洗衣球 5
    格式E（品名多品）：  楊庭瑋 衛生紙30箱 泡澡球10件  ← 無貨號，自動搜尋+確認
    """
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if not lines:
        return None

    # ── 格式B 判斷：第一行符合「姓名訂」且後面沒有產品代碼 ──
    header_m = _STAFF_ORDER_HEADER_RE.match(lines[0])
    if header_m and len(lines) > 1:
        # 確認第一行沒有夾帶產品代碼
        if not _STAFF_ORDER_LINE_RE.search(lines[0]):
            cust_name = header_m.group(1).strip()
            items_raw = []
            for l in lines[1:]:
                im = _STAFF_ORDER_ITEM_RE.search(l)
                if im:
                    items_raw.append((im.group(1).strip(), _parse_qty(im.group(2))))
            if items_raw:
                return _do_order(cust_name, items_raw)

    # ── 格式A：每行各自獨立 ──
    valid = [(l, _STAFF_ORDER_LINE_RE.search(l)) for l in lines]
    valid = [(l, m) for l, m in valid if m]
    if valid:
        results = []
        for _line, m in valid:
            _note_a = m.group(4).strip() if m.lastindex >= 4 else ""
            res = _do_order(
                cust_name_query=m.group(1).strip(),
                items_raw=[(m.group(2).strip(), _parse_qty(m.group(3)))],
                note=_note_a,
            )
            results.append(res)
        return "\n".join(results)

    # ── 格式C：「姓名 產品代碼 數量個 [備註]」（單行，無「訂」關鍵字）──
    if len(lines) == 1:
        m_c = _STAFF_ORDER_DIRECT_RE.match(lines[0])
        if m_c:
            _note_c = m_c.group(4).strip() if m_c.lastindex >= 4 else ""
            return _do_order(
                cust_name_query=m_c.group(1).strip(),
                items_raw=[(m_c.group(2).strip(), _parse_qty(m_c.group(3)))],
                note=_note_c,
            )

    # ── 格式D：「姓名 要/訂/下單 商品名 數量」（品名搜尋下單，無需先查庫存）──
    # 例：「曹竣智 要 洗衣球 5」、「幫曹竣智 訂 洗衣球 5個」
    if len(lines) == 1:
        m_d = _STAFF_ORDER_PROD_NAME_RE.match(lines[0])
        if m_d:
            cust_name_d  = m_d.group(1).strip()
            prod_keyword = m_d.group(2).strip()
            qty_d        = _parse_qty(m_d.group(3))
            # 搜尋商品
            matched_codes = ecount_client.search_products_by_name(prod_keyword)
            if not matched_codes:
                return f"❌ 找不到商品「{prod_keyword}」，請用產品代碼下單"
            # 篩選有庫存的
            stock_hits = []
            for code in matched_codes:
                item = ecount_client.lookup(code)
                if item and (item.get("qty") or 0) > 0:
                    stock_hits.append((code, item.get("name") or code))
            if not stock_hits:
                return f"❌ 「{prod_keyword}」目前無庫存，請確認產品"
            if len(stock_hits) > 1:
                opts = "\n".join(f"  • {c}　{n}" for c, n in stock_hits[:5])
                return f"⚠️ 找到多款「{prod_keyword}」有庫存：\n{opts}\n請用產品代碼指定，例：{cust_name_d} {stock_hits[0][0]} {qty_d}"
            prod_code_d, prod_name_d = stock_hits[0]
            return _do_order(
                cust_name_query=cust_name_d,
                items_raw=[(prod_code_d, qty_d)],
            )

    # ── 格式E：品名多品項下單（無貨號，自動搜尋 + 等待確認）──────────────────
    # 例：「楊庭瑋 衛生紙30箱 泡澡球10件」
    # 偵測條件：無貨號格式，且包含品名+數量+單位（合體或分離格式均支援）
    # 合體：「衛生紙30箱」  分離：「衛生紙 30箱」
    # 也支援：「鄭鉅耀 (大)多色麥克風音響 × 10 個 備註:不要黑色」
    if group_id and not _PROD_CODE_RE.search(text):
        # 預處理：擷取備註、去 × 符號、合併「10 個」→「10個」
        _note_m = re.search(r'備註[:：](.*)', text)
        _note   = _note_m.group(1).strip() if _note_m else ""
        _text_e = re.sub(r'備註[:：].*', '', text).strip()
        _text_e = re.sub(r'[×Xx]\s*', '', _text_e)
        _text_e = re.sub(r'(\d+)\s+([個件盒套箱組])', r'\1\2', _text_e)
        _tokens = _text_e.strip().split()
        if len(_tokens) >= 2:
            # 找第一個品項 token 的起始位置（合體或分離格式）
            _item_start = None
            for _i, _tok in enumerate(_tokens):
                if _i == 0:
                    continue
                if _ITEM_TOKEN_RE.match(_tok):          # 合體：「衛生紙30箱」
                    _item_start = _i
                    break
                if (_QTY_ONLY_RE.match(_tok) and _i >= 2):  # 純數量 token，前一個是品名
                    _item_start = _i - 1
                    break
                if (_i + 1 < len(_tokens) and           # 分離：「衛生紙」+「30箱」
                        _QTY_ONLY_RE.match(_tokens[_i + 1])):
                    _item_start = _i
                    break

            if _item_start:
                _customer = " ".join(_tokens[:_item_start])
                # 解析品項（同時支援合體與分離格式）
                _name_items: list[tuple[str, int, str]] = []
                _j = _item_start
                while _j < len(_tokens):
                    _tok = _tokens[_j]
                    _mc = _ITEM_TOKEN_RE.match(_tok)
                    if _mc:                              # 合體：「衛生紙30箱」
                        _name_items.append((_mc.group(1), int(_mc.group(2)), _mc.group(3) or "個"))
                        _j += 1
                        continue
                    # 分離：「衛生紙」+ 下一個「30箱」
                    if _j + 1 < len(_tokens):
                        _mq = _QTY_ONLY_RE.match(_tokens[_j + 1])
                        if _mq:
                            _name_items.append((_tok, int(_mq.group(1)), _mq.group(2) or "個"))
                            _j += 2
                            continue
                    _j += 1

                if _name_items:
                    # 搜尋每個品名
                    _resolved: list[dict] = []
                    _ambiguous: list[dict] = []
                    _not_found: list[str]  = []

                    from handlers.ordering import resolve_order_qty
                    _IS_BULK = {"箱", "件"}

                    for _name, _qty, _unit in _name_items:
                        _is_bulk = _unit in _IS_BULK

                        # 去掉括號前綴（如「(大)」「(小)」），再做品名搜尋
                        _search_name = re.sub(r'^\([^)]+\)\s*', '', _name).strip() or _name

                        # 品名搜尋：若為箱/件單位，先搜「品名+箱」，找不到再搜「品名」
                        _codes = []
                        if _is_bulk:
                            _codes = ecount_client.search_products_by_name(_search_name + "箱")
                        if not _codes:
                            _codes = ecount_client.search_products_by_name(_search_name)

                        if not _codes:
                            _not_found.append(_name)
                        elif len(_codes) == 1:
                            _it = ecount_client.lookup(_codes[0])
                            _pn = (_it.get("name") if _it else "") or _codes[0]
                            # 箱/件換算
                            _actual_qty = resolve_order_qty(_codes[0], _qty) if _is_bulk else _qty
                            _resolved.append({
                                "query": _name, "code": _codes[0],
                                "name": _pn, "qty": _actual_qty,
                                "display_qty": _qty, "unit": _unit,
                            })
                        else:
                            # 多個結果：若為箱/件，優先選 unit==箱 的變體
                            _auto_code = None
                            if _is_bulk:
                                _box_variants = [
                                    _c for _c in _codes
                                    if (ecount_client.get_product_cache_item(_c) or {}).get("unit") == "箱"
                                ]
                                if len(_box_variants) == 1:
                                    _auto_code = _box_variants[0]
                            if _auto_code:
                                _it = ecount_client.lookup(_auto_code)
                                _pn = (_it.get("name") if _it else "") or _auto_code
                                _actual_qty = resolve_order_qty(_auto_code, _qty)
                                _resolved.append({
                                    "query": _name, "code": _auto_code,
                                    "name": _pn, "qty": _actual_qty,
                                    "display_qty": _qty, "unit": _unit,
                                })
                            else:
                                _cands = []
                                for _c in _codes[:5]:
                                    _ci = ecount_client.lookup(_c)
                                    _cn = (_ci.get("name") if _ci else "") or _c
                                    _cands.append((_c, _cn))
                                _ambiguous.append({
                                    "query": _name, "qty": _qty, "unit": _unit,
                                    "candidates": _cands,
                                })

                    # 有多重符合 → 儲存 state，逐一詢問
                    if _ambiguous:
                        from storage.state import state_manager as _sm
                        _sm.set(group_id, {
                            "action":          "pending_ambiguous_resolve",
                            "customer":        _customer,
                            "resolved":        _resolved,
                            "ambiguous_queue": _ambiguous,
                            "not_found":       _not_found,
                            "note":            _note,
                        })
                        return _build_ambiguous_ask(_ambiguous[0], _resolved, _not_found)

                    # 完全找不到
                    if not _resolved:
                        nf = "、".join(_not_found)
                        return f"❌ 找不到「{nf}」，請確認品名或改用貨號下單"

                    # 全部 resolved → 設 state 等待確認
                    from storage.state import state_manager as _sm
                    _sm.set(group_id, {
                        "action":   "pending_name_order_confirm",
                        "customer": _customer,
                        "items":    _resolved,
                        "note":     _note,
                    })
                    _lines = [f"確認下單嗎？\n\n👤 {_customer}"]
                    for _r in _resolved:
                        _dq = _r.get("display_qty", _r["qty"])
                        _aq = _r["qty"]
                        _u  = _r["unit"]
                        if _dq != _aq:  # 有換算
                            _lines.append(f"  📦 {_r['name']}（{_r['code']}）× {_dq} {_u} = {_aq} 個")
                        else:
                            _lines.append(f"  📦 {_r['name']}（{_r['code']}）× {_aq} {_u}")
                    if _not_found:
                        _lines.append(f"\n⚠️ 找不到：{'、'.join(_not_found)}（已略過）")
                    _lines.append("\n回「確認」建立訂單，「取消」放棄")
                    return "\n".join(_lines)

    return None


# ── 2b. 模糊詢問輔助 ──────────────────────────────────────────────────

def _build_ambiguous_ask(amb: dict, resolved: list, not_found: list) -> str:
    """建立詢問訊息：顯示目前模糊品項的選項清單"""
    lines = [f"❓「{amb['query']}」有多個結果，請選擇："]
    for idx, (_ac, _an) in enumerate(amb["candidates"], 1):
        lines.append(f"  {idx}. {_ac}　{_an}")
    lines.append(f"\n回序號（如 1）或貨號（如 {amb['candidates'][0][0]}）")
    if resolved:
        lines.append("\n已確認：")
        for _r in resolved:
            _dq = _r.get("display_qty", _r["qty"])
            _aq = _r["qty"]
            _u  = _r["unit"]
            if _dq != _aq:
                lines.append(f"  ✅ {_r['name']}（{_r['code']}）× {_dq} {_u} = {_aq} 個")
            else:
                lines.append(f"  ✅ {_r['name']}（{_r['code']}）× {_aq} {_u}")
    return "\n".join(lines)


def handle_ambiguous_resolve(group_id: str, text: str) -> str | None:
    """
    處理模糊品項的選擇回覆。
    state action == "pending_ambiguous_resolve" 時介入，否則回 None。
    支援：序號（1/2/3）、貨號（Z2095）、品名關鍵字（厚衛生紙）
    """
    from storage.state import state_manager as _sm
    from handlers.ordering import resolve_order_qty
    state = _sm.get(group_id)
    if not state or state.get("action") != "pending_ambiguous_resolve":
        return None

    t = text.strip()

    # 取消
    if any(kw in t.lower() for kw in {"取消", "cancel", "❌", "算了", "不用"}):
        _sm.clear(group_id)
        return "❌ 已取消建單"

    customer        = state["customer"]
    resolved        = state["resolved"]
    ambiguous_queue = state["ambiguous_queue"]
    not_found       = state.get("not_found", [])
    current         = ambiguous_queue[0]
    candidates      = current["candidates"]  # [(code, name), ...]
    qty             = current["qty"]
    unit            = current["unit"]
    is_bulk         = unit in {"箱", "件"}

    # 解析選擇
    chosen_code = None

    # ① 序號
    if t.isdigit():
        idx = int(t) - 1
        if 0 <= idx < len(candidates):
            chosen_code = candidates[idx][0]

    # ② 貨號（直接出現在候選清單中）
    if not chosen_code:
        t_up = t.upper().split()[0] if t else ""
        for _ac, _ in candidates:
            if _ac.upper() == t_up:
                chosen_code = _ac
                break

    # ③ 品名關鍵字（在候選品名中模糊比對）
    if not chosen_code:
        t_kw = t.upper()
        for _ac, _an in candidates:
            if t_kw in _an.upper():
                chosen_code = _ac
                break

    if not chosen_code:
        # 無法識別 → 重新顯示選項
        return _build_ambiguous_ask(current, resolved, not_found)

    # 找到選擇 → 換算數量並加入 resolved
    _it = ecount_client.lookup(chosen_code)
    _pn = (_it.get("name") if _it else "") or chosen_code
    _actual_qty = resolve_order_qty(chosen_code, qty) if is_bulk else qty
    resolved = resolved + [{
        "query": current["query"], "code": chosen_code,
        "name": _pn, "qty": _actual_qty,
        "display_qty": qty, "unit": unit,
    }]
    remaining = ambiguous_queue[1:]

    # 還有下一個模糊項目 → 繼續詢問
    if remaining:
        _sm.set(group_id, {
            "action":          "pending_ambiguous_resolve",
            "customer":        customer,
            "resolved":        resolved,
            "ambiguous_queue": remaining,
            "not_found":       not_found,
            "note":            state.get("note", ""),
        })
        return _build_ambiguous_ask(remaining[0], resolved, not_found)

    # 全部解決 → 直接建單
    _sm.clear(group_id)
    units     = {i["code"]: i["unit"] for i in resolved}
    items_raw = [(i["code"], i["qty"]) for i in resolved]
    result    = _do_order(customer, items_raw, units=units, note=state.get("note", ""))
    if not_found:
        result += f"\n⚠️ 找不到「{'、'.join(not_found)}」，已略過"
    return result


# ── 2c. 品名下單確認 ──────────────────────────────────────────────────

_CONFIRM_KW = {"確認", "對", "是", "好", "ok", "yes", "✅", "建單", "確定"}
_CANCEL_KW  = {"取消", "不對", "不是", "不用", "算了", "cancel", "no", "❌"}


def handle_name_order_confirm(group_id: str, text: str) -> str | None:
    """
    處理格式E下單的確認/取消。
    state action == "pending_name_order_confirm" 時才介入，否則回 None。
    """
    from storage.state import state_manager as _sm
    state = _sm.get(group_id)
    if not state or state.get("action") != "pending_name_order_confirm":
        return None

    t = text.strip().lower()
    if any(kw in t for kw in _CANCEL_KW):
        _sm.clear(group_id)
        return "❌ 已取消建單"

    if any(kw in t for kw in _CONFIRM_KW):
        customer  = state["customer"]
        items     = state["items"]   # [{"code","name","qty","unit",...}]
        units     = {i["code"]: i["unit"] for i in items}
        items_raw = [(i["code"], i["qty"]) for i in items]
        note      = state.get("note", "")
        _sm.clear(group_id)
        return _do_order(customer, items_raw, units=units, note=note)

    return None


# ── 4. 手動通知登記 ────────────────────────────────────────────────────

def handle_internal_notify_register(text: str) -> str | None:
    """
    幫客戶登記到貨通知，到貨時自動 push。
    支援多種格式：

        通知登記 張三 T1202 3個        ← 句首「通知登記」
        張三 T1202 需要到貨通知         ← 句尾關鍵字
        張三 T1202 3個 要通知          ← 含數量+句尾關鍵字
        通知 T1202                     ← shorthand 第一行產品代碼
        張三                           ← 後續每行一個客戶名
        李四
        通知登記
        張三 T1202 3個
        楊庭瑋 T1208 8個               ← 多行（名稱+代碼同行）
    """
    t = text.strip()

    has_inline_kw  = any(kw in t for kw in _NOTIFY_REG_INLINE_KW)
    is_start_fmt   = _NOTIFY_REG_START_RE.match(t)
    m_shorthand    = _NOTIFY_REG_SHORTHAND_RE.match(t)

    if not has_inline_kw and not is_start_fmt and not m_shorthand:
        return None

    # ── 格式 A：「通知 T1202\n張三\n李四」（shorthand：代碼在第一行，名字在後面）─
    if m_shorthand and not is_start_fmt:
        inline_name  = m_shorthand.group(1)   # 可能是 None（只有代碼在第一行）
        prod_code_sh = m_shorthand.group(2).upper()

        if inline_name:
            # 單行格式：「通知 張三 T1202」→ 只有一個名字
            names = [re.sub(r'\s+', '', inline_name.strip())]
        else:
            # 多行格式：「通知 T1202\n張三\n李四」→ 後面每行一個名字
            name_lines = [l.strip() for l in t.splitlines()[1:] if l.strip()]
            names = [re.sub(r'\s+', '', n) for n in name_lines]

        if not names:
            return None
        # 使用 shorthand 路徑，每個名字對應同一個 product code
        results = []
        for cust_name_query in names:
            cust_name_query = re.sub(r'\s+', '', cust_name_query)   # 移除名字內空格
            item      = ecount_client.lookup(prod_code_sh)
            prod_name = (item["name"] if item else "") or prod_code_sh
            matches   = customer_store.search_by_name(cust_name_query)
            if not matches:
                results.append(f"❌ 找不到客戶「{cust_name_query}」")
                continue
            if len(matches) > 1:
                names_str = "、".join(r.get("real_name") or r.get("display_name", "?") for r in matches[:5])
                results.append(f"⚠️ 「{cust_name_query}」有多位：{names_str}")
                continue
            cust      = matches[0]
            cust_uid  = cust["line_user_id"]
            cust_label = cust.get("real_name") or cust.get("display_name") or cust_name_query
            notify_id = notify_store.add(user_id=cust_uid, prod_code=prod_code_sh,
                                         prod_name=prod_name, qty_wanted=1)
            print(f"[internal] 通知登記(shorthand): #{notify_id} {cust_label} 等候 {prod_name}({prod_code_sh})")
            results.append(f"✅ {cust_label}｜{prod_name}（{prod_code_sh}）")
        return "\n".join(results) if results else None

    # ── 格式 B：句首「通知登記」──────────────────────────────────────────
    if is_start_fmt:
        first_newline = t.find("\n")
        if first_newline == -1:
            body = t[len("通知登記"):].strip()
            lines = [body] if body else []
        else:
            first_rest = t[:first_newline].replace("通知登記", "").strip()
            rest_lines = [l.strip() for l in t[first_newline:].splitlines() if l.strip()]
            lines = ([first_rest] if first_rest else []) + rest_lines
    else:
        # ── 格式 C：句尾/句中含關鍵字 ────────────────────────────────────
        cleaned = t
        for kw in _NOTIFY_REG_INLINE_KW:
            cleaned = cleaned.replace(kw, "").strip()
        lines = [cleaned] if cleaned else []

    if not lines:
        return None

    results = []
    for line in lines:
        # 名字內空格壓縮（「張 三」→「張三」）
        line = re.sub(r'(?<=[^\x00-\x7F])\s+(?=[^\x00-\x7F])', '', line)
        m = _NOTIFY_REG_LINE_RE.search(line)
        if not m:
            results.append(f"⚠️ 無法解析：「{line}」")
            continue

        cust_name_query = m.group(1).strip()
        prod_code       = m.group(2).upper()
        qty             = _parse_qty(m.group(3)) if m.group(3) else 1

        # 查詢產品名稱
        item      = ecount_client.lookup(prod_code)
        prod_name = (item["name"] if item else "") or prod_code

        # 查詢客戶
        matches = customer_store.search_by_name(cust_name_query)
        if not matches:
            results.append(f"❌ 找不到客戶「{cust_name_query}」")
            continue
        if len(matches) > 1:
            names = "、".join(r.get("real_name") or r.get("display_name", "?") for r in matches[:5])
            results.append(f"⚠️ 「{cust_name_query}」有多位：{names}")
            continue

        cust       = matches[0]
        cust_uid   = cust["line_user_id"]
        cust_label = cust.get("real_name") or cust.get("display_name") or cust_name_query

        notify_id = notify_store.add(
            user_id    = cust_uid,
            prod_code  = prod_code,
            prod_name  = prod_name,
            qty_wanted = qty,
        )
        print(f"[internal] 通知登記: #{notify_id} {cust_label} 等候 {prod_name}({prod_code}) x{qty}")
        results.append(f"✅ {cust_label}｜{prod_name}（{prod_code}）× {qty} 個")

    return "\n".join(results) if results else None


# ── 3. 圖片識別 → PO文 + 等待訂單 ───────────────────────────────────

# 圖片後訂單解析：「[name] [qty]個」或「幫[name]訂[qty]個」
# group(1)=客戶名  group(2)=數量
_INTERNAL_ORDER_RE = re.compile(
    r'^(?:幫\s*)?(.+?)\s+(?:訂\s*)?(\d+)\s*(?:個|件|盒|套|箱|組)?$'
)
# Format D（品名搜尋下單，無需先查庫存）：客戶名 + 要/訂/下單 + 商品關鍵字 + 數量
# 例：「曹竣智 要 洗衣球 5」、「幫曹竣智 訂 洗衣球 5個」
_STAFF_ORDER_PROD_NAME_RE = re.compile(
    r'^(?:幫\s*)?(.+?)\s+(?:要|訂|下單)\s+(.+?)\s+(\d+)\s*(?:個|件|盒|套|箱|組)?$'
)


def handle_internal_image(state_key: str, message_id: str, line_api: MessagingApi) -> str:
    """
    內部群組傳來圖片 → 識別產品 → 回傳 PO文，並設 state 等待「客戶名 N個」。
    state_key 應傳入 group_id（群組層級 state，任何成員都能接訂單）。
    """
    from services.vision import (
        download_image, identify_product, identify_product_weak, ocr_extract_candidates,
    )
    from storage.state import state_manager

    image_bytes = download_image(message_id)
    if not image_bytes:
        return "❌ 圖片讀取失敗，請重新傳一次"

    # ── 識別順序：① pHash 高可信 → ② OCR → ③ pHash 弱命中備援 ──
    prod_code = identify_product(image_bytes)   # pHash diff ≤ 10

    if not prod_code:
        # ② OCR 優先（能讀到貨號文字最準確）
        # 只嘗試「貨號格式」(字母+數字) 或中文詞，跳過純英文短詞（IN/LL/RE 等 OCR 雜訊）
        import re as _re
        _CODE_OR_ZH = _re.compile(r'(?:[A-Za-z]\d{2,}|[\u4e00-\u9fff]{2,})')
        for candidate in ocr_extract_candidates(image_bytes):
            if not _CODE_OR_ZH.search(candidate):
                continue  # 跳過純英文短詞
            matched = ecount_client._resolve_product_code(candidate)
            if matched:
                prod_code = matched
                print(f"[internal-image] OCR 比對成功 → {prod_code}（候選詞：{candidate!r}）")
                break

    if not prod_code:
        # ③ pHash 弱命中備援（diff 10-15）
        prod_code = identify_product_weak(image_bytes)
        if prod_code:
            print(f"[internal-image] pHash 弱命中備援 → {prod_code}")

    if not prod_code:
        return "❌ 無法識別產品，請確認圖片是否清晰"

    # 識別成功 → 強制同步最新庫存，再查 PO文
    _sync_and_wait()
    item      = ecount_client.lookup(prod_code)
    prod_name = (item["name"] if item else "") or prod_code

    po           = _format_po(prod_code)
    stock_detail = _fmt_stock_lines(item)

    # 設定 state，等待「客戶名 N個」指令（以 state_key=group_id 存，任何成員都能接）
    state_manager.set(state_key, {
        "action": "awaiting_internal_order",
        "prod_cd": prod_code,
        "prod_name": prod_name,
    })

    return f"{po}\n{stock_detail}"


def handle_internal_order_from_state(
    state_key: str,
    text: str,
    state: dict,
    line_api: MessagingApi,
) -> str | None:
    """
    staff 在圖片識別後說「客戶名 N個」或「幫客戶名訂N個」→ 建立訂單。
    state_key 為 group_id（群組層級，任何成員都能觸發）。
    非此格式回傳 None（讓其他 handler 繼續嘗試）。
    """
    from storage.state import state_manager

    prod_cd   = state.get("prod_cd", "")
    prod_name = state.get("prod_name") or prod_cd

    # 取消
    if any(kw in text for kw in ["取消", "算了", "不用", "不訂"]):
        state_manager.clear(state_key)
        return "已取消"

    # 含庫存查詢關鍵字 → 不當作訂單，讓其他 handler 處理
    if any(kw in text for kw in ["庫存", "有貨", "幾個", "多少"]):
        return None

    # ── 智慧解析：支援多種格式 ────────────────────────────────
    # ① 「曹竣智 5」/ 「曹竣智 5個」          → 基本格式
    # ② 「曹竣智 要 5」/ 「曹竣智 訂 5個」    → 含動詞
    # ③ 「曹竣智 要 洗衣球 5」                 → 含動詞 + 商品名（從 state 取產品）
    # ④ 「幫曹竣智 訂 洗衣球 5個」             → 含「幫」前綴
    _qty_tail_re = re.compile(r'(\d+)\s*(?:個|件|盒|套|箱|組)?\s*$')
    qty_m = _qty_tail_re.search(text.strip())
    if not qty_m:
        return None  # 沒有數字 → 格式不符

    qty = int(qty_m.group(1))

    # 策略：若含「要/訂/下單」動詞 → 動詞前面的部分就是客戶名
    _verb_sep_re = re.compile(r'\s+(?:要|訂|下單|買)\s+')
    verb_m = _verb_sep_re.search(text.strip())
    if verb_m:
        cust_part = text.strip()[:verb_m.start()].strip()
        if cust_part.startswith("幫"):
            cust_part = cust_part[1:].strip()
        cust_name_query = cust_part
    else:
        # 無動詞 → 數字前面全是客戶名（移除末尾殘留的「訂」「幫」）
        prefix = text.strip()[:qty_m.start()].strip()
        if prefix.startswith("幫"):
            prefix = prefix[1:].strip()
        cust_name_query = re.sub(r'\s+(?:訂|要|下單)\s*$', '', prefix).strip()

    if not cust_name_query:
        return None  # 無客戶名 → 格式不符

    # 客戶名太短（單字元）→ 可能是誤判（如貨號前綴）
    if len(cust_name_query) < 2:
        return None

    # 客戶名看起來是貨號（字母+數字，如 R0101）→ 拒絕，避免誤建單
    if re.match(r'^[A-Za-z]{1,3}\d{3,}$', cust_name_query):
        return None

    # ── 優先查 Ecount 客戶清單（ecount_customers.json）──────────────
    import json as _json
    from pathlib import Path as _Path
    _EC_PATH = _Path(__file__).parent.parent / "data" / "ecount_customers.json"
    cust_code  = ""
    cust_label = cust_name_query
    _phone     = ""

    try:
        ec_list = _json.loads(_EC_PATH.read_text(encoding="utf-8"))
        # 完全比對優先
        ec_match = next((x for x in ec_list if x.get("name", "") == cust_name_query), None)
        # 找不到 → 部分比對
        if not ec_match:
            ec_match = next((x for x in ec_list if cust_name_query in x.get("name", "")), None)
        if ec_match:
            cust_code  = ec_match.get("code", "")
            cust_label = ec_match.get("name", cust_name_query)
            _phone     = ec_match.get("phone", "") or ec_match.get("tel", "") or ""
            print(f"[internal] Ecount 客戶匹配: {cust_label} → {cust_code}", flush=True)
    except Exception as e:
        print(f"[internal] ecount_customers.json 讀取失敗: {e}", flush=True)

    # 找不到 → fallback 查本地 LINE 資料庫
    if not cust_code:
        matches = customer_store.search_by_name(cust_name_query)
        if not matches:
            return f"❌ 找不到客戶「{cust_name_query}」，請確認 Ecount 姓名"
        if len(matches) > 1:
            names = "、".join(
                r.get("real_name") or r.get("display_name", "?") for r in matches[:5]
            )
            return f"⚠️ 找到多位「{cust_name_query}」：{names}\n請輸入更完整的姓名"
        cust       = matches[0]
        cust_uid   = cust["line_user_id"]
        cust_label = cust.get("real_name") or cust.get("display_name") or cust_name_query
        codes = customer_store.get_ecount_codes_by_line_id(cust_uid)
        if codes:
            cust_code = codes[0]["ecount_cust_cd"]
        else:
            existing  = customer_store.get_ecount_cust_code(cust_uid, default="")
            direct_cd = cust.get("ecount_cust_cd", "")
            cust_code = existing or direct_cd or _resolve_cust_code(cust_uid) or settings.ECOUNT_DEFAULT_CUST_CD
        _phone = (customer_store.get_by_line_id(cust_uid) or {}).get("phone", "") or ""
    try:
        slip_no = ecount_client.save_order(
            cust_code=cust_code,
            items=[{"prod_cd": prod_cd, "qty": qty}],
            phone=_phone,
        )
    except Exception as e:
        print(f"[internal] save_order 例外: {e}", flush=True)
        return f"❌ 訂單建立失敗（API 錯誤：{e}）\n客戶：{cust_label}\n商品：{prod_name} × {qty} 個"

    state_manager.clear(state_key)  # 訂單完成，清除 state

    if slip_no:
        print(f"[internal] 圖片代訂成功: {slip_no} | {cust_label} | {prod_name} x{qty}")
        return (
            f"✅ 訂單建立成功\n"
            f"客戶：{cust_label}\n"
            f"商品：{prod_name} × {qty} 個"
        )
    else:
        print(f"[internal] 圖片代訂失敗: {cust_code} | {prod_name} x{qty}")
        return f"❌ 訂單建立失敗，請手動建立\n客戶：{cust_label}\n商品：{prod_name} × {qty} 個"


_PO_TXT_PATH = r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt"


def _get_raw_po_block(prod_code: str) -> str | None:
    """
    從原始 PO文.txt 找出對應產品的完整段落（空白行分隔）。
    找到則回傳段落文字，找不到回傳 None。
    """
    import sys
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass

    from pathlib import Path as _Path
    txt = _Path(_PO_TXT_PATH)
    try:
        if not txt.exists():
            return None
    except OSError:
        return None
    content = None
    for enc in ("cp950", "big5", "utf-8"):
        try:
            content = txt.read_text(encoding=enc)
            break
        except Exception:
            continue
    if content is None:
        return None

    code_upper = prod_code.upper()
    # 以空白行切成段落
    blocks = [b.strip() for b in content.split("\n\n") if b.strip()]
    for block in blocks:
        # 段落中任意一行含有該產品編號就算匹配
        if any(code_upper in line.upper() for line in block.splitlines()):
            return block
    return None


def _format_po(prod_code: str) -> str:
    """
    將產品資料格式化成 PO文：
    優先從原始 PO文.txt 抓完整段落；
    沒有才用 specs.json 的結構化欄位組合。
    """
    # 1. 優先取原始 PO文段落
    raw = _get_raw_po_block(prod_code)
    if raw:
        return raw

    # 2. 從 specs.json 組合
    spec = spec_store.get_by_code(prod_code)
    if not spec:
        item = ecount_client.lookup(prod_code)
        if item:
            return f"【{item['name'] or prod_code}】\n商品編號：{prod_code}"
        return f"商品編號：{prod_code}（查無規格資料）"

    lines = []
    name = spec.get("name") or ""
    if name:
        lines.append(f"【{name}】")
    lines.append(f"商品編號：{prod_code}")
    if spec.get("size"):
        lines.append(f"尺寸：{spec['size']}")
    if spec.get("weight"):
        lines.append(f"重量：{spec['weight']}")
    machines = spec.get("machine") or []
    if machines:
        lines.append(f"適用台型：{'、'.join(machines)}")
    if spec.get("price"):
        lines.append(f"售價：{spec['price']}")

    return "\n".join(lines)


# ── 5-a. 規格搜尋（台型 / 尺寸）────────────────────────────────────────

_MACHINE_TYPES = ["中巨", "標準台", "標準", "K霸", "巨無霸"]
_SPEC_QUERY_KW = ["有哪些", "有什麼", "哪些產品", "什麼產品", "產品", "有哪些產品", "推薦"]

import re as _re_spec
_SIZE_RE = _re_spec.compile(r'(\d+(?:\.\d+)?)\s*公分')


def handle_internal_spec_query(text: str) -> str | None:
    """
    「中巨的產品有哪些」→ 搜尋適合台型的產品 + 庫存
    「13公分的有哪些產品」→ 搜尋尺寸符合的產品 + 庫存
    回傳 None 表示不是此類查詢。
    """
    from storage.specs import get_by_machine, get_by_size

    # 必須含列表意圖關鍵字
    if not any(kw in text for kw in _SPEC_QUERY_KW):
        return None

    # 偵測台型
    matched_machine = next((m for m in _MACHINE_TYPES if m in text), None)
    # 偵測尺寸
    m_size = _SIZE_RE.search(text)
    size_kw = m_size.group(0) if m_size else None  # 例：「13公分」

    if not matched_machine and not size_kw:
        return None

    # 搜尋規格DB
    if matched_machine:
        specs = get_by_machine(matched_machine)
        label = f"「{matched_machine}」台型"
    else:
        specs = get_by_size(size_kw)
        label = f"「{size_kw}」尺寸"

    if not specs:
        return f"🔍 規格DB 目前沒有{label}的產品記錄"

    lines = [f"🔍 {label} 的產品（共 {len(specs)} 筆）：\n"]
    for s in specs:
        code = s.get("code", "")
        name = s.get("name", code)
        size = s.get("size", "")
        price = s.get("price", "")
        # 查庫存
        try:
            item = ecount_client.lookup(code)
            qty      = item.get("qty") if item else None
            preorder = item.get("preorder") if item else None
            stock_str = f"可售：{qty} 個" if qty is not None else "庫存查詢失敗"
            if (preorder or 0) > 0:
                stock_str += f"  預購：{preorder} 個"
        except Exception:
            stock_str = "庫存查詢失敗"
        lines.append(f"• {code}　{name}\n  {size}　{price}　{stock_str}")

    return "\n".join(lines)


# ── 5. 庫存查詢 ────────────────────────────────────────────────────────

_INV_QUERY_KW = [
    "庫存", "有多少", "幾個", "剩幾", "剩多少", "多少個",
    "有幾個", "有幾", "查庫存", "幾件", "幾箱", "幾盒",
    "有沒有貨", "有貨嗎", "還有嗎", "缺貨嗎",
]

_PREORDER_KW = [
    "預購", "可預購", "還可以訂", "還能訂", "能預購",
]

# 「產品有哪些」類型的觸發詞（只列品名，不查庫存）
_PRODUCT_LIST_KW = [
    "有哪些產品", "有什麼產品", "產品有哪些", "品項有哪些",
    "有什麼品項", "有哪些品項", "有哪些東西", "有什麼東西",
]

# 「清單限定詞」：結合庫存詞時 → 篩選 available>0 或 preorder>0
_LIST_KW = ["哪些", "什麼"]


def _fmt_inv_block(item: dict, prod_code: str) -> str:
    """將 lookup() 結果格式化成多行庫存明細"""
    name     = item.get("name") or prod_code
    qty      = item.get("qty")        # 可售庫存（扣bot保留後）
    balance  = item.get("balance")    # 倉庫庫存
    unfilled = item.get("unfilled")   # ERP未出
    incoming = item.get("incoming")   # 總公司未到
    preorder = item.get("preorder")   # 可預購數量
    # OAPI fallback 時 balance 為 None，改用 stock（BAL_QTY）
    if balance is None and item.get("stock") is not None:
        balance = item.get("stock")

    lines = [f"📦 {name}（{prod_code}）"]
    mid_lines = []

    if balance is not None:
        mid_lines.append(f"倉庫庫存：{balance} 個")
    if unfilled is not None:
        mid_lines.append(f"ERP未出：{unfilled} 個")
    if incoming is not None:
        mid_lines.append(f"總公司未到：{incoming} 個")
    if qty is None:
        mid_lines.append("可售庫存：查詢失敗")
    elif qty <= 0:
        mid_lines.append("可售庫存：0 個（缺貨）")
    else:
        mid_lines.append(f"可售庫存：{qty} 個")
    if preorder and preorder > 0:
        mid_lines.append(f"可預購：{preorder} 個")

    for i, ln in enumerate(mid_lines):
        prefix = "  └ " if i == len(mid_lines) - 1 else "  ├ "
        lines.append(prefix + ln)

    return "\n".join(lines)


def _fmt_stock_lines(item: dict) -> str:
    """
    回傳庫存明細純文字（不含產品名稱標題），供 PO文 + 庫存格式使用。
    格式：
      倉庫庫存：X 個
       ERP未出：X 個
       總公司未到：X 個
       可售庫存：X 個
       可預購：X 個
    """
    if not item:
        return "可售庫存：查詢失敗"
    balance  = item.get("balance")
    if balance is None and item.get("stock") is not None:
        balance = item.get("stock")
    unfilled = item.get("unfilled")
    incoming = item.get("incoming")
    qty      = item.get("qty")
    preorder = item.get("preorder")

    lines = []
    if balance is not None:
        lines.append(f"倉庫庫存：{balance} 個")
    if unfilled is not None:
        lines.append(f" ERP未出：{unfilled} 個")
    if incoming is not None:
        lines.append(f" 總公司未到：{incoming} 個")
    if qty is None:
        lines.append(" 可售庫存：查詢失敗")
    else:
        lines.append(f" 可售庫存：{qty} 個")
    if (preorder or 0) > 0:
        lines.append(f" 可預購：{preorder} 個")

    return "\n".join(lines)


_INFO_KW = ["資訊", "info", "INFO", "說明", "介紹"]

def handle_internal_product_info(text: str, state_key: str | None = None) -> str | None:
    """
    「T1102 資訊」→ 回傳 PO文 + 可售庫存 + 可預購庫存（與丟圖片相同格式）
    回傳 None 表示不是此類查詢。
    """
    from storage.state import state_manager
    if not any(kw in text for kw in _INFO_KW):
        return None
    codes = _PROD_CODE_RE.findall(text)
    if not codes:
        return None

    results = []
    last_code, last_name = None, None
    for raw_code in codes:
        prod_code = raw_code.upper()
        po = _format_po(prod_code)
        try:
            item = ecount_client.lookup(prod_code)
        except Exception as e:
            results.append(f"⚠️ {prod_code}：查詢失敗（{e}）")
            continue
        prod_name = (item.get("name") if item else "") or prod_code
        stock_detail = _fmt_stock_lines(item)
        results.append(f"{po}\n{stock_detail}")
        last_code, last_name = prod_code, prod_name

    if state_key and len(codes) == 1 and last_code:
        state_manager.set(state_key, {
            "action":    "awaiting_internal_order",
            "prod_cd":   last_code,
            "prod_name": last_name,
        })

    return "\n\n".join(results) if results else None


def handle_internal_inventory(text: str, state_key: str | None = None) -> str | None:
    from storage.state import state_manager
    """
    內部群庫存查詢：
      「K0236 庫存」「K0236 有多少」→ 完整明細
      「K0236 預購」             → 只回可預購數量
      「迪士尼 庫存有哪些」      → 品名搜尋 + 篩選 available>0 或 preorder>0
      「迪士尼 有哪些產品」      → 品名搜尋，只列品名清單（不查庫存）
    回傳 None 表示不是庫存查詢。
    """
    has_inv          = any(kw in text for kw in _INV_QUERY_KW)
    has_preorder     = any(kw in text for kw in _PREORDER_KW)
    has_product_list = any(kw in text for kw in _PRODUCT_LIST_KW)
    is_list_query    = any(kw in text for kw in _LIST_KW) and (has_inv or has_preorder)

    if not has_inv and not has_preorder and not has_product_list:
        return None

    codes = _PROD_CODE_RE.findall(text)

    # 沒找到產品編號 → 品名搜尋
    if not codes:
        _STRIP_KW = (
            _INV_QUERY_KW + _PREORDER_KW + _PRODUCT_LIST_KW
            + ["有哪些", "有什麼", "哪些", "什麼", "有嗎", "嗎", "有", "？", "?",
               "多少", "還", "個", "數量", "都", "各", "查詢", "產品", "品項"]
        )
        stripped = text
        for kw in sorted(_STRIP_KW, key=len, reverse=True):
            stripped = stripped.replace(kw, " ")
        tokens = [t.strip() for t in stripped.split() if len(t.strip()) >= 2]
        if not tokens:
            return None

        # ── 模式A：「產品有哪些」→ 只列品名，不查庫存 ──────────────
        if has_product_list and not has_inv and not has_preorder:
            matched: dict[str, None] = {}
            for token in tokens:
                for code in ecount_client.search_products_by_name(token):
                    matched[code] = None
            if not matched:
                return f"🔍 找不到包含「{'、'.join(tokens)}」的產品"
            kw_label = "".join(tokens)
            lines = [f"🔍 「{kw_label}」相關產品（共 {len(matched)} 筆）："]
            for code in matched:
                name = ecount_client._get_product_name(code) or ""
                lines.append(f"  • {code}　{name}")
            return "\n".join(lines)

        # ── 模式B：品名搜尋 + 篩選 available>0 或 preorder>0 ─────────
        matched2: dict[str, None] = {}
        for token in tokens:
            for code in ecount_client.search_products_by_name(token):
                matched2[code] = None
        if not matched2:
            return f"🔍 找不到包含「{'、'.join(tokens)}」的產品，請用產品編號查詢"

        results = []
        result_codes = []   # 所有展示的產品（qty>0 或 preorder>0）
        stock_codes  = []   # 只有實際可售庫存（qty>0）的產品，用來設 state
        for raw_code in matched2:
            prod_code = raw_code.upper()
            try:
                item = ecount_client.lookup(prod_code)
            except Exception:
                continue
            if not item:
                continue
            # 預購模式
            if has_preorder and not has_inv:
                preorder = item.get("preorder")
                if (preorder or 0) > 0:
                    name = item.get("name") or prod_code
                    results.append(f"📦 {name}（{prod_code}）\n  可預購：{preorder} 個")
                    result_codes.append((prod_code, name))
            else:
                # 庫存模式：只回 available>0 或 preorder>0
                if (item.get("qty") or 0) > 0 or (item.get("preorder") or 0) > 0:
                    results.append(_fmt_inv_block(item, prod_code))
                    result_codes.append((prod_code, item.get("name") or prod_code))
                # 只記「實際可售庫存 qty > 0」的（供 state 判斷用）
                if (item.get("qty") or 0) > 0:
                    stock_codes.append((prod_code, item.get("name") or prod_code))

        kw_label = "".join(tokens)
        if not results:
            return f"🔍 目前「{kw_label}」相關產品均無庫存或可預購數量"
        print(f"[internal] 品名庫存搜尋「{kw_label}」→ {len(results)} 筆有庫存（其中 {len(stock_codes)} 筆 qty>0）", flush=True)

        # 決定用哪個清單來設 state：
        # 優先用有實際庫存（qty>0）的；若全是預購則用 result_codes
        set_codes = stock_codes if stock_codes else result_codes

        # 剛好找到唯一一款有庫存的產品 → 設 state 讓 staff 直接說「客戶名 N個」下單
        if state_key and len(set_codes) == 1:
            single_code, single_name = set_codes[0]
            state_manager.set(state_key, {
                "action":    "awaiting_internal_order",
                "prod_cd":   single_code,
                "prod_name": single_name,
            })
            print(f"[internal] 品名搜尋設定 awaiting_internal_order: {single_code} for {state_key}", flush=True)
            return "\n\n".join(results) + "\n（已記錄，接著說「客戶名 N個」可直接建單）"
        elif state_key and len(set_codes) > 1:
            print(f"[internal] 品名搜尋找到 {len(set_codes)} 款有庫存，不設 state", flush=True)

        return "\n\n".join(results)

    mode = "預購" if (has_preorder and not has_inv) else "完整"
    print(f"[internal] 庫存查詢({mode}): {codes}", flush=True)
    results = []
    for raw_code in codes:
        prod_code = raw_code.upper()
        try:
            item = ecount_client.lookup(prod_code)
        except Exception as e:
            print(f"[internal] 庫存查詢失敗 {prod_code}: {e}", flush=True)
            results.append(f"⚠️ {prod_code}：查詢失敗（{e}）")
            continue
        if not item:
            results.append(f"❌ {prod_code}：查無此產品")
            continue

        name = item.get("name") or prod_code

        # 僅問預購（且沒有問庫存）→ 只回可預購數量
        if has_preorder and not has_inv:
            preorder = item.get("preorder")
            if preorder is None:
                results.append(f"📦 {name}（{prod_code}）\n  可預購：資料不足（請手動確認）")
            else:
                results.append(f"📦 {name}（{prod_code}）\n  可預購：{preorder} 個")
        else:
            results.append(_fmt_inv_block(item, prod_code))

        print(f"[internal] 庫存結果: {prod_code} qty={item.get('qty')} preorder={item.get('preorder')}", flush=True)

    # 查到剛好一個產品 → 設 state，等待「客戶名 N個」可直接下單
    if state_key and len(codes) == 1 and results:
        single_code = codes[0].upper()
        try:
            _item = ecount_client.lookup(single_code)
            _name = (_item.get("name") if _item else "") or single_code
        except Exception:
            _name = single_code
        state_manager.set(state_key, {
            "action":    "awaiting_internal_order",
            "prod_cd":   single_code,
            "prod_name": _name,
        })
        print(f"[internal] 設定 awaiting_internal_order: {single_code} for {state_key}", flush=True)
        return ("\n\n".join(results) + "\n（已記錄，接著說「客戶名 N個」可直接建單）") if results else None

    return "\n\n".join(results) if results else None


# ── 6. 分類標籤推送 ─────────────────────────────────────────────────────
# 觸發詞：「推送」「群發」「發給」「廣播」
_PUSH_TRIGGER_KW = ["推送", "群發", "發給", "廣播"]
_PUSH_TAG_RE = re.compile(
    r'(?:推送|群發|發給|廣播)\s*(\S+?)\s+(.+)',
    re.DOTALL
)

_IMG_EXTS   = {".jpg", ".jpeg", ".png", ".gif"}
_VIDEO_EXTS = {".mp4", ".mov"}
_ALL_MEDIA_EXTS = _IMG_EXTS | _VIDEO_EXTS


def _get_media_dir() -> Path | None:
    """取得產品照片資料夾 Path，若磁碟機未連線則回 None"""
    import sys as _sys
    if _sys.platform == "win32":
        try:
            import ctypes as _ct
            _ct.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass
    from config import settings as _s
    try:
        p = Path(_s.PRODUCT_MEDIA_PATH)
        if p.is_dir():
            return p
    except OSError:
        pass
    return None


def _stem_to_code(stem: str) -> str:
    """
    從檔名 stem（去副檔名）取出產品代碼。

    規則：尾端單一大寫字母是「版本」，去掉它剩下的才是產品代碼。
      T1102A   → T1102    （版本 A）
      T1102B   → T1102    （版本 B）
      T1102-1A → T1102-1  （版本 A）
      T1102-1B → T1102-1  （版本 B）
      T1102    → T1102    （無版本字母，完全比對）
    """
    s = stem.upper()
    if len(s) >= 2 and s[-1].isalpha() and not s[-2].isalpha():
        # 末尾是字母且倒數第二位不是字母（避免剪掉純字母代碼的最後一碼）
        return s[:-1]
    return s


def _match_product_media_files(prod_code: str, media_dir: Path) -> list[Path]:
    """
    掃描 media_dir，找出所有屬於 prod_code 的媒體檔案。

    識別邏輯：
      先用 _stem_to_code() 把檔名轉成產品代碼，再與目標代碼比對。

      T1102A.jpg   → code=T1102  → 符合 T1102 ✅，不符合 T1102-1 ❌
      T1102-1A.jpg → code=T1102-1 → 符合 T1102-1 ✅，不符合 T1102 ❌
      T1102-1B.mp4 → code=T1102-1 → 符合 T1102-1 ✅
    """
    code_upper = prod_code.upper()
    matches    = []
    try:
        for f in media_dir.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in _ALL_MEDIA_EXTS:
                continue
            if _stem_to_code(f.stem) == code_upper:
                matches.append(f)
    except OSError:
        pass
    return sorted(matches, key=lambda f: f.stem.upper())


def _build_media_messages(prod_code: str, files: list[Path], base_url: str) -> list:
    """
    將媒體檔案清單轉換成 LINE 訊息物件（ImageMessage / VideoMessage）。

    影片配對預覽圖策略：
      1. 同名 jpg/png（T1102A.mp4 → T1102A.jpg）
      2. 同產品任意一張 jpg（T1102*.jpg）
      3. 無圖 → 跳過該影片（LINE VideoMessage 必須有預覽圖）
    """
    base = base_url.rstrip("/")
    # 先建立 stem→圖片 URL 的對應表，以便影片找預覽圖
    stem_to_img_url: dict[str, str] = {}
    any_img_url = ""
    for f in files:
        if f.suffix.lower() in _IMG_EXTS:
            url = f"{base}/product-media/{f.name}"
            stem_to_img_url[f.stem.upper()] = url
            if not any_img_url:
                any_img_url = url

    messages = []
    for f in files:
        ext = f.suffix.lower()
        media_url = f"{base}/product-media/{f.name}"
        if ext in _IMG_EXTS:
            messages.append(ImageMessage(
                original_content_url=media_url,
                preview_image_url=media_url,
            ))
        elif ext in _VIDEO_EXTS:
            # 找預覽圖
            preview_url = (
                stem_to_img_url.get(f.stem.upper())  # 1. 同名 jpg
                or any_img_url                         # 2. 任意 jpg
            )
            if not preview_url:
                print(f"[tag-push] 影片 {f.name} 無預覽圖，跳過")
                continue
            messages.append(VideoMessage(
                original_content_url=media_url,
                preview_image_url=preview_url,
            ))
    return messages


def _push_messages_chunked(
    line_api, uid: str, text_msg: TextMessage, media_msgs: list
) -> None:
    """
    分批 push：第一批 = text + 最多 4 media（LINE 限 5 則/次）；
    後續批次每批最多 5 media。
    """
    first_batch = [text_msg] + media_msgs[:4]
    line_api.push_message(PushMessageRequest(to=uid, messages=first_batch))
    remaining = media_msgs[4:]
    while remaining:
        batch = remaining[:5]
        remaining = remaining[5:]
        line_api.push_message(PushMessageRequest(to=uid, messages=batch))


def _resolve_push_products(prod_query: str) -> list[tuple[str, str]]:
    """
    從推送指令的產品部分解析出所有產品代碼，
    回傳 [(prod_code, prod_name), ...] 清單。
    同時支援：
      - 多個產品代碼（T1202 T1208）
      - 品名搜尋（洗衣球）
      - 混合（T1202 洗衣球）
    """
    text = prod_query.strip()
    found_codes: list[str] = []
    remaining = text

    # 1. 先抓所有明確產品代碼
    for m in _PROD_CODE_RE.finditer(text):
        found_codes.append(m.group(1).upper())
    # 把已找到的代碼從 remaining 中移除，剩下的當品名搜尋
    remaining = _PROD_CODE_RE.sub("", text).strip()

    results: list[tuple[str, str]] = []
    seen: set[str] = set()

    for code in found_codes:
        if code in seen:
            continue
        seen.add(code)
        item = ecount_client.lookup(code)
        name = (item["name"] if item else "") or code
        results.append((code, name))

    # 2. 品名搜尋（移除掉代碼之後的殘餘文字）
    if remaining:
        candidates = ecount_client.search_products_by_name(remaining)
        for code in candidates[:5]:  # 最多 5 款
            uc = code.upper()
            if uc in seen:
                continue
            seen.add(uc)
            item = ecount_client.lookup(uc)
            name = (item["name"] if item else "") or uc
            results.append((uc, name))

    return results


def _get_ngrok_url() -> str:
    """同步查詢 ngrok 本地 API，取得目前公開 HTTPS 網址"""
    try:
        import requests as _req
        r = _req.get("http://localhost:4040/api/tunnels", timeout=2)
        tunnels = r.json().get("tunnels", [])
        for t in tunnels:
            if t.get("proto") == "https":
                return t["public_url"]
        if tunnels:
            return tunnels[0].get("public_url", "")
    except Exception:
        pass
    return ""


def handle_internal_tag_push(text: str, line_api: MessagingApi) -> str | None:
    """
    偵測「推送 [分類] [產品...]」指令，push PO文+圖片給所有有該標籤且有 LINE ID 的客戶。

    支援：
      推送 野獸國 T1202                   ← 單一產品代碼
      推送 VIP T1202 T1208               ← 多個產品代碼
      推送 K霸 洗衣球                    ← 品名搜尋
      推送 標準 T1202 洗衣球             ← 混合
      廣播 中句 T1202                    ← 廣播同義詞

    圖片：static/products/{CODE}/ 資料夾內所有 jpg/png（按檔名排序，最多4張）。
    """
    from storage.tags_config import load_tags

    if not any(kw in text for kw in _PUSH_TRIGGER_KW):
        return None

    m = _PUSH_TAG_RE.search(text.strip())
    if not m:
        all_tags = load_tags()
        return (
            "📣 推送格式：推送 [分類] [產品1] [產品2]...\n"
            "例：推送 野獸國 T1202 T1208\n"
            f"可用分類：{'、'.join(all_tags)}"
        )

    tag    = m.group(1).strip()
    prod_q = m.group(2).strip()

    # 驗證 tag 是否有效
    all_tags = load_tags()
    if tag not in all_tags:
        return (
            f"❌ 找不到分類「{tag}」\n"
            f"可用分類：{'、'.join(all_tags)}\n"
            "如需新增分類請至後台「標籤管理」設定"
        )

    # 解析產品清單
    products = _resolve_push_products(prod_q)
    if not products:
        return f"❌ 找不到任何產品「{prod_q}」，請確認產品名稱或代碼"

    # 取有該標籤的客戶
    customers = customer_store.get_customers_by_tag(tag)
    if not customers:
        return f"📣 找不到分類「{tag}」的客戶，或該分類客戶尚未登入 LINE"

    # 取 ngrok URL（用於圖片 / 影片公開連結）
    ngrok_url = _get_ngrok_url()
    if not ngrok_url:
        print("[internal-tag-push] 警告：無法取得 ngrok URL，媒體不會推送")

    # 取產品照片資料夾
    media_dir = _get_media_dir()
    if not media_dir:
        print("[internal-tag-push] 警告：產品照片磁碟機未連線")

    # 收集每個產品的推送資料：(code, name, po_text, [media_msgs])
    prod_data: list[tuple[str, str, str, list]] = []
    for prod_code, prod_name in products:
        po_text = _format_po(prod_code)
        media_msgs = []
        if ngrok_url and media_dir:
            files = _match_product_media_files(prod_code, media_dir)
            media_msgs = _build_media_messages(prod_code, files, ngrok_url)
            print(f"[internal-tag-push] {prod_code} 媒體檔案 {len(files)} 個 → {len(media_msgs)} 則訊息")
        prod_data.append((prod_code, prod_name, po_text, media_msgs))

    sent = 0
    failed = 0
    for cust in customers:
        uid = cust.get("line_user_id")
        if not uid:
            continue
        cust_name = cust.get("real_name") or cust.get("display_name") or ""
        greeting_prefix = f"老闆 {cust_name}～\n\n" if cust_name else "老闆～\n\n"

        try:
            # 第一款加上問候語，其後各款不重複問候
            for i, (prod_code, prod_name, po_text, media_msgs) in enumerate(prod_data):
                prefix = greeting_prefix if i == 0 else ""
                text_msg = TextMessage(text=prefix + po_text)
                # 分批推送：text + 最多 4 媒體/批，超過自動續批
                _push_messages_chunked(line_api, uid, text_msg, media_msgs)
            sent += 1
            codes_str = "、".join(c for c, *_ in prod_data)
            print(f"[internal-tag-push] {tag}／{codes_str} → {cust_name or uid}")
        except Exception as e:
            failed += 1
            print(f"[internal-tag-push] 推送失敗 {uid}: {e}")

    # 摘要回覆
    def _media_label(media_msgs: list) -> str:
        imgs  = sum(1 for m in media_msgs if isinstance(m, ImageMessage))
        vids  = sum(1 for m in media_msgs if isinstance(m, VideoMessage))
        parts = []
        if imgs:  parts.append(f"🖼×{imgs}")
        if vids:  parts.append(f"🎬×{vids}")
        return "  " + " ".join(parts) if parts else ""

    prod_summary = "\n".join(
        f"  • {name}（{code}）{_media_label(media)}"
        for code, name, po, media in prod_data
    )
    result = (
        f"📣 推送完成！\n"
        f"分類：{tag}　共 {len(customers)} 位客戶\n"
        f"產品：\n{prod_summary}\n"
        f"✅ 成功：{sent} 位"
    )
    if failed:
        result += f"\n❌ 失敗：{failed} 位（請查看 log）"
    if not ngrok_url:
        result += "\n⚠️ ngrok 未啟動，圖片/影片未推送"
    elif not media_dir:
        result += "\n⚠️ 產品照片磁碟機未連線，媒體未推送"
    return result


# ══════════════════════════════════════════════════════════════════════════
# 6. 上架指令系統
#
#  單品快速：上架 + 圖片 + PO文（同一 burst）→ 直接存
#  批次 Session（上架 / 存檔 單獨送出）：
#    → 貨號（純貨號訊息）→ 開新組
#    → 圖片/影片 → 累積到目前組
#    → PO文（含貨號的長文）→ 記錄 PO 並設貨號
#    → 完成 / 好了 → 批次存檔
#  存圖 Z3432      → 傳圖/影片 → 只存照片（不動PO文）
#  存文            → 文字內容 → 只存到 產品PO文.txt
# ══════════════════════════════════════════════════════════════════════════

_PO_FILE = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt")
_IMG_DIR = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品照片")

# 存圖指令正則：「存圖 Z3432」（替換舊圖）
_SAVE_IMG_RE = re.compile(r'存圖\s+([A-Za-z]{1,3}-?\d{3,6})', re.IGNORECASE)

# 加圖指令正則：「加圖 Z3432」（保留舊圖，追加新圖）
_ADD_IMG_RE  = re.compile(r'加圖\s+([A-Za-z]{1,3}-?\d{3,6})', re.IGNORECASE)

# Session 觸發詞與結束詞（「存圖」單獨傳也進 session；含貨號時走單品路徑）
_UPLOAD_TRIGGERS  = {"上架", "存檔", "存圖"}
_UPLOAD_FINISH_RE = re.compile(r'^(完成|好了|結束|done|finish)$', re.IGNORECASE)

# 純貨號偵測（整條訊息只有貨號）
_CODE_ONLY_RE = re.compile(r'^\s*([A-Za-z]{1,3}-?\d{3,6})\s*$', re.IGNORECASE)


# ── 共用：下載並儲存媒體（回傳 saved, failed 清單）────────────────────
def _save_media(code: str, media_items: list[dict],
                replace: bool = True) -> tuple[list[str], list[str]]:
    """
    下載並儲存媒體檔案。
    replace=True（預設）：先刪除該 code 所有舊檔，從 A 重新存。
    replace=False（加圖）：保留舊檔，從下一個可用字母接續存。
    """
    from services.vision import download_image
    saved, failed = [], []
    try:
        _IMG_DIR.mkdir(parents=True, exist_ok=True)
        if replace:
            # 重複上傳同一產品：刪除舊檔，從 A 重新開始
            for f in _IMG_DIR.glob(f"{code}[A-Z].*"):
                try:
                    f.unlink()
                    print(f"[upload] 刪除舊檔 {f.name}")
                except Exception:
                    pass
            letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        else:
            # 追加模式：找出已用字母，從下一個開始
            used = set()
            for f in _IMG_DIR.glob(f"{code}[A-Z].*"):
                stem = f.stem
                if len(stem) > len(code):
                    used.add(stem[len(code)])
            letters = [c for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if c not in used]
        for i, item in enumerate(media_items):
            if i >= len(letters):
                failed.append(f"（字母用完，第{i+1}個跳過）")
                continue
            letter = letters[i]
            ext    = ".jpg" if item["type"] == "image" else ".mp4"
            fname  = f"{code}{letter}{ext}"
            data   = download_image(item["msg_id"])
            if data:
                (_IMG_DIR / fname).write_bytes(data)
                saved.append(fname)
                print(f"[upload] 儲存 {fname} ({len(data)//1024}KB)")
            else:
                failed.append(f"{fname}（下載失敗）")
    except Exception as e:
        print(f"[upload] 媒體儲存錯誤: {e}")
        failed.append(f"（錯誤：{e}）")
    return saved, failed


# ── 共用：追加文字到 PO文.txt ─────────────────────────────────────────
def _append_po_text(content: str) -> bool:
    """
    寫入 PO文.txt。
    若同一貨號已存在，先刪除舊段落再寫入新版本，避免重複。
    """
    try:
        _PO_FILE.parent.mkdir(parents=True, exist_ok=True)
        new_block = content.strip()

        # 抓新內容的貨號（如有）
        m_code = _PROD_CODE_RE.search(new_block)
        new_code = m_code.group(1).upper() if m_code else None

        existing = _PO_FILE.read_text(encoding="utf-8") if _PO_FILE.exists() else ""

        if new_code and existing.strip():
            # 按空行切割既有段落，過濾掉相同貨號的舊段落
            import re as _re
            blocks = _re.split(r"\n{2,}", existing.strip())
            kept = []
            removed = 0
            for blk in blocks:
                m_blk = _PROD_CODE_RE.search(blk)
                if m_blk and m_blk.group(1).upper() == new_code:
                    removed += 1
                    print(f"[upload] PO文 替換舊段落 {new_code}（共 {removed} 筆）")
                else:
                    kept.append(blk)
            existing = "\n\n".join(kept)

        sep = "\n\n" if existing.strip() else ""
        _PO_FILE.write_text(existing + sep + new_block, encoding="utf-8")
        return True
    except Exception as e:
        print(f"[upload] PO文寫入失敗: {e}")
        return False


# ── 共用：觸發重建 ────────────────────────────────────────────────────
def _trigger_rebuild_safe():
    try:
        from services.refresh import trigger_rebuild
        trigger_rebuild()
    except Exception as e:
        print(f"[upload] trigger_rebuild 失敗: {e}")


# ── 共用：同步更新規格庫後生成架上標籤（背景執行緒）────────────────
def _generate_labels_sync(codes: list[str]) -> dict:
    """
    直接在當前執行緒（同步）呼叫 parse_specs + generate_labels。
    不用背景執行緒，確保 uvicorn reload 不會中斷。
    上架 10 個產品約 1-2 秒，可接受。

    回傳 {"pdfs": [...Path], "missing": [...str], "error": str|None}
    """
    result = {"pdfs": [], "missing": [], "error": None}
    try:
        from scripts.import_specs import parse_specs, OUTPUT, SOURCE
        import json as _json

        # 1. 同步解析 PO文，更新 specs.json
        try:
            exists = SOURCE.exists()
        except OSError:
            exists = False
        if exists:
            specs = parse_specs(SOURCE.read_text(encoding="utf-8"))
            OUTPUT.parent.mkdir(exist_ok=True)
            OUTPUT.write_text(
                _json.dumps(specs, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            import storage.specs as _spec_store
            _spec_store.reload()
            print(f"[label] specs 同步更新（{len(specs)} 筆）")
        else:
            # PO文.txt 不可存取，嘗試從現有 specs.json 讀取
            print("[label] PO文.txt 不可存取，使用現有 specs.json")
            try:
                specs = _json.loads(OUTPUT.read_text(encoding="utf-8")) if OUTPUT.exists() else {}
            except Exception:
                specs = {}

        # 2. 找出哪些 code 在 specs 裡找不到（提前偵測，讓 reply 能顯示）
        for c in codes:
            if c.upper() not in specs:
                result["missing"].append(c.upper())

        # 3. 生成架上標籤（generate_labels 內部也會判斷缺規格，但我們已先抓到）
        from scripts.generate_shelf_label import generate_labels
        pdfs = generate_labels(codes)
        result["pdfs"] = pdfs
        if pdfs:
            print(f"[label] 架上標籤已生成：{[p.name for p in pdfs]}")
        else:
            print(f"[label] 架上標籤佇列已更新，待湊滿3個")
    except Exception as _e:
        result["error"] = str(_e)
        print(f"[label] 生成架上標籤失敗：{_e}", flush=True)
    return result


# ══════ 指令 1：上架 ════════════════════════════════════════════════════
def handle_internal_product_upload(
    text: str,
    media_items: list[dict],
    line_api=None,
) -> str:
    """
    指令「上架」：PO文（含貨號）+ 圖片/影片
    → 存照片到 產品照片/ + 追加 PO文.txt + 觸發重建
    """
    # 從文字找貨號（去掉「上架」關鍵字再找）
    clean = text.replace("上架", "").strip()
    m = _PROD_CODE_RE.search(clean)
    if not m:
        return "⚠️ PO文中找不到貨號（格式如 Z3432），請確認後重試"
    code = m.group(1).upper()

    po_ok = _append_po_text(clean)
    saved, failed = _save_media(code, media_items) if media_items else ([], [])
    _trigger_rebuild_safe()
    label_result = _generate_labels_sync([code])

    lines = [f"✅ {code} 上架完成"]
    lines.append(f"• PO文{'已更新' if po_ok else '寫入失敗⚠️'}")
    if saved:
        lines.append(f"• 照片/影片：{', '.join(saved)}")
    if failed:
        lines.append(f"• 失敗：{', '.join(failed)}")
    if label_result["pdfs"]:
        names = "、".join(p.name for p in label_result["pdfs"])
        lines.append(f"🏷️ 架上標籤已生成：{names}")
    elif label_result["missing"]:
        missing_str = "、".join(label_result["missing"])
        lines.append(f"⚠️ 架上標籤規格缺失（{missing_str}），請補 PO文後重新上架")
    else:
        lines.append("📋 架上標籤已加入佇列，待湊滿3個自動生成")
    if label_result["error"]:
        lines.append(f"⚠️ 標籤生成錯誤：{label_result['error']}")
    return "\n".join(lines)


# ══════ 指令 2：存圖 Z3432 ══════════════════════════════════════════════
def handle_internal_save_images(code: str, media_items: list[dict]) -> str:
    """
    指令「存圖 Z3432」：取代所有舊照片，只儲存圖片/影片，不動 PO文.txt
    """
    if not media_items:
        return "⚠️ 沒有收到圖片或影片"
    saved, failed = _save_media(code, media_items, replace=True)
    _trigger_rebuild_safe()

    lines = [f"✅ {code} 圖片儲存完成（舊圖已替換）"]
    if saved:
        lines.append(f"• {', '.join(saved)}")
    if failed:
        lines.append(f"• 失敗：{', '.join(failed)}")
    return "\n".join(lines)


# ══════ 指令 2b：加圖 Z3432 ═════════════════════════════════════════════
def handle_internal_add_images(code: str, media_items: list[dict]) -> str:
    """
    指令「加圖 Z3432」：保留舊照片，追加新圖片/影片（從下一個字母接續）
    """
    if not media_items:
        return "⚠️ 沒有收到圖片或影片"
    saved, failed = _save_media(code, media_items, replace=False)
    _trigger_rebuild_safe()

    lines = [f"✅ {code} 圖片新增完成（舊圖保留）"]
    if saved:
        lines.append(f"• {', '.join(saved)}")
    if failed:
        lines.append(f"• 失敗：{', '.join(failed)}")
    return "\n".join(lines)


# ══════ 指令 3：存文 ════════════════════════════════════════════════════

def _split_po_by_code(text: str) -> list[str]:
    """
    將多段 PO文分割為獨立筆記。

    優先：空白行分段（同一訊息內按多次 Enter）→ 最自然，貨號位置不限。
    備援：無空白行時（分開送的訊息被合併），逐行掃描貨號，貨號換了就換段。
          用 search（非 match），貨號可在行中任意位置（如「編號：T1198」）。
    """
    # ── 優先：空白行分段 ────────────────────────────────────────────────
    paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
    if len(paragraphs) > 1:
        return paragraphs

    # ── 備援：逐行掃描，貨號換了就換段 ─────────────────────────────────
    lines        = text.splitlines()
    blocks       = []
    current      = []
    current_code = None

    for line in lines:
        m = _PROD_CODE_RE.search(line.strip())   # search，貨號不限行首
        if m:
            code = m.group(1).upper()
            if current_code and code != current_code and current:
                blocks.append("\n".join(current).strip())
                current = []
            current_code = code
        current.append(line)

    if current:
        blocks.append("\n".join(current).strip())

    return [b for b in blocks if b.strip()] or [text]


def handle_internal_save_text(content: str) -> str:
    """
    指令「存文」：
    - 單段 PO文 → 直接追加到 產品PO文.txt
    - 多段（不同貨號）→ 各自作為獨立筆記追加
    """
    content = content.replace("存文", "").strip()
    if not content:
        return "⚠️ 沒有文字內容可儲存"

    blocks = _split_po_by_code(content)

    # 若沒有貨號可分割 → 整段當一筆存
    if not blocks:
        blocks = [content]

    saved, failed = [], []
    for block in blocks:
        ok = _append_po_text(block)
        m  = _PROD_CODE_RE.search(block)
        label = m.group(1).upper() if m else f"（{block[:10]}...）"
        if ok:
            saved.append(label)
        else:
            failed.append(label)

    _trigger_rebuild_safe()

    lines = [f"✅ PO文已儲存 {len(saved)} 筆：{'、'.join(saved)}"] if saved else []
    if failed:
        lines.append(f"⚠️ 失敗：{'、'.join(failed)}")
    return "\n".join(lines) if lines else "⚠️ PO文寫入失敗，請確認磁碟機是否連線"


# ══════════════════════════════════════════════════════════════════════════
# 6b. 批次上架 Session（上架 / 存檔 單獨觸發）
#
#  State 結構：
#    action        = "uploading"
#    current_code  = str | None        ← 目前組的貨號
#    current_media = [...]             ← 目前組的圖片/影片（由 append_upload_media 原子追加）
#    current_po    = str               ← 目前組的 PO文
#    groups        = [{"code","media","po"}, ...]  ← 已完成的組
# ══════════════════════════════════════════════════════════════════════════

def handle_internal_upload_start(user_id: str) -> str | None:
    """「上架」/「存檔」單獨送出 → 進入批次上架 Session（靜默開始，完成才通知）"""
    from storage.state import state_manager
    state_manager.set(user_id, {
        "action":        "uploading",
        "current_code":  None,
        "current_media": [],
        "current_po":    "",
        "groups":        [],
    })
    print(f"[upload] {user_id[:10]}... 上架 session 開始", flush=True)
    return None  # 靜默，不通知群組


def handle_internal_upload_add_media(user_id: str, msg_id: str, media_type: str) -> None:
    """在 uploading session 中收到圖片/影片 → 原子追加到 current_media"""
    from storage.state import state_manager
    ok = state_manager.append_upload_media(user_id, {"msg_id": msg_id, "type": media_type})
    if not ok:
        print(f"[upload-session] append_upload_media 失敗（state 已失效）")


def handle_internal_upload_text(user_id: str, combined: str) -> str:
    """
    在 uploading session 中收到文字（5 秒合併後）：
    - 純貨號 → 結束上一組，開新組
    - PO文   → 記錄說明，並從中抓貨號
    - 完成   → （由 caller 判斷，不進此函數）
    """
    from storage.state import state_manager
    state = state_manager.get(user_id)
    if not state:
        return "⚠️ 上架 Session 已過期，請重新傳「上架」"

    combined = combined.strip()
    groups        = state.get("groups", [])
    current_code  = state.get("current_code")
    current_po    = state.get("current_po", "")

    # ── 純貨號：開新組 ────────────────────────────────────────────────
    m_code = _CODE_ONLY_RE.match(combined)
    if m_code:
        code = m_code.group(1).upper()
        # 把上一組推入 groups（current_media 由 state 直接讀取）
        cur_media = state.get("current_media", [])
        if current_code or cur_media:
            groups.append({
                "code":  current_code,
                "media": cur_media,
                "po":    current_po,
            })
        state["groups"]        = groups
        state["current_code"]  = code
        state["current_media"] = []
        state["current_po"]    = ""
        state_manager.set(user_id, state)
        return None  # 靜默，不通知群組

    # ── PO文（含貨號的長文）────────────────────────────────────────────
    m_po = _PROD_CODE_RE.search(combined)
    if m_po:
        code = m_po.group(1).upper()
        cur_media = state.get("current_media", [])
        # 若貨號不同且前一組有內容 → 先存前一組
        if current_code and current_code != code and cur_media:
            groups.append({
                "code":  current_code,
                "media": cur_media,
                "po":    current_po,
            })
            state["groups"]        = groups
            state["current_media"] = []
        state["current_code"] = code
        state["current_po"]   = combined
        state_manager.set(user_id, state)
        return None  # 靜默，不通知群組

    # ── 其他文字（無貨號）→ 當作補充說明 ──────────────────────────────
    state["current_po"] = (current_po + "\n" + combined).strip()
    state_manager.set(user_id, state)
    return None  # 靜默，不通知群組


def handle_internal_upload_finish(user_id: str) -> str:
    """「完成」→ 批次處理所有組，存照片 + PO文"""
    from storage.state import state_manager
    state = state_manager.get(user_id)
    if not state:
        return "沒有進行中的上架作業"

    groups        = state.get("groups", [])
    current_code  = state.get("current_code")
    current_media = state.get("current_media", [])
    current_po    = state.get("current_po", "")

    # 加入最後一組
    if current_code or current_media:
        groups.append({
            "code":  current_code,
            "media": current_media,
            "po":    current_po,
        })

    state_manager.clear(user_id)

    if not groups:
        return "沒有任何內容，上架取消"

    results = []
    uploaded_codes = []
    for g in groups:
        code  = g.get("code")
        media = g.get("media", [])
        po    = g.get("po", "")

        if not code:
            results.append(f"⚠️ 一組沒有貨號，跳過（{len(media)} 個檔案）")
            continue

        po_ok         = _append_po_text(po) if po.strip() else None
        saved, failed = _save_media(code, media) if media else ([], [])

        parts = []
        if po_ok:
            parts.append("PO文✓")
        if saved:
            parts.append("、".join(saved))
        if failed:
            parts.append(f"失敗:{','.join(failed)}")
        results.append(f"✅ {code}：{'  '.join(parts) if parts else '（無內容）'}")
        uploaded_codes.append(code)

    _trigger_rebuild_safe()
    label_result: dict = {"pdfs": [], "missing": [], "error": None}
    if uploaded_codes:
        label_result = _generate_labels_sync(uploaded_codes)

    # 組合架上標籤資訊到回覆末尾
    label_lines = []
    if label_result["pdfs"]:
        names = "、".join(p.name for p in label_result["pdfs"])
        label_lines.append(f"🏷️ 架上標籤已生成：{names}")
    elif label_result["missing"]:
        missing_str = "、".join(label_result["missing"])
        label_lines.append(f"⚠️ 規格缺失，標籤未生成：{missing_str}")
        label_lines.append("請補 PO文（含尺寸/重量/價格）後重新上架")
    else:
        label_lines.append("📋 架上標籤已加入佇列，待湊滿3個自動生成")
    if label_result["error"]:
        label_lines.append(f"⚠️ 標籤生成錯誤：{label_result['error']}")

    suffix = "\n" + "\n".join(label_lines) if label_lines else ""
    return "🏁 上架完成！\n" + "\n".join(results) + suffix


# ── 新增品項 ───────────────────────────────────────────────────────────────
# 格式（單行或多行均支援）：
#   新增品項 Z9999 (原)多色麥克風音響 個 條碼:1234567890 售價:299 規格:30×20cm
#   新增品項 Z9999
#   品名：(大)多色麥克風音響
#   條碼：1234567890
#   售價：299
#   加盟商：250
#   規格：30×20cm
_NEW_PROD_TRIGGER_RE = re.compile(r'^新增品項', re.IGNORECASE)
_NEW_PROD_CODE_RE    = re.compile(r'([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)')
_UNIT_WORDS_NP       = r'個|件|盒|套|箱|組|片|包|瓶|罐|條|支|只|枚|粒|顆|袋|塊'

# CLASS_CD 對應（品名前綴，按長到短排列避免短前綴先匹配）
_CLASS_CD_MAP = [
    (r'^\(原定\)', "00004"),
    (r'^\(定\)',   "00004"),
    (r'^\(原\)',   "00001"),
    (r'^\(大\)',   "00002"),
]

def _detect_class_cd(prod_name: str) -> str:
    """根據品名前綴判斷 CLASS_CD"""
    for pat, cd in _CLASS_CD_MAP:
        if re.match(pat, prod_name):
            return cd
    return ""

def _calc_in_price(class_cd: str, out_price_str: str, in_price_raw: str) -> str:
    """
    計算入庫單價：
    - 若消息內有明確加盟商價 → 直接使用（優先）
    - CLASS_CD 00001 → 售價 × 0.95（原裝，自動計算）
    - CLASS_CD 00002 → 售價 × 0.85（大包裝，自動計算）
    - CLASS_CD 00004 → 使用消息內加盟商價；無則空白
    - 其餘 → 使用消息內加盟商價；無則空白
    """
    if in_price_raw:
        return in_price_raw
    if not out_price_str:
        return ""
    try:
        out = float(out_price_str)
        if class_cd == "00001":
            return str(int(round(out * 0.95)))
        if class_cd == "00002":
            return str(int(round(out * 0.85)))
    except ValueError:
        pass
    return ""

def _parse_new_product_fields(text: str) -> dict | None:
    """
    解析「新增品項」訊息，回傳欄位 dict 或 None（格式不符）。
    支援單行與多行；關鍵字後接 :、：或空格均可。
    """
    if not _NEW_PROD_TRIGGER_RE.match(text.strip()):
        return None

    # 多行合成一行方便搜尋
    flat = " ".join(text.strip().splitlines())

    # 貨號（必填）
    m_code = _NEW_PROD_CODE_RE.search(flat)
    if not m_code:
        return None
    prod_cd = m_code.group(1).upper()

    # 條碼
    bar_code_m = re.search(r'條碼\s*[:：]?\s*(\S+)', flat)
    bar_code   = bar_code_m.group(1) if bar_code_m else prod_cd  # 預設條碼 = 品項編碼（貨號）

    # 售價 / 賣價 / 出庫單價
    out_price_m = re.search(r'(?:售價|賣價|出庫單價)\s*[:：]?\s*([\d.]+)', flat)
    out_price   = out_price_m.group(1) if out_price_m else ""

    # 加盟商價格 / 入庫單價（按長到短排，避免短詞先匹配）
    in_price_m   = re.search(r'(?:加盟商價格|加盟商商價|加盟商價|加盟商|入庫單價|進價)\s*[:：]?\s*([\d.]+)', flat)
    in_price_raw = in_price_m.group(1) if in_price_m else ""

    # 規格（取到行末或下一個關鍵字前）
    size_des_m = re.search(
        r'規格\s*[:：]?\s*(.+?)(?=\s+(?:條碼|售價|賣價|出庫|入庫|加盟|單位|品名|貨號)|$)',
        flat,
    )
    size_des = size_des_m.group(1).strip() if size_des_m else ""

    # 單位
    unit_kw_m = re.search(rf'單位\s*[:：]?\s*({_UNIT_WORDS_NP})', flat)
    if unit_kw_m:
        unit = unit_kw_m.group(1)
    else:
        unit_bare_m = re.search(rf'(?:^|\s)({_UNIT_WORDS_NP})(?:\s|$)', flat)
        unit = unit_bare_m.group(1) if unit_bare_m else "個"

    # 品名：從貨號後開始，剝除所有已識別欄位
    name_part = flat[m_code.end():]
    for strip_pat in [
        r'條碼\s*[:：]?\s*\S+',
        r'(?:售價|賣價|出庫單價)\s*[:：]?\s*[\d.]+',
        r'(?:加盟商價格|加盟商商價|加盟商價|加盟商|入庫單價|進價)\s*[:：]?\s*[\d.]+',
        r'規格\s*[:：]?\s*\S+(?:\s+\S+)*?(?=\s+(?:條碼|售價|賣價|出庫|入庫|加盟)|\s*$)',
        rf'單位\s*[:：]?\s*(?:{_UNIT_WORDS_NP})',
        rf'(?:^|\s)(?:{_UNIT_WORDS_NP})(?:\s|$)',
        r'品名\s*[:：]?\s*',
    ]:
        name_part = re.sub(strip_pat, ' ', name_part)
    prod_name = name_part.strip()

    if not prod_name:
        return None  # 品名必填

    class_cd = _detect_class_cd(prod_name)
    in_price = _calc_in_price(class_cd, out_price, in_price_raw)

    return {
        "prod_cd":   prod_cd,
        "prod_name": prod_name,
        "unit":      unit,
        "bar_code":  bar_code,
        "class_cd":  class_cd,
        "out_price": out_price,
        "in_price":  in_price,
        "size_des":  size_des,
        "cust":      "10003",
    }


_CLASS_LABEL_NP = {"00001": "原裝", "00002": "盒", "00004": "定裝"}
# 以貨號開頭且後接空白的行（貨號+品名同行）
_PROD_LINE_START_RE = re.compile(r'^([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s', re.IGNORECASE)
# 整行只有貨號（貨號獨行格式）
_CODE_ONLY_RE = re.compile(r'^[A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?$', re.IGNORECASE)
# 欄位關鍵字開頭的行（不會是品名）
_FIELD_LINE_RE = re.compile(
    r'^(?:規格|售價|賣價|出庫單價|加盟商|入庫單價|進價|條碼|單位)[:：]',
    re.IGNORECASE,
)


def _split_new_product_entries(text: str) -> list[str]:
    """
    把多筆品項訊息拆成多個單筆字串，每筆都可餵給 _parse_new_product_fields。

    支援三種格式：
    ① 單行：  新增品項 Z9999 (原)多色麥克風音響 個 售價:299
    ② 同行：  新增品項
              Z9999 (原)多色麥克風音響 個 售價:299
              Z0123 (大)泡澡球禮盒 個 售價:399
    ③ 多行：  新增品項
              Z9999
              (原)多色麥克風音響
              規格:12個/箱
              加盟商價:250
              Z0123
              (大)泡澡球禮盒
              規格:50顆
    """
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if not lines:
        return []

    first = re.sub(r'^新增品項\s*', '', lines[0], flags=re.IGNORECASE).strip()
    rest  = lines[1:]
    all_lines = ([first] if first else []) + rest

    if not all_lines:
        return []

    # 單行
    if len(all_lines) == 1:
        return ["新增品項 " + all_lines[0]]

    # 判斷格式：若存在「整行只有貨號」的行 → 格式③（貨號獨行）
    if any(_CODE_ONLY_RE.match(l) for l in all_lines):
        groups: list[list[str]] = []
        current: list[str] = []
        i = 0
        while i < len(all_lines):
            line = all_lines[i]
            if _CODE_ONLY_RE.match(line):
                if current:
                    groups.append(current)
                # 下一行若不是欄位關鍵字也不是貨號 → 視為品名，合併到同一行
                if (i + 1 < len(all_lines)
                        and not _CODE_ONLY_RE.match(all_lines[i + 1])
                        and not _FIELD_LINE_RE.match(all_lines[i + 1])):
                    current = [line + " " + all_lines[i + 1]]
                    i += 2
                else:
                    current = [line]
                    i += 1
            else:
                current.append(line)
                i += 1
        if current:
            groups.append(current)
        return ["新增品項 " + "\n".join(g) for g in groups]

    # 格式②：貨號+品名同行，按「以貨號開頭的行」切割
    groups = []
    current = []
    for line in all_lines:
        if _PROD_LINE_START_RE.match(line):
            if current:
                groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)
    return ["新增品項 " + "\n".join(g) for g in groups]


def _build_one_product(fields: dict) -> str:
    """把單筆解析結果組成回覆行"""
    prod_cd   = fields["prod_cd"]
    prod_name = fields["prod_name"]
    unit      = fields["unit"]
    bar_code  = fields["bar_code"]
    class_cd  = fields["class_cd"]
    out_price = fields["out_price"]
    in_price  = fields["in_price"]
    size_des  = fields["size_des"]

    extra: dict = {
        "PROD_TYPE": "3",
        "BAL_FLAG":  "1",
        "USE_FLAG":  "Y",
        "CUST":      "10003",
    }
    if bar_code:  extra["BAR_CODE"]  = bar_code
    if class_cd:  extra["CLASS_CD"]  = class_cd
    if out_price: extra["OUT_PRICE"] = out_price
    if in_price:  extra["IN_PRICE"]  = in_price
    if size_des:  extra["SIZE_DES"]  = size_des

    result = ecount_client.save_product(
        prod_cd=prod_cd, prod_name=prod_name, unit=unit, extra=extra,
    )

    from storage.new_products import new_products_store
    new_products_store.add(
        prod_cd=prod_cd,   prod_name=prod_name, unit=unit,
        bar_code=bar_code, class_cd=class_cd,   out_price=out_price,
        in_price=in_price, size_des=size_des,   cust="10003",
    )

    icon = "✅" if result else "⚠️"
    details = []
    if out_price: details.append(f"售:{out_price}")
    if in_price:  details.append(f"入:{in_price}")
    if size_des:  details.append(f"規:{size_des}")
    if class_cd:  details.append(_CLASS_LABEL_NP.get(class_cd, class_cd))
    detail_str = "　" + "　".join(details) if details else ""
    return f"{icon} {prod_cd} {prod_name}　{unit}{detail_str}"


def handle_internal_new_product(text: str) -> str | None:
    """
    新增品項指令：支援單筆與多筆，在 Ecount 建立品項並記錄到 admin 待審核清單。

    單筆：新增品項 Z9999 (原)多色麥克風音響 個 售價:299
    多筆：新增品項
          Z9999 (原)多色麥克風音響 個 售價:299
          Z0123 (大)泡澡球禮盒 個 售價:399 加盟商:250
    """
    if not _NEW_PROD_TRIGGER_RE.match(text.strip()):
        return None

    entries = _split_new_product_entries(text)
    if not entries:
        return None

    parsed = []
    for entry in entries:
        f = _parse_new_product_fields(entry)
        if f:
            parsed.append(f)

    if not parsed:
        return None

    print(f"[內部] 新增品項 共 {len(parsed)} 筆")
    result_lines = [_build_one_product(f) for f in parsed]

    header = f"📦 新增品項 {len(parsed)} 筆" if len(parsed) > 1 else ""
    footer = "📋 已記錄至 admin 待審核清單"
    parts  = ([header] if header else []) + result_lines + [footer]
    return "\n".join(parts)
