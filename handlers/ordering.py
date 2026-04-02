"""
下單流程處理

當 Bot 確認有貨並詢問數量後，客戶回覆數量，
Bot 自動在 Ecount 建立訂貨單並回覆確認。
"""

import re
import json
import subprocess
import sys
import threading
import time
from pathlib import Path

from linebot.v3.messaging import MessagingApi

from handlers import tone
from services.ecount import ecount_client
from storage.customers import customer_store
from storage import cart as cart_store
from config import settings

_ECOUNT_CUST_JSON   = Path(__file__).parent.parent / "data" / "ecount_customers.json"
_CUST_SYNC_SCRIPT   = Path(__file__).parent.parent / "scripts" / "sync_cust_from_web.py"
_CUST_JSON_STALE_H  = 4          # 超過幾小時就觸發客戶同步
_CUST_SYNC_TIMEOUT  = 180        # 客戶同步最長等待秒數
_sync_cust_lock     = threading.Lock()


# ---------------------------------------------------------------------------
# 地址比對工具（供 _resolve_cust_code 使用）
# ---------------------------------------------------------------------------

def _norm_phone(ph: str) -> str:
    return ph.replace(" ", "").replace("-", "").strip() if ph else ""


def _addr_key(addr: str) -> str:
    addr = addr.replace(" ", "").replace("\u3000", "")
    m = re.search(
        r'([\u4e00-\u9fff]{1,8}[\u8def\u8857\u9053\u5df7\u5f04])'
        r'(?:[\u4e00-\u9fff\d\u6bb5]*?)(\d+)[\u865f\u53f7]?',
        addr
    )
    return (m.group(1) + m.group(2)) if m else (addr[:12] if len(addr) >= 4 else "")


def _addr_match(la: str, ea: str) -> bool:
    if not la or not ea:
        return False
    la = la.replace(" ", "").replace("\u3000", "")
    ea = ea.replace(" ", "").replace("\u3000", "")
    if la in ea or ea in la:
        return True
    lk, ek = _addr_key(la), _addr_key(ea)
    return bool(lk and ek and len(lk) >= 4 and (lk in ek or ek in lk))


# ---------------------------------------------------------------------------
# 客戶代碼解析（代碼空白時即時比對或建立）
# ---------------------------------------------------------------------------

def _refresh_cust_json_if_stale():
    """
    若 ecount_customers.json 超過 4 小時未更新，同步一次（最多等 90 秒）。
    使用 threading.Lock 避免並發重複觸發。
    """
    if not _ECOUNT_CUST_JSON.exists():
        return
    age_h = (time.time() - _ECOUNT_CUST_JSON.stat().st_mtime) / 3600
    if age_h < _CUST_JSON_STALE_H:
        return
    with _sync_cust_lock:
        # 取到 lock 後再確認一次（可能前一個執行緒已同步完）
        if _ECOUNT_CUST_JSON.exists():
            age_h = (time.time() - _ECOUNT_CUST_JSON.stat().st_mtime) / 3600
            if age_h < _CUST_JSON_STALE_H:
                return
        print(f"[cust_code] ecount_customers.json 已 {age_h:.1f} 小時，觸發同步...")
        try:
            subprocess.run(
                [sys.executable, str(_CUST_SYNC_SCRIPT)],
                creationflags=subprocess.CREATE_NO_WINDOW,
                cwd=str(_CUST_SYNC_SCRIPT.parent.parent),
                timeout=_CUST_SYNC_TIMEOUT,
            )
            print("[cust_code] 客戶資料同步完成")
        except subprocess.TimeoutExpired:
            print(f"[cust_code] 客戶資料同步逾時（{_CUST_SYNC_TIMEOUT} 秒），使用既有資料")
        except Exception as e:
            print(f"[cust_code] 客戶資料同步失敗: {e}")


def _resolve_cust_code(user_id: str, do_refresh: bool = True) -> str | None:
    """
    用 ecount_customers.json 比對客戶代碼（電話/地址）。
    do_refresh=True 時，若 JSON 超過 4 小時自動觸發同步後再比對。
    找不到回傳 None（不建立新客戶）。
    """
    # 必要時先同步客戶資料
    if do_refresh:
        _refresh_cust_json_if_stale()

    cust = customer_store.get_by_line_id(user_id)
    if not cust:
        return None

    name    = (cust.get("real_name") or cust.get("display_name") or "").strip()
    phone   = (cust.get("phone") or "").strip()
    address = (cust.get("address") or "").strip()
    db_id   = cust.get("id")

    if not phone and not address and not name:
        print(f"[cust_code] 無姓名、手機、地址，無法比對")
        return None

    # ── JSON 比對（電話+姓名優先，地址次之）─────────────────
    if not _ECOUNT_CUST_JSON.exists():
        return None
    try:
        ec_list = json.loads(_ECOUNT_CUST_JSON.read_text(encoding="utf-8"))
        norm_ph = _norm_phone(phone)
        matched_code = None
        matched_addr = ""

        for ec in ec_list:
            code   = ec["code"]
            ec_adr = (ec.get("addr") or "").strip()
            ec_ph  = _norm_phone(ec.get("phone") or "")
            ec_tel = _norm_phone(ec.get("tel") or "")
            ec_nm  = (ec.get("name") or "").strip()

            if norm_ph and norm_ph in (ec_ph, ec_tel):
                if not name or name in ec_nm or ec_nm in name:
                    matched_code = code
                    matched_addr = ec_adr
                    break
                elif matched_code is None:
                    matched_code = code
                    matched_addr = ec_adr

        # 電話沒比到 → 用地址比對
        if not matched_code and address:
            for ec in ec_list:
                code   = ec["code"]
                ec_adr = (ec.get("addr") or "").strip()
                if ec_adr and _addr_match(address, ec_adr):
                    matched_code = code
                    matched_addr = ec_adr
                    break

        # 電話地址都沒比到 → 用名字精確比對
        if not matched_code and name:
            for ec in ec_list:
                ec_nm = (ec.get("name") or "").strip()
                if ec_nm and ec_nm == name:
                    matched_code = ec["code"]
                    matched_addr = (ec.get("addr") or "").strip()
                    print(f"[cust_code] 名字精確比對: {name} → {matched_code}")
                    break

        if matched_code:
            print(f"[cust_code] JSON 比對: {name} → {matched_code}")
            customer_store.update_ecount_cust_cd(user_id, matched_code)
            if db_id:
                customer_store.upsert_ecount_code(db_id, matched_code, matched_addr)
            return matched_code
    except Exception as e:
        print(f"[cust_code] JSON 比對例外: {e}")

    return None


def _create_ecount_customer(user_id: str) -> str | None:
    """
    用 Ecount API 建立新客戶（需 DB 有 real_name + phone）。
    成功回傳新代碼，失敗回傳 None。
    """
    cust = customer_store.get_by_line_id(user_id)
    if not cust:
        return None

    name    = (cust.get("real_name") or "").strip()
    phone   = (cust.get("phone") or "").strip()
    address = (cust.get("address") or "").strip()
    db_id   = cust.get("id")

    if not name or not phone:
        print(f"[cust_code] 缺少姓名或手機，無法建立 Ecount 客戶")
        return None

    from datetime import datetime as _dt
    import sqlite3 as _sqlite3
    today  = _dt.now().strftime("%y%m%d")
    prefix = f"M{today}"
    # 每日流水號從 1000 起，查 DB 已有同前綴最大號 +1
    try:
        _db_path = str(Path(__file__).parent.parent / "data" / "customers.db")
        with _sqlite3.connect(_db_path) as _conn:
            _rows = _conn.execute(
                "SELECT ecount_cust_cd FROM customers WHERE ecount_cust_cd LIKE ?",
                (f"{prefix}%",)
            ).fetchall()
        _nums = [int(cd[len(prefix):]) for (cd,) in _rows
                 if cd and cd.startswith(prefix) and cd[len(prefix):].isdigit()]
        serial = max(_nums) + 1 if _nums else 1000
    except Exception:
        serial = 1000
    new_code = f"{prefix}{serial}"

    from datetime import datetime as _dt2
    remarks = f"LINE客戶自動建立 {_dt2.now().strftime('%Y-%m-%d %H:%M')}"

    ok = ecount_client.save_customer(
        business_no=new_code,
        cust_name=name,
        hp_no=phone,
        addr=address,
        remarks=remarks,
    )
    if ok:
        print(f"[cust_code] Ecount 新建客戶: {name} → {new_code}")
        customer_store.update_ecount_cust_cd(user_id, new_code)
        if db_id:
            customer_store.upsert_ecount_code(db_id, new_code, address)
        return new_code

    print(f"[cust_code] Ecount 建立失敗: {name}")
    return None

# 數量詞清單（含箱/件批量單位）
_UNIT_PATTERN = r"(?:個|盒|組|台|條|瓶|套|份|片|包|罐|顆|粒|入|箱|件|支|隻)"

# 保留供外部引用（不再轉真人，直接接受為數量）
BULK_UNITS = ["件", "箱"]

_QTY_PATTERNS = [
    # 「各 X 個/...」→ 每款各幾個（取第一個數字）
    rf"各\s*(\d+)\s*{_UNIT_PATTERN}?",
    # 「X 個/...」
    rf"(\d+)\s*{_UNIT_PATTERN}",
    # 純數字（整段訊息就是數字，或數字在句尾）
    r"^(\d+)$",
    r"(\d+)\s*$",
]

# ── 中文數字轉換 ──────────────────────────────────
_CN_DIGIT = {
    "一": 1, "二": 2, "兩": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
}

def _cn_to_int(text: str) -> int | None:
    """
    從文字中提取中文數字並轉換為整數（支援 1–99）。
    例：「一」→1, 「兩個」→2, 「十二個」→12, 「二十五」→25
    """
    # 含「十」→ 處理 10–99
    m = re.search(r'([一二兩三四五六七八九]?)十([一二兩三四五六七八九]?)', text)
    if m:
        tens = _CN_DIGIT.get(m.group(1), 1)   # 空白 → 十 = 1*10
        ones = _CN_DIGIT.get(m.group(2), 0)   # 空白 → 個位 0
        val = tens * 10 + ones
        return val if val > 0 else None
    # 單個中文數字（一–九）＋量詞，避免「一次」「一下」「一起」等誤判
    _cn_unit = r'(?:個|箱|件|盒|組|台|條|瓶|套|份|片|包|罐|顆|粒)'
    for char, val in _CN_DIGIT.items():
        if re.search(char + _cn_unit, text):
            return val
    # 純中文數字（整段訊息就是一個中文數字，例如「三」）
    stripped = text.strip()
    if len(stripped) == 1 and stripped in _CN_DIGIT:
        return _CN_DIGIT[stripped]
    return None


def extract_quantity(text: str) -> int | None:
    """
    從訊息中提取數量（正整數）。

    支援：
    - 阿拉伯數字：「10個」「3盒」「各5個」「10」
    - 中文數字：「一個」「兩個」「十二個」「二十五」「一」
    """
    t = text.strip()
    # 1. 阿拉伯數字 pattern
    for pat in _QTY_PATTERNS:
        m = re.search(pat, t)
        if m:
            try:
                val = int(m.group(1))
                if val > 0:
                    return val
            except (ValueError, IndexError):
                pass
    # 2. 中文數字 fallback
    return _cn_to_int(t)


def _resolve_case_code(prod_cd: str) -> str | None:
    """查找箱裝版本的品項（如 Z3432 → Z3432-1）"""
    code = prod_cd.upper()
    for suffix in ["-1", "-2"]:
        c = code + suffix
        cache = ecount_client.get_product_cache_item(c)
        if cache and cache.get("unit") == "箱":
            return cache["code"]
    # 如果自己有 -1/-2，查 base code
    if "-" in code:
        base = code.rsplit("-", 1)[0]
        cache = ecount_client.get_product_cache_item(base)
        if cache and cache.get("unit") == "箱":
            return cache["code"]
    return None


def resolve_order_qty(prod_cd: str, input_qty: int) -> int:
    """
    箱/件下單數量換算：
    - 產品單位已是「箱」→ 不換算，回傳 input_qty
    - 產品單位是其他/空 → 回傳 input_qty × box_qty（裝箱數）
    - 查不到產品 or box_qty=0 → 原樣回傳，不換算
    """
    from services.ecount import ecount_client
    item = ecount_client.get_product_cache_item(prod_cd)
    if not item:
        return input_qty
    unit    = item.get("unit", "")
    box_qty = item.get("box_qty", 0)
    if unit == "箱":
        return input_qty          # 產品本身以箱計，不用乘
    if box_qty > 0:
        return input_qty * box_qty  # 5箱 × 12個/箱 = 60個
    return input_qty


def handle_order_quantity(
    user_id: str,
    text: str,
    state: dict,
    line_api: MessagingApi,
) -> str:
    """
    客戶在 awaiting_quantity 狀態下回覆數量時呼叫。

    state 需包含：
        prod_cd   — Ecount 產品編碼
        prod_name — 產品顯示名稱

    流程：
    1. 解析數量
    2. 呼叫 ecount_client.save_order()
    3. 回覆確認 or 通知真人
    """
    prod_cd = state.get("prod_cd", "")
    prod_name = state.get("prod_name") or prod_cd

    # ── 取消訂單 ──────────────────────────────────
    if any(kw in text for kw in ["不要", "算了", "取消", "不訂", "不用"]):
        cart_store.clear_cart(user_id)
        return f"好的{tone.suffix_light()} 已取消，{tone.boss()}有需要再找我哦"

    # ── 解析數量 ──────────────────────────────────
    qty = extract_quantity(text)
    if not qty:
        return tone.ask_quantity(prod_name)

    # ── 箱/件換算 ──────────────────────────────────
    _case_cd = _resolve_case_code(prod_cd)
    if any(u in text for u in BULK_UNITS):
        # 客戶明確說「箱」
        if _case_cd:
            prod_cd = _case_cd
            prod_name = ecount_client.get_product_cache_item(_case_cd).get("name", prod_name)
        else:
            qty = resolve_order_qty(prod_cd, qty)
    elif _case_cd:
        # 有箱裝版但客戶說個數 → 看能不能整除換算
        _case_item = ecount_client.get_product_cache_item(_case_cd)
        _per_box_case = 0
        import re as _re_case
        _sd = (_case_item.get("size_des", "") if _case_item else "")
        _mc = _re_case.search(r'(\d+)\s*(?:盒|個|入|包|罐|條|瓶)\s*/?\s*(?:箱|件)', _sd)
        if _mc:
            _per_box_case = int(_mc.group(1))
        if _per_box_case > 0 and qty >= _per_box_case and qty % _per_box_case == 0:
            # 能整除 → 自動換算箱裝
            _box_count = qty // _per_box_case
            prod_cd = _case_cd
            prod_name = _case_item.get("name", prod_name)
            qty = _box_count
        # 不能整除 → 用個裝品項下個數（不換算）
    else:
        # 無箱裝版，檢查產品本身是否箱裝
        from services.ecount import ecount_client as _ec_unit
        _item_unit = _ec_unit.get_product_cache_item(prod_cd)
        if _item_unit and _item_unit.get("unit") == "箱":
            # 從 Ecount SIZE_DES 或 PO 文提取裝箱數
            _per_box = 0
            import re as _re_box
            # 優先用 Ecount 規格欄（如「100個/箱」「24盒/箱」）
            _size_des = _item_unit.get("size_des", "")
            if _size_des:
                _m_box = _re_box.search(r'(\d+)\s*(?:盒|個|入|包|罐|條|瓶)\s*/?\s*(?:箱|件)', _size_des)
                if _m_box:
                    _per_box = int(_m_box.group(1))
            # 沒有就從 PO 文找
            if not _per_box:
                try:
                    from handlers.internal import _get_raw_po_block
                    _po_box = _get_raw_po_block(prod_cd) or ""
                    _m_box = _re_box.search(r'(\d+)\s*(?:盒|個|入|包|罐|條|瓶)\s*/?\s*(?:箱|件)', _po_box)
                    if not _m_box:
                        _m_box = _re_box.search(r'(?:1?\s*(?:箱|件))\s*(\d+)\s*(?:盒|個|入|包)', _po_box)
                    if not _m_box:
                        _m_box = _re_box.search(r'(\d+)\s*(?:盒|個|入|包)起批', _po_box)
                    if _m_box:
                        _per_box = int(_m_box.group(1))
                except Exception:
                    pass

            # 客戶說了非箱單位（盒/個/包等）→ 嘗試自動換算
            _said_unit = any(u in text for u in ["盒", "個", "入", "包", "罐", "條", "瓶"])
            if _said_unit and _per_box > 1 and qty % _per_box == 0:
                # 能整除 → 自動換算（如 100盒 ÷ 100 = 1箱）
                _box_count = qty // _per_box
                from storage import cart as _cart_auto_box
                _cart_auto_box.add_item(user_id, prod_cd, prod_name, _box_count)
                return f"好的，{qty}盒 = {_box_count}箱\n" + tone.cart_item_added(_cart_auto_box.get_cart(user_id))
            elif _said_unit and _per_box > 1:
                # 不能整除 → 提醒
                return f"這款 1箱={_per_box}盒，{qty}盒無法整除。請問要幾箱呢？"

            # 純數字沒單位 → 問
            from storage.state import state_manager as _sm_unit
            _sm_unit.set(user_id, {
                "action": "awaiting_box_confirm",
                "prod_cd": prod_cd,
                "prod_name": prod_name,
                "input_qty": qty,
            })
            if _per_box > 1:
                return f"這款「{prod_name}」是以箱為單位（1箱={_per_box}盒），請問 {qty} 是幾箱呢？"
            return f"這款「{prod_name}」是以箱為單位喔，請問 {qty} 是幾箱呢？"

    # ── 加入購物車 ────────────────────────────────
    cart = cart_store.add_item(user_id, prod_cd, prod_name, qty)
    print(f"[ordering] 加入購物車: {user_id} | {prod_name} x{qty} | 共 {len(cart)} 項")
    return tone.cart_item_added(cart)


def handle_checkout(
    user_id: str,
    line_api: MessagingApi,
) -> str:
    """
    客戶說「好了」時，將購物車內所有品項一次送出 Ecount 訂貨單。
    """
    cart = cart_store.get_cart(user_id)
    if not cart:
        return tone.cart_empty_checkout()

    # ── 取得客戶所有 Ecount 代碼（含地址）──────────
    codes = customer_store.get_ecount_codes_by_line_id(user_id)

    if len(codes) > 1:
        from storage.state import state_manager as _sm
        # 群組預設地址：詢問「是否送到 XX？」而非列出全部選項
        preferred = _sm.get_group_cust_cd(user_id)
        if preferred:
            pref_label = next(
                (c.get("address_label") or c.get("cust_name") or preferred
                 for c in codes if c["ecount_cust_cd"] == preferred),
                preferred
            )
            _sm.set(user_id, {"action": "awaiting_group_address_confirm"})
            return tone.ask_group_address_confirm(pref_label)
        # 一般多地址：列出全部選項（購物車不清）
        _sm.set(user_id, {"action": "awaiting_address_selection_checkout"})
        return tone.ask_address_selection(codes)

    # ── 決定客戶代碼 ──────────────────────────────
    cust_code = None
    if codes:
        cust_code = codes[0]["ecount_cust_cd"]
    else:
        existing = customer_store.get_ecount_cust_code(user_id, default="")
        if existing:
            cust_code = existing
        else:
            cust_info  = customer_store.get_by_line_id(user_id)
            cust_name  = (cust_info.get("real_name") or "").strip() if cust_info else ""
            cust_phone = (cust_info.get("phone") or "").strip() if cust_info else ""

            if cust_name and cust_phone:
                # 有姓名+電話 → 先 JSON 比對，失敗就 API 建立
                cust_code = _resolve_cust_code(user_id, do_refresh=True)
                if not cust_code:
                    cust_code = _create_ecount_customer(user_id)

            if not cust_code:
                # DB 真的沒有姓名或電話 → 才詢問客戶聯絡資訊
                from storage.state import state_manager
                state_manager.set(user_id, {"action": "awaiting_contact_info_checkout"})
                return tone.ask_contact_info()

    # ── 一次送出所有品項 ──────────────────────────
    from storage.customers import customer_store as _cs_ord
    _phone = (_cs_ord.get_by_line_id(user_id) or {}).get("phone", "") or ""
    items = [{"prod_cd": i["prod_cd"], "qty": i["qty"]} for i in cart]
    slip_no = ecount_client.save_order(cust_code=cust_code, items=items, phone=_phone)

    if slip_no:
        print(f"[ordering] 購物車訂單建立成功: {slip_no} | {cust_code} | {len(cart)} 項")
        # 預購品自動登記到貨通知
        from handlers.inventory import _check_preorder, notify_hq_restock
        from storage.notify import notify_store
        _oos_items = []  # 缺貨品項（需通知總公司調貨）
        _po_items = []   # 預購品項
        for item in cart:
            if _check_preorder(item["prod_cd"]):
                _po_items.append(item)
                notify_store.add(
                    user_id=user_id,
                    prod_code=item["prod_cd"],
                    prod_name=item["prod_name"],
                    source="staff",
                    qty_wanted=item["qty"],
                )
                print(f"[ordering] 預購品自動登記到貨通知: {item['prod_name']} x{item['qty']}")
            else:
                # 非預購品 → 檢查庫存是否足夠
                _item_info = ecount_client.lookup(item["prod_cd"])
                _item_qty = _item_info.get("qty") if _item_info else None
                if not _item_qty or _item_qty < item["qty"]:
                    # 庫存不足（包含 0 和不夠的情況）
                    _short = item["qty"] - (_item_qty or 0)
                    _oos_items.append({**item, "short": _short, "stock": _item_qty or 0})
        # 缺貨品項 → 一次通知總公司 + 一筆待處理
        if _oos_items:
            from storage.issues import issue_store
            _oos_desc = "、".join(
                f"{i['prod_name']}（{i['prod_cd']}）×{i['qty']}" for i in _oos_items)
            issue_store.add(user_id, "restock_inquiry", f"缺貨調貨：{_oos_desc}")
            _notify_hq_restock_batch(_oos_items, line_api)
            print(f"[ordering] 缺貨品批次通知總公司: {_oos_desc}")
        cart_store.clear_cart(user_id)
        return tone.checkout_confirmed(cart, oos_items=_oos_items, po_items=_po_items)
    else:
        print(f"[ordering] 購物車訂單建立失敗: {cust_code}")
        from storage.issues import issue_store
        desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart)
        issue_store.add(user_id, "order_failed", desc)
        return "⚠️ 訂單處理時發生問題，請稍後再試或聯繫客服。"


def _notify_hq_restock_batch(oos_items: list[dict], line_api) -> None:
    """一次通知總公司群組所有缺貨品項"""
    from config import settings
    from linebot.v3.messaging import PushMessageRequest, TextMessage

    if not line_api or not settings.LINE_GROUP_ID_HQ:
        print(f"[總公司通知] 未設定 LINE_GROUP_ID_HQ，跳過")
        return

    lines = ["⚠️ 客戶已下單，以下品項庫存不足，麻煩盡快確認調貨："]
    for item in oos_items:
        short = item.get("short", item["qty"])
        lines.append(f"📦 {item['prod_name']}（{item['prod_cd']}）需調 {short} 個")
    try:
        line_api.push_message(
            PushMessageRequest(
                to=settings.LINE_GROUP_ID_HQ,
                messages=[TextMessage(text="\n".join(lines))],
            )
        )
    except Exception as e:
        print(f"[總公司通知] 批次推送失敗: {e}")


def _notify_staff(
    user_id: str,
    prod_name: str,
    qty: int,
    cust_code: str,
    line_api: MessagingApi | None,
):
    """訂單失敗時通知真人"""
    from config import settings
    from linebot.v3.messaging import PushMessageRequest, TextMessage

    if not line_api or not settings.LINE_GROUP_ID:
        print(f"[ordering] 無群組 ID，僅 log：{cust_code} 要訂 {prod_name} x{qty}")
        return
    try:
        msg = (
            f"⚠️ 訂單需要人工處理\n"
            f"客戶：{cust_code}\n"
            f"LINE ID：{user_id}\n"
            f"商品：{prod_name}\n"
            f"數量：{qty} 個\n"
            f"請確認後手動建立訂單"
        )
        line_api.push_message(
            PushMessageRequest(
                to=settings.LINE_GROUP_ID,
                messages=[TextMessage(text=msg)],
            )
        )
    except Exception as e:
        print(f"[ordering] 通知失敗: {e}")
