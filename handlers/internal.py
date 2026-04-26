"""
內部群組指令處理

1. 到貨通知：「T1202 到貨」→ push 等待通知的客戶
2. 幫訂單：  「幫 張三 訂 BB-232 3個」→ 建立 Ecount 訂單
3. 圖片識別：（在 on_image_message 呼叫）→ 識別產品 + 回 PO文
4. 通知登記：「T1202 通知 張三 3個」→ 手動幫客戶登記到貨通知
5. 庫存查詢：「K0236 庫存」「K0236 有多少」→ 查 Ecount 回覆
6. 上架：傳圖/影片 + PO文 → 儲存照片 + 更新產品PO文.txt
"""

import json as _json
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

# ── Ecount 客戶快取 ────────────────────────────────────────────────────
_EC_PATH = Path(__file__).parent.parent / "data" / "ecount_customers.json"
_ec_customers_cache: list[dict] | None = None
_ec_customers_mtime: float = 0

def _load_ec_customers() -> list[dict]:
    """Load ecount_customers.json with file-mtime cache"""
    global _ec_customers_cache, _ec_customers_mtime
    try:
        mtime = _EC_PATH.stat().st_mtime
        if _ec_customers_cache is not None and mtime == _ec_customers_mtime:
            return _ec_customers_cache
        _ec_customers_cache = _json.loads(_EC_PATH.read_text(encoding="utf-8"))
        _ec_customers_mtime = mtime
        return _ec_customers_cache
    except Exception:
        return _ec_customers_cache or []

def _resolve_customer(name: str) -> dict | None:
    """Resolve customer by name: exact match -> partial match -> None"""
    ec_list = _load_ec_customers()
    clean = name.strip()
    # Exact match
    match = next((x for x in ec_list if x.get("name", "").strip() == clean), None)
    if match:
        return match
    # Partial match
    match = next((x for x in ec_list if clean in x.get("name", "")), None)
    return match

# ── 共用產品代碼 pattern ───────────────────────────────────────────────
_PROD_CODE_PAT = r'[A-Za-z]{1,3}-?\d{3,6}(?:-[A-Za-z0-9]+)*'

# ── 正則 ──────────────────────────────────────────────────────────────
# 商品編號：英文1~3碼（可含 -）+ 數字3~6碼 + 可選後綴（-J-23、-1、-A2 等），例：T1202、Z3323-J-23
# 排除常見非貨號（PD=充電協議、USB、LED、MAX等）
_NOT_PROD_CODE = {"PD", "USB", "LED", "MAX", "MAH", "LCD", "RGB", "GPS", "SOS", "DIY", "ABS", "TPU", "BTS"}
# 排除型號/品牌標示：2-3 字母 + dash + 恰好 3 數字（且無後綴），例：TP-650、VPB-011
# 房號實際格式為 1 字母 + 4 數字（Z3432）或 2-3 字母 + 4+ 數字（NN249），不會用此短格式
_BRAND_MODEL_RE = re.compile(r'^[A-Za-z]{2,3}-\d{3}$')
_PROD_CODE_RE_RAW = re.compile(rf'(?<![A-Za-z\-])({_PROD_CODE_PAT})(?!\d)')
def _is_excluded_code(code: str) -> bool:
    if code[:2].upper() in _NOT_PROD_CODE or code[:3].upper() in _NOT_PROD_CODE:
        return True
    if _BRAND_MODEL_RE.match(code):
        return True
    return False
class _ProdCodeFinder:
    """findall/search 時自動排除非貨號"""
    def findall(self, text):
        return [m for m in _PROD_CODE_RE_RAW.findall(text) if not _is_excluded_code(m)]
    def search(self, text):
        for m in _PROD_CODE_RE_RAW.finditer(text):
            if not _is_excluded_code(m.group(1)):
                return m
        return None
    def finditer(self, text):
        for m in _PROD_CODE_RE_RAW.finditer(text):
            if not _is_excluded_code(m.group(1)):
                yield m
    def sub(self, repl, text):
        def _repl_fn(m):
            if not _is_excluded_code(m.group(1)):
                return repl if isinstance(repl, str) else repl(m)
            return m.group(0)
        return _PROD_CODE_RE_RAW.sub(_repl_fn, text)
_PROD_CODE_RE = _ProdCodeFinder()

# 到貨觸發詞
_ARRIVAL_KW = ["到貨", "到了", "到齊", "收到了", "進來了", "到貨了", "貨到了", "貨到"]

# 格式A（每行一筆）：「張三 訂 T1202 3」
# group(1)=姓名  group(2)=產品代碼  group(3)=數量
_STAFF_ORDER_LINE_RE = re.compile(
    rf'(.+?)\s+(?:訂|下單)\s+({_PROD_CODE_PAT})\s+([零一二三四五六七八九十百千\d]+)\s*(個|件|盒|套|箱|組)?\s*(.*)'
)
# group(5) = 尾段備註（如「不要黑色」），可為空
# 格式B 第一行：「張三訂」或「張三 訂」（後面沒有產品代碼）
# group(1)=姓名
_STAFF_ORDER_HEADER_RE = re.compile(
    r'^(.+?)\s*(?:訂|下單)\s*$'
)
# 格式B 後續每行：「T1202 3個」
# group(1)=產品代碼  group(2)=數量  group(3)=單位（可為空）
# 支援空白、*、× 作為分隔符，例：Z3598 1、Z3598*1、Z3598×1
_STAFF_ORDER_ITEM_RE = re.compile(
    rf'({_PROD_CODE_PAT})[\s×\*xX]+([零一二三四五六七八九十百千\d]+)\s*(個|件|盒|套|箱|組)?'
)
# 格式C（無需「訂」關鍵字）：「姓名 產品代碼 數量個 [備註]」，例：方力緯 Z3562 5個 不要黑色
# group(1)=姓名  group(2)=產品代碼  group(3)=數量  group(4)=單位  group(5)=尾段備註（可為空）
_STAFF_ORDER_DIRECT_RE = re.compile(
    rf'^(.+?)\s+({_PROD_CODE_PAT})[\s×\*xX]+([零一二三四五六七八九十百千\d]+)\s*(個|件|盒|套|箱|組)?\s*(.*?)$'
)
_BULK_UNITS = {"件", "箱"}

# 通知登記觸發詞：句首「通知登記」OR 句中/句尾含以下關鍵字
_NOTIFY_REG_START_RE = re.compile(r'^通知登記')
_NOTIFY_REG_INLINE_KW = ["需要到貨通知", "要到貨通知", "通知登記", "需要通知", "要通知", "登記通知", "要登記", "需要登記"]
# 注意：「到貨通知」不在此清單，它是「到貨觸發」而非登記指令
# 格式：「通知/登記 [姓名] 產品代碼」
# group(1)=可選姓名  group(2)=產品代碼
# 例：「通知 T1202」、「通知 張三 T1202」、「登記 T1202」、「登記 張三 T1202」
_NOTIFY_REG_SHORTHAND_RE = re.compile(
    r'^(?:通知|登記)\s+(?:(.+?)\s+)?([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)', re.IGNORECASE
)
# 每一行：「姓名  產品代碼  [數量]」
_NOTIFY_REG_LINE_RE  = re.compile(
    r'(.+?)\s+([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)(?:\s+([零一二三四五六七八九十百千\d]+)\s*(個|件|盒|套|箱|組)?)?'
)

# 品名下單 token（合體格式）：「衛生紙30箱」「泡澡球10件」
_ITEM_TOKEN_RE = re.compile(
    r'^([\u4e00-\u9fff\w]+?)(\d+)\s*(個|件|盒|套|箱|組)?$'
)
# 純數量 token（分離格式）：「30箱」「10件」「5個」
_QTY_ONLY_RE = re.compile(
    r'^(\d+)\s*(個|件|盒|套|箱|組)$'
)

# 純貨號偵測（整行只有貨號，含字母或純數字格式）
_CODE_ONLY_RE = re.compile(
    rf'^\s*({_PROD_CODE_PAT}|\d{{5,6}}(?:-\d+)?)\s*$',
    re.IGNORECASE
)

# 到貨批量格式：「到貨通知\n{code}\n{name1}\n{name2}」
_ARRIVAL_BATCH_RE = re.compile(
    rf'^到貨通知\s*\n\s*({_PROD_CODE_PAT})\s*\n(.+)',
    re.DOTALL | re.IGNORECASE,
)

# OCR 候選詞過濾：貨號格式(字母+數字) 或中文詞
_CODE_OR_ZH = re.compile(r'(?:[A-Za-z]\d{2,}|[\u4e00-\u9fff]{2,})')

# 登記通知：純貨號格式（用於判斷第二行是貨號還是客戶名）
_NOTIFY_PROD_CODE_PAT = re.compile(rf'^{_PROD_CODE_PAT}$', re.IGNORECASE)

# 圖片代訂：尾部數量
_QTY_TAIL_RE = re.compile(r'(\d+)\s*(?:個|件|盒|套|箱|組)?\s*$')

# 圖片代訂：動詞分隔
_VERB_SEP_RE = re.compile(r'\s+(?:要|訂|下單|買)\s+')

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

_SPEC_FORMAT_RE = re.compile(
    r'(?:產品名稱|品名|商品名稱)[：:].+|(?:編號|貨號)[：:]\s*[A-Za-z0-9\-]+',
    re.IGNORECASE
)

# 規格訊息是否足夠完整（含編號 + 至少一個其他欄位）
_SPEC_HAS_CODE_RE  = re.compile(r'(?:編號|貨號)[：:]\s*([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)', re.IGNORECASE)
_SPEC_HAS_EXTRA_RE = re.compile(r'(?:裝箱量|装箱量|包裝|包装|重量|价格|價格|批量|預購)', re.IGNORECASE)


def handle_internal_spec_inquiry(text: str, group_id: str) -> str | None:
    """
    偵測內部群傳入產品規格訊息（含編號＋其他欄位），
    詢問「要訂貨還是查庫存？」，並把原始規格文字存入 state 等待回覆。
    """
    if not _SPEC_HAS_CODE_RE.search(text):
        return None
    if not _SPEC_HAS_EXTRA_RE.search(text):
        return None
    # 提取貨號
    m = _SPEC_HAS_CODE_RE.search(text)
    prod_code = m.group(1).upper()
    # 儲存狀態，等待回覆
    from storage.state import state_manager
    state_manager.set(group_id, {
        "action":    "spec_inquiry_pending",
        "prod_code": prod_code,
        "spec_text": text,
    })
    return f"收到 {prod_code} 的資料 📋\n請問要訂貨還是查庫存？"


def handle_spec_inquiry_reply(group_id: str, text: str, line_api) -> str | None:
    """
    處理「規格詢問」狀態下的回覆（訂貨 / 查庫存）。
    """
    from storage.state import state_manager
    st = state_manager.get(group_id) or {}
    if st.get("action") != "spec_inquiry_pending":
        return None

    prod_code = st.get("prod_code", "")
    spec_text = st.get("spec_text", "")
    state_manager.clear(group_id)

    t = text.strip()

    # 查庫存
    if any(kw in t for kw in ["庫存", "查", "有多少", "幾個", "幾箱"]):
        from handlers.inventory import _query_single_product
        result = _query_single_product("internal", prod_code)
        return result or f"查不到 {prod_code} 的庫存資料"

    # 訂貨
    if any(kw in t for kw in ["訂貨", "訂", "下單", "要"]):
        state_manager.set(group_id, {
            "action":    "spec_inquiry_order_qty",
            "prod_code": prod_code,
            "spec_text": spec_text,
        })
        from services.ecount import ecount_client
        item = ecount_client.lookup(prod_code)
        name = item["name"] if item else prod_code
        return f"好的，請問 {name}（{prod_code}）要訂幾個？"

    # 無法識別
    state_manager.set(group_id, {
        "action":    "spec_inquiry_pending",
        "prod_code": prod_code,
        "spec_text": spec_text,
    })
    return "請回覆「訂貨」或「查庫存」"


def handle_spec_inquiry_qty(group_id: str, text: str, line_api) -> str | None:
    """
    處理「規格詢問訂貨數量」狀態下的回覆。
    """
    from storage.state import state_manager
    st = state_manager.get(group_id) or {}
    if st.get("action") != "spec_inquiry_order_qty":
        return None

    prod_code = st.get("prod_code", "")
    state_manager.clear(group_id)

    from handlers.ordering import extract_quantity, resolve_unit
    from services.ecount import ecount_client
    import re as _re_unit

    qty = extract_quantity(text)
    if not qty:
        return f"請輸入數量，例如「10個」或「5箱」"

    item = ecount_client.lookup(prod_code)
    if not item:
        return f"❌ 找不到產品 {prod_code}"

    _um = _re_unit.search(r'(\d+)\s*(個|盒|組|台|條|瓶|套|份|片|包|罐|入|箱|件|支|隻)', text)
    _said_unit = _um.group(2) if _um else None
    prod_code, actual_qty, _warn = resolve_unit(prod_code, qty, _said_unit)
    if _warn and _warn.startswith("⚠️"):
        return _warn
    slip_no = ecount_client.save_order(
        cust_code="INTERNAL",
        items=[{"prod_cd": prod_code, "qty": actual_qty}],
    )
    name = item["name"] or prod_code
    if slip_no:
        _u = (item.get("unit") or "個")
        return f"✅ 已建立訂單 {slip_no}\n  {name}（{prod_code}）× {actual_qty} {_u}"
    else:
        return f"❌ {name}（{prod_code}）訂單建立失敗"

def _push_arrival_msg(uid: str, prod_name: str, prod_code: str, qty: int, source: str = "staff", line_api=None) -> bool:
    """push 到貨通知給單一客戶，根據 source 用不同模板"""
    if source == "staff":
        # 內部群登記：訂購到貨格式
        item = ecount_client.get_product_cache_item(prod_code)
        prod_unit = (item.get("unit") or "") if item else ""
        box_qty = (item.get("box_qty") or 0) if item else 0
        if prod_unit == "箱":
            qty_display = f"{qty}箱"
        elif box_qty > 1 and qty >= box_qty and qty % box_qty == 0:
            qty_display = f"{qty // box_qty}箱"
        else:
            qty_display = f"{qty}個"
        msg = (
            f"老闆您好，您之前訂的貨已經到了\n"
            f"{prod_name}（{prod_code}）× {qty_display}"
        )
    else:
        # 客戶自己登記：到貨通知格式
        from handlers.tone import restock_back_in_stock
        msg = restock_back_in_stock(name=prod_name, code=prod_code)
    try:
        line_api.push_message(
            PushMessageRequest(to=uid, messages=[TextMessage(text=msg)])
        )
        return True
    except Exception as e:
        print(f"[internal] 推播失敗 {uid}: {e}")
        return False


def handle_internal_arrival(text: str, line_api: MessagingApi) -> str | None:
    """
    到貨觸發，push 等待該產品的 staff 登記客戶。

    格式 1（一行，單/多貨號）：
        到貨 Z1234
        Z3568 Z7859-1 到貨

    格式 2（多行，指定客戶）：
        到貨通知
        Z4658
        張三
        李四
    """
    # 含產品規格格式（產品名稱：/編號：）→ 是新品資料，不是到貨通知
    if _SPEC_FORMAT_RE.search(text):
        return None

    # 含登記關鍵字 → 是登記指令，不是到貨觸發
    if any(kw in text for kw in _NOTIFY_REG_INLINE_KW):
        return None

    t = text.strip()

    # ── 格式 2：「到貨通知\n{code}\n{name1}\n{name2}」─────────────────────
    m_batch = _ARRIVAL_BATCH_RE.match(t)
    if m_batch:
        prod_code  = m_batch.group(1).upper()
        item       = ecount_client.lookup(prod_code)
        prod_name  = (item["name"] if item else "") or prod_code
        name_lines = [l.strip() for l in m_batch.group(2).splitlines() if l.strip()]

        results = []
        for name in name_lines:
            cust_name_q = re.sub(r'\s+', '', name)
            matches = customer_store.search_by_name(cust_name_q, real_name_only=True)
            if not matches:
                results.append(f"❌ 找不到「{cust_name_q}」")
                continue
            cust = matches[0]
            uid  = cust.get("line_user_id", "")
            lbl  = cust.get("real_name") or cust.get("display_name") or cust_name_q
            if not uid:
                results.append(f"⚠️ 「{lbl}」無 LINE ID")
                continue
            # 取登記數量（有的話），否則預設 1
            pending_rows = notify_store.get_pending_by_code(prod_code, source="staff")
            entry = next((r for r in pending_rows if r["user_id"] == uid), None)
            qty   = entry["qty_wanted"] if entry else 1
            source = entry.get("source", "staff") if entry else "staff"
            ok    = _push_arrival_msg(uid, prod_name, prod_code, qty, source, line_api)
            if ok:
                if entry:
                    notify_store.mark_notified(entry["id"])
                results.append(f"📨 {lbl}｜{prod_name}（{prod_code}）× {qty}")
            else:
                results.append(f"❌ 推播失敗：{lbl}")
        return "\n".join(results) if results else None

    # ── 格式 1：一行含到貨關鍵字 + 貨號 ────────────────────────────────────
    has_arrival_kw = any(kw in t for kw in _ARRIVAL_KW)
    if not has_arrival_kw:
        return None

    codes = _PROD_CODE_RE.findall(t)
    if not codes:
        return None

    results = []
    for raw_code in codes:
        prod_code = raw_code.upper()
        # 取全部 pending（staff + customer）
        all_pending = notify_store.get_pending_by_code(prod_code, source="staff")
        all_pending += notify_store.get_pending_by_code(prod_code, source="customer")
        if not all_pending:
            results.append(f"📦 {prod_code}：沒有待通知的客戶")
            continue

        item      = ecount_client.lookup(prod_code)
        prod_name = (item["name"] if item else "") or prod_code
        notified  = 0
        for entry in all_pending:
            uid = entry["user_id"]
            qty = entry["qty_wanted"]
            source = entry.get("source", "customer")
            ok = _push_arrival_msg(uid, prod_name, prod_code, qty, source, line_api)
            if ok:
                notify_store.mark_notified(entry["id"])
                notified += 1
                print(f"[internal] 到貨通知({source}): {prod_code} → {uid}")

        results.append(f"📦 {prod_code}：已通知 {notified} 位客戶")

    return "\n".join(results) if results else None


# ── 2. 幫訂單 ─────────────────────────────────────────────────────────

def _do_order(
    cust_name_query: str,
    items_raw: list[tuple[str, int]],
    units: dict[str, str] | None = None,   # prod_cd → 單位（箱/件/個…）
    note: str = "",                         # 備註，放每個品項的 REMARK
    group_id: str | None = None,            # 內部群 ID，用於存待確認狀態
) -> str:
    """
    共用下單邏輯。items_raw = [(prod_query, qty), ...]
    優先從 ecount_customers.json 查客戶，找不到再查 LINE 本地 DB。
    units 選填：{prod_cd: "箱"} 可讓訂單訊息顯示正確單位。
    回傳結果文字。
    """
    from pathlib import Path as _Path

    cust_code   = ""
    cust_label  = cust_name_query
    _phone      = ""
    is_new_cust = False

    # 去除括號（「陳怡如(彥鈞)」→「陳怡如」）以利查詢
    cust_name_clean = re.sub(r'[\(（][^\)）]*[\)）]', '', cust_name_query).strip() or cust_name_query

    # 1. 先查 Ecount 客戶清單（原始名 → 去括號名，均精確比對）
    ec_match = _resolve_customer(cust_name_query)
    if not ec_match and cust_name_clean != cust_name_query:
        ec_match = _resolve_customer(cust_name_clean)
    if ec_match:
        cust_code  = ec_match.get("code", "")
        cust_label = ec_match.get("name", cust_name_query)
        _phone     = ec_match.get("phone", "") or ec_match.get("tel", "") or ""
        print(f"[internal] Ecount 客戶: {cust_label} → {cust_code}", flush=True)

    # 2. 找不到 → fallback 查 LINE 本地 DB（精確比對，不模糊）
    if not cust_code:
        # 先用去括號的乾淨名稱查，再用原始名查
        cust_matches = customer_store.search_by_name(cust_name_clean, real_name_only=True)
        if not cust_matches and cust_name_clean != cust_name_query:
            cust_matches = customer_store.search_by_name(cust_name_query, real_name_only=True)
        if not cust_matches:
            # 找不到 → 詢問確認，存入 state 等待回覆「是」
            from storage.state import state_manager as _sm
            if group_id:
                _sm.set(group_id, {
                    "action": "confirm_new_customer",
                    "cust_name": cust_name_clean,
                    "items_raw": items_raw,
                    "units": units or {},
                    "note": note,
                })
                items_desc = "、".join(f"{t[0]}×{t[1]}" for t in items_raw)
                print(f"[internal] 找不到客戶「{cust_name_clean}」，等待確認", flush=True)
                return (
                    f"⚠️ 找不到客戶「{cust_name_clean}」\n"
                    f"訂單內容：{items_desc}\n\n"
                    f"是今天新客人嗎？\n"
                    f"回覆「是」→ 同步客戶資料後建單\n"
                    f"回覆「不是」→ 建立新客人後建單"
                )
            else:
                # 無 group_id fallback：直接用預設代碼
                cust_code   = settings.ECOUNT_DEFAULT_CUST_CD
                cust_label  = cust_name_clean
                is_new_cust = False
                print(f"[internal] 找不到客戶「{cust_name_clean}」，無 group_id，用預設代碼", flush=True)
        elif len(cust_matches) > 1:
            names = "、".join(c.get("real_name") or c.get("chat_label") or c.get("display_name", "?") for c in cust_matches[:5])
            return f"⚠️ 「{cust_name_query}」有多位：{names}"
        else:
            cust       = cust_matches[0]
            user_id    = cust["line_user_id"]
            cust_label = cust.get("real_name") or cust.get("chat_label") or cust.get("display_name") or cust_name_query
            if user_id:
                codes  = customer_store.get_ecount_codes_by_line_id(user_id)
                if codes:
                    cust_code = codes[0]["ecount_cust_cd"]
                else:
                    existing  = customer_store.get_ecount_cust_code(user_id, default="")
                    cust_code = existing or _resolve_cust_code(user_id) or settings.ECOUNT_DEFAULT_CUST_CD
                _phone = (customer_store.get_by_line_id(user_id) or {}).get("phone", "") or ""
            else:
                # line_user_id 空白（僅有 chat_label）：用姓名在 Ecount 再查一次
                ec_name = cust.get("real_name") or cust.get("chat_label", "").split("-")[0].strip()
                if ec_name:
                    ec_match2 = _resolve_customer(ec_name)
                    if ec_match2:
                        cust_code  = ec_match2.get("code", "")
                        cust_label = ec_match2.get("name", ec_name)
                        _phone     = ec_match2.get("phone", "") or ""
                if not cust_code:
                    cust_code = settings.ECOUNT_DEFAULT_CUST_CD

    # 3. 查詢產品，組成 items
    order_items = []
    for entry in items_raw:
        if len(entry) >= 3:
            prod_query, qty, item_note = entry[0], entry[1], entry[2]
        else:
            prod_query, qty = entry[0], entry[1]
            item_note = note  # 全域備註 fallback
        item = ecount_client.lookup(prod_query)
        if not item:
            return f"❌ 找不到產品「{prod_query}」"
        order_items.append({
            "prod_cd":   item["code"],
            "prod_name": item["name"] or item["code"],
            "qty":       qty,
            "note":      item_note,
        })

    # 4. 建立訂單
    slip_no = ecount_client.save_order(
        cust_code=cust_code,
        items=[{"prod_cd": i["prod_cd"], "qty": i["qty"], "note": i.get("note", "")} for i in order_items],
        phone=_phone,
    )

    new_tag = "（新建客戶）" if is_new_cust else ""
    if slip_no:
        detail = "、".join(f"{i['prod_name']}×{i['qty']}" for i in order_items)
        print(f"[internal] 代訂成功: {slip_no} | {cust_label}{new_tag} | {detail}")
        lines_out = [f"✅ {cust_label}{new_tag}｜{slip_no}"]
        for i in order_items:
            unit = (units or {}).get(i["prod_cd"], "")
            if not unit:
                _item_unit = ecount_client.get_product_cache_item(i["prod_cd"])
                unit = (_item_unit.get("unit") if _item_unit else "") or "個"
            note_str = f"（{i['note']}）" if i.get("note") else ""
            lines_out.append(f"  {i['prod_name']}（{i['prod_cd']}）× {i['qty']} {unit}{note_str}")
        # 同時建立到貨通知登記（直接用已知的 cust_code/user_id，避免重名查詢問題）
        _notify_ok = False
        try:
            # 優先用 LINE user_id；沒有則用 ecount: 前綴
            _notify_uid = None
            if cust_code:
                # 從 customers.db 找 LINE ID
                _all_custs = customer_store.search_by_name(cust_label, real_name_only=True)
                _with_uid = [c for c in _all_custs if c.get("line_user_id")]
                if _with_uid:
                    _notify_uid = _with_uid[0]["line_user_id"]
            if not _notify_uid:
                _notify_uid = f"ecount:{cust_label}"
            for i in order_items:
                notify_store.add(
                    user_id=_notify_uid, prod_code=i["prod_cd"],
                    prod_name=i["prod_name"], qty_wanted=i["qty"],
                    source="staff",
                )
            _notify_ok = True
            _tag = "（到貨通知內部群）" if _notify_uid.startswith("ecount:") else ""
            print(f"[internal] 到貨通知已登記: {cust_label}{_tag} | {detail}")
        except Exception as _ne:
            print(f"[internal] 到貨通知登記失敗: {_ne}")
        if _notify_ok:
            lines_out.append("📬 已登記到貨通知")
        return "\n".join(lines_out)
    else:
        detail = "、".join(f"{i['prod_name']}×{i['qty']}" for i in order_items)
        print(f"[internal] 代訂失敗: {cust_code} | {detail}")
        return f"❌ {cust_label}{new_tag}｜訂單建立失敗"


# ── 內部群：客戶購物車管理 ─────────────────────────────────────
# 「購物車 賴柏舟」→ 進入 session：查看 + 後續可直接 T0101*6 加購 / 送出
# 「加購 賴柏舟 S0633*24」→ 不用 session 也能加購
# 「代結帳 賴柏舟」→ 不用 session 也能結帳
_CART_VIEW_RE = re.compile(r"^購物車\s*(\S+)$")
_CART_VIEW_RE_REVERSE = re.compile(r"^(\S+?)\s*購物車$")
_CART_ADD_RE = re.compile(r"^加購\s*(\S+)\s+([A-Za-z]{1,3}-?\d{3,6})\s*[*×xX]\s*(\d+)$")
_CART_CHECKOUT_RE = re.compile(r"^代結帳\s*(\S+)$")
# session 中的快速指令
_CART_QUICK_ADD_RE = re.compile(r"^(?:加購\s*)?([A-Za-z]{1,3}-?\d{3,6})\s*[*×xX]\s*(\d+)$")
_CART_QUICK_SUBMIT = {"送出", "結帳", "好了", "確認"}

# 購物車管理 session（per 內部群使用者）
import threading as _cart_threading
_cart_session: dict[str, dict] = {}  # staff_user_id → {"line_id": ..., "label": ..., "ts": ...}
_cart_session_lock = _cart_threading.Lock()
_CART_SESSION_TIMEOUT = 300  # 5 分鐘無操作自動結束


def _get_cart_session(staff_id: str) -> dict | None:
    """取得有效的購物車 session"""
    import time
    with _cart_session_lock:
        s = _cart_session.get(staff_id)
        if s and time.time() - s["ts"] < _CART_SESSION_TIMEOUT:
            return s
        _cart_session.pop(staff_id, None)
    return None


def _set_cart_session(staff_id: str, line_id: str, label: str) -> None:
    import time
    with _cart_session_lock:
        _cart_session[staff_id] = {"line_id": line_id, "label": label, "ts": time.time()}


def _touch_cart_session(staff_id: str) -> None:
    import time
    with _cart_session_lock:
        if staff_id in _cart_session:
            _cart_session[staff_id]["ts"] = time.time()


def _clear_cart_session(staff_id: str) -> None:
    with _cart_session_lock:
        _cart_session.pop(staff_id, None)


def _resolve_customer_line_id(name: str) -> tuple[str | None, str]:
    """用客戶名找 line_user_id，回傳 (line_id, display_label)"""
    from storage.customers import customer_store
    results = customer_store.search_by_name(name, real_name_only=True)
    if not results:
        results = customer_store.search_by_name(name)
    if not results:
        return None, name
    cust = results[0]
    line_id = cust.get("line_user_id", "")
    label = cust.get("real_name") or cust.get("display_name") or name
    return line_id, label


def _format_cart(label: str, cart: list[dict]) -> str:
    """格式化購物車顯示"""
    lines = [f"📋 {label} 的購物車："]
    for item in cart:
        lines.append(f"  • {item['prod_name']}（{item['prod_cd']}）× {item['qty']}")
    lines.append(f"\n直接輸入「貨號*數量」加購，「送出」結帳")
    return "\n".join(lines)


def handle_internal_cart(text: str, line_api=None, staff_id: str = "") -> str | None:
    """處理內部群購物車管理指令"""
    from storage import cart as cart_store
    from services.ecount import ecount_client
    t = text.strip()

    # ── 查看購物車（進入 session）──
    m = _CART_VIEW_RE.match(t) or _CART_VIEW_RE_REVERSE.match(t)
    if m:
        name = m.group(1)
        line_id, label = _resolve_customer_line_id(name)
        if not line_id:
            return f"❌ 找不到客戶「{name}」"
        _set_cart_session(staff_id or "default", line_id, label)
        cart = cart_store.get_cart(line_id)
        if not cart:
            return f"📋 {label} 的購物車是空的\n直接輸入「貨號*數量」加購"
        return _format_cart(label, cart)

    # ── 加購品項（指定客戶名）──
    m = _CART_ADD_RE.match(t)
    if m:
        name, code, qty_str = m.group(1), m.group(2).upper(), int(m.group(3))
        line_id, label = _resolve_customer_line_id(name)
        if not line_id:
            return f"❌ 找不到客戶「{name}」"
        _set_cart_session(staff_id or "default", line_id, label)
        info = ecount_client.lookup(code)
        prod_name = (info.get("name") if info else None) or code
        cart = cart_store.set_item(line_id, code, prod_name, qty_str)
        return f"✅ 已設定 {prod_name} × {qty_str}\n" + _format_cart(label, cart)

    # ── 代結帳（指定客戶名）──
    m = _CART_CHECKOUT_RE.match(t)
    if m:
        name = m.group(1)
        line_id, label = _resolve_customer_line_id(name)
        if not line_id:
            return f"❌ 找不到客戶「{name}」"
        cart = cart_store.get_cart(line_id)
        if not cart:
            return f"❌ {label} 的購物車是空的，無法結帳"
        _clear_cart_session(staff_id or "default")
        from handlers.ordering import handle_checkout
        result = handle_checkout(line_id, line_api)
        return f"【代 {label} 結帳】\n{result}"

    # ── Session 快速操作 ──
    session = _get_cart_session(staff_id or "default")
    if session:
        line_id = session["line_id"]
        label = session["label"]

        # 快速加購/改數量：T0101*6（同品項覆蓋數量）
        m = _CART_QUICK_ADD_RE.match(t)
        if m:
            code, qty_str = m.group(1).upper(), int(m.group(2))
            _touch_cart_session(staff_id or "default")
            info = ecount_client.lookup(code)
            prod_name = (info.get("name") if info else None) or code
            existing_cart = cart_store.get_cart(line_id)
            _was_in_cart = any(i["prod_cd"].upper() == code for i in existing_cart)
            cart = cart_store.set_item(line_id, code, prod_name, qty_str)
            _verb = "已修改" if _was_in_cart else "已加購"
            return f"✅ {_verb} {prod_name} × {qty_str}\n" + _format_cart(label, cart)

        # 快速改數量：「改6」「改6個」→ 改最後一個品項
        _chg_m = re.match(r'^改\s*(\d+)\s*[個箱件盒組條]?', t)
        if _chg_m:
            _touch_cart_session(staff_id or "default")
            cart = cart_store.get_cart(line_id)
            if not cart:
                return f"❌ {label} 的購物車是空的"
            _last = cart[-1]
            _new_qty = int(_chg_m.group(1))
            cart_store.set_item(line_id, _last["prod_cd"], _last["prod_name"], _new_qty)
            cart = cart_store.get_cart(line_id)
            return f"✅ 已修改 {_last['prod_name']} → {_new_qty}\n" + _format_cart(label, cart)

        # 快速送出
        if t in _CART_QUICK_SUBMIT:
            cart = cart_store.get_cart(line_id)
            if not cart:
                _clear_cart_session(staff_id or "default")
                return f"❌ {label} 的購物車是空的"
            _clear_cart_session(staff_id or "default")
            from handlers.ordering import handle_checkout
            result = handle_checkout(line_id, line_api)
            return f"【代 {label} 結帳】\n{result}"

    return None


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

    from handlers.ordering import resolve_unit

    def _apply_unit(prod_cd: str, qty: int, unit: str | None) -> tuple[str, int]:
        """統一單位換算，回傳 (實際品項編號, 實際數量)。"""
        cd, q, _warn = resolve_unit(prod_cd, qty, unit)
        return cd, q

    # ── 格式B 判斷：第一行符合「姓名訂」且後面沒有產品代碼 ──
    header_m = _STAFF_ORDER_HEADER_RE.match(lines[0])
    if header_m and len(lines) > 1:
        # 確認第一行沒有夾帶產品代碼
        if not _STAFF_ORDER_LINE_RE.search(lines[0]):
            cust_name = header_m.group(1).strip()
            items_raw = []
            _pending_note = ""  # 獨立行備註，歸給上一個品項
            for l in lines[1:]:
                im = _STAFF_ORDER_ITEM_RE.search(l)
                if im:
                    # 先把上一行的獨立備註補給上一個品項
                    if _pending_note and items_raw:
                        prev = items_raw[-1]
                        items_raw[-1] = (prev[0], prev[1], _pending_note)
                        _pending_note = ""
                    prod_cd = im.group(1).strip()
                    qty     = _parse_qty(im.group(2))
                    unit    = im.group(3) if im.lastindex >= 3 else None
                    _cd, _q = _apply_unit(prod_cd, qty, unit)
                    # 檢查同行備註：有「備註:」前綴就剝掉，否則尾段整段當備註
                    _rest = l[im.end():].strip()
                    if re.search(r'備[註誌记]', _rest):
                        _inline_note = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', _rest).strip()
                    else:
                        _inline_note = _rest
                    items_raw.append((_cd, _q, _inline_note))
                elif re.match(r'^備[註誌记]\s*[:：]?\s*', l):
                    _pending_note = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', l).strip()
            # 最後一個獨立備註
            if _pending_note and items_raw:
                prev = items_raw[-1]
                items_raw[-1] = (prev[0], prev[1], _pending_note)
            if items_raw:
                return _do_order(cust_name, items_raw, group_id=group_id)

    # ── 格式B2：第一行只有姓名（無「訂」），後續行有貨號+數量 ──
    # 例：鄭鉅耀\nZ3340 10個\nZ3338 20個\n備註 送松山
    if len(lines) > 1 and not _PROD_CODE_RE.search(lines[0]) and not _STAFF_ORDER_HEADER_RE.match(lines[0]):
        items_b2 = []
        _pending_note_b2 = ""
        for l in lines[1:]:
            im = _STAFF_ORDER_ITEM_RE.search(l)
            if im:
                if _pending_note_b2 and items_b2:
                    prev = items_b2[-1]
                    items_b2[-1] = (prev[0], prev[1], _pending_note_b2)
                    _pending_note_b2 = ""
                _rest = l[im.end():].strip()
                # 尾段有「備註:」前綴就剝掉，否則整段當備註
                if re.search(r'備[註誌记]', _rest):
                    _inline_note = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', _rest).strip()
                else:
                    _inline_note = _rest
                prod_cd = im.group(1).strip()
                qty     = _parse_qty(im.group(2))
                unit    = im.group(3) if im.lastindex >= 3 else None
                _cd2, _q2 = _apply_unit(prod_cd, qty, unit)
                items_b2.append((_cd2, _q2, _inline_note))
            elif re.match(r'^備[註誌记]\s*[:：]?\s*', l):
                _pending_note_b2 = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', l).strip()
        if _pending_note_b2 and items_b2:
            prev = items_b2[-1]
            items_b2[-1] = (prev[0], prev[1], _pending_note_b2)
        if items_b2:
            cust_name_b2 = lines[0].strip()
            return _do_order(cust_name_b2, items_b2, group_id=group_id)

    # ── 格式A：每行各自獨立 ──
    valid = [(l, _STAFF_ORDER_LINE_RE.search(l)) for l in lines]
    valid = [(l, m) for l, m in valid if m]
    if valid:
        results = []
        for _line, m in valid:
            _note_a  = m.group(5).strip() if m.lastindex >= 5 else ""
            _prod_cd = m.group(2).strip()
            _qty     = _parse_qty(m.group(3))
            _unit    = m.group(4) if m.lastindex >= 4 else None
            res = _do_order(
                cust_name_query=m.group(1).strip(),
                items_raw=[_apply_unit(_prod_cd, _qty, _unit)],
                note=_note_a,
            )
            results.append(res)
        return "\n".join(results)

    # ── 格式C：「姓名 產品代碼*數量」（支援多行，第一行有姓名+品項，後續行只有品項）──
    # 例：林銘宇 Z3251*1
    #      HH008-022*3
    m_c = _STAFF_ORDER_DIRECT_RE.match(lines[0])
    if m_c:
        _cust_name_c = m_c.group(1).strip()
        _prod_cd_c = m_c.group(2).strip()
        _qty_c     = _parse_qty(m_c.group(3))
        _unit_c    = m_c.group(4) if m_c.lastindex >= 4 else None
        _note_c    = m_c.group(5).strip() if m_c.lastindex >= 5 else ""
        # 尾段可能包含更多品項（如「U0360*16」），先提取再當備註
        _extra_items_c = list(_STAFF_ORDER_ITEM_RE.finditer(_note_c))
        if _extra_items_c:
            # 尾段有品項 → 提取出來，剩餘才是備註
            _real_note_c = _note_c
            for _ei in reversed(_extra_items_c):
                _real_note_c = _real_note_c[:_ei.start()] + _real_note_c[_ei.end():]
            _note_c = _real_note_c.strip()
        # 備註處理：group(5) 可能含「備註:XXX」
        if re.search(r'備[註誌记]', _note_c):
            _note_c = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', _note_c).strip()
        _cd_c, _q_c = _apply_unit(_prod_cd_c, _qty_c, _unit_c)
        _items_c = [(_cd_c, _q_c, _note_c)]
        # 尾段提取到的額外品項加入
        for _ei in _extra_items_c:
            _ei_cd = _ei.group(1).strip()
            _ei_q  = _parse_qty(_ei.group(2))
            _ei_u  = _ei.group(3) if _ei.lastindex >= 3 else None
            _ei_cd2, _ei_q2 = _apply_unit(_ei_cd, _ei_q, _ei_u)
            _items_c.append((_ei_cd2, _ei_q2, ""))
        # 後續行
        _pending_note_c = ""
        for _lc in lines[1:]:
            _im_c = _STAFF_ORDER_ITEM_RE.search(_lc)
            if _im_c:
                if _pending_note_c and _items_c:
                    prev = _items_c[-1]
                    _items_c[-1] = (prev[0], prev[1], _pending_note_c)
                    _pending_note_c = ""
                _cd = _im_c.group(1).strip()
                _q  = _parse_qty(_im_c.group(2))
                _u  = _im_c.group(3) if _im_c.lastindex >= 3 else None
                _rest_c = _lc[_im_c.end():].strip()
                _inline_c = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', _rest_c).strip() if re.search(r'備[註誌记]', _rest_c) else ""
                _cd2, _q2 = _apply_unit(_cd, _q, _u)
                _items_c.append((_cd2, _q2, _inline_c))
            elif re.match(r'^備[註誌记]\s*[:：]?\s*', _lc):
                _pending_note_c = re.sub(r'^備[註誌记]\s*[:：]?\s*', '', _lc).strip()
        if _pending_note_c and _items_c:
            prev = _items_c[-1]
            _items_c[-1] = (prev[0], prev[1], _pending_note_c)
        return _do_order(
            cust_name_query=_cust_name_c,
            items_raw=_items_c,
            group_id=group_id,
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
                group_id=group_id,
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
                            _final_cd, _final_qty = _apply_unit(_codes[0], _qty, _unit)
                            _it = ecount_client.lookup(_final_cd)
                            _pn = (_it.get("name") if _it else "") or _final_cd
                            _resolved.append({
                                "query": _name, "code": _final_cd,
                                "name": _pn, "qty": _final_qty,
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
                                _resolved.append({
                                    "query": _name, "code": _auto_code,
                                    "name": _pn, "qty": _qty,
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
                        _real_u = (ecount_client.get_product_cache_item(_r["code"]) or {}).get("unit", "個") or "個"
                        if _dq != _aq:
                            _lines.append(f"  📦 {_r['name']}（{_r['code']}）× {_dq} {_u} = {_aq} {_real_u}")
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
            _real_u = (ecount_client.get_product_cache_item(_r["code"]) or {}).get("unit", "個") or "個"
            if _dq != _aq:
                lines.append(f"  ✅ {_r['name']}（{_r['code']}）× {_dq} {_u} = {_aq} {_real_u}")
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
    from handlers.ordering import resolve_unit
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
    chosen_code, _actual_qty, _warn = resolve_unit(chosen_code, qty, unit)
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
        return _do_order(customer, items_raw, units=units, note=note, group_id=group_id)

    return None


def handle_new_customer_confirm(group_id: str, text: str) -> str | None:
    """
    處理「找不到客戶」後的確認：
    - 「是」→ 先同步客戶資料再建單（今天新客人，Ecount 已有但 json 未同步）
    - 「不是」→ 新建 Ecount 客戶再建單
    state action == "confirm_new_customer" 時才介入。
    """
    global _ec_customers_cache, _ec_customers_mtime
    from storage.state import state_manager as _sm
    state = _sm.get(group_id)
    if not state or state.get("action") != "confirm_new_customer":
        return None

    t = text.strip()
    cust_name  = state["cust_name"]
    items_raw  = state["items_raw"]
    units      = state.get("units", {})
    note       = state.get("note", "")

    if t in ("是", "是!", "是！", "對"):
        # 今天新客人 → 同步 ecount_customers.json 再重新建單
        _sm.clear(group_id)
        print(f"[internal] 確認「是」→ 同步客戶資料後重新建單: {cust_name}", flush=True)
        try:
            import subprocess as _sp, sys as _sys_local
            from pathlib import Path as _Path
            _python = _sys_local.executable
            _root   = str(_Path(__file__).parent.parent)
            _flags  = _sp.CREATE_NO_WINDOW if _sys_local.platform == "win32" else 0
            proc = _sp.run(
                [_python, "-m", "scripts.sync_cust_from_web"],
                cwd=_root, capture_output=True, timeout=300, creationflags=_flags,
            )
            if proc.returncode == 0:
                print("[internal] 客戶資料同步完成", flush=True)
                # 重新載入客戶快取
                _ec_customers_cache = None
                _ec_customers_mtime = 0
            else:
                stderr = proc.stderr.decode("utf-8", errors="replace")[-500:] if proc.stderr else ""
                print(f"[internal] 客戶資料同步失敗: {stderr}", flush=True)
                return f"⚠️ 客戶資料同步失敗，請稍後再試"
        except Exception as e:
            print(f"[internal] 客戶資料同步例外: {e}", flush=True)
            return f"⚠️ 客戶資料同步失敗: {e}"
        # 確認同步後找得到客戶
        ec_match = _resolve_customer(cust_name)
        if not ec_match:
            return f"❌ 客戶資料已同步，但仍找不到「{cust_name}」，訂單未建立\n請確認 Ecount 客戶名稱是否正確"
        # 重新建單
        return _do_order(cust_name, items_raw, units=units, note=note, group_id=group_id)

    if t in ("不是", "不是!", "不是！", "否", "新客人"):
        # 不是今天新客人 → 直接新建 Ecount 客戶
        _sm.clear(group_id)
        print(f"[internal] 確認「不是」→ 新建 Ecount 客戶: {cust_name}", flush=True)
        from datetime import datetime as _dt
        import sqlite3 as _sqlite3
        from pathlib import Path as _Path
        _today   = _dt.now().strftime("%y%m%d")
        _prefix  = f"M{_today}"
        _db_path = str(_Path(__file__).parent.parent / "data" / "customers.db")
        try:
            with _sqlite3.connect(_db_path) as _conn:
                _rows = _conn.execute(
                    "SELECT ecount_cust_cd FROM customers WHERE ecount_cust_cd LIKE ?",
                    (f"{_prefix}%",),
                ).fetchall()
            _nums  = [int(cd[len(_prefix):]) for (cd,) in _rows
                      if cd and cd.startswith(_prefix) and cd[len(_prefix):].isdigit()]
            _serial = max(_nums) + 1 if _nums else 1000
        except Exception:
            _serial = 1000
        new_cust_cd = f"{_prefix}{_serial}"
        ok = ecount_client.save_customer(new_cust_cd, cust_name)
        if ok:
            print(f"[internal] 新建 Ecount 客戶: {cust_name} → {new_cust_cd}", flush=True)
            # 刷新快取
            _ec_customers_cache = None
            _ec_customers_mtime = 0
        else:
            return f"❌ 新建客戶「{cust_name}」失敗"
        return _do_order(cust_name, items_raw, units=units, note=note, group_id=group_id)

    # 其他文字不處理
    return None


# ── 4. 手動通知登記 ────────────────────────────────────────────────────

# ── 新增客戶 ────────────────────────────────────────────────────────────
# 支援格式：
#   新增客戶 張三
#   新增客戶 張三 0912345678
#   新增客戶 張三 0912345678 台北市XX路1號
#   新增客戶\n張三\n0912345678\n台北市XX路1號   （名字可換行）
_ADD_CUST_TRIGGER_RE = re.compile(r'^新增客戶[\s\n]+(.+)', re.DOTALL)
_PHONE_RE_EXTRACT    = re.compile(r'(09\d{8})')

def _gen_ecount_cust_cd() -> str:
    """
    自動產生 Ecount 客戶代碼。
    格式：M + 年後兩碼 + 月(2位) + 日(2位) + 流水號(4位，1000起)
    例：M2603171000、M2603171001
    """
    from datetime import datetime
    import sqlite3, os
    now    = datetime.now()
    prefix = f"M{now.strftime('%y%m%d')}"  # e.g. M260317
    db_path = os.path.join(os.path.dirname(__file__), "..", "data", "customers.db")
    try:
        with sqlite3.connect(db_path) as conn:
            rows = conn.execute(
                "SELECT ecount_cust_cd FROM customers WHERE ecount_cust_cd LIKE ?",
                (f"{prefix}%",)
            ).fetchall()
        nums = [int(cd[len(prefix):]) for (cd,) in rows
                if cd and cd.startswith(prefix) and cd[len(prefix):].isdigit()]
        seq = max(nums) + 1 if nums else 1000
    except Exception:
        seq = 1000
    return f"{prefix}{seq}"

def handle_internal_add_customer(text: str) -> str | None:
    """
    內部群新增客戶指令（支援單行/多行）。
    同步建立 Ecount 客戶，回覆帶 Ecount 代碼。
    """
    t = text.strip()
    m = _ADD_CUST_TRIGGER_RE.match(t)
    if not m:
        return None

    rest  = m.group(1).strip()
    lines = [l.strip() for l in rest.splitlines() if l.strip()]
    if not lines:
        return None

    # 第一行：姓名（去掉電話後的第一個詞）
    first   = lines[0]
    phone_m = _PHONE_RE_EXTRACT.search(first)
    phone   = phone_m.group(1) if phone_m else ""
    name    = _PHONE_RE_EXTRACT.sub("", first).strip().split()[0]

    # 地址：第一行電話後的剩餘 + 後續行
    addr_inline = _PHONE_RE_EXTRACT.sub("", first).replace(name, "", 1).strip()
    addr_parts  = [addr_inline] if addr_inline else []

    for ln in lines[1:]:
        p_m = _PHONE_RE_EXTRACT.search(ln)
        if p_m and not phone:
            phone = p_m.group(1)
            remainder = _PHONE_RE_EXTRACT.sub("", ln).strip()
            if remainder:
                addr_parts.append(remainder)
        else:
            addr_parts.append(ln)

    address = " ".join(addr_parts).strip()

    # 若姓名已存在則提示（含現有 Ecount 代碼）
    existing = customer_store.search_by_name(name, real_name_only=True)
    if existing:
        ex       = existing[0]
        ex_phone = ex.get("phone") or "無"
        ex_code  = ex.get("ecount_cust_cd") or "無"
        return f"⚠️ 客戶「{name}」已存在\n📞 {ex_phone}\n🔑 Ecount代碼：{ex_code}"

    # 建立本地客戶
    customer_store.import_from_csv_data(
        display_name=name,
        chat_label=name,
        phones=[phone] if phone else [],
        address=address,
    )

    # 自動產生 Ecount 代碼並同步
    cust_cd = _gen_ecount_cust_cd()
    ec_ok = ecount_client.save_customer(
        business_no=cust_cd,
        cust_name=name,
        hp_no=phone,
        addr=address,
    )

    # 把代碼存回本地 DB
    cust_matches = customer_store.search_by_name(name, real_name_only=True)
    if cust_matches:
        db_id = cust_matches[0]["id"]
        customer_store.update_ecount_cust_cd_by_db_id(db_id, cust_cd)

    parts = [f"✅ 已新增客戶「{name}」"]
    parts.append(f"🔑 Ecount代碼：{cust_cd}" + ("" if ec_ok else "（本地已存，Ecount同步失敗）"))
    if phone:
        parts.append(f"📞 {phone}")
    if address:
        parts.append(f"📍 {address}")
    return "\n".join(parts)


def handle_internal_notify_register(text: str, line_api=None) -> str | None:
    """
    幫客戶登記到貨通知，登記後立即 push 通知客戶（到貨即通知）。
    支援格式：

        格式 A（代碼獨立一行，名字+數量在後）：
            登記通知
            Z3336
            王小明  5
            張三10
            李四                  ← 無數量預設 1

        格式 B（shorthand，代碼與名字同行）：
            通知 T1202            ← 後面每行一個名字
            張三
            李四
            通知 張三 T1202       ← 單行一人

        格式 C（句首「通知登記」，每行名稱+代碼）：
            通知登記 張三 T1202 3個
            通知登記
            張三 T1202 3個
            楊庭瑋 T1208 8個

        格式 D（句尾/句中關鍵字）：
            張三 T1202 需要到貨通知
            張三 T1202 3個 要通知
    """
    t = text.strip()

    has_inline_kw = any(kw in t for kw in _NOTIFY_REG_INLINE_KW)
    is_start_fmt  = _NOTIFY_REG_START_RE.match(t)
    m_shorthand   = _NOTIFY_REG_SHORTHAND_RE.match(t)

    # ── 格式 A：「登記通知/通知登記\n第二行\n後續行...」─────────────────────
    _fmt_a_trigger = None
    if t.startswith("登記通知"):
        _fmt_a_trigger = "登記通知"
    elif t.startswith("通知登記"):
        _fmt_a_trigger = "通知登記"
    if _fmt_a_trigger:
        lines_all = [l.strip() for l in t.splitlines() if l.strip()]
        # 去掉第一行觸發詞
        lines_body = lines_all[1:]
        if not lines_body:
            return None

        line2 = lines_body[0]
        rest  = lines_body[1:]

        # ── A1：第二行是貨號 → 一品多客 ──────────────────────────────
        if _NOTIFY_PROD_CODE_PAT.match(line2):
            prod_code = line2.upper()
            item      = ecount_client.lookup(prod_code)
            prod_name = (item["name"] if item else "") or prod_code
            results   = []
            for nl in rest:
                # 解析「名字  數量 單位」或「名字10箱」
                mq = re.match(r'^(.+?)[\s　]+(\d+)\s*(個|件|盒|套|箱|組)?$', nl)
                if not mq:
                    mq2 = re.match(r'^(.+?)(\d+)\s*(個|件|盒|套|箱|組)?$', nl)
                    cust_name_q = re.sub(r'\s+', '', mq2.group(1)) if mq2 else re.sub(r'\s+', '', nl)
                    qty = int(mq2.group(2)) if mq2 else 1
                    unit = (mq2.group(3) or "") if mq2 else ""
                else:
                    cust_name_q = re.sub(r'\s+', '', mq.group(1))
                    qty = int(mq.group(2))
                    unit = mq.group(3) or ""
                from handlers.ordering import resolve_unit
                _cd, qty, _warn = resolve_unit(prod_code, qty, unit or None)
                results.append(_notify_register_and_push(cust_name_q, _cd, prod_name, qty))
            return "\n".join(results) if results else None

        # ── A2：第二行是客戶名 → 一客多品 ──────────────────────────────
        else:
            cust_name_q = re.sub(r'\s+', '', line2)
            results = []
            for nl in rest:
                # 解析「貨號  數量」或「貨號*數量」
                mp = re.match(
                    r'^([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*[*×\s]\s*([零一二三四五六七八九十百千\d]+)\s*(?:個|件|盒|套|箱|組)?$',
                    nl, re.IGNORECASE
                )
                if not mp:
                    results.append(f"⚠️ 無法解析：「{nl}」")
                    continue
                prod_code = mp.group(1).upper()
                qty       = _parse_qty(mp.group(2))
                item      = ecount_client.lookup(prod_code)
                prod_name = (item["name"] if item else "") or prod_code
                results.append(_notify_register_and_push(cust_name_q, prod_code, prod_name, qty))
            return "\n".join(results) if results else None

    if not has_inline_kw and not is_start_fmt and not m_shorthand:
        return None

    # ── 格式 B：shorthand「通知/登記 [name] CODE」────────────────────────
    if m_shorthand and not is_start_fmt:
        inline_name  = m_shorthand.group(1)
        prod_code_sh = m_shorthand.group(2).upper()
        item         = ecount_client.lookup(prod_code_sh)
        prod_name_sh = (item["name"] if item else "") or prod_code_sh

        if inline_name:
            name_list = [re.sub(r'\s+', '', inline_name.strip())]
        else:
            name_list = [re.sub(r'\s+', '', l.strip())
                         for l in t.splitlines()[1:] if l.strip()]
        if not name_list:
            return None

        results = []
        for cust_name_q in name_list:
            results.append(_notify_register_and_push(
                cust_name_q, prod_code_sh, prod_name_sh, 1, line_api
            ))
        return "\n".join(results)

    # ── 格式 C：句首「通知登記」──────────────────────────────────────────
    if is_start_fmt:
        first_newline = t.find("\n")
        if first_newline == -1:
            body  = t[len("通知登記"):].strip()
            lines = [body] if body else []
        else:
            first_rest = t[:first_newline].replace("通知登記", "").strip()
            rest_lines = [l.strip() for l in t[first_newline:].splitlines() if l.strip()]
            lines = ([first_rest] if first_rest else []) + rest_lines
    else:
        # ── 格式 D：句尾/句中含關鍵字 ─────────────────────────────────────
        cleaned = t
        for kw in _NOTIFY_REG_INLINE_KW:
            cleaned = cleaned.replace(kw, "").strip()
        lines = [cleaned] if cleaned else []

    if not lines:
        return None

    results = []
    for line in lines:
        line = re.sub(r'(?<=[^\x00-\x7F])\s+(?=[^\x00-\x7F])', '', line)
        m = _NOTIFY_REG_LINE_RE.search(line)
        if not m:
            results.append(f"⚠️ 無法解析：「{line}」")
            continue
        cust_name_q = m.group(1).strip()
        prod_code   = m.group(2).upper()
        qty         = _parse_qty(m.group(3)) if m.group(3) else 1
        unit        = m.group(4) or ""
        from handlers.ordering import resolve_unit
        prod_code, qty, _warn = resolve_unit(prod_code, qty, unit or None)
        item        = ecount_client.lookup(prod_code)
        prod_name   = (item["name"] if item else "") or prod_code
        results.append(_notify_register_and_push(
            cust_name_q, prod_code, prod_name, qty, line_api
        ))

    return "\n".join(results) if results else None


def _notify_register_and_push(
    cust_name_q: str, prod_code: str, prod_name: str, qty: int, line_api=None
) -> str:
    """
    查找客戶 → 登記 notify_store → 回傳狀態行。
    到貨時再由排程自動 push（不立即通知）。
    """
    # 取產品單位用於顯示
    _item_unit = ecount_client.get_product_cache_item(prod_code)
    _display_unit = (_item_unit.get("unit") if _item_unit else "") or "個"

    matches = customer_store.search_by_name(cust_name_q, real_name_only=True)
    if not matches:
        # LINE DB 找不到 → 用 ecount: 前綴存，到貨時通知到內部群
        notify_id = notify_store.add(
            user_id=f"ecount:{cust_name_q}", prod_code=prod_code,
            prod_name=prod_name, qty_wanted=qty,
            source="staff",
        )
        print(f"[internal] 通知登記(staff): #{notify_id} {cust_name_q}（到貨通知內部群）← {prod_name}({prod_code}) x{qty}")
        return f"✅ {cust_name_q}｜{prod_name}（{prod_code}）× {qty} {_display_unit}（到貨通知內部群）"
    if len(matches) > 1:
        ns = "、".join(r.get("real_name") or r.get("display_name", "?") for r in matches[:5])
        return f"⚠️ 「{cust_name_q}」有多位：{ns}"

    cust       = matches[0]
    cust_uid   = cust.get("line_user_id", "")
    cust_label = cust.get("real_name") or cust.get("display_name") or cust_name_q

    if not cust_uid:
        # 無 LINE ID → 用 ecount: 前綴存，到貨時通知到內部群
        cust_uid = f"ecount:{cust_label}"

    notify_id = notify_store.add(
        user_id=cust_uid, prod_code=prod_code,
        prod_name=prod_name, qty_wanted=qty,
        source="staff",
    )
    tag = "（到貨通知內部群）" if cust_uid.startswith("ecount:") else ""
    print(f"[internal] 通知登記(staff): #{notify_id} {cust_label}{tag} ← {prod_name}({prod_code}) x{qty}")
    return f"✅ {cust_label}｜{prod_name}（{prod_code}）× {qty} {_display_unit}{tag}"


# ── 3. 圖片識別 → PO文 + 等待訂單 ───────────────────────────────────

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
    stock_detail = _fmt_stock_lines(item, prod_code)

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
    qty_m = _QTY_TAIL_RE.search(text.strip())
    if not qty_m:
        return None  # 沒有數字 → 格式不符

    qty = int(qty_m.group(1))

    # 策略：若含「要/訂/下單」動詞 → 動詞前面的部分就是客戶名
    verb_m = _VERB_SEP_RE.search(text.strip())
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

    # ── 優先查 Ecount 客戶清單 ──────────────────────────────────────
    cust_code  = ""
    cust_label = cust_name_query
    _phone     = ""

    ec_match = _resolve_customer(cust_name_query)
    if ec_match:
        cust_code  = ec_match.get("code", "")
        cust_label = ec_match.get("name", cust_name_query)
        _phone     = ec_match.get("phone", "") or ec_match.get("tel", "") or ""
        print(f"[internal] Ecount 客戶匹配: {cust_label} → {cust_code}", flush=True)

    # 找不到 → fallback 查本地 LINE 資料庫
    if not cust_code:
        matches = customer_store.search_by_name(cust_name_query, real_name_only=True)
        if not matches:
            # 客戶名在任何資料庫都找不到 → 可能是無關訊息（如「開會 3點」）
            # 回傳 None 讓其他 handler 繼續處理，而非直接報錯
            return None
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
        return f"❌ 訂單建立失敗（API 錯誤：{e}）\n客戶：{cust_label}\n商品：{prod_name} × {qty}"

    state_manager.clear(state_key)  # 訂單完成，清除 state

    if slip_no:
        print(f"[internal] 圖片代訂成功: {slip_no} | {cust_label} | {prod_name} x{qty}")
        _item_u = ecount_client.get_product_cache_item(prod_cd)
        _disp_u = (_item_u.get("unit") if _item_u else "") or "個"
        return (
            f"✅ 訂單建立成功\n"
            f"客戶：{cust_label}\n"
            f"商品：{prod_name} × {qty} {_disp_u}"
        )
    else:
        print(f"[internal] 圖片代訂失敗: {cust_code} | {prod_name} x{qty}")
        return f"❌ 訂單建立失敗，請手動建立\n客戶：{cust_label}\n商品：{prod_name} × {qty}"


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
    # 精準匹配：後面不能接字母/數字（避免 Z181 抓 Z1814、Z1814 抓 Z1814A），
    # 也不能接「-數字」（避免 Z1814 抓到 Z1814-1）；
    # 但允許「-非數字」如 `Z3671-(原)...`、`Z3671-中文`
    _code_re = re.compile(
        r'(?<![A-Z0-9])' + re.escape(code_upper) + r'(?![A-Za-z0-9])(?!-\d)'
    )
    # 以空白行切成段落
    blocks = [b.strip() for b in content.split("\n\n") if b.strip()]
    for block in blocks:
        if _code_re.search(block.upper()):
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
_SPEC_QUERY_KW = ["有哪些", "有什麼", "哪些產品", "什麼產品", "產品", "有哪些產品", "推薦", "庫存"]

import re as _re_spec
_SIZE_RE = _re_spec.compile(r'(\d+(?:\.\d+)?)\s*公分')
_PRICE_RE = _re_spec.compile(r'(\d+)\s*(?:元|塊)?以下')


def handle_internal_spec_query(text: str) -> str | None:
    """
    「中巨的產品有哪些」→ 搜尋適合台型的產品 + 庫存
    「13公分的有哪些產品」→ 搜尋尺寸符合的產品 + 庫存
    回傳 None 表示不是此類查詢。
    """
    from storage.specs import get_by_machine, get_by_size

    # 多行訊息（超過 2 行）不是簡單查詢，跳過
    if text.count("\n") >= 2:
        return None

    # 必須含列表意圖關鍵字
    if not any(kw in text for kw in _SPEC_QUERY_KW):
        return None

    # 偵測台型
    matched_machine = next((m for m in _MACHINE_TYPES if m in text), None)
    # 偵測尺寸
    m_size = _SIZE_RE.search(text)
    size_kw = m_size.group(0) if m_size else None  # 例：「13公分」
    # 偵測價格上限
    m_price = _PRICE_RE.search(text)
    price_limit = int(m_price.group(1)) if m_price else None

    if not matched_machine and not size_kw and not price_limit:
        return None

    # 是否只列有庫存的
    stock_only = "庫存" in text

    # 搜尋規格DB
    if matched_machine:
        specs = get_by_machine(matched_machine)
        label = f"「{matched_machine}」台型"
    elif size_kw:
        specs = get_by_size(size_kw)
        label = f"「{size_kw}」尺寸"
    else:
        # 價格搜尋：取全部規格再篩選
        from storage.specs import get_all
        all_specs = get_all()
        specs = list(all_specs.values()) if isinstance(all_specs, dict) else all_specs
        label = f"「{price_limit}元以下」"

    if not specs:
        return f"🔍 規格DB 目前沒有{label}的產品記錄"

    # 排除耗材
    specs = [s for s in specs if not s.get("code", "").upper().startswith("HH")]

    # 依價格由高到低排序
    def _price_val(s):
        nums = re.findall(r'\d+', str(s.get("price", "0")))
        return int(nums[0]) if nums else 0
    specs = sorted(specs, key=_price_val, reverse=True)

    # 價格上限篩選
    if price_limit:
        specs = [s for s in specs if 0 < _price_val(s) <= price_limit]

    in_stock = []
    out_of_stock = []
    for s in specs:
        code = s.get("code", "")
        name = s.get("name", code)
        try:
            item = ecount_client.lookup(code)
            qty = item.get("qty") if item else None
        except Exception:
            qty = None
        if qty is not None and qty > 0:
            in_stock.append(f"  {code}　{name}　可售:{qty}")
        else:
            if stock_only:
                continue
            out_of_stock.append(f"  {code}　{name}　缺貨")

    lines = [f"🔍 {label} 產品"]
    if in_stock:
        lines.append(f"\n✅ 有庫存（{len(in_stock)} 筆）：")
        lines.extend(in_stock)
    if out_of_stock and not stock_only:
        lines.append(f"\n❌ 缺貨（{len(out_of_stock)} 筆）：")
        lines.extend(out_of_stock)
    if not in_stock and not out_of_stock:
        return f"🔍 {label} 目前沒有產品"
    if stock_only and not in_stock:
        return f"🔍 {label} 目前沒有有庫存的產品"
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

# 品名搜尋時要剝除的關鍵字（預編譯正則，長詞優先）
_STRIP_KW_LIST = (
    _INV_QUERY_KW + _PREORDER_KW + _PRODUCT_LIST_KW
    + ["有哪些", "有什麼", "哪些", "什麼", "有嗎", "嗎", "有", "？", "?",
       "多少", "還", "個", "數量", "都", "各", "查詢", "產品", "品項"]
)
_STRIP_KW_RE = re.compile('|'.join(
    re.escape(kw) for kw in sorted(_STRIP_KW_LIST, key=len, reverse=True)
))


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

    # 出庫價格（從 available.json 讀）
    unit_price = None
    try:
        import json as _j
        _avail_path = Path(__file__).parent.parent / "data" / "available.json"
        if _avail_path.exists():
            _avail = _j.loads(_avail_path.read_text(encoding="utf-8"))
            _d = _avail.get(prod_code)
            if isinstance(_d, dict):
                unit_price = _d.get("unit_price")
    except Exception:
        pass

    # 取單位/裝箱量；箱裝變體（unit=箱 且 box_qty=1）改讀兄弟代碼的 box_qty
    _cache = ecount_client.get_product_cache_item(prod_code) or {}
    _unit  = _cache.get("unit") or "個"
    _bq    = _cache.get("box_qty") or 0
    _is_box_variant = (_unit == "箱")
    if _is_box_variant and _bq <= 1:
        _base_code = prod_code.rsplit("-", 1)[0] if "-" in prod_code else prod_code
        _bq = (ecount_client.get_product_cache_item(_base_code) or {}).get("box_qty") or 0

    def _fmt_qty(n: int | float) -> str:
        """依品項單位格式化數量；箱裝變體會同時顯示箱 + 餘數"""
        try:
            n = int(n)
        except Exception:
            return f"{n} {_unit}"
        if _is_box_variant and _bq > 1:
            boxes, rem = divmod(n, _bq)
            if rem == 0:
                return f"{boxes} 箱"
            return f"{boxes} 箱 {rem} 個"
        return f"{n} {_unit}"

    lines = [f"📦 {name}（{prod_code}）"]
    mid_lines = []

    if unit_price and unit_price > 0:
        mid_lines.append(f"出庫單價：${unit_price:,.0f}")
    if balance is not None:
        mid_lines.append(f"倉庫庫存：{_fmt_qty(balance)}")
    if unfilled is not None:
        mid_lines.append(f"ERP未出：{_fmt_qty(unfilled)}")
    if incoming is not None:
        mid_lines.append(f"總公司未到：{_fmt_qty(incoming)}")
    if qty is None:
        mid_lines.append("可售庫存：查詢失敗")
    elif qty <= 0:
        mid_lines.append(f"可售庫存：0 {_unit}（缺貨）")
    else:
        mid_lines.append(f"可售庫存：{_fmt_qty(qty)}")
    if preorder and preorder > 0:
        mid_lines.append(f"可預購：{_fmt_qty(preorder)}")

    for i, ln in enumerate(mid_lines):
        prefix = "  └ " if i == len(mid_lines) - 1 else "  ├ "
        lines.append(prefix + ln)

    return "\n".join(lines)


def _fmt_stock_lines(item: dict, prod_code: str = "") -> str:
    """
    回傳庫存明細純文字（不含產品名稱標題），供 PO文 + 庫存格式使用。
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

    # 出庫單價
    unit_price = None
    if prod_code:
        try:
            import json as _j
            _avail_path = Path(__file__).parent.parent / "data" / "available.json"
            if _avail_path.exists():
                _avail = _j.loads(_avail_path.read_text(encoding="utf-8"))
                _d = _avail.get(prod_code.upper())
                if isinstance(_d, dict):
                    unit_price = _d.get("unit_price")
        except Exception:
            pass

    lines = []
    if unit_price and unit_price > 0:
        lines.append(f"出庫單價：${unit_price:,.0f}")
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


def _fuzzy_product_search(query: str, max_results: int = 50) -> list[str]:
    """
    模糊品名搜尋：整串 → 逐詞 → 逐字縮短，直到有結果。
    回傳產品編號清單（最多 max_results 筆）。
    """
    q = query.strip()
    if len(q) < 2:
        return []
    # 1. 整串搜尋
    codes = ecount_client.search_products_by_name(q)
    if codes:
        return codes[:max_results]
    # 2. 逐詞搜尋（空格分詞）
    tokens = [t for t in q.split() if len(t) >= 2]
    for token in tokens:
        codes = ecount_client.search_products_by_name(token)
        if codes:
            return codes[:max_results]
    # 3. 逐字縮短（從尾巴砍一個字）
    while len(q) > 2:
        q = q[:-1]
        codes = ecount_client.search_products_by_name(q)
        if codes:
            return codes[:max_results]
    return []


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
        # 沒貨號 → 品名搜尋
        query = text.strip()
        for kw in _INFO_KW:
            query = query.replace(kw, "")
        query = query.strip()
        codes = _fuzzy_product_search(query)
        if not codes:
            return f"🔍 找不到「{query}」的相關產品資訊"

    in_stock = []
    out_of_stock = []
    last_code, last_name = None, None
    media_dir = _get_media_dir()
    for raw_code in codes:
        prod_code = raw_code.upper()
        raw_po = _get_raw_po_block(prod_code)
        try:
            item = ecount_client.lookup(prod_code)
        except Exception:
            continue
        prod_name = (item.get("name") if item else "") or prod_code
        qty = item.get("qty") if item else None

        has_po = raw_po is not None
        files = _match_product_media_files(prod_code, media_dir) if media_dir else []
        has_img = len(files) > 0

        # 價格
        _cache_item = ecount_client.get_product_cache_item(prod_code)
        _price = ""
        if _cache_item and _cache_item.get("price") and _cache_item["price"] > 0:
            _price = f"${int(_cache_item['price'])}"

        _name_line = f"  {prod_code}　{prod_name}　{_price}" if _price else f"  {prod_code}　{prod_name}"

        check_parts = []
        check_parts.append(f"PO文：{'✅' if has_po else '❌'}")
        check_parts.append(f"圖片：{'✅' + str(len(files)) + '張' if has_img else '❌'}")

        if qty is not None and qty > 0:
            check_parts.append(f"可售：{qty}")
            in_stock.append(f"{_name_line}\n  {'　'.join(check_parts)}")
        else:
            check_parts.append("缺貨")
            out_of_stock.append(f"{_name_line}\n  {'　'.join(check_parts)}")

        last_code, last_name = prod_code, prod_name

    lines = []
    if in_stock:
        lines.append(f"✅ 有庫存（{len(in_stock)} 筆）：")
        lines.extend(in_stock)
    if out_of_stock:
        lines.append(f"\n❌ 缺貨（{len(out_of_stock)} 筆）：")
        lines.extend(out_of_stock)

    return "\n".join(lines) if lines else None


def handle_internal_product_info_by_name(text: str, state_key: str | None = None) -> str | None:
    """
    純品名 fallback 查詢（不需關鍵字）：
      「大吉盒」→ 搜尋品名含「大吉盒」的產品 → 回傳 PO文 + 庫存（與「Z3524 資訊」相同格式）
    僅在所有其他 handler 都不匹配時觸發（dispatch chain 最後一項）。
    過濾：長度 < 2 或含明顯非查詢詞時回傳 None。
    """
    from storage.state import state_manager

    # 已含貨號 → 讓 handle_internal_product_info 處理
    if _PROD_CODE_RE.search(text):
        return None

    cleaned = text.strip()
    # 太短或太長不處理
    if len(cleaned) < 2 or len(cleaned) > 30:
        return None
    # 含數量詞 → 可能是訂單，不處理
    if re.search(r'\d+\s*(?:個|箱|件|盒|套|組)', cleaned):
        return None

    # 剝掉常見關鍵字，只留品名部分
    _strip_kw = _INFO_KW + _INV_QUERY_KW + _PREORDER_KW + _PRODUCT_LIST_KW + ["查詢", "查", "找", "搜尋", "？", "?"]
    query = cleaned
    for kw in _strip_kw:
        query = query.replace(kw, "")
    query = query.strip()
    if len(query) < 2:
        return None

    # 品名模糊搜尋
    codes = _fuzzy_product_search(query)
    if not codes:
        return None

    # 最多顯示 3 筆，避免過長
    results = []
    last_code, last_name = None, None
    for raw_code in codes[:3]:
        prod_code = raw_code.upper()
        po = _format_po(prod_code)
        try:
            item = ecount_client.lookup(prod_code)
        except Exception as e:
            results.append(f"⚠️ {prod_code}：查詢失敗（{e}）")
            continue
        prod_name = (item.get("name") if item else "") or prod_code
        stock_detail = _fmt_stock_lines(item, prod_code)
        results.append(f"{po}\n{stock_detail}")
        last_code, last_name = prod_code, prod_name

    if len(codes) > 3:
        results.append(f"⋯ 還有 {len(codes) - 3} 筆符合，請用編號查詢更精確")

    return "\n\n".join(results) if results else None


# ── 耗材前綴（HH008 系列）────────────────────────────────────────
_CONSUMABLE_PREFIX = "HH008-"
_CONSUMABLE_KW = ["耗材", "零件", "配件", "爪子", "爪套", "線圈", "馬達", "齒輪",
                  "投幣器", "電源供應器", "搖桿", "電眼", "螢幕", "燈條", "微動",
                  "滑輪", "線輪", "鎖頭", "按鈕", "變壓器"]
_CONSUMABLE_LIST_KW = ["耗材清單", "耗材產品", "耗材列表", "耗材有哪些", "零件清單", "配件清單",
                       "耗材庫存", "零件庫存", "配件庫存",
                       "耗材 清單", "耗材 產品", "耗材 列表", "耗材 有哪些",
                       "零件 清單", "配件 清單",
                       "耗材 庫存", "零件 庫存", "配件 庫存"]


def handle_internal_consumable(text: str, state_key: str | None = None) -> str | None:
    """
    內部群耗材查詢：
      「耗材清單」「耗材庫存」→ 列出全部 HH008 系列 + 庫存
      「投幣器 庫存」「爪子」  → 搜尋特定耗材
    """
    is_list = any(kw in text for kw in _CONSUMABLE_LIST_KW)
    # 特定耗材搜尋需要同時含耗材名 + 查詢詞（「爪子庫存」✓，單獨「爪子」✗）
    _query_suffix = ["庫存", "有嗎", "有沒有", "還有", "查詢", "多少", "幾個"]
    is_consumable = (any(kw in text for kw in _CONSUMABLE_KW)
                     and any(kw in text for kw in _query_suffix))

    if not is_list and not is_consumable:
        return None

    ecount_client._ensure_product_cache()

    if is_list:
        _t = text.replace(" ", "")  # 忽略空格
        is_stock_query = any(kw in _t for kw in ["耗材庫存", "零件庫存", "配件庫存"])
        items = [p for p in ecount_client._product_cache
                 if p["code"].upper().startswith(_CONSUMABLE_PREFIX)]
        if not items:
            return "目前沒有耗材資料"

        if is_stock_query:
            # 只列有庫存的
            in_stock = []
            for p in items:
                inv = ecount_client.lookup(p["code"])
                qty = inv.get("qty", 0) if inv else 0
                if qty > 0:
                    price = p.get("price")
                    price_str = f" ${int(float(price))}" if price else ""
                    in_stock.append(f"  {p['code']}　{p['name']}{price_str} 庫存:{qty}")
            if not in_stock:
                return "🔧 目前所有耗材都無庫存"
            lines = [f"🔧 有庫存的耗材（{len(in_stock)} 筆）："] + in_stock
        else:
            # 全部清單
            lines = [f"🔧 耗材清單（共 {len(items)} 筆）："]
            for p in items:
                inv = ecount_client.lookup(p["code"])
                qty = inv.get("qty", 0) if inv else 0
                price = p.get("price")
                price_str = f" ${int(float(price))}" if price else ""
                qty_str = f" 庫存:{qty}" if qty > 0 else " ⛔無庫存"
                lines.append(f"  {p['code']}　{p['name']}{price_str}{qty_str}")
        return "\n".join(lines)

    # 特定耗材搜尋
    keyword = text.strip()
    for kw in ["庫存", "有嗎", "有沒有", "還有", "查詢", "查"]:
        keyword = keyword.replace(kw, "")
    keyword = keyword.strip()

    if len(keyword) < 2:
        return None

    # 先在耗材裡搜
    matched = [p for p in ecount_client._product_cache
               if p["code"].upper().startswith(_CONSUMABLE_PREFIX)
               and keyword.upper() in p["name"].upper()]

    if not matched:
        return None

    lines = [f"🔧 「{keyword}」相關耗材（{len(matched)} 筆）："]
    for p in matched[:10]:
        inv = ecount_client.lookup(p["code"])
        qty = inv.get("qty", 0) if inv else 0
        price = p.get("price")
        price_str = f" ${int(float(price))}" if price else ""
        qty_str = f" 庫存:{qty}" if qty > 0 else " ⛔無庫存"
        lines.append(f"  {p['code']}　{p['name']}{price_str}{qty_str}")
    if len(matched) > 10:
        lines.append(f"  ⋯ 還有 {len(matched) - 10} 筆")
    return "\n".join(lines)


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
        stripped = _STRIP_KW_RE.sub(' ', text)
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

        in_stock_results  = []   # qty>0 的明細
        out_stock_results = []   # qty<=0 / None 的明細
        preorder_results  = []   # 預購模式結果
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
                    preorder_results.append(f"📦 {name}（{prod_code}）\n  可預購：{preorder} 個")
                    result_codes.append((prod_code, name))
            else:
                # 庫存模式：有可售庫存的放前面
                block = _fmt_inv_block(item, prod_code)
                result_codes.append((prod_code, item.get("name") or prod_code))
                if (item.get("qty") or 0) > 0:
                    in_stock_results.append(block)
                    stock_codes.append((prod_code, item.get("name") or prod_code))
                else:
                    out_stock_results.append(block)

        results = preorder_results + in_stock_results + out_stock_results

        kw_label = "".join(tokens)
        if not results:
            return f"🔍 找不到「{kw_label}」的庫存資料"
        print(f"[internal] 品名庫存搜尋「{kw_label}」→ {len(results)} 筆（其中 {len(stock_codes)} 筆 qty>0）", flush=True)

        # 決定用哪個清單來設 state：
        # 優先用有實際庫存（qty>0）的；若全是預購則用 result_codes
        set_codes = stock_codes if stock_codes else result_codes

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

    return "\n\n".join(results) if results else None


# ── 6. 價格查詢 ──────────────────────────────────────────────────────────

_PRICE_QUERY_RE = re.compile(
    r'([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*(?:價格|多少錢|售價|多少|賣多少)',
    re.IGNORECASE
)
_PRICE_RANGE_RE = re.compile(
    r'(\d+)\s*元\s*(?:以下|以內|內|below)',
    re.IGNORECASE
)


def _load_available_json() -> dict:
    """讀取 data/available.json，回傳 dict（失敗時回 {}）"""
    import json as _j
    from pathlib import Path as _P
    try:
        p = _P(__file__).parent.parent / "data" / "available.json"
        return _j.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _load_specs_json() -> dict:
    """讀取 data/specs.json，回傳 dict（失敗時回 {}）"""
    import json as _j
    from pathlib import Path as _P
    try:
        p = _P(__file__).parent.parent / "data" / "specs.json"
        return _j.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def handle_internal_price_query(text: str) -> str | None:
    """
    內部群價格查詢：
    1. 單品：「Z3191-1 價格」「Z3191-1 多少錢」
    2. 範圍：「150元以下產品」「150元以下庫存」

    價格優先順序：
      available.json unit_price → specs.json → Ecount API
    """
    # ── 1. 單品價格 ──────────────────────────────
    m = _PRICE_QUERY_RE.search(text)
    if m:
        prod_cd = m.group(1).upper()
        import httpx as _httpx
        from config import settings as _s

        name  = prod_cd
        price = 0.0

        # Step 0：從 available.json 讀出庫單價（離線也能用，最快）
        _avail = _load_available_json()
        _entry = _avail.get(prod_cd, {})
        _up = float(_entry.get("unit_price") or 0)
        if _up > 0:
            price = _up
        qty = _entry.get("available")  # 同時取可售庫存

        # 品名從產品快取取
        _ci = ecount_client.get_product_cache_item(prod_cd)
        if _ci and _ci.get("name"):
            name = _ci["name"]

        # Step 1：specs.json 補價格（若 available.json 無單價）
        if not price:
            _sp = _load_specs_json().get(prod_cd) or {}
            _nums = re.findall(r'\d+', str(_sp.get("price", "0")))
            if _nums:
                price = int(_nums[0])
            if (not name or name == prod_cd) and _sp.get("name"):
                name = _sp["name"]

        # Step 2：Ecount API（前兩步都無價格才呼叫，避免非營業時間 412）
        if not price:
            try:
                _sid = ecount_client._ensure_session()
                _resp = _httpx.post(
                    _s.ECOUNT_BASE_URL + '/OAPI/V2/InventoryBasic/GetBasicProductsList',
                    params={'SESSION_ID': _sid},
                    json={'PROD_CD': prod_cd},
                    timeout=10,
                )
                _data = ecount_client._safe_json(_resp)
                _r = (_data.get('Data') or {}).get('Result') or []
                if _r:
                    if not name or name == prod_cd:
                        name = _r[0].get('PROD_DES') or prod_cd
                    _op = float(_r[0].get('OUT_PRICE') or 0)
                    if _op > 0:
                        price = _op
            except Exception:
                pass

        # qty 補充（若 available.json 沒有，試 lookup）
        if qty is None:
            _item = ecount_client.lookup(prod_cd)
            if _item:
                qty = _item.get("qty")
                if (not name or name == prod_cd) and _item.get("name"):
                    name = _item["name"]

        stock_str = f"\n  可售庫存：{'缺貨' if not qty or qty <= 0 else f'{qty} 個'}"
        price_str = f"{int(price)} 元" if price else "未設定"
        return f"💰 {name}（{prod_cd}）\n  出庫單價：{price_str}{stock_str}"

    # ── 2. 價格範圍 ──────────────────────────────
    mr = _PRICE_RANGE_RE.search(text)
    if not mr:
        return None
    limit = int(mr.group(1))
    want_stock = any(kw in text for kw in ["庫存", "有貨", "有庫存"])

    # 偵測台型關鍵字（中巨／巨無霸／標準／K霸 等）
    matched_machine = next((m for m in _MACHINE_TYPES if m in text), None)

    avail = _load_available_json()
    specs = _load_specs_json()
    results = []

    if matched_machine:
        # ── 台型模式：先從規格DB找適合此台型的產品，再過濾價格/庫存 ──
        from storage.specs import get_by_machine as _gbm
        machine_specs = _gbm(matched_machine)
        for s in machine_specs:
            code = s.get("code", "").upper()
            if not code:
                continue
            # 取出庫單價（available.json 優先，再 specs.json price 欄）
            up = float((avail.get(code) or {}).get("unit_price") or 0)
            if not up:
                nums = re.findall(r'\d+', str(s.get("price", "0")))
                up = int(nums[0]) if nums else 0
            if up <= 0 or up > limit:
                continue
            name = s.get("name", code)
            av_qty = (avail.get(code) or {}).get("available")
            if want_stock:
                if not av_qty or av_qty <= 0:
                    continue
                results.append((up, f"  {code}  {int(up)}元  {name}（庫存 {av_qty} 個）"))
            else:
                qty_str = f"（庫存 {av_qty} 個）" if av_qty is not None else ""
                results.append((up, f"  {code}  {int(up)}元  {name}{qty_str}"))
        machine_label = f"【{matched_machine}台型】"
    else:
        # ── 全品項模式：合併 available.json + specs.json ──
        machine_label = ""
        all_codes: set[str] = set(avail.keys()) | set(specs.keys())
        for code in sorted(all_codes):
            up = float((avail.get(code) or {}).get("unit_price") or 0)
            if not up:
                sp = specs.get(code) or {}
                nums = re.findall(r'\d+', str(sp.get("price", "0")))
                up = int(nums[0]) if nums else 0
            if up <= 0 or up > limit:
                continue
            _ci = ecount_client.get_product_cache_item(code)
            name = (_ci or {}).get("name") or (specs.get(code) or {}).get("name") or code
            av_qty = (avail.get(code) or {}).get("available")
            if want_stock:
                if not av_qty or av_qty <= 0:
                    continue
                results.append((up, f"  {code}  {int(up)}元  {name}（庫存 {av_qty} 個）"))
            else:
                results.append((up, f"  {code}  {int(up)}元  {name}"))

    # 依單價由高到低排序
    results.sort(key=lambda x: x[0], reverse=True)
    lines = [r[1] for r in results]

    if not lines:
        suffix = "（有庫存）" if want_stock else ""
        return f"找不到 {machine_label}{limit} 元以下的產品{suffix}"

    header = f"💰 {machine_label}{limit} 元以下{'有庫存' if want_stock else ''}產品，共 {len(lines)} 筆："
    note = f"\n（只顯示前 20 筆，共 {len(lines)} 筆）" if len(lines) > 20 else ""
    return header + "\n" + "\n".join(lines[:20]) + note


# ── 7. 分類標籤推送 ─────────────────────────────────────────────────────
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
            url = f"{base}/{f.name}"
            stem_to_img_url[f.stem.upper()] = url
            if not any_img_url:
                any_img_url = url

    messages = []
    for f in files:
        ext = f.suffix.lower()
        media_url = f"{base}/{f.name}"
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
    line_api, uid: str, text_msg: TextMessage, media_msgs: list,
    prod_code: str | None = None,
) -> None:
    """
    單次 push：text + 最多 4 media（LINE 限 5 則/次），多的不送。
    prod_code 有給且 push response 帶 sent_messages → 記錄圖片 msg_id → 貨號，
    讓客戶日後 tag 回覆這張圖時 bot 能辨識是哪個產品（跟 reply 路徑行為一致）。
    """
    media_batch = media_msgs[:4]
    batch = [text_msg] + media_batch
    resp = line_api.push_message(PushMessageRequest(to=uid, messages=batch))
    if prod_code and media_batch and hasattr(resp, 'sent_messages') and resp.sent_messages:
        from main import _store_sent_image_ids
        _store_sent_image_ids(resp.sent_messages, [prod_code] * len(media_batch))


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


# ---------------------------------------------------------------------------
# 產品圖片查詢：「圖片 Z3555」「照片 T1202」
# ---------------------------------------------------------------------------
_PHOTO_RE = re.compile(
    rf'^(?:圖片|照片|圖)\s*(.+)',
    re.DOTALL
)


def handle_internal_product_photo(text: str, line_api) -> str | None:
    """
    「圖片 Z3555」→ 推送產品照片到內部群。
    """
    m = _PHOTO_RE.match(text.strip())
    if not m:
        return None
    codes = _PROD_CODE_RE.findall(m.group(1))
    if not codes:
        return None

    ngrok_url = _get_ngrok_url()
    media_dir = _get_media_dir()
    if not media_dir:
        return "⚠️ 產品照片磁碟機未連線"

    prod_code = codes[0].upper()
    files = _match_product_media_files(prod_code, media_dir)
    if not files:
        return f"❌ {prod_code} 無照片"
    base = ngrok_url.rstrip("/")
    image_urls = [f"{base}/{f.name}" for f in files[:4] if f.suffix.lower() in _IMG_EXTS]
    if image_urls:
        return (f"📷 {prod_code} 共 {len(files)} 張", image_urls[:4])
    return f"📷 {prod_code} 共 {len(files)} 張（無圖片格式）"


# ---------------------------------------------------------------------------
# 產品圖文查詢：「Z3555圖文」「Z3555 T1202圖文」
# ---------------------------------------------------------------------------
_PO_PHOTO_KW = ["圖文"]


def handle_internal_product_po_photo(text: str, line_api) -> str | None:
    """
    「Z3555圖文」→ 推送 PO 文 + 圖片到內部群。支援多產品。
    """
    if not any(kw in text for kw in _PO_PHOTO_KW):
        return None
    codes = _PROD_CODE_RE.findall(text)
    if not codes:
        return None

    ngrok_url = _get_ngrok_url()
    media_dir = _get_media_dir()

    prod_code = codes[0].upper()
    raw_po = _get_raw_po_block(prod_code)
    files = _match_product_media_files(prod_code, media_dir) if media_dir else []

    if not raw_po and not files:
        return f"❌ {prod_code} 無 PO 文、無圖片"
    if not raw_po:
        return f"❌ {prod_code} 無 PO 文"
    if not files:
        return f"❌ {prod_code} 無圖片"

    base = ngrok_url.rstrip("/")
    image_urls = [f"{base}/{f.name}" for f in files[:4] if f.suffix.lower() in _IMG_EXTS]
    if image_urls:
        return (raw_po, image_urls[:4])
    return raw_po


def _get_ngrok_url() -> str:
    """取得公開 HTTPS 網址（DuckDNA）"""
    return "https://xmnline.duckdns.org/product-photo"


# ---------------------------------------------------------------------------
# 同業比價：「比價 Z3754」→ 對 dingshang.com.tw 抓相似品名+尺寸的競品價格
# ---------------------------------------------------------------------------
_COMPETITOR_PRICE_RE = re.compile(r'^比[價对対]\s*([A-Za-z]\d{3,5}[A-Za-z\d-]*)$', re.I)


# ---------------------------------------------------------------------------
# 採購單：「採購單<換行>Z3754*50<換行>T1138*20」→ dry-run 預覽
#         「送出」（5 分鐘內）→ 真送總公司訂貨單 API
# ---------------------------------------------------------------------------
_PURCHASE_TRIGGER_RE = re.compile(r'^採購單\s*$', re.MULTILINE)
# 行格式：「Z3754*50」或「Z3754*50 不要黑色」（空白後到行尾為該品項摘要）
_PURCHASE_LINE_RE = re.compile(rf'({_PROD_CODE_PAT})\s*[\*xX×]\s*(\d+)(?:\s+(\S.*))?', re.IGNORECASE)


def _parse_purchase_lines(body: str) -> list[dict]:
    """從 body 抓出 [{"prod_cd", "qty", "note"}, ...]，支援 *、x、X、× 為分隔符
    每行可選用 `#備註` 結尾標 per-item 摘要"""
    items = []
    seen_codes = set()
    for line in body.splitlines():
        line = line.strip()
        if not line:
            continue
        m = _PURCHASE_LINE_RE.search(line)
        if not m:
            continue
        code = m.group(1).upper()
        if code in seen_codes:
            continue  # 同貨號去重，第一筆為準
        seen_codes.add(code)
        try:
            qty = int(m.group(2))
        except ValueError:
            continue
        if qty <= 0:
            continue
        note = (m.group(3) or "").strip()
        items.append({"prod_cd": code, "qty": qty, "note": note})
    return items


def _run_async(coro, timeout: int = 120):
    """在 sync handler 內跑 async coroutine（dispatch chain 是 sync）"""
    import asyncio as _aio
    import threading as _th
    result = {"value": None, "error": None}

    def _runner():
        try:
            result["value"] = _aio.run(coro)
        except Exception as e:
            result["error"] = e

    t = _th.Thread(target=_runner, daemon=True)
    t.start()
    t.join(timeout=timeout)
    if result["error"]:
        raise result["error"]
    return result["value"]


def handle_internal_purchase_order(text: str) -> str | None:
    """
    「採購單<換行>Z3525*50 不要黑色<換行>Z3245*20 訂單30」
    → 直接執行：
        1. Chrome 建小蠻牛採購單（供應商 10003）
        2. 成功才送總公司訂貨單 API（客戶 A05 / 倉庫 200）
    → 回覆「採購單建立成功 / 總部訂單建立成功」（含單號）
    """
    stripped = text.strip()
    if not stripped.startswith("採購單"):
        return None
    body = "\n".join(stripped.splitlines()[1:]).strip()
    if not body:
        return (
            "📋 採購單格式：\n"
            "採購單\n"
            "Z3525*50 不要黑色\n"
            "Z3245*20 訂單30"
        )

    items = _parse_purchase_lines(body)
    if not items:
        return "❌ 採購單沒解析到任何品項，格式：Z3525*50 不要黑色"

    total_qty = sum(it["qty"] for it in items)
    header = f"📋 採購單處理中（{len(items)} 品 / {total_qty} 件）..."
    print(f"[purchase-order] {header} items={items}", flush=True)

    # ── Step 1: Chrome 建小蠻牛採購單 ──
    async def _do_chrome():
        from services.ecount_chrome import get_ecount_page, create_purchase_order
        async with get_ecount_page() as page:
            if not page:
                return {"ok": False, "slip_no": None, "error": "Chrome 連不上"}
            return await create_purchase_order(page, items=items, remark="LINE 採購單")

    try:
        chrome_res = _run_async(_do_chrome(), timeout=120)
    except Exception as e:
        chrome_res = {"ok": False, "slip_no": None, "error": f"Chrome exception: {e}"}

    if not chrome_res or not chrome_res.get("ok"):
        return (
            f"❌ 小蠻牛採購單建立失敗（未送總公司）：\n"
            f"   {chrome_res.get('error') if chrome_res else '未知錯誤'}"
        )

    # ── Step 2: API 送總公司訂貨單 ──
    from services.ecount_hq import ecount_hq
    api_res = ecount_hq.save_purchase_order(items=items, remark="LINE 採購單", dry_run=False)

    lines = [f"📋 採購單已建立（{len(items)} 品 / {total_qty} 件）"]
    lines.append(f"1️⃣ 小蠻牛採購單：✅ {chrome_res.get('slip_no')}")
    if api_res["ok"]:
        lines.append(f"2️⃣ 總公司訂貨單：✅ 預留 {api_res['slip_no']}（實際以後台為準）")
    else:
        lines.append(f"2️⃣ 總公司訂貨單：❌ {api_res['error']}")
    return "\n".join(lines)


def handle_internal_competitor_price(text: str) -> str | None:
    """「比價 Z3754」→ 比對小蠻牛 vs dingshang 同業相似商品價格"""
    m = _COMPETITOR_PRICE_RE.match(text.strip())
    if not m:
        return None
    code = m.group(1).upper()

    from services.competitor_match import find_similar_by_code, format_for_line
    import json as _json
    from pathlib import Path as _P
    specs_path = _P(__file__).resolve().parents[1] / "data" / "specs.json"
    try:
        specs = _json.loads(specs_path.read_text(encoding="utf-8"))
    except Exception:
        specs = {}
    sp = specs.get(code)
    if not sp:
        return f"❌ 找不到貨號 {code} 的規格"

    matches = find_similar_by_code(code, limit=5)
    return format_for_line(
        spec_code=code,
        spec_name=sp.get("name", ""),
        spec_price=sp.get("price", "未知"),
        matches=matches,
    )


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
    skipped: list[str] = []
    for prod_code, prod_name in products:
        raw_po = _get_raw_po_block(prod_code)
        if not raw_po:
            skipped.append(f"{prod_name}（{prod_code}）無 PO 文")
            continue
        media_msgs = []
        files = []
        if ngrok_url and media_dir:
            files = _match_product_media_files(prod_code, media_dir)
            media_msgs = _build_media_messages(prod_code, files, ngrok_url)
            print(f"[internal-tag-push] {prod_code} 媒體檔案 {len(files)} 個 → {len(media_msgs)} 則訊息")
        if not files:
            skipped.append(f"{prod_name}（{prod_code}）無圖片")
            continue
        po_text = _format_po(prod_code)
        prod_data.append((prod_code, prod_name, po_text, media_msgs))

    if not prod_data:
        msg = "❌ 所有產品都不符合推送條件（需有 PO 文 + 圖片）"
        if skipped:
            msg += "\n⚠️ 跳過：\n" + "\n".join(f"  • {s}" for s in skipped)
        return msg

    sent = 0
    failed = 0
    for cust in customers:
        uid = cust.get("line_user_id")
        if not uid:
            continue
        import random
        greeting_prefix = random.choice([
            "老闆您好~給您送新品來了~\n\n",
            "哈囉~老闆~挖來啊~給你送新品來了!!\n\n",
            "老闆好~新品到了唷~快來看看~\n\n",
            "嗨~老闆~又有好東西來了!!\n\n",
            "老闆~新品來報到囉~看看有沒有喜歡的~\n\n",
            "老闆您好~熱騰騰的新品來了~\n\n",
            "嘿~老闆~新品上架啦~來看看吧~\n\n",
            "老闆~好東西來了~快來瞧瞧~\n\n",
        ])

        try:
            # 第一款加上問候語，其後各款不重複問候
            import time
            for i, (prod_code, prod_name, po_text, media_msgs) in enumerate(prod_data):
                if i > 0:
                    time.sleep(3)
                prefix = greeting_prefix if i == 0 else ""
                text_msg = TextMessage(text=prefix + po_text)
                _push_messages_chunked(line_api, uid, text_msg, media_msgs, prod_code=prod_code)
            sent += 1
            codes_str = "、".join(c for c, *_ in prod_data)
            print(f"[internal-tag-push] {tag}／{codes_str} → {cust.get('name') or uid}")
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
    if skipped:
        result += "\n⚠️ 跳過：\n" + "\n".join(f"  • {s}" for s in skipped)
    return result


# ---------------------------------------------------------------------------
# 看貨群推送
# ---------------------------------------------------------------------------

_SHOWCASE_TRIGGER = "看貨群"
_SHOWCASE_GROUP_NAME = "小蠻牛新北-新品看貨群"
_CONTACT_GROUP_TRIGGER = "聯絡群組"


def handle_internal_contact_group_push(text: str, line_api) -> str | None:
    """
    「聯絡群組 Z3456 T1122」→ 把 PO 文+圖片推送到所有名稱含「聯絡群組」的聊天室。
    """
    if _CONTACT_GROUP_TRIGGER not in text:
        return None
    _cg_text = text.replace(_CONTACT_GROUP_TRIGGER, "", 1).strip()
    _cg_codes = _PROD_CODE_RE.findall(_cg_text)
    if not _cg_codes:
        return None

    import threading as _t_cg
    _cg_codes_upper = list(dict.fromkeys(c.upper() for c in _cg_codes))

    def _do_contact_push():
        from services.line_oa_chat import send_many_to_chat_sync
        from config import settings as _cfg_cg
        chat_names = _cfg_cg.contact_group_chats_list()
        if not chat_names:
            print(f"[contact-group] 未設定 CONTACT_GROUP_CHATS", flush=True)
            return
        print(f"[contact-group] 推送到 {len(chat_names)} 個聯絡群組：{chat_names}", flush=True)

        # 預先準備每個貨號的 PO 文 + 圖片（只讀一次）
        media_dir = _get_media_dir()
        payloads = []
        for code in _cg_codes_upper:
            po_text = _format_po(code)
            if not po_text:
                print(f"[contact-group] ❌ {code} 無 PO 文，跳過", flush=True)
                continue
            img_paths = []
            if media_dir:
                files = _match_product_media_files(code, media_dir)
                img_paths = [str(media_dir / f) for f in files[:4]]
            payloads.append((code, po_text, img_paths))

        if not payloads:
            print(f"[contact-group] 無可推送貨號", flush=True)
            return

        # 群組優先：每個群組開一次聊天室，連發所有貨號
        items = [
            {"text": po_text, "image_paths": img_paths}
            for _, po_text, img_paths in payloads
        ]
        for cname in chat_names:
            results = send_many_to_chat_sync(cname, items, delay_sec=2.0)
            for (code, _po, _imgs), ok in zip(payloads, results):
                mark = "✅" if ok else "❌"
                print(f"[contact-group] {mark} {cname} ← {code}", flush=True)

    _t_cg.Thread(target=_do_contact_push, daemon=True).start()
    codes_str = "、".join(_cg_codes_upper)
    return f"📣 正在推送到所有聯絡群組：{codes_str}\n（透過 LINE OA，不消耗 API 額度）"


_RECOMMEND_TRIGGER = "推薦"


def handle_internal_recommend_push(text: str, line_api) -> str | None:
    """
    「推薦 武林 R0135 Z3353」→ 在 LINE OA 用「武林」搜尋聊天室，連發 PO 文+圖片（多則）。
    與「聯絡群組」指令差別：目標是單一自訂聊天室（搜尋詞模糊比對），不是設定檔裡的清單。
    """
    _strip = text.strip()
    if not _strip.startswith(_RECOMMEND_TRIGGER):
        return None
    _body = _strip[len(_RECOMMEND_TRIGGER):].strip()
    _first = _PROD_CODE_RE.search(_body)
    if not _first:
        return None
    _search_term = _body[:_first.start()].strip()
    if not _search_term:
        return None
    _codes_upper = list(dict.fromkeys(c.upper() for c in _PROD_CODE_RE.findall(_body)))
    if not _codes_upper:
        return None

    import threading as _t_rec

    def _do_recommend_push():
        from services.line_oa_chat import send_many_to_chat_sync
        media_dir = _get_media_dir()
        payloads = []
        for code in _codes_upper:
            po_text = _format_po(code)
            if not po_text:
                print(f"[recommend-push] ❌ {code} 無 PO 文，跳過", flush=True)
                continue
            img_paths = []
            if media_dir:
                files = _match_product_media_files(code, media_dir)
                img_paths = [str(media_dir / f) for f in files[:4]]
            payloads.append((code, po_text, img_paths))

        if not payloads:
            print(f"[recommend-push] 無可推送貨號", flush=True)
            return

        items = [
            {"text": po_text, "image_paths": img_paths}
            for _, po_text, img_paths in payloads
        ]
        results = send_many_to_chat_sync(_search_term, items, delay_sec=2.0)
        for (code, _po, _imgs), ok in zip(payloads, results):
            mark = "✅" if ok else "❌"
            print(f"[recommend-push] {mark} {_search_term} ← {code}", flush=True)

    _t_rec.Thread(target=_do_recommend_push, daemon=True).start()
    codes_str = "、".join(_codes_upper)
    return f"📣 正在推送到「{_search_term}」：{codes_str}\n（透過 LINE OA，不消耗 API 額度）"


def handle_internal_showcase_push(text: str, line_api) -> str | None:
    """
    偵測指令：
    - 「最近新品」 → 列出最近 14 天新品
    - 「看貨群 T1122 Z3456」→ 透過 LINE OA Chrome 推送 PO文+圖片到看貨群
    """
    _t_strip = text.strip()
    # 「最近新品」→ 列新品
    if _t_strip == "最近新品":
        _sc_codes = []
        _sc_text = ""
    else:
        if _SHOWCASE_TRIGGER not in text:
            return None
        # 提取產品代碼
        _sc_text = text.replace(_SHOWCASE_TRIGGER, "").strip()
        _sc_codes = _PROD_CODE_RE.findall(_sc_text)
        # 「看貨群」無貨號 → 不處理
        if not _sc_codes:
            return None

    # ── 有產品代碼 → 推送到看貨群 ──
    if _sc_codes:
        import threading as _t_sc
        _sc_codes_upper = list(dict.fromkeys(c.upper() for c in _sc_codes))

        def _do_showcase_push():
            from services.line_oa_chat import send_many_to_chat_sync

            # 準備所有 (text, images) — 無 PO 文的跳過
            items = []
            code_for_item: list[str] = []
            failed = []
            media_dir = _get_media_dir()
            for code in _sc_codes_upper:
                po_text = _format_po(code)
                if not po_text:
                    failed.append(f"{code}（無 PO 文）")
                    continue
                img_paths = []
                if media_dir:
                    files = _match_product_media_files(code, media_dir)
                    img_paths = [str(media_dir / f) for f in files[:4]]
                items.append({"text": po_text, "image_paths": img_paths})
                code_for_item.append(code)

            if not items:
                print(f"[showcase] ⚠️ 無可推送項目（全部缺 PO 文）", flush=True)
                return

            # 一次開啟聊天室、連續發送
            results = send_many_to_chat_sync(_SHOWCASE_GROUP_NAME, items, delay_sec=2.0)
            for code, ok in zip(code_for_item, results):
                if ok:
                    print(f"[showcase] ✅ {code} 已推送到看貨群", flush=True)
                else:
                    print(f"[showcase] ❌ {code} 推送失敗", flush=True)

        # 背景執行推送
        _t_sc.Thread(target=_do_showcase_push, daemon=True).start()
        codes_str = "、".join(_sc_codes_upper)
        return f"📣 正在推送到看貨群：{codes_str}\n（透過 LINE OA，不消耗 API 額度）"

    # ── 無產品代碼 → 列出新品 ──
    import json as _json_sc
    from datetime import datetime, timedelta
    from pathlib import Path as _Path_sc

    np_path = _Path_sc(__file__).parent.parent / "data" / "new_products.json"
    if not np_path.exists():
        return "📋 目前沒有新品資料（尚未偵測到新品項）"

    try:
        all_products = _json_sc.loads(np_path.read_text(encoding="utf-8"))
    except Exception:
        return "❌ 讀取新品資料失敗"

    cutoff = datetime.now() - timedelta(days=30)
    recent = []
    for code, info in all_products.items():
        first_seen = info.get("first_seen", "")
        try:
            dt = datetime.strptime(first_seen, "%Y-%m-%d %H:%M")
        except (ValueError, TypeError):
            continue
        if dt >= cutoff:
            recent.append((code, info.get("name", ""), info.get("price", 0), dt))

    if not recent:
        return "📋 最近 30 天沒有新增品項"

    # 補上可售庫存、圖片、PO 文狀態
    media_dir = _get_media_dir()
    enriched = []
    for code, name, price, dt in recent:
        _stk = ecount_client.lookup(code)
        qty = int(_stk["qty"]) if (_stk and _stk.get("qty") is not None) else -1
        has_img = bool(_match_product_media_files(code, media_dir)) if media_dir else False
        has_po  = _get_raw_po_block(code) is not None
        enriched.append((code, name, price, dt, qty, has_img, has_po))

    # 依可售庫存由高到低（無庫存資訊排最後）
    enriched.sort(key=lambda x: (x[4] if x[4] >= 0 else -1), reverse=True)

    lines = [f"📋 最近 30 天新品（共 {len(enriched)} 筆，依可售庫存排序）："]
    for code, name, price, dt, qty, has_img, has_po in enriched:
        price_str = f"　${int(price)}" if price else ""
        qty_str = f"　可售{qty}" if qty >= 0 else "　可售-"
        mark = f"[{'圖' if has_img else '✗圖'}/{'文' if has_po else '✗文'}]"
        lines.append(f"  {mark} {code}　{name}{price_str}{qty_str}")

    return "\n".join(lines)


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
_SAVE_IMG_RE = re.compile(
    r'(?:存圖\s*([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)|([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*存圖)',
    re.IGNORECASE,
)

# 加圖指令正則：「加圖 Z3432」（保留舊圖，追加新圖）
_ADD_IMG_RE  = re.compile(
    r'(?:(?:加圖|補圖)\s*([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)|([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*(?:加圖|補圖))',
    re.IGNORECASE,
)

# 秒殺指令正則：「沒貨 Z3432」標記、「有貨 Z3432」取消
_SOLD_OUT_RE = re.compile(
    r'(?:沒貨\s*([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)|([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*沒貨)',
    re.IGNORECASE,
)
_RESTOCK_RE = re.compile(
    r'(?:有貨\s*([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)|([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)\s*有貨)',
    re.IGNORECASE,
)

# Session 觸發詞與結束詞（「存圖」單獨傳也進 session；含貨號時走單品路徑）
_UPLOAD_TRIGGERS  = {"上架", "存檔", "存圖"}
_UPLOAD_FINISH_RE = re.compile(r'^(完成|好了|結束|done|finish)$', re.IGNORECASE)
_UPLOAD_CANCEL_RE = re.compile(r'^(取消|cancel)$', re.IGNORECASE)


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
        # 刪除空白行（保留單行換行）
        new_block = "\n".join(line for line in content.splitlines() if line.strip())

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
        from scripts.import_specs import parse_specs, _enrich_from_ecount, OUTPUT, SOURCE
        import json as _json

        # 1. 同步解析 PO文，用 Ecount 覆蓋品名/價格，更新 specs.json
        try:
            exists = SOURCE.exists()
        except OSError:
            exists = False
        if exists:
            specs = parse_specs(SOURCE.read_text(encoding="utf-8"))
            specs = _enrich_from_ecount(specs)
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

        # 2. 找出哪些 code 在 specs 裡找不到或 Ecount 無品名
        ecount_client._cache_expires = 0  # 強制刷新，確保新增品項後能查到
        result["no_name"] = []
        for c in codes:
            uc = c.upper()
            if uc not in specs:
                result["missing"].append(uc)
            else:
                # 有規格但 Ecount 無品名 → 需先新增品項
                item = ecount_client.lookup(uc)
                if not item or not (item.get("name") or "").strip():
                    result["no_name"].append(uc)

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


# ══════ 指令 2c：沒貨/有貨 Z3432（秒殺標記/取消）═══════════════════════
def handle_internal_mark_sold_out(code: str) -> str:
    """指令「沒貨 Z3432」：標記為秒殺，客戶下單會被擋"""
    from storage import sold_out
    code = code.strip().upper()
    item = ecount_client.get_product_cache_item(code)
    name = (item.get("name") if item else None) or code
    sold_out.add(code)
    return f"✅ {code} {name} → 已標記秒殺，客戶下單會被擋"


def handle_internal_unmark_sold_out(code: str) -> str:
    """指令「有貨 Z3432」：恢復可訂"""
    from storage import sold_out
    code = code.strip().upper()
    item = ecount_client.get_product_cache_item(code)
    name = (item.get("name") if item else None) or code
    if sold_out.remove(code):
        return f"✅ {code} {name} → 已恢復可訂"
    return f"⚠️ {code} {name} 原本就沒標記秒殺"


# ══════ 指令 3：存文 ════════════════════════════════════════════════════

def _split_po_by_code(text: str) -> list[str]:
    """
    將多段 PO文分割為獨立筆記。

    優先：空白行分段（同一訊息內按多次 Enter）→ 最自然，貨號位置不限。
    備援：無空白行時（分開送的訊息被合併），逐行掃描貨號，貨號換了就換段。
          用 search（非 match），貨號可在行中任意位置（如「編號：T1198」）。
    """
    # 只有一個貨號 → 不拆分（整段就是一筆 PO文）
    all_codes = list(dict.fromkeys(c.upper() for c in _PROD_CODE_RE.findall(text)))
    if len(all_codes) <= 1:
        return [text] if all_codes else []

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
    existing = state_manager.get(user_id)
    if existing and existing.get("action") == "uploading":
        # 已有進行中的上架 session
        n = len(existing.get("groups", []))
        has_current = bool(existing.get("current_code") or existing.get("current_media"))
        total = n + (1 if has_current else 0)
        return f"⚠️ 你已有上架作業進行中（{total} 組），請先傳「完成」結束目前的作業"
    if existing and existing.get("action") == "new_product_session":
        return "⚠️ 你正在新增品項中，請先傳「完成」結束目前的作業"
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


def _split_po_segments(text: str, codes: list[str]) -> list[tuple[str, str]]:
    """將含多個貨號的 PO文拆分為 [(code, segment_text), ...]"""
    lines = text.split('\n')
    # 找出每個貨號所在的行號
    code_line_map: dict[str, int] = {}
    for i, line in enumerate(lines):
        for m in _PROD_CODE_RE.finditer(line):
            c = m.group(1).upper()
            if c in codes and c not in code_line_map:
                code_line_map[c] = i

    # 按行號排序
    ordered = sorted(
        [(code_line_map.get(c, len(lines)), c) for c in codes],
        key=lambda x: x[0],
    )

    segments = []
    for idx, (line_idx, code) in enumerate(ordered):
        if idx == 0:
            start = 0
        else:
            prev_line = ordered[idx - 1][0]
            # 在前一個貨號行和本貨號行之間找分割點（中點偏後）
            mid = (prev_line + line_idx + 1) // 2
            start = mid
        if idx + 1 < len(ordered):
            next_line = ordered[idx + 1][0]
            mid = (line_idx + next_line + 1) // 2
            end = mid
        else:
            end = len(lines)
        seg = '\n'.join(lines[start:end]).strip()
        segments.append((code, seg))

    return segments


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
    # 排除品名行裡的假貨號（如「品名:野獸國 VPB-011SP」的 VPB-011）
    # 先找編號行的貨號，品名行裡出現但編號行沒出現的 → 假貨號
    _code_line_codes = set()
    _name_line_codes = set()
    for _line in combined.splitlines():
        _ls = _line.strip()
        if re.match(r'^(?:編號|貨號|產品編號|商品編號)[：:]', _ls):
            _code_line_codes.update(c.upper() for c in _PROD_CODE_RE_RAW.findall(_ls))
        elif re.match(r'^(?:品名|名稱|商品名|產品名)[：:]', _ls):
            _name_line_codes.update(c.upper() for c in _PROD_CODE_RE_RAW.findall(_ls))
    _fake_codes = _name_line_codes - _code_line_codes  # 只在品名出現、不在編號行的
    all_po_matches = list(_PROD_CODE_RE.finditer(combined))
    if all_po_matches:
        # 取得不重複的貨號（保持順序），排除假貨號
        seen = set()
        unique_codes = []
        for m in all_po_matches:
            c = m.group(1).upper()
            if c not in seen and c not in _fake_codes:
                seen.add(c)
                unique_codes.append(c)

        if len(unique_codes) > 1:
            # ── 多組 PO文合併在同一訊息 → 逐組拆分 ──
            segments = _split_po_segments(combined, unique_codes)
            cur_media = state.get("current_media", [])
            # 先存前一組（如有）
            if current_code and (cur_media or current_po):
                groups.append({
                    "code":  current_code,
                    "media": cur_media,
                    "po":    current_po,
                })
                state["current_media"] = []
            # 前 N-1 組直接存入 groups（media 為空，圖片會在 finish 時分配）
            for seg_code, seg_text in segments[:-1]:
                groups.append({
                    "code":  seg_code,
                    "media": [],
                    "po":    seg_text,
                })
            # 最後一組設為 current（後續圖片會進 current_media）
            last_code, last_text = segments[-1]
            state["groups"]        = groups
            state["current_code"]  = last_code
            state["current_po"]    = last_text
            state_manager.set(user_id, state)
            print(f"[upload] 偵測到 {len(unique_codes)} 組 PO文：{'、'.join(unique_codes)}", flush=True)
            return None

        # ── 單一貨號 ──
        code = unique_codes[0]
        cur_media = state.get("current_media", [])
        # 若貨號不同且前一組有內容 → 先存前一組
        if current_code and current_code != code and (cur_media or current_po):
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
    if label_result.get("no_name"):
        no_name_str = "、".join(label_result["no_name"])
        label_lines.append(f"⚠️ Ecount 無品名，標籤未生成：{no_name_str}")
        label_lines.append("請先「新增品項」建立品名，完成後自動加入標籤佇列")
    if label_result["error"]:
        label_lines.append(f"⚠️ 標籤生成錯誤：{label_result['error']}")

    # 上架完成後更新預購清單（PO文可能新增預購品）
    try:
        from handlers.inventory import refresh_preorder_list
        refresh_preorder_list()
    except Exception:
        pass

    suffix = "\n" + "\n".join(label_lines) if label_lines else ""
    return "🏁 上架完成！\n" + "\n".join(results) + suffix


def handle_internal_upload_cancel(user_id: str) -> str:
    """「取消」→ 丟棄 uploading session 累積的 PO文/媒體，清 state（不存任何東西）"""
    from storage.state import state_manager
    state = state_manager.get(user_id)
    if not state or state.get("action") != "uploading":
        return "沒有進行中的上架作業"
    groups = state.get("groups", [])
    has_current = bool(state.get("current_code") or state.get("current_media") or state.get("current_po"))
    n = len(groups) + (1 if has_current else 0)
    state_manager.clear(user_id)
    return f"❌ 已取消上架（清除 {n} 組未完成）" if n else "❌ 已取消上架"


# ── 新增品項 ───────────────────────────────────────────────────────────────
# 格式（單行或多行均支援）：
#   新增品項 Z9999 (原)多色麥克風音響 個 條碼:1234567890 售價:299 規格:30×20cm
#   新增品項 Z9999
#   品名：(大)多色麥克風音響
#   條碼：1234567890
#   售價：299
#   加盟商：250
#   規格：30×20cm
_NEW_PROD_TRIGGER_RE = re.compile(r'^(?:新增|新建)品項', re.IGNORECASE)
_NEW_PROD_CODE_RE    = re.compile(r'([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?|\d{5,6}(?:-\d+)?)')
_UNIT_WORDS_NP       = r'個|件|盒|套|箱|組|片|包|瓶|罐|條|支|只|枚|粒|顆|袋|塊'

# CLASS_CD 對應（品名前綴，按長到短排列避免短前綴先匹配）
_CLASS_CD_MAP = [
    (r'^[（(]原定[)）]', "00004"),
    (r'^[（(]定[)）]',   "00004"),
    (r'^[（(]原[)）]',   "00001"),
    (r'^[（(]大[)）]',   "00002"),
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

    # LINE emoji 數字殘留：(three)(nine) → 39；($) → $ 等
    _EN_DIGIT_MAP = {
        "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
        "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
    }
    def _emoji_num_sub(m):
        w = m.group(1).lower()
        return _EN_DIGIT_MAP.get(w, m.group(0))
    flat = re.sub(r'\(([A-Za-z]+)\)', _emoji_num_sub, flat)
    # ($) → $
    flat = flat.replace("($)", "$")
    # 數字 emoji alternate text：(1)(9)(9) → 199（LINE 有些 emoji 替代文字是 (N) 格式）
    flat = re.sub(r'\((\d)\)', r'\1', flat)
    # Unicode keycap emoji：1️⃣9️⃣9️⃣ → 199（移除 VS16 U+FE0F 和 combining enclosing keycap U+20E3）
    flat = re.sub(r'[\ufe0f\u20e3]', '', flat)
    # 全形數字 → 半形
    flat = flat.translate(str.maketrans("０１２３４５６７８９", "0123456789"))

    # 貨號（必填）
    m_code = _NEW_PROD_CODE_RE.search(flat)
    if not m_code:
        return None
    prod_cd = m_code.group(1).upper()

    # 條碼
    bar_code_m = re.search(r'條碼\s*[:：]?\s*(\S+)', flat)
    bar_code   = bar_code_m.group(1) if bar_code_m else prod_cd  # 預設條碼 = 品項編碼（貨號）

    # 售價 / 賣價 / 出庫單價
    # 優先：「單盒N元」「單個N元」格式（最明確）
    out_price_m = re.search(r'(?:單盒|单盒|單個|单个|每盒|每個|每个)\s*([\d.]+)\s*元', flat)
    if not out_price_m:
        # 特價/售價類關鍵字：容忍 emoji/中文/$，結尾可為「元」或直接數字（例：「超特價$129」「現在特價‼️299元」）
        # 必須是明確的特價/售價類，不能只是「價」（會誤配「價格：公司原價」）
        out_price_m = re.search(
            r'(?:'
            r'(?:超|大|限時|限定|現在|出清|下殺|促銷|活動|爆殺|破盤|超級|特大)?(?:特價|優惠價|優惠|特惠|預購價|預售價)'
            r'|售價|賣價|產品售價|出庫單價|批價|零售價'
            r'|現在只要|只要|最低價|只需'
            r')'
            r'[^\d\n]{0,10}?([\d]+(?:\.\d+)?)\s*(?:元)?(?!\S*起批)',
            flat,
        )
    if not out_price_m:
        # 「價格：」後方容忍中文（常有「原價/公司原價/廠商建議售價」等前綴）
        out_price_m = re.search(r'(?:價格|售)\s*[:：][^\d\n]{0,15}?([\d.]+)\s*元?', flat)
    if not out_price_m:
        # fallback：任意位置「$N」或「N元」（取第一個），但跳過「原價/建議售價/市價/定價/折前」前綴
        _ORIG_PRICE_RE = re.compile(r'(?:原價|建議售價|公司.{0,4}售價|市價|定價|折前)\s*[$＄]?\s*$')
        _fallback_m = None
        for _cm in re.finditer(r'[$＄]\s*([\d]+(?:\.\d+)?)|([\d]+(?:\.\d+)?)\s*元', flat):
            _before_ctx = flat[max(0, _cm.start()-10):_cm.start()]
            if _ORIG_PRICE_RE.search(_before_ctx):
                continue
            _fallback_m = _cm
            break
        if _fallback_m is not None:
            class _M:
                def __init__(self, v): self._v = v
                def group(self, n): return self._v
            out_price_m = _M(_fallback_m.group(1) or _fallback_m.group(2))
    if not out_price_m:
        # fallback：一行只有純數字，視為售價
        out_price_m = re.search(r'(?:^|\s)([\d.]+)(?=\s|$)', flat)
    out_price = out_price_m.group(1) if out_price_m else ""

    # 加盟商價格 / 入庫單價 / 批發價（按長到短排，避免短詞先匹配）
    in_price_m   = re.search(r'(?:加盟商價格|加盟商商價|加盟商價|加盟商|入庫單價|批發價|進價)\s*[:：]?\s*[$＄]?\s*([\d.]+)', flat)
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

    # 品名：優先從「品名：」「產品名稱：」「名稱：」標籤取值
    # 逐行掃描：若某行以標籤開頭，只取該行內容（避免吃到下一行描述）
    prod_name = ""
    for _ln in text.strip().splitlines():
        _ln_m = re.match(r'\s*(?:產品名稱|品名|名稱)\s*[:：]\s*(.+)', _ln)
        if _ln_m:
            prod_name = _ln_m.group(1).strip()
            break
    name_label_m = re.search(r'(?:產品名稱|品名|名稱)\s*[:：]\s*(.+?)(?=\s+(?:產品編號|編號|貨號|條碼|售價|賣價|限時特價|特價|價格|出庫|入庫|加盟|單位|規格|尺寸|重量|單品|建議|包裝)|$)', flat) if not prod_name else None
    if name_label_m:
        prod_name = name_label_m.group(1).strip()
    # 裝別前綴優先：flat 出現 (原)/(大)/(定)/(原定) → 直接當品名
    # 應對「描述1 描述2 (大)真品名 編號:XXX ...」格式（品名寫在貨號前且無 label）
    if not prod_name:
        _class_prefix_m = re.search(
            r'([（(](?:原定|原|大|定)[）)].+?)'
            rf'(?=\s+(?:產品編號|編號|貨號|條碼|售價|賣價|限時特價|特價|價格|出庫|入庫|加盟|單位|規格|尺寸|重量|單品|建議|包裝|{_UNIT_WORDS_NP})(?:\s|[:：$＄\d]|$)|$)',
            flat,
        )
        if _class_prefix_m:
            prod_name = _class_prefix_m.group(1).strip()
    if prod_name:
        pass  # 已透過 label / 裝別前綴取得
    else:
        # fallback：從貨號後開始，剝除所有已識別欄位
        name_part = flat[m_code.end():]
        _strip_pats = [
            r'條碼\s*[:：]?\s*\S+',
            r'(?:加盟商價格|加盟商商價|加盟商價|加盟商|入庫單價|批發價|進價)\s*[:：]?\s*[$＄]?\s*[\d.]+\s*元?',
            r'(?:產品售價|產品價格|限時特價|售價|賣價|價格|出庫單價|特價|批價|零售價|售)\s*[:：]?\s*[$＄]?\s*\(?\$?\)?\s*[\d.]+\s*(?:元|/\S*)?',
            r'規格\s*[:：]?\s*\S+(?:\s+\S+)*?(?=\s+(?:條碼|售價|賣價|出庫|入庫|加盟)|\s*$)',
            r'(?:產品|包裝)?尺寸\s*[-:：]?\s*約?\s*\S+',
            r'(?:單品|產品)?重量\s*[:：]?\s*約?\s*\S+',
            r'建議\s*[:：]?\s*\S+',
            rf'單位\s*[:：]?\s*(?:{_UNIT_WORDS_NP})',
            rf'(?:^|\s)(?:{_UNIT_WORDS_NP})(?:\s|$)',
            r'(?:產品名稱|品名|名稱)\s*[:：]?\s*',
            r'編號\s*[:：]?\s*\S+',
            r'(?:^|\s)[\d.]+\s*元?(?=\s|$)',  # 裸數字+元（售價 fallback）
        ]
        for strip_pat in _strip_pats:
            name_part = re.sub(strip_pat, ' ', name_part)
        prod_name = name_part.strip()

        # 貨號後面找不到品名 → 嘗試從貨號前面找（品名寫在編號前的情況）
        if not prod_name:
            # 去掉「新增品項」「新建品項」前綴和「編號：XXX」
            _before = flat[:m_code.start()]
            _before = re.sub(r'^(?:新增|新建)品項\s*', '', _before)
            _before = re.sub(r'(?:產品)?(?:編號|貨號)\s*[:：]?\s*$', '', _before)
            for strip_pat in _strip_pats:
                _before = re.sub(strip_pat, ' ', _before)
            prod_name = _before.strip()

    if not prod_name:
        return None  # 品名必填

    class_cd = _detect_class_cd(prod_name)
    in_price = _calc_in_price(class_cd, out_price, in_price_raw)

    # 抓尺寸和重量（寫入 REMARKS_WIN / CONT1）
    _sz_m = re.search(r'(?:包裝)?尺寸\s*[-:：]?\s*約?\s*(\S+)', flat)
    _wt_m = re.search(r'(?:單品|產品)?重量\s*[:：]?\s*約?\s*(\S+)', flat)

    return {
        "prod_cd":      prod_cd,
        "prod_name":    prod_name,
        "unit":         unit,
        "bar_code":     bar_code,
        "class_cd":     class_cd,
        "out_price":    out_price,
        "in_price":     in_price,
        "size_des":     size_des,
        "prod_size":    _sz_m.group(1) if _sz_m else "",
        "prod_weight":  _wt_m.group(1) if _wt_m else "",
        "cust":         "10003",
    }


_CLASS_LABEL_NP = {"00001": "原裝", "00002": "改裝", "00004": "定裝"}
# 以貨號開頭且後接空白的行（貨號+品名同行）
_PROD_LINE_START_RE = re.compile(r'^([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)[\s:：]', re.IGNORECASE)
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

    first = re.sub(r'^(?:新增|新建)品項\s*', '', lines[0], flags=re.IGNORECASE).strip()
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

    _prod_size   = fields.get("prod_size", "")
    _prod_weight = fields.get("prod_weight", "")

    extra: dict = {
        "PROD_TYPE": "3",
        "BAL_FLAG":  "1",
        "USE_FLAG":  "Y",
        "CUST":      "10003",
    }
    if bar_code:     extra["BAR_CODE"]     = bar_code
    if class_cd:     extra["CLASS_CD"]     = class_cd
    if out_price:    extra["OUT_PRICE"]    = out_price
    if in_price:     extra["IN_PRICE"]     = in_price
    if size_des:     extra["SIZE_DES"]     = size_des
    # 從規格提取裝箱數（如「20個/箱」「24/箱」）
    _box_m = re.search(r'(\d+)\s*(?:個|盒|入)?\s*/\s*(?:箱|件)', size_des)
    if _box_m:
        if unit == "箱":
            # 單位是箱 → 寫入規格（SIZE_DES），不寫 EXCH_RATE
            if not size_des:
                extra["SIZE_DES"] = _box_m.group(0)
        else:
            # 單位是個 → 寫入包裝數（EXCH_RATE）
            extra["EXCH_RATE"] = _box_m.group(1)
    if _prod_size:   extra["REMARKS_WIN"]  = _prod_size     # 尺寸 → REMARKS_WIN
    if _prod_weight: extra["CONT1"]        = _prod_weight   # 重量 → CONT1

    # 先檢查品項是否已存在（cache 命中先 force refresh 一次，避免 user 剛在 ecount 後台刪/改品項時 cache stale）
    _existing = ecount_client.get_product_cache_item(prod_cd)
    if _existing:
        ecount_client.force_refresh_product_cache()
        _existing = ecount_client.get_product_cache_item(prod_cd)
    already_existed = False
    existing_name = ""
    if _existing:
        existing_name = _existing.get("name", "") or ""
        print(f"[內部] 品項已存在，跳過新增: {prod_cd} ({existing_name})")
        already_existed = True
        ok = True
        error_msg = ""
        result = {"ok": True, "slip": prod_cd}
    else:
        result = ecount_client.save_product(
            prod_cd=prod_cd, prod_name=prod_name, unit=unit, extra=extra,
        )
        ok = isinstance(result, dict) and result.get("ok")
        error_msg = result.get("error", "") if isinstance(result, dict) else ""

    # 已存在時不要再 add to new_products_store / label queue（避免打錯貨號污染清單）
    label_result = {}
    if ok and not already_existed:
        # 標記品項快取過期，下次查詢時自動刷新（避免連續新增時重複呼叫 API）
        ecount_client._cache_expires = 0
        from storage.new_products import new_products_store
        new_products_store.add(
            prod_cd=prod_cd,   prod_name=prod_name, unit=unit,
            bar_code=bar_code, class_cd=class_cd,   out_price=out_price,
            in_price=in_price, size_des=size_des,   cust="10003",
        )
        # 自動嘗試加入架上標籤佇列（品項建立後補印）
        try:
            label_result = _generate_labels_sync([prod_cd])
            if label_result["pdfs"]:
                print(f"[label] 新增品項後自動生成標籤：{[p.name for p in label_result['pdfs']]}")
        except Exception as _le:
            print(f"[label] 新增品項後標籤處理失敗：{_le}")

    # 已存在時：比對既有品名跟使用者輸入，不一致就警告（典型 typo 撞到別的貨號）
    if already_existed:
        if existing_name and existing_name.strip() != prod_name.strip():
            return (
                f"⚠️ {prod_cd} 已存在為「{existing_name}」\n"
                f"   你輸入的是「{prod_name}」— 若打錯貨號請改用正確的貨號重送"
            )
        # 同名重送視為冪等，不動作也不刷 label
        return f"ℹ️ {prod_cd} {existing_name or prod_name}　已存在，未重新建立"

    icon = "✅" if ok else "❌"
    details = []
    details.append(f"售:{out_price}" if out_price else "售:-")
    details.append(f"入:{in_price}" if in_price else "入:-")
    if size_des:  details.append(f"規:{size_des}")
    if class_cd:  details.append(_CLASS_LABEL_NP.get(class_cd, class_cd))
    detail_str = "　" + "　".join(details) if details else ""
    line = f"{icon} {prod_cd} {prod_name}　{unit}{detail_str}"
    if not ok and error_msg:
        line += f"\n   ⚠️ 原因：{error_msg}"
    if ok and label_result.get("pdfs"):
        line += "\n   🏷️ 架上標籤已自動生成"
    elif ok and label_result and not label_result.get("missing"):
        line += "\n   📋 已加入標籤佇列"
    return line


def handle_internal_label_queue(text: str, state_key: str | None = None) -> str | None:
    """
    手動加入標籤佇列：
      「標籤 Z3594」         → 加入 1 個
      「標籤 Z3594 T1135 Z3555」 → 加入多個
    """
    t = text.strip()
    lines = t.splitlines()
    if not lines[0].strip().startswith("標籤"):
        return None

    remaining = t.replace("標籤", "").strip()
    codes = _PROD_CODE_RE.findall(remaining)
    if not codes:
        return "❌ 請指定產品編碼\n格式：標籤 Z3594 T1135 Z3555"

    codes = [c.upper() for c in codes]

    # 先檢查哪些有效（有規格+有品名）
    from scripts.generate_shelf_label import (
        _load_specs, _build_product_data, _generate_one_pdf,
        _load_queue, _save_queue, QUEUE_FILE, OUTPUT_DIR, _queue_lock,
    )
    from scripts.import_specs import parse_specs, OUTPUT as SPECS_OUTPUT, SOURCE as SPECS_SOURCE
    import json as _json2

    # 同步 specs
    try:
        if SPECS_SOURCE.exists():
            specs_data = parse_specs(SPECS_SOURCE.read_text(encoding="utf-8"))
            SPECS_OUTPUT.parent.mkdir(exist_ok=True)
            SPECS_OUTPUT.write_text(_json2.dumps(specs_data, ensure_ascii=False, indent=2), encoding="utf-8")
            import storage.specs as _spec_store2
            _spec_store2.reload()
        else:
            specs_data = _json2.loads(SPECS_OUTPUT.read_text(encoding="utf-8")) if SPECS_OUTPUT.exists() else {}
    except Exception:
        specs_data = {}

    # 強制刷新 Ecount 品項快取，確保新增品項後能立即查到
    ecount_client._cache_expires = 0

    valid = []
    missing = []
    no_name = []
    for c in codes:
        d = _build_product_data(c, specs_data)
        if d:
            valid.append(d)
        elif c not in specs_data:
            missing.append(c)
        else:
            no_name.append(c)

    result_lines = []
    pdfs = []

    if len(valid) >= 3:
        # ≥ 3 個：直接生成 PDF，不混入佇列
        from datetime import datetime as _dt2
        ts = _dt2.now().strftime("%Y%m%d")
        while len(valid) >= 3:
            batch = valid[:3]
            valid = valid[3:]
            codes_str = "_".join(p["商品編號"] for p in batch)
            out_path = OUTPUT_DIR / f"架上標_{ts}_{codes_str}.pdf"
            try:
                _generate_one_pdf(batch, out_path)
                pdfs.append(out_path)
            except Exception as e:
                result_lines.append(f"❌ PDF 生成失敗：{e}")
        # 剩餘不足 3 個的加入佇列
        if valid:
            with _queue_lock:
                queue = _load_queue()
                for p in valid:
                    queue.append(p)
                _save_queue(queue)
            result_lines.append(f"📋 {'、'.join(p['商品編號'] for p in valid)} 加入佇列，待湊滿 3 個")
    elif valid:
        # < 3 個：加入佇列
        label_result = _generate_labels_sync([p["商品編號"] for p in valid])
        pdfs = label_result.get("pdfs", [])
        added_codes = [p["商品編號"] for p in valid]
        if pdfs:
            pass  # 下面統一顯示
        else:
            result_lines.append(f"📋 {'、'.join(added_codes)} 加入佇列，待湊滿 3 個")

    if pdfs:
        names = "、".join(p.name for p in pdfs)
        result_lines.insert(0, f"🏷️ 架上標籤已生成：{names}")
    if missing:
        result_lines.append(f"⚠️ 規格缺失：{'、'.join(missing)}")
    if no_name:
        result_lines.append(f"⚠️ Ecount 無品名：{'、'.join(no_name)}")
    if not result_lines:
        result_lines.append("❌ 沒有產品可加入標籤佇列")

    return "\n".join(result_lines)


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


# ---------------------------------------------------------------------------
# 回饋金查詢
# ---------------------------------------------------------------------------

_REBATE_KW = ["回饋金資料", "回饋金"]
_REBATE_PUSH_KW = ["推送回饋金", "發送回饋金", "推回饋金", "回饋金推送", "回饋金通知"]
_REBATE_SET_KW = "設定回饋金推送"


def handle_internal_rebate_push(text: str, line_api) -> str | None:
    """內部群指令「推送回饋金」：手動觸發回饋金通知推播。
    帶「預覽/dry/測試」→ dry_run（不實際推送，只回報）"""
    t = text.strip()
    if not any(kw in t for kw in _REBATE_PUSH_KW):
        return None
    from services.rebate_push import run as rebate_push_run, build_admin_report

    dry_run = any(k in t for k in ["預覽", "dry", "測試"])
    results = rebate_push_run(line_api, dry_run=dry_run, send_admin_report=False)
    header = "🧪 預覽（未實際推送）\n" if dry_run else "📤 推送完成\n"
    return header + build_admin_report(results)


def handle_internal_set_rebate_target(text: str) -> str | None:
    """內部群指令「設定回饋金推送 <組名> <客戶名>」：備註合併組推送對象"""
    t = text.strip()
    if _REBATE_SET_KW not in t:
        return None
    m = re.match(rf'{_REBATE_SET_KW}\s+(\S+)\s+(\S+)', t)
    if not m:
        return f"❌ 格式錯誤\n請用：{_REBATE_SET_KW} <組名> <客戶名>\n例：{_REBATE_SET_KW} WEI丞 WEI"
    group_name = m.group(1)
    customer_name = m.group(2)

    from storage.customers import customer_store
    rows = customer_store.search_by_name(customer_name, real_name_only=True)
    if not rows:
        return f"❌ 找不到客戶「{customer_name}」\n請確認客戶 real_name 或 chat_label 欄位"
    if len(rows) > 1:
        names = [r.get("real_name") or r.get("display_name") or "?" for r in rows[:5]]
        return f"❌「{customer_name}」match 到 {len(rows)} 筆：{', '.join(names)}\n請用更精確名稱"
    uid = rows[0].get("line_user_id")
    if not uid:
        return f"❌ 客戶「{customer_name}」無 LINE UID（尚未互動過）"

    from storage import rebate_notify
    rebate_notify.set_target(group_name, uid, customer_name)
    return f"✅ 已設定：合併組「{group_name}」→ 推送給「{customer_name}」"


def handle_internal_rebate(text: str, state_key: str | None = None) -> str | None:
    """
    內部群回饋金查詢：
      「回饋金」「查回饋」  → 顯示當月回饋金總表
      「XXX 回饋金」       → 查詢特定客戶的回饋金
    """
    t = text.strip()
    if not any(kw in t for kw in _REBATE_KW):
        return None

    # 提取查詢客戶名
    query_name = t
    for kw in _REBATE_KW + ["查", "查詢", "多少", "？", "?", " "]:
        query_name = query_name.replace(kw, "")
    query_name = query_name.strip()

    # 單純「回饋金」不觸發，需要「回饋金資料」（總表）或「XXX回饋金」（客戶查詢）
    if not query_name and "回饋金資料" not in t:
        return None

    from services.rebate import calculate_rebates, get_approaching_customers

    result = calculate_rebates()
    groups = result.get("groups", [])
    summary = result.get("summary", {})
    month = result.get("month", "")

    if not groups:
        return f"📊 {month} 回饋金\n目前無銷貨資料"

    if query_name:
        # 特定客戶查詢
        matched = [g for g in groups
                   if query_name in g["group_name"]
                   or any(query_name in m["name"] for m in g["members"])]
        if not matched:
            return f"📊 找不到「{query_name}」的回饋金資料"

        lines = [f"📊 {month} 「{query_name}」回饋金"]
        for g in matched:
            lines.append(f"\n👤 {g['group_name']}　合計 ${g['total']:,.0f}")
            lines.append(f"   級距：{g['tier']}　回饋金：${g['rebate']:,.0f}")
            if len(g["members"]) > 1:
                for m in g["members"]:
                    rebate_str = f" → ${m['rebate']:,.0f}" if m["rebate"] > 0 else ""
                    lines.append(f"   　{m['name']}　${m['amount']:,.0f}{rebate_str}")
            # 快達標提示
            thresholds = [(30000, 17000, "3萬"), (60000, 45000, "6萬"), (100000, 75000, "10萬")]
            for target, floor, label in thresholds:
                if g["total"] < target and g["total"] >= floor:
                    lines.append(f"   ⚡ 差 ${target - g['total']:,.0f} 達 {label}")
                    break
        return "\n".join(lines)

    # 總表：列出有達標的 + 快接近的
    lines = [f"📊 {month} 回饋金總表"]
    lines.append(f"總銷售：${summary['total_sales']:,.0f}　總回饋：${summary['total_rebate']:,.0f}")

    # 有回饋金的
    with_rebate = [g for g in groups if g["rebate"] > 0]
    if with_rebate:
        lines.append(f"\n✅ 已達標（{len(with_rebate)} 組）：")
        for g in with_rebate:
            if len(g["members"]) > 1:
                lines.append(
                    f"  {g['group_name']}　${g['total']:,.0f}　"
                    f"{g['tier']}　→${g['rebate']:,.0f}"
                )
                for m in g["members"]:
                    rebate_str = f" →${m['rebate']:,.0f}" if m["rebate"] > 0 else ""
                    lines.append(f"    {m['name']}　${m['amount']:,.0f}{rebate_str}")
                    # 合併組顯示各店
                    if m.get("stores") and len(m["stores"]) > 1:
                        for s in m["stores"]:
                            lines.append(f"      {s['name']}　${s['amount']:,.0f}")
            else:
                lines.append(
                    f"  {g['group_name']}　${g['total']:,.0f}　"
                    f"{g['tier']}　→${g['rebate']:,.0f}"
                )

    # 根據日期決定顯示內容
    from datetime import datetime as _dt
    day = _dt.now().day
    if day < 15:
        # 1~14日：顯示上月達標
        from services.rebate import get_last_month_achievers
        last = get_last_month_achievers()
        if last["achievers"]:
            lines.append(f"\n✅ {last['month']} 確定達標（{len(last['achievers'])} 組）：")
            for g in last["achievers"]:
                lines.append(
                    f"  {g['group_name']}　${g['total']:,.0f}　"
                    f"{g['tier']}　→${g['rebate']:,.0f}"
                )
    else:
        # 15日起：顯示快接近達成
        approaching = get_approaching_customers()
        if approaching:
            lines.append(f"\n⚡ 快達標：")
            for a in approaching:
                lines.append(
                    f"  {a['group_name']}　${a['total']:,.0f}　"
                    f"差 ${a['gap']:,.0f} 達 {a['next_tier']}"
                )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 未處理訂單查詢
# ---------------------------------------------------------------------------

_UNFULFILLED_PATH = Path(__file__).parent.parent / "data" / "unfulfilled_orders.json"

_UNFULFILLED_KW = ["未備貨"]
_UNFULFILLED_ALL_KW = ["未備貨訂單", "未備貨資料"]


def _load_unfulfilled() -> list[dict]:
    """載入未備貨訂單資料"""
    if not _UNFULFILLED_PATH.exists():
        return []
    try:
        return _json.loads(_UNFULFILLED_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []


def _unfulfilled_needs_refresh() -> bool:
    """檔案超過 30 分鐘需要更新"""
    if not _UNFULFILLED_PATH.exists():
        return True
    import time
    age = time.time() - _UNFULFILLED_PATH.stat().st_mtime
    return age > 30 * 60


def _refresh_unfulfilled():
    """同步未處理訂單（同步執行）"""
    try:
        import asyncio
        from scripts.sync_unfulfilled import sync_unfulfilled
        asyncio.run(sync_unfulfilled())
    except Exception as e:
        print(f"[unfulfilled] 自動更新失敗: {e}")


def handle_internal_unfulfilled(text: str, state_key: str | None = None) -> str | None:
    """
    內部群未備貨訂單查詢：
      「未備貨資料」      → 列出全部未備貨訂單
      「XXX 未備貨」      → 查特定產品或客戶的未備貨訂單
    """
    t = text.strip()
    if not any(kw in t for kw in _UNFULFILLED_KW):
        return None

    # 檔案超過 30 分鐘自動更新
    if _unfulfilled_needs_refresh():
        print("[unfulfilled] 資料超過 30 分鐘，自動更新...")
        _refresh_unfulfilled()

    orders = _load_unfulfilled()
    if not orders:
        return "📋 目前沒有未備貨訂單資料"

    # 「未備貨資料」→ 按產品別分組
    if any(kw in t for kw in _UNFULFILLED_ALL_KW):
        prods: dict[str, dict] = {}  # code → {name, customers: [(cust, qty)]}
        for o in orders:
            code = o.get("code", "")
            if code not in prods:
                prods[code] = {"name": o.get("name", ""), "customers": []}
            prods[code]["customers"].append((o.get("customer", ""), o.get("qty", 0)))

        lines = [f"📋 全部未備貨訂單（{len(orders)} 筆，{len(prods)} 品項）"]
        for code, data in sorted(prods.items()):
            total = sum(q for _, q in data["customers"])
            lines.append(f"\n{code} {data['name'][:20]} 共{total:g}個")
            for cust, qty in data["customers"]:
                lines.append(f"  {cust}*{qty:g}")
        return "\n".join(lines)

    # 提取查詢關鍵字
    query = t
    for kw in _UNFULFILLED_KW + ["查", "查詢", "訂單", "？", "?", " "]:
        query = query.replace(kw, "")
    query = query.strip()

    # 沒有查詢對象 → 不觸發
    if not query:
        return None

    # 搜尋產品或客戶
    matched = [o for o in orders
               if query.upper() in o["code"].upper()
               or query in o["name"]
               or query in o["customer"]]
    if not matched:
        return f"📋 找不到「{query}」的未備貨訂單"

    # 判斷是否為產品查詢（所有結果同一產品）
    total_qty = sum(o["qty"] for o in matched)
    codes = set(o["code"] for o in matched)
    if len(codes) == 1:
        first = matched[0]
        lines = [f"{first['code']} {first['name']} 未備貨訂單({len(matched)}筆)"]
        for o in matched:
            note_str = f" {o['note']}" if o.get("note") else ""
            lines.append(f"{o['customer']} *{o['qty']:g}{note_str}")
    else:
        lines = [f"「{query}」未備貨訂單({len(matched)}筆)"]
        for o in matched:
            note_str = f" {o['note']}" if o.get("note") else ""
            lines.append(f"{o['code']} {o['name'][:18]} *{o['qty']:g}{note_str}")
    lines.append(f"合計：{total_qty:g}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 未取訂單查詢
# ---------------------------------------------------------------------------

_UNCLAIMED_PATH = Path(__file__).parent.parent / "data" / "unclaimed_orders.json"

_UNCLAIMED_KW = ["未取訂單", "未取資料", "已備貨資料", "已備貨訂單", "未取", "已備貨"]


def _load_unclaimed() -> list[dict]:
    if not _UNCLAIMED_PATH.exists():
        return []
    try:
        return _json.loads(_UNCLAIMED_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []


def _unclaimed_needs_refresh() -> bool:
    if not _UNCLAIMED_PATH.exists():
        return True
    import time
    return time.time() - _UNCLAIMED_PATH.stat().st_mtime > 30 * 60


def _refresh_unclaimed():
    try:
        import asyncio
        from scripts.sync_unfulfilled import sync_unclaimed
        asyncio.run(sync_unclaimed())
    except Exception as e:
        print(f"[unclaimed] 自動更新失敗: {e}")


_READY_PICKUP_KW = ["可通知取貨"]


def handle_internal_ready_for_pickup(text: str, state_key: str | None = None) -> str | None:
    """
    內部群「可通知取貨」：列出全部品項（含預購）都備妥的客戶，輸出明細。
    合併同名客戶（「林子翔-基隆」「林子翔-樹林」視為同一人「林子翔」）。
    """
    t = text.strip()
    if not any(kw in t for kw in _READY_PICKUP_KW):
        return None

    # 資料過期自動刷
    if _unfulfilled_needs_refresh():
        print("[ready-pickup] 未備貨資料過期，自動更新...")
        _refresh_unfulfilled()
    if _unclaimed_needs_refresh():
        print("[ready-pickup] 未取資料過期，自動更新...")
        _refresh_unclaimed()

    unfulfilled = _load_unfulfilled()
    unclaimed = _load_unclaimed()

    # 合併同名客戶 — 用 base_name（去掉 -後綴 / 括號）
    from services.rebate import _get_base_name
    pending_custs: set[str] = set()
    for o in unfulfilled:
        base = _get_base_name(o.get("customer", ""))
        if base:
            pending_custs.add(base)

    # 按 base_name 分組已備貨未取，排除任何還有未備貨的客戶（含預購）
    by_customer: dict[str, list[dict]] = {}
    for o in unclaimed:
        base = _get_base_name(o.get("customer", ""))
        if not base or base in pending_custs:
            continue
        by_customer.setdefault(base, []).append(o)

    if not by_customer:
        return "📋 目前沒有「全品項都備妥（含預購）」的客戶"

    # 先計算每位客戶的最早訂單日，再按日期升冪排（越早越前）
    def _earliest_date(orders: list[dict]) -> str:
        ds = []
        for o in orders:
            dn = (o.get("date_no") or "").split()[0].strip()
            if dn and len(dn) >= 10:
                ds.append(dn)
        return min(ds) if ds else "9999/99/99"  # 無日期排最後

    lines = [f"✅ 全品項已備妥（含預購），可通知取貨 {len(by_customer)} 位："]
    for cust, cust_orders in sorted(by_customer.items(), key=lambda x: _earliest_date(x[1])):
        total_qty = sum(o.get("qty", 0) for o in cust_orders)
        earliest = _earliest_date(cust_orders)
        date_note = f"，{earliest[5:]}起" if earliest != "9999/99/99" else ""
        lines.append(f"\n{cust}（{len(cust_orders)} 筆，共 {total_qty:g} 件{date_note}）")
        for o in cust_orders:
            lines.append(f"  {o.get('product','')[:30]} *{o.get('qty',0):g}")
    return "\n".join(lines)


def handle_internal_unclaimed(text: str, state_key: str | None = None) -> str | None:
    """
    內部群未取訂單查詢：
      「未取資料」    → 全部未取訂單摘要
      「XXX 未取」   → 查特定客戶的未取訂單
    """
    t = text.strip()
    if not any(kw in t for kw in _UNCLAIMED_KW):
        return None

    # 檔案超過 30 分鐘自動更新
    if _unclaimed_needs_refresh():
        print("[unclaimed] 資料超過 30 分鐘，自動更新...")
        _refresh_unclaimed()

    orders = _load_unclaimed()
    if not orders:
        return "📋 目前沒有未取訂單"

    # 提取查詢關鍵字
    query = t
    for kw in _UNCLAIMED_KW + ["查", "查詢", "？", "?", " "]:
        query = query.replace(kw, "")
    query = query.strip()

    if query:
        # 特定客戶查詢
        matched = [o for o in orders if query in o["customer"] or query in o["product"]]
        if not matched:
            return f"📋 找不到「{query}」的未取訂單"
        total_qty = sum(o["qty"] for o in matched)
        lines = [f"「{query}」未取訂單({len(matched)}筆，共 {total_qty:g} 件)"]
        for o in matched:
            lines.append(f"{o['product'][:20]} *{o['qty']:g}")
        return "\n".join(lines)

    # 全部未取 — 按客戶分組
    by_customer: dict[str, list[dict]] = {}
    for o in orders:
        by_customer.setdefault(o["customer"], []).append(o)

    lines = [f"📋 未取訂單（共 {len(orders)} 筆，{len(by_customer)} 位客戶）"]
    for cust, cust_orders in sorted(by_customer.items(), key=lambda x: -len(x[1])):
        total_qty = sum(o["qty"] for o in cust_orders)
        lines.append(f"\n{cust}（{len(cust_orders)} 筆，共 {total_qty:g} 件）")
        for o in cust_orders:
            lines.append(f"  {o['product'][:20]} *{o['qty']:g}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 客戶訂單查詢（未備貨 + 已備貨未取）
# ---------------------------------------------------------------------------

_ORDER_KW = ["訂單"]


def handle_internal_customer_orders(
    text: str, state_key: str | None = None,
    as_customer_reply: bool = False,
) -> str | None:
    """
    內部群客戶訂單查詢：
      「鄭家展訂單」→ 列出該客戶的未備貨 + 已備貨未取訂單
    as_customer_reply=True：給客戶看的版本，頭行會直接給結論
      （全都在已備貨未取 → 「貨都到囉，隨時可以來取」）
    """
    t = text.strip()
    if not any(kw in t for kw in _ORDER_KW):
        return None

    # 提取客戶名
    query = t
    for kw in _ORDER_KW + ["查", "查詢", "？", "?", " "]:
        query = query.replace(kw, "")
    query = query.strip()
    if not query:
        return None  # 沒有客戶名，不處理

    # 確保兩個資料檔都是最新的
    if _unfulfilled_needs_refresh():
        print("[customer-orders] 未備貨資料過期，自動更新...")
        _refresh_unfulfilled()
    if _unclaimed_needs_refresh():
        print("[customer-orders] 未取資料過期，自動更新...")
        _refresh_unclaimed()

    # 載入未備貨
    unfulfilled = _load_unfulfilled()
    uf_matched = [o for o in unfulfilled if query in o.get("customer", "")]

    # 載入已備貨未取
    unclaimed = _load_unclaimed()
    uc_matched = [o for o in unclaimed if query in o.get("customer", "")]

    if not uf_matched and not uc_matched:
        if as_customer_reply:
            return f"目前沒有查到您的訂單，請跟我們確認一下唷～"
        return f"📋 找不到「{query}」的訂單"

    lines = []
    # 客戶視角：頭行直接給結論
    if as_customer_reply:
        if uc_matched and not uf_matched:
            lines.append("✅ 貨都到囉，隨時可以來取唷～")
        elif uf_matched and not uc_matched:
            lines.append("📋 您的訂單目前還在備貨中：")
        else:
            lines.append("📋 您的訂單狀態如下（部分已到可取，部分備貨中）：")
    else:
        lines.append(f"📋「{query}」訂單")

    if uf_matched:
        total_qty = sum(o.get("qty", 0) for o in uf_matched)
        lines.append(f"\n⏳ 未備貨（{len(uf_matched)}筆，共 {total_qty:g} 件）")
        for o in uf_matched:
            name = o.get("name", "")[:20]
            code = o.get("code", "")
            qty = o.get("qty", 0)
            lines.append(f"  {code} {name} *{qty:g}")

    if uc_matched:
        total_qty = sum(o.get("qty", 0) for o in uc_matched)
        lines.append(f"\n✅ 已備貨未取（{len(uc_matched)}筆，共 {total_qty:g} 件）")
        for o in uc_matched:
            product = o.get("product", "")[:20]
            qty = o.get("qty", 0)
            lines.append(f"  {product} *{qty:g}")

    return "\n".join(lines)


# ── 廣告圖查詢 ──────────────────────────────────────────────────────────
_AD_QUERY_KW = ["廣告圖查詢", "廣告圖 查詢", "查廣告圖"]
_AD_FILENAME_CODE_RE = re.compile(
    r'([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)', re.IGNORECASE
)


def _scan_ad_image_codes() -> set[str]:
    """掃 AD_OUTPUT_DIR 取得已有廣告圖的貨號集合（大寫）"""
    from handlers.ad_maker import AD_OUTPUT_DIR
    codes: set[str] = set()
    try:
        if not AD_OUTPUT_DIR.exists():
            return codes
        for p in AD_OUTPUT_DIR.iterdir():
            if not p.is_file():
                continue
            m = _AD_FILENAME_CODE_RE.search(p.stem)
            if m:
                codes.add(m.group(1).upper())
    except Exception as e:
        print(f"[ad_query] 掃描廣告圖資料夾失敗: {e}")
    return codes


def handle_internal_ad_query(text: str, state_key: str | None = None) -> str | None:
    """
    廣告圖查詢：列出有可售庫存的貨號，依庫存分級並標記是否已有廣告圖。
    分級：>=200 / 100-199 / 50-99 / 1-49
    """
    t = text.strip()
    if not any(kw in t for kw in _AD_QUERY_KW):
        return None

    try:
        avail = _json.loads(
            (Path(__file__).parent.parent / "data" / "available.json")
            .read_text(encoding="utf-8")
        )
    except Exception as e:
        return f"❌ 讀取庫存資料失敗：{e}"

    ad_codes = _scan_ad_image_codes()

    buckets: dict[str, list[tuple[str, int, bool]]] = {
        "≥200":    [],
        "100-199": [],
        "50-99":   [],
        "1-49":    [],
    }

    for code, entry in avail.items():
        if not code or code.upper().startswith("HH"):
            continue
        qty = entry.get("available", 0) if isinstance(entry, dict) else int(entry or 0)
        try:
            qty = int(qty)
        except (TypeError, ValueError):
            continue
        if qty <= 0:
            continue

        has_ad = code.upper() in ad_codes
        row = (code, qty, has_ad)
        if qty >= 200:
            buckets["≥200"].append(row)
        elif qty >= 100:
            buckets["100-199"].append(row)
        elif qty >= 50:
            buckets["50-99"].append(row)
        else:
            buckets["1-49"].append(row)

    for k in buckets:
        buckets[k].sort(key=lambda r: -r[1])

    total = sum(len(v) for v in buckets.values())
    with_ad = sum(1 for v in buckets.values() for r in v if r[2])
    lines = [
        f"📸 廣告圖查詢（可售 {total} 品項，已有廣告圖 {with_ad}）"
    ]
    for label, rows in buckets.items():
        if not rows:
            continue
        miss = sum(1 for r in rows if not r[2])
        lines.append(f"\n▶ {label}（{len(rows)}，缺 {miss}）")
        for code, qty, has_ad in rows:
            mark = "✓" if has_ad else "✗"
            lines.append(f"  {mark} {code}  可售{qty}")

    out = "\n".join(lines)
    # LINE 文字訊息上限約 5000 字，超過就截斷
    if len(out) > 4900:
        out = out[:4850] + "\n\n⚠️ 內容過長已截斷，請改用 Admin 介面或縮小範圍"
    return out
