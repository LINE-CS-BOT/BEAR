"""
價格查詢處理

客戶詢問「XX多少錢」→ 查 Ecount 品項清單的 OUT_PRICE 回答。
"""

import re
import sqlite3
from pathlib import Path

from services.ecount import ecount_client
from handlers import tone


_PROD_CODE_RE = re.compile(r'[A-Za-z]{1,3}-?\d{3,6}')
_CHAT_DB = Path(__file__).parent.parent / "data" / "chat_history.db"


def _find_recent_product_names(user_id: str, max_rows: int = 4) -> list[str]:
    """從最近幾則 bot 訊息抓產品名稱/貨號 (剛推過的到貨通知、產品推薦等)。
    優先抓貨號；若沒有，抓「(xxx)xxx」格式的品名。"""
    try:
        with sqlite3.connect(_CHAT_DB) as conn:
            rows = conn.execute(
                "SELECT text FROM chat_history WHERE user_id=? AND role='bot' "
                "ORDER BY id DESC LIMIT ?",
                (user_id, max_rows),
            ).fetchall()
    except Exception:
        return []

    found: list[str] = []
    seen: set[str] = set()
    for (txt,) in rows:
        # 先抓貨號
        for m in _PROD_CODE_RE.findall(txt or ""):
            key = m.upper()
            if key not in seen:
                seen.add(key)
                found.append(key)
        # 再抓「•品名 × N」這種到貨通知行
        for m in re.finditer(r'[•・･]\s*([^\n×xX]{3,40})(?:\s*[×xX]\s*\d+)?', txt or ""):
            name = m.group(1).strip()
            if name and name not in seen:
                seen.add(name)
                found.append(name)
    return found


def handle_price(user_id: str, text: str) -> str:
    """查詢產品售價並回覆"""
    product = _extract_product(text)

    # 沒指定產品 → 從最近對話找剛推過的（到貨通知/產品推薦），1 個就自動用
    if not product:
        recent = _find_recent_product_names(user_id)
        if len(recent) == 1:
            product = recent[0]
            print(f"[price] 從 chat_history 取用最近產品 → {product}", flush=True)
        elif len(recent) > 1:
            listed = "\n".join(f"  • {n}" for n in recent[:5])
            return (
                f"是要問剛剛提到的這幾款嗎？\n{listed}\n"
                f"請跟我說貨號或名稱，我幫{tone.boss()}查一下 👍"
            )

    if not product:
        return f"請問{tone.boss()}要查哪款的價格呢？（輸入產品編號或名稱）"

    result = ecount_client.get_price(product)

    if result is None:
        return tone.product_not_found(product)

    name  = result["name"] or product
    price = result["price"]
    unit  = result["unit"] or "個"

    if price <= 0:
        # 有品項但沒設定售價 → 轉人工
        return (
            f"「{name}」的價格需要確認一下{tone.suffix_light()} "
            f"稍後幫{tone.boss()}回覆哦"
        )

    return tone.price_reply(name, price, unit)


def _extract_product(text: str) -> str:
    """從訊息中提取產品名稱/編號（去除價格詢問詞）"""
    t = re.sub(
        r"^(?:請問|想問|問一下|查一下|你們|你們的)\s*", "", text
    )
    t = re.sub(
        r"\s*(?:多少錢|多少钱|幾錢|幾塊|價格|単價|售價|多少|報價|價位|價錢"
        r"|一個多少|一箱多少|一盒多少|一個|一箱|一盒)\s*$",
        "", t,
    )
    t = re.sub(r"\s*嗎\s*$", "", t)
    t = t.strip()
    return t if t and t != text else ""
