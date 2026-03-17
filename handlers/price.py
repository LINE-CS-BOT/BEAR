"""
價格查詢處理

客戶詢問「XX多少錢」→ 查 Ecount 品項清單的 OUT_PRICE 回答。
"""

import re

from services.ecount import ecount_client
from handlers import tone


def handle_price(user_id: str, text: str) -> str:
    """查詢產品售價並回覆"""
    product = _extract_product(text)

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
