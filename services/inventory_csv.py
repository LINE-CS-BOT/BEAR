"""
CSV 庫存查詢（Ecount API 未接通前的替代方案）

員工直接編輯 data/inventory.csv 更新庫存。
Bot 每次查詢都讀最新的檔案，不需重啟。

CSV 格式：
    產品編號,產品名稱,庫存數量,備註
    P0101,蝸牛面膜,50,
    P0102,保濕乳液,0,預計3/10到貨
"""

import csv
from pathlib import Path

INVENTORY_FILE = Path("data/inventory.csv")


def lookup(product_input: str) -> dict | None:
    """
    查詢庫存。用產品編號或產品名稱（模糊）搜尋。

    Returns:
        dict  — {"code": str, "name": str, "qty": int, "note": str}
        None  — 查無此產品
    """
    if not INVENTORY_FILE.exists():
        return None

    keyword = product_input.strip().upper()
    best = None

    with open(INVENTORY_FILE, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = row.get("產品編號", "").strip().upper()
            name = row.get("產品名稱", "").strip()
            # 完全符合編號優先
            if code == keyword:
                best = row
                break
            # 名稱包含關鍵字次之
            if keyword in name.upper() or keyword in code:
                best = row

    if not best:
        return None

    return {
        "code": best.get("產品編號", "").strip(),
        "name": best.get("產品名稱", "").strip(),
        "qty":  int(best.get("庫存數量", 0) or 0),
        "note": best.get("備註", "").strip(),
    }
