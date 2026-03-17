"""
從 Ecount 庫存情況報表 Excel 同步未出貨數量到 data/unfulfilled.json

使用方式：
1. 登入 Ecount ERP → 進銷存 → 報表 → 庫存情況
2. 倉庫選 101，按查詢(F8)
3. 點「Excel」下載，把 xlsx 放到 data/ 或直接指定路徑
4. 執行：python scripts/sync_unfulfilled.py [路徑/到/檔案.xlsx]
"""

import json
import re
import sys
from pathlib import Path

# 預設路徑
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_PATH = DATA_DIR / "unfulfilled.json"


def parse_excel(xlsx_path: Path) -> dict[str, int]:
    try:
        import openpyxl
    except ImportError:
        print("請先安裝 openpyxl: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # 欄位：品項編碼(0) 品項名稱(1) 出庫單價(2) 安全庫存(3) 最少採購單位(4)
    #        未進貨(5) 未出貨(6) 庫存數量(7) 可售庫存(8)
    unfulfilled: dict[str, int] = {}
    for row in rows[2:]:  # 跳過前兩列標題
        code = row[0]
        if not code or not isinstance(code, str):
            continue
        code = code.strip()
        # 只處理合法貨號（英數字 + 橫線）
        if not re.match(r'^[A-Za-z0-9\-]+$', code):
            continue
        qty = row[6]  # 未出貨欄
        if qty and float(qty) > 0:
            unfulfilled[code.upper()] = int(float(qty))

    return unfulfilled


def main():
    if len(sys.argv) > 1:
        xlsx_path = Path(sys.argv[1])
    else:
        # 自動找 data/ 目錄裡最新的 xlsx
        xlsx_files = sorted(DATA_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not xlsx_files:
            print(f"找不到 xlsx 檔案，請放到 {DATA_DIR} 或指定路徑")
            sys.exit(1)
        xlsx_path = xlsx_files[0]
        print(f"使用最新檔案：{xlsx_path.name}")

    if not xlsx_path.exists():
        print(f"檔案不存在：{xlsx_path}")
        sys.exit(1)

    unfulfilled = parse_excel(xlsx_path)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(json.dumps(unfulfilled, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] Synced {len(unfulfilled)} items -> {OUTPUT_PATH}")
    print()
    # 顯示未出貨數量較大的品項
    top = sorted(unfulfilled.items(), key=lambda x: x[1], reverse=True)[:10]
    print("Top unfulfilled items:")
    for code, qty in top:
        print(f"  {code:<15} unfulfilled={qty}")


if __name__ == "__main__":
    main()
