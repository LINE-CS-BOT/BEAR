"""
從 Ecount 銷貨單明細同步完整出貨紀錄 → data/sales_detail.db (SQLite)

用途：銷售分析、滯銷偵測、客戶分析、補貨預測、價位帶分析

執行方式：
  python -m scripts.sync_sales_detail              # 同步本月
  python -m scripts.sync_sales_detail --months 3   # 同步最近 3 個月
"""

import asyncio
import json
import re
import sys
import sqlite3
import argparse
from datetime import datetime, timedelta
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
DB_PATH = ROOT / "data" / "sales_detail.db"

sys.path.insert(0, str(ROOT))
from scripts._chrome_helper import (
    launch_chrome_if_needed,
    connect_get_page,
    ensure_logged_in,
    load_web_config,
    save_web_config,
    ERP_URL,
)

# 銷貨單明細 hash
_SALE_DETAIL_HASH = (
    "#menuType=MENUTREE_000004"
    "&menuSeq=MENUTREE_000494"
    "&groupSeq=MENUTREE_000030"
    "&prgId=E040207&depth=4"
)

_SALE_DETAIL_MENU_TEXTS = ["銷貨單明細"]


# ---------------------------------------------------------------------------
# DB
# ---------------------------------------------------------------------------

def _init_db():
    """建立 sales_detail 表"""
    with sqlite3.connect(str(DB_PATH)) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sales_detail (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                slip_no TEXT,
                customer TEXT NOT NULL,
                prod_cd TEXT NOT NULL,
                prod_name TEXT,
                qty INTEGER DEFAULT 0,
                unit_price REAL DEFAULT 0,
                amount REAL DEFAULT 0,
                warehouse TEXT,
                synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(date, slip_no, prod_cd, customer)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sd_date ON sales_detail(date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sd_prod ON sales_detail(prod_cd)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sd_cust ON sales_detail(customer)")


def _upsert_rows(rows: list[dict]) -> int:
    """批次寫入（重複的跳過），回傳新增筆數"""
    inserted = 0
    with sqlite3.connect(str(DB_PATH)) as conn:
        for r in rows:
            try:
                conn.execute("""
                    INSERT OR IGNORE INTO sales_detail
                    (date, slip_no, customer, prod_cd, prod_name, qty, unit_price, amount, warehouse)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    r.get("date", ""),
                    r.get("slip_no", ""),
                    r.get("customer", ""),
                    r.get("prod_cd", ""),
                    r.get("prod_name", ""),
                    r.get("qty", 0),
                    r.get("unit_price", 0),
                    r.get("amount", 0),
                    r.get("warehouse", ""),
                ))
                inserted += conn.total_changes
            except Exception:
                pass
        conn.commit()
    return inserted


# ---------------------------------------------------------------------------
# 導航
# ---------------------------------------------------------------------------

async def _navigate_to_sale_detail(page, ec_sid: str) -> bool:
    """導航到銷貨單明細頁"""
    # 策略 1：用 saved hash
    cfg = load_web_config()
    saved_hash = cfg.get("sale_detail_hash", _SALE_DETAIL_HASH)
    try:
        url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}{saved_hash}"
        print(f"  導航到銷貨單明細...")
        await page.goto("about:blank", timeout=5000)
        await page.goto(url, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(3000)
        if await page.query_selector("#header_search"):
            print("  ✓ 銷貨單明細頁面載入成功")
            return True
    except Exception as e:
        print(f"  [warn] 導航失敗: {e}")

    # 策略 2：點選單
    for menu_text in _SALE_DETAIL_MENU_TEXTS:
        try:
            loc = page.locator(f'text="{menu_text}"').first
            if await loc.is_visible(timeout=3000):
                await loc.click()
                await page.wait_for_load_state("networkidle", timeout=15000)
                await page.wait_for_timeout(3000)
                if await page.query_selector("#header_search"):
                    # 儲存成功的 hash
                    new_hash = re.search(r"(#.+)", page.url)
                    if new_hash:
                        cfg["sale_detail_hash"] = new_hash.group(1)
                        save_web_config(cfg)
                    print(f"  ✓ 透過選單「{menu_text}」載入成功")
                    return True
        except Exception:
            pass

    print("  ✗ 無法導航到銷貨單明細")
    return False


async def _set_date_range(page, start_date: str, end_date: str):
    """設定查詢日期範圍 (格式: YYYY/MM/DD)"""
    try:
        # 找到日期輸入欄
        start_inputs = await page.query_selector_all('input[type="text"]')
        date_inputs = []
        for inp in start_inputs:
            val = await inp.get_attribute("value") or ""
            placeholder = await inp.get_attribute("placeholder") or ""
            if re.match(r"\d{4}/\d{2}/\d{2}", val) or "日期" in placeholder:
                date_inputs.append(inp)

        if len(date_inputs) >= 2:
            # 清除並填入開始日期
            await date_inputs[0].click(click_count=3)
            await date_inputs[0].fill(start_date)
            await page.wait_for_timeout(500)
            # 清除並填入結束日期
            await date_inputs[1].click(click_count=3)
            await date_inputs[1].fill(end_date)
            await page.wait_for_timeout(500)
            print(f"  日期範圍：{start_date} ~ {end_date}")
        else:
            print(f"  [warn] 找到 {len(date_inputs)} 個日期欄位，嘗試 JS 設定...")
            await page.evaluate(f"""
                () => {{
                    const inputs = document.querySelectorAll('input[type="text"]');
                    for (const inp of inputs) {{
                        if (/\\d{{4}}\\/\\d{{2}}\\/\\d{{2}}/.test(inp.value)) {{
                            if (!window._dateSet1) {{
                                inp.value = '{start_date}';
                                inp.dispatchEvent(new Event('change'));
                                window._dateSet1 = true;
                            }} else {{
                                inp.value = '{end_date}';
                                inp.dispatchEvent(new Event('change'));
                                break;
                            }}
                        }}
                    }}
                }}
            """)
    except Exception as e:
        print(f"  [warn] 設定日期失敗: {e}")


async def _click_query(page):
    """點查詢按鈕"""
    for sel in ["#header_search", 'button:has-text("查詢")', 'text="查詢(F8)"']:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=2000):
                await loc.click()
                print("  點擊查詢...")
                await page.wait_for_timeout(5000)
                return
        except Exception:
            pass
    # fallback: F8
    try:
        await page.keyboard.press("F8")
        print("  按 F8 查詢...")
        await page.wait_for_timeout(5000)
    except Exception:
        pass


async def _download_excel(page) -> Path | None:
    """下載 Excel"""
    tmp_path = ROOT / "data" / "_tmp_sales_detail.xlsx"
    for sel in [
        '[data-cid*="Excel"]',
        'text="Excel(畫面)"',
        ':text("Excel")',
        'button:has-text("Excel")',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=3000):
                async with page.expect_download(timeout=30000) as dl:
                    await loc.click()
                download = await dl.value
                await download.save_as(str(tmp_path))
                print(f"  ✓ 已下載 Excel → {tmp_path.name}")
                return tmp_path
        except Exception:
            continue
    print("  ✗ Excel 下載失敗")
    return None


def _parse_excel(path: Path) -> list[dict]:
    """解析銷貨單明細 Excel"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        ws = wb.active

        rows_data = list(ws.iter_rows(values_only=True))
        if not rows_data:
            return []

        # 找表頭（含「品項編碼」或「品名」的那行）
        header_idx = -1
        headers = []
        for i, row in enumerate(rows_data):
            row_str = " ".join(str(c or "") for c in row)
            if "品項編碼" in row_str or "品名" in row_str:
                headers = [str(c or "").strip() for c in row]
                header_idx = i
                break

        if header_idx < 0:
            # 嘗試第二行當表頭
            if len(rows_data) > 1:
                headers = [str(c or "").strip() for c in rows_data[1]]
                header_idx = 1
            else:
                print(f"  [warn] 找不到表頭")
                return []

        print(f"  Excel 表頭 (第{header_idx+1}行): {headers[:10]}")

        # 欄位對應
        col_map = {}
        for ci, h in enumerate(headers):
            h_lower = h.lower().replace(" ", "")
            if "日期" in h and "date" not in col_map:
                col_map["date"] = ci
            elif "單據號碼" in h or "單號" in h:
                col_map["slip_no"] = ci
            elif "客戶" in h and "名稱" in h:
                col_map["customer"] = ci
            elif "品項編碼" in h or "品項代碼" in h:
                col_map["prod_cd"] = ci
            elif "品名" in h or "品項名稱" in h:
                col_map["prod_name"] = ci
            elif "數量" in h and "qty" not in col_map:
                col_map["qty"] = ci
            elif "單價" in h:
                col_map["unit_price"] = ci
            elif "金額" in h or "合計" in h:
                col_map["amount"] = ci
            elif "倉庫" in h:
                col_map["warehouse"] = ci

        print(f"  欄位對應: {col_map}")

        results = []
        for row in rows_data[header_idx + 1:]:
            if not row or all(c is None for c in row):
                continue

            def _get(key):
                idx = col_map.get(key)
                if idx is not None and idx < len(row):
                    return row[idx]
                return None

            prod_cd = str(_get("prod_cd") or "").strip()
            if not prod_cd or not re.match(r"[A-Za-z]", prod_cd):
                continue  # 跳過非產品行（合計行等）

            date_val = _get("date")
            if isinstance(date_val, datetime):
                date_str = date_val.strftime("%Y-%m-%d")
                slip_from_date = ""
            else:
                raw = str(date_val or "").strip()
                # 「日期-號碼」格式：「2026/02/01 -1」→ 日期 + 單號
                dm = re.match(r"(\d{4}/\d{2}/\d{2})\s*(-?\d+)?", raw)
                if dm:
                    date_str = dm.group(1).replace("/", "-")
                    slip_from_date = dm.group(2) or ""
                else:
                    date_str = raw.replace("/", "-")
                    slip_from_date = ""

            # 單號：優先用 slip_no 欄，沒有則用日期欄裡的
            slip_no = str(_get("slip_no") or "").strip() or slip_from_date

            qty = _get("qty")
            try:
                qty = int(float(qty)) if qty else 0
            except (ValueError, TypeError):
                qty = 0

            unit_price = _get("unit_price")
            try:
                unit_price = float(str(unit_price).replace(",", "")) if unit_price else 0
            except (ValueError, TypeError):
                unit_price = 0

            amount = _get("amount")
            try:
                amount = float(str(amount).replace(",", "")) if amount else 0
            except (ValueError, TypeError):
                amount = 0

            results.append({
                "date": date_str,
                "slip_no": slip_no,
                "customer": str(_get("customer") or "").strip(),
                "prod_cd": prod_cd.upper(),
                "prod_name": str(_get("prod_name") or "").strip(),
                "qty": qty,
                "unit_price": unit_price,
                "amount": amount,
                "warehouse": str(_get("warehouse") or "").strip(),
            })

        print(f"  Excel 解析完成：{len(results)} 筆")
        return results

    except Exception as e:
        print(f"  Excel 解析失敗: {e}")
        import traceback; traceback.print_exc()
        return []


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

async def sync(months: int = 1):
    """同步銷貨明細"""
    from playwright.async_api import async_playwright

    _init_db()
    launch_chrome_if_needed()

    now = datetime.now()
    end_date = now.strftime("%Y/%m/%d")
    start_dt = now.replace(day=1) - timedelta(days=(months - 1) * 30)
    start_dt = start_dt.replace(day=1)
    start_date = start_dt.strftime("%Y/%m/%d")

    print(f"[sales] 同步銷貨明細：{start_date} ~ {end_date}")

    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        if not page:
            print("[sales] ✗ 無法連接 Chrome")
            return False

        ec_sid = await ensure_logged_in(page)
        if not ec_sid:
            print("[sales] ✗ 未登入 Ecount")
            return False

        # 1. 導航
        if not await _navigate_to_sale_detail(page, ec_sid):
            return False

        # 2. 設定日期
        await _set_date_range(page, start_date, end_date)

        # 3. 查詢
        await _click_query(page)
        await page.wait_for_timeout(3000)

        # 4. 下載 Excel
        xlsx = await _download_excel(page)
        if not xlsx:
            print("[sales] ✗ 無法下載 Excel")
            return False

    # 5. 解析 + 寫入 DB
    rows = _parse_excel(xlsx)
    if not rows:
        print("[sales] ✗ 沒有解析到資料")
        return False

    inserted = _upsert_rows(rows)
    print(f"[sales] ✓ 共 {len(rows)} 筆，新增 {inserted} 筆")

    # 統計
    customers = set(r["customer"] for r in rows)
    products = set(r["prod_cd"] for r in rows)
    total_amount = sum(r["amount"] for r in rows)
    print(f"[sales]   客戶 {len(customers)} 位，品項 {len(products)} 個，總額 ${total_amount:,.0f}")

    # 清理暫存
    try:
        xlsx.unlink(missing_ok=True)
    except Exception:
        pass

    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--months", type=int, default=1, help="同步幾個月（預設 1）")
    args = parser.parse_args()

    ok = asyncio.run(sync(months=args.months))
    if not ok:
        print("❌ 同步失敗")
        sys.exit(1)


if __name__ == "__main__":
    main()
