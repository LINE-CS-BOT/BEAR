"""
自動從 Ecount 訂貨單出貨處理同步未處理訂單 → data/unfulfilled_orders.json

執行方式：
  python scripts/sync_unfulfilled.py
"""

import asyncio
import json
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
OUTPUT_PATH = ROOT / "data" / "unfulfilled_orders.json"

sys.path.insert(0, str(ROOT))
from scripts._chrome_helper import (
    launch_chrome_if_needed,
    connect_get_page,
    ensure_logged_in,
    ERP_URL,
)

_ORDER_HASH = (
    "#menuType=MENUTREE_000004"
    "&menuSeq=836HEBN5VVQ8A2M"
    "&groupSeq=MENUTREE_000030"
    "&prgId=E040230&depth=3&version=V3"
    "&menuUrl=%2FECERP%2FSVC%2FESD%2FESD030M"
    "&isFavMenu=Y"
)

_PAGE_LOADED_SEL = "#tabIng"


async def sync_unfulfilled():
    """同步未處理訂單資料"""
    from playwright.async_api import async_playwright

    launch_chrome_if_needed()
    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        if not page:
            print("[unfulfilled] ✗ 無法連接 Chrome")
            return False

        ec_sid = await ensure_logged_in(page)
        if not ec_sid:
            print("[unfulfilled] ✗ 未登入 Ecount")
            return False

        # 1. 導航到訂貨單出貨處理
        url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}{_ORDER_HASH}"
        print("[unfulfilled] 導航訂貨單出貨處理...")
        try:
            await page.goto("about:blank", timeout=5000)
            await page.goto(url, timeout=20000)
            await page.wait_for_load_state("networkidle", timeout=12000)
            await page.wait_for_timeout(3000)
        except Exception as e:
            print(f"[unfulfilled] ✗ 導航失敗: {e}")
            return False

        if not await page.query_selector(_PAGE_LOADED_SEL):
            print("[unfulfilled] ✗ 頁面未載入")
            return False

        # 2. 點擊「未處理」tab
        try:
            await page.locator("a#tabIng").click(timeout=5000)
            print("[unfulfilled] ✓ 已點擊「未處理」tab")
            await page.wait_for_timeout(5000)
        except Exception as e:
            print(f"[unfulfilled] ✗ 點擊未處理 tab 失敗: {e}")
            return False

        # 3. 滾動載入全部
        for _ in range(10):
            try:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(1000)
            except Exception:
                break

        # 4. 解析表格
        # 欄位：序號, 品項編碼, 品項名稱[規格], 日期-號碼, 客戶名稱, 餘量, 出貨數量, 交付日期, 摘要
        data = await page.evaluate(r"""() => {
            const results = [];
            const rows = document.querySelectorAll('tr');
            for (const row of rows) {
                const cells = Array.from(row.querySelectorAll('td'));
                if (cells.length < 6) continue;
                const no = cells[0]?.textContent?.trim() || '';
                if (!no || isNaN(parseInt(no))) continue;

                const code = cells[1]?.textContent?.trim() || '';
                const name = cells[2]?.textContent?.trim() || '';
                const dateNo = cells[3]?.textContent?.trim() || '';
                const customer = cells[4]?.textContent?.trim() || '';
                const qty = parseFloat((cells[5]?.textContent?.trim() || '0').replace(/,/g, '')) || 0;
                const deliveryDate = cells[7]?.textContent?.trim() || '';
                const note = cells[8]?.textContent?.trim() || '';

                if (code && customer) {
                    results.push({ code, name, date_no: dateNo, customer, qty, delivery_date: deliveryDate, note });
                }
            }
            return results;
        }""")

    if data:
        OUTPUT_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"[unfulfilled] ✓ 未處理訂單已存：{len(data)} 筆")
        print(f"[unfulfilled]   → {OUTPUT_PATH}")
        for d in data[:5]:
            print(f"  {d['code']:10s} {d['name'][:20]:20s} {d['customer']:12s} x{d['qty']}")
        return True
    else:
        print("[unfulfilled] ✗ 無法取得未處理訂單資料")
        return False


def main():
    asyncio.run(sync_unfulfilled())


if __name__ == "__main__":
    main()
