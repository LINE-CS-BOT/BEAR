"""
自動從 Ecount 銷貨單明細同步回饋金資料 → data/rebate_sales.json

執行方式：
  python scripts/sync_rebate.py              # 同步本月
  python scripts/sync_rebate.py --last-month  # 同步上個月

導航策略：
  1. 頂部導航列點「銷貨單明細」
  2. 點「回饋金總計」tab
  3. 點「本月(~今天)」或設定日期
  4. 等待結果 → 解析表格 → 存 JSON
"""

import asyncio
import io
import json
import re
import sys
import argparse
from datetime import datetime, timedelta
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
OUTPUT_PATH = ROOT / "data" / "rebate_sales.json"
LAST_MONTH_PATH = ROOT / "data" / "rebate_sales_lastmonth.json"

sys.path.insert(0, str(ROOT))
from scripts._chrome_helper import (
    launch_chrome_if_needed,
    connect_get_page,
    ensure_logged_in,
    load_web_config,
    save_web_config,
    ERP_URL,
)

# 銷貨單明細的正確 hash（含 menuSeq/groupSeq）
_SALE_DETAIL_HASH = (
    "#menuType=MENUTREE_000004"
    "&menuSeq=MENUTREE_000494"
    "&groupSeq=MENUTREE_000030"
    "&prgId=E040207&depth=4"
)

# 頁面載入成功的判斷元素（查詢(F8) 按鈕）
_PAGE_LOADED_SEL = "#header_search"

_SALE_DETAIL_MENU_TEXTS = ["銷貨單明細"]


# ---------------------------------------------------------------------------
# 導航到銷貨單明細
# ---------------------------------------------------------------------------

async def _click_sale_detail_menu(page) -> bool:
    """點頂部導航列「銷貨單明細」書籤"""
    for menu_text in _SALE_DETAIL_MENU_TEXTS:
        selectors = [
            f'a:has-text("{menu_text}")',
            f'text="{menu_text}"',
        ]
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible(timeout=3000):
                    await loc.click(timeout=5000)
                    print(f"[rebate] 點擊「{menu_text}」，等待頁面...")
                    await page.wait_for_selector(_PAGE_LOADED_SEL, timeout=15000)
                    print("[rebate] ✓ 銷貨單明細頁面已開啟")
                    return True
            except Exception:
                continue
    return False


async def _navigate_to_sale_detail(page, ec_sid: str) -> bool:
    """導向銷貨單明細頁面"""
    # 策略 1：直接用已知的正確 hash
    url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}{_SALE_DETAIL_HASH}"
    print("[rebate] 導航銷貨單明細 (prgId=E040207)...")
    try:
        await page.goto("about:blank", timeout=5000)
        await page.goto(url, timeout=20000)
        await page.wait_for_load_state("networkidle", timeout=12000)
        await page.wait_for_timeout(3000)
        if await page.query_selector(_PAGE_LOADED_SEL):
            print("[rebate] ✓ 銷貨單明細頁面已開啟")
            return True
        print("[rebate] 直接導航未偵測到頁面，嘗試其他策略...")
    except Exception as e:
        print(f"[rebate] 直接導航失敗: {e}")

    # 策略 2：回首頁點書籤選單
    print("[rebate] 回 ERP 首頁再點「銷貨單明細」書籤...")
    try:
        url_home = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
        await page.goto(url_home, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"[rebate] ERP 首頁導向失敗: {e}")

    if await _click_sale_detail_menu(page):
        return True

    print("[rebate] ✗ 無法導航到銷貨單明細")
    return False


# ---------------------------------------------------------------------------
# 點擊「回饋金總計」tab 並設定查詢條件
# ---------------------------------------------------------------------------

async def _click_rebate_tab(page) -> bool:
    """點擊「回饋金總計」tab（li.preset id="1"）"""
    try:
        for sel in [
            'li.preset:has-text("回饋金總計")',
            'li:has-text("回饋金總計")',
            'a:has-text("回饋金總計")',
        ]:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible(timeout=3000):
                    await loc.click()
                    print("[rebate] ✓ 已點擊「回饋金總計」tab")
                    await page.wait_for_timeout(2000)
                    return True
            except Exception:
                continue
        print("[rebate] ✗ 找不到「回饋金總計」tab")
        return False
    except Exception as e:
        print(f"[rebate] 點擊 tab 失敗: {e}")
        return False


async def _click_this_month(page) -> bool:
    """點擊「本月(~今天)」快速日期選擇"""
    try:
        loc = page.locator('button:has-text("本月(~今天)")').first
        if await loc.is_visible(timeout=3000):
            await loc.click()
            print("[rebate] ✓ 已點擊「本月(~今天)」")
            await page.wait_for_timeout(1000)
            return True
    except Exception:
        pass
    print("[rebate] ✗ 找不到「本月(~今天)」按鈕")
    return False


async def _click_last_month(page) -> bool:
    """點擊「上個月」快速日期選擇"""
    try:
        # 精確匹配「上個月」，排除「上個月+本月」
        locs = await page.locator('button:has-text("上個月")').all()
        for loc in locs:
            text = (await loc.text_content() or "").strip()
            if text == "上個月" and await loc.is_visible():
                await loc.click()
                print("[rebate] ✓ 已點擊「上個月」")
                await page.wait_for_timeout(1000)
                return True
    except Exception:
        pass
    print("[rebate] ✗ 找不到「上個月」按鈕")
    return False


async def _click_query(page) -> bool:
    """點擊查詢(F8)"""
    try:
        loc = page.locator("#header_search")
        if await loc.is_visible(timeout=3000):
            await loc.click()
            print("[rebate] ✓ 點擊查詢(F8)")
            return True
    except Exception:
        pass
    # Fallback: F8 鍵
    try:
        await page.keyboard.press("F8")
        print("[rebate] ✓ 按 F8 查詢")
        return True
    except Exception as e:
        print(f"[rebate] 查詢失敗: {e}")
        return False


# ---------------------------------------------------------------------------
# 解析結果表格
# ---------------------------------------------------------------------------

async def _parse_results(page) -> list[dict]:
    """解析結果表格，取出客戶名稱和合計金額"""
    print("[rebate] 等待結果...")
    await page.wait_for_timeout(5000)  # 等資料載入

    # 嘗試滾動載入全部
    for _ in range(10):
        try:
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await page.wait_for_timeout(1000)
        except Exception:
            break

    # 解析表格
    _PARSE_JS = r"""
        () => {
            const results = [];
            // 找結果表格（grid-main 或 class 含 grid 的 table）
            const tables = document.querySelectorAll('table');
            for (const table of tables) {
                const rows = table.querySelectorAll('tr');
                for (const row of rows) {
                    const cells = Array.from(row.querySelectorAll('td'));
                    if (cells.length < 4) continue;

                    // 找「改盒」在第一欄
                    const col0 = cells[0]?.textContent?.trim() || '';
                    if (col0 !== '改盒') continue;

                    // 第二欄：客戶名稱
                    const customer = cells[1]?.textContent?.trim() || '';
                    if (!customer) continue;

                    // 最後一欄（合計）的數字
                    const lastCell = cells[cells.length - 1]?.textContent?.trim() || '0';
                    const amount = parseFloat(lastCell.replace(/,/g, '')) || 0;

                    if (amount !== 0) {
                        results.push({ customer, amount });
                    }
                }
            }
            return results;
        }
    """

    try:
        data = await page.evaluate(_PARSE_JS)
        print(f"[rebate] 解析到 {len(data)} 筆改盒客戶資料")
        return data
    except Exception as e:
        print(f"[rebate] 解析表格失敗: {e}")
        return []


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

async def sync_rebate(last_month: bool = False):
    """同步回饋金資料"""
    from playwright.async_api import async_playwright

    launch_chrome_if_needed()
    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        if not page:
            print("[rebate] ✗ 無法連接 Chrome")
            return False

        ec_sid = await ensure_logged_in(page)
        if not ec_sid:
            print("[rebate] ✗ 未登入 Ecount")
            return False

        # 1. 導航到銷貨單明細
        if not await _navigate_to_sale_detail(page, ec_sid):
            return False

        # 2. 點擊「回饋金總計」tab
        await page.wait_for_timeout(2000)
        if not await _click_rebate_tab(page):
            print("[rebate] 嘗試不點 tab 直接查詢...")

        # 3. 設定日期範圍
        await page.wait_for_timeout(1000)
        if last_month:
            await _click_last_month(page)
        else:
            await _click_this_month(page)

        # 4. 查詢
        await page.wait_for_timeout(1000)
        await _click_query(page)

        # 5. 解析結果
        await page.wait_for_timeout(3000)
        data = await _parse_results(page)

        if not data:
            print("[rebate] ⚠ 沒有取得資料，嘗試用 Excel 下載...")
            data = await _try_excel_download(page)

    if data:
        month_label = "上月" if last_month else "本月"
        target_path = LAST_MONTH_PATH if last_month else OUTPUT_PATH
        target_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        total = sum(d["amount"] for d in data)
        print(f"[rebate] ✓ {month_label}回饋金資料已存：{len(data)} 筆客戶，合計 ${total:,.0f}")
        print(f"[rebate]   → {target_path}")

        for d in sorted(data, key=lambda x: -x["amount"])[:5]:
            print(f"  {d['customer']:20s} ${d['amount']:>10,.0f}")
        return True
    else:
        print("[rebate] ✗ 無法取得回饋金資料")
        return False


async def _try_excel_download(page) -> list[dict]:
    """嘗試點 Excel(畫面) 按鈕下載資料"""
    try:
        for sel in [
            'text="Excel(畫面)"',
            ':text("Excel")',
            'button:has-text("Excel")',
        ]:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible(timeout=3000):
                    # 等待下載
                    async with page.expect_download(timeout=30000) as dl:
                        await loc.click()
                    download = await dl.value
                    tmp_path = ROOT / "data" / "_tmp_rebate.xlsx"
                    await download.save_as(str(tmp_path))
                    print(f"[rebate] 已下載 Excel → {tmp_path.name}")
                    return _parse_excel(tmp_path)
            except Exception:
                continue
    except Exception as e:
        print(f"[rebate] Excel 下載失敗: {e}")
    return []


def _parse_excel(path: Path) -> list[dict]:
    """解析下載的 Excel"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active

        results = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 4:
                continue
            col0 = str(row[0] or "").strip()
            if col0 != "改盒":
                continue
            customer = str(row[1] or "").strip()
            # 合計在最後一欄
            amount = 0
            for cell in reversed(row):
                try:
                    val = float(str(cell).replace(",", ""))
                    if val != 0:
                        amount = val
                        break
                except (ValueError, TypeError):
                    continue
            if customer and amount > 0:
                results.append({"customer": customer, "amount": amount})

        wb.close()
        print(f"[rebate] Excel 解析完成：{len(results)} 筆")
        return results
    except Exception as e:
        print(f"[rebate] Excel 解析失敗: {e}")
        return []


def main():
    parser = argparse.ArgumentParser(description="同步 Ecount 回饋金資料")
    parser.add_argument("--last-month", action="store_true", help="同步上個月")
    args = parser.parse_args()

    asyncio.run(sync_rebate(last_month=args.last_month))


if __name__ == "__main__":
    main()
