"""診斷：嘗試用真實點擊選單導航到目標頁面"""
import asyncio, re, sys
from pathlib import Path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from scripts._chrome_helper import connect_get_page, ensure_logged_in, ERP_URL
from playwright.async_api import async_playwright


async def diag():
    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        ec_sid = await ensure_logged_in(page)

        base_url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
        await page.goto(base_url, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        print("主頁已載入，title:", await page.title())

        # ── 找並印出所有含目標文字的可點選元素 ──────────────────
        targets = {
            "客戶/供應商列表": "customer",
            "客戶清單":         "customer",
            "거래처":            "customer",
            "庫存情況":          "inventory",
        }

        for text, kind in targets.items():
            elements = await page.locator(f'text="{text}"').all()
            print(f"\n找到 {len(elements)} 個 '{text}' 元素:")
            for i, el in enumerate(elements[:3]):
                try:
                    tag = await el.evaluate("e => e.tagName")
                    visible = await el.is_visible()
                    box = await el.bounding_box()
                    print(f"  [{i}] tag={tag} visible={visible} box={box}")
                except Exception as e:
                    print(f"  [{i}] error: {e}")

        # ── 嘗試點擊「客戶/供應商列表」 ──────────────────────────
        print("\n\n--- 嘗試點擊「客戶/供應商列表」---")
        try:
            loc = page.locator('text="客戶/供應商列表"').first
            await loc.wait_for(state="visible", timeout=5000)
            await loc.click(timeout=5000)
            print("點擊成功，等待頁面...")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await page.wait_for_timeout(3000)
            print("URL:", page.url[:100])
            has_grid = await page.evaluate("() => !!document.getElementById('grid-main')")
            print("#grid-main:", has_grid)
            m = re.search(r"prgId=([^&#]+)", page.url)
            print("prgId:", m.group(1) if m else "not found")
        except Exception as e:
            print(f"點擊失敗: {e}")

        # ── 回到主頁再試點擊「庫存情況」 ─────────────────────────
        print("\n--- 嘗試點擊「庫存情況」---")
        await page.goto(base_url, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        try:
            # 找第一個 visible 的「庫存情況」連結
            locs = await page.locator('text="庫存情況"').all()
            clicked = False
            for loc in locs:
                if await loc.is_visible():
                    await loc.click(timeout=5000)
                    clicked = True
                    print("點擊成功，等待頁面...")
                    break
            if not clicked:
                print("找不到 visible 的「庫存情況」")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await page.wait_for_timeout(3000)
            print("URL:", page.url[:120])
            has_sg = await page.evaluate("() => !!document.getElementById('searchGroup')")
            print("#searchGroup:", has_sg)
            m = re.search(r"prgId=([^&#]+)", page.url)
            print("prgId:", m.group(1) if m else "not found")
        except Exception as e:
            print(f"點擊失敗: {e}")

        await browser.close()


asyncio.run(diag())
