"""診斷：ERP 主頁載入後的 DOM 狀態"""
import asyncio, sys
from pathlib import Path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from scripts._chrome_helper import connect_get_page, ensure_logged_in, ERP_URL
from playwright.async_api import async_playwright


async def diag():
    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        ec_sid = await ensure_logged_in(page)
        print("sid:", ec_sid[:15] if ec_sid else None)

        # Load ERP base page
        base_url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
        print("Going to:", base_url[:80])
        await page.goto(base_url, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(3000)

        print("URL after load:", page.url[:100])
        title = await page.title()
        print("Title:", title)

        # Check major elements
        for sel in ["#grid-main", "#searchGroup", "#frameApps",
                    "#contents", ".menuList", "#leftMenu", "#menu", "#gnb"]:
            exists = await page.evaluate(f"() => !!document.querySelector('{sel}')")
            print(f"  {sel}: {exists}")

        # Hash change test
        print("\nTrying location.hash change to E010101...")
        await page.evaluate("location.hash = '#menuType=MENUTREE_000001&prgId=E010101&depth=2'")
        await page.wait_for_timeout(5000)
        has_grid = await page.evaluate("() => !!document.getElementById('grid-main')")
        print(f"  #grid-main after hashchange: {has_grid}")
        cur_url = page.url
        print(f"  URL: {cur_url[:100]}")

        # Print page body text snippet
        body_text = await page.evaluate(
            "() => document.body.innerText.replace(/\\s+/g,' ').slice(0, 300)"
        )
        print("\nPage body text (300 chars):")
        print(body_text)

        await browser.close()


asyncio.run(diag())
