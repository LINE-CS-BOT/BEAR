"""
從 Ecount 品項基本資料頁抓取產品圖片，存到 H:\\ 產品圖資料夾

用法：
  python scripts/fetch_ecount_product_images.py Z3300 Z3331 T1202
  python scripts/fetch_ecount_product_images.py --all-cheap   # 所有 150 元以下

流程：
  1. 連接 Chrome CDP（複用 auto_sync 的 helper）
  2. 導航到品項基本資料頁（BA000001 / BA017 / BA0001）
  3. 逐一搜尋貨號 → 開啟詳情 → 抓 <img> → 存檔
"""

import asyncio
import io
import re
import sys
import json
import httpx
from pathlib import Path
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from scripts._chrome_helper import (
    launch_chrome_if_needed,
    connect_get_page,
    ensure_logged_in,
    load_web_config,
    save_web_config,
    ERP_URL,
)
from handlers.ad_maker import AD_IMAGE_DIR

# 品項基本資料候選 prgId
_BASIC_PRGIDS = [
    "BA000001", "BA0001", "BA001",
    "E000401",  "E000400",
    "INV_BASIC", "PROD001",
]

# 150 元以下有 specs 的產品清單
CHEAP_CODES = [
    "Z3348","T1188","Z3558","Z3495","Z3300","Z1995","Z3340","Z1941-1",
    "R0147","Z3331","Z3325","Z3363","Z1822","T1202","Z3207","R0143",
    "Z1950","Z3212","Z3234","P0142","Z3329","Q0449","T1134","R0152",
]


# ---------------------------------------------------------------------------
# 導航到品項基本資料
# ---------------------------------------------------------------------------

async def _navigate_to_product_basic(page, ec_sid: str) -> bool:
    """嘗試導航到「品項基本資料」頁面"""

    cfg = load_web_config()
    saved_hash = cfg.get("product_basic_hash")

    # 策略 1：使用已儲存 hash
    if saved_hash:
        url = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}{saved_hash}"
        try:
            await page.goto("about:blank", timeout=5000)
            await page.goto(url, timeout=25000)
            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(2000)
            if await _is_product_basic_page(page):
                print("[img] ✓ 已儲存 hash 導航成功")
                return True
        except Exception as e:
            print(f"[img] 已儲存 hash 失敗: {e}")

    # 策略 2：逐一嘗試 prgId
    print("[img] 嘗試候選 prgId...")
    for prgid in _BASIC_PRGIDS:
        url = (
            f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
            f"#menuType=MENUTREE_000001&prgId={prgid}&depth=3"
        )
        try:
            await page.goto("about:blank", timeout=5000)
            await page.goto(url, timeout=25000)
            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(3000)
            if await _is_product_basic_page(page):
                h = f"#menuType=MENUTREE_000001&prgId={prgid}&depth=3"
                save_web_config({"product_basic_hash": h})
                print(f"[img] ✓ prgId={prgid} 成功，已儲存 hash")
                return True
            print(f"[img]   prgId={prgid} → 不是品項頁")
        except Exception as e:
            print(f"[img]   prgId={prgid} 失敗: {e}")

    # 策略 3：嘗試點選選單
    print("[img] 嘗試選單點擊...")
    menu_texts = ["品項基本資料", "제품기본사항", "Basic Products", "품목기본정보"]
    try:
        url_home = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
        await page.goto(url_home, timeout=25000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(2000)
        for text in menu_texts:
            locs = await page.locator(f':text("{text}")').all()
            for loc in locs:
                try:
                    if await loc.is_visible():
                        await loc.click(timeout=3000)
                        await page.wait_for_timeout(3000)
                        if await _is_product_basic_page(page):
                            h = "#" + page.url.split("#", 1)[1] if "#" in page.url else ""
                            if h:
                                save_web_config({"product_basic_hash": h})
                            print(f"[img] ✓ 選單點擊 '{text}' 成功")
                            return True
                except Exception:
                    pass
    except Exception as e:
        print(f"[img] 選單點擊失敗: {e}")

    print("[img] ✗ 無法導航到品項基本資料頁")
    print("[img]   請手動在 Ecount 開啟「品項基本資料」頁面，再重跑腳本")
    return False


async def _is_product_basic_page(page) -> bool:
    """判斷是否在品項基本資料頁（有搜尋表單 + 貨號輸入框）"""
    for sel in [
        'input[id*="PROD_CD"]', 'input[name*="PROD_CD"]',
        'input[placeholder*="貨號"]', 'input[placeholder*="품목"]',
        '#searchGroup input', '.search-area input',
    ]:
        try:
            if await page.query_selector(sel):
                return True
        except Exception:
            pass
    return False


# ---------------------------------------------------------------------------
# 搜尋並取得圖片
# ---------------------------------------------------------------------------

async def _get_product_image(page, code: str, out_dir: Path) -> bool:
    """在品項基本資料頁搜尋 code，進詳情頁，抓取並儲存圖片"""

    print(f"[img] 搜尋 {code}...")

    # 找貨號輸入框並填入
    prod_cd_sel = None
    for sel in [
        'input[id*="PROD_CD"]', 'input[name*="PROD_CD"]',
        'input[placeholder*="貨號"]', '#PROD_CD', 'input[id="PROD_CD"]',
    ]:
        try:
            el = await page.query_selector(sel)
            if el:
                prod_cd_sel = sel
                break
        except Exception:
            pass

    if not prod_cd_sel:
        # 嘗試截圖看看頁面
        print(f"[img]   ✗ 找不到貨號輸入框，截圖診斷...")
        await page.screenshot(path=str(ROOT / "data" / f"debug_basic_{code}.png"))
        return False

    try:
        await page.fill(prod_cd_sel, "", timeout=5000)
        await page.fill(prod_cd_sel, code, timeout=5000)
    except Exception as e:
        print(f"[img]   ✗ 填入貨號失敗: {e}")
        return False

    # 按查詢按鈕
    for btn_sel in [
        'button[id*="search"]', 'button:has-text("查詢")', 'button:has-text("검색")',
        '#btnSearch', 'input[type="button"][value*="查詢"]',
        'a:has-text("查詢")', '.btn-search',
    ]:
        try:
            btn = await page.query_selector(btn_sel)
            if btn and await btn.is_visible():
                await btn.click(timeout=3000)
                await page.wait_for_timeout(2000)
                break
        except Exception:
            pass
    else:
        # 嘗試按 Enter
        try:
            inp = await page.query_selector(prod_cd_sel)
            if inp:
                await inp.press("Enter")
                await page.wait_for_timeout(2000)
        except Exception:
            pass

    # 等待搜尋結果（找列表 row）
    await page.wait_for_timeout(1500)

    # 點第一筆結果
    clicked = False
    for row_sel in [
        'tr.rows td:first-child', 'td[data-col="PROD_CD"]',
        f'td:has-text("{code}")', f'a:has-text("{code}")',
        '.grid-row:first-child', 'tbody tr:first-child td',
    ]:
        try:
            rows = await page.locator(row_sel).all()
            for r in rows:
                txt = (await r.inner_text()).strip()
                if code in txt:
                    await r.click(timeout=3000)
                    clicked = True
                    print(f"[img]   點擊結果 row")
                    await page.wait_for_timeout(2000)
                    break
            if clicked:
                break
        except Exception:
            pass

    if not clicked:
        print(f"[img]   ✗ 找不到 {code} 的搜尋結果")
        await page.screenshot(path=str(ROOT / "data" / f"debug_result_{code}.png"))
        return False

    # 在詳情頁找圖片
    return await _extract_and_save_image(page, code, out_dir)


async def _extract_and_save_image(page, code: str, out_dir: Path) -> bool:
    """在目前頁面中找產品圖並儲存"""

    await page.wait_for_timeout(1000)

    # 找圖片 img 元素（各種可能的 selector）
    img_sels = [
        'img[id*="IMG"]', 'img[id*="img"]', 'img[id*="PHOTO"]',
        'img[id*="photo"]', 'img[id*="prod"]', 'img[id*="PROD"]',
        '.product-img img', '.prod-img img', '.item-img img',
        '#img_area img', '#photo_area img', '.detail-img img',
        'img[src*="upload"]', 'img[src*="file"]', 'img[src*="product"]',
    ]

    img_src = None
    for sel in img_sels:
        try:
            els = await page.query_selector_all(sel)
            for el in els:
                src = await el.get_attribute("src") or ""
                if src and not src.endswith("no_image.gif") and len(src) > 10:
                    if "ecount.com" in src or src.startswith("http"):
                        img_src = src
                        break
                    elif src.startswith("data:image"):
                        img_src = src
                        break
            if img_src:
                break
        except Exception:
            pass

    if not img_src:
        # 用 JS 找所有圖片
        imgs = await page.evaluate("""() => {
            return Array.from(document.querySelectorAll('img'))
                .map(i => ({src: i.src, w: i.naturalWidth, h: i.naturalHeight, id: i.id, cls: i.className}))
                .filter(i => i.w > 50 && i.h > 50 && !i.src.includes('icon') && !i.src.includes('logo'))
        }""")
        print(f"[img]   找到 {len(imgs)} 張圖（>50px）")
        for img in imgs:
            print(f"    {img['w']}x{img['h']} id={img['id']} {img['src'][:80]}")
        if imgs:
            # 取最大的那張
            imgs.sort(key=lambda x: x['w'] * x['h'], reverse=True)
            img_src = imgs[0]['src']

    if not img_src:
        print(f"[img]   ✗ {code} 詳情頁找不到圖片")
        await page.screenshot(path=str(ROOT / "data" / f"debug_detail_{code}.png"))
        return False

    out_path = out_dir / f"{code}A.jpg"
    print(f"[img]   圖片 src: {img_src[:80]}")

    # data URI
    if img_src.startswith("data:image"):
        import base64
        m = re.match(r"data:image/(\w+);base64,(.+)", img_src, re.DOTALL)
        if m:
            ext = m.group(1)
            data = base64.b64decode(m.group(2))
            out_path = out_dir / f"{code}A.{ext}"
            out_path.write_bytes(data)
            print(f"[img]   ✅ 已儲存（data URI）: {out_path.name}")
            return True

    # HTTP 下載
    try:
        cookies_list = await page.context.cookies()
        cookie_dict = {c["name"]: c["value"] for c in cookies_list}
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": "https://oapiIB.ecount.com/",
        }
        # 補全相對路徑
        if img_src.startswith("//"):
            img_src = "https:" + img_src
        elif img_src.startswith("/"):
            img_src = "https://oapiIB.ecount.com" + img_src

        resp = httpx.get(img_src, cookies=cookie_dict, headers=headers,
                         timeout=20, follow_redirects=True)
        if resp.status_code == 200 and len(resp.content) > 500:
            # 偵測副檔名
            ct = resp.headers.get("content-type", "")
            ext = "jpg"
            if "png" in ct:
                ext = "png"
            elif "webp" in ct:
                ext = "webp"
            elif "gif" in ct:
                ext = "gif"
            out_path = out_dir / f"{code}A.{ext}"
            out_path.write_bytes(resp.content)
            print(f"[img]   ✅ 已儲存（HTTP {resp.status_code}）: {out_path.name} {len(resp.content)//1024}KB")
            return True
        else:
            print(f"[img]   ✗ HTTP {resp.status_code} / {len(resp.content)} bytes")
    except Exception as e:
        print(f"[img]   ✗ 下載失敗: {e}")

    # 最後備案：截元素圖
    try:
        for sel in img_sels:
            el = await page.query_selector(sel)
            if el:
                await el.screenshot(path=str(out_dir / f"{code}A.png"))
                print(f"[img]   ✅ 已儲存（元素截圖）: {code}A.png")
                return True
    except Exception:
        pass

    return False


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

async def run(codes: list[str]):
    from playwright.async_api import async_playwright

    out_dir = AD_IMAGE_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"[img] 輸出目錄: {out_dir}")
    print(f"[img] 待下載: {codes}")

    # 過濾已有圖片的貨號
    need = []
    for code in codes:
        existing = list(out_dir.glob(f"{code}*.jpg")) + list(out_dir.glob(f"{code}*.png")) + list(out_dir.glob(f"{code}*.webp"))
        if existing:
            print(f"[img] ⏭  {code} 已有圖片：{[f.name for f in existing]}")
        else:
            need.append(code)

    if not need:
        print("[img] 所有貨號都已有圖片，無需下載")
        return

    print(f"[img] 需要下載 {len(need)} 筆：{need}")

    if not launch_chrome_if_needed():
        print("[img] ✗ Chrome 無法啟動")
        return

    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        ec_sid = await ensure_logged_in(page)
        if not ec_sid:
            print("[img] ✗ 無法取得 Ecount session")
            await browser.close()
            return

        ok = await _navigate_to_product_basic(page, ec_sid)
        if not ok:
            await browser.close()
            return

        success, fail = [], []
        for code in need:
            result = await _get_product_image(page, code, out_dir)
            (success if result else fail).append(code)
            await page.wait_for_timeout(500)

        print(f"\n{'='*50}")
        print(f"✅ 成功 {len(success)} 筆：{success}")
        if fail:
            print(f"❌ 失敗 {len(fail)} 筆：{fail}")
            print(f"   失敗的貨號請手動截圖後放到 {out_dir}")

        await browser.close()


def main():
    if "--all-cheap" in sys.argv:
        codes = CHEAP_CODES
    else:
        codes = [a for a in sys.argv[1:] if not a.startswith("--")]
        if not codes:
            print("用法: python scripts/fetch_ecount_product_images.py Z3300 Z3331 ...")
            print("      python scripts/fetch_ecount_product_images.py --all-cheap")
            sys.exit(1)

    asyncio.run(run(codes))


if __name__ == "__main__":
    main()
