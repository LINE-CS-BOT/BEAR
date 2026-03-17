"""
CDP 診斷腳本 v6 - 直接填表查詢
前提：Chrome 已手動開啟「庫存情況」頁面（有倉庫輸入欄）
執行：python scripts/diagnose_sync.py
"""
import asyncio
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

CDP_URL = "http://localhost:9222"


async def diagnose():
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        print("=== 1. CDP 連線 ===")
        browser = await p.chromium.connect_over_cdp(CDP_URL)
        print("    ✓ 連線成功")

        print("\n=== 2. 找 Ecount 分頁 ===")
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "ecount.com" in pg.url:
                    page = pg
                    break
            if page:
                break
        if not page:
            print("    ✗ 找不到 Ecount tab")
            await browser.close()
            return
        print(f"    URL: {page.url[:100]}")

        print("\n=== 3. 確認表單已載入 ===")
        search_btn = await page.query_selector('#searchGroup')
        if not search_btn:
            print("    ✗ 找不到 #searchGroup，請先手動點開「庫存情況」頁面")
            await browser.close()
            return
        print("    ✓ 找到 #searchGroup 查詢按鈕")

        print("\n=== 4. 顯示所有 SELECT 元素（找 FN 包含0）===")
        selects = await page.evaluate("""
            () => {
                const result = [];
                document.querySelectorAll('select').forEach(sel => {
                    const opts = Array.from(sel.options).map(o => o.text.trim() + '=' + o.value);
                    result.push({
                        id: sel.id, name: sel.name,
                        options: opts,
                        currentValue: sel.value
                    });
                });
                return result;
            }
        """)
        print(f"    SELECT 元素共 {len(selects)} 個:")
        for s in selects:
            print(f"      id={s['id']!r} name={s['name']!r} opts={s['options']} current={s['currentValue']!r}")

        print("\n=== 5. 顯示所有 INPUT[type=text/number] 元素 ===")
        text_inputs = await page.evaluate("""
            () => {
                const result = [];
                document.querySelectorAll('input[type=text], input[type=number], input:not([type])').forEach(el => {
                    result.push({
                        id: el.id, name: el.name,
                        placeholder: el.placeholder,
                        value: el.value,
                        class: el.className.substring(0, 40)
                    });
                });
                return result;
            }
        """)
        print(f"    文字 INPUT 共 {len(text_inputs)} 個:")
        for el in text_inputs:
            print(f"      {el}")

        print("\n=== 6. 填入倉庫 = 101 ===")
        wh_filled = False
        # 嘗試各種方式找倉庫欄位
        wh_candidates = [
            "input[placeholder='倉庫']",
            "input[placeholder*='倉庫']",
        ]
        for sel in wh_candidates:
            try:
                els = await page.query_selector_all(sel)
                if els:
                    print(f"    找到 {len(els)} 個 {sel!r}")
                    el = els[0]
                    await el.triple_click()
                    await el.type("101")
                    val = await el.input_value()
                    print(f"    ✓ 填入後值 = {val!r}")
                    wh_filled = True
                    # 觸發 blur/change 事件
                    await el.press("Tab")
                    await page.wait_for_timeout(500)
                    break
            except Exception as e:
                print(f"    {sel} 失敗: {e}")

        if not wh_filled:
            print("    ✗ 找不到倉庫欄位，繼續嘗試查詢...")

        print("\n=== 7. 設定「包含0」===")
        # 嘗試找 FN (數量) 相關的 select
        fn_result = await page.evaluate("""
            () => {
                // 找所有 select，印出所有選項
                const allSelects = [];
                document.querySelectorAll('select').forEach(sel => {
                    allSelects.push({
                        id: sel.id,
                        name: sel.name,
                        options: Array.from(sel.options).map(o => ({t: o.text.trim(), v: o.value}))
                    });
                });

                // 嘗試找包含「包含」或「0」的 option
                for (const s of allSelects) {
                    for (const opt of s.options) {
                        if (opt.t.includes('包含') || opt.t === '包含0') {
                            const el = document.querySelector(`select[id="${s.id}"]`) ||
                                       document.querySelector(`select[name="${s.name}"]`);
                            if (el) {
                                el.value = opt.v;
                                el.dispatchEvent(new Event('change', {bubbles: true}));
                                return `✓ 設定 ${s.id}/${s.name} = ${opt.v} (${opt.t})`;
                            }
                        }
                    }
                }

                // 找 label 含「包含」的旁邊 input
                const allLabels = Array.from(document.querySelectorAll('label, span, td, th'));
                for (const lbl of allLabels) {
                    if (lbl.textContent.includes('包含') && lbl.textContent.length < 30) {
                        return `找到包含文字: "${lbl.textContent.trim()}" tag=${lbl.tagName}`;
                    }
                }

                return `未找到包含0選項，select 列表: ${JSON.stringify(allSelects.map(s=>s.id+'/'+s.name))}`;
            }
        """)
        print(f"    {fn_result}")

        print("\n=== 8. 點擊查詢按鈕 ===")
        try:
            btn = page.locator('#searchGroup')
            cnt = await btn.count()
            print(f"    找到 #searchGroup x{cnt}")
            await btn.last.click(timeout=5000)
            print("    ✓ 點擊完成")
        except Exception as e:
            print(f"    ✗ 點擊失敗: {e}")

        print("\n=== 9. 等待查詢結果（最多 30 秒）===")
        found = False
        for attempt in range(30):
            await page.wait_for_timeout(1000)
            try:
                el = await page.query_selector("#grid-main")
                if el:
                    rows = await page.evaluate(
                        "() => { const t=document.getElementById('grid-main'); return t?t.rows.length:0; }"
                    )
                    if rows > 1:
                        print(f"    ✓ {attempt+1}s: 找到 #grid-main，{rows} 行")
                        found = True
                        break
            except Exception:
                pass
            if attempt % 5 == 4:
                print(f"    {attempt+1}s: 等待中...")

        if not found:
            print("    ✗ 30 秒後仍未找到結果")
            await browser.close()
            return

        print("\n=== 10. 確認欄位結構 ===")
        headers = await page.evaluate("""
            () => {
                const t = document.getElementById('grid-main');
                if (!t || !t.rows[0]) return [];
                return Array.from(t.rows[0].cells).map(c => c.textContent.trim().substring(0, 20));
            }
        """)
        print(f"    表頭: {headers}")

        sample = await page.evaluate("""
            () => {
                const t = document.getElementById('grid-main');
                if (!t) return [];
                const rows = [];
                for (let i = 1; i < Math.min(t.rows.length, 5); i++) {
                    const cells = Array.from(t.rows[i].cells).map(c => c.textContent.trim().substring(0, 15));
                    rows.push(cells);
                }
                return rows;
            }
        """)
        print(f"    前幾行: {sample}")

        # 找可售庫存欄位
        sale_col = next((i for i, h in enumerate(headers) if '可售' in h), None)
        qty_col  = next((i for i, h in enumerate(headers) if '庫存' in h and '可售' not in h), None)
        print(f"\n    → 可售庫存欄: {sale_col}（{headers[sale_col] if sale_col is not None else 'N/A'}）")
        print(f"    → 庫存欄:     {qty_col}（{headers[qty_col] if qty_col is not None else 'N/A'}）")
        print(f"\n    ✅ 診斷成功！sync 腳本應使用欄 {sale_col} 提取可售庫存")

        print("\n診斷完成")
        await browser.close()


asyncio.run(diagnose())
