"""
從 Ecount 庫存收支明細同步庫存變更紀錄 → data/sales_detail.db (inventory_changes 表)

紀錄所有入庫、出庫、調整等異動，用於分析銷售速度、補貨預測等。

執行方式：
  python -m scripts.sync_inventory_changes              # 同步本月
  python -m scripts.sync_inventory_changes --months 3   # 同步最近 3 個月
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

# 可能的頁面名稱
_MENU_TEXTS = ["庫存收支明細", "庫存異動明細", "庫存變更紀錄", "入出庫明細"]


# ---------------------------------------------------------------------------
# DB
# ---------------------------------------------------------------------------

def _init_db():
    with sqlite3.connect(str(DB_PATH)) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS inventory_changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                slip_no TEXT,
                type TEXT,
                prod_cd TEXT NOT NULL,
                prod_name TEXT,
                warehouse TEXT,
                qty_in INTEGER DEFAULT 0,
                qty_out INTEGER DEFAULT 0,
                balance INTEGER DEFAULT 0,
                customer TEXT,
                note TEXT,
                synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(date, slip_no, prod_cd, type)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ic_date ON inventory_changes(date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ic_prod ON inventory_changes(prod_cd)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ic_type ON inventory_changes(type)")


def _upsert_rows(rows: list[dict]) -> int:
    inserted = 0
    with sqlite3.connect(str(DB_PATH)) as conn:
        for r in rows:
            try:
                conn.execute("""
                    INSERT OR IGNORE INTO inventory_changes
                    (date, slip_no, type, prod_cd, prod_name, warehouse, qty_in, qty_out, balance, customer, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    r.get("date", ""),
                    r.get("slip_no", ""),
                    r.get("type", ""),
                    r.get("prod_cd", ""),
                    r.get("prod_name", ""),
                    r.get("warehouse", ""),
                    r.get("qty_in", 0),
                    r.get("qty_out", 0),
                    r.get("balance", 0),
                    r.get("customer", ""),
                    r.get("note", ""),
                ))
            except Exception:
                pass
        conn.commit()
        inserted = conn.total_changes
    return inserted


# ---------------------------------------------------------------------------
# 導航
# ---------------------------------------------------------------------------

async def _find_inventory_change_page(page, ec_sid: str) -> bool:
    """嘗試找到庫存收支明細頁面"""
    cfg = load_web_config()

    # 不用 saved hash（saved hash 會直接載入結果，跳過設定畫面）
    # 直接從首頁點書籤，確保回到設定畫面
    for menu_text in _MENU_TEXTS:
        try:
            clicked = await page.evaluate(f"""
                () => {{
                    const links = document.querySelectorAll('a');
                    for (const a of links) {{
                        if (a.textContent.trim() === '{menu_text}' && a.id && a.id.includes('bookmark')) {{
                            a.click();
                            return true;
                        }}
                    }}
                    // fallback: 任何包含文字的 a
                    for (const a of links) {{
                        if (a.textContent.trim() === '{menu_text}' && a.offsetParent !== null) {{
                            a.click();
                            return true;
                        }}
                    }}
                    return false;
                }}
            """)
            if clicked:
                print(f"  點擊「{menu_text}」書籤...")
                await page.wait_for_load_state("networkidle", timeout=15000)
                await page.wait_for_timeout(3000)
                # 這個頁面用「查詢(F8)」按鈕，不是 #header_search
                has_page = await page.evaluate("""
                    () => {
                        const btns = document.querySelectorAll('button, input[type="button"]');
                        for (const b of btns) {
                            if (b.textContent.includes('查詢')) return true;
                        }
                        return false;
                    }
                """)
                if has_page:
                    m = re.search(r"(#.+)", page.url)
                    if m:
                        cfg["inventory_change_hash"] = m.group(1)
                        save_web_config(cfg)
                    print(f"  ✓ 「{menu_text}」載入成功")
                    return True
        except Exception as e:
            print(f"  [warn] 點擊「{menu_text}」失敗: {e}")

    print("  ✗ 無法找到庫存收支明細頁面")
    print("  請在 Ecount Chrome 中手動打開「庫存收支明細」頁面，然後重新執行")
    return False


async def _set_date_range(page, start_date: str, end_date: str):
    """設定查詢日期"""
    try:
        await page.evaluate(f"""
            () => {{
                const inputs = document.querySelectorAll('input[type="text"]');
                let set1 = false;
                for (const inp of inputs) {{
                    if (/\\d{{4}}\\/\\d{{2}}\\/\\d{{2}}/.test(inp.value)) {{
                        if (!set1) {{
                            inp.value = '{start_date}';
                            inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                            set1 = true;
                        }} else {{
                            inp.value = '{end_date}';
                            inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                            break;
                        }}
                    }}
                }}
            }}
        """)
        print(f"  日期範圍：{start_date} ~ {end_date}")
    except Exception as e:
        print(f"  [warn] 設定日期失敗: {e}")


async def _select_by_date(page):
    """點選「按日期」radio"""
    try:
        clicked = await page.evaluate("""
            () => {
                // 找「按日期」radio 或 label
                const labels = document.querySelectorAll('label, span');
                for (const el of labels) {
                    if (el.textContent.trim() === '按日期') {
                        el.click();
                        return true;
                    }
                }
                // fallback: 找 radio input
                const radios = document.querySelectorAll('input[type="radio"]');
                for (const r of radios) {
                    const lbl = r.closest('label') || r.parentElement;
                    if (lbl && lbl.textContent.includes('按日期')) {
                        r.click();
                        return true;
                    }
                }
                return false;
            }
        """)
        if clicked:
            print("  ✓ 選擇「按日期」")
            await page.wait_for_timeout(1000)
        else:
            print("  [warn] 找不到「按日期」選項")
    except Exception as e:
        print(f"  [warn] 選擇按日期失敗: {e}")


async def _set_warehouse(page, warehouse_code: str):
    """倉庫欄填入代碼（如 101）"""
    try:
        filled = await page.evaluate(f"""
            () => {{
                // 找 placeholder 含「倉庫」的 input
                const inputs = document.querySelectorAll('input[type="text"]');
                for (const inp of inputs) {{
                    if ((inp.placeholder || '').includes('倉庫')) {{
                        inp.value = '{warehouse_code}';
                        inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                        inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                        return 'placeholder';
                    }}
                }}
                // fallback: 找 label/span 含「倉庫」旁邊的 input
                const els = document.querySelectorAll('td, th, label, span, div');
                for (const el of els) {{
                    const t = (el.textContent || '').trim();
                    if (t === '倉庫' || t.startsWith('倉庫')) {{
                        // 同一行或下一個兄弟的 input
                        const row = el.closest('tr') || el.closest('div') || el.parentElement;
                        if (row) {{
                            const inp = row.querySelector('input[type="text"]');
                            if (inp) {{
                                inp.value = '{warehouse_code}';
                                inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                                inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                                return 'label';
                            }}
                        }}
                    }}
                }}
                return '';
            }}
        """)
        if filled:
            print(f"  倉庫：{warehouse_code}（{filled}）")
        else:
            print(f"  [warn] 找不到倉庫欄位")
    except Exception as e:
        print(f"  [warn] 設定倉庫失敗: {e}")


async def _click_query(page):
    # 這個頁面的查詢按鈕文字是「查詢(F8)」
    try:
        clicked = await page.evaluate("""
            () => {
                const btns = document.querySelectorAll('button, input[type="button"]');
                for (const b of btns) {
                    if (b.textContent.includes('查詢')) {
                        b.click();
                        return true;
                    }
                }
                return false;
            }
        """)
        if clicked:
            print("  點擊查詢...")
            await page.wait_for_timeout(8000)
            return
    except Exception:
        pass
    # fallback: F8
    try:
        await page.keyboard.press("F8")
        print("  按 F8 查詢...")
        await page.wait_for_timeout(8000)
    except Exception:
        pass


async def _download_excel(page) -> Path | None:
    tmp_path = ROOT / "data" / "_tmp_inv_changes.xlsx"
    for sel in [
        'button:has-text("Excel")',
        'text="Excel"',
        '[data-cid*="Excel"]',
        ':text("Excel")',
    ]:
        try:
            locs = await page.locator(sel).all()
            for loc in locs:
                if await loc.is_visible():
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
    """解析庫存收支明細 Excel"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        ws = wb.active
        rows_data = list(ws.iter_rows(values_only=True))
        if not rows_data:
            return []

        # 找表頭
        header_idx = -1
        headers = []
        for i, row in enumerate(rows_data):
            row_str = " ".join(str(c or "") for c in row)
            if "品項" in row_str or "入庫" in row_str or "出庫" in row_str:
                headers = [str(c or "").strip() for c in row]
                header_idx = i
                break
        if header_idx < 0 and len(rows_data) > 1:
            headers = [str(c or "").strip() for c in rows_data[1]]
            header_idx = 1

        print(f"  Excel 表頭 (第{header_idx+1}行): {headers[:15]}")

        # 欄位對應
        col_map = {}
        for ci, h in enumerate(headers):
            hl = h.replace(" ", "")
            if "日期" in hl and "date" not in col_map:
                col_map["date"] = ci
            elif "單據" in hl or "單號" in hl or "號碼" in hl:
                col_map["slip_no"] = ci
            elif "類型" in hl or "區分" in hl or "單據種類" in hl:
                col_map["type"] = ci
            elif "品項編碼" in hl or "品項代碼" in hl:
                col_map["prod_cd"] = ci
            elif "品名" in hl or "品項名稱" in hl:
                col_map["prod_name"] = ci
            elif "倉庫" in hl:
                col_map["warehouse"] = ci
            elif "入庫" in hl and "qty_in" not in col_map:
                col_map["qty_in"] = ci
            elif "出庫" in hl and "qty_out" not in col_map:
                col_map["qty_out"] = ci
            elif "庫存" in hl or "餘額" in hl or "結存" in hl:
                col_map["balance"] = ci
            elif "客戶" in hl or "供應商" in hl:
                col_map["customer"] = ci
            elif "備註" in hl or "摘要" in hl:
                col_map["note"] = ci

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
                continue

            date_val = _get("date")
            if isinstance(date_val, datetime):
                date_str = date_val.strftime("%Y-%m-%d")
            else:
                raw = str(date_val or "").strip()
                dm = re.match(r"(\d{4}[/-]\d{2}[/-]\d{2})", raw)
                date_str = dm.group(1).replace("/", "-") if dm else raw

            def _int(key):
                v = _get(key)
                try:
                    return int(float(v)) if v else 0
                except (ValueError, TypeError):
                    return 0

            results.append({
                "date": date_str,
                "slip_no": str(_get("slip_no") or "").strip(),
                "type": str(_get("type") or "").strip(),
                "prod_cd": prod_cd.upper(),
                "prod_name": str(_get("prod_name") or "").strip(),
                "warehouse": str(_get("warehouse") or "").strip(),
                "qty_in": _int("qty_in"),
                "qty_out": _int("qty_out"),
                "balance": _int("balance"),
                "customer": str(_get("customer") or "").strip(),
                "note": str(_get("note") or "").strip(),
            })

        print(f"  Excel 解析完成：{len(results)} 筆")
        return results

    except Exception as e:
        print(f"  Excel 解析失敗: {e}")
        import traceback; traceback.print_exc()
        return []


async def _parse_table(page) -> list[dict]:
    """從頁面表格直接解析庫存變更資料"""
    # 先滾動載入全部
    for _ in range(30):
        try:
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await page.wait_for_timeout(500)
        except Exception:
            break

    data = await page.evaluate(r"""
        () => {
            const results = [];
            const tables = document.querySelectorAll('table');

            for (const table of tables) {
                const rows = table.querySelectorAll('tr');
                if (rows.length < 10) continue;

                // 找表頭
                const headers = Array.from(rows[0].querySelectorAll('th, td'))
                    .map(c => (c.textContent || '').trim());
                if (!headers.some(h => h.includes('日期')) || !headers.some(h => h.includes('入庫'))) continue;

                // 表頭: ['品項編碼', '品項名稱', '規格', '日期', '入庫數量', '出庫數量', '餘量']
                // 品項行: 7 欄 [code, name, spec, date, qty_in, qty_out, balance]
                // 日期行: 5 欄 ['', date, qty_in, qty_out, balance]（品項+名稱+規格被 rowspan 合併）

                let currentCode = '';
                let currentName = '';
                for (let i = 1; i < rows.length; i++) {
                    const cells = Array.from(rows[i].querySelectorAll('td'));
                    if (cells.length < 3) continue;

                    let dateStr, qtyIn, qtyOut, balance;

                    if (cells.length >= 7) {
                        // 品項行（完整 7 欄）
                        const code = (cells[0].textContent || '').trim();
                        if (code && /^[A-Za-z]/.test(code)) {
                            currentCode = code.toUpperCase();
                            currentName = (cells[1].textContent || '').trim();
                        }
                        dateStr = (cells[3].textContent || '').trim();
                        qtyIn = (cells[4].textContent || '').trim();
                        qtyOut = (cells[5].textContent || '').trim();
                        balance = (cells[6].textContent || '').trim();
                    } else if (cells.length >= 4 && cells.length <= 6) {
                        // 日期行（合併後 4~5 欄）
                        // 找第一個含日期格式的 cell
                        let dateIdx = -1;
                        for (let j = 0; j < cells.length; j++) {
                            if (/\d{4}\/\d{2}\/\d{2}/.test(cells[j].textContent)) {
                                dateIdx = j;
                                break;
                            }
                        }
                        if (dateIdx < 0) continue;
                        dateStr = (cells[dateIdx].textContent || '').trim();
                        qtyIn = (cells[dateIdx + 1] ? cells[dateIdx + 1].textContent : '').trim();
                        qtyOut = (cells[dateIdx + 2] ? cells[dateIdx + 2].textContent : '').trim();
                        balance = (cells[dateIdx + 3] ? cells[dateIdx + 3].textContent : '').trim();
                    } else {
                        continue;
                    }

                    if (!currentCode) continue;
                    if (!dateStr || !/\d{4}\/\d{2}\/\d{2}/.test(dateStr)) continue;

                    const qi = parseFloat((qtyIn || '0').replace(/,/g, '')) || 0;
                    const qo = parseFloat((qtyOut || '0').replace(/,/g, '')) || 0;
                    const bal = parseFloat((balance || '0').replace(/,/g, '')) || 0;

                    if (qi === 0 && qo === 0) continue;

                    results.push({
                        date: dateStr.replace(/\//g, '-'),
                        prod_cd: currentCode,
                        prod_name: currentName,
                        qty_in: qi,
                        qty_out: qo,
                        balance: bal,
                        slip_no: '',
                        type: qi > 0 ? '入庫' : '出庫',
                        warehouse: '101',
                        customer: '',
                        note: ''
                    });
                }

                if (results.length > 0) break;
            }
            return results;
        }
    """)

    print(f"  頁面表格解析：{len(data)} 筆")
    return data


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

async def sync(months: int = 1):
    from playwright.async_api import async_playwright

    _init_db()
    launch_chrome_if_needed()

    now = datetime.now()
    end_date = now.strftime("%Y/%m/%d")
    start_dt = now.replace(day=1) - timedelta(days=(months - 1) * 30)
    start_dt = start_dt.replace(day=1)
    start_date = start_dt.strftime("%Y/%m/%d")

    print(f"[inv-change] 同步庫存變更：{start_date} ~ {end_date}")

    async with async_playwright() as p:
        browser, page = await connect_get_page(p)
        if not page:
            print("[inv-change] ✗ 無法連接 Chrome")
            return False

        ec_sid = await ensure_logged_in(page)
        if not ec_sid:
            print("[inv-change] ✗ 未登入 Ecount")
            return False

        # 先載入 ERP 首頁（讓書籤列出現）
        try:
            url_base = f"{ERP_URL}?w_flag=1&ec_req_sid={ec_sid}"
            await page.goto("about:blank", timeout=5000)
            await page.goto(url_base, timeout=25000)
            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(3000)
        except Exception as e:
            print(f"  [warn] 載入首頁: {e}")

        if not await _find_inventory_change_page(page, ec_sid):
            return False

        # 用 Playwright 操作：選按日期 → 倉庫 101 → 查詢
        print("  設定查詢條件...")

        # 1. 點「按日期」label（force=True 繞過 radio input 攔截）
        try:
            loc = page.locator('label:has-text("按日期")').first
            await loc.click(force=True)
            await page.wait_for_timeout(2000)
            print("  ✓ 選擇「按日期」")
        except Exception as e:
            print(f"  [warn] 按日期: {e}")

        # 2. 倉庫填 101（找 placeholder 含「倉庫」的 input，用 fill 填入）
        try:
            wh_input = page.locator('input[placeholder*="倉庫"]').first
            if await wh_input.is_visible(timeout=3000):
                await wh_input.click()
                await wh_input.fill("101")
                await page.keyboard.press("Enter")
                await page.wait_for_timeout(1000)
                print("  ✓ 倉庫：101")
            else:
                print("  [warn] 找不到倉庫輸入框")
        except Exception as e:
            print(f"  [warn] 倉庫: {e}")

        # 3. 按 F8 查詢（查詢後按鈕會變成「查詢/篩選(F3)」）
        await page.keyboard.press("F8")
        print("  查詢中...")
        # 等結果表格出現
        for _ in range(20):
            await page.wait_for_timeout(1000)
            has_data = await page.evaluate("""
                () => {
                    const tds = document.querySelectorAll('td');
                    for (const td of tds) {
                        if (/\\d{4}\\/\\d{2}\\/\\d{2}/.test(td.textContent)) return true;
                    }
                    return false;
                }
            """)
            if has_data:
                print("  ✓ 查詢結果已載入")
                break
        await page.wait_for_timeout(2000)

        # 直接從頁面表格解析（不依賴 Excel 下載）
        print("  解析頁面表格...")
        rows = await _parse_table(page)

    if not rows:
        print("[inv-change] ✗ 沒有解析到資料")
        return False

    inserted = _upsert_rows(rows)
    total_in = sum(r["qty_in"] for r in rows)
    total_out = sum(r["qty_out"] for r in rows)
    products = set(r["prod_cd"] for r in rows)
    print(f"[inv-change] ✓ 共 {len(rows)} 筆，新增 {inserted} 筆")
    print(f"[inv-change]   品項 {len(products)} 個，入庫 {total_in}，出庫 {total_out}")

    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--months", type=int, default=1, help="同步幾個月")
    args = parser.parse_args()
    ok = asyncio.run(sync(months=args.months))
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
