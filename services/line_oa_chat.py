"""
LINE OA Manager 對話紀錄讀取

透過 Playwright CDP 連接 LINE OA Manager Chrome（port 9223），
搜尋客戶並讀取聊天紀錄。

用途：
  - 真人接管釋放時，自動讀取接管期間的對話
  - 補全 chat_history.db 中缺少的真人回覆
"""

import asyncio
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

_LINE_OA_CDP = "http://127.0.0.1:9223"
_LINE_OA_EMAIL = "1127bear@ms93.url.com.tw"
_LINE_OA_PASS = "Bear671027"


async def _ensure_logged_in(page) -> bool:
    """確認已登入 LINE OA，未登入則自動登入"""
    url = page.url
    if "chat.line.biz" in url and "login" not in url:
        return True  # 已登入

    print("[line-oa] 未登入，嘗試自動登入...")
    await page.goto("https://chat.line.biz/", timeout=20000)
    await page.wait_for_timeout(5000)
    url = page.url

    # access.line.me SSO 頁面 — 直接點「登录」按鈕
    if "access.line.me" in url:
        try:
            btn = page.locator('button:has-text("登录"), button:has-text("Log in"), button:has-text("登入")').first
            await btn.click()
            print("[line-oa] 已點 SSO 登录")
            await page.wait_for_timeout(8000)
            if "chat.line.biz" in page.url and "login" not in page.url:
                print("[line-oa] ✓ SSO 登入成功")
                return True
        except Exception as e:
            print(f"[line-oa] SSO 登入失敗: {e}")

    # LINE Business ID 登入頁 — 點 LINE account → SSO 登入
    if "account.line.biz" in page.url:
        try:
            # 點綠色的「LINE account」按鈕
            line_btn = page.locator('button:has-text("LINE account"), a:has-text("LINE account")').first
            if await line_btn.count() > 0:
                await line_btn.click()
                print("[line-oa] 已點 LINE account")
                await page.wait_for_timeout(5000)

            # 跳到 SSO 頁面 — 點「登录」
            if "access.line.me" in page.url:
                btn = page.locator('button:has-text("登录"), button:has-text("Log in"), button:has-text("登入")').first
                await btn.click()
                print("[line-oa] 已點 SSO 登录")
                await page.wait_for_timeout(8000)
                if "chat.line.biz" in page.url and "login" not in page.url:
                    print("[line-oa] ✓ 自動登入成功")
                    return True
        except Exception as e:
            print(f"[line-oa] 登入失敗: {e}")

    print(f"[line-oa] ✗ 自動登入失敗，目前 URL: {page.url[:80]}")
    return False


def _spawn_line_oa_chrome() -> bool:
    """Chrome 沒開就自動啟動（用 start.bat 同一組參數）。回傳是否成功 ready。"""
    import subprocess, time, urllib.request
    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if not Path(chrome_exe).exists():
        chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    user_data = str(Path(__file__).parent.parent / "data" / "line_chrome_session")
    try:
        subprocess.Popen(
            [chrome_exe, "--remote-debugging-port=9223",
             f"--user-data-dir={user_data}",
             "--no-first-run", "--disable-default-apps",
             "https://chat.line.biz/"],
            creationflags=0x00000008,  # DETACHED_PROCESS
        )
    except Exception as e:
        print(f"[line-oa] 啟動 Chrome 失敗：{e}")
        return False
    # poll 最多 20 秒等 CDP ready
    for _ in range(40):
        try:
            urllib.request.urlopen("http://127.0.0.1:9223/json/version", timeout=1)
            print("[line-oa] ✓ Chrome 已啟動，CDP ready")
            return True
        except Exception:
            time.sleep(0.5)
    print("[line-oa] ✗ Chrome 啟動後 CDP 仍未 ready")
    return False


async def _get_line_oa_page():
    """連接 LINE OA Manager Chrome，回傳 (browser, page) 或 (None, None)"""
    from playwright.async_api import async_playwright
    p = await async_playwright().start()
    try:
        browser = await p.chromium.connect_over_cdp(_LINE_OA_CDP)
    except Exception as e:
        print(f"[line-oa] 無法連接 Chrome (port 9223): {e}，嘗試自動啟動…")
        if not await asyncio.to_thread(_spawn_line_oa_chrome):
            await p.stop()
            return None, None
        try:
            browser = await p.chromium.connect_over_cdp(_LINE_OA_CDP)
        except Exception as e2:
            print(f"[line-oa] 自動啟動後仍連不上：{e2}")
            await p.stop()
            return None, None

    # 找 LINE OA tab（剛 spawn 時可能還在載入，retry 最多 15 秒）
    page = None
    for _retry in range(30):
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "line.biz" in pg.url or "line.me" in pg.url:
                    page = pg
                    break
                if "chrome://" not in pg.url and pg.url not in ("about:blank", ""):
                    page = pg
            if page:
                break
        if page and "line.biz" in page.url:
            break
        await asyncio.sleep(0.5)

    if not page:
        print("[line-oa] 找不到可用的 tab")
        return None, None

    # 全自動點掉啟動時的 modal（背景 Chrome，使用者看不到）
    await _auto_accept_modals(page, rounds=6)

    # 確認已登入
    if not await _ensure_logged_in(page):
        return None, None

    # 登入後再清一次（登入後可能有公告 modal）
    await _auto_accept_modals(page, rounds=3)

    return browser, page


async def _auto_accept_modals(page, rounds: int = 3):
    """
    背景 Chrome 不讓使用者手動點，這裡自動按掉所有「確定/OK/好/繼續/知道了」類 button。
    每輪間隔 500ms，最多 rounds 輪。會 log 實際點到的按鈕文字以便除錯。
    同時處理 Playwright native dialog（beforeunload/alert/confirm）。
    """
    # Playwright native dialog
    async def _on_dialog(d):
        try:
            print(f"[line-oa] 自動 accept native dialog: {d.type} / {d.message[:60]}")
            await d.accept()
        except Exception:
            pass
    try:
        page.on("dialog", lambda d: asyncio.create_task(_on_dialog(d)))
    except Exception:
        pass

    total_clicked = 0
    for i in range(rounds):
        try:
            clicked_texts = await page.evaluate("""() => {
                const texts = ['確定','OK','好','好的','關閉','知道了','繼續','我知道了','下次再說','稍後','Got it','Continue','Close','略過'];
                const clicked = [];
                const sel = 'button, [role="button"], a.btn, .btn, [role="dialog"] *, .modal *';
                const nodes = Array.from(document.querySelectorAll(sel));
                for (const b of nodes) {
                    const rect = b.getBoundingClientRect();
                    if (rect.width === 0 || rect.height === 0) continue;
                    const style = window.getComputedStyle(b);
                    if (style.visibility === 'hidden' || style.display === 'none') continue;
                    const t = (b.innerText || b.textContent || '').trim();
                    if (!t || t.length > 20) continue;
                    if (texts.some(x => t === x || t.includes(x))) {
                        try { b.click(); clicked.push(t); } catch(e) {}
                    }
                }
                return clicked;
            }""")
            if clicked_texts:
                total_clicked += len(clicked_texts)
                print(f"[line-oa] 自動點 modal button (round {i+1}): {clicked_texts}")
        except Exception as e:
            print(f"[line-oa] _auto_accept_modals round {i+1} 失敗: {e}")
        await page.wait_for_timeout(500)

    # 殘餘遮罩清掉
    try:
        await page.evaluate("""() => {
            document.querySelectorAll('.modal, .modal-backdrop, [role="dialog"]').forEach(el => {
                const style = window.getComputedStyle(el);
                if (style.position === 'fixed' || style.position === 'absolute') {
                    el.style.display = 'none';
                    el.style.pointerEvents = 'none';
                }
            });
            document.body.classList.remove('modal-open');
            document.body.style.overflow = 'auto';
            document.body.style.paddingRight = '0';
        }""")
    except Exception:
        pass

    if total_clicked == 0:
        print("[line-oa] 無 modal 需要處理")


async def read_customer_chat(customer_name: str, max_messages: int = 30) -> list[dict]:
    """
    搜尋客戶並讀取聊天紀錄。

    回傳 [{"role": "customer"/"staff", "text": "...", "time": "..."}, ...]
    """
    browser, page = await _get_line_oa_page()
    if not page:
        return []

    try:
        # 1. 搜尋客戶（用 JS 繞過可能的元素遮擋）
        await page.evaluate("""() => {
            const el = document.querySelector('input[placeholder*="搜尋"], #chatListSearchInput');
            if (el) { el.focus(); el.click(); }
        }""")
        await page.wait_for_timeout(500)
        search = page.locator('input[placeholder*="搜尋"], #chatListSearchInput').first
        await search.fill("")
        await search.fill(customer_name)
        await page.wait_for_timeout(2000)

        # 2. 點擊搜尋結果
        result = page.locator(f'.list-group-item:has-text("{customer_name}")').first
        if await result.count() == 0:
            print(f"[line-oa] 找不到客戶「{customer_name}」")
            return []
        await result.click()
        await page.wait_for_timeout(3000)

        # 3. 讀取聊天內容
        messages = await page.evaluate(r"""(maxMsgs) => {
            const msgs = [];
            const items = document.querySelectorAll('.chat');
            for (const el of items) {
                // chat-reverse = 客服/bot 發的（右邊綠色）
                // chat-secondary = 客戶發的（左邊灰色）
                const isStaff = el.classList.contains('chat-reverse');
                const body = el.querySelector('.chat-body');
                if (!body) continue;

                // 取純文字（排除時間戳等）
                const moreEl = body.querySelector('.more');
                const textEl = moreEl || body;
                let text = '';
                // 取第一個文字節點或 p 標籤
                const pEl = textEl.querySelector('p');
                if (pEl) {
                    text = pEl.textContent?.trim() || '';
                } else {
                    text = textEl.textContent?.trim() || '';
                }

                // 取時間
                const timeEl = el.querySelector('.chat-date');
                const time = timeEl ? timeEl.textContent?.trim() : '';

                // 過濾空訊息和系統訊息
                if (!text || text.length < 1) continue;
                // 去掉尾部的時間和「已讀」
                text = text.replace(/已讀\s*$/, '').replace(/\d{1,2}:\d{2}\s*$/, '').trim();
                if (!text) continue;

                msgs.push({
                    role: isStaff ? 'staff' : 'customer',
                    text: text.slice(0, 500),
                    time: time,
                });
            }
            return msgs.slice(-maxMsgs);
        }""", max_messages)

        # 4. 清空搜尋（恢復原本的聊天列表）
        try:
            await search.fill("")
            await page.wait_for_timeout(500)
        except Exception:
            pass

        return messages

    except Exception as e:
        print(f"[line-oa] 讀取對話失敗: {e}")
        return []


def read_chat_sync(customer_name: str, max_messages: int = 30) -> list[dict]:
    """同步版本的 read_customer_chat"""
    return asyncio.run(read_customer_chat(customer_name, max_messages))


# ── 發送訊息 ─────────────────────────────────────────────


async def _dismiss_modals(page):
    """先點掉「確定」類按鈕（如過期提示），再隱藏殘餘遮罩。"""
    try:
        # 1. 嘗試點可見的 確定/OK/關閉 按鈕（處理登入過期提示等）
        await page.evaluate("""() => {
            const texts = ['確定', 'OK', '好', '關閉', '知道了'];
            const btns = Array.from(document.querySelectorAll(
                '.modal button, .modal a.btn, .modal .btn, [role="dialog"] button'
            ));
            for (const b of btns) {
                const t = (b.innerText || b.textContent || '').trim();
                if (texts.some(x => t === x || t.includes(x))) {
                    try { b.click(); } catch (e) {}
                }
            }
        }""")
        await page.wait_for_timeout(300)

        # 2. CSS 隱藏殘餘遮罩
        await page.evaluate("""() => {
            document.querySelectorAll('.modal, .modal-backdrop').forEach(el => {
                el.style.display = 'none';
                el.style.pointerEvents = 'none';
            });
            document.body.classList.remove('modal-open');
            document.body.style.overflow = 'auto';
            document.body.style.paddingRight = '0';
        }""")
    except Exception:
        pass


async def _open_chat(page, chat_name: str) -> bool:
    """搜尋並開啟聊天室（1:1 或群組），回傳是否成功"""
    try:
        # 先關掉彈窗和搜尋模式
        await _dismiss_modals(page)

        # 若目前卡在搜尋模式（只顯示最近搜尋，沒顯示聊天列表），按 × 退出
        try:
            await page.evaluate("""() => {
                const inp = document.querySelector('input[placeholder*="搜尋"], #chatListSearchInput');
                if (inp && inp.value) {
                    inp.value = '';
                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                }
                // 點 × 清除按鈕退出搜尋模式
                const clear = document.querySelector('button[aria-label*="清"], .btn-clear, .search-clear');
                if (clear) clear.click();
                // Blur 搜尋框，讓頁面回到聊天列表
                if (inp) inp.blur();
            }""")
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(400)
        except Exception:
            pass

        # 等聊天室列表渲染（釘選的應該很快出現）
        try:
            await page.wait_for_selector('.list-group-item', timeout=8000)
        except Exception:
            pass

        # 方法1：先在聊天列表直接找（含釘選的群組）— 多試幾次等非同步載入
        items = page.locator(f'.list-group-item:has-text("{chat_name}")')
        for _ in range(5):
            if await items.count() > 0:
                break
            await page.wait_for_timeout(500)
        if await items.count() > 0:
            await items.first.click()
            await page.wait_for_timeout(2000)
            print(f"[line-oa] 在聊天列表找到「{chat_name}」")
            return True

        # 方法2：搜尋
        await page.evaluate("""() => {
            const el = document.querySelector('input[placeholder*="搜尋"], #chatListSearchInput');
            if (el) { el.focus(); el.click(); }
        }""")
        await page.wait_for_timeout(500)
        search = page.locator('input[placeholder*="搜尋"], #chatListSearchInput').first
        await search.fill("")
        await search.fill(chat_name)
        await page.wait_for_timeout(2000)

        result = page.locator(f'.list-group-item:has-text("{chat_name}")').first
        if await result.count() > 0:
            await result.click()
            await page.wait_for_timeout(2000)
            # 清掉搜尋
            try:
                await search.fill("")
            except Exception:
                pass
            return True

        # 清掉搜尋
        try:
            await search.fill("")
            await page.keyboard.press("Escape")
        except Exception:
            pass

        print(f"[line-oa] 找不到聊天室「{chat_name}」")
        return False
    except Exception as e:
        print(f"[line-oa] 開啟聊天室失敗: {e}")
        return False


async def _send_text(page, text: str) -> bool:
    """在已開啟的聊天室發送文字（用剪貼簿貼上支援多行）"""
    try:
        await _dismiss_modals(page)
        ta = page.locator('textarea').first
        if await ta.count() == 0:
            print("[line-oa] 找不到輸入框")
            return False

        await ta.click()
        await ta.fill('')
        await page.wait_for_timeout(200)
        # 用剪貼簿貼上（支援多行文字）
        await page.evaluate('(t) => navigator.clipboard.writeText(t)', text)
        await ta.click()
        await page.keyboard.press('Control+v')
        await page.wait_for_timeout(500)
        # Enter 傳送
        await ta.press('Enter')
        await page.wait_for_timeout(1000)
        return True
    except Exception as e:
        print(f"[line-oa] 發送文字失敗: {e}")
        return False


async def _send_images(page, image_paths: list[str]) -> bool:
    """在已開啟的聊天室一次發送多張圖片"""
    try:
        await _dismiss_modals(page)
        file_input = page.locator('input[type="file"]').first
        if await file_input.count() == 0:
            print("[line-oa] 找不到檔案上傳 input")
            return False

        # 一次上傳所有圖片
        await file_input.set_input_files(image_paths)
        await page.wait_for_timeout(2000)

        # 點傳送按鈕（圖片預覽頁的）
        send_btn = page.locator('button.btn-primary:has-text("傳送")').first
        if await send_btn.count() > 0:
            await send_btn.click()
            await page.wait_for_timeout(3000)
            return True

        print("[line-oa] 找不到圖片傳送按鈕")
        return False
    except Exception as e:
        print(f"[line-oa] 發送圖片失敗: {e}")
        return False


async def send_to_chat(chat_name: str, text: str = "", image_paths: list[str] = None) -> bool:
    """
    發送訊息到指定聊天室（1:1 或群組）。
    先發圖片，再發文字。
    """
    browser, page = await _get_line_oa_page()
    if not page:
        return False

    try:
        if not await _open_chat(page, chat_name):
            return False

        ok = True
        # 先發圖片（一次全部上傳）
        if image_paths:
            if not await _send_images(page, image_paths):
                print(f"[line-oa] 圖片發送失敗")
                ok = False

        # 再發文字
        if text:
            if not await _send_text(page, text):
                ok = False

        # 清空搜尋
        try:
            search = page.locator('input[placeholder*="搜尋"], #chatListSearchInput').first
            await search.fill("")
            await page.wait_for_timeout(500)
        except Exception:
            pass

        return ok
    except Exception as e:
        print(f"[line-oa] send_to_chat 失敗: {e}")
        return False


def send_to_chat_sync(chat_name: str, text: str = "", image_paths: list[str] = None) -> bool:
    """同步版本"""
    return asyncio.run(send_to_chat(chat_name, text, image_paths))


async def send_many_to_chat(chat_name: str, items: list[dict], delay_sec: float = 2.0) -> list[bool]:
    """
    一次開啟聊天室，連續發送多則（text + images）。
    items: [{"text": str, "image_paths": list[str]}, ...]
    回傳每筆是否成功。
    """
    browser, page = await _get_line_oa_page()
    if not page:
        return [False] * len(items)

    results: list[bool] = []
    try:
        if not await _open_chat(page, chat_name):
            return [False] * len(items)

        for idx, it in enumerate(items):
            text = it.get("text", "")
            image_paths = it.get("image_paths") or []
            ok = True
            if image_paths:
                if not await _send_images(page, image_paths):
                    ok = False
            if text:
                if not await _send_text(page, text):
                    ok = False
            results.append(ok)
            if idx < len(items) - 1 and delay_sec > 0:
                await page.wait_for_timeout(int(delay_sec * 1000))
        return results
    except Exception as e:
        print(f"[line-oa] send_many_to_chat 失敗: {e}")
        results.extend([False] * (len(items) - len(results)))
        return results


def send_many_to_chat_sync(chat_name: str, items: list[dict], delay_sec: float = 2.0) -> list[bool]:
    return asyncio.run(send_many_to_chat(chat_name, items, delay_sec))


async def list_chats_matching(keyword: str) -> list[str]:
    """搜尋聊天室名稱包含 keyword 的全部聊天室（回傳名稱清單）。"""
    browser, page = await _get_line_oa_page()
    if not page:
        return []
    try:
        await _dismiss_modals(page)
        await page.evaluate("""() => {
            const el = document.querySelector('input[placeholder*="搜尋"], #chatListSearchInput');
            if (el) { el.focus(); el.click(); }
        }""")
        await page.wait_for_timeout(500)
        search = page.locator('input[placeholder*="搜尋"], #chatListSearchInput').first
        await search.fill("")
        await search.fill(keyword)
        await page.wait_for_timeout(2500)

        items = page.locator(f'.list-group-item:has-text("{keyword}")')
        count = await items.count()
        print(f"[line-oa] 搜尋『{keyword}』.list-group-item 匹配數: {count}")
        names = []
        for i in range(count):
            try:
                raw = (await items.nth(i).inner_text()).strip()
                # 取第一行（群組名）
                for ln in raw.replace("\r", "\n").split("\n"):
                    ln = ln.strip()
                    if ln and keyword in ln and ln not in names:
                        names.append(ln)
                        break
            except Exception:
                continue

        try:
            await search.fill("")
            await page.keyboard.press("Escape")
        except Exception:
            pass
        return names
    except Exception as e:
        print(f"[line-oa] list_chats_matching 失敗: {e}")
        return []


def list_chats_matching_sync(keyword: str) -> list[str]:
    """同步版本"""
    return asyncio.run(list_chats_matching(keyword))


if __name__ == "__main__":
    name = sys.argv[1] if len(sys.argv) > 1 else "賴柏舟"
    msgs = read_chat_sync(name)
    print(f"\n=== {name} 的對話（{len(msgs)} 則）===")
    for m in msgs:
        role = "客服" if m["role"] == "staff" else "客戶"
        print(f"[{role}] {m['text'][:80]}")
