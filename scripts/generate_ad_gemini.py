"""
Playwright 自動化：開啟 Gemini → 上傳產品圖 → 貼提示詞 → 下載廣告圖

用法（由 handlers/ad_maker.py 呼叫）：
  python scripts/generate_ad_gemini.py --payload '{"codes":[...],...}'

直接執行測試：
  python scripts/generate_ad_gemini.py --test Z3432
"""

import argparse
import asyncio
import base64
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── 同步寫入 log 檔案 ─────────────────────────────────────────────────────────
import io
_LOG_PATH = ROOT / "data" / "ad_gemini.log"
_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

class _Tee:
    """同時輸出到 stdout 和 log 檔"""
    def __init__(self, stream, log_path):
        self._stream  = stream
        self._logfile = open(log_path, "a", encoding="utf-8", errors="replace")
    def write(self, data):
        self._stream.write(data)
        self._logfile.write(data)
        return len(data)
    def flush(self):
        self._stream.flush()
        self._logfile.flush()
    @property
    def closed(self):
        return self._stream.closed
    def __getattr__(self, name):
        return getattr(self._stream, name)

sys.stdout = _Tee(sys.stdout, _LOG_PATH)
sys.stderr = _Tee(sys.stderr, _LOG_PATH)

GEMINI_URL     = "https://gemini.google.com/app"
GEMINI_PROFILE = str(ROOT / "data" / "gemini_chrome_session")
PLATFORMS      = ["line", "fb"]

# ── LINE API 推送通知 ─────────────────────────────────────────────────────────

def _push_notify(group_id: str, text: str) -> None:
    """不推送通知（節省 push 額度）"""
    return
    try:
        import os
        os.chdir(str(ROOT))
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass
    try:
        from linebot.v3 import WebhookHandler
        from linebot.v3.messaging import (
            Configuration, ApiClient, MessagingApi,
            PushMessageRequest, TextMessage,
        )
        from config import settings
        cfg = Configuration(access_token=settings.LINE_CHANNEL_ACCESS_TOKEN)
        with ApiClient(cfg) as api:
            MessagingApi(api).push_message(PushMessageRequest(
                to=_ADMIN_UID,
                messages=[TextMessage(text=text)],
            ))
    except Exception as e:
        print(f"[notify] 推送失敗：{e}")


# ── 單一產品 × 單一平台生成 ───────────────────────────────────────────────────

async def generate_one(page, prod_code: str, image_paths: list[str],
                       prompt: str, output_folder: str,
                       platform: str) -> str | None:
    """
    對單一產品+平台執行 Gemini 生成。
    image_paths：產品圖路徑列表（可多張）。
    回傳儲存路徑，失敗回傳 None。
    """
    out_path = Path(output_folder) / f"{prod_code}_{platform}.png"
    print(f"[gemini] 開始生成 {prod_code}_{platform} ...", flush=True)

    try:
        # 1. 確認在 Gemini 頁面
        if "gemini.google.com" not in page.url:
            await page.goto(GEMINI_URL, timeout=30000)
            await page.wait_for_load_state("domcontentloaded", timeout=20000)
            await page.wait_for_timeout(2000)

        # 2. 找輸入框（多個選擇器備援）
        input_sels = [
            'rich-textarea div[contenteditable="true"]',
            'div[aria-label="Enter a prompt here"]',
            'textarea[placeholder]',
            '.ql-editor',
        ]
        input_box = None
        for sel in input_sels:
            loc = page.locator(sel)
            if await loc.count() > 0:
                input_box = loc.first
                break
        if not input_box:
            print(f"[gemini] ❌ 找不到輸入框，可能需要重新登入 Gemini")
            return None

        # 3. 上傳產品圖（全部張數）
        # 若圖片在網路磁碟機（如 H:\），先複製到本機 temp 再上傳
        import shutil, tempfile
        local_imgs = []
        for img in image_paths:
            p = Path(img)
            if not p.exists():
                continue
            if p.drive.upper() not in ("C:", "D:", "E:"):
                try:
                    tmp = Path(tempfile.gettempdir()) / p.name
                    shutil.copy2(str(p), str(tmp))
                    local_imgs.append(str(tmp))
                    print(f"[gemini] 圖片已複製到本機：{tmp.name}")
                except Exception as _ce:
                    print(f"[gemini] 複製圖片失敗（{_ce}），使用原路徑")
                    local_imgs.append(img)
            else:
                local_imgs.append(img)

        if local_imgs:
            try:
                attached = False
                names = ", ".join(Path(p).name for p in local_imgs)

                # Gemini 兩步驟：① 點 + 按鈕 → ② 點「上傳檔案」→ file chooser
                plus_sels = [
                    'button[aria-controls="upload-file-menu"]',
                    'button.upload-card-button',
                    'button[aria-label="新增檔案"]',
                    'button[aria-label="Add files"]',
                ]
                for sel in plus_sels:
                    btn = page.locator(sel)
                    if await btn.count() > 0:
                        await btn.first.click()
                        await page.wait_for_timeout(800)
                        upload_item_sels = [
                            'button[role="menuitem"]:has-text("上傳檔案")',
                            '[role="menuitem"]:has-text("上傳檔案")',
                            '[role="menuitem"]:has-text("Upload")',
                        ]
                        for usel in upload_item_sels:
                            uitem = page.locator(usel)
                            if await uitem.count() > 0:
                                try:
                                    async with page.expect_file_chooser(timeout=5000) as fc_info:
                                        await uitem.first.click()
                                    fc = await fc_info.value
                                    await fc.set_files(local_imgs)  # 一次上傳全部
                                    attached = True
                                    print(f"[gemini] ✅ 已上傳 {len(local_imgs)} 張圖片：{names}")
                                    await page.wait_for_timeout(2000)
                                except Exception as ue:
                                    print(f"[gemini] 上傳選單失敗（{ue}）")
                                break
                        if attached:
                            break

                # 備援：直接找 file input
                if not attached:
                    fi = page.locator('input[type="file"]')
                    if await fi.count() > 0:
                        await fi.first.set_input_files(local_imgs)
                        attached = True
                        print(f"[gemini] ✅ 已上傳圖片（備援）：{names}")
                        await page.wait_for_timeout(2000)
                if not attached:
                    print(f"[gemini] ⚠️  找不到圖片上傳按鈕，僅使用文字提示詞")
            except Exception as e:
                print(f"[gemini] ⚠️  圖片上傳失敗（{e}），繼續用文字提示詞")

        # 4. 清空並輸入提示詞
        await input_box.click()
        await page.keyboard.press("Control+a")
        await page.keyboard.press("Delete")
        await page.wait_for_timeout(300)
        # 用 JavaScript 填入避免中文輸入問題（直接在元素上 evaluate）
        await input_box.evaluate(
            "(el, txt) => { el.innerText = txt; "
            "el.dispatchEvent(new Event('input', {bubbles:true})); }",
            prompt
        )
        await page.wait_for_timeout(500)

        # 5. 送出（Enter 或按鈕）
        send_sels = [
            'button[aria-label="Send message"]',
            'button[aria-label="傳送"]',
            'button[aria-label="傳送訊息"]',
            'button[data-test-id="send-button"]',
            'button.send-button',
        ]
        sent = False
        for sel in send_sels:
            btn = page.locator(sel)
            if await btn.count() > 0:
                try:
                    await btn.first.click(timeout=3000)
                    sent = True
                    break
                except Exception:
                    pass
        if not sent:
            await page.keyboard.press("Enter")

        print(f"[gemini] 提示詞已送出，等待生成...", flush=True)

        # ── 送出前先記錄現有圖片 src，之後只抓「新出現的」 ──────────────────
        _img_sels = [
            'model-response img', '.response-container img',
            '.message-content img', 'chat-message:last-child img',
            '[data-message-id] img', 'article img',
        ]
        _valid_prefixes = ("data:image", "blob:", "https://lh3",
                           "https://generativelanguage", "https://gg/")
        existing_srcs: set[str] = set()
        for sel in _img_sels:
            loc = page.locator(sel)
            cnt = await loc.count()
            for i in range(cnt):
                src = await loc.nth(i).get_attribute("src") or ""
                if src:
                    existing_srcs.add(src)
        print(f"[gemini] 現有圖片 {len(existing_srcs)} 張，等待新圖...", flush=True)

        await page.wait_for_timeout(4000)

        # 6. 等待「新」圖片出現（最多 120 秒）
        generated_src = None
        for attempt in range(40):  # 40 × 3s = 120s
            await page.wait_for_timeout(3000)

            for sel in _img_sels:
                imgs = page.locator(sel)
                cnt  = await imgs.count()
                for i in range(cnt - 1, -1, -1):
                    src = await imgs.nth(i).get_attribute("src") or ""
                    if (src and src not in existing_srcs
                            and any(src.startswith(p) for p in _valid_prefixes)):
                        generated_src = src
                        print(f"[gemini] ✅ 偵測到新生成圖片（attempt {attempt+1}）")
                        break
                if generated_src:
                    break
            if generated_src:
                break

            if attempt % 5 == 4:
                print(f"[gemini] ... 等待中 {(attempt+1)*3}s", flush=True)

        if not generated_src:
            print(f"[gemini] ❌ 等待逾時（120s），未偵測到生成圖片")
            return None

        print(f"[gemini] 圖片 src 類型：{generated_src[:60]}", flush=True)

        # 7. 下載 / 存圖
        saved = False

        # 方法 A：點圖片 → 右鍵另存新檔
        try:
            img_loc = None
            for sel in ['model-response img', '.response-container img', 'article img']:
                loc = page.locator(sel)
                cnt = await loc.count()
                for i in range(cnt - 1, -1, -1):
                    src = await loc.nth(i).get_attribute("src") or ""
                    if src == generated_src:
                        img_loc = loc.nth(i)
                        break
                if img_loc:
                    break
            if not img_loc:
                for sel in ['model-response img', '.response-container img', 'article img']:
                    loc = page.locator(sel)
                    if await loc.count() > 0:
                        img_loc = loc.last
                        break
            if img_loc:
                # 點圖片打開大圖
                await img_loc.click()
                await page.wait_for_timeout(1000)
                # 右鍵另存新檔
                try:
                    # 找大圖的 img 元素
                    big_img = page.locator('img[style*="max-width"], img[style*="max-height"], .lightbox img, [role="dialog"] img').last
                    if await big_img.count() == 0:
                        big_img = img_loc
                    async with page.expect_download(timeout=20000) as dl_info:
                        await big_img.click(button="right")
                        await page.wait_for_timeout(500)
                        # 點「另存圖片」選項
                        save_sels = [
                            'text="另存圖片為..."',
                            'text="Save image as..."',
                            'text="另存圖片"',
                        ]
                        for ss in save_sels:
                            sl = page.locator(ss)
                            if await sl.count() > 0:
                                await sl.first.click()
                                break
                    dl = await dl_info.value
                    await dl.save_as(str(out_path))
                    saved = True
                    print(f"[gemini] ✅ 方法A（右鍵另存）成功")
                except Exception as e:
                    print(f"[gemini] 方法A 右鍵另存失敗（{e}），嘗試其他方法")
                    # 按 Esc 關閉可能的大圖彈窗
                    await page.keyboard.press("Escape")
                    await page.wait_for_timeout(500)
        except Exception as e:
            print(f"[gemini] 方法A 失敗（{e}）")

        # 方法 B：data URI → 直接解碼
        if not saved and generated_src.startswith("data:image"):
            try:
                _, b64data = generated_src.split(",", 1)
                out_path.write_bytes(base64.b64decode(b64data))
                saved = True
                print(f"[gemini] ✅ 方法B（data URI）成功")
            except Exception as e:
                print(f"[gemini] data URI 解碼失敗：{e}")

        # 方法 C：blob URL → 透過 JS 轉 base64
        if not saved and generated_src.startswith("blob:"):
            try:
                b64 = await page.evaluate("""async (blobUrl) => {
                    const resp = await fetch(blobUrl);
                    const buf  = await resp.arrayBuffer();
                    const arr  = new Uint8Array(buf);
                    let bin = '';
                    arr.forEach(b => bin += String.fromCharCode(b));
                    return btoa(bin);
                }""", generated_src)
                out_path.write_bytes(base64.b64decode(b64))
                saved = True
                print(f"[gemini] ✅ 方法B（blob→base64）成功")
            except Exception as e:
                print(f"[gemini] blob 轉換失敗：{e}")

        # 方法 D：從 browser context 取 cookies，用 httpx 帶 cookie 下載
        if not saved and generated_src.startswith("https://"):
            try:
                import httpx
                cookies_list = await page.context.cookies()
                cookie_dict  = {c["name"]: c["value"] for c in cookies_list}
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Referer":    "https://gemini.google.com/",
                }
                resp = httpx.get(generated_src, cookies=cookie_dict,
                                 headers=headers, timeout=30, follow_redirects=True)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    out_path.write_bytes(resp.content)
                    saved = True
                    print(f"[gemini] ✅ 方法D（帶Cookie下載）成功，{len(resp.content)//1024}KB")
                else:
                    print(f"[gemini] 方法D 失敗（HTTP {resp.status_code}，{len(resp.content)} bytes）")
            except Exception as e:
                print(f"[gemini] 方法D 失敗（{e}）")

        # 方法 C：找生成圖片元素，截取該元素的截圖
        if not saved:
            try:
                img_sels_save = [
                    'model-response img', '.response-container img',
                    'article img', 'chat-message:last-child img',
                ]
                for sel in img_sels_save:
                    imgs = page.locator(sel)
                    cnt = await imgs.count()
                    for idx in range(cnt - 1, -1, -1):
                        src = await imgs.nth(idx).get_attribute("src") or ""
                        if src.startswith(("data:image", "blob:", "https://")):
                            await imgs.nth(idx).screenshot(path=str(out_path))
                            saved = True
                            print(f"[gemini] ✅ 方法C（element screenshot）成功")
                            break
                    if saved:
                        break
            except Exception as e:
                print(f"[gemini] element screenshot 失敗：{e}")


        # 方法 E：JS Canvas → base64（避免 CORS 問題）
        if not saved:
            try:
                b64 = await page.evaluate("""async () => {
                    const sels = ['model-response img', '.response-container img',
                                  'article img', 'chat-message:last-child img'];
                    let imgEl = null;
                    for (const sel of sels) {
                        const els = document.querySelectorAll(sel);
                        for (let i = els.length - 1; i >= 0; i--) {
                            if (els[i].src && els[i].naturalWidth > 0) {
                                imgEl = els[i]; break;
                            }
                        }
                        if (imgEl) break;
                    }
                    if (!imgEl) return null;
                    const canvas = document.createElement('canvas');
                    canvas.width  = imgEl.naturalWidth  || imgEl.width;
                    canvas.height = imgEl.naturalHeight || imgEl.height;
                    canvas.getContext('2d').drawImage(imgEl, 0, 0);
                    return canvas.toDataURL('image/png').split(',')[1];
                }""")
                if b64:
                    out_path.write_bytes(base64.b64decode(b64))
                    saved = True
                    print(f"[gemini] ✅ 方法E（JS Canvas）成功")
            except Exception as e:
                print(f"[gemini] 方法E JS Canvas 失敗：{e}")

        if saved:
            print(f"[gemini] ✅ 已儲存：{out_path}", flush=True)
            return str(out_path)
        else:
            print(f"[gemini] ❌ 無法儲存圖片")
            return None

    except Exception as e:
        import traceback
        print(f"[gemini] ❌ 生成失敗 {prod_code}/{platform}: {e}")
        traceback.print_exc()
        return None


# ── 主流程 ────────────────────────────────────────────────────────────────────

async def main(payload: dict) -> None:
    from playwright.async_api import async_playwright

    codes         = payload["codes"]
    images        = payload["images"]       # {code: [path, ...]}
    output_folder = payload["output_folder"]
    notify_group  = payload.get("notify_group", "")

    Path(output_folder).mkdir(parents=True, exist_ok=True)

    # 不去背，直接用原始素材圖
    import tempfile

    # 預先計算所有提示詞
    from handlers.ad_maker import build_gemini_prompt
    PROMPT_DIR = ROOT / "data" / "ad_prompts"
    tasks = []
    for code in codes:
        img_list = images.get(code, [])
        for platform in PLATFORMS:
            txt = PROMPT_DIR / f"{code}_{platform}.txt"
            if txt.exists():
                prompt = txt.read_text(encoding="utf-8").strip()
                print(f"[prompt] {code}/{platform} ← 讀取預存提示詞（{len(prompt)} 字）")
            else:
                prompt = build_gemini_prompt(code, platform)
                print(f"[prompt] {code}/{platform} ← 使用預設模板")
            # 提醒 Gemini 保留原圖
            orig_note = (
                "\n\n⚠️ 重要：我上傳的是原始產品照片，"
                "必須完整保留商品盒型外觀，四個角清楚可見，"
                "不可裁切、變形、更改盒子上的圖案和文字。"
            )
            tasks.append((code, img_list, prompt + orig_note, platform))

    total   = len(tasks)
    success = 0
    failed  = []

    print(f"[ad] 共 {total} 張待生成（{len(codes)} 產品 × 2 平台）", flush=True)

    async with async_playwright() as p:
        # 使用獨立的 Gemini Chrome Profile（有頭模式，需手動登入過一次）
        context = await p.chromium.launch_persistent_context(
            user_data_dir=GEMINI_PROFILE,
            headless=False,
            channel="chrome",
            viewport={"width": 1280, "height": 900},
            args=[
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--excludeSwitches=enable-automation",
            ],
            ignore_default_args=["--enable-automation"],
        )

        page = context.pages[0] if context.pages else await context.new_page()

        # 確認 Gemini 可以存取
        await page.goto(GEMINI_URL, timeout=30000)
        await page.wait_for_load_state("domcontentloaded", timeout=20000)
        await page.wait_for_timeout(3000)

        # 若在登入頁，等使用者手動登入（最多 3 分鐘）
        if "accounts.google.com" in page.url or "signin" in page.url or "gemini.google.com" not in page.url:
            print("[gemini] ⚠️  請在瀏覽器完成 Google 登入（等待最多 3 分鐘）...")
            for i in range(90):  # 90 × 2s = 180s
                await page.wait_for_timeout(2000)
                cur = page.url
                if "gemini.google.com" in cur:
                    print("[gemini] ✅ 登入成功，進入 Gemini")
                    await page.wait_for_timeout(3000)
                    break
                if i % 15 == 14:
                    elapsed = (i + 1) * 2
                    print(f"[gemini] ... 等待登入中 {elapsed}s / 180s", flush=True)
            else:
                print("[gemini] ❌ 登入逾時（180s），結束")
                await context.close()
                return

        # 依序生成（每次開新對話）
        for i, (code, img_list, prompt, platform) in enumerate(tasks):
            print(f"\n[ad] [{i+1}/{total}] {code}_{platform}", flush=True)

            # 開新對話（直接導航到 Gemini 首頁）
            try:
                await page.goto(GEMINI_URL, timeout=20000)
                await page.wait_for_load_state("domcontentloaded", timeout=15000)
                await page.wait_for_timeout(2000)
                print("[gemini] ✅ 已開新對話")

                # 點「建立圖像」按鈕
                _img_btn_sels = [
                    'button:has-text("建立圖像")',
                    'button:has-text("Create image")',
                    'button:has-text("Generate image")',
                    '[data-test-id="image-generation"]',
                ]
                for _ib_sel in _img_btn_sels:
                    _ib = page.locator(_ib_sel)
                    if await _ib.count() > 0:
                        await _ib.first.click()
                        await page.wait_for_timeout(1500)
                        print("[gemini] ✅ 已點選「建立圖像」")
                        break
            except Exception as e:
                print(f"[gemini] ⚠️  開新對話失敗（{e}），使用目前頁面")

            result = await generate_one(
                page, code, img_list, prompt, output_folder, platform
            )
            if result:
                success += 1
            else:
                failed.append(f"{code}_{platform}")

            # 每張圖之間稍作停頓
            await page.wait_for_timeout(2000)

        await context.close()

    # 通知結果
    ts = datetime.now().strftime("%m/%d %H:%M")
    if failed:
        fail_list = "\n".join(f"  • {f}" for f in failed)
        msg = (
            f"🖼️ 廣告圖生成完成（{ts}）\n"
            f"✅ 成功 {success} 張 / ❌ 失敗 {len(failed)} 張\n\n"
            f"失敗項目：\n{fail_list}\n\n"
            f"存檔位置：\n{output_folder}"
        )
    else:
        msg = (
            f"🖼️ 廣告圖生成完成（{ts}）\n"
            f"✅ 全部成功，共 {success} 張\n\n"
            f"存檔位置：\n{output_folder}"
        )

    print(f"\n{msg}", flush=True)
    _push_notify(notify_group, msg)


# ── 入口 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--payload", type=str, help="JSON payload")
    parser.add_argument("--test",    type=str, help="測試用貨號，例：Z3432")
    args = parser.parse_args()

    if args.test:
        # 快速測試模式
        from handlers.ad_maker import AD_IMAGE_DIR, AD_OUTPUT_DIR
        code     = args.test.upper()
        imgs     = [str(f) for f in AD_IMAGE_DIR.glob("*")
                    if f.stem.upper().startswith(code)]
        payload  = {
            "codes":         [code],
            "images":        {code: imgs},
            "output_folder": str(AD_OUTPUT_DIR),
            "notify_group":  "",
        }
    elif args.payload:
        payload = json.loads(args.payload)
    else:
        print("用法：--payload JSON 或 --test 貨號")
        sys.exit(1)

    asyncio.run(main(payload))
