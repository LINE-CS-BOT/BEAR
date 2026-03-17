"""
共用 Chrome CDP 工具
提供：自動啟動 Chrome、自動登入、CDP 連線 + Ecount tab 取得

被 auto_sync_unfulfilled.py 和 sync_cust_from_web.py 共用
"""

import io
import json
import re
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT           = Path(__file__).parent.parent
CHROME_EXE     = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE = str(ROOT / "data" / "chrome_ecount_session")
ERP_URL        = "https://loginib.ecount.com/ec5/view/erp"
LOGIN_HOST     = "login.ecount.com"
CDP_URL        = "http://localhost:9222"
WEB_CONFIG     = ROOT / "data" / "ecount_web_config.json"


# ---------------------------------------------------------------------------
# .env 讀取
# ---------------------------------------------------------------------------

def load_credentials() -> tuple[str, str, str]:
    """回傳 (company_no, user_id, web_password)"""
    env: dict[str, str] = {}
    env_path = ROOT / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                env[k.strip()] = v.strip()
    return (
        env.get("ECOUNT_COMPANY_NO", ""),
        env.get("ECOUNT_USER_ID", ""),
        env.get("ECOUNT_WEB_PASSWORD", ""),
    )


# ---------------------------------------------------------------------------
# config 讀寫
# ---------------------------------------------------------------------------

def load_web_config() -> dict:
    try:
        if WEB_CONFIG.exists():
            return json.loads(WEB_CONFIG.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_web_config(data: dict):
    existing = load_web_config()
    existing.update(data)
    existing["last_updated"] = datetime.now().isoformat()
    WEB_CONFIG.write_text(
        json.dumps(existing, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Chrome 啟動
# ---------------------------------------------------------------------------

def is_port_open(host: str = "127.0.0.1", port: int = 9222, timeout: float = 1.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def launch_chrome_if_needed() -> bool:
    """若 port 9222 未開，自動殺掉既有 Chrome 並重新啟動，回傳是否成功"""
    if is_port_open():
        print("[chrome] CDP 已在運行 (port 9222)")
        return True

    print("[chrome] 啟動 Chrome...", end="", flush=True)
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True)
    time.sleep(2)

    subprocess.Popen(
        [
            CHROME_EXE,
            "--headless=new",
            "--remote-debugging-port=9222",
            f"--user-data-dir={CHROME_PROFILE}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-session-crashed-bubble",
            ERP_URL,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    for i in range(30):
        time.sleep(1)
        if is_port_open():
            print(f" OK ({i + 1}s)")
            time.sleep(3)   # 等頁面初始載入
            return True
        if i % 5 == 4:
            print(f" {i + 1}s...", end="", flush=True)

    print()
    print("[chrome] ✗ Chrome 啟動逾時（30s）")
    return False


# ---------------------------------------------------------------------------
# CDP 連線 + page 取得
# ---------------------------------------------------------------------------

async def connect_get_page(playwright):
    """
    連接 Chrome CDP，找或建立 Ecount tab，設定 1440x900 viewport。
    回傳 (browser, page)
    """
    browser = await playwright.chromium.connect_over_cdp(CDP_URL)
    contexts = browser.contexts
    if not contexts:
        raise RuntimeError("找不到任何 browser context")

    page = None
    for ctx in contexts:
        for pg in ctx.pages:
            if "ecount.com" in pg.url:
                page = pg
                print("[chrome] 找到 Ecount tab")
                break
        if page:
            break

    if not page:
        print("[chrome] 開啟新 Ecount 分頁...")
        page = await contexts[0].new_page()
        await page.goto(ERP_URL, timeout=30000)

    try:
        await page.set_viewport_size({"width": 1440, "height": 900})
        await page.bring_to_front()
    except Exception:
        pass

    return browser, page


# ---------------------------------------------------------------------------
# 自動登入
# ---------------------------------------------------------------------------

async def auto_login(page) -> bool:
    """
    若目前在登入頁（login.ecount.com）則自動填表登入。
    回傳 True = 已登入（或原本就不在登入頁）
    """
    if LOGIN_HOST not in page.url:
        return True

    company, user, password = load_credentials()
    if not password:
        print("[chrome] ✗ .env 未設定 ECOUNT_WEB_PASSWORD")
        return False

    print("[chrome] 偵測到登入頁，自動填表...")
    try:
        await page.fill("#com_code", company, timeout=5000)
        await page.wait_for_timeout(200)
        await page.fill("#id",       user,    timeout=5000)
        await page.wait_for_timeout(200)
        await page.fill("#passwd",   password, timeout=5000)
        await page.wait_for_timeout(200)

        # 記住登入（讓 session 更久）
        try:
            cb = page.locator("#loginck")
            if await cb.count() > 0 and not await cb.is_checked():
                await cb.click(timeout=2000)
        except Exception:
            pass

        await page.click("#save", timeout=5000)

        for attempt in range(30):
            await page.wait_for_timeout(1000)
            if LOGIN_HOST not in page.url:
                print(f"[chrome] ✓ 登入成功（{attempt + 1}s）")
                await page.wait_for_timeout(2000)
                return True
            if attempt % 5 == 4:
                print(f"[chrome]   等待跳轉... {attempt + 1}s")

        print("[chrome] ✗ 登入逾時（密碼可能錯誤）")
        return False

    except Exception as e:
        print(f"[chrome] ✗ 登入失敗: {e}")
        return False


# ---------------------------------------------------------------------------
# 確認已登入 → 回傳 ec_req_sid
# ---------------------------------------------------------------------------

async def ensure_logged_in(page) -> str | None:
    """
    完整登入流程：
      1. 等頁面載入
      2. 若在登入頁 → auto_login
      3. 若無 ec_req_sid → goto ERP_URL → 再試一次
    回傳 ec_req_sid 或 None（失敗）
    """
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass

    if LOGIN_HOST in page.url:
        if not await auto_login(page):
            return None

    m = re.search(r"ec_req_sid=([^&#]+)", page.url)
    if not m:
        print("[chrome] 導向 ERP 主頁取得 session...")
        await page.goto(ERP_URL, timeout=30000)
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        await page.wait_for_timeout(2000)

        if LOGIN_HOST in page.url:
            if not await auto_login(page):
                return None

        m = re.search(r"ec_req_sid=([^&#]+)", page.url)

    if not m:
        print("[chrome] ✗ 無法取得 ec_req_sid（登入失敗）")
        return None

    sid = m.group(1)
    print(f"[chrome] ✓ 已登入 (sid: {sid[:10]}...)")
    return sid
