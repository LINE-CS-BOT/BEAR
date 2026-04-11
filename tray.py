"""
LINE Bot 系統匣程式
- 右下角常駐圖示
- 右鍵選單：開啟 Admin、查看 ngrok URL、停止
- 重複啟動防護（同時只允許一個 instance）
"""

import subprocess
import sys
import os
import webbrowser
import time
import threading
import urllib.request
import json

import pystray
from PIL import Image, ImageDraw

# ── 路徑設定 ──────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
# 固定用 python.exe（不用 pythonw.exe），確保 uvicorn --reload 正常運作
PYTHON      = os.path.join(os.path.dirname(sys.executable), "python.exe")
if not os.path.exists(PYTHON):          # fallback
    PYTHON  = sys.executable
LOCK_FILE   = os.path.join(BASE_DIR, "data", "tray.lock")

os.environ["WATCHFILES_FORCE_POLLING"] = "true"
UVICORN_CMD = [PYTHON, "main.py"]

CADDY_EXE   = os.path.join(BASE_DIR, "caddy.exe")
CADDY_CMD   = [CADDY_EXE, "run", "--config", os.path.join(BASE_DIR, "Caddyfile")]
WEBHOOK_URL = "https://xmnline.duckdns.org/webhook"

# ── 重複啟動防護 ──────────────────────────────────────
_lock_handle = None

def _acquire_lock() -> bool:
    """回傳 True 表示取得鎖（第一個 instance）；False 表示已有 instance 在跑"""
    global _lock_handle
    os.makedirs(os.path.dirname(LOCK_FILE), exist_ok=True)
    try:
        import msvcrt
        _lock_handle = open(LOCK_FILE, "w")
        msvcrt.locking(_lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
        _lock_handle.write(str(os.getpid()))
        _lock_handle.flush()
        return True
    except (OSError, IOError):
        return False

def _release_lock():
    global _lock_handle
    if _lock_handle:
        try:
            import msvcrt
            _lock_handle.seek(0)
            msvcrt.locking(_lock_handle.fileno(), msvcrt.LK_UNLCK, 1)
            _lock_handle.close()
        except Exception:
            pass
        try:
            os.remove(LOCK_FILE)
        except Exception:
            pass

# ── 子程序 ────────────────────────────────────────────
_procs: dict[str, subprocess.Popen] = {}
CREATE_NO_WINDOW = 0x08000000

def _start_uvicorn():
    _kill_port_8000()   # 啟動前先清掉殘留 process
    _wait_port_free(8000, timeout=10)  # 等 port 真正釋放
    # 不重導 stdout，避免 uvicorn reload 時 file handle 關閉導致 crash
    p = subprocess.Popen(
        UVICORN_CMD,
        cwd=BASE_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW,
    )
    _procs["uvicorn"] = p

def _start_caddy():
    """啟動 Caddy 反向代理（自動 HTTPS，取代 ngrok）"""
    p = subprocess.Popen(
        CADDY_CMD,
        cwd=BASE_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW,
    )
    _procs["caddy"] = p

def _kill_port_8000():
    """強制釋放 port 8000（殺掉佔用的 process tree）"""
    try:
        import subprocess as _sp
        result = _sp.run(
            ["netstat", "-ano"],
            capture_output=True, text=True,
            creationflags=CREATE_NO_WINDOW,
        )
        killed = set()
        for line in result.stdout.splitlines():
            if ":8000 " in line and "LISTENING" in line:
                pid = line.strip().split()[-1]
                if pid.isdigit() and pid not in killed:
                    killed.add(pid)
                    # /T 殺掉整個 process tree
                    _sp.run(
                        ["taskkill", "/F", "/T", "/PID", pid],
                        creationflags=CREATE_NO_WINDOW,
                        capture_output=True,
                        timeout=5,
                    )
    except Exception:
        pass


def _wait_port_free(port=8000, timeout=10):
    """等待 port 完全釋放，最多等 timeout 秒"""
    import subprocess as _sp
    for _ in range(timeout * 2):
        result = _sp.run(
            ["netstat", "-ano"],
            capture_output=True, text=True,
            creationflags=CREATE_NO_WINDOW,
        )
        if not any(f":{port} " in l and "LISTENING" in l for l in result.stdout.splitlines()):
            return True
        time.sleep(0.5)
    return False


_watchdog_running = False

def _watchdog_loop():
    """背景執行緒：每 5 秒檢查 uvicorn + caddy，若死掉就自動重啟"""
    global _watchdog_running
    while _watchdog_running:
        time.sleep(5)
        if not _watchdog_running:
            break
        p = _procs.get("uvicorn")
        if p is not None and p.poll() is not None:
            _kill_port_8000()
            _wait_port_free(8000, timeout=10)
            _start_uvicorn()
            # 等待 server 健康
            for _ in range(20):
                time.sleep(1)
                try:
                    r = urllib.request.urlopen("http://localhost:8000/health", timeout=2)
                    if r.status == 200:
                        break
                except Exception:
                    pass
        c = _procs.get("caddy")
        if c is not None and c.poll() is not None:
            time.sleep(1)
            _start_caddy()

def _stop_all():
    global _watchdog_running
    _watchdog_running = False
    for name, p in list(_procs.items()):
        try:
            p.terminate()
            p.wait(timeout=3)
        except Exception:
            pass
    _procs.clear()
    # 強制釋放 port 8000（防殭屍 process）
    _kill_port_8000()
    # 也確保 caddy 停止
    subprocess.Popen(
        "taskkill /F /IM caddy.exe >nul 2>&1",
        shell=True,
        creationflags=CREATE_NO_WINDOW,
    )

# ── 圖示（綠色圓點）──────────────────────────────────
def _make_icon() -> Image.Image:
    size = 64
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([2,  2,  size-2,  size-2],  fill=(50, 50, 50))
    draw.ellipse([12, 12, size-12, size-12], fill=(0, 200, 80))
    return img

# ── 選單動作 ──────────────────────────────────────────
def action_open_admin(icon, item):
    webbrowser.open("http://localhost:8000/admin")

def action_show_webhook(icon, item):
    subprocess.run(
        ["clip"],
        input=WEBHOOK_URL.encode("utf-8"),
        creationflags=CREATE_NO_WINDOW,
    )
    icon.notify(f"Webhook URL:\n{WEBHOOK_URL}\n（已複製到剪貼簿）", "LINE Bot")

def action_restart_server(icon, item):
    if "uvicorn" in _procs:
        try:
            _procs["uvicorn"].terminate()
            _procs["uvicorn"].wait(timeout=5)
        except Exception:
            pass
        del _procs["uvicorn"]
    # 強制釋放 port 8000，確保舊 process 已死（shell=True 時 terminate 只殺 cmd.exe）
    _kill_port_8000()
    if not _wait_port_free(8000, timeout=15):
        # 第二輪嘗試
        _kill_port_8000()
        _wait_port_free(8000, timeout=10)
    _start_uvicorn()
    # 等待 server 健康
    for _ in range(20):
        time.sleep(1)
        try:
            r = urllib.request.urlopen("http://localhost:8000/health", timeout=2)
            if r.status == 200:
                icon.notify("Server 已重啟", "LINE Bot")
                return
        except Exception:
            pass
    icon.notify("Server 重啟中，請稍候...", "LINE Bot")

def action_quit(icon, item):
    _stop_all()
    _release_lock()
    icon.stop()

# ── 主程式 ────────────────────────────────────────────
def main():
    if not _acquire_lock():
        # 已有 instance，靜默退出
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            "LINE Bot 已經在執行中！\n（右下角查看綠色圖示）",
            "LINE Bot",
            0x30  # MB_ICONWARNING
        )
        sys.exit(0)

    os.chdir(BASE_DIR)
    _start_uvicorn()
    _start_caddy()

    # ── 啟動 watchdog 背景執行緒 ──
    global _watchdog_running
    _watchdog_running = True
    t = threading.Thread(target=_watchdog_loop, daemon=True)
    t.start()

    menu = pystray.Menu(
        pystray.MenuItem("🌐 開啟 Admin 介面", action_open_admin),
        pystray.MenuItem("🔗 複製 Webhook URL", action_show_webhook),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("🔄 重啟 Server",      action_restart_server),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("❌ 停止並離開",        action_quit),
    )

    icon = pystray.Icon(
        name="line-bot",
        icon=_make_icon(),
        title="LINE Bot Server",
        menu=menu,
    )
    icon.run()

if __name__ == "__main__":
    main()
