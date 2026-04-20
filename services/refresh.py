"""
資料庫自動刷新服務

刷新條件（任一成立就立即跑腳本，無冷卻限制）：
  specs.json        → 來源 產品PO文.txt 比 output 新，或 output 不存在
  image_hashes.json → 圖片庫中有比 output 更新的圖片，或 output 不存在

時間限制：
  晚上 23:00 ~ 早上 11:00 → 跳過刷新（非營業時間）

呼叫位置：
  main.py lifespan 啟動時執行一次
  _refresh_data_loop 每 2 小時背景檢查一次
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


def _suppress_windows_drive_dialogs():
    """
    防止 Windows 在存取無法連線的磁碟機時彈出錯誤對話框。
    只在 Windows 上有效。
    SEM_FAILCRITICALERRORS (0x0001) + SEM_NOOPENFILEERRORBOX (0x8000)
    """
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass

# ── 路徑設定 ─────────────────────────────────────────────
PYTHON        = r"C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe"
_PROJECT_ROOT = Path(__file__).parent.parent

SPECS_JSON     = _PROJECT_ROOT / "data" / "specs.json"
HASHES_JSON    = _PROJECT_ROOT / "data" / "image_hashes.json"
AVAILABLE_JSON = _PROJECT_ROOT / "data" / "available.json"
SPECS_SOURCE   = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt")
IMAGE_DIR      = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品照片")

# available.json 超過幾秒就重新同步（預設 2 小時）
AVAILABLE_MAX_AGE = 2 * 3600

# 刷新時段：11:00 ~ 23:00（其他時間跳過）
REFRESH_HOUR_START = 11
REFRESH_HOUR_END   = 23


# ── 工具函數 ──────────────────────────────────────────────
def _in_active_hours() -> bool:
    """是否在允許刷新的時段（11:00 ~ 23:00）"""
    h = datetime.now().hour
    return REFRESH_HOUR_START <= h < REFRESH_HOUR_END


def _mtime(path: Path) -> float:
    """取得檔案修改時間，檔案不存在回傳 0"""
    try:
        return path.stat().st_mtime if path.exists() else 0.0
    except Exception:
        return 0.0


def _newest_image_mtime() -> float:
    """圖片庫中最新圖片的修改時間"""
    if not IMAGE_DIR.exists():
        return 0.0
    exts = {".jpg", ".jpeg", ".png", ".webp"}
    mtimes = [
        f.stat().st_mtime
        for f in IMAGE_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in exts
    ]
    return max(mtimes) if mtimes else 0.0


def _run_script(script_rel: str) -> bool:
    """執行腳本，回傳是否成功"""
    script = _PROJECT_ROOT / script_rel
    try:
        result = subprocess.run(
            [PYTHON, str(script)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(_PROJECT_ROOT),
            timeout=120,
        )
        if result.returncode == 0:
            print(f"[refresh] OK {script_rel} 完成")
            return True
        else:
            stderr_text = result.stderr.decode("utf-8", errors="replace")[-500:] if result.stderr else ""
            print(f"[refresh] FAIL {script_rel} 失敗: {stderr_text}")
            return False
    except subprocess.TimeoutExpired:
        print(f"[refresh] TIMEOUT {script_rel} 超時（120s）")
        return False
    except Exception as e:
        print(f"[refresh] ERROR {script_rel} 執行錯誤: {e}")
        return False


# ── 主要檢查函數 ──────────────────────────────────────────
def check_and_refresh():
    """
    檢查兩個資料庫是否需要刷新，條件成立就立即執行。
    晚上 23:00 ~ 早上 11:00 之間跳過。
    """
    _suppress_windows_drive_dialogs()   # 避免 Windows 彈出磁碟機錯誤對話框

    if not _in_active_hours():
        h = datetime.now().hour
        print(f"[refresh] 非營業時段（{h:02d}:xx），跳過刷新")
        return

    # ── 規格 DB（specs.json）──────────────────────────────
    specs_out = _mtime(SPECS_JSON)
    specs_src = _mtime(SPECS_SOURCE)

    if specs_out == 0 or specs_src > specs_out:
        reason = "output 不存在" if specs_out == 0 else f"來源新了 {int((specs_src - specs_out)/60)} 分"
        print(f"[refresh] 規格DB 需更新（{reason}），執行 import_specs.py ...")
        if _run_script("scripts/import_specs.py"):
            try:
                import storage.specs as spec_store
                spec_store.reload()
            except Exception as e:
                print(f"[refresh] specs 快取重載失敗: {e}")
    else:
        age = int((time.time() - specs_out) / 60)
        print(f"[refresh] 規格DB 無需刷新（{age} 分前更新）")

    # ── 圖片雜湊 DB（image_hashes.json）──────────────────
    hashes_out = _mtime(HASHES_JSON)
    newest_img = _newest_image_mtime()

    if hashes_out == 0 or newest_img > hashes_out:
        reason = "output 不存在" if hashes_out == 0 else f"有新圖片（+{int((newest_img - hashes_out)/60)} 分）"
        print(f"[refresh] 圖片DB 需更新（{reason}），執行 build_image_hashes.py ...")
        _run_script("scripts/build_image_hashes.py")
    else:
        age = int((time.time() - hashes_out) / 60)
        print(f"[refresh] 圖片DB 無需刷新（{age} 分前更新）")

    # ── 可售庫存（available.json）──────────────────────────────────────────
    avail_out = _mtime(AVAILABLE_JSON)
    avail_age = time.time() - avail_out if avail_out > 0 else float("inf")

    if avail_age > AVAILABLE_MAX_AGE:
        reason = "output 不存在" if avail_out == 0 else f"已 {int(avail_age/60)} 分未更新"
        print(f"[refresh] 可售庫存 需更新（{reason}），執行 auto_sync_unfulfilled.py ...")
        _run_script("scripts/auto_sync_unfulfilled.py")
    else:
        print(f"[refresh] 可售庫存 無需刷新（{int(avail_age/60)} 分前更新）")


def trigger_rebuild(callback=None):
    """
    立即重建 specs.json + image_hashes.json（背景執行緒，不阻塞主線程）。
    callback: 重建完成後呼叫的函式（無參數），用於 label 生成等後續動作。
    """
    import threading

    def _rebuild():
        print("[refresh] 觸發立即重建 specs + image_hashes ...")
        _suppress_windows_drive_dialogs()
        if _run_script("scripts/import_specs.py"):
            try:
                import storage.specs as spec_store
                spec_store.reload()
            except Exception as e:
                print(f"[refresh] specs 快取重載失敗: {e}")
        _run_script("scripts/build_image_hashes.py")
        print("[refresh] 立即重建完成")
        if callback:
            try:
                callback()
            except Exception as _e:
                print(f"[refresh] callback 執行失敗: {_e}")

    t = threading.Thread(target=_rebuild, daemon=True)
    t.start()
