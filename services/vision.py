"""
圖片辨識服務

支援兩種識別方式：
1. 感知雜湊（pHash）比對：與本地圖片庫比對，找出相似產品
2. OCR 文字擷取：Tesseract，從圖片讀取產品編號 / 品名

圖片庫：H:\其他電腦\我的電腦\小蠻牛\產品照片\
雜湊資料庫：data/image_hashes.json（由 scripts/build_image_hashes.py 生成）

使用流程：
  image_bytes = download_image(message_id)
  prod_code   = identify_product(image_bytes)
  if prod_code:
      # 查詢 Ecount 庫存 + specs.json 規格
  else:
      # OCR fallback
      text = ocr_extract_text(image_bytes)
"""

import io
import json
import sys
from pathlib import Path

import httpx

from config import settings

# 感知雜湊比對閾值：值越小越嚴格（0=完全相同）
# 實測建議值：
#   <= 8  → 高度可信（相同產品同角度）
#   9-12  → 中度可信（同產品不同角度）
#   13+   → 不可信（容易誤判，停用）
HASH_THRESHOLD        = 6    # pHash 直接命中（更嚴格避免誤判，其餘走 Claude）
HASH_THRESHOLD_WEAK   = 10   # 弱命中（OCR 無結果時才採用）

HASH_DB_PATH = Path("data/image_hashes.json")

try:
    from PIL import Image, ImageFilter, ImageOps
    import imagehash as _imagehash
    import numpy as np
    _VISION_OK = True
except ImportError:
    _VISION_OK = False
    print("[vision] ⚠️  Pillow/imagehash 未安裝，圖片辨識功能停用")
    print("[vision]    執行：pip install Pillow imagehash")

# ── Tesseract OCR ─────────────────────────────────────────────────────────────
_TESSERACT_OK = False
try:
    import pytesseract

    # Windows 常見安裝路徑自動偵測
    if sys.platform == "win32":
        _TESS_CANDIDATES = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for _p in _TESS_CANDIDATES:
            if Path(_p).exists():
                pytesseract.pytesseract.tesseract_cmd = _p
                break

    # 自訂 tessdata 路徑（chi_tra 放在使用者目錄，避免 Program Files 權限問題）
    _USER_TESSDATA = Path.home() / "tessdata"
    if _USER_TESSDATA.exists():
        import os
        os.environ.setdefault("TESSDATA_PREFIX", str(_USER_TESSDATA))

    pytesseract.get_tesseract_version()   # 若找不到 binary 這裡會拋例外
    _TESSERACT_OK = True
    print("[vision] Tesseract OCR 已啟用")
except Exception:
    print("[vision] ⚠️  Tesseract 未安裝，OCR 功能停用")
    print("[vision]    下載安裝：https://github.com/UB-Mannheim/tesseract/wiki")
    print("[vision]    安裝後執行：pip install pytesseract")


def download_image(message_id: str) -> bytes | None:
    """從 LINE 下載圖片內容"""
    try:
        resp = httpx.get(
            f"https://api-data.line.me/v2/bot/message/{message_id}/content",
            headers={"Authorization": f"Bearer {settings.LINE_CHANNEL_ACCESS_TOKEN}"},
            timeout=15,
        )
        if resp.status_code == 200:
            return resp.content
        print(f"[vision] 圖片下載失敗 HTTP {resp.status_code}")
    except Exception as e:
        print(f"[vision] 圖片下載錯誤: {e}")
    return None


# 轉帳截圖偵測關鍵字（常見台灣銀行轉帳成功頁面出現的文字）
_TRANSFER_KEYWORDS = [
    "轉帳", "交易成功", "轉入帳號", "轉入戶名", "轉帳金額",
    "匯款", "付款成功", "交易序號", "TWD", "NTD",
    "轉出", "收款", "交易完成", "轉入行庫",
]


def is_transfer_screenshot(image_bytes: bytes) -> bool:
    """
    偵測圖片是否為銀行/支付轉帳截圖（不需 OCR，支援所有銀行顏色）。

    判斷條件（三項都符合才回傳 True）：
    1. 直向手機截圖（高 / 寬 > 1.5）
    2. 上方 1/4 有顯著的彩色 header（飽和度高的色塊，占 25%+）
       → 任何顏色皆可：綠、藍、紅、橘、紫等各銀行主色調
    3. 中段 1/3 有大面積白/淺色區（交易明細白底，占 35%+）

    適用銀行：將來、台新、國泰、中信、玉山、台銀、富邦、
              Line Pay、街口、全支付 … 等各種介面
    """
    if not _VISION_OK:
        return False
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        w, h = img.size

        # 1. 直向截圖
        if h / max(w, 1) < 1.5:
            return False

        # 2. 上方 1/4 有高飽和度的彩色 header（任何顏色）
        #    判斷方式：max(R,G,B) - min(R,G,B) > 50（色彩飽和）
        #    且亮度不能太暗（避免黑色截圖誤判）
        top = img.crop((0, 0, w, h // 4))
        top_arr = np.array(top)  # shape: (h, w, 3)
        max_rgb = top_arr.max(axis=2)
        min_rgb = top_arr.min(axis=2)
        diff = max_rgb - min_rgb
        colorful = int(np.sum((diff > 50) & (max_rgb > 80)))
        total_pixels = top_arr.shape[0] * top_arr.shape[1]
        if colorful / max(total_pixels, 1) < 0.25:
            return False

        # 3. 中段有大面積白/淺色區（銀行 UI 交易明細區）
        mid_start = h // 4
        mid_end   = h * 3 // 4
        mid = img.crop((0, mid_start, w, mid_end))
        mid_pixels = list(mid.getdata())
        light = sum(1 for r, g, b in mid_pixels if min(r, g, b) > 185)
        if light / max(len(mid_pixels), 1) < 0.35:
            return False

        return True
    except Exception:
        return False


def _auto_rebuild_if_stale() -> None:
    """
    比較 image_hashes.json 與產品照片資料夾的修改時間，
    若資料夾有比 DB 還新的圖片 → 自動重建雜湊庫。
    """
    import importlib, sys
    # 抑制 Windows 磁碟機未連線時的彈窗
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass

    image_dir = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品照片")
    try:
        if not image_dir.exists():
            return
    except OSError:
        return

    image_exts = {".jpg", ".jpeg", ".png", ".webp"}

    # hash DB 的修改時間（不存在視為 0）
    db_mtime = HASH_DB_PATH.stat().st_mtime if HASH_DB_PATH.exists() else 0.0

    # 照片資料夾中最新的圖片時間
    latest = max(
        (f.stat().st_mtime for f in image_dir.iterdir() if f.suffix.lower() in image_exts),
        default=0.0,
    )

    if latest <= db_mtime:
        return  # 資料庫是最新的，不需要重建

    print("[vision] 偵測到新圖片，自動重建雜湊庫...")
    try:
        # 直接 import 並執行 build_image_hashes 的 build()
        spec_mod = importlib.util.spec_from_file_location(
            "build_image_hashes",
            Path(__file__).parent.parent / "scripts" / "build_image_hashes.py",
        )
        mod = importlib.util.module_from_spec(spec_mod)
        spec_mod.loader.exec_module(mod)
        mod.build()
        print("[vision] 雜湊庫重建完成")
    except Exception as e:
        print(f"[vision] 自動重建失敗: {e}")


def identify_product(image_bytes: bytes) -> str | None:
    """
    識別圖片中的產品，回傳產品編號（PROD_CD）。
    差值 ≤ HASH_THRESHOLD（10）才直接回傳；
    差值在 10-15 之間屬於「弱命中」，需搭配 OCR 使用
    → 請改呼叫 identify_product_smart()。
    """
    code, diff = _identify_product_raw(image_bytes)
    if code and diff <= HASH_THRESHOLD:
        return code
    return None


def identify_product_weak(image_bytes: bytes) -> str | None:
    """
    傳回「弱命中」結果（差值 10-15）。
    當 OCR 也找不到任何候選詞時，才使用此結果作為最後備援。
    """
    code, diff = _identify_product_raw(image_bytes)
    if code and diff <= HASH_THRESHOLD_WEAK:
        return code
    return None


def _identify_product_raw(image_bytes: bytes) -> tuple[str | None, int]:
    """
    pHash 比對核心，回傳 (best_code, best_diff)。
    """
    if not _VISION_OK:
        return None, 999

    _auto_rebuild_if_stale()

    if not HASH_DB_PATH.exists():
        print("[vision] 雜湊資料庫不存在，請先執行 scripts/build_image_hashes.py")
        return None, 999

    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        query_phash = _imagehash.phash(img)
        query_chash = _imagehash.colorhash(img)

        with open(HASH_DB_PATH, encoding="utf-8") as f:
            db: list[dict] = json.load(f)

        if not db:
            return None, 999

        best_code = None
        best_diff = 999

        # 第一輪：pHash 比對
        for entry in db:
            stored_p = _imagehash.hex_to_hash(entry["hash"])
            diff = query_phash - stored_p
            if diff < best_diff:
                best_diff = diff
                best_code = entry["code"]

        if best_diff <= HASH_THRESHOLD:
            print(f"[vision] pHash 高可信 → {best_code}（差值={best_diff}，閾值={HASH_THRESHOLD}）")
            return best_code, best_diff
        elif best_diff <= HASH_THRESHOLD_WEAK:
            print(f"[vision] pHash 弱命中 → {best_code}（差值={best_diff}，僅在 OCR 無結果時採用）")
            return best_code, best_diff

        # 第二輪：pHash 失敗 → colorhash 補救
        import numpy as _np
        query_ch_flat = query_chash.hash.flatten().tolist()
        ch_best_code = None
        ch_best_diff = 999
        for entry in db:
            stored_ch = entry.get("chash")
            if not stored_ch or not isinstance(stored_ch, list):
                continue
            # 比對：計算不同 bit 數
            diff_c = sum(a != b for a, b in zip(query_ch_flat, stored_ch))
            if diff_c < ch_best_diff:
                ch_best_diff = diff_c
                ch_best_code = entry["code"]

        if ch_best_diff <= 1 and best_diff <= HASH_THRESHOLD:
            # colorhash 補救：顏色完全相同 + pHash 也在高可信範圍才採用
            print(f"[vision] colorHash 補救命中 → {ch_best_code}（色差={ch_best_diff}，pHash={best_diff}）")
            return ch_best_code, ch_best_diff
        elif ch_best_diff <= 2:
            print(f"[vision] colorHash 色調相近但 pHash 差太遠（{best_diff}），不採用 {ch_best_code}")

        print(f"[vision] 無匹配（pHash 最近={best_code}/{best_diff}，colorHash 最近={ch_best_code}/{ch_best_diff}）")
        return best_code, best_diff

    except Exception as e:
        print(f"[vision] 圖片辨識失敗: {e}")
        return None, 999


# ── OCR ───────────────────────────────────────────────────────────────────────

def ocr_extract_text(image_bytes: bytes) -> str:
    """
    用 Tesseract OCR 辨識圖片中的文字，回傳辨識結果字串。
    Tesseract 未安裝時回傳空字串。

    預處理流程：
    1. 放大 2x（小字清晰度提升）
    2. 轉灰階 → 二值化（OTSU 自動閾值）
    3. 優先嘗試繁中 + 英文，若無語言包則 fallback 英文
    """
    if not _TESSERACT_OK or not _VISION_OK:
        return ""

    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")

        # 1. 條件放大（僅小圖放大，且限制最大尺寸避免記憶體浪費）
        MAX_OCR_DIM = 2000
        w, h = img.size
        scale = 1
        if max(w, h) < 800:
            scale = 2
        if max(w * scale, h * scale) > MAX_OCR_DIM:
            scale = MAX_OCR_DIM / max(w, h)
        if scale != 1:
            img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

        # 2. 灰階 + 二值化（讓文字更清晰）
        gray = ImageOps.grayscale(img)
        # Pillow 無內建 OTSU，改用固定閾值 128（效果已足夠）
        bw = gray.point(lambda p: 255 if p > 128 else 0)

        # 3. OCR（繁中優先，fallback 英文）
        try:
            text = pytesseract.image_to_string(bw, lang="chi_tra+eng",
                                               config="--psm 6")
        except pytesseract.TesseractError:
            text = pytesseract.image_to_string(bw, lang="eng",
                                               config="--psm 6")

        result = text.strip()
        print(f"[vision] OCR 結果（{len(result)} 字）: {result[:80]!r}")
        return result

    except Exception as e:
        print(f"[vision] OCR 失敗: {e}")
        return ""


def ocr_extract_candidates(image_bytes: bytes) -> list[str]:
    """
    從圖片 OCR 結果中萃取「貨號 / 品名」候選詞列表，依優先序排列：

    優先順序：
      1. 英數貨號（如 BK-001、V3、S-100、K霸）
      2. 英數＋中文混合（如 V3籃球）
      3. 純中文品名片段（如 籃球、洗衣球）
      4. 整行 fallback（去掉太長 / 太短的行）

    回傳 list[str]，呼叫端逐一比對 Ecount，第一個匹配即停止。
    Tesseract 未啟用時回傳空 list。
    """
    import re

    text = ocr_extract_text(image_bytes)
    if not text:
        return []

    seen: set[str] = set()
    candidates: list[str] = []

    def add(s: str) -> None:
        s = s.strip()
        if s and len(s) >= 2 and s not in seen:
            seen.add(s)
            candidates.append(s)

    # ── 1. 英數貨號：字母開頭，可含數字 / 橫線 / 底線，2~15 字元
    #    e.g. BK-001、V3、S-100、KBA
    for m in re.finditer(r'[A-Za-z][A-Za-z0-9\-_]{1,14}', text):
        add(m.group().upper())

    # ── 2. 純數字貨號（部分廠商用純數字編號，如 001、1001）
    for m in re.finditer(r'\b\d{3,8}\b', text):
        add(m.group())

    # ── 3. 英數＋中文 / 中文＋英數混合（如 V3籃球、K霸台用）
    for m in re.finditer(r'[A-Za-z0-9]{1,6}[\u4e00-\u9fff]{1,8}', text):
        add(m.group())
    for m in re.finditer(r'[\u4e00-\u9fff]{1,8}[A-Za-z0-9]{1,6}', text):
        add(m.group())

    # ── 4. 純中文品名：連續中文 2~10 字
    for m in re.finditer(r'[\u4e00-\u9fff]{2,10}', text):
        add(m.group())

    # ── 5. 整行 fallback（排除超長行 / 純符號）
    for line in text.splitlines():
        line = line.strip()
        if 2 <= len(line) <= 30 and re.search(r'[A-Za-z0-9\u4e00-\u9fff]', line):
            add(line)

    print(f"[vision] OCR 候選詞 ({len(candidates)}): {candidates[:10]}")
    return candidates
