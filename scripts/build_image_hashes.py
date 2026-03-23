"""
掃描產品圖片庫，建立感知雜湊（pHash）資料庫

輸出：data/image_hashes.json
圖片庫：H:\其他電腦\我的電腦\小蠻牛\產品照片\

辨識規則：
  - 檔名以「字母 + 4位以上數字」開頭者視為有編號的圖片
    例如：Z3469.jpg → Z3469、T1186-1.jpg → T1186、S0626A.jpg → S0626
  - 其他命名格式（S__xxx、342xxx 等）跳過，不納入雜湊庫

執行方式（新增圖片後重跑以更新資料庫）：
  python scripts/build_image_hashes.py
"""

import io
import json
import re
import sys
from pathlib import Path

# Windows cmd 編碼修正
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# 壓制 Windows 磁碟機未連線彈窗（H: 雲端磁碟離線時靜默失敗）
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
    except Exception:
        pass

try:
    from PIL import Image
    import imagehash
except ImportError:
    print("需要安裝 Pillow 和 imagehash：")
    print("  pip install Pillow imagehash")
    sys.exit(1)

IMAGE_DIR = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品照片")
OUTPUT    = Path(__file__).parent.parent / "data" / "image_hashes.json"

# 從檔名提取產品編號：T1186、Z3495、S0626A → 取字母+數字部分
_CODE_RE = re.compile(r"^([A-Z]\d{4,})", re.IGNORECASE)


def extract_code(stem: str) -> str | None:
    """從檔名（不含副檔名）提取產品編號"""
    m = _CODE_RE.match(stem.upper())
    return m.group(1).upper() if m else None


def build():
    try:
        exists = IMAGE_DIR.exists()
    except OSError:
        exists = False
    if not exists:
        print(f"找不到圖片庫資料夾：{IMAGE_DIR}")
        sys.exit(1)

    image_exts = {".jpg", ".jpeg", ".png", ".webp"}
    entries: list[dict] = []
    ok = skip = err = 0

    all_images = [f for f in sorted(IMAGE_DIR.iterdir()) if f.suffix.lower() in image_exts]
    print(f"共掃描 {len(all_images)} 張圖片...\n")

    for f in all_images:
        code = extract_code(f.stem)
        if not code:
            print(f"  ⏭  略過（無法辨識編號）：{f.name}")
            skip += 1
            continue

        try:
            img = Image.open(f).convert("RGB")
            h   = imagehash.phash(img)
            ch  = imagehash.colorhash(img)
            entries.append({
                "code": code,
                "file": f.name,
                "hash": str(h),
                "chash": ch.hash.flatten().tolist(),
            })
            print(f"  ✓  {f.name} → {code}")
            ok += 1
        except Exception as e:
            print(f"  ✗  {f.name}: {e}")
            err += 1

    OUTPUT.parent.mkdir(exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as out:
        json.dump(entries, out, ensure_ascii=False, indent=2)

    print(f"\n完成！✓ {ok} 張已建立 | ⏭ {skip} 張略過 | ✗ {err} 張失敗")
    print(f"雜湊資料庫儲存至：{OUTPUT}")
    print(f"\n💡 無法辨識編號的圖片，可手動重新命名（格式：XXXXX.jpg）後重跑此腳本")


if __name__ == "__main__":
    build()
