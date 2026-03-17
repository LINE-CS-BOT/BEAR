"""
從 LINE OA 匯出的 CSV（ZIP 或解壓目錄）批次匯入客戶資料到 customers.db

用法：
    python scripts/import_customers.py                          # 預設 ZIP
    python scripts/import_customers.py path/to/file.zip        # 指定 ZIP
    python scripts/import_customers.py path/to/csv_dir         # 指定目錄
"""

import csv
import glob
import io
import os
import re
import sys
import zipfile
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent.parent))
from storage.customers import customer_store

DEFAULT_ZIP = r"C:\Users\bear\Desktop\code\line_oa_chat_csv_260311_033638.zip"

_PHONE_RE = re.compile(r"09\d{8}")
_ADDR_RE  = re.compile(
    r"[\u4e00-\u9fff]{2,5}[市縣][\u4e00-\u9fff]{1,5}[區鄉鎮市]"
    r"[\u4e00-\u9fff\d\-]+[路街][^\n,，。]{0,40}"
)


def extract_phones(text: str) -> list[str]:
    return list(dict.fromkeys(_PHONE_RE.findall(text)))


def extract_address(text: str) -> str:
    m = _ADDR_RE.search(text)
    return m.group(0).strip() if m else ""


def extract_vip_note(chat_label: str) -> str:
    return "VIP" if chat_label.startswith("VIP") else ""


def parse_rows(rows: list[list[str]], chat_label: str) -> dict:
    """解析 CSV 資料列，回傳客戶資訊"""
    display_name = ""
    phones: list[str] = []
    address = ""

    # 從檔名抓電話
    phones.extend(extract_phones(chat_label))

    for row in rows:
        if len(row) < 5:
            continue
        sender_type, sender_name, _, _, content = row[0], row[1], row[2], row[3], row[4]

        if sender_type == "User" and sender_name not in ("Unknown", "") and not display_name:
            display_name = sender_name

        # 從所有訊息內容抓電話
        phones.extend(extract_phones(content))

        # 從 User 訊息抓地址
        if sender_type == "User" and not address:
            address = extract_address(content)

    phones = list(dict.fromkeys(phones))
    return {
        "display_name": display_name or chat_label,
        "chat_label": chat_label,
        "phones": phones,
        "address": address,
        "note": extract_vip_note(chat_label),
    }


def parse_csv_content(content_bytes: bytes, filename: str) -> dict:
    parts = filename.replace(".csv", "").split("_", 3)
    chat_label = parts[3] if len(parts) >= 4 else filename

    try:
        text = content_bytes.decode("utf-8-sig", errors="replace")
        lines = text.splitlines()
        reader = list(csv.reader(lines[4:]))  # 跳過前 4 行 header
        return parse_rows(reader, chat_label)
    except Exception as e:
        print(f"  ⚠️  解析失敗 {filename}: {e}")
        return {}


def iter_from_zip(zip_path: str):
    """從 ZIP 逐一 yield (filename, content_bytes)"""
    with zipfile.ZipFile(zip_path) as z:
        for info in z.infolist():
            if info.filename.endswith(".csv"):
                yield info.filename, z.read(info.filename)


def iter_from_dir(dir_path: str):
    """從目錄逐一 yield (filename, content_bytes)"""
    for fpath in sorted(glob.glob(os.path.join(dir_path, "*.csv"))):
        with open(fpath, "rb") as f:
            yield os.path.basename(fpath), f.read()


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_ZIP

    if os.path.isdir(source):
        file_iter = iter_from_dir(source)
        print(f"讀取目錄：{source}")
    elif zipfile.is_zipfile(source):
        file_iter = iter_from_zip(source)
        print(f"讀取 ZIP：{source}")
    else:
        print(f"找不到有效的 ZIP 或目錄：{source}")
        sys.exit(1)

    print()
    imported = 0
    skipped = 0

    for filename, content_bytes in file_iter:
        info = parse_csv_content(content_bytes, filename)
        if not info:
            skipped += 1
            continue

        cid = customer_store.import_from_csv_data(
            display_name=info["display_name"],
            chat_label=info["chat_label"],
            phones=info["phones"],
            address=info["address"],
            note=info["note"],
        )

        phone_str = ", ".join(info["phones"]) if info["phones"] else "（無電話）"
        addr_str  = info["address"][:20] + "…" if len(info["address"]) > 20 else info["address"] or "（無地址）"
        print(f"  OK [{cid:>4}] {info['display_name']:<25} | {phone_str:<14} | {addr_str}")
        imported += 1

    total = customer_store.count()
    print(f"\n匯入完成！新增/更新 {imported} 筆，跳過 {skipped} 筆")
    print(f"資料庫目前共 {total} 位客戶")


if __name__ == "__main__":
    main()
