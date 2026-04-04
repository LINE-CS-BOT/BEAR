"""
Ecount ↔ customers.db 客戶同步腳本

功能：
1. 從 Ecount 抓取所有客戶清單
2. 與 customers.db 的 LINE 客戶比對（手機 → 姓名）
3. 已在 Ecount 的 → 寫入 ecount_cust_cd
4. 不在 Ecount 的 LINE 客戶 → 自動在 Ecount 建立，並寫入 ecount_cust_cd

使用方式：
    python scripts/sync_cust_ecount.py
    python scripts/sync_cust_ecount.py --dry-run   # 只顯示結果，不寫入
"""

import sys
import io
import time
import argparse
from pathlib import Path

# Windows 終端機強制 UTF-8 輸出
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 讓 import 能找到上層模組
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from services.ecount import ecount_client
from storage.customers import customer_store


# --------------------------------------------------------------------------
# DB 更新輔助（支援 line_user_id 或 db id 兩種 key）
# --------------------------------------------------------------------------

def _update_cust_code(uid: str, db_id, ecount_code: str) -> bool:
    """
    更新 ecount_cust_cd。
    優先用 line_user_id；若 uid 是空字串（CSV 匯入客戶）則用 DB id。
    """
    if uid:
        return customer_store.update_ecount_cust_cd(uid, ecount_code)
    if db_id:
        return customer_store.update_ecount_cust_cd_by_db_id(db_id, ecount_code)
    return False


# --------------------------------------------------------------------------
# 工具函式
# --------------------------------------------------------------------------

def _normalize_phone(phone: str) -> str:
    """去除空白、dash，統一格式（台灣手機：09xxxxxxxx）"""
    if not phone:
        return ""
    return phone.replace(" ", "").replace("-", "").strip()


def _build_ecount_index(ecount_custs: list[dict]) -> tuple[dict, dict]:
    """
    建立兩個 lookup table：
        phone_index : normalized_phone → ecount_code
        name_index  : cust_name → ecount_code
    """
    phone_index: dict[str, str] = {}
    name_index:  dict[str, str] = {}

    for c in ecount_custs:
        code = c["code"]
        for ph_field in (c["phone"], c["tel"]):
            norm = _normalize_phone(ph_field)
            if norm:
                phone_index[norm] = code
        if c["name"]:
            name_index[c["name"]] = code

    return phone_index, name_index


def _generate_ecount_code(line_user_id: str, phone: str, name: str) -> str:
    """
    為新 LINE 客戶產生 Ecount 客戶代碼。
    同一支手機 → 同一代碼（同一個人的多個 LINE 帳號）。

    優先順序：
        1. 手機後 9 碼 (L-9xxxxxxxx)
        2. LINE user_id 末 8 碼 (L-Uxxxxxxx)
        3. 顯示名稱前 8 字元 (L-XXXXXXXX)
    前綴 'L-' 避免與現有代碼衝突。
    """
    if phone:
        norm = _normalize_phone(phone)
        suffix = norm[-9:] if len(norm) >= 9 else norm
        return f"L-{suffix}"
    if line_user_id:
        return f"L-{line_user_id[-8:]}"
    if name:
        safe = "".join(c for c in name if c.isalnum())[:8]
        if safe:
            return f"L-{safe}"
    import time
    return f"L-{int(time.time() * 1000) % 100000000}"


# --------------------------------------------------------------------------
# 主要流程
# --------------------------------------------------------------------------

def sync(dry_run: bool = False):
    print("=" * 60)
    print("Ecount ↔ customers.db 客戶同步")
    print("=" * 60)

    # 1. 抓 Ecount 客戶清單
    print("\n[1/4] 從 Ecount 抓取客戶清單...")
    ecount_custs = ecount_client.get_customers_list()
    if not ecount_custs:
        print("  ⚠️  無法取得 Ecount 客戶清單（API 未設定或回傳空）")
        print("  → 僅執行「找不到 → 建立」步驟仍會跳過（需要 API）")
    else:
        print(f"  ✅ 取得 {len(ecount_custs)} 筆 Ecount 客戶")

    phone_index, name_index = _build_ecount_index(ecount_custs)
    # code → name 反向索引
    code_to_name = {c["code"]: c["name"] for c in ecount_custs if c.get("name")}

    # 2. 讀取 LINE 客戶
    print("\n[2/4] 讀取 customers.db...")
    line_custs = customer_store.all(limit=9999)
    print(f"  共 {len(line_custs)} 筆 LINE 客戶")

    # 3. 比對 & 更新
    print("\n[3/4] 比對客戶...")
    matched   = []
    to_create = []
    already   = []

    for c in line_custs:
        uid  = c.get("line_user_id") or ""
        name = (c.get("display_name") or "").strip()
        phone = _normalize_phone(c.get("phone") or "")
        db_id = c.get("id")

        # 已有 ecount_cust_cd → 檢查 real_name 是否需要補填
        if c.get("ecount_cust_cd"):
            already.append((name, c["ecount_cust_cd"]))
            if uid and not c.get("real_name") and not dry_run:
                ec_name = code_to_name.get(c["ecount_cust_cd"], "")
                if ec_name:
                    customer_store.update_real_name(uid, ec_name)
                    print(f"  [補填] {name or uid} → 真實姓名：{ec_name}")
            continue

        ecount_code = None

        # 優先：手機比對 Ecount 已有客戶
        if phone and phone in phone_index:
            ecount_code = phone_index[phone]

        # 次之：姓名完全比對 Ecount 已有客戶
        if not ecount_code and name and name in name_index:
            ecount_code = name_index[name]

        if ecount_code:
            matched.append((uid, db_id, name, phone, ecount_code))
        else:
            new_code = _generate_ecount_code(uid, phone, name)
            to_create.append((uid, db_id, name, phone, new_code))

    print(f"  已綁定: {len(already)} 筆（跳過）")
    print(f"  命中比對: {len(matched)} 筆 → 更新 ecount_cust_cd")
    print(f"  未命中: {len(to_create)} 筆（僅記錄，不自動建立）")

    # 4. 寫入比對結果（只同步已比對到的，不建立新客戶）
    print("\n[4/4] 寫入比對結果...")

    updated = 0
    for uid, db_id, name, phone, ecount_code in matched:
        ec_name = code_to_name.get(ecount_code, "")
        print(f"  [比對] {name or uid}  →  {ecount_code}（{ec_name}）")
        if not dry_run:
            ok = _update_cust_code(uid, db_id, ecount_code)
            if ok:
                updated += 1
            # real_name 為空 → 用 Ecount 客戶名填上
            if ec_name and uid:
                cust = customer_store.get_by_line_id(uid)
                if cust and not cust.get("real_name"):
                    customer_store.update_real_name(uid, ec_name)
                    print(f"    → 填入真實姓名：{ec_name}")

    # 列出未命中的客戶（僅記錄，不建立）
    if to_create:
        print(f"\n  以下 {len(to_create)} 位 LINE 客戶未在 Ecount 找到（不自動建立）：")
        for uid, db_id, name, phone, new_code in to_create:
            print(f"    • {name or uid}  電話:{phone or '無'}")

    # 摘要
    print("\n" + "=" * 60)
    if dry_run:
        print("【DRY RUN 模式，未實際寫入】")
    print(f"已綁定（跳過）: {len(already)}")
    print(f"比對命中並更新: {updated if not dry_run else len(matched)}")
    print(f"未命中（未建立）: {len(to_create)}")
    print("=" * 60)


# --------------------------------------------------------------------------
# 入口
# --------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="同步 Ecount ↔ LINE 客戶")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="只顯示結果，不實際寫入 DB 或呼叫 Ecount API"
    )
    args = parser.parse_args()
    sync(dry_run=args.dry_run)
