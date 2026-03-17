"""
測試 Ecount API - 建立品項（InventoryBasic/SaveBasicProduct）

用法：
    python scripts/test_create_product.py

測試流程：
  1. 登入取得 SESSION_ID
  2. 嘗試建立一筆測試品項
  3. 印出完整 API 回應（方便確認欄位名稱與格式）
"""

import sys
import os
import json

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import httpx
from dotenv import load_dotenv
load_dotenv()

from config import settings

BASE_URL = settings.ECOUNT_BASE_URL


# ─────────────────────────────────────────────────────────────
# 1. 登入
# ─────────────────────────────────────────────────────────────
def step1_login() -> str | None:
    print("\n[Step 1] 登入 Ecount API...")
    resp = httpx.post(
        f"{BASE_URL}/OAPI/V2/OAPILogin",
        json={
            "COM_CODE":     settings.ECOUNT_COMPANY_NO,
            "USER_ID":      settings.ECOUNT_USER_ID,
            "API_CERT_KEY": settings.ECOUNT_API_CERT_KEY,
            "LAN_TYPE":     "zh-TW",
            "ZONE":         settings.ECOUNT_ZONE,
        },
        timeout=10,
    )
    data = resp.json()
    if str(data.get("Status")) == "200" and str(data.get("Data", {}).get("Code", "")) == "00":
        sid = data["Data"]["Datas"]["SESSION_ID"]
        print(f"  ✅ 登入成功  SESSION_ID: {sid[:20]}...")
        return sid
    print(f"  ❌ 登入失敗: {data}")
    return None


# ─────────────────────────────────────────────────────────────
# 2. 建立品項
# ─────────────────────────────────────────────────────────────
def step2_create_product(session_id: str, prod_cd: str, prod_name: str) -> None:
    print(f"\n[Step 2] 建立品項 PROD_CD={prod_cd}  品名={prod_name}")

    payload = {
        "PROD_CD":          prod_cd,       # 貨號（必填）
        "PROD_DES":         prod_name,     # 品名（必填）
        "BAL_UNIT":         "個",          # 單位
        "SET_FLAG":         "N",           # 非套件
        "USE_FLAG":         "Y",           # 使用中
    }

    print(f"\n  送出 Payload：")
    print(json.dumps(payload, ensure_ascii=False, indent=4))

    resp = httpx.post(
        f"{BASE_URL}/OAPI/V2/InventoryBasic/SaveBasicProduct",
        params={"SESSION_ID": session_id},
        json=payload,
        timeout=15,
    )

    print(f"\n  HTTP {resp.status_code}")
    try:
        data = resp.json()
        print(f"  完整回應：")
        print(json.dumps(data, ensure_ascii=False, indent=4))

        status = str(data.get("Status"))
        if status == "200":
            print("\n  ✅ 品項建立成功！")
        else:
            errors = data.get("Data", {}).get("Errors", "")
            print(f"\n  ❌ 建立失敗  Status={status}  Errors={errors}")
    except Exception as e:
        print(f"  ❌ 解析回應失敗: {e}")
        print(f"  原始回應: {resp.text[:500]}")


# ─────────────────────────────────────────────────────────────
# 3. 查詢品項（確認是否建立成功）
# ─────────────────────────────────────────────────────────────
def step3_query_product(session_id: str, prod_cd: str) -> None:
    print(f"\n[Step 3] 查詢品項 PROD_CD={prod_cd}...")

    resp = httpx.post(
        f"{BASE_URL}/OAPI/V2/InventoryBasic/GetBasicProductsList",
        params={"SESSION_ID": session_id},
        json={"PROD_CD": prod_cd},
        timeout=15,
    )

    try:
        data = resp.json()
        results = data.get("Data", {}).get("Result", [])
        if results:
            print(f"  ✅ 查到 {len(results)} 筆：")
            for r in results:
                print(f"     {r.get('PROD_CD','')}  {r.get('PROD_DES','')}  單位={r.get('BAL_UNIT','')}")
        else:
            print(f"  ⚠️ 查無此品項（可能建立失敗或有延遲）")
            print(f"  完整回應: {json.dumps(data, ensure_ascii=False)}")
    except Exception as e:
        print(f"  ❌ 錯誤: {e}")


# ─────────────────────────────────────────────────────────────
# 執行
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 55)
    print("  Ecount API 測試 - 建立品項")
    print("=" * 55)

    session_id = step1_login()
    if not session_id:
        print("\n⛔ 登入失敗，請確認 .env 設定")
        sys.exit(1)

    print("\n請輸入要建立的測試品項（輸入後才會真正建立）：")
    prod_cd   = input("  貨號 (PROD_CD)，例 TEST001：").strip().upper()
    prod_name = input("  品名 (PROD_DES)，例 測試品項：").strip()

    if not prod_cd or not prod_name:
        print("⛔ 貨號與品名不能空白")
        sys.exit(1)

    ans = input(f"\n  確定要在 Ecount 建立「{prod_cd} {prod_name}」？[y/N] ").strip().lower()
    if ans != "y":
        print("  已取消")
        sys.exit(0)

    step2_create_product(session_id, prod_cd, prod_name)
    step3_query_product(session_id, prod_cd)

    print("\n完成 ✅")
