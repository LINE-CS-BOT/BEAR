"""
測試 Ecount API - 登入 + 建立訂單

用法：
    python scripts/test_save_order.py

會依序測試：
  1. 登入（Session）
  2. 取得品項清單，印出前 5 筆
  3. 用第一筆品項建立測試訂單（qty=1），確認成功後印出單號
"""

import sys
import os

# 強制 UTF-8 輸出（解決 Windows cp950 問題）
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# 讓腳本可以從 scripts/ 以外的地方 import
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
    try:
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
        status = str(data.get("Status"))
        code   = str(data.get("Data", {}).get("Code", ""))

        if status == "200" and code == "00":
            session_id = data["Data"]["Datas"]["SESSION_ID"]
            print(f"  ✅ 登入成功！SESSION_ID: {session_id[:20]}...")
            return session_id
        else:
            print(f"  ❌ 登入失敗 Status={status} Code={code}")
            print(f"     回應：{data}")
            return None
    except Exception as e:
        print(f"  ❌ 連線錯誤: {e}")
        return None


# ─────────────────────────────────────────────────────────────
# 2. 取得品項清單（只取前 5 筆）
# ─────────────────────────────────────────────────────────────
def step2_get_products(session_id: str) -> list[dict]:
    print("\n[Step 2] 取得品項清單（前 5 筆）...")
    try:
        resp = httpx.post(
            f"{BASE_URL}/OAPI/V2/InventoryBasic/GetBasicProductsList",
            params={"SESSION_ID": session_id},
            json={},
            timeout=20,
        )
        data = resp.json()
        if str(data.get("Status")) == "200":
            results = data.get("Data", {}).get("Result", [])
            products = [
                {"code": r.get("PROD_CD", "").strip(),
                 "name": r.get("PROD_DES", "").strip()}
                for r in results[:5] if r.get("PROD_CD")
            ]
            print(f"  ✅ 取得成功！共 {len(results)} 筆，前 5 筆：")
            for p in products:
                print(f"     {p['code']:15s}  {p['name']}")
            return products
        else:
            print(f"  ❌ 失敗: {data.get('Status')} {data.get('Data', {}).get('Errors', '')}")
            return []
    except Exception as e:
        print(f"  ❌ 錯誤: {e}")
        return []


# ─────────────────────────────────────────────────────────────
# 3. 建立測試訂單
# ─────────────────────────────────────────────────────────────
def step3_save_order(session_id: str, prod_cd: str) -> None:
    print(f"\n[Step 3] 建立測試訂單（PROD_CD={prod_cd}, QTY=1）...")
    from datetime import datetime
    today = datetime.now().strftime("%Y%m%d")

    try:
        resp = httpx.post(
            f"{BASE_URL}/OAPI/V2/SaleOrder/SaveSaleOrder",
            params={"SESSION_ID": session_id},
            json={"SaleOrderList": [{
                "BulkDatas": {
                    "IO_DATE":       today,
                    "UPLOAD_SER_NO": "1",
                    "CUST":          "TEST_LINE_BOT",   # 測試用客戶名
                    "WH_CD":         "101",
                    "PROD_CD":       prod_cd,
                    "QTY":           "1",
                    "PRICE":         "",
                }
            }]},
            timeout=10,
        )
        data = resp.json()
        print(f"  HTTP {resp.status_code}  Status={data.get('Status')}")

        if str(data.get("Status")) == "200":
            slip_nos = data.get("Data", {}).get("SlipNos", [])
            success  = data.get("Data", {}).get("SuccessCnt", 0)
            details  = data.get("Data", {}).get("ResultDetails", "")
            print(f"  ✅ 訂單建立成功！SuccessCnt={success}")
            print(f"     SlipNos: {slip_nos}")
            if details:
                print(f"     Details: {details}")
        else:
            errors  = data.get("Data", {}).get("Errors", "")
            details = data.get("Data", {}).get("ResultDetails", "")
            print(f"  ❌ 訂單建立失敗！")
            print(f"     Errors: {errors}")
            print(f"     Details: {details}")
            print(f"     完整回應: {data}")
    except Exception as e:
        print(f"  ❌ 錯誤: {e}")


# ─────────────────────────────────────────────────────────────
# 執行
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 55)
    print("  Ecount API 測試 - save_order")
    print("=" * 55)

    session_id = step1_login()
    if not session_id:
        print("\n⛔ 登入失敗，請確認 .env 設定後再試")
        sys.exit(1)

    products = step2_get_products(session_id)

    if not products:
        print("\n⛔ 無法取得品項，無法繼續測試訂單")
        sys.exit(1)

    # 用第一筆品項建立測試訂單
    test_prod = products[0]["code"]
    print(f"\n  → 使用 {test_prod}（{products[0]['name']}）建立測試訂單")
    ans = input("  確定要在 Ecount 建立這筆測試訂單嗎？[y/N] ").strip().lower()
    if ans == "y":
        step3_save_order(session_id, test_prod)
    else:
        print("  跳過訂單測試。")

    print("\n完成 ✅")
