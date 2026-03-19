"""
測試：建立單一品項到 Ecount
執行方式：python scripts/test_create_one_product.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import httpx
from config import settings

# ── 直接貼 JSON 在這裡 ────────────────────────────────
PAYLOAD = {
  "ProductList": [{
    "BulkDatas": {
      "PROD_CD":  "TESTBOT010",
      "PROD_DES": "LINE Bot測試品",
      "UNIT":     "個"
    }
  }]
}
# ─────────────────────────────────────────────────────


def get_session_id():
    resp = httpx.post(
        f"{settings.ECOUNT_BASE_URL}/OAPI/V2/OAPILogin",
        json={
            "ZONE":         settings.ECOUNT_ZONE,
            "COMPANY_NO":   settings.ECOUNT_COMPANY_NO,
            "USER_ID":      settings.ECOUNT_USER_ID,
            "API_CERT_KEY": settings.ECOUNT_API_CERT_KEY,
            "LAN_TYPE":     "zh-TW",
        },
        timeout=10,
    )
    data = resp.json()
    session_id = data.get("Data", {}).get("Datas", {}).get("SESSION_ID")
    if not session_id:
        print(f"❌ 登入失敗: {data}")
        return None
    print(f"✅ 登入成功")
    return session_id


def create_product(session_id):
    print(f"\n📦 傳送資料: {PAYLOAD}\n")

    resp = httpx.post(
        f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBasic/SaveBasicProduct",
        params={"SESSION_ID": session_id},
        json=PAYLOAD,
        timeout=15,
    )
    data = resp.json()

    print(f"HTTP 狀態: {resp.status_code}")
    print(f"API Status: {data.get('Status')}")
    print(f"Message:    {data.get('Message', '')}")
    if data.get("Data"):
        print(f"Data:       {data['Data']}")

    if str(data.get("Status")) == "200" and not data.get("Errors"):
        print(f"\n✅ 建立成功！")
    else:
        print(f"\n❌ 失敗")


if __name__ == "__main__":
    print("=== Ecount 建立品項測試 ===\n")
    session_id = get_session_id()
    if session_id:
        create_product(session_id)
