"""
測試 Ecount 建立品項 - 使用測試區(sboapi)
Test URL: https://sboapi{ZONE}.ecount.com/OAPI/V2/InventoryBasic/SaveBasicProduct
"""
import sys, os, json
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import httpx
from dotenv import load_dotenv
load_dotenv()
from config import settings

# 測試區 URL（sboapi 取代 oapi）
ZONE         = settings.ECOUNT_ZONE   # "IB"
TEST_BASE    = f"https://sboapi{ZONE}.ecount.com"
PROD_BASE    = settings.ECOUNT_BASE_URL   # https://oapiIB.ecount.com

print(f"測試區 URL : {TEST_BASE}")
print(f"正式區 URL : {PROD_BASE}\n")

def login(base_url: str) -> str | None:
    resp = httpx.post(f"{base_url}/OAPI/V2/OAPILogin", json={
        "COM_CODE":     settings.ECOUNT_COMPANY_NO,
        "USER_ID":      settings.ECOUNT_USER_ID,
        "API_CERT_KEY": settings.ECOUNT_API_CERT_KEY,
        "LAN_TYPE":     "zh-TW",
        "ZONE":         settings.ECOUNT_ZONE,
    }, timeout=10)
    d = resp.json()
    if str(d.get("Status")) == "200":
        return d["Data"]["Datas"]["SESSION_ID"]
    print(f"  登入失敗: {d}")
    return None

def create_product(base_url: str, sid: str, prod_cd: str, prod_name: str, unit: str = "個"):
    payload = {
        "ProductList": [{
            "BulkDatas": {
                "PROD_CD":  prod_cd,
                "PROD_DES": prod_name,
                "UNIT":     unit,
            }
        }]
    }
    print(f"Payload:\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n")

    r = httpx.post(
        f"{base_url}/OAPI/V2/InventoryBasic/SaveBasicProduct",
        params={"SESSION_ID": sid},
        json=payload,
        timeout=15,
    )
    data = r.json()
    print(f"HTTP {r.status_code}")
    print(json.dumps(data, ensure_ascii=False, indent=2))

    status  = str(data.get("Status"))
    errors  = data.get("Errors", [])
    success = data.get("Data", {}).get("SuccessCnt", 0) if isinstance(data.get("Data"), dict) else 0

    if status == "200" and not errors and success > 0:
        print("\n✅ 建立品項成功！")
    elif errors:
        for e in errors:
            print(f"\n❌ {e.get('Code')} : {e.get('Message')}")
    else:
        print(f"\n⚠️ Status={status}  SuccessCnt={success}")


# ── 執行 ──────────────────────────────────────────────
print("=" * 55)
print("  STEP 1：測試區登入")
print("=" * 55)
sid = login(TEST_BASE)
if not sid:
    print("⛔ 測試區登入失敗（測試區帳號設定可能不同）")
    print("   改用正式區登入驗證格式...")
    sid = login(PROD_BASE)
    if not sid:
        sys.exit(1)
    base = PROD_BASE
else:
    base = TEST_BASE
    print(f"✅ 登入成功  SESSION_ID: {sid[:20]}...")

print(f"\n使用伺服器: {base}\n")

print("=" * 55)
print("  STEP 2：建立品項 TESTBOT001")
print("=" * 55)
create_product(base, sid, "TESTBOT001", "LINE Bot測試品", "個")
