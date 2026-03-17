"""Ecount API 測試 - 正式伺服器 + 倉庫 101"""
import httpx, json
from datetime import datetime

BASE  = "https://oapiIB.ecount.com"
TODAY = datetime.now().strftime("%Y%m%d")

# 登入
login = httpx.post(f"{BASE}/OAPI/V2/OAPILogin", json={
    "COM_CODE":     "851759",
    "USER_ID":      "1127BEAR",
    "API_CERT_KEY": "24953cdf029014c08981aa8d4ec48a1e31",
    "LAN_TYPE":     "zh-TW",
    "ZONE":         "IB",
}, timeout=10)
ld = login.json()
assert ld["Data"]["Code"] == "00", f"Login failed: {ld}"
sid = ld["Data"]["Datas"]["SESSION_ID"]
print(f"[OK] 登入成功 SID={sid[:20]}...")

# 庫存查詢（倉庫 101，查全部品項）
resp = httpx.post(
    f"{BASE}/OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus",
    params={"SESSION_ID": sid},
    json={"PROD_CD": "", "WH_CD": "101", "BASE_DATE": TODAY, "ZERO_FLAG": "Y"},
    timeout=15,
)
d = resp.json()
print(f"[庫存] Status={d.get('Status')} TotalCnt={d.get('Data',{}).get('TotalCnt')}")

results = d.get("Data", {}).get("Result", [])
for row in results[:10]:
    print(f"  {row.get('PROD_CD'):15s}  庫存={row.get('BAL_QTY')}")
if len(results) > 10:
    print(f"  ... 共 {len(results)} 筆")
if d.get("Status") != "200":
    print("Errors:", json.dumps(d.get("Errors"), ensure_ascii=False))
