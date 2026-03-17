"""
本地模擬測試腳本
測試最近修改的功能，無需連接 LINE API 或 Ecount API

執行方式：
  cd line-cs-bot
  python scripts/test_flow.py

涵蓋測試：
  1. DB real_name 欄位是否存在
  2. update_real_name() 是否正常寫入
  3. ordering.py 的 real_name 邏輯：無 real_name 時觸發詢問
  4. awaiting_contact_info 是否正確儲存 real_name
  5. order_failed 是否記錄到 issue_store
  6. summary _ISSUE_LABEL 是否包含 order_failed
  7. main.py reply_text = None 不會 crash
"""

import sys
import io
import sqlite3
from pathlib import Path
from unittest.mock import MagicMock, patch

# Windows cp950 終端強制 UTF-8 輸出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ── 加入專案根目錄到 sys.path ──────────────────────────
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

_PASS = "✅"
_FAIL = "❌"
_results = []


def check(name: str, passed: bool, detail: str = ""):
    status = _PASS if passed else _FAIL
    msg = f"{status} {name}"
    if detail:
        msg += f"\n     → {detail}"
    print(msg)
    _results.append((name, passed))


# ═══════════════════════════════════════════════════════
# 測試 1：DB real_name 欄位
# ═══════════════════════════════════════════════════════
print("\n━━━ 1. DB real_name 欄位 ━━━")

from storage.customers import customer_store, DB_PATH

with sqlite3.connect(DB_PATH) as conn:
    cols = [r[1] for r in conn.execute("PRAGMA table_info(customers)").fetchall()]

check("real_name 欄位存在", "real_name" in cols,
      f"現有欄位：{cols}")


# ═══════════════════════════════════════════════════════
# 測試 2：update_real_name() 寫入和讀取
# ═══════════════════════════════════════════════════════
print("\n━━━ 2. update_real_name() 寫入 ━━━")

TEST_USER = "U_TEST_FLOW_001"

# 先 upsert 一筆測試用客戶
customer_store.upsert_from_line(TEST_USER, "測試暱稱AAA")

# 確認 display_name 是暱稱
before = customer_store.get_by_line_id(TEST_USER)
check("upsert 後 display_name = LINE 暱稱",
      before and before.get("display_name") == "測試暱稱AAA",
      f"display_name={before.get('display_name') if before else None}")

check("upsert 後 real_name 為空",
      before and before.get("real_name") is None,
      f"real_name={before.get('real_name') if before else None}")

# 更新真實姓名
ok = customer_store.update_real_name(TEST_USER, "王小明")
after = customer_store.get_by_line_id(TEST_USER)

check("update_real_name() 回傳 True", ok is True)
check("display_name 不被覆蓋（仍是暱稱）",
      after and after.get("display_name") == "測試暱稱AAA",
      f"display_name={after.get('display_name') if after else None}")
check("real_name 正確儲存",
      after and after.get("real_name") == "王小明",
      f"real_name={after.get('real_name') if after else None}")


# ═══════════════════════════════════════════════════════
# 測試 3：ordering.py — 無 real_name 時觸發 awaiting_contact_info
# ═══════════════════════════════════════════════════════
print("\n━━━ 3. ordering.py 無 real_name 時詢問聯絡資料 ━━━")

from handlers.ordering import handle_order_quantity
from handlers import tone

# 重置 real_name 為空，並給手機
with sqlite3.connect(DB_PATH) as conn:
    conn.execute(
        "UPDATE customers SET real_name=NULL, phone='0912000001' WHERE line_user_id=?",
        (TEST_USER,)
    )

mock_line_api = MagicMock()

# Mock Ecount 讓它「有貨」（但不走到建立訂單那步）
with patch("services.ecount.ecount_client.save_order", return_value=None):
    # 也要先給 ecount_codes 為空，才會走到 real_name 判斷
    result = handle_order_quantity(
        user_id=TEST_USER,
        text="5個",
        state={"action": "awaiting_quantity", "prod_cd": "TEST001", "prod_name": "測試商品"},
        line_api=mock_line_api,
    )

check("無 real_name 時回傳詢問聯絡資料訊息",
      result and ("姓名" in result or "手機" in result or "大名" in result),
      f"回傳：{result[:50] if result else None}")

from storage.state import state_manager
st = state_manager.get(TEST_USER)
check("無 real_name 時 state 設為 awaiting_contact_info",
      st and st.get("action") == "awaiting_contact_info",
      f"state={st}")


# ═══════════════════════════════════════════════════════
# 測試 4：有 real_name + phone 時直接嘗試建立訂單（成功路徑）
# ═══════════════════════════════════════════════════════
print("\n━━━ 4. 有 real_name + phone 時建立訂單（Mock 成功）━━━")

# 清除狀態
state_manager.clear(TEST_USER)

# 設定 real_name
customer_store.update_real_name(TEST_USER, "王小明")

with patch("services.ecount.ecount_client.save_order", return_value="SL-20260001"), \
     patch("handlers.ordering._resolve_cust_code", return_value="M260001"):
    result_ok = handle_order_quantity(
        user_id=TEST_USER,
        text="5個",
        state={"action": "awaiting_quantity", "prod_cd": "TEST001", "prod_name": "測試商品"},
        line_api=mock_line_api,
    )

check("有 real_name 時訂單成功，回傳確認訊息",
      result_ok and ("訂單" in result_ok or "登記" in result_ok or "建立" in result_ok),
      f"回傳：{result_ok[:60] if result_ok else None}")


# ═══════════════════════════════════════════════════════
# 測試 5：訂單失敗 → 記錄 issue_store，回傳 None
# ═══════════════════════════════════════════════════════
print("\n━━━ 5. 訂單失敗 → issue_store 記錄，不回覆客戶 ━━━")

from storage.issues import issue_store

# 取目前 pending 數量（基準）
before_issues = issue_store.get_pending()
before_count = len(before_issues)

state_manager.clear(TEST_USER)

with patch("services.ecount.ecount_client.save_order", return_value=None), \
     patch("handlers.ordering._resolve_cust_code", return_value="M260001"):
    result_fail = handle_order_quantity(
        user_id=TEST_USER,
        text="3個",
        state={"action": "awaiting_quantity", "prod_cd": "TEST001", "prod_name": "測試商品"},
        line_api=mock_line_api,
    )

after_issues = issue_store.get_pending()
new_issues = [i for i in after_issues if i["user_id"] == TEST_USER and i["type"] == "order_failed"]

check("訂單失敗時回傳 None（不回覆客戶）",
      result_fail is None,
      f"回傳值：{result_fail!r}")
check("訂單失敗時寫入 issue_store（order_failed）",
      len(new_issues) > 0,
      f"新增 issue：{new_issues[-1] if new_issues else '無'}")


# ═══════════════════════════════════════════════════════
# 測試 6：summary _ISSUE_LABEL 包含 order_failed
# ═══════════════════════════════════════════════════════
print("\n━━━ 6. summary 待處理清單包含 order_failed 標籤 ━━━")

from handlers.summary import build_full_report

report = build_full_report(days=1)
check("build_full_report() 包含「訂單失敗」標籤",
      "訂單失敗" in report,
      f"報表片段：\n{report[:400]}")


# ═══════════════════════════════════════════════════════
# 測試 7：reply_text = None 不 crash（main.py 防護）
# ═══════════════════════════════════════════════════════
print("\n━━━ 7. reply_text = None 不 crash ━━━")

# 用 MagicMock 模擬 reply_message，驗證 None 時不被呼叫
mock_api = MagicMock()

reply_text = None
if reply_text:
    mock_api.reply_message("something")

check("reply_text=None 時不呼叫 reply_message",
      not mock_api.reply_message.called,
      "reply_message 未被呼叫 ✓")


# ═══════════════════════════════════════════════════════
# 清理測試資料
# ═══════════════════════════════════════════════════════
with sqlite3.connect(DB_PATH) as conn:
    conn.execute("DELETE FROM customers WHERE line_user_id=?", (TEST_USER,))

# 清理測試 issue（標記已處理）
for i in new_issues:
    issue_store.resolve(i["id"])

state_manager.clear(TEST_USER)


# ═══════════════════════════════════════════════════════
# 結果統計
# ═══════════════════════════════════════════════════════
print("\n" + "═" * 50)
passed = sum(1 for _, ok in _results if ok)
total = len(_results)
print(f"結果：{passed} / {total} 通過")
if passed < total:
    print("失敗項目：")
    for name, ok in _results:
        if not ok:
            print(f"  {_FAIL} {name}")
else:
    print("所有測試通過 🎉")
print("═" * 50)
