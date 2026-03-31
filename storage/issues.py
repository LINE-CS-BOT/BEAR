"""
客服問題記錄（SQLite）

統一存放三類需要人工處理的問題：
  return         — 退換貨申請
  complaint      — 投訴 / 商品問題
  address_change — 地址 / 收件人更改

內部人員確認後在群組輸入「✅ I{id}」標記完成。

status:
    pending  — 等待人工處理
    resolved — 已由人員確認完成
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "issues.db"


class IssueStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS issues (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    type        TEXT    NOT NULL,
                    user_id     TEXT    NOT NULL,
                    text_raw    TEXT    NOT NULL,
                    status      TEXT    NOT NULL DEFAULT 'pending',
                    created_at  TEXT    NOT NULL,
                    resolved_at TEXT
                )
            """)

    def add(self, user_id: str, issue_type: str, text_raw: str) -> int:
        """新增問題記錄，回傳 ID。issue_type: 'return' / 'complaint' / 'address_change'"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "INSERT INTO issues (type, user_id, text_raw, created_at) VALUES (?, ?, ?, ?)",
                (issue_type, user_id, text_raw, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_pending(self) -> list[dict]:
        """取得所有待處理問題"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM issues WHERE status = 'pending' ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def resolve(self, issue_id: int) -> bool:
        """標記為已處理，回傳是否成功"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE issues SET status = 'resolved', resolved_at = ? "
                "WHERE id = ? AND status = 'pending'",
                (datetime.now().isoformat(), issue_id),
            )
            return cur.rowcount > 0

    # 這些類型的待處理不會觸發 bot 靜默（客戶仍可正常互動）
    _NO_SILENCE_TYPES = {"restock_inquiry", "payment_screenshot", "claude_unsure"}

    def has_pending_issue(self, user_id: str) -> bool:
        """檢查該客戶是否有需要靜默的未處理問題（restock_inquiry 不算）"""
        placeholders = ",".join("?" for _ in self._NO_SILENCE_TYPES)
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                f"SELECT 1 FROM issues WHERE user_id=? AND status='pending'"
                f" AND type NOT IN ({placeholders}) LIMIT 1",
                (user_id, *self._NO_SILENCE_TYPES),
            ).fetchone()
        return row is not None

    def get_pending_for_user(self, user_id: str) -> list[dict]:
        """取得該客戶的所有未處理問題"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM issues WHERE user_id=? AND status='pending' ORDER BY created_at",
                (user_id,),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_recent_resolved(self, days: int = 3) -> list[dict]:
        """取得近 N 天已處理的問題"""
        from datetime import timedelta
        cutoff = (datetime.now() - timedelta(days=days)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM issues WHERE status = 'resolved' AND created_at >= ? "
                "ORDER BY resolved_at DESC",
                (cutoff,),
            ).fetchall()
        return [dict(r) for r in rows]


issue_store = IssueStore()
