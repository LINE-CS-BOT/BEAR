"""
轉帳確認記錄（SQLite）

當客戶傳來轉帳確認訊息時記錄在這裡。
內部人員確認後可在群組輸入「✅ P{id}」標記完成。

status:
    pending  — 等待人工確認
    resolved — 已確認完成
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "payment_confirmations.db"


class PaymentStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS payments (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id     TEXT    NOT NULL,
                    text_raw    TEXT    NOT NULL,
                    status      TEXT    NOT NULL DEFAULT 'pending',
                    created_at  TEXT    NOT NULL,
                    resolved_at TEXT
                )
            """)

    def add(self, user_id: str, text_raw: str) -> int:
        """新增轉帳確認記錄，回傳 ID"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "INSERT INTO payments (user_id, text_raw, created_at) VALUES (?, ?, ?)",
                (user_id, text_raw, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_pending(self) -> list[dict]:
        """取得所有待確認的轉帳記錄"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM payments WHERE status = 'pending' ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def resolve(self, payment_id: int) -> bool:
        """標記為已確認，回傳是否成功"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE payments SET status = 'resolved', resolved_at = ? "
                "WHERE id = ? AND status = 'pending'",
                (datetime.now().isoformat(), payment_id),
            )
            return cur.rowcount > 0

    def get_recent_resolved(self, days: int = 3) -> list[dict]:
        """取得近 N 天已確認的轉帳記錄"""
        from datetime import timedelta
        cutoff = (datetime.now() - timedelta(days=days)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM payments WHERE status = 'resolved' AND created_at >= ? "
                "ORDER BY resolved_at DESC",
                (cutoff,),
            ).fetchall()
        return [dict(r) for r in rows]


payment_store = PaymentStore()
