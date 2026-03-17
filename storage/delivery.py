"""
配送詢問記錄（SQLite）

客戶詢問配送時間時記錄在這裡，由內部人員跟司機確認後手動回覆。
內部人員確認後在群組輸入「✅ D{id}」標記完成。

status:
    pending  — 等待人工確認並回覆客戶
    resolved — 已由人員確認完成
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path("data/delivery_inquiries.db")


class DeliveryStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS deliveries (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id     TEXT    NOT NULL,
                    text_raw    TEXT    NOT NULL,
                    status      TEXT    NOT NULL DEFAULT 'pending',
                    created_at  TEXT    NOT NULL,
                    resolved_at TEXT
                )
            """)

    def add(self, user_id: str, text_raw: str) -> int:
        """新增配送詢問記錄，回傳 ID"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "INSERT INTO deliveries (user_id, text_raw, created_at) VALUES (?, ?, ?)",
                (user_id, text_raw, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_pending(self) -> list[dict]:
        """取得所有待確認的配送詢問"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM deliveries WHERE status = 'pending' ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def has_pending(self, user_id: str) -> bool:
        """此客戶是否已有未處理的配送詢問"""
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                "SELECT 1 FROM deliveries WHERE user_id = ? AND status = 'pending' LIMIT 1",
                (user_id,),
            ).fetchone()
        return row is not None

    def resolve(self, delivery_id: int) -> bool:
        """標記為已確認，回傳是否成功"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE deliveries SET status = 'resolved', resolved_at = ? "
                "WHERE id = ? AND status = 'pending'",
                (datetime.now().isoformat(), delivery_id),
            )
            return cur.rowcount > 0

    def get_recent_resolved(self, days: int = 3) -> list[dict]:
        """取得近 N 天已確認的配送詢問"""
        from datetime import timedelta
        cutoff = (datetime.now() - timedelta(days=days)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM deliveries WHERE status = 'resolved' AND created_at >= ? "
                "ORDER BY resolved_at DESC",
                (cutoff,),
            ).fetchall()
        return [dict(r) for r in rows]


delivery_store = DeliveryStore()
