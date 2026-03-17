"""
到貨通知記錄（SQLite）

當客戶因等待時間過久取消叫貨，
自動登記「有貨時通知我」，到貨後 push 一次訊息。

status:
    pending   — 等待到貨，尚未通知
    notified  — 已推送到貨通知
    cancelled — 已手動取消（admin 移除）
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path("data/notify_requests.db")


class NotifyStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS notify (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id      TEXT    NOT NULL,
                    prod_code    TEXT    NOT NULL,
                    prod_name    TEXT    NOT NULL,
                    qty_wanted   INTEGER NOT NULL DEFAULT 1,
                    status       TEXT    NOT NULL DEFAULT 'pending',
                    created_at   TEXT    NOT NULL,
                    notified_at  TEXT
                )
            """)

    def add(self, user_id: str, prod_code: str, prod_name: str, qty_wanted: int = 1) -> int:
        """
        登記到貨通知。若同一客戶對同一產品已有 pending 記錄，更新 qty 即可，不重複新增。
        回傳 id。
        """
        with sqlite3.connect(DB_PATH) as conn:
            # 已有 pending → 更新數量
            existing = conn.execute(
                "SELECT id FROM notify WHERE user_id=? AND prod_code=? AND status='pending'",
                (user_id, prod_code),
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE notify SET qty_wanted=?, created_at=? WHERE id=?",
                    (qty_wanted, datetime.now().isoformat(), existing[0]),
                )
                return existing[0]
            # 新增
            cur = conn.execute(
                "INSERT INTO notify (user_id, prod_code, prod_name, qty_wanted, created_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (user_id, prod_code, prod_name, qty_wanted, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_pending(self) -> list[dict]:
        """取得所有等待通知的記錄"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM notify WHERE status='pending' ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def mark_notified(self, notify_id: int) -> bool:
        """標記為已通知"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE notify SET status='notified', notified_at=? WHERE id=?",
                (datetime.now().isoformat(), notify_id),
            )
            return cur.rowcount > 0

    def cancel(self, notify_id: int) -> bool:
        """手動取消通知（admin 用）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE notify SET status='cancelled' WHERE id=?",
                (notify_id,),
            )
            return cur.rowcount > 0

    def count_pending(self) -> int:
        """目前等待通知的總筆數"""
        with sqlite3.connect(DB_PATH) as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM notify WHERE status='pending'"
            ).fetchone()[0]


notify_store = NotifyStore()
