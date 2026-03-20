"""
待確認庫存查詢記錄（SQLite）

當客戶詢問庫存不足的產品時，記錄在這裡。
工作人員確認後可查看待回覆清單。
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "pending_queries.db"


class PendingStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS pending (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id     TEXT    NOT NULL,
                    product     TEXT    NOT NULL,
                    created_at  TEXT    NOT NULL,
                    answered    INTEGER NOT NULL DEFAULT 0
                )
            """)

    def add(self, user_id: str, product: str):
        """新增一筆待確認記錄"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                "INSERT INTO pending (user_id, product, created_at) VALUES (?, ?, ?)",
                (user_id, product, datetime.now().isoformat()),
            )

    def get_pending(self) -> list[dict]:
        """取得所有尚未回覆的查詢"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM pending WHERE answered = 0 ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def mark_answered(self, pending_id: int) -> bool:
        """標記為已回覆，回傳是否成功"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE pending SET answered = 1 WHERE id = ? AND answered = 0", (pending_id,)
            )
            return cur.rowcount > 0

    def has_pending(self, user_id: str) -> bool:
        """檢查該客戶是否有未回覆的查無商品記錄（真人介入中，bot 凍結）"""
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                "SELECT 1 FROM pending WHERE user_id=? AND answered=0 LIMIT 1",
                (user_id,),
            ).fetchone()
        return row is not None


pending_store = PendingStore()
