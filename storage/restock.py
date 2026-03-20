"""
調貨請求記錄（SQLite）

當庫存不足客戶詢問數量時，記錄在這裡。
HQ 群組回覆後依 status 更新流程。

status:
    pending    — 已詢問 HQ，等待回覆
    available  — HQ 表示有貨可調（訂單建立中）
    ordering   — HQ 表示需要叫貨，等待客戶確認是否能等
    confirmed  — 客戶確認等待，訂單已建立
    cancelled  — 客戶取消
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "restock_requests.db"


class RestockStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS restock (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id     TEXT    NOT NULL,
                    prod_name   TEXT    NOT NULL,
                    prod_cd     TEXT    NOT NULL DEFAULT '',
                    qty         INTEGER NOT NULL,
                    status      TEXT    NOT NULL DEFAULT 'pending',
                    wait_time   TEXT    NOT NULL DEFAULT '',
                    created_at  TEXT    NOT NULL
                )
            """)

    def add(self, user_id: str, prod_name: str, prod_cd: str, qty: int) -> int:
        """新增調貨請求，回傳 ID"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "INSERT INTO restock (user_id, prod_name, prod_cd, qty, created_at) VALUES (?, ?, ?, ?, ?)",
                (user_id, prod_name, prod_cd, qty, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_latest_pending(self) -> dict | None:
        """取得最新一筆 pending 調貨請求"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT * FROM restock WHERE status = 'pending' ORDER BY created_at DESC LIMIT 1"
            ).fetchone()
        return dict(row) if row else None

    def find_pending_by_product(self, keyword: str) -> dict | None:
        """模糊比對產品名/編號，找最新的 pending 請求"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT * FROM restock WHERE status = 'pending' "
                "AND (prod_name LIKE ? OR prod_cd LIKE ?) "
                "ORDER BY created_at DESC LIMIT 1",
                (f"%{keyword}%", f"%{keyword}%"),
            ).fetchone()
        return dict(row) if row else None

    def get_unresolved(self) -> list[dict]:
        """取得所有尚未完成的調貨請求（pending / ordering）"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM restock WHERE status IN ('pending', 'ordering') ORDER BY created_at"
            ).fetchall()
        return [dict(r) for r in rows]

    def update_status(self, restock_id: int, status: str, wait_time: str = "") -> bool:
        """更新狀態，回傳是否有實際變更（True = 找到並更新）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE restock SET status = ?, wait_time = ? WHERE id = ?",
                (status, wait_time, restock_id),
            )
            return cur.rowcount > 0

    def get_recent_completed(self, days: int = 3) -> list[dict]:
        """取得近 N 天已完成（confirmed / cancelled）的調貨記錄"""
        from datetime import timedelta
        cutoff = (datetime.now() - timedelta(days=days)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM restock WHERE status IN ('confirmed', 'cancelled') "
                "AND created_at >= ? ORDER BY created_at DESC",
                (cutoff,),
            ).fetchall()
        return [dict(r) for r in rows]


restock_store = RestockStore()
