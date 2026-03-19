"""
到貨通知記錄（SQLite）

source:
    customer — 客戶自己說「有貨通知我」→ 20:00 排程自動 push
    staff    — 內部群組代客登記 → 不走排程，由員工手動觸發

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
                    source       TEXT    NOT NULL DEFAULT 'customer',
                    status       TEXT    NOT NULL DEFAULT 'pending',
                    created_at   TEXT    NOT NULL,
                    notified_at  TEXT
                )
            """)
            # 舊資料庫升級：若 source 欄不存在則新增
            try:
                conn.execute("ALTER TABLE notify ADD COLUMN source TEXT NOT NULL DEFAULT 'customer'")
            except Exception:
                pass  # 欄位已存在，略過

    def add(self, user_id: str, prod_code: str, prod_name: str,
            qty_wanted: int = 1, source: str = "customer") -> int:
        """
        登記到貨通知。
        source = 'customer'：客戶自己登記（20:00 排程自動通知）
        source = 'staff'：內部群代客登記（不走排程，手動觸發）
        若同一客戶對同一產品已有 pending 記錄，更新 qty/source 即可，不重複新增。
        回傳 id。
        """
        with sqlite3.connect(DB_PATH) as conn:
            existing = conn.execute(
                "SELECT id FROM notify WHERE user_id=? AND prod_code=? AND status='pending'",
                (user_id, prod_code),
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE notify SET qty_wanted=?, source=?, created_at=? WHERE id=?",
                    (qty_wanted, source, datetime.now().isoformat(), existing[0]),
                )
                return existing[0]
            cur = conn.execute(
                "INSERT INTO notify (user_id, prod_code, prod_name, qty_wanted, source, created_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (user_id, prod_code, prod_name, qty_wanted, source, datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_pending_by_code(self, prod_code: str, source: str = "staff") -> list[dict]:
        """取得特定貨號、特定來源的待通知記錄"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM notify WHERE status='pending' AND prod_code=? AND source=? ORDER BY created_at",
                (prod_code.upper(), source),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_pending(self, source: str = "customer") -> list[dict]:
        """
        取得等待通知的記錄。
        source='customer'：只取客戶自己登記的（排程用）
        source='staff'：只取員工代登記的
        source=None：取全部
        """
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            if source is None:
                rows = conn.execute(
                    "SELECT * FROM notify WHERE status='pending' ORDER BY created_at"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM notify WHERE status='pending' AND source=? ORDER BY created_at",
                    (source,),
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
        """目前等待通知的總筆數（所有來源）"""
        with sqlite3.connect(DB_PATH) as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM notify WHERE status='pending'"
            ).fetchone()[0]


notify_store = NotifyStore()
