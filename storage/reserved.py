"""
訂單保留庫存追蹤（SQLite）

當透過 bot 建立訂貨單時，記錄保留數量。
查詢庫存時：可售數量 = Ecount庫存 - 本地未出貨保留數量

staff 出貨後可透過 release() 或 release_by_slip() 釋放保留。
超過 30 天的未釋放保留自動清除（avoid stale data）。
"""

import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "pending_queries.db"


class ReserveStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS order_reserves (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    slip_no     TEXT,           -- Ecount 訂貨單號
                    prod_cd     TEXT NOT NULL,  -- 品項編號
                    qty         INTEGER NOT NULL,
                    cust_name   TEXT,           -- 客戶名稱（可選）
                    created_at  TEXT NOT NULL,
                    released    INTEGER NOT NULL DEFAULT 0,
                    released_at TEXT
                )
            """)

    def reserve(self, prod_cd: str, qty: int,
                slip_no: str = "", cust_name: str = "") -> int:
        """新增保留記錄，回傳 id"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                """INSERT INTO order_reserves
                   (slip_no, prod_cd, qty, cust_name, created_at)
                   VALUES (?, ?, ?, ?, ?)""",
                (slip_no, prod_cd.upper(), qty, cust_name,
                 datetime.now().isoformat()),
            )
            return cur.lastrowid

    def get_reserved_qty(self, prod_cd: str) -> int:
        """取得某品項目前保留的總數量（未釋放、30天內）"""
        cutoff = (datetime.now() - timedelta(days=30)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                """SELECT COALESCE(SUM(qty), 0)
                   FROM order_reserves
                   WHERE prod_cd = ?
                     AND released = 0
                     AND created_at > ?""",
                (prod_cd.upper(), cutoff),
            ).fetchone()
        return int(row[0]) if row else 0

    def release(self, reserve_id: int):
        """手動釋放保留（出貨後呼叫）"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                """UPDATE order_reserves
                   SET released = 1, released_at = ?
                   WHERE id = ?""",
                (datetime.now().isoformat(), reserve_id),
            )

    def release_by_slip(self, slip_no: str):
        """依訂貨單號釋放保留"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                """UPDATE order_reserves
                   SET released = 1, released_at = ?
                   WHERE slip_no = ? AND released = 0""",
                (datetime.now().isoformat(), slip_no),
            )

    def list_active(self) -> list[dict]:
        """列出所有未釋放的保留記錄"""
        cutoff = (datetime.now() - timedelta(days=30)).isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """SELECT * FROM order_reserves
                   WHERE released = 0 AND created_at > ?
                   ORDER BY created_at DESC""",
                (cutoff,),
            ).fetchall()
        return [dict(r) for r in rows]


reserve_store = ReserveStore()
