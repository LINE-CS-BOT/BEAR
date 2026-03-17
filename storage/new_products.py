"""
待審核新品項記錄（SQLite）

由內部群組「新增品項」指令觸發後，先存入此表等待人工在 admin 介面確認。
確認後記錄保留（status='confirmed'），不再顯示在待審核清單。
"""

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path("data/new_products.db")


class NewProductStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS new_products (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    prod_cd      TEXT NOT NULL,
                    prod_name    TEXT NOT NULL,
                    unit         TEXT DEFAULT '個',
                    bar_code     TEXT DEFAULT '',
                    class_cd     TEXT DEFAULT '',
                    out_price    TEXT DEFAULT '',
                    in_price     TEXT DEFAULT '',
                    size_des     TEXT DEFAULT '',
                    cust         TEXT DEFAULT '10003',
                    status       TEXT DEFAULT 'pending',
                    created_at   TEXT NOT NULL,
                    confirmed_at TEXT
                )
            """)

    def add(
        self,
        prod_cd:   str,
        prod_name: str,
        unit:      str = "個",
        bar_code:  str = "",
        class_cd:  str = "",
        out_price: str = "",
        in_price:  str = "",
        size_des:  str = "",
        cust:      str = "10003",
    ) -> int:
        """新增待審核品項，回傳 ID。"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                """INSERT INTO new_products
                   (prod_cd, prod_name, unit, bar_code, class_cd,
                    out_price, in_price, size_des, cust, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    prod_cd, prod_name, unit, bar_code, class_cd,
                    out_price, in_price, size_des, cust,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            return cur.lastrowid

    def get_pending(self) -> list[dict]:
        """取得所有 status='pending' 的品項。"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM new_products WHERE status='pending' ORDER BY created_at DESC"
            ).fetchall()
        return [dict(r) for r in rows]

    def confirm(self, item_id: int) -> bool:
        """人工確認，標記為 confirmed。"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE new_products SET status='confirmed', confirmed_at=? WHERE id=?",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), item_id),
            )
            return cur.rowcount > 0

    def delete(self, item_id: int) -> bool:
        """直接刪除（棄用）。"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute("DELETE FROM new_products WHERE id=?", (item_id,))
            return cur.rowcount > 0


new_products_store = NewProductStore()
