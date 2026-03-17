"""
客戶到店預約記錄（data/visits.db）

客戶說「下星期去拿」「明天過去」等 → 自動記錄預計到店日期
"""

import sqlite3
from pathlib import Path

DB_PATH = Path("data/visits.db")


def _conn():
    DB_PATH.parent.mkdir(exist_ok=True)
    c = sqlite3.connect(str(DB_PATH))
    c.row_factory = sqlite3.Row
    return c


def init():
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS customer_visits (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id       TEXT NOT NULL,
                display_name  TEXT,
                visit_text    TEXT,
                visit_date    TEXT,
                visit_note    TEXT,
                status        TEXT DEFAULT 'pending',
                created_at    TEXT DEFAULT (datetime('now','localtime'))
            )
        """)


def add(user_id: str, display_name: str, visit_text: str,
        visit_date: str | None, visit_note: str) -> int:
    with _conn() as c:
        cur = c.execute("""
            INSERT INTO customer_visits
                (user_id, display_name, visit_text, visit_date, visit_note)
            VALUES (?, ?, ?, ?, ?)
        """, (user_id, display_name, visit_text, visit_date, visit_note))
        return cur.lastrowid


def get_pending() -> list[dict]:
    """取得所有未到店的記錄，依日期排序（無日期排最後）"""
    with _conn() as c:
        rows = c.execute("""
            SELECT * FROM customer_visits
            WHERE status = 'pending'
            ORDER BY
                CASE WHEN visit_date IS NULL THEN 1 ELSE 0 END,
                visit_date ASC,
                created_at DESC
        """).fetchall()
        return [dict(r) for r in rows]


def mark_visited(visit_id: int) -> bool:
    with _conn() as c:
        c.execute(
            "UPDATE customer_visits SET status='visited' WHERE id=?",
            (visit_id,)
        )
        return c.total_changes > 0


def get_recent_visited(days: int = 7) -> list[dict]:
    with _conn() as c:
        rows = c.execute("""
            SELECT * FROM customer_visits
            WHERE status = 'visited'
              AND created_at >= datetime('now', ?, 'localtime')
            ORDER BY created_at DESC
        """, (f"-{days} days",)).fetchall()
        return [dict(r) for r in rows]
