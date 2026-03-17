"""
離峰訊息佇列

收集時段：00:30 ~ 11:30
補處理時間：12:45（自動排程）

欄位：
  id         INTEGER PK
  user_id    TEXT
  msg_type   TEXT   'text' / 'image'
  content    TEXT   文字訊息內容（text 類型用）
  msg_id     TEXT   LINE message ID（image 類型用，下載圖片用）
  processed  INTEGER  0 = 待處理 / 1 = 已處理
  created_at TEXT   ISO 格式
"""

import sqlite3
from datetime import datetime
from pathlib import Path

_DB = Path(__file__).parent.parent / "data" / "message_queue.db"


def _conn():
    c = sqlite3.connect(str(_DB))
    c.row_factory = sqlite3.Row
    return c


def init():
    """建立 table（若不存在）"""
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS message_queue (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    TEXT    NOT NULL,
                msg_type   TEXT    NOT NULL,
                content    TEXT    DEFAULT '',
                msg_id     TEXT    DEFAULT '',
                processed  INTEGER DEFAULT 0,
                created_at TEXT    NOT NULL
            )
        """)


def add(user_id: str, msg_type: str, content: str = "", msg_id: str = "") -> int:
    """新增一則離峰訊息，回傳 id"""
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO message_queue (user_id, msg_type, content, msg_id, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (user_id, msg_type, content, msg_id, datetime.now().isoformat()),
        )
        return cur.lastrowid


def get_unprocessed() -> list[dict]:
    """取出所有未處理的訊息（依時間排序）"""
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM message_queue WHERE processed = 0 ORDER BY created_at"
        ).fetchall()
        return [dict(r) for r in rows]


def mark_processed(row_id: int):
    """標記某則訊息為已處理"""
    with _conn() as c:
        c.execute("UPDATE message_queue SET processed = 1 WHERE id = ?", (row_id,))


def count_unprocessed() -> int:
    """查詢待處理數量"""
    with _conn() as c:
        return c.execute(
            "SELECT COUNT(*) FROM message_queue WHERE processed = 0"
        ).fetchone()[0]
