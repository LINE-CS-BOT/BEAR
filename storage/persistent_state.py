"""
持久化對話狀態 SQLite 後端

只保存客戶端「等待回覆」的長效狀態，重啟後可恢復。
支援 24小時提醒 / 48小時自動清除。
"""

import json
import sqlite3
import threading
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "conversation_states.db"

# 需要持久化的狀態（客戶等待中）
PERSISTENT_ACTIONS = {
    "awaiting_quantity",
    "awaiting_restock_qty",
    "awaiting_wait_confirm",
    "awaiting_order_confirm",
    "awaiting_group_address_confirm",
    "awaiting_order_id",
    "awaiting_product_clarify",
    "human_takeover",
    "pending_add_img",
    "pending_save_img",
}

REMIND_AFTER_HOURS  = 24   # 超過 N 小時沒回 → 提醒
EXPIRE_AFTER_HOURS  = 48   # 超過 N 小時沒回 → 清除


class PersistentStateStore:
    def __init__(self):
        self._lock = threading.Lock()
        self._init_db()

    def _conn(self):
        return sqlite3.connect(DB_PATH)

    def _init_db(self):
        with self._conn() as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS conv_states (
                    user_id         TEXT PRIMARY KEY,
                    action          TEXT NOT NULL,
                    state_json      TEXT NOT NULL,
                    created_at      TEXT NOT NULL,
                    updated_at      TEXT NOT NULL,
                    last_reminded_at TEXT
                )
            """)

    # ── 寫入 ──────────────────────────────────────────────
    def save(self, user_id: str, state: dict):
        now = datetime.now().isoformat()
        action = state.get("action", "")
        state_json = json.dumps(state, ensure_ascii=False)
        with self._lock:
            with self._conn() as conn:
                # 如果已存在，保留 created_at；否則用 now
                existing = conn.execute(
                    "SELECT created_at FROM conv_states WHERE user_id=?", (user_id,)
                ).fetchone()
                created_at = existing[0] if existing else now
                conn.execute("""
                    INSERT OR REPLACE INTO conv_states
                    (user_id, action, state_json, created_at, updated_at, last_reminded_at)
                    VALUES (?, ?, ?, ?, ?, NULL)
                """, (user_id, action, state_json, created_at, now))

    def delete(self, user_id: str):
        with self._lock:
            with self._conn() as conn:
                conn.execute("DELETE FROM conv_states WHERE user_id=?", (user_id,))

    def load_all(self) -> list[tuple[str, dict]]:
        """啟動時載入全部狀態"""
        with self._lock:
            with self._conn() as conn:
                rows = conn.execute(
                    "SELECT user_id, state_json FROM conv_states"
                ).fetchall()
                return [(r[0], json.loads(r[1])) for r in rows]

    # ── 提醒 / 過期查詢 ───────────────────────────────────
    def get_need_remind(self) -> list[dict]:
        """取得超過 24h 且尚未提醒過（或上次提醒超過 24h）的狀態"""
        cutoff = (datetime.now() - timedelta(hours=REMIND_AFTER_HOURS)).isoformat()
        with self._lock:
            with self._conn() as conn:
                rows = conn.execute("""
                    SELECT user_id, action, state_json, updated_at, last_reminded_at
                    FROM conv_states
                    WHERE updated_at < ?
                      AND (last_reminded_at IS NULL OR last_reminded_at < ?)
                """, (cutoff, cutoff)).fetchall()
                return [
                    {
                        "user_id":          r[0],
                        "action":           r[1],
                        "state":            json.loads(r[2]),
                        "updated_at":       r[3],
                        "last_reminded_at": r[4],
                    }
                    for r in rows
                ]

    def get_expired(self) -> list[str]:
        """取得超過 48h 的 user_id 列表"""
        cutoff = (datetime.now() - timedelta(hours=EXPIRE_AFTER_HOURS)).isoformat()
        with self._lock:
            with self._conn() as conn:
                rows = conn.execute(
                    "SELECT user_id FROM conv_states WHERE updated_at < ?", (cutoff,)
                ).fetchall()
                return [r[0] for r in rows]

    def mark_reminded(self, user_id: str):
        now = datetime.now().isoformat()
        with self._lock:
            with self._conn() as conn:
                conn.execute(
                    "UPDATE conv_states SET last_reminded_at=? WHERE user_id=?",
                    (now, user_id)
                )


persistent_state_store = PersistentStateStore()
