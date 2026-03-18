"""
對話狀態管理（in-memory + SQLite 持久化）

用途：支援多輪對話
例：客戶說「有貨嗎」但沒說產品，Bot 問「請問哪個產品？」
    下一則訊息就能知道是在回答庫存問題。

客戶等待中的狀態（awaiting_*）會持久化到 SQLite，
重啟 server 後可自動恢復，並支援 24h 提醒 / 48h 清除。
"""

import threading
from datetime import datetime, timedelta


# 需要持久化的長效狀態（客戶等待回覆）
_LONG_LIVED_TTL = timedelta(hours=72)   # 記憶體 TTL：72小時（SQLite 會在 48h 清除）


class StateManager:
    def __init__(self, ttl_minutes: int = 10):
        self._store: dict[str, dict] = {}
        self._group_pref: dict[str, str] = {}  # user_id → preferred Ecount cust_cd（群組訂單用）
        self.ttl = timedelta(minutes=ttl_minutes)
        self._lock = threading.Lock()

    def _is_persistent_action(self, action: str) -> bool:
        from storage.persistent_state import PERSISTENT_ACTIONS
        return action in PERSISTENT_ACTIONS

    def set(self, user_id: str, state: dict):
        """設定對話狀態，TTL 內有效；長效狀態同時持久化到 SQLite"""
        action = state.get("action", "")
        ttl = _LONG_LIVED_TTL if self._is_persistent_action(action) else self.ttl
        with self._lock:
            self._store[user_id] = {
                **state,
                "_expires_at": datetime.now() + ttl,
            }
        # 持久化到 SQLite（不在 lock 內避免死鎖）
        if self._is_persistent_action(action):
            try:
                from storage.persistent_state import persistent_state_store
                persistent_state_store.save(user_id, state)
            except Exception as e:
                print(f"[state] 持久化失敗: {e}")

    def get(self, user_id: str) -> dict | None:
        """取得狀態，過期自動清除"""
        with self._lock:
            entry = self._store.get(user_id)
            if not entry:
                return None
            if datetime.now() > entry["_expires_at"]:
                del self._store[user_id]
                return None
            return {k: v for k, v in entry.items() if k != "_expires_at"}

    def append_upload_media(self, user_id: str, item: dict) -> bool:
        """
        原子性地將媒體項目追加到 uploading session 的 current_media。
        回傳 True 表示成功（state 存在且是 uploading），False 表示無效。
        """
        with self._lock:
            entry = self._store.get(user_id)
            if not entry or entry.get("action") != "uploading":
                return False
            if datetime.now() > entry["_expires_at"]:
                del self._store[user_id]
                return False
            entry.setdefault("current_media", []).append(item)
            # 每次加入媒體就重置 TTL
            entry["_expires_at"] = datetime.now() + self.ttl
            return True

    def clear(self, user_id: str):
        """清除狀態（同時清除 SQLite 持久化）"""
        with self._lock:
            self._store.pop(user_id, None)
        try:
            from storage.persistent_state import persistent_state_store
            persistent_state_store.delete(user_id)
        except Exception:
            pass

    def restore_from_db(self):
        """啟動時從 SQLite 恢復所有持久化狀態到記憶體"""
        try:
            from storage.persistent_state import persistent_state_store
            items = persistent_state_store.load_all()
            restored = 0
            with self._lock:
                for user_id, state in items:
                    self._store[user_id] = {
                        **state,
                        "_expires_at": datetime.now() + _LONG_LIVED_TTL,
                    }
                    restored += 1
            if restored:
                print(f"[state] 從 SQLite 恢復 {restored} 筆對話狀態")
        except Exception as e:
            print(f"[state] 狀態恢復失敗: {e}")

    # ── 群組預設地址（獨立於 TTL 狀態，訂單完成後手動清除）──────────────
    def set_group_cust_cd(self, user_id: str, cust_cd: str) -> None:
        """記錄此用戶的群組預設 Ecount 代碼"""
        self._group_pref[user_id] = cust_cd

    def get_group_cust_cd(self, user_id: str) -> str | None:
        """取得此用戶的群組預設 Ecount 代碼"""
        return self._group_pref.get(user_id)

    def clear_group_cust_cd(self, user_id: str) -> None:
        """清除群組預設代碼（訂單完成或取消後呼叫）"""
        self._group_pref.pop(user_id, None)

    def all_states(self) -> dict[str, dict]:
        """回傳所有尚未過期的狀態（供接手面板列舉用）"""
        now = datetime.now()
        with self._lock:
            return {
                uid: {k: v for k, v in st.items() if k != "_expires_at"}
                for uid, st in self._store.items()
                if st.get("_expires_at", now) > now
            }


state_manager = StateManager()
