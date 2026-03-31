"""
購物車（in-memory + JSON 持久化）

每位客戶獨立購物車，結帳後清空。
重啟時自動從 data/carts.json 恢復。
items 格式：[{"prod_cd": str, "prod_name": str, "qty": int}, ...]
"""

import json
import time
from pathlib import Path
from threading import Lock

_PERSIST_PATH = Path(__file__).parent.parent / "data" / "carts.json"
_carts: dict[str, list[dict]] = {}
_cart_timestamps: dict[str, float] = {}
_lock = Lock()


def _save() -> None:
    """寫入 JSON（需在 _lock 內呼叫）"""
    try:
        data = {
            uid: {"items": items, "ts": _cart_timestamps.get(uid, time.time())}
            for uid, items in _carts.items() if items
        }
        _PERSIST_PATH.parent.mkdir(parents=True, exist_ok=True)
        _PERSIST_PATH.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        print(f"[cart] 持久化失敗: {e}", flush=True)


def _load() -> None:
    """啟動時從 JSON 恢復"""
    global _carts, _cart_timestamps
    if not _PERSIST_PATH.exists():
        return
    try:
        data = json.loads(_PERSIST_PATH.read_text(encoding="utf-8"))
        cutoff = time.time() - 48 * 3600  # 超過 48 小時的不恢復
        for uid, entry in data.items():
            ts = entry.get("ts", 0)
            if ts > cutoff:
                _carts[uid] = entry.get("items", [])
                _cart_timestamps[uid] = ts
        if _carts:
            print(f"[cart] 恢復 {len(_carts)} 個購物車", flush=True)
    except Exception as e:
        print(f"[cart] 恢復失敗: {e}", flush=True)


# 啟動時自動恢復
_load()


def add_item(user_id: str, prod_cd: str, prod_name: str, qty: int) -> list[dict]:
    """加入品項（相同 prod_cd 則累加數量），回傳目前購物車"""
    prod_cd = prod_cd.upper()
    with _lock:
        cart = _carts.setdefault(user_id, [])
        for item in cart:
            if item["prod_cd"].upper() == prod_cd:
                item["qty"] += qty
                _cart_timestamps[user_id] = time.time()
                _save()
                return list(cart)
        cart.append({"prod_cd": prod_cd, "prod_name": prod_name, "qty": qty})
        _cart_timestamps[user_id] = time.time()
        _save()
        return list(cart)


def get_cart(user_id: str) -> list[dict]:
    """取得目前購物車（空時回傳 []）"""
    with _lock:
        return list(_carts.get(user_id, []))


def clear_cart(user_id: str) -> None:
    """清空購物車"""
    with _lock:
        _carts.pop(user_id, None)
        _cart_timestamps.pop(user_id, None)
        _save()


def is_empty(user_id: str) -> bool:
    with _lock:
        return len(_carts.get(user_id, [])) == 0


def cleanup_expired(max_age_hours: int = 48) -> int:
    """移除超過 max_age_hours 未修改的購物車，回傳清除數量"""
    cutoff = time.time() - max_age_hours * 3600
    removed = 0
    with _lock:
        expired_users = [
            uid for uid, ts in _cart_timestamps.items() if ts < cutoff
        ]
        for uid in expired_users:
            _carts.pop(uid, None)
            _cart_timestamps.pop(uid, None)
            removed += 1
        if removed:
            _save()
    return removed
