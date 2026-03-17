"""
購物車（in-memory，per user）

每位客戶獨立購物車，結帳後清空。
items 格式：[{"prod_cd": str, "prod_name": str, "qty": int}, ...]
"""

from threading import Lock

_carts: dict[str, list[dict]] = {}
_lock = Lock()


def add_item(user_id: str, prod_cd: str, prod_name: str, qty: int) -> list[dict]:
    """加入品項（相同 prod_cd 則累加數量），回傳目前購物車"""
    prod_cd = prod_cd.upper()   # 統一大寫，避免 k0216 vs K0216 重複加入
    with _lock:
        cart = _carts.setdefault(user_id, [])
        for item in cart:
            if item["prod_cd"].upper() == prod_cd:
                item["qty"] += qty
                return list(cart)
        cart.append({"prod_cd": prod_cd, "prod_name": prod_name, "qty": qty})
        return list(cart)


def get_cart(user_id: str) -> list[dict]:
    """取得目前購物車（空時回傳 []）"""
    with _lock:
        return list(_carts.get(user_id, []))


def clear_cart(user_id: str) -> None:
    """清空購物車"""
    with _lock:
        _carts.pop(user_id, None)


def is_empty(user_id: str) -> bool:
    with _lock:
        return len(_carts.get(user_id, [])) == 0
