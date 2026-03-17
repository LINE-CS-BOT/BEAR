"""
規格資料庫（data/specs.json）

由 scripts/import_specs.py 從產品PO文.txt 生成。
若要更新規格資料，請修改PO文.txt後重新執行 import_specs.py。

查詢方式：
  get_by_code("T1202")             → 依產品編號（完全符合）
  get_by_name("摩托車")            → 依品名關鍵字（模糊）
  get_by_machine("標準台")         → 列出適用某台型的所有產品

回傳格式：
  {
    "code":    "T1202",
    "name":    "杜卡迪合金回力摩托車",
    "size":    "18X9X9公分",
    "weight":  "237公克",
    "machine": ["標準台"],
    "price":   "109元",
  }
"""

import json
from pathlib import Path

SPECS_PATH = Path("data/specs.json")

_cache: dict = {}


def _load() -> dict:
    global _cache
    if not _cache and SPECS_PATH.exists():
        try:
            _cache = json.loads(SPECS_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"[specs] 載入失敗: {e}")
    return _cache


def get_by_code(code: str) -> dict | None:
    """以產品編號查詢規格（完全符合）"""
    return _load().get(code.upper())


def get_by_name(keyword: str) -> dict | None:
    """以關鍵字模糊搜尋產品名稱，回傳第一筆符合的記錄"""
    kw = keyword.strip().upper()
    if not kw:
        return None
    for s in _load().values():
        if kw in s.get("name", "").upper():
            return s
    return None


def get_by_machine(machine_type: str) -> list[dict]:
    """列出所有適用指定台型的產品"""
    result = []
    for s in _load().values():
        if any(machine_type in m for m in s.get("machine", [])):
            result.append(s)
    return result


def get_by_size(keyword: str) -> list[dict]:
    """
    以尺寸關鍵字搜尋，回傳符合的產品列表。
    支援：
      - 「28公分」→ 提取 28，比對尺寸欄位中有 28 這個數字（邊界比對，不匹配 128/280）
      - 「28」    → 同上
    """
    import re as _re
    kw = keyword.strip()
    if not kw:
        return []
    # 提取數字部分（如「28公分」→「28」，「28.5公分」→「28.5」）
    m = _re.search(r'\d+(?:\.\d+)?', kw)
    num = m.group(0) if m else kw
    # 用邊界比對，避免 「28」 匹配到 「128」
    pattern = _re.compile(r'(?<![0-9])' + _re.escape(num) + r'(?![0-9])')
    result = []
    for s in _load().values():
        if pattern.search(s.get("size", "")):
            result.append(s)
    return result


def get_all() -> dict:
    """回傳全部規格資料 dict（code → info）"""
    return _load()


def reload():
    """強制重新載入（更新PO文後使用）"""
    global _cache
    _cache = {}
    _load()
    print(f"[specs] 已重新載入，共 {len(_cache)} 筆")
