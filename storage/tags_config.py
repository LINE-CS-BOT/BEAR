"""
客戶分類標籤設定（data/tags_config.json）
預設：["VIP", "野獸國", "標準", "中句", "K霸"]

使用方式：
  from storage.tags_config import load_tags, add_tag, remove_tag
"""
import json
from pathlib import Path

_CONFIG_PATH = Path(__file__).parent.parent / "data" / "tags_config.json"
_DEFAULT_TAGS = ["VIP", "野獸國", "標準", "中句", "K霸"]


def load_tags() -> list[str]:
    """載入標籤清單（若無設定檔則回傳預設值）"""
    if _CONFIG_PATH.exists():
        try:
            data = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return [str(t) for t in data if t]
        except Exception as e:
            print(f"[tags_config] 載入失敗，使用預設值: {e}")
    return list(_DEFAULT_TAGS)


def save_tags(tags: list[str]) -> None:
    """儲存標籤清單"""
    _CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    _CONFIG_PATH.write_text(
        json.dumps(tags, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"[tags_config] 已儲存 {len(tags)} 個標籤: {tags}")


def add_tag(tag: str) -> list[str]:
    """新增標籤（若已存在則忽略），回傳最新清單"""
    tag = tag.strip()
    if not tag:
        raise ValueError("標籤名稱不可為空")
    tags = load_tags()
    if tag not in tags:
        tags.append(tag)
        save_tags(tags)
    return tags


def remove_tag(tag: str) -> list[str]:
    """移除標籤，回傳最新清單"""
    tags = load_tags()
    if tag in tags:
        tags.remove(tag)
        save_tags(tags)
    return tags
