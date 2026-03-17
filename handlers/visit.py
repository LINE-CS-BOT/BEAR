"""
客戶到店預告處理

偵測：「下星期去拿」「明天過去看看」「過幾天去」「有空找時間去」
→ 解析日期 → 記錄到 visits.db → 回覆確認
→ 內部群查詢：「誰要來」「哪些客人要來」→ 列出待到店清單
"""

import re
from datetime import datetime, timedelta

# ── 到店意圖關鍵字 ────────────────────────────────────────────────────
VISIT_KEYWORDS = [
    "去拿", "來拿", "來店", "來看看", "去看看", "去一下", "有空去",
    "找時間去", "去看", "過去", "去取", "自取", "去買", "去那邊",
    "下星期", "下週", "下周", "明天去", "後天去", "這週末", "週末去",
    "明天過去", "後天過去", "過去拿", "過去看",
]

# ── 內部群查詢觸發詞 ──────────────────────────────────────────────────
VISIT_QUERY_KEYWORDS = [
    "誰要來", "要來拿", "哪些客人", "哪些人要來", "客人要來",
    "最近來", "誰來拿", "來拿貨", "預約到店", "到店名單",
    "客人來", "要來的客人", "來的客人", "要來客人",
]

# ── 星期對照 ──────────────────────────────────────────────────────────
_WEEKDAY_MAP = {
    "一": 0, "二": 1, "三": 2, "四": 3, "五": 4, "六": 5,
    "日": 6, "天": 6,
}
_WEEKDAY_LABEL = ["一", "二", "三", "四", "五", "六", "日"]


def is_visit_message(text: str) -> bool:
    return any(kw in text for kw in VISIT_KEYWORDS)


def is_visit_query(text: str) -> bool:
    return any(kw in text for kw in VISIT_QUERY_KEYWORDS)


def parse_visit_date(text: str) -> tuple[str | None, str]:
    """
    解析到店日期。
    回傳 (ISO 日期字串 or None, 人讀說明)
    """
    today = datetime.now()

    # 今天
    if re.search(r"今天|今日", text):
        return today.strftime("%Y-%m-%d"), "今天"

    # 明天
    if "明天" in text:
        d = today + timedelta(days=1)
        return d.strftime("%Y-%m-%d"), "明天"

    # 後天
    if "後天" in text:
        d = today + timedelta(days=2)
        return d.strftime("%Y-%m-%d"), "後天"

    # 下星期X / 下週X（指定星期幾）
    m = re.search(r"下(?:星期|週|周)([一二三四五六日天])", text)
    if m:
        wd = _WEEKDAY_MAP.get(m.group(1), 0)
        days_to_next_monday = (7 - today.weekday()) % 7
        if days_to_next_monday == 0:
            days_to_next_monday = 7
        next_monday = today + timedelta(days=days_to_next_monday)
        target = next_monday + timedelta(days=wd)
        label = f"下星期{_WEEKDAY_LABEL[wd]}"
        return target.strftime("%Y-%m-%d"), label

    # 下星期 / 下週（未指定星期幾 → 取下週一）
    if re.search(r"下(?:星期|週|周)", text):
        days_to_next_monday = (7 - today.weekday()) % 7
        if days_to_next_monday == 0:
            days_to_next_monday = 7
        next_monday = today + timedelta(days=days_to_next_monday)
        return next_monday.strftime("%Y-%m-%d"), "下星期（約）"

    # 這週末 / 週末
    if re.search(r"(?:這|本)?(?:週末|周末)", text):
        days_to_sat = (5 - today.weekday()) % 7
        if days_to_sat == 0:
            days_to_sat = 7
        d = today + timedelta(days=days_to_sat)
        return d.strftime("%Y-%m-%d"), "這週末"

    # X月X號/日
    m = re.search(r"(\d{1,2})月(\d{1,2})[號日]", text)
    if m:
        try:
            month, day = int(m.group(1)), int(m.group(2))
            d = datetime(today.year, month, day)
            if d.date() < today.date():
                d = datetime(today.year + 1, month, day)
            return d.strftime("%Y-%m-%d"), f"{month}月{day}號"
        except Exception:
            pass

    # 過幾天（模糊）
    if re.search(r"過幾天|幾天後|近幾天|這幾天", text):
        return None, "近幾天（未定）"

    # 有空 / 找時間（非常模糊）
    if re.search(r"有空|找時間|空的時候|方便時|有機會", text):
        return None, "有空再去（未定）"

    return None, "未指定時間"


def handle_visit(user_id: str, text: str, display_name: str = "") -> str:
    """
    1:1 客戶說要來店 → 記錄 + 回覆
    """
    from storage import visits as visit_store
    visit_date, visit_note = parse_visit_date(text)
    visit_store.add(
        user_id,
        display_name or user_id[:8],
        text,
        visit_date,
        visit_note,
    )
    if visit_date:
        # 計算幾天後
        try:
            d = datetime.strptime(visit_date, "%Y-%m-%d")
            diff = (d.date() - datetime.now().date()).days
            if diff == 0:
                timing = "今天"
            elif diff == 1:
                timing = "明天"
            else:
                timing = visit_note
        except Exception:
            timing = visit_note
        return f"好的！{timing}見 😊 有需要再告訴我～"
    else:
        return "好的！歡迎來，有任何問題再告訴我 😊"


def handle_visit_query() -> str:
    """
    內部群查詢「誰要來 / 哪些客人要來」→ 列出待到店清單
    """
    from storage import visits as visit_store
    visits = visit_store.get_pending()
    if not visits:
        return "目前沒有客人預告要來店"

    today_str = datetime.now().strftime("%Y-%m-%d")
    lines = [f"📅 預計到店客人（共 {len(visits)} 位）："]
    for v in visits:
        name = v.get("display_name") or "客人"
        note = v.get("visit_note") or "未定"
        date = v.get("visit_date") or ""

        if date:
            try:
                d = datetime.strptime(date, "%Y-%m-%d")
                diff = (d.date() - datetime.now().date()).days
                if diff < 0:
                    date_label = f"{date}（已過期）"
                elif diff == 0:
                    date_label = "今天"
                elif diff == 1:
                    date_label = "明天"
                else:
                    date_label = f"{date}（{note}）"
            except Exception:
                date_label = note
        else:
            date_label = note

        raw = v.get("visit_text") or ""
        short = raw[:20] + ("…" if len(raw) > 20 else "")
        lines.append(f"• #{v['id']} {name}｜{date_label}｜「{short}」")

    lines.append("\n✅ V{id} → 標記已到店")
    return "\n".join(lines)
