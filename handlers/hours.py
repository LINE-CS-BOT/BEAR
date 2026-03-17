import re
from datetime import datetime, timedelta
import pytz

from config import settings
from handlers import tone

_WEEKDAY_NAMES = {1: '週一', 2: '週二', 3: '週三', 4: '週四', 5: '週五', 6: '週六', 7: '週日'}
_WEEKDAY_MAP = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '日': 7, '天': 7}


def _parse_date(text: str, now: datetime) -> datetime | None:
    """從訊息抽取指定日期，回傳 datetime 或 None"""
    tz = now.tzinfo

    # M/D、M-D、M月D日/號
    m = re.search(r'(\d{1,2})[/\-月](\d{1,2})[號日]?', text)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            dt = now.replace(month=month, day=day, hour=0, minute=0, second=0, microsecond=0)
            if dt.date() < now.date():
                dt = dt.replace(year=dt.year + 1)
            return dt
        except ValueError:
            pass

    # D號（無月份）
    m = re.search(r'(\d{1,2})[號日]', text)
    if m:
        day = int(m.group(1))
        try:
            dt = now.replace(day=day, hour=0, minute=0, second=0, microsecond=0)
            if dt.date() < now.date():
                next_month = now.month % 12 + 1
                year = now.year + (1 if now.month == 12 else 0)
                dt = dt.replace(year=year, month=next_month)
            return dt
        except ValueError:
            pass

    # 明天、後天、大後天
    if '明天' in text or '明日' in text:
        return now + timedelta(days=1)
    if '後天' in text:
        return now + timedelta(days=2)
    if '大後天' in text:
        return now + timedelta(days=3)

    # 下週X
    m = re.search(r'下[週周]([一二三四五六日天])', text)
    if m:
        target_wd = _WEEKDAY_MAP[m.group(1)]
        days_ahead = (target_wd - now.isoweekday()) % 7 or 7
        return now + timedelta(days=days_ahead + 7)

    return None


def handle_business_hours(text: str = "") -> str:
    """回覆營業時間，若訊息含指定日期則針對該日回覆"""
    tz = pytz.timezone(settings.BUSINESS_TZ)
    now = datetime.now(tz)

    s = settings.BUSINESS_HOURS_START
    e = settings.BUSINESS_HOURS_END
    addr = settings.STORE_ADDRESS
    biz_days = settings.business_days_list()

    target = _parse_date(text, now) if text else None

    if target:
        wd = target.isoweekday()
        date_label = f"{target.month}/{target.day}（{_WEEKDAY_NAMES[wd]}）"
        if wd not in biz_days:
            return tone.business_hours_specific_closed(date_label, s, e)
        return tone.business_hours_specific_open(date_label, s, e, addr)

    # 無指定日期 → 看今天 + 現在時段
    if now.isoweekday() not in biz_days:
        return tone.business_hours_holiday(s, e, addr)

    sh, sm = map(int, s.split(":"))
    eh, em = map(int, e.split(":"))
    open_time  = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
    close_time = now.replace(hour=eh, minute=em, second=0, microsecond=0)

    if now < open_time:
        return tone.business_hours_not_open_yet(s, e, addr)
    if now > close_time:
        return tone.business_hours_after_close(s, e)
    return tone.business_hours_open(s, e, addr)


def _is_open_now(now: datetime) -> bool:
    today_iso = now.isoweekday()  # 1=週一 … 7=週日
    if today_iso not in settings.business_days_list():
        return False

    sh, sm = map(int, settings.BUSINESS_HOURS_START.split(":"))
    eh, em = map(int, settings.BUSINESS_HOURS_END.split(":"))

    start = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
    end = now.replace(hour=eh, minute=em, second=0, microsecond=0)

    return start <= now <= end
