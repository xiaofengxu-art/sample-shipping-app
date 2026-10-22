"""统一使用北京时间，避免云服务器（UTC）导致时间显示、文件命名、支付日期偏移 8 小时。"""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone

try:
    from zoneinfo import ZoneInfo

    SHANGHAI_TZ = ZoneInfo("Asia/Shanghai")
except Exception:  # 极少数环境缺少 tzdata 时退回固定 UTC+8
    SHANGHAI_TZ = timezone(timedelta(hours=8), name="Asia/Shanghai")


def now_shanghai() -> datetime:
    """当前北京时间（带时区信息）。"""
    return datetime.now(SHANGHAI_TZ)


def today_shanghai() -> date:
    """当前北京时间的日期，用于写入支付日期。"""
    return now_shanghai().date()
