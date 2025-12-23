"""Utility functions: name normalization, date parsing/resolution, holiday checks"""
from typing import List, Optional
import re
from datetime import datetime, date, timedelta
from dateutil import parser as dateparser
import jpholiday


def normalize_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name)
    # trim and normalize spaces
    s = s.strip()
    s = re.sub(r"[ 　]+", " ", s)
    return s


def get_month_dates(year: int, month: int) -> List[date]:
    d = date(year, month, 1)
    res = []
    while d.month == month:
        res.append(d)
        d += timedelta(days=1)
    return res


def is_holiday(d: date) -> bool:
    # jpholiday returns True for national holidays
    return jpholiday.is_holiday_name(d) is not None


def parse_single_date_token(token: str, target_year: int, target_month: int) -> Optional[date]:
    """Parse tokens like '1', '1日', '2026/1/1', '2026-01-01' into date object within target month when possible."""
    if token is None:
        return None
    t = str(token).strip()
    if t == "":
        return None

    # Direct parse attempt
    try:
        dt = dateparser.parse(t, dayfirst=False, yearfirst=False, default=datetime(target_year, target_month, 1))
        # ensure same month/year
        if dt.year == target_year and dt.month == target_month:
            return dt.date()
    except Exception:
        pass

    # Try to extract number like '1' or '1日'
    m = re.search(r"(\d{1,2})", t)
    if m:
        day = int(m.group(1))
        try:
            return date(target_year, target_month, day)
        except Exception:
            return None

    return None


def normalize_date_input(text: str, month: str) -> List[str]:
    """Given an input text (possibly multiple tokens) and target month 'YYYY-MM', return list of 'YYYY-MM-DD' strings.

    Supports: '1', '1日', '2026/1/1', '1,2,3', '1-3' (range)
    """
    if text is None:
        return []
    ts = str(text).strip()
    if ts == "":
        return []
    year, mon = [int(x) for x in month.split("-")]
    out = set()

    # split by commas or whitespace
    parts = re.split(r"[，,\s]+", ts)
    for p in parts:
        p = p.strip()
        if p == "":
            continue
        # range like 1-3
        if re.match(r"^\d{1,2}-\d{1,2}$", p):
            a, b = [int(x) for x in p.split("-")]
            for d in range(a, b+1):
                try:
                    out.add(date(year, mon, d).isoformat())
                except Exception:
                    continue
            continue

        single = parse_single_date_token(p, year, mon)
        if single:
            out.add(single.isoformat())

    return sorted(list(out))
