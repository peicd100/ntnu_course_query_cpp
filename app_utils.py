from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional, Set, Tuple

import numpy as np
import pandas as pd

from app_constants import BITS_PER_DAY, DAY_INDEX, DAYS, PERIOD_INDEX, PERIODS


@dataclass
class ParsedTime:
    slots: Set[str]
    tba: bool


def format_cid4(cid: int) -> str:
    try:
        n = int(cid)
    except Exception:
        return str(cid)
    return f"{n:04d}" if 0 <= n < 10000 else str(n)


def parse_cid_to_int(val) -> Optional[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def strip_bracket_text_for_timetable(name: str) -> str:
    s = str(name or "")
    s = re.sub(r"\[.*?\]", "", s)
    s = re.sub(r"【.*?】", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def sorted_array_from_set_int(s: Set[int]) -> np.ndarray:
    if not s:
        return np.empty((0,), dtype=np.int64)
    return np.array(sorted((int(x) for x in s)), dtype=np.int64)


def slot_to_mask(day: str, period: str) -> Tuple[np.uint64, np.uint64]:
    if day not in DAY_INDEX:
        return np.uint64(0), np.uint64(0)
    if period not in PERIOD_INDEX:
        return np.uint64(0), np.uint64(0)
    idx = DAY_INDEX[day] * BITS_PER_DAY + PERIOD_INDEX[period]
    if idx < 64:
        return np.uint64(1) << np.uint64(idx), np.uint64(0)
    return np.uint64(0), np.uint64(1) << np.uint64(idx - 64)


def slots_set_to_masks(slots: Set[str]) -> Tuple[np.uint64, np.uint64]:
    lo = np.uint64(0)
    hi = np.uint64(0)
    if not slots:
        return lo, hi
    for s in slots:
        try:
            d, p = s.split("-", 1)
        except Exception:
            continue
        mlo, mhi = slot_to_mask(d, p)
        lo |= mlo
        hi |= mhi
    return lo, hi


def expand_period_range(start: str, end: Optional[str]) -> List[str]:
    start_u = str(start).strip().upper()
    end_u = str(end).strip().upper() if end is not None else None

    if start_u not in PERIOD_INDEX:
        raise ValueError(f"未知節次：{start_u}")
    if end_u is None:
        return [start_u]
    if end_u not in PERIOD_INDEX:
        raise ValueError(f"未知節次：{end_u}")

    i, j = PERIOD_INDEX[start_u], PERIOD_INDEX[end_u]
    if j < i:
        i, j = j, i
    return PERIODS[i : j + 1]


def _extract_first_day(s: str) -> Optional[str]:
    for ch in s:
        if ch in ("一", "二", "三", "四", "五", "六", "日", "天"):
            return "日" if ch == "天" else ch
    return None


def _extract_first_token(s: str) -> Optional[str]:
    t = s.strip().upper()
    if not t:
        return None
    if t.startswith("10"):
        return "10"
    ch = t[0]
    if ch.isdigit():
        return ch
    if ch in ("A", "B", "C", "D"):
        return ch
    return None


def parse_time_text(text: str) -> ParsedTime:
    if text is None:
        return ParsedTime(slots=set(), tba=True)

    s0 = str(text).strip()
    if not s0 or s0 == "nan":
        return ParsedTime(slots=set(), tba=True)

    parts: List[str] = []
    buf = ""
    for ch in s0:
        if ch in (",", "，", ";", "；"):
            if buf.strip():
                parts.append(buf.strip())
            buf = ""
        else:
            buf += ch
    if buf.strip():
        parts.append(buf.strip())

    slots: Set[str] = set()
    matched_any = False

    for part in parts:
        day = _extract_first_day(part)
        if day is None:
            continue

        t = part.replace(" ", "")
        day_pos = t.find(day)
        if day_pos < 0:
            continue
        rest = t[day_pos + 1 :]
        if not rest:
            continue

        matched_any = True

        if "-" in rest:
            left, right = rest.split("-", 1)
            start = _extract_first_token(left)
            end = _extract_first_token(right)
            if start is None:
                continue
            try:
                expanded = expand_period_range(start, end)
            except Exception:
                continue
            for p in expanded:
                slots.add(f"{day}-{p}")
        else:
            start = _extract_first_token(rest)
            if start is None:
                continue
            slots.add(f"{day}-{start}")

    if not matched_any:
        return ParsedTime(slots=set(), tba=True)

    return ParsedTime(slots=slots, tba=False if slots else True)


def _slot_sort_key(slot: str) -> Tuple[int, int]:
    day, per = slot.split("-", 1)
    return (DAYS.index(day) if day in DAYS else 99, PERIOD_INDEX.get(per, 999))


def parse_gened_categories_from_course_name(course_name: str) -> List[str]:
    s = str(course_name or "").strip()
    if not s:
        return []

    matches_sq = list(re.finditer(r"\[(.*?)\]", s))
    matches_fw = list(re.finditer(r"【(.*?)】", s))

    picked = None
    if matches_sq and matches_fw:
        picked = matches_sq[-1] if matches_sq[-1].start() > matches_fw[-1].start() else matches_fw[-1]
    elif matches_sq:
        picked = matches_sq[-1]
    elif matches_fw:
        picked = matches_fw[-1]

    if not picked:
        return []

    inner = (picked.group(1) or "").strip()
    if not inner:
        return []

    parts = [p.strip() for p in re.split(r"[；;]", inner) if p.strip()]
    if not parts:
        return []

    last = parts[-1]
    if "：" in last:
        tail = last.split("：")[-1].strip()
    elif ":" in last:
        tail = last.split(":")[-1].strip()
    else:
        tail = last.strip()

    tail = re.sub(r"\s+", " ", tail).strip()
    if not tail:
        return []

    return [t.strip() for t in tail.split(" ") if t.strip()]


def sanitize_folder_name(name: str) -> str:
    s = (name or "").strip() or "未命名使用者"
    s = re.sub(r"[\\/:*?\[\]]+", "_", s)
    s = s.replace("\r", " ").replace("\n", " ").strip()
    return s[:80] if len(s) > 80 else s


def truthy_flag(v) -> bool:
    return v in (True, 1, "1", "Y", "y", "yes", "YES", "True", "true")
