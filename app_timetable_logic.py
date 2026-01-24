from __future__ import annotations

from typing import Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd
from PySide6.QtGui import QColor

from app_constants import DAY_LABEL, DAYS, PERIODS, PERIOD_INDEX
from app_utils import strip_bracket_text_for_timetable, format_cid4


def darken(color: QColor, factor: float) -> QColor:
    r = max(0, min(255, int(color.red() * factor)))
    g = max(0, min(255, int(color.green() * factor)))
    b = max(0, min(255, int(color.blue() * factor)))
    return QColor(r, g, b)


def _subset_courses_by_ids(
    courses_df: pd.DataFrame, included_ids_sorted: np.ndarray, cols: List[str]
) -> pd.DataFrame:
    if courses_df is None or courses_df.empty:
        return pd.DataFrame(columns=cols)
    if included_ids_sorted is None or included_ids_sorted.size == 0:
        return courses_df.iloc[0:0][cols]

    ids = included_ids_sorted.astype(np.int64, copy=False)
    cid_arr = courses_df["_cid"].to_numpy(dtype=np.int64, copy=False)
    mask = np.isin(cid_arr, ids)
    if not mask.any():
        return courses_df.iloc[0:0][cols]
    return courses_df.loc[mask, cols]


def occupied_masks_sorted(courses_df: pd.DataFrame, included_ids_sorted: np.ndarray) -> Tuple[np.uint64, np.uint64]:
    if included_ids_sorted is None or included_ids_sorted.size == 0:
        return np.uint64(0), np.uint64(0)
    sub = _subset_courses_by_ids(courses_df, included_ids_sorted, ["_mask_lo", "_mask_hi"])
    if sub.empty:
        return np.uint64(0), np.uint64(0)

    arr_lo = sub["_mask_lo"].to_numpy(dtype="uint64", copy=False)
    arr_hi = sub["_mask_hi"].to_numpy(dtype="uint64", copy=False)
    occ_lo = np.bitwise_or.reduce(arr_lo) if arr_lo.size else np.uint64(0)
    occ_hi = np.bitwise_or.reduce(arr_hi) if arr_hi.size else np.uint64(0)
    return np.uint64(occ_lo), np.uint64(occ_hi)


def build_lane_assignment_sorted(courses_df: pd.DataFrame, included_ids_sorted: np.ndarray) -> Tuple[Dict[int, int], int]:
    if included_ids_sorted is None or included_ids_sorted.size == 0:
        return {}, 1
    sub = _subset_courses_by_ids(courses_df, included_ids_sorted, ["_cid", "_slots_set"])
    if sub.empty:
        return {}, 1

    used_by_slot: Dict[str, int] = {}
    lane_of: Dict[int, int] = {}
    max_lane = 1

    rows = list(sub.itertuples(index=False, name=None))
    rows.sort(key=lambda x: (-(len(x[1]) if isinstance(x[1], set) else 0), int(x[0])))

    for cid, slots in rows:
        cid_i = int(cid)
        slots_set = slots if isinstance(slots, set) else set()
        if not slots_set:
            lane_of[cid_i] = 1
            continue

        used = 0
        for s in slots_set:
            used |= used_by_slot.get(s, 0)

        bitlen = used.bit_length() + 2
        inv = (~used) & ((1 << bitlen) - 1)
        lsb = inv & -inv
        lane = lsb.bit_length()

        lane_of[cid_i] = lane
        if lane > max_lane:
            max_lane = lane

        bit = 1 << (lane - 1)
        for s in slots_set:
            used_by_slot[s] = used_by_slot.get(s, 0) | bit

    return lane_of, max_lane


def build_timetable_matrix_per_day_lanes_sorted(
    courses_df: pd.DataFrame,
    included_ids_sorted: np.ndarray,
    locked_ids_set: Set[int],
    show_days: List[str],
) -> Tuple[
    List[List[str]],
    List[str],
    Dict[str, int],
    List[int],
    List[List[Optional[int]]],
    List[List[bool]],
]:
    lane_map, _ = build_lane_assignment_sorted(courses_df, included_ids_sorted)
    day_lanes: Dict[str, int] = {d: 1 for d in show_days}

    subset_meta = _subset_courses_by_ids(
        courses_df,
        included_ids_sorted,
        ["_cid", "銝剜?隤脩??迂", "_slots_set", "_tt_label"],
    )

    for cid, _cname, slots in subset_meta.itertuples(index=False, name=None):
        cid_i = int(cid)
        lane = lane_map.get(cid_i, 1)
        slots_set = slots if isinstance(slots, set) else set()
        for slot in slots_set:
            day, _per = slot.split("-", 1)
            if day in day_lanes and lane > day_lanes[day]:
                day_lanes[day] = lane

    day_offset: Dict[str, int] = {}
    col_day_idx: List[int] = []
    offset = 0
    for day_idx, d in enumerate(show_days):
        day_offset[d] = offset
        for _ in range(int(day_lanes[d])):
            col_day_idx.append(day_idx)
        offset += int(day_lanes[d])

    cols = offset
    matrix = [["" for _ in range(cols)] for _ in PERIODS]
    id_matrix: List[List[Optional[int]]] = [[None for _ in range(cols)] for _ in PERIODS]
    locked_matrix: List[List[bool]] = [[False for _ in range(cols)] for _ in PERIODS]
    conflict_slots: List[str] = []

    locked_ids_set = set(int(x) for x in locked_ids_set)

    for cid, cname, slots, tt_label in subset_meta.itertuples(index=False, name=None):
        cid_i = int(cid)
        tt_label_text = str(tt_label).strip() if tt_label else ""
        if tt_label_text:
            label = tt_label_text
        else:
            cname_show = strip_bracket_text_for_timetable(str(cname).strip())
            label = f"{cname_show}\n{format_cid4(cid_i)}".strip()

        slots_set = slots if isinstance(slots, set) else set()
        lane = lane_map.get(cid_i, 1)
        is_locked = (cid_i in locked_ids_set)

        for slot in slots_set:
            day, per = slot.split("-", 1)
            if day not in day_offset or per not in PERIOD_INDEX:
                continue

            r = PERIOD_INDEX.get(per)
            if r is None:
                continue
            c = day_offset[day] + (lane - 1)

            if matrix[r][c] and label not in matrix[r][c]:
                conflict_slots.append(f"{DAY_LABEL.get(day, day)} 蝚洌per}蝭")
                matrix[r][c] = matrix[r][c] + "\n---\n" + label
                if is_locked:
                    locked_matrix[r][c] = True
            else:
                matrix[r][c] = label
                if id_matrix[r][c] is None:
                    id_matrix[r][c] = cid_i
                if is_locked:
                    locked_matrix[r][c] = True\n
    uniq: List[str] = []
    seen: Set[str] = set()
    for x in conflict_slots:
        if x not in seen:
            uniq.append(x)
            seen.add(x)

    return matrix, uniq, day_lanes, col_day_idx, id_matrix, locked_matrix



