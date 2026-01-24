from __future__ import annotations

import sys
from pathlib import Path

# ====== runtime / 路徑相關 ======
WORKSPACE_ROOT = Path(__file__).resolve().parent
USER_DATA_ROOT_DIRNAME = "user_data"
USER_DATA_STORE_DIRNAME = "user_schedules"
COURSE_INPUT_DIRNAME = "course_inputs"


def runtime_root_path() -> Path:
    """開發期即 workspace root，打包後為 exe 同層。"""
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else WORKSPACE_ROOT


def user_data_root_path() -> Path:
    """取得程式目前所應遵守的 user_data 根資料夾。"""
    return runtime_root_path() / USER_DATA_ROOT_DIRNAME


def user_data_store_path() -> Path:
    """實際保存使用者資料的子資料夾（user_data/user_schedules）。"""
    return user_data_root_path() / USER_DATA_STORE_DIRNAME


def course_input_dir_path() -> Path:
    """儲存課程輸入 Excel 的目錄（user_data/course_inputs）。"""
    path = user_data_root_path() / COURSE_INPUT_DIRNAME
    path.mkdir(parents=True, exist_ok=True)
    return path


# ====== Excel / 資料欄位 ======
COURSE_SHEET_CANDIDATES = ["課程", "Courses", "Sheet1"]
REQUIRED_COLUMNS = [
    "開課序號",
    "開課代碼",
    "系所",
    "中文課程名稱",
    "教師",
    "學分",
    "必/選",
    "全/半",
    "地點時間",
    "限修人數",
    "選修人數",
]

# ====== 星期 / 節次 ======
DAYS = ["一", "二", "三", "四", "五", "六", "日"]
DAY_LABEL = {
    "一": "星期一",
    "二": "星期二",
    "三": "星期三",
    "四": "星期四",
    "五": "星期五",
    "六": "星期六",
    "日": "星期日",
    "天": "星期日",
}

PERIODS = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "A", "B", "C", "D"]
PERIOD_INDEX = {p: i for i, p in enumerate(PERIODS)}
DAY_INDEX = {d: i for i, d in enumerate(DAYS)}
BITS_PER_DAY = len(PERIODS)

# 節次對應時間
PERIOD_TIME = {
    "0": "07:10 ~ 08:00",
    "1": "08:10 ~ 09:00",
    "2": "09:10 ~ 10:00",
    "3": "10:20 ~ 11:10",
    "4": "11:20 ~ 12:10",
    "5": "12:20 ~ 13:10",
    "6": "13:20 ~ 14:10",
    "7": "14:20 ~ 15:10",
    "8": "15:30 ~ 16:20",
    "9": "16:30 ~ 17:20",
    "10": "17:30 ~ 18:20",
    "A": "18:40 ~ 19:30",
    "B": "19:35 ~ 20:25",
    "C": "20:30 ~ 21:20",
    "D": "21:25 ~ 22:15",
}

# ====== 特殊篩選 ======
GENED_DEPT_NAME = "通識"
SPORT_DEPT_NAME = "普通體育"
TEACHING_NAME_TOKEN = "（教）"

GENED_CORE_OPTIONS = [
    "人文藝術",
    "社會科學",
    "自然科學",
    "邏輯運算",
    "學院共同課程",
    "跨域專業探索課程",
    "大學入門",
    "專題探索",
    "MOOCs",
    "所有通識",
]
