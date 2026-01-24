from __future__ import annotations

import os
import shutil
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple
from functools import partial

import numpy as np
import pandas as pd

from PySide6.QtCore import (
    Qt,
    QEvent,
    QSignalBlocker,
    QTimer,
    QSortFilterProxyModel,
    QThreadPool,
)
from PySide6.QtGui import QAction, QBrush, QColor, QFontMetrics
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QCompleter,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app_constants import (
    DAY_LABEL,
    GENED_CORE_OPTIONS,
    GENED_DEPT_NAME,
    PERIODS,
    PERIOD_TIME,
    SPORT_DEPT_NAME,
    TEACHING_NAME_TOKEN,
    course_input_dir_path,
)
from app_excel import ensure_excel_readable, load_courses_auto
from app_timetable_logic import build_timetable_matrix_per_day_lanes_sorted, darken, occupied_masks_sorted
from app_user_data import (
    best_schedule_dir_path,
    list_all_users,
    list_user_history_files,
    load_best_schedule_cache,
    load_user_file,
    unique_login_file_path,
    user_dir,
    user_root_dir,
    save_user_file,
)
from app_utils import (
    parse_gened_categories_from_course_name,
    sanitize_folder_name,
    slot_to_mask,
    sorted_array_from_set_int,
    format_cid4,
)
from app_widgets import (
    FavoritesTableWidget,
    FloatSortItem,
    IntSortItem,
    ResultsFrozenView,
    ResultsModel,
    TimetableWidget,
    TTTimeSelectDelegate,
)
from app_workers import BestScheduleWorker, SaveWorker

FAV_CID_ROLE = Qt.UserRole + 1


def _is_excel_file_name(fn: str) -> bool:
    lf = fn.lower()
    if not (lf.endswith(".xls") or lf.endswith(".xlsx")):
        return False
    if lf.startswith("~$"):
        return False
    return True


def find_lex_last_excel(search_dirs: Sequence[str]) -> Optional[str]:
    candidates: List[str] = []
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        try:
            for fn in os.listdir(d):
                if not _is_excel_file_name(fn):
                    continue
                p = os.path.join(d, fn)
                if os.path.isfile(p):
                    candidates.append(p)
        except Exception:
            continue

    if not candidates:
        return None

    # 以檔名（字典序）為準，取最後
    best = max(candidates, key=lambda p: os.path.basename(p).casefold())
    return best


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("課程查詢與排課系統（桌面版）")
        self.resize(1600, 900)

        self.excel_path: str = ""
        self.course_sheet_name: str = ""
        self.courses_df: Optional[pd.DataFrame] = None
        self.display_columns: List[str] = []
        self.filtered_df: pd.DataFrame = pd.DataFrame()

        self.username: str = ""
        self.user_dir_path: str = ""
        self.session_file_path: str = ""
        self.readonly_mode: bool = False

        self.favorites_ids: Set[int] = set()
        self.included_ids: Set[int] = set()
        self.locked_ids: Set[int] = set()

        self.fav_seq: Dict[int, int] = {}
        self._fav_seq_next: int = 1

        self._favorites_sorted_cache = np.empty((0,), dtype=np.int64)
        self._favorites_sorted_dirty = True
        self._included_sorted_cache = np.empty((0,), dtype=np.int64)
        self._included_sorted_dirty = True
        self._locked_sorted_cache = np.empty((0,), dtype=np.int64)
        self._locked_sorted_dirty = True

        self._sel_lo = np.uint64(0)
        self._sel_hi = np.uint64(0)

        self.show_saturday = False
        self.show_time = False
        self.show_days: List[str] = self._calc_show_days()
        self._slot_mask_lo: List[List[np.uint64]] = []
        self._slot_mask_hi: List[List[np.uint64]] = []
        self._rebuild_slot_masks()

        self._tt_col_day_idx: List[int] = []
        self._tt_first_lane_col: Dict[int, int] = {}
        self._tt_locked_matrix: List[List[bool]] = []

        self._tt_dragging = False
        self._tt_drag_state: Optional[bool] = None
        self._tt_drag_start_day: Optional[int] = None
        self._tt_drag_start_row: Optional[int] = None
        self._tt_drag_base_lo = np.uint64(0)
        self._tt_drag_base_hi = np.uint64(0)
        self._tt_drag_last_rect: Optional[Tuple[int, int, int, int]] = None
        self._tt_drag_has_moved = False
        self._tt_drag_initial_rect: Optional[Tuple[int, int, int, int]] = None

        self._session_fav_backup: Optional[Set[int]] = None
        self._session_inc_backup: Optional[Set[int]] = None
        self._session_lock_backup: Optional[Set[int]] = None
        self._session_seq_backup: Optional[Dict[int, int]] = None
        self._history_selected_file: str = ""
        self._history_mode: Optional[str] = None
        self._best_worker: Optional[BestScheduleWorker] = None
        self._best_running = False
        self._best_token = 0
        self._best_files: List[str] = []
        self._history_selected_brush = QBrush(QColor("#FFB74D"))
        self.tbl_tt_preview: Optional[TimetableWidget] = None
        self._history_preview_snapshot: Optional[Dict[str, Any]] = None
        self._history_layout_active = False

        self._day_bg_base: List[QColor] = [
            QColor("#F2F7FF"),
            QColor("#F2FFF7"),
            QColor("#FFF7F2"),
            QColor("#F9F2FF"),
            QColor("#FFFDF2"),
            QColor("#F2FFFF"),
        ]
        self._block_dark_factor = 0.82

        self._cid_sorted: Optional[np.ndarray] = None
        self._name_sorted: Optional[np.ndarray] = None
        self._teacher_sorted: Optional[np.ndarray] = None
        self._credit_sorted: Optional[np.ndarray] = None
        self._all_depts: Set[str] = set()
        self._cid_arr: Optional[np.ndarray] = None
        self._mask_lo_arr: Optional[np.ndarray] = None
        self._mask_hi_arr: Optional[np.ndarray] = None
        self._tba_arr: Optional[np.ndarray] = None
        self._dept_arr: Optional[np.ndarray] = None
        self._slots_by_cid: Dict[int, Set[str]] = {}

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._do_search_now)

        self.threadpool = QThreadPool.globalInstance()
        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.timeout.connect(self._autosave_now)

        self._save_token = 0
        self._save_latest_token = 0
        self._save_inflight = False
        self._save_pending = False
        self._save_pending_snapshot: Optional[Tuple[Set[int], Set[int], Set[int], Dict[int, int]]] = None

        self._default_sizes_applied = False
        self._fav_default_width_applied = False

        self.FAV_COL_HANDLE = 0
        self.FAV_COL_SCHEDULE = 1
        self.FAV_COL_LOCK = 2
        self.FAV_COL_ID = 3
        self.FAV_COL_NAME = 4
        self.FAV_COL_TEACHER = 5
        self.FAV_COL_CREDIT = 6
        self.FAV_COL_RANK = 7
        self.FAV_COL_DELETE = 8
        self._fav_sort_section = self.FAV_COL_RANK
        self._fav_sort_order = Qt.AscendingOrder

        self._build_menu()
        self._build_ui()
        self._try_autoload_default_excel()

    # ====== UI / Menu ======
    def showEvent(self, event):
        super().showEvent(event)
        if not self._default_sizes_applied:
            self._default_sizes_applied = True
            QTimer.singleShot(0, self._apply_default_panel_sizes)

    def _build_menu(self) -> None:
        open_act = QAction("開啟課程 Excel…", self)
        open_act.triggered.connect(self.on_open_excel)

        reload_act = QAction("重新載入", self)
        reload_act.triggered.connect(self.on_reload_excel)

        menubar = self.menuBar()
        file_menu = menubar.addMenu("檔案")
        file_menu.addAction(open_act)
        file_menu.addAction(reload_act)

    def _configure_combo_searchable(self, cb: QComboBox, placeholder: str) -> None:
        cb.setEditable(True)
        cb.setInsertPolicy(QComboBox.NoInsert)
        if cb.lineEdit():
            cb.lineEdit().setPlaceholderText(placeholder)
            cb.lineEdit().setClearButtonEnabled(True)

        comp = QCompleter(cb.model(), cb)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        comp.setFilterMode(Qt.MatchContains)
        comp.setCompletionMode(QCompleter.PopupCompletion)
        cb.setCompleter(comp)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)

        top = QHBoxLayout()
        root_layout.addLayout(top)

        self.lbl_excel = QLabel("課程 Excel：尚未載入")
        self.lbl_excel.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top.addWidget(self.lbl_excel, 1)

        top.addWidget(QLabel("選擇使用者："))
        self.cb_users = QComboBox()
        self._configure_combo_searchable(self.cb_users, "搜尋/選擇已建立使用者")
        self.cb_users.addItem("(未選擇)")
        self.cb_users.setMinimumWidth(220)
        top.addWidget(self.cb_users)

        top.addWidget(QLabel("新增使用者："))
        self.ed_new_user = QLineEdit()
        self.ed_new_user.setPlaceholderText("輸入新使用者名稱（建立新使用者）")
        self.ed_new_user.setMinimumWidth(220)
        top.addWidget(self.ed_new_user)

        self.btn_login = QPushButton("新增/切換")
        self.btn_login.clicked.connect(self.on_login)
        top.addWidget(self.btn_login)

        self.lbl_user = QLabel("尚未登入")
        self.lbl_user.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top.addWidget(self.lbl_user)

        self.split_main = QSplitter(Qt.Horizontal)
        root_layout.addWidget(self.split_main, 1)

        left = QWidget()
        right = QWidget()
        self.split_main.addWidget(left)
        self.split_main.addWidget(right)
        self.split_main.setSizes([800, 800])
        self.split_main.setChildrenCollapsible(False)

        left_layout = QVBoxLayout(left)
        right_layout = QVBoxLayout(right)

        self.split_left_v = QSplitter(Qt.Vertical)
        left_layout.addWidget(self.split_left_v, 1)
        self.split_left_v.setSizes([450, 450])
        self.split_left_v.setChildrenCollapsible(False)

        self.left_top = QWidget()
        self.left_bottom = QWidget()
        self.split_left_v.addWidget(self.left_top)
        self.split_left_v.addWidget(self.left_bottom)

        self.split_left_top = QSplitter(Qt.Horizontal)
        lt_layout = QVBoxLayout(self.left_top)
        lt_layout.setContentsMargins(0, 0, 0, 0)
        lt_layout.addWidget(self.split_left_top, 1)
        self.split_left_top.setSizes([420, 380])
        self.split_left_top.setChildrenCollapsible(False)

        lb_layout = QVBoxLayout(self.left_bottom)
        lb_layout.setContentsMargins(0, 0, 0, 0)

        # ====== 課表 ======
        self.gb_tt = QGroupBox("課表檢視")
        right_layout.addWidget(self.gb_tt, 1)
        tt_layout = QVBoxLayout(self.gb_tt)

        self.tt_splitter = QSplitter(Qt.Horizontal)
        self.tt_splitter.setChildrenCollapsible(False)

        self.gb_tt_primary = QGroupBox("檢視檔案")
        primary_layout = QVBoxLayout(self.gb_tt_primary)

        header_row = QHBoxLayout()
        header_row.addStretch(1)

        self.btn_show_time = QPushButton("顯示時間")
        self.btn_show_time.setCheckable(True)
        self.btn_show_time.setChecked(self.show_time)
        self.btn_show_time.toggled.connect(self._set_show_time)
        header_row.addWidget(self.btn_show_time, 0, Qt.AlignRight)

        self.btn_show_sat = QPushButton("顯示週六")
        self.btn_show_sat.setCheckable(True)
        self.btn_show_sat.setChecked(self.show_saturday)
        self.btn_show_sat.toggled.connect(self._set_show_saturday)
        header_row.addWidget(self.btn_show_sat, 0, Qt.AlignRight)

        self.lbl_total_credits = QLabel("已選總學分：0")
        self.lbl_total_credits.setStyleSheet(
            "QLabel { background: #FFF59D; color: #000000; padding: 4px 10px; border-radius: 6px; font-weight: 700; }"
        )
        self.lbl_total_credits.setAlignment(Qt.AlignCenter)
        header_row.addWidget(self.lbl_total_credits, 0, Qt.AlignRight)

        self.lbl_locked_credits = QLabel("已鎖定學分：0")
        self.lbl_locked_credits.setStyleSheet(
            "QLabel { background: #B0BEC5; color: #000000; padding: 4px 10px; border-radius: 6px; font-weight: 700; }"
        )
        self.lbl_locked_credits.setAlignment(Qt.AlignCenter)
        header_row.addWidget(self.lbl_locked_credits, 0, Qt.AlignRight)

        primary_layout.addLayout(header_row)

        self.tbl_tt = TimetableWidget()
        self.tbl_tt_preview = TimetableWidget()
        self._configure_timetable_widget(self.tbl_tt, use_delegate=True)
        self._configure_timetable_widget(self.tbl_tt_preview, use_delegate=False)

        primary_layout.addWidget(self.tbl_tt, 1)

        self.lbl_conflicts = QLabel("")
        self.lbl_conflicts.setWordWrap(True)
        primary_layout.addWidget(self.lbl_conflicts)

        self.gb_tt_preview = QGroupBox("目前方案")
        preview_layout = QVBoxLayout(self.gb_tt_preview)
        preview_header = QHBoxLayout()
        preview_header.addStretch(1)

        self.btn_show_time_preview = QPushButton("顯示時間")
        self.btn_show_time_preview.setCheckable(True)
        self.btn_show_time_preview.setChecked(self.show_time)
        self.btn_show_time_preview.toggled.connect(self._set_show_time)
        preview_header.addWidget(self.btn_show_time_preview, 0, Qt.AlignRight)

        self.btn_show_sat_preview = QPushButton("顯示週六")
        self.btn_show_sat_preview.setCheckable(True)
        self.btn_show_sat_preview.setChecked(self.show_saturday)
        self.btn_show_sat_preview.toggled.connect(self._set_show_saturday)
        preview_header.addWidget(self.btn_show_sat_preview, 0, Qt.AlignRight)

        self.lbl_total_credits_preview = QLabel("已選總學分：0")
        self.lbl_total_credits_preview.setStyleSheet(self.lbl_total_credits.styleSheet())
        self.lbl_total_credits_preview.setAlignment(Qt.AlignCenter)
        preview_header.addWidget(self.lbl_total_credits_preview, 0, Qt.AlignRight)

        self.lbl_locked_credits_preview = QLabel("已鎖定學分：0")
        self.lbl_locked_credits_preview.setStyleSheet(self.lbl_locked_credits.styleSheet())
        self.lbl_locked_credits_preview.setAlignment(Qt.AlignCenter)
        preview_header.addWidget(self.lbl_locked_credits_preview, 0, Qt.AlignRight)

        preview_layout.addLayout(preview_header)
        preview_layout.addWidget(self.tbl_tt_preview, 1)

        self.gb_tt_preview.setVisible(False)

        self.tt_splitter.addWidget(self.gb_tt_primary)
        self.tt_splitter.addWidget(self.gb_tt_preview)
        self.tt_splitter.setSizes([1, 0])

        tt_layout.addWidget(self.tt_splitter, 1)

        self.tbl_tt.zoomChanged.connect(self._on_tt_zoom_changed)
        self.tbl_tt_preview.zoomChanged.connect(self._on_tt_zoom_changed)

        # ====== 查詢條件 ======
        self.gb_filters = QGroupBox("查詢條件")
        self.gb_filters.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Minimum)
        self.split_left_top.addWidget(self.gb_filters)

        vbox = QVBoxLayout(self.gb_filters)
        form_widget = QWidget()
        form = QFormLayout(form_widget)

        self.ed_serial = QLineEdit()
        self.ed_serial.setPlaceholderText("例如：0001 或 0123 0456")
        form.addRow("開課序號", self.ed_serial)

        self.ed_course_code = QLineEdit()
        self.ed_course_code.setPlaceholderText("例如：00UP013（可空白分隔多關鍵字）")
        form.addRow("開課代碼", self.ed_course_code)

        self.ed_cname = QLineEdit()
        self.ed_cname.setPlaceholderText("課名關鍵字")
        form.addRow("科目名稱", self.ed_cname)

        self.ed_teacher = QLineEdit()
        self.ed_teacher.setPlaceholderText("教師關鍵字")
        form.addRow("教師", self.ed_teacher)

        self.ed_full = QLineEdit()
        self.ed_full.setPlaceholderText("例如：A B（空白分隔多關鍵字；需全部命中）")
        form.addRow("全面搜尋", self.ed_full)

        self.cb_dept = QComboBox()
        self._configure_combo_searchable(self.cb_dept, "輸入系所關鍵字搜尋")
        form.addRow("系所", self.cb_dept)

        self.ck_gened = QCheckBox("通識課程")
        self.ck_sport = QCheckBox("一般體育科目")
        self.ck_teaching = QCheckBox("教育學程科目")

        self.ck_not_full = QCheckBox("未滿額")
        self.ck_exclude_conflict = QCheckBox("先排除與課表衝突")
        self.ck_exclude_conflict.setChecked(False)
        self.ck_exclude_selected = QCheckBox("排除已選課程")
        self.ck_show_tba = QCheckBox("顯示無時間課程")
        self.ck_show_tba.setChecked(False)

        opt_container = QWidget()
        opt_v = QVBoxLayout(opt_container)
        opt_v.setContentsMargins(0, 0, 0, 0)
        opt_v.setSpacing(6)

        opt_row1 = QHBoxLayout()
        opt_row1.addWidget(self.ck_gened)
        opt_row1.addWidget(self.ck_sport)
        opt_row1.addWidget(self.ck_teaching)
        opt_row1.addStretch(1)

        opt_row2 = QHBoxLayout()
        opt_row2.addWidget(self.ck_not_full)
        opt_row2.addWidget(self.ck_exclude_conflict)
        opt_row2.addStretch(1)

        opt_row3 = QHBoxLayout()
        opt_row3.addWidget(self.ck_exclude_selected)
        opt_row3.addWidget(self.ck_show_tba)
        opt_row3.addStretch(1)

        opt_v.addLayout(opt_row1)
        opt_v.addLayout(opt_row2)
        opt_v.addLayout(opt_row3)
        form.addRow("選項", opt_container)

        self.lbl_gened_core_disabled = QLabel("未選擇通識課程")
        self.cb_gened_core = QComboBox()
        self.cb_gened_core.addItems(GENED_CORE_OPTIONS)
        self.cb_gened_core.setCurrentText("所有通識")

        self.stk_gened_core = QStackedWidget()
        self.stk_gened_core.addWidget(self.lbl_gened_core_disabled)
        self.stk_gened_core.addWidget(self.cb_gened_core)
        self.stk_gened_core.setCurrentIndex(0)
        self.stk_gened_core.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        form.addRow("核心通識範圍", self.stk_gened_core)

        self.cb_match_mode = QComboBox()
        self.cb_match_mode.addItems(["課程時間需完全落在勾選時段內", "只要與勾選時段有交集即可"])
        self.cb_match_mode.setCurrentIndex(1)
        form.addRow("時間匹配", self.cb_match_mode)

        btn_row_widget = QWidget()
        btn_row = QHBoxLayout(btn_row_widget)
        btn_row.setContentsMargins(0, 0, 0, 0)
        self.btn_clear_time = QPushButton("清空時間選取")
        self.btn_clear_time.clicked.connect(self.on_clear_time_selection)

        self.btn_clear_all = QPushButton("清空所有條件")
        self.btn_clear_all.clicked.connect(self.on_clear_all_conditions)

        btn_row.addWidget(self.btn_clear_time)
        btn_row.addWidget(self.btn_clear_all)
        btn_row.addStretch(1)
        form.addRow("操作", btn_row_widget)

        vbox.addWidget(form_widget)
        vbox.addStretch(1)

        # ====== 我的最愛 ======
        self.gb_fav = QGroupBox("我的最愛（勾選顯示於課表）")
        self.split_left_top.addWidget(self.gb_fav)
        fav_layout = QVBoxLayout(self.gb_fav)

        self.lbl_readonly = QLabel("無法修改")
        self.lbl_readonly.setStyleSheet("color: #b00020; font-weight: 600;")
        self.lbl_readonly.setVisible(False)
        fav_layout.addWidget(self.lbl_readonly)

        self.tbl_fav = FavoritesTableWidget(
            move_column=self.FAV_COL_HANDLE,
            cid_column=self.FAV_COL_ID,
        )
        self.tbl_fav.setColumnCount(9)
        self.tbl_fav.setHorizontalHeaderLabels(
            ["拖曳排序", "課表", "鎖定", "開課序號", "中文課程名稱", "教師", "學分數", "優先度", "刪除"]
        )
        self.tbl_fav.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_fav.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl_fav.cellChanged.connect(self.on_fav_cell_changed)
        self.tbl_fav.orderChanged.connect(self._on_favorites_reordered)
        self.tbl_fav.dragSelectionFinished.connect(self._on_fav_drag_selection_finished)
        # If a drop completed but did not change order, request a visual refresh
        # to avoid temporary disappearance of the dragged row.
        self.tbl_fav.dropCompleted.connect(self._on_fav_drop_completed)

        hh = self.tbl_fav.horizontalHeader()
        hh.sortIndicatorChanged.connect(self._on_fav_sort_changed)
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setStretchLastSection(False)

        self.tbl_fav.setSortingEnabled(True)
        hh.setSortIndicator(self._fav_sort_section, self._fav_sort_order)
        self._update_fav_drag_state()
        fav_layout.addWidget(self.tbl_fav, 1)

        fav_btn_row = QHBoxLayout()
        self.btn_refresh_fav = QPushButton("重新整理")
        self.btn_refresh_fav.clicked.connect(self._refresh_favorites_table)
        self.btn_refresh_fav.setToolTip("重新整理「我的最愛」列表")
        
        self.btn_clear_fav = QPushButton("清空最愛")
        self.btn_clear_fav.clicked.connect(self.on_clear_favorites)
        self.btn_clear_missing_fav = QPushButton("清空不在課表中的最愛")
        self.btn_clear_missing_fav.clicked.connect(self.on_clear_out_of_schedule_favorites)

        self.btn_history = QPushButton("檢視歷史紀錄")
        self.btn_history.clicked.connect(self.on_toggle_history)

        self.btn_best_schedule = QPushButton("最佳選課")
        self.btn_best_schedule.clicked.connect(self.on_start_best_schedule)

        fav_btn_row.addWidget(self.btn_refresh_fav)
        fav_btn_row.addWidget(self.btn_clear_fav)
        fav_btn_row.addWidget(self.btn_clear_missing_fav)
        fav_btn_row.addWidget(self.btn_history)
        fav_btn_row.addWidget(self.btn_best_schedule)
        fav_btn_row.addStretch(1)
        fav_layout.addLayout(fav_btn_row)

        self.lbl_user_file = QLabel("")
        self.lbl_user_file.setWordWrap(True)
        fav_layout.addWidget(self.lbl_user_file)

        self.gb_history = QGroupBox("歷史紀錄（選取後只讀檢視）")
        self.gb_history.setVisible(False)
        fav_layout.addWidget(self.gb_history)

        hist_v = QVBoxLayout(self.gb_history)
        self.best_progress = QProgressBar()
        self.best_progress.setRange(0, 100)
        self.best_progress.setValue(0)
        self.best_progress.setTextVisible(True)
        self.best_progress.setFormat("完成度：%p%")
        self.best_progress.setVisible(False)
        hist_v.addWidget(self.best_progress)
        self.tbl_history = QTableWidget()
        self.tbl_history.setStyleSheet("QTableWidget::item:selected { background-color: #FFB74D; color: black; }")
        self.tbl_history.setColumnCount(2)
        self.tbl_history.setHorizontalHeaderLabels(["檔案", "路徑"])
        self.tbl_history.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_history.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl_history.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_history.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl_history.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_history.itemSelectionChanged.connect(self.on_history_selected)
        hist_v.addWidget(self.tbl_history)

        hist_btn = QHBoxLayout()
        self.btn_use_history = QPushButton("使用此紀錄")
        self.btn_use_history.clicked.connect(self.on_use_selected_history)

        self.btn_back_session = QPushButton("回到目前編輯")
        self.btn_back_session.clicked.connect(self.on_back_to_session)

        hist_btn.addWidget(self.btn_use_history)
        hist_btn.addWidget(self.btn_back_session)
        hist_btn.addStretch(1)
        hist_v.addLayout(hist_btn)

        # ====== 查詢結果 ======
        self.gb_results = QGroupBox("查詢結果（左欄固定勾選我的最愛；右欄可排序）")
        lb_layout.addWidget(self.gb_results, 1)
        res_layout = QVBoxLayout(self.gb_results)

        self.results_frozen = ResultsFrozenView()
        res_layout.addWidget(self.results_frozen, 1)

        self.model_results = ResultsModel(pd.DataFrame(), self.favorites_ids)
        self.model_results.favoriteToggled.connect(self.on_result_favorite_toggled)

        self.proxy_results = QSortFilterProxyModel(self)
        self.proxy_results.setSourceModel(self.model_results)
        self.proxy_results.setSortRole(Qt.UserRole)
        self.proxy_results.setDynamicSortFilter(True)

        self.results_frozen.setModel(self.proxy_results)

        # signals
        self.ed_serial.textChanged.connect(lambda: self.schedule_search(80))
        self.ed_course_code.textChanged.connect(lambda: self.schedule_search(80))
        self.ed_cname.textChanged.connect(lambda: self.schedule_search(80))
        self.ed_teacher.textChanged.connect(lambda: self.schedule_search(80))
        self.ed_full.textChanged.connect(lambda: self.schedule_search(100))

        self.cb_dept.currentTextChanged.connect(lambda: self.schedule_search(60))
        if self.cb_dept.lineEdit() is not None:
            self.cb_dept.lineEdit().textEdited.connect(lambda _txt: self.schedule_search(80))

        self.ck_not_full.stateChanged.connect(lambda _v: self.schedule_search(0))
        self.ck_exclude_conflict.stateChanged.connect(lambda _v: self.schedule_search(0))
        self.ck_exclude_selected.stateChanged.connect(lambda _v: self.schedule_search(0))
        self.ck_show_tba.stateChanged.connect(lambda _v: self.schedule_search(0))

        self.cb_match_mode.currentIndexChanged.connect(lambda _v: self.schedule_search(0))
        self.cb_gened_core.currentTextChanged.connect(lambda _v: self.schedule_search(0))

        self.ck_gened.toggled.connect(self.on_special_option_toggled)
        self.ck_sport.toggled.connect(self.on_special_option_toggled)
        self.ck_teaching.toggled.connect(self.on_special_option_toggled)

        QTimer.singleShot(0, self._apply_fav_default_column_widths_once)

    def _configure_timetable_widget(self, widget: TimetableWidget, *, use_delegate: bool) -> None:
        widget.setEditTriggers(QTableWidget.NoEditTriggers)
        widget.setWordWrap(True)
        widget.setTextElideMode(Qt.ElideNone)
        widget.horizontalHeader().setStretchLastSection(True)
        if use_delegate:
            widget.setItemDelegate(TTTimeSelectDelegate(self, widget))
            widget.viewport().installEventFilter(self)

        widget.setSelectionMode(QAbstractItemView.NoSelection)
        widget.setSelectionBehavior(QAbstractItemView.SelectItems)
        widget.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        widget.verticalHeader().setStyleSheet(
            "QHeaderView::section { background: #F5F5F5; color: #000000; }"
        )
        widget.setStyleSheet(
            "QTableWidget { color: #000000; }"
            "QHeaderView::section { color: #000000; }"
            "QTableWidget::item { padding: 1px; }"
        )

    def _apply_default_panel_sizes(self) -> None:
        if not hasattr(self, "gb_filters"):
            return

        ms = self.gb_filters.minimumSizeHint()
        sh = self.gb_filters.sizeHint()

        minw = int(ms.width()) if ms.width() > 0 else int(sh.width())
        minh = int(ms.height()) if ms.height() > 0 else int(sh.height())

        minw = max(minw, 260)
        minh = max(minh, 260)

        self.gb_filters.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Minimum)
        self.gb_filters.setMinimumWidth(minw)
        self.gb_filters.setMaximumWidth(minw)
        self.gb_filters.setMinimumHeight(minh)

        if hasattr(self, "split_left_top") and isinstance(self.split_left_top, QSplitter):
            total_w = int(self.split_left_top.size().width()) or 800
            left_w = min(minw, max(200, total_w - 260))
            right_w = max(1, total_w - left_w)
            self.split_left_top.setSizes([left_w, right_w])

        if hasattr(self, "split_left_v") and isinstance(self.split_left_v, QSplitter):
            total_h = int(self.split_left_v.size().height()) or 800
            top_h = min(minh + 10, max(220, total_h - 260))
            bot_h = max(1, total_h - top_h)
            self.split_left_v.setSizes([top_h, bot_h])

    def _set_history_mode_layout(self, active: bool) -> None:
        if active == self._history_layout_active:
            return
        self._history_layout_active = active
        self.gb_filters.setVisible(not active)
        self.gb_results.setVisible(not active)
        preview_visible = active and self._history_preview_snapshot is not None
        if getattr(self, "gb_tt_preview", None) is not None:
            self.gb_tt_preview.setVisible(preview_visible)
        if active:
            total = max(1, self.split_main.size().width())
            fav_min = 1
            if hasattr(self, "gb_fav"):
                fav_min = max(1, self.gb_fav.minimumSizeHint().width())
            available = max(1, total - 1)
            left = min(fav_min, available)
            left = max(1, left)
            right = max(1, total - left)
            self.split_main.setSizes([left, right])
            if hasattr(self, "split_left_top"):
                self.split_left_top.setSizes([0, 1])
            if hasattr(self, "split_left_v"):
                self.split_left_v.setSizes([self.split_left_v.size().height(), 0])
            if hasattr(self, "tt_splitter"):
                self.tt_splitter.setSizes([1, 1])
        else:
            self.split_main.setSizes([1, 1])
            if hasattr(self, "split_left_top"):
                self.split_left_top.setSizes([1, 1])
            if hasattr(self, "split_left_v"):
                self.split_left_v.setSizes([1, 1])
            if hasattr(self, "tt_splitter"):
                self.tt_splitter.setSizes([1, 0])
        self._refresh_history_preview_timetable()

    # ====== Excel 載入：自動找字典序最後的 xls/xlsx ======
    def _try_autoload_default_excel(self) -> None:
        try:
            input_dir = course_input_dir_path()
        except Exception:
            input_dir = None
        found = self._first_excel_in_course_inputs(input_dir) if input_dir else None
        if found:
            try:
                self._load_excel(found)
                return
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "讀取失敗",
                    f"自動載入 user_data/course_inputs 課程檔失敗：\n{found}\n\n錯誤：{e}",
                )
                return

        self.lbl_excel.setText("課程 Excel：尚未載入（請使用「檔案 → 開啟課程 Excel…」）")
        QMessageBox.information(
            self,
            "找不到課程檔",
            "在 user_data/course_inputs 中找不到任何 .xls 或 .xlsx。\n\n"
            "請將課程檔放在 user_data/course_inputs，或使用「檔案 → 開啟課程 Excel…」選取（選擇後會自動複製到該資料夾）。",
        )

    def _ensure_course_input_file(self, candidate: str) -> str:
        src = Path(candidate)
        dest_dir = course_input_dir_path()
        dest = dest_dir / src.name
        try:
            if src.resolve() == dest.resolve():
                return str(dest)
        except Exception:
            pass
        dest_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dest))
        return str(dest)

    def _first_excel_in_course_inputs(self, folder: Optional[Path]) -> Optional[str]:
        if not folder or not folder.exists():
            return None
        try:
            entries = [
                p
                for p in folder.iterdir()
                if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"} and not p.name.startswith("~$")
            ]
        except Exception:
            return None
        if not entries:
            return None
        entries.sort(key=lambda p: p.name.casefold())
        return str(entries[0])

    def on_open_excel(self) -> None:
        start_dir = os.fspath(course_input_dir_path())
        path, _ = QFileDialog.getOpenFileName(self, "選擇課程 Excel", start_dir, "Excel Files (*.xls *.xlsx)")
        if not path:
            return
        try:
            target = self._ensure_course_input_file(path)
        except Exception as e:
            QMessageBox.critical(self, "複製失敗", f"無法將課程檔複製到 user_data/course_inputs：\n{e}")
            return
        try:
            self._load_excel(target)
        except Exception as e:
            QMessageBox.critical(self, "讀取失敗", f"讀取課程資料失敗：\n{e}")

    def on_reload_excel(self) -> None:
        if not self.excel_path or not os.path.exists(self.excel_path):
            return
        try:
            self._load_excel(self.excel_path)
        except Exception as e:
            QMessageBox.critical(self, "重新載入失敗", f"重新載入失敗：\n{e}")

    # ====== 資料索引 ======
    def _build_course_binary_index(self) -> None:
        if self.courses_df is None or self.courses_df.empty:
            self._cid_sorted = None
            self._name_sorted = None
            self._teacher_sorted = None
            self._credit_sorted = None
            self._cid_arr = None
            self._mask_lo_arr = None
            self._mask_hi_arr = None
            self._tba_arr = None
            self._dept_arr = None
            self._slots_by_cid = {}
            return

        cids = self.courses_df["_cid"].to_numpy(dtype=np.int64, copy=True)
        order = np.argsort(cids, kind="mergesort")
        self._cid_sorted = cids[order]
        self._name_sorted = self.courses_df["中文課程名稱"].to_numpy(dtype=object, copy=False)[order]
        self._teacher_sorted = self.courses_df["教師"].to_numpy(dtype=object, copy=False)[order]
        self._credit_sorted = self.courses_df["學分"].to_numpy(dtype=float, copy=False)[order]

        self._cid_arr = self.courses_df["_cid"].to_numpy(dtype=np.int64, copy=False)
        self._mask_lo_arr = self.courses_df["_mask_lo"].to_numpy(dtype="uint64", copy=False)
        self._mask_hi_arr = self.courses_df["_mask_hi"].to_numpy(dtype="uint64", copy=False)
        self._tba_arr = self.courses_df["_tba"].to_numpy(dtype=bool, copy=False)
        if "系所" in self.courses_df.columns:
            self._dept_arr = self.courses_df["系所"].to_numpy(dtype=object, copy=False)
        else:
            self._dept_arr = None
        self._slots_by_cid = {}
        for cid, slots in zip(self.courses_df["_cid"].tolist(), self.courses_df["_slots_set"].tolist()):
            try:
                cid_i = int(cid)
            except Exception:
                continue
            if isinstance(slots, set):
                self._slots_by_cid[cid_i] = slots

    def _load_excel(self, path: str) -> None:
        ensure_excel_readable(path)
        df, sheet = load_courses_auto(path)

        self.excel_path = path
        self.courses_df = df
        self.course_sheet_name = sheet

        self._build_course_binary_index()

        base_display = [c for c in df.columns if not str(c).startswith("_")]
        preferred = ["開課序號", "開課代碼", "中文課程名稱", "教師"]
        ordered = [c for c in preferred if c in base_display]
        for c in base_display:
            if c not in ordered:
                ordered.append(c)
        self.display_columns = ordered

        n = len(df)
        self.lbl_excel.setText(f"課程 Excel：{self.excel_path}（課程工作表：{self.course_sheet_name}；課程筆數：{n}）")

        self.cb_dept.blockSignals(True)
        self.cb_dept.clear()
        self.cb_dept.addItem("(全部)")
        if "系所" in df.columns:
            self._all_depts = set(df["系所"].dropna().astype(str).unique().tolist())
            for d in sorted(self._all_depts):
                self.cb_dept.addItem(str(d))
        else:
            self._all_depts = set()
        self.cb_dept.blockSignals(False)

        if self.cb_dept.completer() is not None:
            self.cb_dept.completer().setModel(self.cb_dept.model())

        self._refresh_user_selector()

        self.filtered_df = df.loc[:, self.display_columns]
        self.model_results.set_df(self.filtered_df)
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.schedule_search(0)

    def _refresh_user_selector(self, prefer_select: Optional[str] = None) -> None:
        if not self.excel_path:
            return
        try:
            _ = user_root_dir(self.excel_path)
        except Exception:
            return

        try:
            users = list_all_users(self.excel_path)
        except Exception:
            users = []

        cur_text = prefer_select or self.cb_users.currentText().strip()

        self.cb_users.blockSignals(True)
        self.cb_users.clear()
        self.cb_users.addItem("(未選擇)")
        for u in users:
            self.cb_users.addItem(u)
        self.cb_users.blockSignals(False)

        target = prefer_select or cur_text
        if target:
            idx = self.cb_users.findText(target, Qt.MatchExactly)
            self.cb_users.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            self.cb_users.setCurrentIndex(0)

        if self.cb_users.completer() is not None:
            self.cb_users.completer().setModel(self.cb_users.model())

    # ====== 時間選取（拖曳表格格子） ======
    def _calc_show_days(self) -> List[str]:
        return ["一", "二", "三", "四", "五", "六"] if self.show_saturday else ["一", "二", "三", "四", "五"]

    def _rebuild_slot_masks(self) -> None:
        self.show_days = self._calc_show_days()
        self._slot_mask_lo = []
        self._slot_mask_hi = []
        for d in self.show_days:
            row_lo = []
            row_hi = []
            for p in PERIODS:
                mlo, mhi = slot_to_mask(d, p)
                row_lo.append(np.uint64(mlo))
                row_hi.append(np.uint64(mhi))
            self._slot_mask_lo.append(row_lo)
            self._slot_mask_hi.append(row_hi)

    def _clear_saturday_selection_bits(self) -> None:
        lo = np.uint64(self._sel_lo)
        hi = np.uint64(self._sel_hi)
        for p in PERIODS:
            mlo, mhi = slot_to_mask("六", p)
            lo &= (~np.uint64(mlo))
            hi &= (~np.uint64(mhi))
        self._sel_lo = np.uint64(lo)
        self._sel_hi = np.uint64(hi)

    def tt_day_idx_from_col(self, col: int) -> Optional[int]:
        if col < 0 or col >= len(self._tt_col_day_idx):
            return None
        di = self._tt_col_day_idx[col]
        return int(di) if 0 <= di < len(self.show_days) else None

    def tt_is_time_selected(self, day_idx: int, row: int) -> bool:
        if not (0 <= day_idx < len(self.show_days) and 0 <= row < len(PERIODS)):
            return False
        mlo = self._slot_mask_lo[day_idx][row]
        mhi = self._slot_mask_hi[day_idx][row]
        return (self._sel_lo & mlo) != 0 or (self._sel_hi & mhi) != 0

    def tt_cell_locked(self, row: int, col: int) -> bool:
        if row < 0 or col < 0:
            return False
        if not self._tt_locked_matrix:
            return False
        if row >= len(self._tt_locked_matrix):
            return False
        if col >= len(self._tt_locked_matrix[row]):
            return False
        return bool(self._tt_locked_matrix[row][col])

    def tt_cell_has_selector_box(self, row: int, col: int) -> bool:
        di = self.tt_day_idx_from_col(col)
        if di is None:
            return False
        first_col = self._tt_first_lane_col.get(di, None)
        return first_col == col and 0 <= row < len(PERIODS)

    def _rebuild_tt_first_lane_cols(self) -> None:
        self._tt_first_lane_col = {}
        for c, di in enumerate(self._tt_col_day_idx):
            if di not in self._tt_first_lane_col:
                self._tt_first_lane_col[int(di)] = c

    def schedule_search(self, delay_ms: int = 60) -> None:
        if self.courses_df is None:
            return
        delay_ms = int(max(0, delay_ms))
        if delay_ms == 0:
            self._do_search_now()
            return
        self._search_timer.start(delay_ms)

    def _do_search_now(self) -> None:
        self.on_search()

    def eventFilter(self, watched, event):
        if watched is self.tbl_tt.viewport():
            et = event.type()

            # 你要求：拖曳表格格子 = 拖曳小方框
            if et == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                idx = self.tbl_tt.indexAt(event.pos())
                if not idx.isValid():
                    return False
                r, c = idx.row(), idx.column()
                di = self.tt_day_idx_from_col(c)
                if di is None:
                    return False

                cur_selected = self.tt_is_time_selected(di, r)
                self._tt_drag_state = (not cur_selected)
                self._tt_dragging = True
                self._tt_drag_start_day = int(di)
                self._tt_drag_start_row = int(r)
                self._tt_drag_base_lo = np.uint64(self._sel_lo)
                self._tt_drag_base_hi = np.uint64(self._sel_hi)
                initial_rect = (int(r), int(r), int(di), int(di))
                self._tt_drag_initial_rect = initial_rect
                self._tt_drag_last_rect = initial_rect
                self._tt_drag_has_moved = False

                self._apply_tt_drag_rect(int(r), int(r), int(di), int(di))
                self.tbl_tt.viewport().update()
                return True

            if et == QEvent.MouseMove and self._tt_dragging:
                if not (event.buttons() & Qt.LeftButton):
                    return True

                idx = self.tbl_tt.indexAt(event.pos())
                if not idx.isValid():
                    return True

                r, c = idx.row(), idx.column()
                di = self.tt_day_idx_from_col(c)
                if di is None:
                    return True

                sr = self._tt_drag_start_row
                sd = self._tt_drag_start_day
                if sr is None or sd is None:
                    return True

                r0, r1 = (sr, r) if sr <= r else (r, sr)
                d0, d1 = (sd, di) if sd <= di else (di, sd)

                rect = (int(r0), int(r1), int(d0), int(d1))
                if self._tt_drag_last_rect == rect:
                    return True
                self._tt_drag_last_rect = rect

                self._apply_tt_drag_rect(int(r0), int(r1), int(d0), int(d1))
                if self._tt_drag_initial_rect is not None and rect != self._tt_drag_initial_rect:
                    self._tt_drag_has_moved = True
                if self._tt_drag_has_moved:
                    self.schedule_search(0)
                self.tbl_tt.viewport().update()
                return True

            if et == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton and self._tt_dragging:
                if (
                    self._tt_drag_has_moved
                    and self._tt_drag_initial_rect is not None
                    and self._tt_drag_last_rect == self._tt_drag_initial_rect
                ):
                    self._sel_lo = np.uint64(self._tt_drag_base_lo)
                    self._sel_hi = np.uint64(self._tt_drag_base_hi)
                    self.tbl_tt.viewport().update()
                self._tt_dragging = False
                self._tt_drag_state = None
                self._tt_drag_start_day = None
                self._tt_drag_start_row = None
                self._tt_drag_last_rect = None
                self._tt_drag_initial_rect = None
                self._tt_drag_has_moved = False
                self.schedule_search(0)
                return True

        return super().eventFilter(watched, event)

    def _apply_tt_drag_rect(self, r0: int, r1: int, d0: int, d1: int) -> None:
        state = bool(self._tt_drag_state)
        lo = np.uint64(self._tt_drag_base_lo)
        hi = np.uint64(self._tt_drag_base_hi)

        r0 = max(0, min(len(PERIODS) - 1, r0))
        r1 = max(0, min(len(PERIODS) - 1, r1))
        d0 = max(0, min(len(self.show_days) - 1, d0))
        d1 = max(0, min(len(self.show_days) - 1, d1))

        for di in range(d0, d1 + 1):
            for rr in range(r0, r1 + 1):
                mlo = self._slot_mask_lo[di][rr]
                mhi = self._slot_mask_hi[di][rr]
                if state:
                    lo |= mlo
                    hi |= mhi
                else:
                    lo &= (~mlo)
                    hi &= (~mhi)

        self._sel_lo = np.uint64(lo)
        self._sel_hi = np.uint64(hi)

    # ====== 其餘功能（查詢/最愛/課表刷新/儲存/歷史等） ======
    # 下面為保持完整功能，基本沿用你原本 main.py 的邏輯，只做必要搬移與少量調整。

    def _mark_favorites_dirty(self) -> None:
        self._favorites_sorted_dirty = True

    def _mark_included_dirty(self) -> None:
        self._included_sorted_dirty = True

    def _mark_locked_dirty(self) -> None:
        self._locked_sorted_dirty = True

    def _get_favorites_sorted(self) -> np.ndarray:
        if self._favorites_sorted_dirty:
            self._favorites_sorted_cache = sorted_array_from_set_int(self.favorites_ids)
            self._favorites_sorted_dirty = False
        return self._favorites_sorted_cache

    def _get_included_sorted(self) -> np.ndarray:
        if self._included_sorted_dirty:
            self.included_ids |= self.locked_ids
            self._included_sorted_cache = sorted_array_from_set_int(self.included_ids)
            self._included_sorted_dirty = False
        return self._included_sorted_cache

    def _get_locked_sorted(self) -> np.ndarray:
        if self._locked_sorted_dirty:
            self._locked_sorted_cache = sorted_array_from_set_int(self.locked_ids)
            self._locked_sorted_dirty = False
        return self._locked_sorted_cache

    def _favorites_has(self, cid: int) -> bool:
        arr = self._get_favorites_sorted()
        if arr.size == 0:
            return False
        x = np.int64(int(cid))
        p = int(np.searchsorted(arr, x, side="left"))
        return 0 <= p < arr.size and int(arr[p]) == int(x)

    def _included_has(self, cid: int) -> bool:
        arr = self._get_included_sorted()
        if arr.size == 0:
            return False
        x = np.int64(int(cid))
        p = int(np.searchsorted(arr, x, side="left"))
        return 0 <= p < arr.size and int(arr[p]) == int(x)

    def _locked_has(self, cid: int) -> bool:
        arr = self._get_locked_sorted()
        if arr.size == 0:
            return False
        x = np.int64(int(cid))
        p = int(np.searchsorted(arr, x, side="left"))
        return 0 <= p < arr.size and int(arr[p]) == int(x)

    def _sync_toggle_buttons(self, checked: bool, *buttons: Optional[QPushButton]) -> None:
        for btn in buttons:
            if btn is None:
                continue
            with QSignalBlocker(btn):
                btn.setChecked(checked)

    def _set_show_saturday(self, checked: bool) -> None:
        if self.show_saturday == checked:
            self._sync_toggle_buttons(checked, self.btn_show_sat, self.btn_show_sat_preview)
            return
        self.show_saturday = checked
        if not self.show_saturday:
            self._clear_saturday_selection_bits()
        self._rebuild_slot_masks()
        self._sync_toggle_buttons(checked, self.btn_show_sat, self.btn_show_sat_preview)
        self._refresh_timetable()
        self.tbl_tt.viewport().update()
        self.schedule_search(0)

    def _set_show_time(self, checked: bool) -> None:
        if self.show_time == checked:
            self._sync_toggle_buttons(checked, self.btn_show_time, self.btn_show_time_preview)
            return
        self.show_time = checked
        self._sync_toggle_buttons(checked, self.btn_show_time, self.btn_show_time_preview)
        self._refresh_timetable()

    def _timetable_min_row_height(self) -> int:
        fm = QFontMetrics(self.tbl_tt.font())
        line_h = max(1, int(fm.lineSpacing()))
        return int(line_h * 2 + 8)

    def _apply_timetable_row_heights(self) -> None:
        widgets = [self.tbl_tt]
        if self.tbl_tt_preview is not None and getattr(self, "gb_tt_preview", None) and self.gb_tt_preview.isVisible():
            widgets.append(self.tbl_tt_preview)
        min_h = self._timetable_min_row_height()
        for widget in widgets:
            widget.resizeRowsToContents()
            for r in range(widget.rowCount()):
                if widget.rowHeight(r) < min_h:
                    widget.setRowHeight(r, min_h)

        if len(widgets) <= 1:
            return

        common_rows = min(widget.rowCount() for widget in widgets)
        if common_rows == 0:
            return

        for r in range(common_rows):
            target_height = min_h
            for widget in widgets:
                target_height = max(target_height, widget.rowHeight(r))
            for widget in widgets:
                if widget.rowHeight(r) != target_height:
                    widget.setRowHeight(r, target_height)

    def _sync_timetable_font_size(self, source: Optional[TimetableWidget]) -> None:
        if source is None:
            return
        target: Optional[TimetableWidget]
        base_size: int
        if source is self.tbl_tt and self.tbl_tt_preview is not None:
            target = self.tbl_tt_preview
            base_size = self.tbl_tt.point_size
        elif source is self.tbl_tt_preview:
            target = self.tbl_tt
            base_size = self.tbl_tt_preview.point_size
        else:
            return
        if target is None:
            return
        if target.point_size == base_size:
            return
        target.set_point_size(base_size, emit_zoom=False)

    def _on_tt_zoom_changed(self, _size: int) -> None:
        sender_widget = self.sender()
        if isinstance(sender_widget, TimetableWidget):
            self._sync_timetable_font_size(sender_widget)
        self._apply_timetable_row_heights()

    def on_clear_time_selection(self) -> None:
        self._sel_lo = np.uint64(0)
        self._sel_hi = np.uint64(0)
        self.tbl_tt.viewport().update()
        self.schedule_search(0)

    def on_clear_all_conditions(self) -> None:
        self._search_timer.stop()

        blockers = []

        def _blk(obj):
            prev = obj.blockSignals(True)
            blockers.append((obj, prev))

        _blk(self.ed_serial)
        _blk(self.ed_course_code)
        _blk(self.ed_cname)
        _blk(self.ed_teacher)
        _blk(self.ed_full)
        _blk(self.cb_dept)
        _blk(self.cb_match_mode)
        _blk(self.ck_gened)
        _blk(self.ck_sport)
        _blk(self.ck_teaching)
        _blk(self.ck_not_full)
        _blk(self.ck_exclude_conflict)
        _blk(self.ck_exclude_selected)
        _blk(self.ck_show_tba)
        _blk(self.cb_gened_core)

        self.ed_serial.clear()
        self.ed_course_code.clear()
        self.ed_cname.clear()
        self.ed_teacher.clear()
        self.ed_full.clear()

        self.cb_dept.setCurrentIndex(0)
        if self.cb_dept.lineEdit() is not None:
            self.cb_dept.lineEdit().clear()

        self.ck_gened.setChecked(False)
        self.ck_sport.setChecked(False)
        self.ck_teaching.setChecked(False)

        self.ck_not_full.setChecked(False)
        self.ck_exclude_conflict.setChecked(False)
        self.ck_exclude_selected.setChecked(False)
        self.ck_show_tba.setChecked(False)

        self.cb_match_mode.setCurrentIndex(1)
        self.cb_gened_core.setCurrentText("所有通識")
        self.stk_gened_core.setCurrentIndex(0)

        for obj, prev in blockers:
            obj.blockSignals(prev)

        self.on_clear_time_selection()

    def on_special_option_toggled(self, checked: bool) -> None:
        sender = self.sender()
        if not isinstance(sender, QCheckBox):
            self.schedule_search(0)
            return

        if checked:
            others = []
            if sender is not self.ck_gened:
                others.append(self.ck_gened)
            if sender is not self.ck_sport:
                others.append(self.ck_sport)
            if sender is not self.ck_teaching:
                others.append(self.ck_teaching)

            for ck in others:
                ck.blockSignals(True)
                ck.setChecked(False)
                ck.blockSignals(False)

        self.stk_gened_core.setCurrentIndex(1 if self.ck_gened.isChecked() else 0)
        self.schedule_search(0)

    # ====== 登入/歷史/最愛/儲存/查詢/課表刷新：為避免回覆爆量，保留原邏輯結構 ======
    # 注意：你若要我把這段也再細拆到更多檔案，我也可以，但目前已符合多檔案模式與你本次三個改動。

    # ---- 以下方法與你原版一致（僅做 import/呼叫路徑調整），因此不再逐段加註解 ----

    def _resolve_username_from_ui(self) -> Optional[str]:
        new_name = (self.ed_new_user.text() or "").strip()
        if new_name:
            return new_name

        sel = (self.cb_users.currentText() or "").strip()
        if sel and sel != "(未選擇)":
            return sel

        if self.cb_users.lineEdit() is not None:
            typed = (self.cb_users.lineEdit().text() or "").strip()
            if typed and typed != "(未選擇)":
                return typed

        return None

    def _replace_sets(self, fav: Set[int], inc: Set[int], lock: Set[int], seq: Optional[Dict[int, int]] = None) -> None:
        self.favorites_ids.clear()
        self.favorites_ids.update(set(fav))

        self.locked_ids.clear()
        self.locked_ids.update(set(lock) & set(fav))

        self.included_ids.clear()
        self.included_ids.update((set(inc) & set(fav)) | set(self.locked_ids))

        self.fav_seq.clear()
        if seq:
            for cid in self.favorites_ids:
                if int(cid) in seq:
                    self.fav_seq[int(cid)] = int(seq[int(cid)])

        if len(self.fav_seq) != len(self.favorites_ids):
            missing = [int(cid) for cid in self.favorites_ids if int(cid) not in self.fav_seq]
            missing.sort()
            start = 1
            if self.fav_seq:
                start = max(self.fav_seq.values()) + 1
            for cid in missing:
                self.fav_seq[cid] = start
                start += 1

        self._fav_seq_next = (max(self.fav_seq.values()) + 1) if self.fav_seq else 1
        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()
        self._fav_sort_section = self.FAV_COL_RANK
        self._fav_sort_order = Qt.AscendingOrder

    def _ensure_seq_for_new_fav(self, cid: int) -> None:
        cid_i = int(cid)
        if cid_i not in self.fav_seq:
            self.fav_seq[cid_i] = int(self._fav_seq_next)
            self._fav_seq_next += 1

    def on_login(self) -> None:
        if self.courses_df is None:
            QMessageBox.warning(self, "尚未載入", "請先載入課程 Excel。")
            return

        uname = self._resolve_username_from_ui()
        if not uname:
            QMessageBox.warning(self, "需要使用者名稱", "請在「新增使用者」輸入新名稱，或於「選擇使用者」選擇既有使用者。")
            return

        uname_safe = sanitize_folder_name(uname)
        self.username = uname_safe

        self.user_dir_path = user_dir(self.excel_path, self.username)

        inherited_fav: Set[int] = set()
        inherited_inc: Set[int] = set()
        inherited_lock: Set[int] = set()
        inherited_seq: Dict[int, int] = {}
        try:
            history = list_user_history_files(self.user_dir_path)
            if history:
                fav, inc, seq, lock = load_user_file(history[0])
                inherited_fav = set(fav)
                inherited_inc = set(inc)
                inherited_seq = dict(seq)
                inherited_lock = set(lock)
        except Exception:
            pass

        ts = time.strftime("%Y%m%d_%H%M%S")
        self.session_file_path = unique_login_file_path(self.excel_path, self.username, ts)
        self._replace_sets(inherited_fav, inherited_inc, inherited_lock, inherited_seq)

        try:
            save_user_file(
                self.session_file_path,
                self.username,
                self.favorites_ids,
                self._get_included_sorted(),
                self._get_locked_sorted(),
                self.fav_seq,
                self.courses_df,
            )
        except Exception as e:
            QMessageBox.critical(self, "建立使用者檔案失敗", f"無法建立本次登入檔案：\n{self.session_file_path}\n\n錯誤：{e}")
            self.session_file_path = ""
            return

        self._refresh_user_selector(prefer_select=self.username)
        self.ed_new_user.blockSignals(True)
        self.ed_new_user.clear()
        self.ed_new_user.blockSignals(False)

        self._session_fav_backup = None
        self._session_inc_backup = None
        self._session_lock_backup = None
        self._session_seq_backup = None
        self._history_selected_file = ""
        self._set_readonly(False)

        self.lbl_user.setText(f"使用者：{self.username}")
        self._set_user_file_label()
        self._refresh_history_list()

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
        self.schedule_search(0)

    def _set_readonly(self, readonly: bool) -> None:
        self.readonly_mode = bool(readonly)
        self.model_results.set_readonly(self.readonly_mode)

        self.lbl_readonly.setVisible(self.readonly_mode)
        self.tbl_fav.setDisabled(self.readonly_mode)
        self.btn_clear_fav.setDisabled(self.readonly_mode)
        self._update_fav_drag_state()

    def _set_user_file_label(self) -> None:
        if not self.username:
            self.lbl_user_file.setText("")
            return
        msg = f"使用者資料夾：{self.user_dir_path}"
        if self.session_file_path:
            msg += f"\n本次登入檔案：{self.session_file_path}"
        self.lbl_user_file.setText(msg)

    def _compress_join_order_map(self) -> Dict[int, int]:
        items = []
        for cid in self.favorites_ids:
            cid_i = int(cid)
            seq = int(self.fav_seq.get(cid_i, 10**12))
            items.append((seq, cid_i))
        items.sort()
        total_items = len(items)

        out: Dict[int, int] = {}
        k = 1
        for _seq, cid in items:
            out[int(cid)] = k
            k += 1
        return out

    def _course_name_by_id(self, cid: int) -> str:
        if self._cid_sorted is None or self._name_sorted is None:
            return ""
        x = np.int64(int(cid))
        pos = int(np.searchsorted(self._cid_sorted, x, side="left"))
        if 0 <= pos < self._cid_sorted.size and int(self._cid_sorted[pos]) == int(x):
            return str(self._name_sorted[pos] or "").strip()
        return ""

    def _teacher_by_id(self, cid: int) -> str:
        if self._cid_sorted is None or self._teacher_sorted is None:
            return ""
        x = np.int64(int(cid))
        pos = int(np.searchsorted(self._cid_sorted, x, side="left"))
        if 0 <= pos < self._cid_sorted.size and int(self._cid_sorted[pos]) == int(x):
            return str(self._teacher_sorted[pos] or "").strip()
        return ""

    def _credit_by_id(self, cid: int) -> float:
        if self._cid_sorted is None or self._credit_sorted is None:
            return 0.0
        x = np.int64(int(cid))
        pos = int(np.searchsorted(self._cid_sorted, x, side="left"))
        if 0 <= pos < self._cid_sorted.size and int(self._cid_sorted[pos]) == int(x):
            v = self._credit_sorted[pos]
            try:
                if np.isnan(v):
                    return 0.0
            except Exception:
                pass
            try:
                return float(v)
            except Exception:
                return 0.0
        return 0.0

    def _refresh_favorites_table(self) -> None:
        # Before clearing, save the vertical scroll position
        scrollbar = self.tbl_fav.verticalScrollBar()
        scroll_pos = scrollbar.value()
        updates_enabled = self.tbl_fav.updatesEnabled()

        # Before clearing, save the current visual order of course IDs
        visual_order = []
        for r in range(self.tbl_fav.rowCount()):
            id_item = self.tbl_fav.item(r, self.FAV_COL_ID)
            if id_item:
                cid_data = id_item.data(Qt.UserRole)
                if cid_data is not None:
                    visual_order.append(int(cid_data))

        sorting_enabled = self.tbl_fav.isSortingEnabled()
        self.tbl_fav.setSortingEnabled(False)
        
        was_drag_enabled = self.tbl_fav.is_drag_enabled()
        self.tbl_fav.set_drag_enabled(False)

        self.tbl_fav.setUpdatesEnabled(False)
        self.tbl_fav.blockSignals(True)
        self.tbl_fav.setRowCount(0)

        join_rank = self._compress_join_order_map()

        # If visual order was captured, use it only when it matches current favorites.
        # Otherwise, fall back to sorting by rank from favorites_ids.
        fav_set = set(int(x) for x in self.favorites_ids)
        cids_to_render = [cid for cid in visual_order if cid in fav_set]
        if len(cids_to_render) != len(fav_set):
            items = []
            for cid in self.favorites_ids:
                cid_i = int(cid)
                seq = int(self.fav_seq.get(cid_i, 10**12))
                items.append((seq, cid_i))
            items.sort()
            cids_to_render = [cid for _, cid in items]

        move_controls_enabled = self._is_fav_join_sort_active() and not self.readonly_mode
        total_items = len(cids_to_render)

        for r_idx, cid_i in enumerate(cids_to_render):
            r = self.tbl_fav.rowCount()
            self.tbl_fav.insertRow(r)

            handle_widget = self._build_move_widget(cid_i, r, total_items, enabled=move_controls_enabled)
            self.tbl_fav.setCellWidget(r, self.FAV_COL_HANDLE, handle_widget)

            is_lock = cid_i in self.locked_ids

            in_course = True if is_lock else self._included_has(cid_i)
            ck_item = IntSortItem("", 1 if in_course else 0)
            if is_lock:
                ck_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                ck_item.setCheckState(Qt.Checked)
            else:
                ck_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                ck_item.setCheckState(Qt.Checked if in_course else Qt.Unchecked)
            ck_item.setData(FAV_CID_ROLE, cid_i)
            self.tbl_fav.setItem(r, self.FAV_COL_SCHEDULE, ck_item)

            lock_item = IntSortItem("", 1 if is_lock else 0)
            lock_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            lock_item.setCheckState(Qt.Checked if is_lock else Qt.Unchecked)
            lock_item.setData(FAV_CID_ROLE, cid_i)
            self.tbl_fav.setItem(r, self.FAV_COL_LOCK, lock_item)

            id_item = IntSortItem(format_cid4(cid_i), cid_i)
            id_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            id_item.setData(Qt.UserRole, cid_i)
            self.tbl_fav.setItem(r, self.FAV_COL_ID, id_item)

            name_item = QTableWidgetItem(self._course_name_by_id(cid_i))
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.tbl_fav.setItem(r, self.FAV_COL_NAME, name_item)

            t_item = QTableWidgetItem(self._teacher_by_id(cid_i))
            t_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.tbl_fav.setItem(r, self.FAV_COL_TEACHER, t_item)

            cr = self._credit_by_id(cid_i)
            cr_text = "" if cr == 0.0 else (str(int(cr)) if abs(cr - round(cr)) < 1e-9 else f"{cr:g}")
            cr_item = FloatSortItem(cr_text, cr)
            cr_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.tbl_fav.setItem(r, self.FAV_COL_CREDIT, cr_item)

            rank = int(join_rank.get(cid_i, 0))
            rank_item = IntSortItem(str(rank) if rank > 0 else "", rank)
            rank_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            rank_item.setData(FAV_CID_ROLE, cid_i)
            self.tbl_fav.setItem(r, self.FAV_COL_RANK, rank_item)

            if is_lock:
                self.tbl_fav.setCellWidget(r, self.FAV_COL_DELETE, QWidget())
            else:
                btn = QPushButton("×")
                btn.setProperty("cid", cid_i)
                btn.setFixedWidth(32)
                btn.setToolTip("刪除此最愛")
                btn.clicked.connect(self.on_delete_favorite_button_clicked)
                btn.setEnabled(not self.readonly_mode)
                self.tbl_fav.setCellWidget(r, self.FAV_COL_DELETE, btn)

        self.tbl_fav.blockSignals(False)
        self.tbl_fav.setUpdatesEnabled(updates_enabled)

        if was_drag_enabled:
            self.tbl_fav.set_drag_enabled(True)
        self._update_fav_drag_state()

        self.tbl_fav.setSortingEnabled(sorting_enabled)

        # Restore the scroll position
        scrollbar.setValue(scroll_pos)

    def _try_append_favorite_row(self, cid_i: int) -> bool:
        if getattr(self, "tbl_fav", None) is None:
            return False
        # If the table is out of sync, fall back to full refresh.
        if self.tbl_fav.rowCount() != (len(self.favorites_ids) - 1):
            return False

        sorting_enabled = self.tbl_fav.isSortingEnabled()
        updates_enabled = self.tbl_fav.updatesEnabled()
        self.tbl_fav.setUpdatesEnabled(False)
        self.tbl_fav.blockSignals(True)

        row = self.tbl_fav.rowCount()
        self.tbl_fav.insertRow(row)

        move_controls_enabled = self._is_fav_join_sort_active() and not self.readonly_mode
        total_items = row + 1

        handle_widget = self._build_move_widget(cid_i, row, total_items, enabled=move_controls_enabled)
        self.tbl_fav.setCellWidget(row, self.FAV_COL_HANDLE, handle_widget)

        is_lock = cid_i in self.locked_ids
        in_course = True if is_lock else self._included_has(cid_i)

        ck_item = IntSortItem("", 1 if in_course else 0)
        if is_lock:
            ck_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            ck_item.setCheckState(Qt.Checked)
        else:
            ck_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            ck_item.setCheckState(Qt.Checked if in_course else Qt.Unchecked)
        ck_item.setData(FAV_CID_ROLE, cid_i)
        self.tbl_fav.setItem(row, self.FAV_COL_SCHEDULE, ck_item)

        lock_item = IntSortItem("", 1 if is_lock else 0)
        lock_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        lock_item.setCheckState(Qt.Checked if is_lock else Qt.Unchecked)
        lock_item.setData(FAV_CID_ROLE, cid_i)
        self.tbl_fav.setItem(row, self.FAV_COL_LOCK, lock_item)

        id_item = IntSortItem(format_cid4(cid_i), cid_i)
        id_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        id_item.setData(Qt.UserRole, cid_i)
        self.tbl_fav.setItem(row, self.FAV_COL_ID, id_item)

        name_item = QTableWidgetItem(self._course_name_by_id(cid_i))
        name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.tbl_fav.setItem(row, self.FAV_COL_NAME, name_item)

        t_item = QTableWidgetItem(self._teacher_by_id(cid_i))
        t_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.tbl_fav.setItem(row, self.FAV_COL_TEACHER, t_item)

        cr = self._credit_by_id(cid_i)
        cr_text = "" if cr == 0.0 else (str(int(cr)) if abs(cr - round(cr)) < 1e-9 else f"{cr:g}")
        cr_item = FloatSortItem(cr_text, cr)
        cr_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.tbl_fav.setItem(row, self.FAV_COL_CREDIT, cr_item)

        rank = len(self.favorites_ids)
        rank_item = IntSortItem(str(rank), rank)
        rank_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        rank_item.setData(FAV_CID_ROLE, cid_i)
        self.tbl_fav.setItem(row, self.FAV_COL_RANK, rank_item)

        if is_lock:
            self.tbl_fav.setCellWidget(row, self.FAV_COL_DELETE, QWidget())
        else:
            btn = QPushButton("×")
            btn.setProperty("cid", cid_i)
            btn.setFixedWidth(32)
            btn.setToolTip("刪除此最愛")
            btn.clicked.connect(self.on_delete_favorite_button_clicked)
            btn.setEnabled(not self.readonly_mode)
            self.tbl_fav.setCellWidget(row, self.FAV_COL_DELETE, btn)

        self.tbl_fav.blockSignals(False)
        self.tbl_fav.setUpdatesEnabled(updates_enabled)

        if sorting_enabled:
            self.tbl_fav.sortItems(self._fav_sort_section, self._fav_sort_order)

        self._update_fav_drag_state()
        return True

    def _is_fav_join_sort_active(self) -> bool:
        return self._fav_sort_section == self.FAV_COL_RANK and self._fav_sort_order == Qt.AscendingOrder

    def _update_fav_drag_state(self) -> None:
        enabled = self._is_fav_join_sort_active() and not self.readonly_mode
        self.tbl_fav.set_drag_enabled(enabled)

    def _on_fav_sort_changed(self, section: int, order: Qt.SortOrder) -> None:
        self._fav_sort_section = section
        self._fav_sort_order = order
        self._update_fav_drag_state()

    def _build_move_widget(self, cid: int, row: int, total: int, *, enabled: bool = True) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(2, 0, 2, 0)
        layout.setSpacing(0)
        handle = QLabel("☰")
        handle.setAlignment(Qt.AlignCenter)
        layout.addWidget(handle)
        tooltip = "按住並拖曳此欄位以調整優先度" if enabled else "切換至加入順序並解除唯讀後才能調整優先度"
        container.setToolTip(tooltip)
        return container

    def _favorite_order_list(self) -> List[int]:
        order: List[Tuple[int, int]] = []
        for cid in self.favorites_ids:
            cid_i = int(cid)
            seq = int(self.fav_seq.get(cid_i, 10**12))
            order.append((seq, cid_i))
        order.sort()
        return [cid for _seq, cid in order]

    def _move_favorite(self, cid: int, direction: int) -> None:
        order = self._favorite_order_list()
        try:
            idx = order.index(cid)
        except ValueError:
            return
        target = idx + direction
        if target < 0 or target >= len(order):
            return
        new_order = order.copy()
        new_order[idx], new_order[target] = new_order[target], new_order[idx]
        for rank, cid_i in enumerate(new_order, start=1):
            self.fav_seq[cid_i] = rank
        self._fav_seq_next = len(new_order) + 1

        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()
        self._fav_sort_section = self.FAV_COL_RANK
        self._fav_sort_order = Qt.AscendingOrder

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
        self.schedule_autosave(250)
        self.schedule_search(0)

    def _sync_fav_seq_from_view(self) -> List[int]:
        order: List[int] = []
        self.tbl_fav.blockSignals(True)
        seq = 1
        for row in range(self.tbl_fav.rowCount()):
            id_item = self.tbl_fav.item(row, self.FAV_COL_ID)
            if id_item is None:
                continue
            cid_data = id_item.data(Qt.UserRole)
            if cid_data is None:
                continue
            try:
                cid_i = int(cid_data)
            except Exception:
                continue
            order.append(cid_i)
            self.fav_seq[cid_i] = seq
            rank_item = self.tbl_fav.item(row, self.FAV_COL_RANK)
            if rank_item is not None:
                rank_item.setText(str(seq))
                rank_item.setData(FAV_CID_ROLE, cid_i)
            seq += 1
        self._fav_seq_next = seq
        self.tbl_fav.blockSignals(False)
        return order

    def _apply_fav_default_column_widths_once(self) -> None:
        if self._fav_default_width_applied:
            return
        self._fav_default_width_applied = True
        hh = self.tbl_fav.horizontalHeader()
        fm = QFontMetrics(hh.font())
        one_char = max(1, fm.horizontalAdvance("字"))
        two_char = max(1, fm.horizontalAdvance("字字"))
        three_char = max(1, fm.horizontalAdvance("字字字"))
        four_char = max(1, fm.horizontalAdvance("字字字字"))
        for c in range(self.tbl_fav.columnCount()):
            it = self.tbl_fav.horizontalHeaderItem(c)
            title = it.text() if it else ""
            if title == "拖曳排序":
                w = three_char + 8
            elif title == "學分數":
                w = two_char + 8
            elif title == "優先度":
                w = four_char + 8
            elif title in ("課表", "鎖定", "開課序號", "中文課程名稱"):
                w = fm.horizontalAdvance(title) + one_char + 18 - one_char
                w = max(w, 40)
            else:
                w = fm.horizontalAdvance(title) + one_char + 18
                w = max(w, 40)
            self.tbl_fav.setColumnWidth(c, int(w))

    def on_delete_favorite_button_clicked(self) -> None:
        if self.readonly_mode:
            return
        btn = self.sender()
        if btn is None:
            return
        cid = btn.property("cid")
        if cid is None:
            return
        cid_i = int(cid)

        if cid_i in self.locked_ids:
            return

        self.favorites_ids.discard(cid_i)
        self.included_ids.discard(cid_i)
        self.locked_ids.discard(cid_i)
        self.fav_seq.pop(cid_i, None)
        self._fav_seq_next = (max(self.fav_seq.values()) + 1) if self.fav_seq else 1

        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()

        self.schedule_autosave(250)
        self.schedule_search(0)

    def on_fav_cell_changed(self, row: int, col: int) -> None:
        if self.readonly_mode:
            return

        if col not in (self.FAV_COL_SCHEDULE, self.FAV_COL_LOCK):
            return

        item = self.tbl_fav.item(row, col)
        if not item:
            return

        cid = item.data(FAV_CID_ROLE)
        if cid is None:
            return
        cid_i = int(cid)

        if not self._favorites_has(cid_i):
            return

        if col == self.FAV_COL_SCHEDULE:
            if cid_i in self.locked_ids:
                self.tbl_fav.blockSignals(True)
                item.setCheckState(Qt.Checked)
                self.tbl_fav.blockSignals(False)
                return

            if item.checkState() == Qt.Checked:
                self.included_ids.add(cid_i)
            else:
                self.included_ids.discard(cid_i)
            self._mark_included_dirty()

        elif col == self.FAV_COL_LOCK:
            if item.checkState() == Qt.Checked:
                self.locked_ids.add(cid_i)
                self.included_ids.add(cid_i)
                self._mark_included_dirty()
            else:
                self.locked_ids.discard(cid_i)
            self._mark_locked_dirty()

        self._refresh_favorites_table()
        self._refresh_timetable()

        self.schedule_autosave(250)
        self.schedule_search(0)

    def _on_fav_drag_selection_finished(self, column: int, affected_rows: set) -> None:
        if self.readonly_mode:
            return

        if column not in (self.FAV_COL_SCHEDULE, self.FAV_COL_LOCK):
            return
            
        changed = False
        for row in affected_rows:
            item = self.tbl_fav.item(row, column)
            if not item:
                continue

            cid_data = item.data(FAV_CID_ROLE)
            if cid_data is None:
                continue
            
            cid_i = int(cid_data)
            if not self._favorites_has(cid_i):
                continue
            
            is_checked = item.checkState() == Qt.Checked
            changed = True

            if column == self.FAV_COL_SCHEDULE:
                if is_checked:
                    self.included_ids.add(cid_i)
                else:
                    # Do not uncheck if it's locked
                    if cid_i not in self.locked_ids:
                        self.included_ids.discard(cid_i)

            elif column == self.FAV_COL_LOCK:
                if is_checked:
                    self.locked_ids.add(cid_i)
                    self.included_ids.add(cid_i) # Locking also includes it in schedule
                else:
                    self.locked_ids.discard(cid_i)

        if changed:
            self._mark_included_dirty()
            self._mark_locked_dirty()
            
            # Refresh everything once after all changes
            self._refresh_favorites_table()
            self._refresh_timetable()
            self.schedule_autosave(250)
            self.schedule_search(0)


    def _on_favorites_reordered(self, new_order: List[int]) -> None:
        if not new_order:
            return
        
        # 驗證：確保所有 new_order 中的課程都在 favorites_ids 中
        valid_order = []
        for cid in new_order:
            try:
                cid_i = int(cid)
            except Exception:
                continue
            if cid_i in self.favorites_ids:
                valid_order.append(cid_i)
        
        # 如果有效順序為空，則不進行更新
        if not valid_order:
            return
        
        # 檢查新順序是否與當前順序不同（如果相同則表示沒有移動）
        current_order = self._favorite_order_list()
        if valid_order == current_order:
            # 順序未改變，不進行任何操作
            return
        
        # 防止表格刷新時拖曳被打斷：先暫時禁用拖曳
        was_drag_enabled = self.tbl_fav.is_drag_enabled()
        self.tbl_fav.set_drag_enabled(False)
        
        try:
            seq = 1
            for cid_i in valid_order:
                self.fav_seq[cid_i] = seq
                seq += 1
            self._fav_seq_next = seq
            self._mark_favorites_dirty()
            self._fav_sort_section = self.FAV_COL_RANK
            self._fav_sort_order = Qt.AscendingOrder
            
            self._refresh_favorites_table()
            self.model_results.notify_favorites_changed()
            self.proxy_results.invalidate()
            self.schedule_autosave(250)
            self.schedule_search(0)
        finally:
            # 恢復拖曳狀態
            if was_drag_enabled:
                self.tbl_fav.set_drag_enabled(True)

    def _on_fav_drop_completed(self, changed: bool) -> None:
        # If the drop did not change the underlying favorites order, the
        # view can still be visually corrupted by the drag. Force a
        # lightweight refresh to restore correct visuals.
        if not changed:
            # Only refresh the favorites table UI; this rebuild is cheap.
            self._refresh_favorites_table()

    def on_clear_out_of_schedule_favorites(self) -> None:
        if self.readonly_mode:
            return
        missing = {cid for cid in self.favorites_ids if cid not in self.included_ids and cid not in self.locked_ids}
        if not missing:
            QMessageBox.information(self, "資訊", "目前沒有不在課表中的我的最愛。")
            return

        ret = QMessageBox.question(
            self,
            "清除確認",
            f"確定要清空 {len(missing)} 筆不在課表中的我的最愛嗎？",
        )
        if ret != QMessageBox.Yes:
            return

        for cid in missing:
            self.favorites_ids.discard(cid)
            self.included_ids.discard(cid)
            self.locked_ids.discard(cid)
            self.fav_seq.pop(cid, None)
        self._fav_seq_next = (max(self.fav_seq.values()) + 1) if self.fav_seq else 1

        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()

        self.schedule_autosave(250)
        self.schedule_search(0)

    def on_clear_favorites(self) -> None:
        if self.readonly_mode:
            return
        if not self.favorites_ids:
            return
        if self.locked_ids:
            QMessageBox.information(self, "不可清空", "目前存在鎖定課程，為避免誤刪，暫不允許一鍵清空。\n如需清空，請先取消鎖定。")
            return

        ret = QMessageBox.question(self, "確認", "確定要清空我的最愛嗎？")
        if ret != QMessageBox.Yes:
            return

        self.favorites_ids.clear()
        self.included_ids.clear()
        self.locked_ids.clear()
        self.fav_seq.clear()
        self._fav_seq_next = 1

        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()

        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()

        self.schedule_autosave(250)
        self.schedule_search(0)

    def on_result_favorite_toggled(self, cid: int, checked: bool) -> None:
        if not self.username or not self.session_file_path:
            QMessageBox.warning(self, "需要登入", "請先使用「新增/切換」建立或切換使用者，才能使用我的最愛。")
            self.model_results.notify_favorites_changed()
            return

        if self.readonly_mode:
            self.model_results.notify_favorites_changed()
            return

        cid_i = int(cid)
        if checked:
            self.favorites_ids.add(cid_i)
            self.included_ids.add(cid_i)
            self.locked_ids.discard(cid_i)
            self._ensure_seq_for_new_fav(cid_i)
        else:
            if cid_i in self.locked_ids:
                self.model_results.notify_favorites_changed()
                return

            self.favorites_ids.discard(cid_i)
            self.included_ids.discard(cid_i)
            self.locked_ids.discard(cid_i)
            self.fav_seq.pop(cid_i, None)
            self._fav_seq_next = (max(self.fav_seq.values()) + 1) if self.fav_seq else 1

        self.included_ids |= self.locked_ids

        self._mark_favorites_dirty()
        self._mark_included_dirty()
        self._mark_locked_dirty()
        if checked:
            if not self._try_append_favorite_row(cid_i):
                self._refresh_favorites_table()
        else:
            self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()

        self.schedule_autosave(250)
        self.schedule_search(0)

    def schedule_autosave(self, delay_ms: int = 350) -> None:
        if self.readonly_mode:
            return
        if not self.username or not self.session_file_path:
            return
        if self.courses_df is None:
            return

        self.included_ids |= self.locked_ids
        self._mark_included_dirty()

        self._save_token += 1
        self._save_latest_token = self._save_token
        self._save_pending_snapshot = (set(self.favorites_ids), set(self.included_ids), set(self.locked_ids), dict(self.fav_seq))
        self._save_pending = True

        delay_ms = int(max(0, delay_ms))
        if delay_ms == 0:
            self._autosave_now()
            return
        self._autosave_timer.start(delay_ms)

    def _autosave_now(self) -> None:
        if self.readonly_mode:
            return
        if not self._save_pending:
            return
        if self.courses_df is None or not self.username or not self.session_file_path:
            return
        if self._save_inflight:
            return

        snapshot = self._save_pending_snapshot
        if snapshot is None:
            return

        fav, inc, lock, seq = snapshot
        token = self._save_latest_token

        self._save_pending = False
        self._save_inflight = True

        included_sorted = sorted_array_from_set_int(inc)
        locked_sorted = sorted_array_from_set_int(lock)

        worker = SaveWorker(
            token,
            self.session_file_path,
            self.username,
            fav,
            included_sorted,
            locked_sorted,
            seq,
            self.courses_df,
        )
        worker.finished.connect(self._on_save_finished)
        self.threadpool.start(worker)

    def _on_save_finished(self, token: int, ok: bool, msg: str) -> None:
        self._save_inflight = False
        if self._save_pending:
            QTimer.singleShot(0, self._autosave_now)
        if not ok:
            QMessageBox.warning(self, "自動儲存失敗", f"自動儲存失敗：\n{msg}")
            return
        if token == self._save_latest_token:
            self._refresh_history_list()

    def _refresh_history_list(self) -> None:
        if self._history_mode == "best":
            return
        self.tbl_history.blockSignals(True)
        self.tbl_history.setRowCount(0)

        if not self.user_dir_path:
            self.tbl_history.blockSignals(False)
            return

        cache = load_best_schedule_cache(self.user_dir_path)
        best_names = {os.path.basename(f) for f in cache.get("files", []) if f} if cache else set()
        current_best_names = {os.path.basename(p) for p in self._best_files if p}
        session_abs = os.path.abspath(self.session_file_path) if self.session_file_path else None

        files = list_user_history_files(self.user_dir_path)
        for p in files:
            if session_abs and os.path.abspath(p) == session_abs:
                continue
            name = os.path.basename(p)
            if name in best_names or name in current_best_names:
                continue
            r = self.tbl_history.rowCount()
            self.tbl_history.insertRow(r)

            fn = name
            it0 = QTableWidgetItem(fn)
            it0.setData(Qt.UserRole, p)
            self.tbl_history.setItem(r, 0, it0)

            it1 = QTableWidgetItem(p)
            self.tbl_history.setItem(r, 1, it1)

        self.tbl_history.blockSignals(False)
        self._update_history_highlights()
        self._auto_select_history_first_row()

    def _refresh_best_schedule_list(self) -> None:
        self.tbl_history.blockSignals(True)
        self.tbl_history.setRowCount(0)

        for p in self._best_files:
            r = self.tbl_history.rowCount()
            self.tbl_history.insertRow(r)

            fn = os.path.basename(p)
            it0 = QTableWidgetItem(fn)
            it0.setData(Qt.UserRole, p)
            self.tbl_history.setItem(r, 0, it0)

            it1 = QTableWidgetItem(p)
            self.tbl_history.setItem(r, 1, it1)

        self.tbl_history.blockSignals(False)
        self._auto_select_history_first_row()

    def _auto_select_history_first_row(self) -> None:
        if self.tbl_history.rowCount() == 0 or self._history_mode is None:
            return
        self.tbl_history.blockSignals(True)
        self.tbl_history.setCurrentCell(0, 0)
        self.tbl_history.blockSignals(False)
        QTimer.singleShot(0, self.on_history_selected)

    def _capture_session_snapshot(self, *, force: bool = False) -> None:
        if not force and self._session_fav_backup is not None:
            return
        self._session_fav_backup = set(self.favorites_ids)
        self._session_inc_backup = set(self.included_ids)
        self._session_lock_backup = set(self.locked_ids)
        self._session_seq_backup = dict(self.fav_seq)
        self._history_preview_snapshot = {
            "favorites": set(self.favorites_ids),
            "included": set(self.included_ids),
            "locked": set(self.locked_ids),
            "seq": dict(self.fav_seq),
        }

    def _restore_session_state(self) -> bool:
        if (
            self._session_fav_backup is None
            or self._session_inc_backup is None
            or self._session_lock_backup is None
            or self._session_seq_backup is None
        ):
            return False

        self._replace_sets(
            self._session_fav_backup,
            self._session_inc_backup,
            self._session_lock_backup,
            self._session_seq_backup,
        )
        self._history_selected_file = ""
        self._set_readonly(False)
        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
        self.schedule_search(0)
        self._update_history_highlights()
        return True

    def _update_history_highlights(self) -> None:
        target = os.path.abspath(self._history_selected_file) if self._history_selected_file else None
        for row in range(self.tbl_history.rowCount()):
            it0 = self.tbl_history.item(row, 0)
            it1 = self.tbl_history.item(row, 1)
            brush = QBrush()
            if target and it0 is not None:
                path = it0.data(Qt.UserRole)
                if path:
                    try:
                        if os.path.abspath(str(path)) == target:
                            brush = self._history_selected_brush
                    except Exception:
                        pass
            if it0 is not None:
                it0.setBackground(brush)
            if it1 is not None:
                it1.setBackground(brush)

    def _enter_history_panel(self) -> None:
        self._history_mode = "history"
        self._best_running = False
        self._best_files = []
        self._capture_session_snapshot(force=True)
        self.gb_history.setTitle("歷史紀錄（選取後只讀檢視）")
        self.best_progress.setVisible(False)
        self.btn_use_history.setText("使用此紀錄")
        self.btn_use_history.setEnabled(True)
        self.btn_back_session.setEnabled(True)
        self.btn_back_session.setText("關閉歷史紀錄視窗")
        self.tbl_history.setEnabled(True)
        self.gb_history.setVisible(True)
        self.btn_history.setEnabled(True)
        self.btn_best_schedule.setEnabled(False)
        self._refresh_history_list()
        self._set_history_mode_layout(True)

    def _enter_best_schedule_running(self) -> None:
        self._capture_session_snapshot(force=True)
        self._history_mode = "best"
        self._best_running = True
        self.gb_history.setTitle("最佳選課（選取後只讀檢視）")
        self.best_progress.setRange(0, 100)
        self.best_progress.setValue(0)
        self.best_progress.setVisible(True)
        self.tbl_history.setRowCount(0)
        self.tbl_history.setEnabled(False)
        self.btn_use_history.setText("中斷最佳選課")
        self.btn_use_history.setEnabled(True)
        self.btn_back_session.setEnabled(False)
        self.btn_back_session.setText("關閉最佳選課視窗")
        self.gb_history.setVisible(True)
        self.btn_history.setEnabled(False)
        self.btn_best_schedule.setEnabled(False)
        self._set_history_mode_layout(True)

    def _enter_best_schedule_results(self) -> None:
        self._capture_session_snapshot(force=True)
        self._history_mode = "best"
        self._best_running = False
        self.gb_history.setTitle("最佳選課（選取後只讀檢視）")
        self.best_progress.setVisible(False)
        self.btn_use_history.setText("使用此規劃")
        self.btn_use_history.setEnabled(True)
        self.btn_back_session.setEnabled(True)
        self.btn_back_session.setText("關閉最佳選課視窗")
        self.tbl_history.setEnabled(True)
        self.gb_history.setVisible(True)
        self.btn_history.setEnabled(False)
        self.btn_best_schedule.setEnabled(False)
        self._refresh_best_schedule_list()
        self._set_history_mode_layout(True)

    def _close_history_panel(self) -> None:
        self.gb_history.setVisible(False)
        self._set_history_mode_layout(False)
        self._history_mode = None
        self._best_running = False
        self._best_files = []
        self.btn_history.setEnabled(True)
        self.btn_best_schedule.setEnabled(True)
        self.best_progress.setVisible(False)
        self._history_selected_file = ""
        self._history_preview_snapshot = None
        self._session_fav_backup = None
        self._session_inc_backup = None
        self._session_lock_backup = None
        self._session_seq_backup = None
        self._update_history_highlights()

    def _cancel_best_schedule(self) -> None:
        if self._best_worker is not None:
            self._best_worker.cancel()
        self._best_token += 1
        self._best_running = False
        self._best_files = []
        self._history_selected_file = ""
        self._close_history_panel()

    def on_start_best_schedule(self) -> None:
        if self.courses_df is None:
            QMessageBox.warning(self, "尚未載入", "請先載入課程 Excel。")
            return
        if not self.username or not self.session_file_path:
            QMessageBox.warning(self, "尚未登入", "請先使用「新增/切換」建立本次登入檔案。")
            return
        if not self.favorites_ids:
            QMessageBox.information(self, "無資料", "目前沒有任何我的最愛課程。")
            return
        if self._best_running:
            return

        best_dir = best_schedule_dir_path(self.user_dir_path)

        fav_sorted = sorted(int(x) for x in self.favorites_ids)
        lock_sorted = sorted(int(x) for x in self.locked_ids)
        cache = load_best_schedule_cache(self.user_dir_path)
        if cache:
            cached_fav = [int(x) for x in cache.get("favorites", [])]
            cached_lock = [int(x) for x in cache.get("locked", [])]
            if cached_fav == fav_sorted and cached_lock == lock_sorted:
                filenames = [str(x) for x in cache.get("files", []) if str(x)]
                files = [os.path.join(best_dir or self.user_dir_path, fn) for fn in filenames]
                if files and all(os.path.exists(p) for p in files):
                    QMessageBox.information(self, "沿用結果", "課程相同，沿用上次最佳選課結果。")
                    self._history_selected_file = ""
                    self._best_files = files
                    self._enter_best_schedule_results()
                    return

        self._session_fav_backup = None
        self._session_inc_backup = None
        self._session_lock_backup = None
        self._session_seq_backup = None
        self._history_selected_file = ""

        self._enter_best_schedule_running()
        self._best_token += 1
        worker = BestScheduleWorker(
            self._best_token,
            self.user_dir_path,
            self.username,
            self.favorites_ids,
            self.locked_ids,
            self.included_ids,
            self.fav_seq,
            self.courses_df,
        )
        self._best_worker = worker
        worker.progress.connect(self._on_best_schedule_progress)
        worker.finished.connect(self._on_best_schedule_finished)
        self.threadpool.start(worker)

    def _on_best_schedule_finished(self, token: int, ok: bool, cancelled: bool, files: list, msg: str) -> None:
        if token != self._best_token:
            return
        self._best_worker = None
        self._best_running = False

        if cancelled:
            self._close_history_panel()
            return

        if not ok:
            QMessageBox.warning(self, "最佳選課失敗", f"最佳選課失敗：\n{msg or '未知錯誤'}")
            self._close_history_panel()
            return

        if not files:
            QMessageBox.information(self, "最佳選課完成", "找不到可用的最佳選課結果。")
            self._close_history_panel()
            return

        self._best_files = list(files)
        self._enter_best_schedule_results()

    def _on_best_schedule_progress(self, token: int, value: int) -> None:
        if token != self._best_token:
            return
        if not self._best_running:
            return
        try:
            v = int(value)
        except Exception:
            v = 0
        self.best_progress.setValue(max(0, min(100, v)))

    def on_toggle_history(self) -> None:
        if self._best_running or self._history_mode == "best":
            return
        show = not self.gb_history.isVisible()
        if show:
            self._session_fav_backup = None
            self._session_inc_backup = None
            self._session_lock_backup = None
            self._session_seq_backup = None
            self._history_selected_file = ""
            self._enter_history_panel()
        else:
            self._close_history_panel()

    def on_history_selected(self) -> None:
        if self._history_mode is None:
            return
        if self._history_mode == "best" and self._best_running:
            return
        it0 = self.tbl_history.item(self.tbl_history.currentRow(), 0)
        if it0 is None:
            return
        path = it0.data(Qt.UserRole)
        if not path:
            return
        path = str(path)

        if not os.path.exists(path):
            QMessageBox.warning(self, "檔案不存在", f"檔案不存在：\n{path}")
            return

        if (
            self._session_fav_backup is None
            or self._session_inc_backup is None
            or self._session_lock_backup is None
            or self._session_seq_backup is None
        ):
            self._session_fav_backup = set(self.favorites_ids)
            self._session_inc_backup = set(self.included_ids)
            self._session_lock_backup = set(self.locked_ids)
            self._session_seq_backup = dict(self.fav_seq)

        try:
            fav, inc, seq, lock = load_user_file(path)
        except Exception as e:
            title = "讀取歷史失敗" if self._history_mode == "history" else "讀取規劃失敗"
            QMessageBox.critical(self, title, f"讀取失敗：\n{e}")
            return

        self._history_selected_file = os.path.abspath(path)
        self._replace_sets(fav, inc, lock, seq)

        self._set_readonly(True)
        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
        self.schedule_search(0)
        self._update_history_highlights()

    def on_back_to_session(self) -> None:
        self.tbl_history.blockSignals(True)
        self.tbl_history.clearSelection()
        self.tbl_history.blockSignals(False)

        restored = self._restore_session_state()
        self._close_history_panel()
        if not restored:
            self._history_selected_file = ""
            self._set_readonly(False)

    def on_use_selected_history(self) -> None:
        if self._history_mode == "best":
            if self._best_running:
                self._cancel_best_schedule()
                return
            if not self._history_selected_file:
                QMessageBox.information(self, "未選擇", "請先在最佳選課中選擇一個檔案。")
                return
        else:
            if not self._history_selected_file:
                QMessageBox.information(self, "未選擇", "請先在歷史紀錄中選擇一個檔案。")
                return
        if not self.session_file_path:
            QMessageBox.warning(self, "尚未登入", "請先使用「新增/切換」建立本次登入檔案。")
            return
        if self.courses_df is None:
            return

        try:
            save_user_file(
                self.session_file_path,
                self.username,
                self.favorites_ids,
                self._get_included_sorted(),
                self._get_locked_sorted(),
                self.fav_seq,
                self.courses_df,
            )
        except Exception as e:
            QMessageBox.critical(self, "覆蓋失敗", f"覆蓋失敗：\n{e}")
            return

        self._session_fav_backup = set(self.favorites_ids)
        self._session_inc_backup = set(self.included_ids)
        self._session_lock_backup = set(self.locked_ids)
        self._session_seq_backup = dict(self.fav_seq)

        self._history_selected_file = ""
        self._set_readonly(False)
        self._refresh_favorites_table()
        self._refresh_timetable()
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
        current_mode = self._history_mode
        self._close_history_panel()
        if current_mode == "history":
            QMessageBox.information(self, "完成", "已使用此歷史紀錄覆蓋本次登入檔案。")
        else:
            QMessageBox.information(self, "完成", "已使用此規劃覆蓋本次登入檔案。")

    def _compute_total_credits(self, ids_sorted: np.ndarray) -> float:
        if self.courses_df is None:
            return 0.0
        if ids_sorted is None or ids_sorted.size == 0:
            return 0.0

        if self._cid_sorted is not None and self._credit_sorted is not None:
            ids = ids_sorted.astype(np.int64, copy=False)
            pos = np.searchsorted(self._cid_sorted, ids)
            if pos.size == 0:
                return 0.0
            valid = (pos >= 0) & (pos < self._cid_sorted.size)
            if not np.any(valid):
                return 0.0
            pos = pos[valid]
            ids = ids[valid]
            match = self._cid_sorted[pos] == ids
            if not np.any(match):
                return 0.0
            credits = self._credit_sorted[pos[match]]
            total = float(np.nansum(credits)) if credits.size else 0.0
            return total

        sub = self.courses_df[self.courses_df["_cid"].isin(ids_sorted)][["學分"]]
        if sub.empty:
            return 0.0
        s = pd.to_numeric(sub["學分"], errors="coerce").fillna(0).sum()
        try:
            return float(s)
        except Exception:
            return 0.0

    def _collect_slots_for_ids(self, ids: Optional[Set[int]]) -> Set[Tuple[str, str]]:
        slots: Set[Tuple[str, str]] = set()
        if self.courses_df is None or not ids:
            return slots
        if self._slots_by_cid:
            for cid in ids:
                try:
                    cid_i = int(cid)
                except Exception:
                    continue
                slots_set = self._slots_by_cid.get(cid_i)
                if not slots_set:
                    continue
                for slot in slots_set:
                    if not isinstance(slot, str) or "-" not in slot:
                        continue
                    day, per = slot.split("-", 1)
                    if day and per:
                        slots.add((day, per))
            return slots

        try:
            ids_list = [int(x) for x in ids]
        except Exception:
            ids_list = []
        if not ids_list:
            return slots
        subset = self.courses_df[self.courses_df["_cid"].isin(ids_list)]
        if subset.empty or "_slots_set" not in subset:
            return slots
        for entry in subset["_slots_set"]:
            slots_set = entry if isinstance(entry, set) else set()
            for slot in slots_set:
                if not isinstance(slot, str) or "-" not in slot:
                    continue
                day, per = slot.split("-", 1)
                if day and per:
                    slots.add((day, per))
        return slots

    def _apply_timetable_background(
        self,
        widget: TimetableWidget,
        col_day_idx: List[int],
        id_matrix: List[List[Optional[int]]],
        locked_matrix: List[List[bool]],
        *,
        diff_added_cids: Optional[Set[int]] = None,
        diff_removed_cells: Optional[Set[Tuple[int, int]]] = None,
    ) -> None:
        cols = widget.columnCount()
        rows = widget.rowCount()
        if cols <= 0 or rows <= 0:
            return
        if len(col_day_idx) != cols:
            return

        for c in range(cols):
            day_idx = col_day_idx[c]
            base = self._day_bg_base[day_idx % len(self._day_bg_base)]
            header_item = widget.horizontalHeaderItem(c)
            if header_item is not None:
                header_item.setBackground(QBrush(base))

        black_bg = QBrush(QColor("#000000"))
        white_fg = QBrush(QColor("#FFFFFF"))
        added_brush = QBrush(QColor("#1B5E20"))
        added_fg = QBrush(QColor("#FFFFFF"))
        deleted_brush = QBrush(QColor("#B71C1C"))

        for c in range(cols):
            day_idx = col_day_idx[c]
            base = self._day_bg_base[day_idx % len(self._day_bg_base)]
            darker = darken(base, self._block_dark_factor)

            course_to_shade: Dict[int, int] = {}
            next_shade = 0
            prev_cid: Optional[int] = None

            for r in range(rows):
                cid = id_matrix[r][c] if (r < len(id_matrix) and c < len(id_matrix[r])) else None
                it = widget.item(r, c)
                if it is None:
                    it = QTableWidgetItem("")
                    widget.setItem(r, c, it)

                locked = False
                if locked_matrix and r < len(locked_matrix) and c < len(locked_matrix[r]):
                    locked = bool(locked_matrix[r][c])

                if locked and (it.text() or "").strip():
                    it.setBackground(black_bg)
                    it.setForeground(white_fg)
                    prev_cid = cid
                    continue

                if diff_added_cids and cid is not None and cid in diff_added_cids:
                    it.setBackground(added_brush)
                    it.setForeground(added_fg)
                    prev_cid = cid
                    continue

                it.setForeground(QBrush(QColor("#000000")))

                if cid is None:
                    it.setBackground(QBrush(base))
                    prev_cid = None
                    continue

                if cid != prev_cid:
                    if cid not in course_to_shade:
                        course_to_shade[cid] = next_shade % 2
                        next_shade += 1
                shade = course_to_shade.get(cid, 0)
                it.setBackground(QBrush(base if shade == 0 else darker))
                prev_cid = cid

        if diff_removed_cells:
            for row, col in diff_removed_cells:
                if 0 <= row < rows and 0 <= col < cols:
                    cell = widget.item(row, col)
                    if cell is not None:
                        cell.setBackground(deleted_brush)

    def _render_timetable(
        self,
        widget: TimetableWidget,
        included_sorted: np.ndarray,
        locked_ids: Set[int],
        *,
        store_state: bool,
        baseline_slots: Optional[Set[Tuple[str, str]]] = None,
        diff_added_cids: Optional[Set[int]] = None,
    ) -> List[str]:
        matrix, conflicts, day_lanes, col_day_idx, id_matrix, locked_matrix = build_timetable_matrix_per_day_lanes_sorted(
            self.courses_df,
            included_sorted,
            locked_ids,
            self.show_days,
        )
        cols = len(col_day_idx)
        if store_state:
            self._tt_col_day_idx = list(col_day_idx)
            self._rebuild_tt_first_lane_cols()
            self._tt_locked_matrix = locked_matrix

        widget.setRowCount(len(PERIODS))
        widget.setColumnCount(cols)

        headers: List[str] = []
        for d in self.show_days:
            lanes = int(day_lanes.get(d, 1))
            if lanes <= 1:
                headers.append(DAY_LABEL[d])
            else:
                for k in range(1, lanes + 1):
                    headers.append(f"{DAY_LABEL[d]} {k}")

        widget.setHorizontalHeaderLabels(headers)

        if self.show_time:
            vlabels = [f"{p}\n{PERIOD_TIME.get(p, '')}".rstrip() for p in PERIODS]
        else:
            vlabels = PERIODS
        widget.setVerticalHeaderLabels(vlabels)

        slot_map: Set[Tuple[str, str]] = set()
        day_map = {idx: self.show_days[idx] for idx in range(len(self.show_days))}
        for c in range(cols):
            day_idx = col_day_idx[c] if c < len(col_day_idx) else None
            day_name = day_map.get(day_idx)
            if day_name is None:
                continue
            for r in range(len(PERIODS)):
                cid = id_matrix[r][c] if (r < len(id_matrix) and c < len(id_matrix[r])) else None
                if cid is not None and 0 <= r < len(PERIODS):
                    slot_map.add((day_name, PERIODS[r]))

        for r in range(len(PERIODS)):
            for c in range(cols):
                it = QTableWidgetItem(matrix[r][c])
                it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                it.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
                widget.setItem(r, c, it)

        deleted_cells: Optional[Set[Tuple[int, int]]] = None
        if baseline_slots:
            removed_slots = baseline_slots - slot_map
            if removed_slots:
                day_index = {day: idx for idx, day in enumerate(self.show_days)}
                temp_cells: Set[Tuple[int, int]] = set()
                for day, per in removed_slots:
                    day_idx = day_index.get(day)
                    if day_idx is None:
                        continue
                    try:
                        row_idx = PERIODS.index(per)
                    except ValueError:
                        continue
                    for c in range(cols):
                        if c >= len(col_day_idx) or col_day_idx[c] != day_idx:
                            continue
                        cell = widget.item(row_idx, c)
                        if cell is None:
                            continue
                        if (cell.text() or "").strip():
                            continue
                        temp_cells.add((row_idx, c))
                if temp_cells:
                    deleted_cells = temp_cells

        self._apply_timetable_background(
            widget,
            col_day_idx,
            id_matrix,
            locked_matrix,
            diff_added_cids=diff_added_cids,
            diff_removed_cells=deleted_cells,
        )

        if cols <= 12:
            widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        else:
            widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        return conflicts

    def _refresh_history_preview_timetable(self) -> None:
        if not getattr(self, "_history_layout_active", False):
            if getattr(self, "gb_tt_preview", None) is not None:
                self.gb_tt_preview.setVisible(False)
            return

        snapshot = self._history_preview_snapshot
        if snapshot is None or self.courses_df is None:
            if getattr(self, "gb_tt_preview", None) is not None:
                self.gb_tt_preview.setVisible(False)
            return

        included = sorted_array_from_set_int(snapshot.get("included", set()))
        locked = snapshot.get("locked", set())
        if getattr(self, "gb_tt_preview", None) is not None:
            self.gb_tt_preview.setVisible(True)
        self._render_timetable(
            self.tbl_tt_preview,
            included,
            locked,
            store_state=False,
        )

        total_credits = self._compute_total_credits(included)
        total_txt = (
            str(int(round(total_credits)))
            if abs(total_credits - round(total_credits)) < 1e-9
            else f"{total_credits:.1f}".rstrip("0").rstrip(".")
        )
        if self.lbl_total_credits_preview is not None:
            self.lbl_total_credits_preview.setText(f"已選總學分：{total_txt}")

        locked_arr = sorted_array_from_set_int(set(locked))
        locked_credits = self._compute_total_credits(locked_arr)
        locked_txt = (
            str(int(round(locked_credits)))
            if abs(locked_credits - round(locked_credits)) < 1e-9
            else f"{locked_credits:.1f}".rstrip("0").rstrip(".")
        )
        if self.lbl_locked_credits_preview is not None:
            self.lbl_locked_credits_preview.setText(f"已鎖定學分：{locked_txt}")

        self._sync_timetable_font_size(self.tbl_tt)
        self._apply_timetable_row_heights()

    def _refresh_timetable(self) -> None:
        if self.courses_df is None:
            return

        self.included_ids |= self.locked_ids
        self._mark_included_dirty()

        inc_sorted = self._get_included_sorted()

        total_credits = self._compute_total_credits(inc_sorted)
        total_txt = str(int(round(total_credits))) if abs(total_credits - round(total_credits)) < 1e-9 else f"{total_credits:.1f}".rstrip("0").rstrip(".")
        self.lbl_total_credits.setText(f"已選總學分：{total_txt}")

        locked_included = sorted_array_from_set_int(set(self.locked_ids) & set(self.included_ids))
        locked_credits = self._compute_total_credits(locked_included)
        lock_txt = str(int(round(locked_credits))) if abs(locked_credits - round(locked_credits)) < 1e-9 else f"{locked_credits:.1f}".rstrip("0").rstrip(".")
        self.lbl_locked_credits.setText(f"已鎖定學分：{lock_txt}")

        baseline_slots: Optional[Set[Tuple[str, str]]] = None
        diff_added: Optional[Set[int]] = None
        snapshot = self._history_preview_snapshot
        if snapshot:
            preview_included: Set[int] = set(snapshot.get("included", set()))
            if preview_included:
                baseline_slots = self._collect_slots_for_ids(preview_included)
                diff_added = set(self.included_ids) - preview_included

        self._history_preview_slot_map = baseline_slots

        conflicts = self._render_timetable(
            self.tbl_tt,
            inc_sorted,
            self.locked_ids,
            store_state=True,
            baseline_slots=baseline_slots,
            diff_added_cids=diff_added,
        )

        self._refresh_history_preview_timetable()
        self._apply_timetable_row_heights()

        if conflicts:
            self.lbl_conflicts.setText("衝堂提示：\n" + "\n".join(conflicts))
            self.lbl_conflicts.setStyleSheet("color: #b00020;")
        else:
            self.lbl_conflicts.setText("")
            self.lbl_conflicts.setStyleSheet("")

        self.tbl_tt.viewport().update()

    def on_search(self) -> None:
        if self.courses_df is None:
            return

        df = self.courses_df
        n = len(df)
        mask = np.ones(n, dtype=bool)

        full = (self.ed_full.text() or "").strip()
        if full:
            tokens = [t.strip().lower() for t in full.split() if t.strip()]
            if tokens:
                s = df["_alltext"]
                for tok in tokens:
                    mask &= s.str.contains(tok, regex=False, na=False).to_numpy()

        special_gened = self.ck_gened.isChecked()
        special_sport = self.ck_sport.isChecked()
        special_teaching = self.ck_teaching.isChecked()

        serial = self.ed_serial.text().strip()
        if serial:
            ids: List[int] = []
            raw = serial.replace(",", " ").replace("，", " ")
            for tok in raw.split():
                try:
                    ids.append(int(tok))
                except Exception:
                    continue
            if ids:
                cid_arr = self._cid_arr if self._cid_arr is not None else df["_cid"].to_numpy(dtype=np.int64, copy=False)
                mask &= np.isin(cid_arr, np.array(ids, dtype=np.int64))

        code_q = (self.ed_course_code.text() or "").strip()
        if code_q and "開課代碼" in df.columns:
            tokens = [t.strip().lower() for t in code_q.split() if t.strip()]
            if tokens:
                s = df["_code_lc"] if "_code_lc" in df.columns else df["開課代碼"].astype(str).str.lower()
                for tok in tokens:
                    mask &= s.str.contains(tok, regex=False, na=False).to_numpy()

        cname = self.ed_cname.text().strip()
        if cname and "中文課程名稱" in df.columns:
            cname_lc = cname.lower()
            if "_cname_lc" in df.columns:
                mask &= df["_cname_lc"].str.contains(cname_lc, regex=False, na=False).to_numpy()
            else:
                mask &= df["中文課程名稱"].astype(str).str.contains(cname, na=False).to_numpy()

        teacher = self.ed_teacher.text().strip()
        if teacher and "教師" in df.columns:
            teacher_lc = teacher.lower()
            if "_teacher_lc" in df.columns:
                mask &= df["_teacher_lc"].str.contains(teacher_lc, regex=False, na=False).to_numpy()
            else:
                mask &= df["教師"].astype(str).str.contains(teacher, na=False).to_numpy()

        apply_dept_filter = not (special_gened or special_sport)
        if apply_dept_filter:
            dept_text = self.cb_dept.currentText().strip()
            if dept_text and dept_text != "(全部)" and "系所" in df.columns:
                if dept_text in self._all_depts:
                    if self._dept_arr is not None:
                        mask &= (self._dept_arr == dept_text)
                    else:
                        mask &= (df["系所"] == dept_text).to_numpy()
                else:
                    dept_lc = dept_text.lower()
                    if "_dept_lc" in df.columns:
                        mask &= df["_dept_lc"].str.contains(dept_lc, regex=False, na=False).to_numpy()
                    else:
                        mask &= df["系所"].astype(str).str.contains(dept_text, na=False).to_numpy()

        elif special_gened and "系所" in df.columns:
            if self._dept_arr is not None:
                mask &= (self._dept_arr == GENED_DEPT_NAME)
            else:
                mask &= (df["系所"] == GENED_DEPT_NAME).to_numpy()
            core_choice = self.cb_gened_core.currentText().strip()
            if core_choice and core_choice != "所有通識" and "_gened_cats" in df.columns:
                core_mask = np.array([core_choice in cats for cats in df["_gened_cats"]], dtype=bool)
                mask &= core_mask

        elif special_sport and "系所" in df.columns:
            if self._dept_arr is not None:
                mask &= (self._dept_arr == SPORT_DEPT_NAME)
            else:
                mask &= (df["系所"] == SPORT_DEPT_NAME).to_numpy()

        elif special_teaching:
            if "中文課程名稱" in df.columns:
                mask &= df["中文課程名稱"].astype(str).str.contains(TEACHING_NAME_TOKEN, na=False).to_numpy()

        if self.ck_not_full.isChecked():
            if "限修人數" in df.columns and "選修人數" in df.columns:
                nf = (df["限修人數"].notna()) & (df["選修人數"].notna()) & (df["選修人數"] < df["限修人數"])
                mask &= nf.to_numpy()

        if self.ck_exclude_selected.isChecked():
            if self.included_ids:
                cid_arr = self._cid_arr if self._cid_arr is not None else df["_cid"].to_numpy(dtype=np.int64, copy=False)
                mask &= ~np.isin(cid_arr, np.array(list(self.included_ids), dtype=np.int64))

        if not self.ck_show_tba.isChecked():
            tba = self._tba_arr if self._tba_arr is not None else df["_tba"].to_numpy(dtype=bool, copy=False)
            mask &= ~tba

        if (self._sel_lo != 0) or (self._sel_hi != 0):
            sel_lo = np.uint64(self._sel_lo)
            sel_hi = np.uint64(self._sel_hi)

            mlo = self._mask_lo_arr if self._mask_lo_arr is not None else df["_mask_lo"].to_numpy(dtype="uint64", copy=False)
            mhi = self._mask_hi_arr if self._mask_hi_arr is not None else df["_mask_hi"].to_numpy(dtype="uint64", copy=False)

            mode = self.cb_match_mode.currentIndex()
            if mode == 0:
                cond = (np.bitwise_and(mlo, np.bitwise_not(sel_lo)) == 0) & (np.bitwise_and(mhi, np.bitwise_not(sel_hi)) == 0)
            else:
                cond = (np.bitwise_and(mlo, sel_lo) != 0) | (np.bitwise_and(mhi, sel_hi) != 0)

            mask &= cond

        if self.ck_exclude_conflict.isChecked():
            inc_sorted = self._get_included_sorted()
            if inc_sorted.size:
                occ_lo, occ_hi = occupied_masks_sorted(self.courses_df, inc_sorted)
                occ_lo = np.uint64(occ_lo)
                occ_hi = np.uint64(occ_hi)

                mlo = self._mask_lo_arr if self._mask_lo_arr is not None else df["_mask_lo"].to_numpy(dtype="uint64", copy=False)
                mhi = self._mask_hi_arr if self._mask_hi_arr is not None else df["_mask_hi"].to_numpy(dtype="uint64", copy=False)
                tba = self._tba_arr if self._tba_arr is not None else df["_tba"].to_numpy(dtype=bool, copy=False)

                cond_no_conflict = (np.bitwise_and(mlo, occ_lo) == 0) & (np.bitwise_and(mhi, occ_hi) == 0)
                cond = cond_no_conflict | tba
                mask &= cond

        cols = self.display_columns if self.display_columns else [c for c in df.columns if not str(c).startswith("_")]
        self.filtered_df = df.loc[mask, cols]
        self.model_results.set_df(self.filtered_df)
        self.model_results.notify_favorites_changed()
        self.proxy_results.invalidate()
