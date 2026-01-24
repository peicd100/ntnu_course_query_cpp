from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Protocol, Set, Tuple

import numpy as np
import pandas as pd
from PySide6.QtCore import (
    Qt,
    QAbstractTableModel,
    QModelIndex,
    QEvent,
    QSortFilterProxyModel,
    QRect,
    Signal,
)
from PySide6.QtGui import QColor, QPainter, QPen, QFontMetrics, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QStyledItemDelegate,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from app_utils import sorted_array_from_set_int


class IntSortItem(QTableWidgetItem):
    def __init__(self, text: str = "", value: int = 0):
        super().__init__(text)
        self.setData(Qt.UserRole, int(value))

    def __lt__(self, other: "QTableWidgetItem") -> bool:
        try:
            a = int(self.data(Qt.UserRole) or 0)
        except Exception:
            a = 0
        try:
            b = int(other.data(Qt.UserRole) or 0)
        except Exception:
            b = 0
        return a < b


class FloatSortItem(QTableWidgetItem):
    def __init__(self, text: str = "", value: float = 0.0):
        super().__init__(text)
        try:
            self.setData(Qt.UserRole, float(value))
        except Exception:
            self.setData(Qt.UserRole, 0.0)

    def __lt__(self, other: "QTableWidgetItem") -> bool:
        try:
            a = float(self.data(Qt.UserRole) or 0.0)
        except Exception:
            a = 0.0
        try:
            b = float(other.data(Qt.UserRole) or 0.0)
        except Exception:
            b = 0.0
        return a < b


class FavoritesTableWidget(QTableWidget):
    orderChanged = Signal(list)
    # Emitted after a drop completes. Argument: changed (bool) indicating whether
    # the underlying order actually changed.
    dropCompleted = Signal(bool)

    # Emitted when a drag-selection on checkboxes is finished.
    # Arguments: column index, set of affected row indices.
    dragSelectionFinished = Signal(int, set)

    def __init__(self, move_column: int = 0, cid_column: int = 2):
        super().__init__()
        self._move_column = move_column
        self._cid_column = cid_column
        self._drag_enabled = False
        self._drag_source_row: Optional[int] = None
        self._drag_snapshot: List[int] = []
        self._drag_started = False
        self._sorting_was_enabled = False
        self._drop_indicator_rect: Optional[QRect] = None
        self._drop_indicator_color = QColor("#FF8C00")  # DarkOrange, more visible

        self._is_drag_selecting = False
        self._drag_select_column = -1
        self._drag_select_start_row = -1
        self._drag_select_initial_state: Optional[Qt.CheckState] = None
        self._drag_select_affected_rows: Set[int] = set()
        self._last_drag_select_row = -1

        self.setMouseTracking(True)
        self.setDragDropOverwriteMode(False)
        self.setDropIndicatorShown(False)
        self.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.setDragEnabled(False)
        self.setAcceptDrops(False)
        self.setDefaultDropAction(Qt.MoveAction)

    def cid_at_row(self, row: int) -> Optional[int]:
        item = self.item(row, self._cid_column)
        if not item:
            return None
        data = item.data(Qt.UserRole)
        if data is None:
            return None
        try:
            return int(data)
        except Exception:
            return None

    def set_drag_enabled(self, enabled: bool) -> None:
        self._drag_enabled = bool(enabled)
        if not self._drag_enabled:
            self._drag_source_row = None
            self._drag_snapshot = []
            self._drag_started = False
        mode = QAbstractItemView.DragDrop if self._drag_enabled else QAbstractItemView.NoDragDrop
        self.setDragDropMode(mode)
        self.setDragEnabled(self._drag_enabled)
        self.setAcceptDrops(self._drag_enabled)
        self.setDropIndicatorShown(self._drag_enabled)
        self.setDefaultDropAction(Qt.MoveAction)
        if not self._drag_enabled:
            self.viewport().unsetCursor()

    def is_drag_enabled(self) -> bool:
        return self._drag_enabled

    def _current_order(self) -> List[int]:
        order: List[int] = []
        for row in range(self.rowCount()):
            cid = self.cid_at_row(row)
            if cid is not None:
                order.append(cid)
        return order

    def _event_pos(self, event):
        if hasattr(event, "position"):
            pos = event.position()
            return pos.toPoint() if hasattr(pos, "toPoint") else pos
        return event.pos()

    def _drop_insert_index(self, event) -> int:
        pos = self._event_pos(event)
        idx = self.indexAt(pos)
        if not idx.isValid():
            return len(self._drag_snapshot)
        row = idx.row()
        indicator = self.dropIndicatorPosition()
        if indicator == QAbstractItemView.OnViewport:
            return len(self._drag_snapshot)
        if indicator == QAbstractItemView.BelowItem:
            return row + 1
        return row

    def _reordered_sequence(self, dest_index: int) -> List[int]:
        order = list(self._drag_snapshot)
        if not order:
            return order
        if self._drag_source_row is None:
            return order
        if not (0 <= self._drag_source_row < len(order)):
            return order
        
        # Clamp dest_index to valid range
        dest_index = max(0, min(dest_index, len(order)))
        
        # If source and dest are the same, it's no move
        if dest_index == self._drag_source_row:
            return order
        
        # Remove from source position
        cid = order.pop(self._drag_source_row)
        
        # Adjust dest_index if needed (because we removed an element before it)
        if dest_index > self._drag_source_row:
            dest_index -= 1
        
        # Ensure dest_index is in valid range after adjustment
        dest_index = max(0, min(dest_index, len(order)))
        
        # Insert at destination
        order.insert(dest_index, cid)
        
        # Verify no courses were lost
        if len(order) != len(self._drag_snapshot):
            # Safety check: if count doesn't match, return original
            return list(self._drag_snapshot)
        
        # Verify all course IDs are preserved
        original_set = set(self._drag_snapshot)
        new_set = set(order)
        if original_set != new_set:
            # Safety check: if any course is missing, return original
            return list(self._drag_snapshot)
        
        return order


    def _clear_drag_state(self) -> None:
        self._drag_source_row = None
        self._drag_snapshot = []
        self._drag_started = False
        self.viewport().unsetCursor()

        # Also clear drag-selection state
        self._is_drag_selecting = False
        self._drag_select_column = -1
        self._drag_select_start_row = -1
        self._drag_select_initial_state = None
        self._drag_select_affected_rows.clear()
        self._last_drag_select_row = -1

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            pos = self._event_pos(event)
            row = self.rowAt(pos.y())
            col = self.columnAt(pos.x())

            if row >= 0 and col in (1, 2):  # Schedule or Lock column
                self._is_drag_selecting = True
                self._drag_select_column = col
                self._drag_select_start_row = row
                self._last_drag_select_row = row
                item = self.item(row, col)
                if item and item.flags() & Qt.ItemIsUserCheckable:
                    # Toggle the state for a single click
                    current_state = item.checkState()
                    new_state = Qt.Unchecked if current_state == Qt.Checked else Qt.Checked
                    self._drag_select_initial_state = new_state
                    item.setCheckState(new_state)
                    self._drag_select_affected_rows.add(row)
                else:
                    self._drag_select_initial_state = None
                
                # Prevent row drag-and-drop from starting
                self._drag_source_row = None
                event.accept()
                return

            if self._drag_enabled and row >= 0 and col == self._move_column:
                self._drag_source_row = row
                self._drag_snapshot = self._current_order()
                self._drag_started = False
                # If sorting is enabled, temporarily disable it during drag
                try:
                    self._sorting_was_enabled = self.isSortingEnabled()
                    if self._sorting_was_enabled:
                        self.setSortingEnabled(False)
                except Exception:
                    self._sorting_was_enabled = False
                self.viewport().setCursor(Qt.ClosedHandCursor)
                super().mousePressEvent(event)
                return

        super().mousePressEvent(event)

    def startDrag(self, supportedActions) -> None:
        if self._drag_source_row is None:
            return
        self._drag_started = True
        super().startDrag(supportedActions)

    def mouseMoveEvent(self, event) -> None:
        if self._is_drag_selecting:
            if not (event.buttons() & Qt.LeftButton):
                return
            
            pos = self._event_pos(event)
            current_row = self.rowAt(pos.y())
            
            if current_row >= 0 and current_row != self._last_drag_select_row:
                start_row = self._drag_select_start_row
                end_row = current_row
                
                # Determine the range of rows to apply the state to
                rows_to_process = range(min(start_row, end_row), max(start_row, end_row) + 1)
                
                for r in rows_to_process:
                    # Check if this row was already processed in this drag
                    if r in self._drag_select_affected_rows and self.item(r, self._drag_select_column).checkState() == self._drag_select_initial_state:
                        continue

                    item = self.item(r, self._drag_select_column)
                    if item and item.flags() & Qt.ItemIsUserCheckable and self._drag_select_initial_state is not None:
                        # Prevent changing state on locked items when dragging schedule column
                        if self._drag_select_column == 1: # Schedule column
                            lock_item = self.item(r, 2) # Lock column
                            if lock_item and lock_item.checkState() == Qt.Checked:
                                continue
                        
                        item.setCheckState(self._drag_select_initial_state)
                        self._drag_select_affected_rows.add(r)

                self._last_drag_select_row = current_row
            
            event.accept()
            return

        if self._drag_enabled:
            if self._drag_source_row is None:
                pos = self._event_pos(event)
                row = self.rowAt(pos.y())
                col = self.columnAt(pos.x())
                if row >= 0 and col == self._move_column:
                    self.viewport().setCursor(Qt.OpenHandCursor)
                else:
                    self.viewport().unsetCursor()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._is_drag_selecting:
            if self._drag_select_column != -1 and self._drag_select_affected_rows:
                # On single click without drag, cellChanged handles it.
                # On drag, we need to signal the batch update.
                is_drag = (len(self._drag_select_affected_rows) > 1) or (self._drag_select_start_row != self._last_drag_select_row and self._last_drag_select_row != -1)
                if is_drag:
                    self.dragSelectionFinished.emit(self._drag_select_column, self._drag_select_affected_rows)
            self._clear_drag_state()
            event.accept()
            return

        if self._drag_source_row is not None and not self._drag_started:
            self._clear_drag_state()
            # restore sorting if it was disabled earlier
            try:
                if getattr(self, "_sorting_was_enabled", False):
                    self.setSortingEnabled(True)
            except Exception:
                pass
        super().mouseReleaseEvent(event)

    def dragEnterEvent(self, event) -> None:
        if self._drag_enabled and self._drag_source_row is not None:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._drop_indicator_rect = None
        self.viewport().update()
        event.accept()

    def dragMoveEvent(self, event) -> None:
        if self._drag_enabled and self._drag_source_row is not None:
            dest_index = self._drop_insert_index(event)
            if dest_index < 0 or dest_index > self.rowCount():
                self._drop_indicator_rect = None
            else:
                if dest_index == self.rowCount():
                    # Dropping at the very end
                    if self.rowCount() > 0:
                        last_row_rect = self.visualRect(self.model().index(self.rowCount() - 1, 0))
                        self._drop_indicator_rect = QRect(
                            last_row_rect.left(),
                            last_row_rect.bottom() - 1,
                            self.viewport().width(),
                            3,
                        )
                    else:
                        self._drop_indicator_rect = QRect(0, 0, self.viewport().width(), 3)
                else:
                    # Dropping between rows
                    row_rect = self.visualRect(self.model().index(dest_index, 0))
                    self._drop_indicator_rect = QRect(
                        row_rect.left(),
                        row_rect.top() - 1,
                        self.viewport().width(),
                        3,
                    )
            self.viewport().update()
            event.acceptProposedAction()
        else:
            self._drop_indicator_rect = None
            event.ignore()

    def dropEvent(self, event) -> None:
        if not self._drag_enabled or self._drag_source_row is None:
            event.ignore()
            self._clear_drag_state()
            return

        dest_index = self._drop_insert_index(event)
        new_order = self._reordered_sequence(dest_index)
        changed = False
        if new_order and new_order != self._drag_snapshot:
            self.orderChanged.emit(new_order)
            changed = True

        self._drop_indicator_rect = None
        event.acceptProposedAction()
        # restore sorting if it was disabled during drag
        try:
            if getattr(self, "_sorting_was_enabled", False):
                self.setSortingEnabled(True)
        except Exception:
            pass

        # Notify caller that a drop completed; if there was no underlying
        # order change, the caller can choose to refresh the view to avoid
        # temporary visual artefacts from the drag.
        try:
            self.dropCompleted.emit(changed)
        except Exception:
            pass

        self._clear_drag_state()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self._drop_indicator_rect:
            painter = QPainter(self.viewport())
            painter.fillRect(self._drop_indicator_rect, self._drop_indicator_color)



class ResultsModel(QAbstractTableModel):
    favoriteToggled = Signal(int, bool)

    def __init__(self, df: pd.DataFrame, favorites_ref: Set[int]):
        super().__init__()
        self._df = df
        self._favorites = favorites_ref
        self._readonly = False
        self._fav_sorted = np.empty((0,), dtype=np.int64)
        self._rebuild_fav_sorted()

    def _rebuild_fav_sorted(self) -> None:
        self._fav_sorted = sorted_array_from_set_int(self._favorites)

    def _fav_has(self, cid: int) -> bool:
        arr = self._fav_sorted
        if arr.size == 0:
            return False
        x = np.int64(int(cid))
        pos = int(np.searchsorted(arr, x, side="left"))
        return 0 <= pos < arr.size and int(arr[pos]) == int(x)

    def set_readonly(self, readonly: bool) -> None:
        self._readonly = bool(readonly)
        if self.rowCount() > 0:
            top = self.index(0, 0)
            bot = self.index(self.rowCount() - 1, 0)
            self.dataChanged.emit(top, bot, [Qt.CheckStateRole])

    def set_df(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df.columns) + 1

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if section == 0:
                return "我的最愛"
            return str(self._df.columns[section - 1])
        return str(section + 1)

    def _course_id_at_row(self, row: int) -> Optional[int]:
        if row < 0 or row >= len(self._df):
            return None
        try:
            v = self._df.iloc[row]["開課序號"]
            return int(str(v).strip())
        except Exception:
            return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None

        r = index.row()
        c = index.column()

        if role == Qt.UserRole:
            if c == 0:
                cid = self._course_id_at_row(r)
                return 1 if (cid is not None and self._fav_has(cid)) else 0

            col_name = self._df.columns[c - 1]
            v = self._df.iat[r, c - 1]

            if str(col_name) == "開課序號":
                try:
                    return int(str(v).strip())
                except Exception:
                    return 10**9

            if isinstance(v, (int, float, np.integer, np.floating)) and not (isinstance(v, float) and pd.isna(v)):
                try:
                    return float(v)
                except Exception:
                    return 0.0

            try:
                return "" if pd.isna(v) else str(v)
            except Exception:
                return str(v)

        if c == 0:
            if role == Qt.CheckStateRole:
                cid = self._course_id_at_row(r)
                if cid is None:
                    return Qt.Unchecked
                return Qt.Checked if self._fav_has(cid) else Qt.Unchecked
            if role in (Qt.DisplayRole, Qt.ToolTipRole):
                return ""
            return None

        c2 = c - 1
        if c2 < 0 or c2 >= len(self._df.columns):
            return None

        if role in (Qt.DisplayRole, Qt.ToolTipRole):
            v = self._df.iat[r, c2]
            return "" if pd.isna(v) else str(v)
        return None

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags

        if index.column() == 0:
            if self._readonly:
                return Qt.ItemIsEnabled | Qt.ItemIsSelectable
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable

        return Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if self._readonly:
            return False
        if not index.isValid():
            return False
        if index.column() != 0:
            return False
        if role != Qt.CheckStateRole:
            return False

        cid = self._course_id_at_row(index.row())
        if cid is None:
            return False

        checked = (value == Qt.Checked)
        self.favoriteToggled.emit(cid, checked)
        self.dataChanged.emit(index, index, [Qt.CheckStateRole, Qt.UserRole])
        return True

    def notify_favorites_changed(self) -> None:
        self._rebuild_fav_sorted()
        if self.rowCount() <= 0:
            return
        top_left = self.index(0, 0)
        bottom_right = self.index(self.rowCount() - 1, 0)
        self.dataChanged.emit(top_left, bottom_right, [Qt.CheckStateRole, Qt.UserRole])


class CheckBoxClickDelegate(QStyledItemDelegate):
    def editorEvent(self, event, model, option, index):
        if event.type() in (QEvent.MouseButtonRelease, QEvent.MouseButtonDblClick):
            if index.flags() & Qt.ItemIsUserCheckable:
                current = model.data(index, Qt.CheckStateRole)
                new_state = Qt.Unchecked if current == Qt.Checked else Qt.Checked
                return model.setData(index, new_state, Qt.CheckStateRole)
        return super().editorEvent(event, model, option, index)


class ResultsFrozenView(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        self.frozen_view = QTableView()
        self.main_view = QTableView()

        # 避免選取顏色過深
        self.frozen_view.setStyleSheet(
            "QTableView { border: none; }"
            "QTableView::item:selected { background: rgba(255,255,255,22); }"
        )
        self.main_view.setStyleSheet(
            "QTableView { border: none; }"
            "QTableView::item:selected { background: rgba(255,255,255,22); }"
        )

        lay.addWidget(self.frozen_view, 0)
        lay.addWidget(self.main_view, 1)

        self.main_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.main_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.main_view.setSortingEnabled(True)
        self.main_view.horizontalHeader().setSortIndicatorShown(True)
        self.main_view.horizontalHeader().setStretchLastSection(True)
        self.main_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        self.frozen_view.setFocusPolicy(Qt.NoFocus)
        self.frozen_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.frozen_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.frozen_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.frozen_view.horizontalHeader().setStretchLastSection(False)
        self.frozen_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        self._checkbox_delegate = CheckBoxClickDelegate(self.frozen_view)
        self.frozen_view.setItemDelegateForColumn(0, self._checkbox_delegate)

        self._model: Optional[QSortFilterProxyModel] = None

        self.main_view.verticalScrollBar().valueChanged.connect(self.frozen_view.verticalScrollBar().setValue)
        self.frozen_view.verticalScrollBar().valueChanged.connect(self.main_view.verticalScrollBar().setValue)
        self.main_view.verticalHeader().sectionResized.connect(self._sync_row_height_from_main)

        self.frozen_view.verticalHeader().geometriesChanged.connect(self._update_frozen_width)
        self.frozen_view.horizontalHeader().sectionResized.connect(self._update_frozen_width)
        self.frozen_view.horizontalHeader().geometriesChanged.connect(self._update_frozen_width)

        self._last_header_signature: Optional[Tuple[str, ...]] = None

    def setModel(self, proxy_model: QSortFilterProxyModel) -> None:
        self._model = proxy_model

        self.main_view.setModel(proxy_model)
        self.frozen_view.setModel(proxy_model)

        self.frozen_view.setSelectionModel(self.main_view.selectionModel())

        self._apply_column_visibility()
        self._apply_result_default_column_widths_if_needed()

        self.main_view.verticalHeader().hide()
        self.frozen_view.verticalHeader().setVisible(True)

        self.frozen_view.setSortingEnabled(False)
        self._update_frozen_width()

        proxy_model.modelReset.connect(self._on_proxy_reset_like)
        proxy_model.layoutChanged.connect(self._on_proxy_reset_like)
        proxy_model.columnsInserted.connect(self._on_proxy_reset_like)
        proxy_model.columnsRemoved.connect(self._on_proxy_reset_like)

    def _on_proxy_reset_like(self, *_args) -> None:
        self._apply_column_visibility()
        self._apply_result_default_column_widths_if_needed()

    def _apply_column_visibility(self) -> None:
        if self._model is None:
            return

        for c in range(self._model.columnCount()):
            self.frozen_view.setColumnHidden(c, c != 0)

        self.main_view.setColumnHidden(0, True)
        for c in range(1, self._model.columnCount()):
            self.main_view.setColumnHidden(c, False)

        self.frozen_view.resizeColumnToContents(0)
        self._update_frozen_width()

        ds = self.main_view.verticalHeader().defaultSectionSize()
        self.frozen_view.verticalHeader().setDefaultSectionSize(ds)

    def _sync_row_height_from_main(self, logicalIndex: int, _old: int, newSize: int) -> None:
        self.frozen_view.setRowHeight(logicalIndex, newSize)

    def _update_frozen_width(self, *_args) -> None:
        vh = self.frozen_view.verticalHeader()
        col_w = self.frozen_view.columnWidth(0)
        extra = self.frozen_view.frameWidth() * 2 + 2
        total = int(vh.width() + col_w + extra)
        total = max(total, 60)
        self.frozen_view.setFixedWidth(total)

    def _apply_result_default_column_widths_if_needed(self) -> None:
        if self._model is None:
            return
        col_count = self._model.columnCount()
        if col_count <= 0:
            return

        headers: List[str] = []
        for c in range(col_count):
            h = self._model.headerData(c, Qt.Horizontal, Qt.DisplayRole)
            headers.append(str(h) if h is not None else "")
        sig = tuple(headers)

        if self._last_header_signature == sig:
            return
        self._last_header_signature = sig

        hh = self.main_view.horizontalHeader()
        fm = QFontMetrics(hh.font())
        one_char = max(1, fm.horizontalAdvance("字"))

        title0 = headers[0] if headers else "我的最愛"
        w0 = fm.horizontalAdvance(title0) + one_char + 22
        w0 = max(w0, 60)
        self.frozen_view.setColumnWidth(0, int(w0))

        for c in range(1, col_count):
            title = headers[c]
            w = fm.horizontalAdvance(title) + one_char + 22
            w = max(w, 60)
            self.main_view.setColumnWidth(c, int(w))

        self._update_frozen_width()


class TimetableWidget(QTableWidget):
    zoomChanged = Signal(int)

    def __init__(self):
        super().__init__()
        ps = self.font().pointSize()
        ps = ps if ps and ps > 0 else 10
        self._current_point_size = max(ps, 12)
        self._min_point_size = 6
        self._max_point_size = 26
        self._apply_font_size(self._current_point_size, update_items=False)

    def _apply_font_size(self, size: int, update_items: bool, emit_signal: bool = False) -> None:
        size = max(self._min_point_size, min(self._max_point_size, int(size)))

        f = self.font()
        f.setPointSize(size)
        self.setFont(f)
        self.horizontalHeader().setFont(f)
        self.verticalHeader().setFont(f)

        if update_items:
            for r in range(self.rowCount()):
                for c in range(self.columnCount()):
                    it = self.item(r, c)
                    if it is None:
                        continue
                    fi = it.font()
                    fi.setPointSize(size)
                    it.setFont(fi)

        if emit_signal:
            self.zoomChanged.emit(size)

    def set_point_size(self, size: int, *, update_items: bool = True, emit_zoom: bool = False) -> None:
        size = max(self._min_point_size, min(self._max_point_size, int(size)))
        self._current_point_size = size
        self._apply_font_size(size, update_items=update_items, emit_signal=emit_zoom)

    @property
    def point_size(self) -> int:
        return self._current_point_size

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            dy = event.angleDelta().y()
            if dy == 0:
                event.accept()
                return

            step = 1 if dy > 0 else -1
            new_size = self._current_point_size + step
            new_size = max(self._min_point_size, min(self._max_point_size, new_size))

            if new_size != self._current_point_size:
                self._current_point_size = new_size
                self._apply_font_size(self._current_point_size, update_items=True, emit_signal=True)

            event.accept()
            return

        super().wheelEvent(event)


class MainWindowLike(Protocol):
    def tt_cell_has_selector_box(self, row: int, col: int) -> bool: ...
    def tt_day_idx_from_col(self, col: int) -> Optional[int]: ...
    def tt_is_time_selected(self, day_idx: int, row: int) -> bool: ...
    def tt_cell_locked(self, row: int, col: int) -> bool: ...


class TTTimeSelectDelegate(QStyledItemDelegate):
    def __init__(self, mw: MainWindowLike, parent=None):
        super().__init__(parent)
        self.mw = mw
        self.box_size = 12
        self.box_margin = 2

    def paint(self, painter: QPainter, option, index: QModelIndex):
        super().paint(painter, option, index)

        r = index.row()
        c = index.column()

        if not self.mw.tt_cell_has_selector_box(r, c):
            return

        rect = option.rect
        box = QRect(rect.left() + self.box_margin, rect.top() + self.box_margin, self.box_size, self.box_size)

        day_idx = self.mw.tt_day_idx_from_col(c)
        if day_idx is None:
            return

        selected = self.mw.tt_is_time_selected(day_idx, r)
        locked_cell = self.mw.tt_cell_locked(r, c)

        painter.save()

        if selected:
            painter.setPen(QPen(QColor(220, 0, 0), 1))
            painter.drawRect(box)
            inner = box.adjusted(2, 2, -2, -2)
            painter.fillRect(inner, QColor(220, 0, 0))
        else:
            pen_color = QColor(140, 140, 140) if locked_cell else QColor(0, 0, 0)
            painter.setPen(QPen(pen_color, 1))
            painter.drawRect(box)

        painter.restore()
