from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from PySide6.QtCore import QEvent, QObject, QSettings, QTimer, Qt, Signal
from PySide6.QtGui import QFont

from app.ui.table_headers import (
    calculate_gui_header_base_height,
    install_gui_header_style,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QLineEdit,
    QSpinBox,
    QDoubleSpinBox,
    QTableView,
    QTableWidget,
    QWidget,
)


MIN_TABLE_SCALE = 60
MAX_TABLE_SCALE = 160
TABLE_SCALE_STEP = 10
DEFAULT_TABLE_SCALE = 90
BASE_ROW_HEIGHT = 22
FALLBACK_ITEM_FONT_PT = 10.0
BASE_HEADER_FONT_PT = 10.0
BASE_HEADER_PADDING_PX = 3


@dataclass
class _TableScaleState:
    table: QTableView
    base_column_widths: dict[int, int] = field(default_factory=dict)
    base_header_height: int = 24
    base_item_font_pt: float = FALLBACK_ITEM_FONT_PT
    base_style_sheet: str = ""
    original_resize_columns_to_contents: Callable | None = None
    original_set_column_width: Callable | None = None
    applying: bool = False


class TableScaleManager(QObject):
    """One shared scale for every table in the application."""

    scale_changed = Signal(int)

    def __init__(self, app: QApplication):
        super().__init__(app)
        self._app = app
        self._settings = QSettings("Phoenix Lubricants", "Procurement App")
        self._scale_percent = self._normalise_scale(
            self._settings.value(
                "ui/table_scale_percent",
                DEFAULT_TABLE_SCALE,
                type=int,
            )
        )
        self._states: dict[int, _TableScaleState] = {}
        app.installEventFilter(self)

    @property
    def scale_percent(self) -> int:
        return self._scale_percent

    @property
    def scale_factor(self) -> float:
        return self._scale_percent / 100.0

    def increase(self) -> None:
        self.set_scale(self._scale_percent + TABLE_SCALE_STEP)

    def decrease(self) -> None:
        self.set_scale(self._scale_percent - TABLE_SCALE_STEP)

    def reset(self) -> None:
        self.set_scale(DEFAULT_TABLE_SCALE)

    def set_scale(self, value: int) -> None:
        value = self._normalise_scale(value)
        if value == self._scale_percent:
            return

        self._scale_percent = value
        self._settings.setValue("ui/table_scale_percent", value)

        for state in list(self._states.values()):
            try:
                self._apply_table_scale(state)
            except RuntimeError:
                self._remove_state(id(state.table))

        self.scale_changed.emit(value)

    def register_tables(self, root: QWidget) -> None:
        if isinstance(root, QTableView):
            self.register_table(root)

        for table in root.findChildren(QTableView):
            self.register_table(table)

    def register_table(self, table: QTableView) -> None:
        install_gui_header_style(table)
        table_id = id(table)
        if table_id in self._states:
            self._capture_missing_columns(self._states[table_id])
            self._apply_table_scale(self._states[table_id])
            return

        header = table.horizontalHeader()
        header_height = max(header.height(), header.sizeHint().height(), 24)
        original_item_font_pt = table.font().pointSizeF()
        if original_item_font_pt <= 0:
            original_item_font_pt = FALLBACK_ITEM_FONT_PT

        state = _TableScaleState(
            table=table,
            base_header_height=header_height,
            base_item_font_pt=float(original_item_font_pt),
            base_style_sheet=table.styleSheet(),
            original_resize_columns_to_contents=table.resizeColumnsToContents,
            original_set_column_width=table.setColumnWidth,
        )
        self._states[table_id] = state

        self._capture_column_widths(state, sizes_are_scaled=False)
        self._install_width_wrappers(state)
        self._connect_model_signals(state)

        header.sectionResized.connect(
            lambda logical_index, old_size, new_size, current_id=table_id: self._on_section_resized(
                current_id,
                logical_index,
                old_size,
                new_size,
            )
        )
        header.sectionCountChanged.connect(
            lambda _old_count, _new_count, current_id=table_id: QTimer.singleShot(
                0,
                lambda: self._refresh_columns(current_id),
            )
        )
        table.destroyed.connect(lambda _obj=None, current_id=table_id: self._remove_state(current_id))

        self._apply_table_scale(state)

    def refresh_table(self, table: QTableView) -> None:
        state = self._states.get(id(table))
        if state is None:
            self.register_table(table)
            return
        self._capture_column_widths(state, sizes_are_scaled=True)
        self._apply_table_scale(state)

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        event_type = event.type()

        if event_type == QEvent.Type.Show and isinstance(obj, QTableView):
            self.register_table(obj)

        if event_type == QEvent.Type.Wheel:
            table = self._find_parent_table(obj)
            if table is not None and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
                if event.angleDelta().y() > 0:
                    self.increase()
                elif event.angleDelta().y() < 0:
                    self.decrease()
                event.accept()
                return True

        if event_type in (QEvent.Type.Show, QEvent.Type.Polish):
            widget = obj if isinstance(obj, QWidget) else None
            if widget is not None:
                table = self._find_parent_table(widget)
                if table is not None and table is not widget:
                    self._scale_table_editor(widget)

        return super().eventFilter(obj, event)


    def _connect_model_signals(self, state: _TableScaleState) -> None:
        model = state.table.model()
        if model is None:
            return

        table_id = id(state.table)

        def schedule_refresh(*_args) -> None:
            QTimer.singleShot(0, lambda: self._refresh_header(table_id))

        model.headerDataChanged.connect(schedule_refresh)
        model.modelReset.connect(schedule_refresh)
        model.columnsInserted.connect(schedule_refresh)
        model.columnsRemoved.connect(schedule_refresh)

    def _refresh_header(self, table_id: int) -> None:
        state = self._states.get(table_id)
        if state is None:
            return
        self._apply_table_scale(state)

    def _install_width_wrappers(self, state: _TableScaleState) -> None:
        table = state.table
        table_id = id(table)

        def scaled_resize_columns_to_contents() -> None:
            current_state = self._states.get(table_id)
            if current_state is None or current_state.original_resize_columns_to_contents is None:
                return

            current_state.applying = True
            try:
                current_state.original_resize_columns_to_contents()
            finally:
                current_state.applying = False

            self._capture_column_widths(current_state, sizes_are_scaled=True)
            self._apply_column_widths(current_state)

        def scaled_set_column_width(column: int, width: int) -> None:
            current_state = self._states.get(table_id)
            if current_state is None or current_state.original_set_column_width is None:
                return

            base_width = max(1, int(width))
            current_state.base_column_widths[int(column)] = base_width
            scaled_width = max(1, round(base_width * self.scale_factor))

            current_state.applying = True
            try:
                current_state.original_set_column_width(int(column), scaled_width)
            finally:
                current_state.applying = False

        # Pages already call these methods directly. Wrapping them here keeps all
        # existing page code unchanged while treating its dimensions as 100% values.
        table.resizeColumnsToContents = scaled_resize_columns_to_contents
        table.setColumnWidth = scaled_set_column_width

    def _apply_table_scale(self, state: _TableScaleState) -> None:
        table = state.table
        factor = self.scale_factor
        row_height = max(13, round(BASE_ROW_HEIGHT * factor))
        required_header_height = calculate_gui_header_base_height(
            table,
            base_font_point_size=BASE_HEADER_FONT_PT,
            minimum_height=state.base_header_height,
        )
        header_height = max(18, round(required_header_height * factor))

        state.applying = True
        try:
            vertical_header = table.verticalHeader()
            vertical_header.setDefaultSectionSize(row_height)
            vertical_header.setMinimumSectionSize(row_height)
            vertical_header.setMaximumSectionSize(row_height)

            horizontal_header = table.horizontalHeader()
            horizontal_header.setMinimumHeight(header_height)
            horizontal_header.setMaximumHeight(header_height)

            table_font = QFont(table.font())
            if table_font.pointSizeF() > 0:
                table_font.setPointSizeF(max(5.0, state.base_item_font_pt * factor))
                table.setFont(table_font)

            header_font = QFont(horizontal_header.font())
            if header_font.pointSizeF() > 0:
                header_font.setPointSizeF(max(6.0, BASE_HEADER_FONT_PT * factor))
                horizontal_header.setFont(header_font)

            table.setStyleSheet(self._build_scaled_style(state))
        finally:
            state.applying = False

        self._apply_column_widths(state)
        self._scale_existing_editors(table)
        table.viewport().update()
        table.horizontalHeader().viewport().update()

    def _apply_column_widths(self, state: _TableScaleState) -> None:
        if state.original_set_column_width is None:
            return

        self._capture_missing_columns(state)
        state.applying = True
        try:
            for column, base_width in state.base_column_widths.items():
                state.original_set_column_width(
                    column,
                    max(1, round(base_width * self.scale_factor)),
                )
        finally:
            state.applying = False

    def _capture_column_widths(
        self,
        state: _TableScaleState,
        *,
        sizes_are_scaled: bool,
    ) -> None:
        table = state.table
        factor = self.scale_factor if sizes_are_scaled else 1.0
        factor = factor or 1.0

        for column in range(self._column_count(table)):
            width = max(1, table.columnWidth(column))
            state.base_column_widths[column] = max(1, round(width / factor))

    def _capture_missing_columns(self, state: _TableScaleState) -> None:
        table = state.table
        for column in range(self._column_count(table)):
            if column not in state.base_column_widths:
                # A newly inserted section starts with Qt's unscaled default width.
                state.base_column_widths[column] = max(1, table.columnWidth(column))

        valid_columns = set(range(self._column_count(table)))
        for column in list(state.base_column_widths):
            if column not in valid_columns:
                state.base_column_widths.pop(column, None)

    def _on_section_resized(
        self,
        table_id: int,
        logical_index: int,
        _old_size: int,
        new_size: int,
    ) -> None:
        state = self._states.get(table_id)
        if state is None or state.applying:
            return

        if QApplication.mouseButtons() & Qt.MouseButton.LeftButton:
            # Manual resize: the user changes the width visible at the current scale.
            state.base_column_widths[logical_index] = max(
                1,
                round(new_size / (self.scale_factor or 1.0)),
            )
            return

        # A direct programmatic resizeSection call is treated as a 100% dimension.
        state.base_column_widths[logical_index] = max(1, int(new_size))
        self._apply_column_widths(state)

    def _refresh_columns(self, table_id: int) -> None:
        state = self._states.get(table_id)
        if state is None:
            return
        self._capture_missing_columns(state)
        self._apply_column_widths(state)

    def _build_scaled_style(self, state: _TableScaleState) -> str:
        table = state.table
        factor = self.scale_factor
        object_name = table.objectName()
        selector = f"QTableWidget#{object_name}" if isinstance(table, QTableWidget) and object_name else "QTableWidget"
        if not isinstance(table, QTableWidget):
            selector = f"QTableView#{object_name}" if object_name else "QTableView"

        item_font = max(5.0, state.base_item_font_pt * factor)
        header_font = max(6.0, BASE_HEADER_FONT_PT * factor)
        header_padding = max(1, round(BASE_HEADER_PADDING_PX * factor))
        editor_height = max(13, round(BASE_ROW_HEIGHT * factor) - 2)

        scale_style = f"""
/* TABLE_SCALE_RULES */
{selector}::item {{
    font-size: {item_font:.2f}pt;
}}
{selector} QHeaderView::section:horizontal {{
    font-size: {header_font:.2f}pt;
    padding: {header_padding}px;
}}
{selector} QLineEdit,
{selector} QComboBox,
{selector} QSpinBox,
{selector} QDoubleSpinBox {{
    font-size: {item_font:.2f}pt;
    min-height: {editor_height}px;
    max-height: {editor_height}px;
}}
"""
        return f"{state.base_style_sheet}\n{scale_style}".strip()

    def _scale_existing_editors(self, table: QTableView) -> None:
        # PySide6 QObject.findChildren() accepts one Qt type at a time.
        # Passing a Python tuple, as with isinstance(), raises TypeError.
        for editor_type in (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox):
            for editor in table.findChildren(editor_type):
                self._scale_table_editor(editor)

    def _scale_table_editor(self, widget: QWidget) -> None:
        if not isinstance(widget, (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox)):
            return

        base_font_size = widget.property("table_scale_base_font_size")
        current_font = QFont(widget.font())
        if base_font_size is None:
            table = self._find_parent_table(widget)
            table_state = self._states.get(id(table)) if table is not None else None
            if table_state is not None:
                base_font_size = table_state.base_item_font_pt
            else:
                point_size = current_font.pointSizeF()
                base_font_size = point_size if point_size > 0 else FALLBACK_ITEM_FONT_PT
            widget.setProperty("table_scale_base_font_size", float(base_font_size))

        current_font.setPointSizeF(max(5.0, float(base_font_size) * self.scale_factor))
        widget.setFont(current_font)

    @staticmethod
    def _column_count(table: QTableView) -> int:
        if isinstance(table, QTableWidget):
            return table.columnCount()
        model = table.model()
        return model.columnCount() if model is not None else 0

    @staticmethod
    def _find_parent_table(obj: QObject) -> QTableView | None:
        current = obj if isinstance(obj, QWidget) else None
        while current is not None:
            if isinstance(current, QTableView):
                return current
            current = current.parentWidget()
        return None

    def _remove_state(self, table_id: int) -> None:
        self._states.pop(table_id, None)

    @staticmethod
    def _normalise_scale(value: int) -> int:
        try:
            value = int(value)
        except (TypeError, ValueError):
            value = DEFAULT_TABLE_SCALE

        value = max(MIN_TABLE_SCALE, min(MAX_TABLE_SCALE, value))
        return int(round(value / TABLE_SCALE_STEP) * TABLE_SCALE_STEP)


_manager: TableScaleManager | None = None


def initialise_table_scale_manager(app: QApplication) -> TableScaleManager:
    global _manager
    if _manager is None:
        _manager = TableScaleManager(app)
    return _manager


def get_table_scale_manager() -> TableScaleManager | None:
    return _manager


def register_table_for_scaling(table: QAbstractItemView) -> None:
    if _manager is not None and isinstance(table, QTableView):
        _manager.register_table(table)
