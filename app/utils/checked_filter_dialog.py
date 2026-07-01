from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Sequence

from PySide6.QtCore import QFile, Qt, QSize
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QAbstractItemView, QDialog, QListWidgetItem, QWidget


BASE_DIR = Path(__file__).resolve().parents[2]
CHECKED_FILTER_DIALOG_UI = BASE_DIR / "app" / "ui" / "windows" / "checked_filter_dialog.ui"


def load_ui(ui_path: Path, parent=None):
    loader = QUiLoader()
    ui_file = QFile(str(ui_path))
    if not ui_file.open(QFile.ReadOnly):
        raise RuntimeError(f"Не удалось открыть UI: {ui_path}")
    try:
        widget = loader.load(ui_file, parent)
    finally:
        ui_file.close()

    if widget is None:
        raise RuntimeError(f"Не удалось загрузить UI: {ui_path}")

    return widget


@dataclass(frozen=True)
class FilterOption:
    key: Any
    label: str
    search_text: str = ""


class CheckedFilterDialog:
    def __init__(
        self,
        parent: QWidget,
        *,
        title: str,
        options: Sequence[FilterOption],
        selected_keys: Optional[set[Any]],
    ):
        self.ui = load_ui(CHECKED_FILTER_DIALOG_UI, parent)
        self.options = list(options)
        self.selected_keys = None if selected_keys is None else set(selected_keys)

        self.ui.setWindowTitle(title)
        self.ui.label_Title.setText(title)
        self.ui.lst_Values.setSelectionMode(QAbstractItemView.NoSelection)
        self.ui.lst_Values.setSpacing(0)

        self._fill_values()
        self.ui.line_Search.textChanged.connect(self._apply_search_filter)
        self.ui.btn_SelectAll.clicked.connect(lambda: self._set_visible_items_checked(True))
        self.ui.btn_ClearAll.clicked.connect(lambda: self._set_visible_items_checked(False))
        self.ui.btn_Ok.clicked.connect(self.ui.accept)
        self.ui.btn_Cancel.clicked.connect(self.ui.reject)
        self.ui.lst_Values.itemChanged.connect(lambda _item: self._update_count_label())

        self._update_count_label()

    def exec_and_get_selection(self) -> tuple[bool, Optional[set[Any]]]:
        if self.ui.exec() != QDialog.Accepted:
            return False, None

        checked_keys: set[Any] = set()
        for index in range(self.ui.lst_Values.count()):
            item = self.ui.lst_Values.item(index)
            if item and item.checkState() == Qt.Checked:
                checked_keys.add(item.data(Qt.UserRole))

        if len(checked_keys) == len(self.options):
            return True, None
        return True, checked_keys

    def _fill_values(self) -> None:
        all_checked = self.selected_keys is None
        self.ui.lst_Values.blockSignals(True)
        self.ui.lst_Values.clear()

        for option in self.options:
            item = QListWidgetItem(option.label)
            item.setData(Qt.UserRole, option.key)
            item.setToolTip(option.search_text or option.label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            item.setCheckState(Qt.Checked if all_checked or option.key in self.selected_keys else Qt.Unchecked)
            item.setSizeHint(QSize(0, 20))
            self.ui.lst_Values.addItem(item)

        self.ui.lst_Values.blockSignals(False)

    def _apply_search_filter(self) -> None:
        # Search works the same way as product search fields above dropdown lists:
        # the entered text is treated as one normalized phrase and is matched
        # against the visible option label.
        #
        # Previous logic split the query into separate tokens and searched in
        # the extended tooltip text. For product filters this returned too many
        # rows: e.g. "Omala S4 GXV 460" could also show products where
        # "Omala S4 GXV" came from family and "460" came from another
        # field.
        query = self._normalize_search_text(self.ui.line_Search.text())

        for index in range(self.ui.lst_Values.count()):
            item = self.ui.lst_Values.item(index)
            if not item:
                continue
            haystack = self._normalize_search_text(item.text())
            item.setHidden(bool(query) and query not in haystack)

        self._update_count_label()

    @staticmethod
    def _normalize_search_text(value: object) -> str:
        return " ".join(str(value or "").split()).casefold()

    def _set_visible_items_checked(self, checked: bool) -> None:
        state = Qt.Checked if checked else Qt.Unchecked
        self.ui.lst_Values.blockSignals(True)
        for index in range(self.ui.lst_Values.count()):
            item = self.ui.lst_Values.item(index)
            if item and not item.isHidden():
                item.setCheckState(state)
        self.ui.lst_Values.blockSignals(False)
        self._update_count_label()

    def _update_count_label(self) -> None:
        total = self.ui.lst_Values.count()
        visible = 0
        checked = 0
        for index in range(total):
            item = self.ui.lst_Values.item(index)
            if not item:
                continue
            if not item.isHidden():
                visible += 1
            if item.checkState() == Qt.Checked:
                checked += 1

        self.ui.label_Count.setText(f"Выбрано: {checked} из {total}. Показано: {visible}")
