from pathlib import Path
from decimal import Decimal, InvalidOperation

from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QMessageBox,
    QHeaderView,
    QTableWidget,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QApplication,
    QVBoxLayout,
    QComboBox,
    QLineEdit,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import MarkingRate
from app.db.db import SessionLocal
from app.ui.table_style import *


BASE_DIR = Path(__file__).resolve().parents[2]
MARKING_RATES_UI = BASE_DIR / "app" / "ui" / "windows" / "marking_rates.ui"


def load_ui(ui_path: Path):
    loader = QUiLoader()
    ui_file = QFile(str(ui_path))
    if not ui_file.open(QFile.ReadOnly):
        raise RuntimeError(f"Не удалось открыть UI: {ui_path}")
    try:
        widget = loader.load(ui_file)
    finally:
        ui_file.close()

    if widget is None:
        raise RuntimeError(f"Не удалось загрузить UI: {ui_path}")

    return widget


class MarkingRatesPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(MARKING_RATES_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self.filter_pack_type = self._pick_widget(
            ["line_PackType", "line_Pack_name", "line_Name"],
            required=False,
        )

        self._updating_table = False
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

        self.columns = ["pack_type", "cost_per_l"]
        self.headers = ["Pack type", "Cost per L"]
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def _pick_widget(self, names, required=True):
        for name in names:
            widget = getattr(self.ui, name, None)
            if widget is not None:
                return widget
        if required:
            raise AttributeError(f"Не найден widget: {', '.join(names)}")
        return None

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.btn_Search.clicked.connect(self.find_marking_rates)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

    def get_session(self):
        return SessionLocal()

    def _get_filter_value(self, widget):
        if widget is None:
            return "-"
        if isinstance(widget, QComboBox):
            return widget.currentText()
        if isinstance(widget, QLineEdit):
            text = widget.text().strip()
            return text if text else "-"
        return "-"

    def _set_filter_items(self, widget, items):
        if widget is None or not isinstance(widget, QComboBox):
            return
        current_value = widget.currentText()
        widget.blockSignals(True)
        widget.clear()
        widget.addItem("-")
        if items:
            widget.addItems(sorted(items))
        if current_value in items:
            widget.setCurrentText(current_value)
        widget.blockSignals(False)

    def _to_decimal(self, value, field_name):
        if isinstance(value, Decimal):
            return value

        text = str(value).strip().replace(",", ".")
        if text == "":
            return Decimal("0")

        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            raise Exception(f"Поле '{field_name}' должно быть числом")

    def _get_row_key(self, row):
        item = self.table.item(row, 0)
        if not item:
            return None
        return item.data(Qt.UserRole)

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        delete_action = menu.addAction("Удалить строку")
        apply_action = menu.addAction("Применить изменения")
        revert_action = menu.addAction("Отменить изменения")

        copy_action.triggered.connect(self.copy_cell_content)
        delete_action.triggered.connect(self.delete_selected_row)
        apply_action.triggered.connect(self.apply_pending_changes)
        revert_action.triggered.connect(self.revert_changes)

        menu.exec_(self.table.viewport().mapToGlobal(position))

    def on_item_changed(self, item):
        if self._updating_table:
            return

        try:
            row = item.row()
            header_item = self.table.horizontalHeaderItem(item.column())
            if not header_item:
                return

            row_key = self._get_row_key(row)
            if row_key is None:
                return

            column_name = self.header_to_column.get(header_item.text())
            if not column_name:
                return

            if row_key not in self._pending_changes:
                self._pending_changes[row_key] = {}

            self._pending_changes[row_key][column_name] = item.text()

        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        clipboard = QApplication.clipboard()

        if len(selected_items) == 1:
            text = selected_items[0].text()
        else:
            rows = {}
            for item in selected_items:
                rows.setdefault(item.row(), {})
                rows[item.row()][item.column()] = item.text()

            text = "\n".join(
                "\t".join(value for _, value in sorted(cols.items()))
                for _, cols in sorted(rows.items())
            )

        clipboard.setText(text.strip())
        self.show_message("Скопировано")

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0:
            self.show_error_message("Не выбрана строка для удаления")
            return

        row_key = self._get_row_key(row)
        if row_key is None:
            self.show_error_message("Не удалось определить строку")
            return

        if row_key in self._new_rows:
            self._new_rows.discard(row_key)
            self._pending_changes.pop(row_key, None)
        else:
            self._pending_deletes.add(row_key)

        self.table.removeRow(row)
        self.show_message("Строка помечена на удаление")

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self.has_active_filters():
                self.find_marking_rates()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Изменения отменены")
        except Exception as e:
            self.show_error_message(f"Ошибка отката: {str(e)}")

    def apply_pending_changes(self):
        if not self._pending_changes and not self._pending_deletes:
            self.show_message("Нет изменений для применения")
            return

        try:
            with self.get_session() as session:
                if self._pending_deletes:
                    session.query(MarkingRate).filter(
                        MarkingRate.pack_type.in_(list(self._pending_deletes))
                    ).delete(synchronize_session=False)

                for row_key, changes in self._pending_changes.items():
                    if row_key in self._new_rows:
                        pack_type = str(changes.get("pack_type", "")).strip()
                        if not pack_type:
                            raise Exception("Для новой строки поле Pack type обязательно")

                        existing = (
                            session.query(MarkingRate)
                            .filter(MarkingRate.pack_type == pack_type)
                            .first()
                        )
                        if existing:
                            raise Exception(f"Pack type '{pack_type}' уже существует")

                        row = MarkingRate(
                            pack_type=pack_type,
                            cost_per_l=self._to_decimal(changes.get("cost_per_l", 0), "Cost per L"),
                        )
                        session.add(row)
                    else:
                        row = (
                            session.query(MarkingRate)
                            .filter(MarkingRate.pack_type == row_key)
                            .first()
                        )
                        if not row:
                            raise Exception(f"Не найден MarkingRate '{row_key}'")

                        if "cost_per_l" in changes:
                            row.cost_per_l = self._to_decimal(changes["cost_per_l"], "Cost per L")

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.refresh_all_comboboxes()

            if self.has_active_filters():
                self.find_marking_rates()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def refresh_all_comboboxes(self):
        if self.filter_pack_type is None:
            return

        try:
            with self.get_session() as session:
                rows = session.query(MarkingRate.pack_type).order_by(MarkingRate.pack_type).all()

            items = [row[0] for row in rows if row[0]]
            self._set_filter_items(self.filter_pack_type, items)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении Pack type: {str(e)}")

    def get_marking_rates_from_db(self):
        with self.get_session() as session:
            rows = session.query(MarkingRate).order_by(MarkingRate.pack_type).all()

            data = []
            for row in rows:
                data.append({
                    "pack_type": row.pack_type,
                    "cost_per_l": row.cost_per_l,
                })

        return data

    def find_marking_rates(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        data = self.get_marking_rates_from_db()
        if not data:
            self.show_message("Нет данных для отображения")
            return

        pack_type_filter = self._get_filter_value(self.filter_pack_type)
        if pack_type_filter != "-":
            data = [row for row in data if (row["pack_type"] or "") == pack_type_filter]

        self._display_data(data)

        if not data:
            self.show_message("Нет данных по заданным фильтрам")

    def _build_item(self, value, editable=True, align_left=False, row_key=None):
        item = QTableWidgetItem(format_table_value(value))
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        if row_key is not None:
            item.setData(Qt.UserRole, row_key)
        return item

    def _display_data(self, data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        if not data:
            self.show_message("Нет данных для отображения")
            return

        self._updating_table = True

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(self.headers)

        for row_index, row_data in enumerate(data):
            row_key = row_data["pack_type"]

            pack_item = self._build_item(
                row_data["pack_type"],
                editable=False,
                align_left=True,
                row_key=row_key,
            )
            cost_item = self._build_item(
                row_data["cost_per_l"],
                editable=True,
                align_left=False,
                row_key=row_key,
            )

            self.table.setItem(row_index, 0, pack_item)
            self.table.setItem(row_index, 1, cost_item)

        self.table.resizeColumnsToContents()
        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False
        self.table.setSortingEnabled(True)

    def add_line(self):
        self._updating_table = True

        self.table.setSortingEnabled(False)
        self.table.insertRow(0)

        row_key = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_key)

        self._pending_changes[row_key] = {
            "pack_type": "",
            "cost_per_l": "0",
        }

        self.table.setItem(0, 0, self._build_item("", editable=True, align_left=True, row_key=row_key))
        self.table.setItem(0, 1, self._build_item("0", editable=True, align_left=False, row_key=row_key))

        self._updating_table = False
        self.table.setCurrentCell(0, 0)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return self._get_filter_value(self.filter_pack_type) != "-"

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setProperty("active", True)
        self.ui.label_msg.style().unpolish(self.ui.label_msg)
        self.ui.label_msg.style().polish(self.ui.label_msg)
        self.ui.label_msg.setVisible(True)

    def clear_message(self):
        self.ui.label_msg.setText("")
        self.ui.label_msg.setProperty("active", False)
        self.ui.label_msg.style().unpolish(self.ui.label_msg)
        self.ui.label_msg.style().polish(self.ui.label_msg)
        self.ui.label_msg.setVisible(False)

    def show_error_message(self, text):
        msg = QMessageBox()
        msg.setWindowTitle("Ошибка")
        msg.setIcon(QMessageBox.Critical)
        msg.setMinimumSize(900, 600)

        if len(text) > 500:
            msg.setText("Произошла ошибка. Подробности ниже (используйте кнопку 'Show Details')")
            msg.setDetailedText(text)
        else:
            msg.setText(text)

        copy_button = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)

        def copy_text():
            QApplication.clipboard().setText(text)

        copy_button.clicked.connect(copy_text)
        msg.exec_()