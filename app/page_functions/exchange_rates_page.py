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

from app.db.models import ExchangeRate
from app.db.db import SessionLocal
from app.ui.table_style import *
from app.utils.output_headers import display_headers


BASE_DIR = Path(__file__).resolve().parents[2]
EXCHANGE_RATES_UI = BASE_DIR / "app" / "ui" / "windows" / "exchange_rates.ui"


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


class ExchangeRatesPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(EXCHANGE_RATES_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self.filter_currency = self._pick_widget(["line_Currency", "line_currency", "line_Code"], required=False)

        self._updating_table = False
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

        self.columns = ["currency_code", "rate_to_rub"]
        self.headers = ["Currency", "Rate to RUB"]
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
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.btn_Search.clicked.connect(self.find_exchange_rates)
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

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self.has_active_filters():
                self.find_exchange_rates()
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
                    session.query(ExchangeRate).filter(
                        ExchangeRate.currency_code.in_(list(self._pending_deletes))
                    ).delete(synchronize_session=False)

                for row_key, changes in self._pending_changes.items():
                    if row_key in self._new_rows:
                        currency_code = str(changes.get("currency_code", "")).strip().upper()
                        if not currency_code:
                            raise Exception("Для новой строки поле Currency обязательно")

                        existing = (
                            session.query(ExchangeRate)
                            .filter(ExchangeRate.currency_code == currency_code)
                            .first()
                        )
                        if existing:
                            raise Exception(f"Валюта '{currency_code}' уже существует")

                        row = ExchangeRate(
                            currency_code=currency_code,
                            rate_to_rub=self._to_decimal(changes.get("rate_to_rub", 0), "Rate to RUB"),
                        )
                        session.add(row)
                    else:
                        row = (
                            session.query(ExchangeRate)
                            .filter(ExchangeRate.currency_code == row_key)
                            .first()
                        )
                        if not row:
                            raise Exception(f"Не найдена валюта '{row_key}'")

                        if "rate_to_rub" in changes:
                            row.rate_to_rub = self._to_decimal(changes["rate_to_rub"], "Rate to RUB")

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.refresh_all_comboboxes()

            if self.has_active_filters():
                self.find_exchange_rates()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def refresh_all_comboboxes(self):
        if self.filter_currency is None:
            return

        try:
            with self.get_session() as session:
                rows = session.query(ExchangeRate.currency_code).order_by(ExchangeRate.currency_code).all()

            items = [row[0] for row in rows if row[0]]
            self._set_filter_items(self.filter_currency, items)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении валют: {str(e)}")

    def get_exchange_rates_from_db(self):
        with self.get_session() as session:
            rows = session.query(ExchangeRate).all()

            data = []
            for row in rows:
                data.append({
                    "currency_code": row.currency_code,
                    "rate_to_rub": row.rate_to_rub,
                })

        return data

    def find_exchange_rates(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        data = self.get_exchange_rates_from_db()

        if not data:
            self.show_message("Нет данных для отображения")
            return

        currency_filter = self._get_filter_value(self.filter_currency)
        if currency_filter != "-":
            data = [row for row in data if (row["currency_code"] or "") == currency_filter]

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
        self.table.setHorizontalHeaderLabels(display_headers(self.headers))

        for row_index, row_data in enumerate(data):
            row_key = row_data["currency_code"]

            currency_item = self._build_item(
                row_data["currency_code"],
                editable=False,
                align_left=True,
                row_key=row_key,
            )
            rate_item = self._build_item(
                row_data["rate_to_rub"],
                editable=True,
                align_left=False,
                row_key=row_key,
            )

            self.table.setItem(row_index, 0, currency_item)
            self.table.setItem(row_index, 1, rate_item)

        self.table.resizeColumnsToContents()
        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False

    def add_line(self):
        self._updating_table = True

        self.table.setSortingEnabled(False)
        self.table.insertRow(0)

        row_key = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_key)

        self._pending_changes[row_key] = {
            "currency_code": "",
            "rate_to_rub": "0",
        }

        self.table.setItem(0, 0, self._build_item("", editable=True, align_left=True, row_key=row_key))
        self.table.setItem(0, 1, self._build_item("0", editable=True, align_left=False, row_key=row_key))

        self._updating_table = False
        self.table.setCurrentCell(0, 0)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return self._get_filter_value(self.filter_currency) != "-"

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
