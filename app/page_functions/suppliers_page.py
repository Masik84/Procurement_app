from pathlib import Path
from decimal import Decimal, InvalidOperation

from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QMessageBox,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QApplication,
    QVBoxLayout,
    QCheckBox,
    QHBoxLayout,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import Supplier
from app.db.db import SessionLocal
from app.ui.table_style import *


BASE_DIR = Path(__file__).resolve().parents[2]
SUPPLIERS_UI = BASE_DIR / "app" / "ui" / "windows" / "supplier.ui"


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


class SuppliersPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(SUPPLIERS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._updating_table = False
        self._original_values = {}
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

        self.columns = [
            "id",
            "name",
            "base_currency",
            "transport_cost_per_l",
            "reexport_percent",
            "fx_rate_markup",
            "agent_fee",
            "is_via_novo",
            "has_import_duty",
            "rating_calc",
            "marks_for_us",
            "is_rf",
            "country",
        ]

        self.headers = [
            "id",
            "Supplier name",
            "Base currency",
            "Transport cost per L",
            "Reexport percent",
            "FX rate markup",
            "Agent fee",
            "Via Novo",
            "Import duty",
            "Rating calc",
            "Marks for us",
            "RF",
            "Country",
        ]

        self.text_columns = {"name", "base_currency", "country"}
        self.numeric_columns = {
            "transport_cost_per_l",
            "reexport_percent",
            "fx_rate_markup",
            "agent_fee",
        }
        self.bool_columns = {
            "is_via_novo",
            "has_import_duty",
            "rating_calc",
            "marks_for_us",
            "is_rf",
        }
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=True)
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.btn_Search.clicked.connect(self.find_supplier)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

    def get_session(self):
        return SessionLocal()

    def on_item_changed(self, item):
        if self._updating_table:
            return

        try:
            row = item.row()
            column = item.column()
            header_item = self.table.horizontalHeaderItem(column)
            id_item = self.table.item(row, 0)

            if not header_item or not id_item:
                return

            row_id_text = id_item.text().strip()
            if not row_id_text:
                return

            row_id = int(row_id_text)
            header = header_item.text()
            column_name = self.header_to_column.get(header)

            if not column_name or column_name == "id":
                return

            new_value = item.text()

            if row_id not in self._pending_changes:
                self._pending_changes[row_id] = {}

            self._pending_changes[row_id][column_name] = new_value

        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def on_checkbox_changed(self, row_id, column_name, checked):
        if self._updating_table:
            return

        try:
            if row_id not in self._pending_changes:
                self._pending_changes[row_id] = {}

            self._pending_changes[row_id][column_name] = bool(checked)

        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def _get_cell_display_text(self, row, col):
        column_name = self.columns[col]

        if column_name in self.bool_columns:
            checkbox = self._get_checkbox_from_cell(row, col)
            return "Да" if checkbox and checkbox.isChecked() else "Нет"

        item = self.table.item(row, col)
        return item.text() if item else ""

    def _get_checkbox_from_cell(self, row, col):
        container = self.table.cellWidget(row, col)
        if not container:
            return None
        return container.findChild(QCheckBox)

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self.has_active_filters():
                self.find_supplier()
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
                    session.query(Supplier).filter(
                        Supplier.id.in_(self._pending_deletes)
                    ).delete(synchronize_session=False)

                for row_id, changes in self._pending_changes.items():
                    if row_id in self._new_rows:
                        self._insert_supplier(session, changes)
                    else:
                        self._update_supplier(session, row_id, changes)

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.refresh_all_comboboxes()

            if self.has_active_filters():
                self.find_supplier()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def _insert_supplier(self, session, changes):
        name = str(changes.get("name", "")).strip()
        base_currency = str(changes.get("base_currency", "")).strip()
        country = str(changes.get("country", "")).strip()

        if not name:
            raise Exception("Для новой строки поле Supplier name обязательно")

        if not base_currency:
            raise Exception("Для новой строки поле Base currency обязательно")

        existing = session.query(Supplier).filter(Supplier.name == name).first()
        if existing:
            raise Exception(f"Поставщик с name '{name}' уже существует")

        supplier = Supplier(
            name=name,
            base_currency=base_currency,
            transport_cost_per_l=self._to_decimal(
                changes.get("transport_cost_per_l", 0),
                "Transport cost per L",
            ),
            reexport_percent=self._to_percent_decimal(
                changes.get("reexport_percent", 0),
                "Reexport percent",
            ),
            fx_rate_markup=self._to_percent_decimal(
                changes.get("fx_rate_markup", 0),
                "FX rate markup",
            ),
            agent_fee=self._to_decimal(
                changes.get("agent_fee", 0),
                "Agent fee",
            ),
            is_via_novo=bool(changes.get("is_via_novo", False)),
            has_import_duty=bool(changes.get("has_import_duty", False)),
            rating_calc=bool(changes.get("rating_calc", True)),
            marks_for_us=bool(changes.get("marks_for_us", False)),
            is_rf=bool(changes.get("is_rf", False)),
            country=country if country else None,
        )
        session.add(supplier)

    def _update_supplier(self, session, row_id, changes):
        supplier = session.query(Supplier).filter(Supplier.id == row_id).first()
        if not supplier:
            raise Exception(f"Не найден supplier id={row_id}")

        if "name" in changes:
            new_name = str(changes["name"]).strip()
            if not new_name:
                raise Exception("Поле Supplier name не может быть пустым")

            duplicate = (
                session.query(Supplier)
                .filter(Supplier.name == new_name, Supplier.id != row_id)
                .first()
            )
            if duplicate:
                raise Exception(f"Поставщик с name '{new_name}' уже существует")

            supplier.name = new_name

        if "base_currency" in changes:
            value = str(changes["base_currency"]).strip()
            if not value:
                raise Exception("Поле Base currency не может быть пустым")
            supplier.base_currency = value

        if "transport_cost_per_l" in changes:
            supplier.transport_cost_per_l = self._to_decimal(
                changes["transport_cost_per_l"],
                "Transport cost per L",
            )

        if "reexport_percent" in changes:
            supplier.reexport_percent = self._to_percent_decimal(
                changes["reexport_percent"],
                "Reexport percent",
            )

        if "fx_rate_markup" in changes:
            supplier.fx_rate_markup = self._to_percent_decimal(
                changes["fx_rate_markup"],
                "FX rate markup",
            )

        if "agent_fee" in changes:
            supplier.agent_fee = self._to_decimal(
                changes["agent_fee"],
                "Agent fee",
            )

        if "is_via_novo" in changes:
            supplier.is_via_novo = bool(changes["is_via_novo"])

        if "has_import_duty" in changes:
            supplier.has_import_duty = bool(changes["has_import_duty"])

        if "rating_calc" in changes:
            supplier.rating_calc = bool(changes["rating_calc"])

        if "marks_for_us" in changes:
            supplier.marks_for_us = bool(changes["marks_for_us"])

        if "is_rf" in changes:
            supplier.is_rf = bool(changes["is_rf"])

        if "country" in changes:
            value = str(changes["country"]).strip()
            supplier.country = value if value else None

    def _to_decimal(self, value, field_name):
        if isinstance(value, Decimal):
            return value

        text = str(value).strip()

        if "%" in text:
            text = text.replace("%", "").strip()
            text = text.replace(",", ".")
            return Decimal(text) / Decimal("100")

        text = text.replace(",", ".")

        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            raise Exception(f"Поле '{field_name}' должно быть числом")

    def _to_percent_decimal(self, value, field_name):
        if isinstance(value, Decimal):
            decimal_value = value
            has_percent_sign = False
        else:
            text = str(value).strip()
            has_percent_sign = "%" in text
            text = text.replace("%", "").strip().replace(",", ".")
            try:
                decimal_value = Decimal(text)
            except (InvalidOperation, ValueError):
                raise Exception(f"Поле '{field_name}' должно быть числом")

        # В БД проценты хранятся долей: 1% = 0.01, 100% = 1.
        # Поэтому ввод в процентных колонках трактуем именно как проценты.
        # 1 / 1% -> 0.01; 100 / 100% -> 1; 0,01 остается 0,01.
        if has_percent_sign or abs(decimal_value) >= Decimal("1"):
            decimal_value = decimal_value / Decimal("100")
        return decimal_value

    def refresh_all_comboboxes(self):
        self.fill_in_supplier_list()

    def fill_in_supplier_list(self):
        current_value = self.ui.line_Suppl1.currentText()

        try:
            with self.get_session() as session:
                suppliers = (
                    session.query(Supplier.name)
                    .filter(Supplier.name.isnot(None), Supplier.name != "", Supplier.name != "Manual")
                    .distinct()
                    .order_by(Supplier.name)
                    .all()
                )

            supplier_names = [row[0] for row in suppliers if row[0]]
            self._fill_combobox(self.ui.line_Suppl1, supplier_names)

            if current_value in supplier_names:
                self.ui.line_Suppl1.setCurrentText(current_value)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {str(e)}")
            self._fill_combobox(self.ui.line_Suppl1, [])

    def _fill_combobox(self, combobox, items):
        combobox.blockSignals(True)
        combobox.clear()
        combobox.addItem("-")
        if items:
            combobox.addItems(sorted(items))
        combobox.blockSignals(False)

    def get_suppliers_from_db(self):
        with self.get_session() as session:
            rows = session.query(Supplier).filter(Supplier.name != "Manual").order_by(Supplier.name).all()

            data = []
            for row in rows:
                data.append({
                    "id": row.id,
                    "name": row.name,
                    "base_currency": row.base_currency,
                    "transport_cost_per_l": row.transport_cost_per_l,
                    "reexport_percent": row.reexport_percent,
                    "fx_rate_markup": row.fx_rate_markup,
                    "agent_fee": row.agent_fee,
                    "is_via_novo": bool(row.is_via_novo),
                    "has_import_duty": bool(row.has_import_duty),
                    "rating_calc": bool(row.rating_calc),
                    "marks_for_us": bool(row.marks_for_us),
                    "is_rf": bool(row.is_rf),
                    "country": row.country,
                })

        return data

    def find_supplier(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        supplier_data = self.get_suppliers_from_db()

        if not supplier_data:
            self.show_message("Нет данных для отображения")
            return

        supplier_name = self.ui.line_Suppl1.currentText()

        if supplier_name != "-":
            supplier_data = [row for row in supplier_data if (row["name"] or "") == supplier_name]

        self._display_data(supplier_data)

        if not supplier_data:
            self.show_message("Нет данных по заданным фильтрам")

    def _display_data(self, data):
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

        if not data:
            self.show_message("Нет данных для отображения")
            self.table.setSortingEnabled(True)
            return

        self._updating_table = True
        self._original_values.clear()

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(self.headers)

        setup_data_table(self.table, sorting=True)

        for row_index, row_data in enumerate(data):
            row_id = int(row_data["id"])
            self._original_values[row_id] = {}

            for col_index, col_name in enumerate(self.columns):
                value = row_data.get(col_name)

                if col_name in self.bool_columns:
                    checked = bool(value)
                    self._original_values[row_id][col_name] = checked
                    self.table.setCellWidget(
                        row_index,
                        col_index,
                        self._build_checkbox_widget(row_id, col_name, checked)
                    )
                    continue

                if col_name in {"reexport_percent", "fx_rate_markup"}:
                    text_value = self._format_percent(value)
                else:
                    text_value = "" if value is None else str(value).replace(".", ",")

                item = self._build_table_item(col_name, text_value)
                self._original_values[row_id][col_name] = text_value
                self.table.setItem(row_index, col_index, item)

        self.table.resizeColumnsToContents()

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 100:
                self.table.setColumnWidth(i, 100)

        self._updating_table = False
        self.table.setSortingEnabled(True)

    def _build_checkbox_widget(self, row_id, column_name, checked):
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setStyleSheet("""
            QCheckBox {
                background: transparent;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
            }
        """)
        checkbox.toggled.connect(
            lambda state, rid=row_id, col=column_name: self.on_checkbox_changed(rid, col, state)
        )

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)

        return container

    def _build_table_item(self, col_name, value):
        item = QTableWidgetItem(value)

        if col_name == "id":
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            item.setTextAlignment(Qt.AlignCenter)
            return item

        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)

        if col_name in self.numeric_columns:
            item.setTextAlignment(Qt.AlignCenter)
        elif col_name in self.text_columns:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        else:
            item.setTextAlignment(Qt.AlignCenter)

        return item

    def _format_percent(self, value):
        if value is None:
            return ""

        try:
            return f"{float(value) * 100:.1f}".replace(".", ",") + "%"
        except Exception:
            return str(value)

    def add_line(self):
        self._updating_table = True

        self.table.setSortingEnabled(False)

        if self.table.columnCount() == 0:
            self.table.setColumnCount(len(self.headers))
            self.table.setHorizontalHeaderLabels(self.headers)
            setup_data_table(self.table, sorting=True)

        self.table.insertRow(0)

        row_id = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_id)

        values = {
            "id": str(row_id),
            "name": "",
            "base_currency": "",
            "transport_cost_per_l": "0",
            "reexport_percent": "0",
            "fx_rate_markup": "0",
            "agent_fee": "0",
            "is_via_novo": False,
            "has_import_duty": False,
            "rating_calc": True,
            "marks_for_us": False,
            "is_rf": False,
            "country": "",
        }

        self._pending_changes[row_id] = {
            "name": "",
            "base_currency": "",
            "transport_cost_per_l": "0",
            "reexport_percent": "0",
            "fx_rate_markup": "0",
            "agent_fee": "0",
            "is_via_novo": False,
            "has_import_duty": False,
            "rating_calc": True,
            "marks_for_us": False,
            "is_rf": False,
            "country": "",
        }

        for col_index, col_name in enumerate(self.columns):
            if col_name in self.bool_columns:
                self.table.setCellWidget(
                    0,
                    col_index,
                    self._build_checkbox_widget(row_id, col_name, bool(values[col_name]))
                )
                continue

            item = self._build_table_item(col_name, values.get(col_name, ""))
            self.table.setItem(0, col_index, item)

        self._updating_table = False
        self.table.setSortingEnabled(True)
        self.table.setCurrentCell(0, 1)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return self.ui.line_Suppl1.currentText() != "-"

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
