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
from app.utils.parsers import parse_user_percent
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption
from app.utils.money import parse_decimal_field


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
        self._selected_supplier_ids: set[int] | None = None
        self._selected_country_values: set[str] | None = None

        self.columns = [
            "id",
            "name",
            "base_currency",
            "transport_cost_per_l",
            "reexport_percent",
            "insurance_percent",
            "fx_rate_markup",
            "fx_rate_markup_abs",
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
            "Re-export %",
            "Insurance %",
            "FX markup %",
            "FX markup abs",
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
            "insurance_percent",
            "fx_rate_markup",
            "fx_rate_markup_abs",
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

        self.ui.btn_FilterSupplier.clicked.connect(self.open_supplier_filter)
        self.ui.btn_FilterCountry.clicked.connect(self.open_country_filter)
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
            header = table_header_name(self.table, column)
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
            insurance_percent=self._to_percent_decimal(
                changes.get("insurance_percent", 0),
                "Insurance %",
            ),
            fx_rate_markup=self._to_percent_decimal(
                changes.get("fx_rate_markup", 0),
                "FX markup %",
            ),
            fx_rate_markup_abs=self._to_decimal(
                changes.get("fx_rate_markup_abs", 0),
                "FX markup abs",
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

        if "insurance_percent" in changes:
            supplier.insurance_percent = self._to_percent_decimal(
                changes["insurance_percent"],
                "Insurance %",
            )

        if "fx_rate_markup" in changes:
            supplier.fx_rate_markup = self._to_percent_decimal(
                changes["fx_rate_markup"],
                "FX markup %",
            )

        if "fx_rate_markup_abs" in changes:
            supplier.fx_rate_markup_abs = self._to_decimal(
                changes["fx_rate_markup_abs"],
                "FX markup abs",
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
        # Only called from apply_pending_changes (final Save). Supports a
        # trailing "%" shorthand (e.g. "3,5%" -> 0.035); base parsing and the
        # NOT NULL empty check are delegated to the shared helper.
        if isinstance(value, Decimal):
            return value

        text = str(value).strip()
        if "%" in text:
            return parse_decimal_field(
                text.replace("%", "").strip(), field_name, empty="raise"
            ) / Decimal("100")

        return parse_decimal_field(value, field_name, empty="raise")

    def _to_percent_decimal(self, value, field_name):
        decimal_value = parse_user_percent(value)
        if decimal_value is None:
            raise Exception(f"Поле '{field_name}' должно быть числом")
        return decimal_value

    def refresh_all_comboboxes(self):
        self._prune_filter_selections()
        self._refresh_filter_buttons()

    def _filter_option_rows(self, *, ignore: str | None = None) -> list[dict]:
        rows = self.get_suppliers_from_db()
        if ignore != "supplier" and self._selected_supplier_ids is not None:
            rows = [row for row in rows if int(row["id"]) in self._selected_supplier_ids]
        if ignore != "country" and self._selected_country_values is not None:
            rows = [row for row in rows if (row["country"] or "") in self._selected_country_values]
        return rows

    def open_supplier_filter(self):
        rows = self._filter_option_rows(ignore="supplier")
        options = [
            FilterOption(key=int(row["id"]), label=row["name"], search_text=row["name"])
            for row in sorted(rows, key=lambda item: (item["name"] or "").casefold())
            if row["name"]
        ]
        accepted, selected = CheckedFilterDialog(
            self,
            title="Фильтр по поставщикам",
            options=options,
            selected_keys=self._selected_supplier_ids,
        ).exec_and_get_selection()
        if accepted:
            self._selected_supplier_ids = None if selected is None else {int(value) for value in selected}
            self._prune_filter_selections()
            self._refresh_filter_buttons()

    def open_country_filter(self):
        countries = sorted({
            row["country"] for row in self._filter_option_rows(ignore="country") if row["country"]
        }, key=str.casefold)
        accepted, selected = CheckedFilterDialog(
            self,
            title="Фильтр по странам",
            options=[FilterOption(key=value, label=value, search_text=value) for value in countries],
            selected_keys=self._selected_country_values,
        ).exec_and_get_selection()
        if accepted:
            self._selected_country_values = None if selected is None else {str(value) for value in selected}
            self._prune_filter_selections()
            self._refresh_filter_buttons()

    def _prune_filter_selections(self):
        rows = self.get_suppliers_from_db()
        supplier_ids = {int(row["id"]) for row in rows}
        countries = {row["country"] for row in rows if row["country"]}
        if self._selected_supplier_ids is not None:
            self._selected_supplier_ids &= supplier_ids
        if self._selected_country_values is not None:
            self._selected_country_values &= countries

    def _refresh_filter_buttons(self):
        self._set_filter_button_text(self.ui.btn_FilterSupplier, "все Поставщики", self._selected_supplier_ids)
        self._set_filter_button_text(self.ui.btn_FilterCountry, "все Страны", self._selected_country_values)

    @staticmethod
    def _set_filter_button_text(button, all_text: str, selected: set | None):
        button.setText(all_text if selected is None else f"{all_text} ({len(selected)})")

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
                    "insurance_percent": row.insurance_percent,
                    "fx_rate_markup": row.fx_rate_markup,
                    "fx_rate_markup_abs": row.fx_rate_markup_abs,
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

        if self._selected_supplier_ids is not None:
            supplier_data = [row for row in supplier_data if int(row["id"]) in self._selected_supplier_ids]

        if self._selected_country_values is not None:
            supplier_data = [
                row for row in supplier_data if (row["country"] or "") in self._selected_country_values
            ]

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

                if col_name in {"reexport_percent", "insurance_percent", "fx_rate_markup"}:
                    text_value = self._format_percent(value)
                else:
                    text_value = "" if value is None else str(value).replace(".", ",")

                item = self._build_table_item(col_name, text_value)
                self._original_values[row_id][col_name] = text_value
                self.table.setItem(row_index, col_index, item)

        resize_columns_for_multiline_headers(self.table)

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
        item_class = NumericTableWidgetItem if col_name == "id" else QTableWidgetItem
        item = item_class(value)

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
            "insurance_percent": "0",
            "fx_rate_markup": "0",
            "fx_rate_markup_abs": "0",
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
            "insurance_percent": "0",
            "fx_rate_markup": "0",
            "fx_rate_markup_abs": "0",
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
        return self._selected_supplier_ids is not None or self._selected_country_values is not None

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
