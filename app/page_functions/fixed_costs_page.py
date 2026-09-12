from pathlib import Path
from decimal import Decimal, InvalidOperation

from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QHeaderView,
    QTableWidget,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QVBoxLayout,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import FixedCosts
from app.db.db import SessionLocal
from app.ui.table_style import *
from app.utils.output_headers import display_headers
from app.utils.money import parse_decimal_field
from app.utils.parsers import parse_user_percent
from app.utils.message_dialogs import show_error


BASE_DIR = Path(__file__).resolve().parents[2]
FIXED_COSTS_UI = BASE_DIR / "app" / "ui" / "windows" / "fixed_costs.ui"


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


class FixedCostsPage(QWidget):
    # These values are stored as fractions and used as multipliers in the cost
    # formulas (1 + value). The remaining fixed-cost fields are RUB/L amounts.
    PERCENT_COLUMNS = {"customs_clearance", "vat", "bank_fee", "money"}

    def __init__(self):
        super().__init__()

        self.ui = load_ui(FIXED_COSTS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table

        self._updating_table = False
        self._pending_changes = {}
        self._row_id = None

        self.columns = [
            "id",
            "customs_clearance",
            "additional_customs",
            "excise",
            "eco_fee",
            "vat",
            "customs_fee",
            "bank_fee",
            "money",
            "storage",
            "move",
        ]
        self.headers = [
            "id",
            "Customs clearance %",
            "Additional customs",
            "Excise",
            "Eco fee",
            "VAT %",
            "Customs fee",
            "Bank fee %",
            "Money %",
            "Storage",
            "Move",
        ]
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.find_fixed_costs()

    def setup_ui(self):
        setup_data_table(self.table, sorting=False)
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

    def get_session(self):
        return SessionLocal()

    def _to_decimal(self, value, field_name):
        if field_name in self.PERCENT_COLUMNS:
            parsed = parse_user_percent(value)
            if parsed is None:
                raise Exception(f"Поле '{field_name}' должно быть числом")
            return parsed

        # Only called from apply_pending_changes (final Save), never live per
        # keystroke, so it's safe to reject an empty NOT NULL field right away
        # instead of silently saving it as 0.
        return parse_decimal_field(value, field_name, empty="raise")

    def on_item_changed(self, item):
        if self._updating_table:
            return

        try:
            header_item = self.table.horizontalHeaderItem(item.column())
            if not header_item:
                return

            column_name = self.header_to_column.get(table_header_name(self.table, item.column()))
            if not column_name or column_name == "id":
                return

            self._pending_changes[column_name] = item.text()

        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self.find_fixed_costs()
            self.show_message("Изменения отменены")
        except Exception as e:
            self.show_error_message(f"Ошибка отката: {str(e)}")

    def apply_pending_changes(self):
        if not self._pending_changes:
            self.show_message("Нет изменений для применения")
            return

        try:
            with self.get_session() as session:
                row = session.query(FixedCosts).first()

                if not row:
                    row = FixedCosts(
                        customs_clearance=0,
                        additional_customs=0,
                        excise=0,
                        eco_fee=0,
                        vat=0,
                        customs_fee=0,
                        bank_fee=0,
                        money=0,
                        storage=0,
                        move=0,
                    )
                    session.add(row)
                    session.flush()

                for column_name, value in self._pending_changes.items():
                    setattr(row, column_name, self._to_decimal(value, column_name))

                session.commit()

            self._pending_changes.clear()
            self.find_fixed_costs()
            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def get_fixed_costs_from_db(self):
        with self.get_session() as session:
            row = session.query(FixedCosts).first()

            if not row:
                row = FixedCosts(
                    customs_clearance=0,
                    additional_customs=0,
                    excise=0,
                    eco_fee=0,
                    vat=0,
                    customs_fee=0,
                    bank_fee=0,
                    money=0,
                    storage=0,
                    move=0,
                )
                session.add(row)
                session.commit()
                session.refresh(row)

            return {
                "id": row.id,
                "customs_clearance": row.customs_clearance,
                "additional_customs": row.additional_customs,
                "excise": row.excise,
                "eco_fee": row.eco_fee,
                "vat": row.vat,
                "customs_fee": row.customs_fee,
                "bank_fee": row.bank_fee,
                "money": row.money,
                "storage": row.storage,
                "move": row.move,
            }

    def find_fixed_costs(self):
        data = self.get_fixed_costs_from_db()
        self._display_data(data)

    def _build_item(self, value, editable=True, align_left=False, numeric_sort=False):
        item_class = NumericTableWidgetItem if numeric_sort else QTableWidgetItem
        item = item_class(format_table_value(value))
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    @staticmethod
    def _format_percent(value) -> str:
        if value is None:
            return ""
        text = f"{float(value) * 100:.2f}".rstrip("0").rstrip(".").replace(".", ",")
        return f"{text}%"

    def _display_data(self, row_data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        self._updating_table = True
        self._row_id = row_data["id"]

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(1)
        self.table.setHorizontalHeaderLabels(display_headers(self.headers))

        for col_index, col_name in enumerate(self.columns):
            editable = col_name != "id"
            value = row_data[col_name]
            if col_name in self.PERCENT_COLUMNS:
                value = self._format_percent(value)
            item = self._build_item(
                value,
                editable=editable,
                align_left=False,
                numeric_sort=col_name == "id",
            )
            self.table.setItem(0, col_index, item)

        resize_columns_for_multiline_headers(self.table)

        self._updating_table = False

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
        show_error(self, text)
