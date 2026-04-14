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
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import FixedCosts
from app.db.db import SessionLocal


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
            "move_novo_tamozh",
            "move_tamozh_chekhov",
        ]
        self.headers = [
            "id",
            "Customs clearance",
            "Additional customs",
            "Excise",
            "Eco fee",
            "VAT",
            "Customs fee",
            "Bank fee",
            "Money",
            "Storage",
            "Move Novo-Tamozh",
            "Move Tamozh-Chekhov",
        ]
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.find_fixed_costs()

    def setup_ui(self):
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(False)
        self.table.setWordWrap(False)
        self.table.setTextElideMode(Qt.TextElideMode.ElideRight)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f0f0;
                selection-background-color: #3daee9;
                selection-color: black;
            }
            QTableWidget::item {
                padding: 3px;
            }
            QTableWidget::item:editable {
                background-color: #fff5cc;
                border: 1px solid #ffcc66;
            }
            QTableWidget::item:focus {
                background-color: #f28223;
                border: 1px solid #ff9900;
                padding: 1px;
            }
            QTableWidget QLineEdit {
                background-color: #fff2cc;
                color: black;
                border: 1px solid #ff9900;
                padding: 1px;
            }
        """)

        self.table.setTabKeyNavigation(True)
        self.table.setCornerButtonEnabled(False)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

    def get_session(self):
        return SessionLocal()

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

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        apply_action = menu.addAction("Применить изменения")
        revert_action = menu.addAction("Отменить изменения")

        copy_action.triggered.connect(self.copy_cell_content)
        apply_action.triggered.connect(self.apply_pending_changes)
        revert_action.triggered.connect(self.revert_changes)

        menu.exec_(self.table.viewport().mapToGlobal(position))

    def on_item_changed(self, item):
        if self._updating_table:
            return

        try:
            header_item = self.table.horizontalHeaderItem(item.column())
            if not header_item:
                return

            column_name = self.header_to_column.get(header_item.text())
            if not column_name or column_name == "id":
                return

            self._pending_changes[column_name] = item.text()

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
                        move_novo_tamozh=0,
                        move_tamozh_chekhov=0,
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
                    move_novo_tamozh=0,
                    move_tamozh_chekhov=0,
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
                "move_novo_tamozh": row.move_novo_tamozh,
                "move_tamozh_chekhov": row.move_tamozh_chekhov,
            }

    def find_fixed_costs(self):
        data = self.get_fixed_costs_from_db()
        self._display_data(data)

    def _build_item(self, value, editable=True, align_left=False):
        item = QTableWidgetItem("" if value is None else str(value))
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def _display_data(self, row_data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        self._updating_table = True
        self._row_id = row_data["id"]

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(1)
        self.table.setHorizontalHeaderLabels(self.headers)

        for col_index, col_name in enumerate(self.columns):
            editable = col_name != "id"
            item = self._build_item(row_data[col_name], editable=editable, align_left=False)
            self.table.setItem(0, col_index, item)

        self.table.resizeColumnsToContents()

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setStyleSheet("""
            QLabel {
                background-color: #CCFF99;
                color: #12501A;
                border: 2px solid #12501A;
                border-radius: 5px;
                padding: 8px;
                font: 10pt "Tahoma";
                margin: 2px;
            }
        """)
        self.ui.label_msg.setVisible(True)

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