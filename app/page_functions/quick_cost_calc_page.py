from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import QFile, QEvent, QPoint, Qt
from PySide6.QtWidgets import QApplication, QMessageBox, QToolTip, QVBoxLayout, QWidget
from PySide6.QtUiTools import QUiLoader

from app.db.db import SessionLocal
from app.db.models import ExchangeRate, MarkingRate, Supplier
from app.services.quick_cost_calculation_service import QuickCostCalculationService
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


BASE_DIR = Path(__file__).resolve().parents[2]
QUICK_COST_CALC_UI = BASE_DIR / "app" / "ui" / "windows" / "quick_prices_calc.ui"


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


class QuickCostCalcPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(QUICK_COST_CALC_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.ui.line_CostNovo.setAlignment(Qt.AlignRight)
        self.ui.line_FullCost.setAlignment(Qt.AlignRight)
        self.ui.line_Price.setAlignment(Qt.AlignRight)

        self.setup_ui()
        self.setup_connections()
        self.load_initial_state()

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        self.ui.line_CostNovo.setReadOnly(True)
        self.ui.line_FullCost.setReadOnly(True)

        self._setup_number_field(self.ui.line_ExchangeRate, "Формат: 82,0000")
        self._setup_number_field(self.ui.line_Transport, "Формат: 1,2500")
        self._setup_number_field(self.ui.line_AgentFee, "Формат: 2,6500")
        self._setup_number_field(self.ui.line_Reexport, "Формат: 3,5% / 0,24%")
        self._setup_number_field(self.ui.line_FXMarkup, "Формат: 3,5% / 0,24%")
        self._setup_number_field(self.ui.line_FXMarkupAbs, "Формат: 1,5000")
        self._setup_number_field(self.ui.line_Price, "Формат: 125,4500")

    def setup_connections(self):
        self.ui.cbo_SupplName.currentIndexChanged.connect(self.on_supplier_changed)
        self.ui.cbo_Currency.currentIndexChanged.connect(self.on_currency_changed)

        self.ui.line_ExchangeRate.editingFinished.connect(self.normalize_exchange_rate)
        self.ui.line_Transport.editingFinished.connect(self.normalize_transport)
        self.ui.line_AgentFee.editingFinished.connect(self.normalize_agent_fee)
        self.ui.line_Reexport.editingFinished.connect(self.normalize_reexport)
        self.ui.line_FXMarkup.editingFinished.connect(self.normalize_fx_markup)
        self.ui.line_FXMarkupAbs.editingFinished.connect(self.normalize_fx_markup_abs)
        self.ui.line_Price.editingFinished.connect(self.normalize_price)

        self.ui.btn_Calc.clicked.connect(self.calculate_costs)
        self.ui.btn_Reset.clicked.connect(self.reset_form)

    def _setup_number_field(self, widget, tooltip_text: str):
        widget.setToolTip(tooltip_text)
        widget.installEventFilter(self)

    def eventFilter(self, watched, event):
        watched_fields = {
            self.ui.line_ExchangeRate,
            self.ui.line_Transport,
            self.ui.line_AgentFee,
            self.ui.line_Reexport,
            self.ui.line_FXMarkup,
            self.ui.line_FXMarkupAbs,
            self.ui.line_Price,
        }
        if watched in watched_fields:
            if event.type() in {QEvent.Enter, QEvent.FocusIn, QEvent.MouseButtonPress}:
                QToolTip.showText(
                    watched.mapToGlobal(QPoint(0, watched.height())),
                    watched.toolTip(),
                    watched,
                )

        return super().eventFilter(watched, event)

    def set_combo_text(self, combo, value: str):
        index = combo.findText(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif combo.isEditable():
            combo.setCurrentText(value)

    def load_initial_state(self):
        self.prepare_static_combos()
        self.load_suppliers()
        self.load_currencies()
        self.load_pack_types()
        self.apply_default_values()

    def prepare_static_combos(self):
        self._prepend_dash_if_missing(self.ui.line_Excise)

    def _prepend_dash_if_missing(self, combo):
        if combo.findText("-") >= 0:
            return
        combo.insertItem(0, "-")

    def apply_default_values(self):
        self.set_combo_text(self.ui.cbo_SupplName, "-")
        self.set_combo_text(self.ui.cbo_Currency, "-")
        self.set_combo_text(self.ui.line_PackType, "-")
        self.set_combo_text(self.ui.line_Excise, "-")

        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
        self.set_combo_text(self.ui.cbo_Customs, "да")
        self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
        self.set_combo_text(self.ui.cbo_Marking, "Феникс")

        self.ui.line_ExchangeRate.clear()
        self.ui.line_Transport.clear()
        self.ui.line_AgentFee.clear()
        self.ui.line_Reexport.setText("0,0%")
        self.ui.line_FXMarkup.setText("0,0%")
        self.ui.line_FXMarkupAbs.setText("0,0000")
        self.ui.line_Price.clear()
        self.ui.line_CostNovo.clear()
        self.ui.line_FullCost.clear()

    def load_suppliers(self):
        current_text = self.ui.cbo_SupplName.currentText().strip()

        with self.get_session() as session:
            rows = (
                session.query(Supplier)
                .filter(Supplier.name != "Manual")
                .order_by(Supplier.name.asc())
                .all()
            )

        self.ui.cbo_SupplName.blockSignals(True)
        self.ui.cbo_SupplName.clear()
        self.ui.cbo_SupplName.addItem("-", None)
        for row in rows:
            self.ui.cbo_SupplName.addItem(row.name, row.id)
        self.ui.cbo_SupplName.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_SupplName, current_text)

    def load_currencies(self):
        current_text = self.ui.cbo_Currency.currentText().strip()

        with self.get_session() as session:
            rows = (
                session.query(ExchangeRate.currency_code)
                .order_by(ExchangeRate.currency_code.asc())
                .all()
            )

        self.ui.cbo_Currency.blockSignals(True)
        self.ui.cbo_Currency.clear()
        self.ui.cbo_Currency.addItem("-")
        for row in rows:
            if row[0]:
                self.ui.cbo_Currency.addItem(row[0])
        self.ui.cbo_Currency.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_Currency, current_text)

    def load_pack_types(self):
        current_text = self.ui.line_PackType.currentText().strip()

        with self.get_session() as session:
            rows = (
                session.query(MarkingRate.pack_type)
                .filter(MarkingRate.pack_type.isnot(None), MarkingRate.pack_type != "")
                .order_by(MarkingRate.pack_type.asc())
                .all()
            )

        self.ui.line_PackType.blockSignals(True)
        self.ui.line_PackType.clear()
        self.ui.line_PackType.addItem("-")
        for row in rows:
            self.ui.line_PackType.addItem(row[0])
        self.ui.line_PackType.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.line_PackType, current_text)

    def on_supplier_changed(self):
        supplier_id = self.ui.cbo_SupplName.currentData()

        self.ui.line_CostNovo.clear()
        self.ui.line_FullCost.clear()

        if supplier_id is None:
            self.set_combo_text(self.ui.cbo_Currency, "-")
            self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
            self.set_combo_text(self.ui.cbo_Customs, "да")
            self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
            self.set_combo_text(self.ui.cbo_Marking, "Феникс")
            self.ui.line_ExchangeRate.clear()
            self.ui.line_Transport.clear()
            self.ui.line_AgentFee.clear()
            self.ui.line_Reexport.setText("0,0%")
            self.ui.line_FXMarkup.setText("0,0%")
            self.ui.line_FXMarkupAbs.setText("0,0000")
            return

        with self.get_session() as session:
            supplier = (
                session.query(Supplier)
                .filter(Supplier.id == int(supplier_id))
                .first()
            )
            if supplier is None:
                self.show_error_message("Поставщик не найден")
                return

            rate_row = (
                session.query(ExchangeRate)
                .filter(ExchangeRate.currency_code == supplier.base_currency)
                .first()
            )

        self.set_combo_text(self.ui.cbo_Currency, supplier.base_currency or "-")
        self.ui.line_ExchangeRate.setText(
            self.format_number(rate_row.rate_to_rub, 4) if rate_row is not None else ""
        )
        self.ui.line_Transport.setText(self.format_number(supplier.transport_cost_per_l, 4))
        self.ui.line_AgentFee.setText(self.format_number(getattr(supplier, "agent_fee", None), 4))
        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново" if supplier.is_via_novo else "в Мск")
        self.ui.line_Reexport.setText(self.format_percent(supplier.reexport_percent))
        self.ui.line_FXMarkup.setText(self.format_percent(supplier.fx_rate_markup))
        self.ui.line_FXMarkupAbs.setText(self.format_number(supplier.fx_rate_markup_abs, 4))
        self.set_combo_text(self.ui.cbo_Customs, "да" if supplier.has_import_duty else "нет")
        self.set_combo_text(self.ui.cbo_SupplierRF, "да" if supplier.is_rf else "нет")
        self.set_combo_text(self.ui.cbo_Marking, "Поставщик" if supplier.marks_for_us else "Феникс")
        self.show_message("Поставщик загружен")

    def on_currency_changed(self):
        currency_code = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
        if not currency_code or currency_code == "-":
            return

        with self.get_session() as session:
            row = (
                session.query(ExchangeRate)
                .filter(ExchangeRate.currency_code == currency_code)
                .first()
            )

        if row is not None:
            self.ui.line_ExchangeRate.setText(self.format_number(row.rate_to_rub, 4))

    def parse_decimal_field(self, widget, field_name: str, allow_zero: bool = True) -> Decimal:
        text = clean_multi_spaces(widget.text())
        if not text:
            raise ValueError(f"Заполни поле '{field_name}'")

        value = parse_loose_number(text)
        if value is None:
            raise ValueError(f"Поле '{field_name}' должно быть числом")

        if not allow_zero and value == Decimal("0"):
            raise ValueError(f"Поле '{field_name}' не может быть равно 0")

        return Decimal(str(value))

    def parse_percent_field(self, widget, field_name: str) -> Decimal:
        text = clean_multi_spaces(widget.text())
        if not text:
            return Decimal("0")

        cleaned = text.replace("%", "")
        value = parse_loose_number(cleaned)
        if value is None:
            raise ValueError(f"Поле '{field_name}' должно быть числом или процентом")

        decimal_value = Decimal(str(value))
        if abs(decimal_value) > Decimal("1"):
            decimal_value = decimal_value / Decimal("100")
        return decimal_value

    def format_number(self, value: object, digits: int = 4) -> str:
        number = parse_loose_number(value)
        if number is None:
            return ""
        return f"{float(number):,.{digits}f}".replace(",", "_").replace(".", ",").replace("_", "")

    def format_percent(self, value: object) -> str:
        number = parse_loose_number(value)
        if number is None:
            return "0,0%"
        return f"{float(number) * 100:.1f}".replace(".", ",") + "%"

    def normalize_exchange_rate(self):
        self._normalize_number_widget(self.ui.line_ExchangeRate, digits=4)

    def normalize_transport(self):
        self._normalize_number_widget(self.ui.line_Transport, digits=4)

    def normalize_agent_fee(self):
        self._normalize_number_widget(self.ui.line_AgentFee, digits=4)

    def normalize_price(self):
        self._normalize_number_widget(self.ui.line_Price, digits=4)

    def normalize_reexport(self):
        self._normalize_percent_widget(self.ui.line_Reexport)

    def normalize_fx_markup(self):
        self._normalize_percent_widget(self.ui.line_FXMarkup)

    def normalize_fx_markup_abs(self):
        self._normalize_number_widget(self.ui.line_FXMarkupAbs, digits=4)

    def _normalize_number_widget(self, widget, digits: int):
        text = clean_multi_spaces(widget.text())
        if not text:
            return
        value = parse_loose_number(text)
        if value is None:
            self.show_error_message("Проверь число")
            return
        widget.setText(self.format_number(value, digits))

    def _normalize_percent_widget(self, widget):
        text = clean_multi_spaces(widget.text())
        if not text:
            widget.setText("0,0%")
            return
        cleaned = text.replace("%", "")
        value = parse_loose_number(cleaned)
        if value is None:
            self.show_error_message("Проверь процент")
            return
        decimal_value = Decimal(str(value))
        if abs(decimal_value) > Decimal("1"):
            decimal_value = decimal_value / Decimal("100")
        widget.setText(self.format_percent(decimal_value))

    def validate_required_combos(self):
        required_fields = [
            (self.ui.cbo_SupplName.currentText(), "Поставщик"),
            (self.ui.cbo_Currency.currentText(), "Валюта"),
            (self.ui.line_PackType.currentText(), "Тип упаковки"),
            (self.ui.line_Excise.currentText(), "Акциз"),
            (self.ui.cbo_viaNovo.currentText(), "Путь"),
            (self.ui.cbo_Customs.currentText(), "Пошлина"),
            (self.ui.cbo_SupplierRF.currentText(), "Поставщик РФ"),
            (self.ui.cbo_Marking.currentText(), "Маркировка"),
        ]

        for value, caption in required_fields:
            if clean_multi_spaces(value) in {"", "-"}:
                raise ValueError(f"Заполни поле '{caption}'")

    def calculate_costs(self):
        try:
            self.validate_required_combos()

            supplier_id = self.ui.cbo_SupplName.currentData()
            if supplier_id is None:
                raise ValueError("Выбери поставщика")

            supplier_price = self.parse_decimal_field(self.ui.line_Price, "Цена поставщика", allow_zero=False)
            fx_rate = self.parse_decimal_field(self.ui.line_ExchangeRate, "Курс", allow_zero=False)
            transport = self.parse_decimal_field(self.ui.line_Transport, "Транспорт")
            agent_fee = self.parse_decimal_field(self.ui.line_AgentFee, "Agent fee")
            reexport = self.parse_percent_field(self.ui.line_Reexport, "Реэкспорт")
            fx_markup = self.parse_percent_field(self.ui.line_FXMarkup, "FX markup %")
            fx_markup_abs = self.parse_decimal_field(self.ui.line_FXMarkupAbs, "FX markup abs")

            pack_type_name = clean_multi_spaces(self.ui.line_PackType.currentText())
            has_customs = self.ui.cbo_Customs.currentText() == "да"
            via_novo = self.ui.cbo_viaNovo.currentText() == "через Ново"
            supplier_is_rf = self.ui.cbo_SupplierRF.currentText() == "да"
            marks_for_us = self.ui.cbo_Marking.currentText() == "Поставщик"
            is_excise = self.ui.line_Excise.currentText() == "да"

            with self.get_session() as session:
                service = QuickCostCalculationService(session)
                result = service.calculate(
                    supplier_price=supplier_price,
                    supplier_id=int(supplier_id),
                    pack_type_name=pack_type_name,
                    fx_rate=fx_rate,
                    transport=transport,
                    agent_fee=agent_fee,
                    reexport=reexport,
                    fx_markup=fx_markup,
                    fx_markup_abs=fx_markup_abs,
                    has_customs=has_customs,
                    via_novo=via_novo,
                    supplier_is_rf=supplier_is_rf,
                    marks_for_us=marks_for_us,
                    is_excise=is_excise,
                )

            self.ui.line_CostNovo.setText(self.format_number(result.cost_novo_wvat, 4))
            self.ui.line_FullCost.setText(self.format_number(result.full_cost_msk, 4))
            self.show_message("Расчет выполнен")
        except Exception as e:
            self.ui.line_CostNovo.clear()
            self.ui.line_FullCost.clear()
            self.show_error_message(str(e))

    def reset_form(self):
        self.apply_default_values()
        self.show_message("Форма очищена")

    def show_message(self, text: str):
        label = getattr(self.ui, "label_msg", None)
        if label is not None:
            label.setText(text)
            label.setProperty("active", True)
            label.style().unpolish(label)
            label.style().polish(label)
            label.setVisible(True)
            return

        anchor = getattr(self.ui, "btn_Calc", self)
        QToolTip.showText(
            anchor.mapToGlobal(QPoint(0, anchor.height())),
            text,
            anchor,
        )

    def show_error_message(self, text: str):
        msg = QMessageBox(self)
        msg.setWindowTitle("Ошибка")
        msg.setIcon(QMessageBox.Critical)
        msg.setMinimumSize(700, 400)

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
        msg.exec()

