from __future__ import annotations

"""Единые правила форматирования Excel.

Все рабочие форматы хранятся в локальном виде, как в старых рабочих выгрузках:
- дата: ДД.ММ.ГГ;@
- Price, L / Price, pack: # ##0,00_ ;[Red]-# ##0,00_ ;'-'
- рубли: # ##0 ₽
- FX rate: # ##0

Сначала используется NumberFormatLocal. Если локальная маска не принимается,
внутри helper строится отдельный invariant-вариант для NumberFormat.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ExcelFormats:
    GENERAL: str = "General"
    TEXT: str = "@"
    DATE: str = "ДД.ММ.ГГ;@"
    INTEGER: str = '# ##0;[Red]-# ##0;"-"'
    DECIMAL_1: str = '# ##0,0;[Red]-# ##0,0;"-"'
    DECIMAL_2: str = '# ##0,00;[Red]-# ##0,00;"-"'
    DECIMAL_2_SIMPLE: str = "0,00"
    DECIMAL_2_PLAIN: str = "# ##0,00"
    PRICE_DECIMAL: str = "# ##0,00_ ;[Red]-# ##0,00_ ;'-'"
    DECIMAL_4: str = '# ##0,0000;[Red]-# ##0,0000;"-"'
    DECIMAL_FLEX: str = '# ##0,00##;[Red]-# ##0,00##;0'
    MONEY_RUB: str = '# ##0 ₽;[Red]-# ##0 ₽;"-"'
    MONEY_RUB_SIMPLE: str = "# ##0 ₽"
    PERCENT_1: str = "0,0%"
    PERCENT_2: str = "0,00%"
    PERCENT_FLEX: str = "0,0#%"
    FX_INTEGER: str = "# ##0"
    FX_FLEX: str = "# ##0,0###"


FORMATS = ExcelFormats()
LOCAL_FORMATS = FORMATS


class ExcelNumberFormatError(RuntimeError):
    """Оставлено для совместимости со старым кодом."""


def cost_calc_headers(
    *,
    quick_order_months: int | None,
    safe_stock_months: int | None,
) -> tuple[list[str], str]:
    """Return the approved CostCalc_ column order.

    The quick-order calculation remains in the exporter, but its visible column
    is temporarily disabled only for CostCalc_. Other reports keep their own
    quick-order columns unchanged.
    """
    standard_order_header = (
        f"к Заказу, л ({safe_stock_months} м)"
        if safe_stock_months is not None
        else "к Заказу, л"
    )
    headers = [
        "Supplier Article",
        "Supplier Product Name",
        "Our Product Name",
        "Pack",
        "Категория ABC",
        "Qty, pcs",
        "Volume, L",
        "Ср.Продажи мес",
        standard_order_header,
        "Volume to take",
        "Price, L",
        "Price, pack",
        "Currency",
        "FX rate",
        "Cost Novo with VAT",
        "Full Cost Msk",
        "uC3",
        "Target price, L",
        "uC3 PY",
        "uC3 3 mnth",
        "last update (prev)",
        "Price, L (prev)",
        "Cost Novo with VAT (prev)",
        "Full Cost Msk (prev)",
        "Дистр цена",
        "Промо цена",
        "curr LPC",
        "curr Landed cost",
        "min uC3 stock",
        "Best Suppl",
        "Best full Price, L",
        "last update Best1",
        "FX rate Best1",
        "Currency Best1",
        "Best Suppl 2",
        "Best full Price, L 2",
        "last update Best2",
        "FX rate Best2",
        "Currency Best2",
        "Volume PY",
        "Volume 3 mnth",
        "Stock",
        "Transit",
        "Purchase Order",
        "Order IS",
        "Stock IS",
        "Reserve cust",
        "Reserve E-Comm",
        "Damaged",
        # Быстрый заказ временно скрыт только в CostCalc_. Расчёт и
        # значение в строках сохранены, чтобы колонку было легко вернуть.
        # f"к Быстрому заказу, л ({quick_order_months} м)"
        # if quick_order_months is not None else "к Быстрому заказу, л",
    ]
    return headers, standard_order_header


def to_invariant_number_format(format_code: str | None) -> str | None:
    if not format_code:
        return format_code
    return (
        str(format_code)
        .replace("ДД", "dd")
        .replace("ММ", "mm")
        .replace("ГГГГ", "yyyy")
        .replace("ГГ", "yy")
        .replace(",0000", ".0000")
        .replace(",00##", ".00##")
        .replace(",00", ".00")
        .replace(",0###", ".0###")
        .replace(",0#", ".0#")
        .replace(",0", ".0")
        .replace("# ##", "#,##")
    )


def to_local_number_format(
    format_code: str | None,
    *,
    decimal_separator: str = ",",
    thousands_separator: str = " ",
) -> str | None:
    return format_code


def set_number_format_safe(
    target: Any,
    format_en: str = FORMATS.GENERAL,
    format_local: str | None = None,
    *,
    verify: bool = False,
) -> str:
    local_code = format_local or format_en or FORMATS.GENERAL

    if format_local is not None and format_en and format_en != local_code:
        invariant_code = format_en
    else:
        invariant_code = to_invariant_number_format(local_code) or FORMATS.GENERAL

    candidates = [
        ("NumberFormatLocal", local_code),
        ("NumberFormat", invariant_code),
        ("NumberFormat", FORMATS.GENERAL),
    ]

    seen = set()
    for attr, fmt in candidates:
        key = (attr, fmt)
        if not fmt or key in seen:
            continue
        seen.add(key)
        try:
            setattr(target, attr, fmt)
            return fmt
        except Exception:
            pass

    return FORMATS.GENERAL


def ensure_xlsx_path(file_path: str | Path) -> Path:
    path = Path(file_path)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    return path


def save_workbook_xlsx(workbook: Any, file_path: str | Path) -> Path:
    target_path = ensure_xlsx_path(file_path)
    workbook.SaveAs(str(target_path))
    return target_path
