from __future__ import annotations

"""Single source of truth for Excel number formats used by every exporter.

Excel COM exposes two locale-sensitive properties:

* ``NumberFormat`` is documented as invariant, but some localized COM builds
  still interpret its separators using the Excel locale;
* ``NumberFormatLocal`` expects the current Excel separators and UI locale.

All application rules below are stored in invariant form. The helper reads the
actual separators from the running Excel instance, creates a local mask,
applies it and verifies the value that Excel retained. This prevents a
successful-looking COM call from silently leaving a column in ``General`` or
interpreting a thousands separator as a decimal separator.
"""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ExcelFormats:
    GENERAL: str = "General"
    TEXT: str = "@"
    DATE: str = "dd.mm.yy;@"
    INTEGER: str = r'#,##0;[Red]\-#,##0;"-"'
    DECIMAL_1: str = r'#,##0.0;[Red]\-#,##0.0;"-"'
    DECIMAL_2: str = r'#,##0.00;[Red]\-#,##0.00;"-"'
    DECIMAL_2_SIMPLE: str = "0.00"
    DECIMAL_2_PLAIN: str = "#,##0.00"
    PRICE_DECIMAL: str = r'#,##0.00_ ;[Red]\-#,##0.00_ ;"-"'
    DECIMAL_4: str = r'#,##0.0000;[Red]\-#,##0.0000;"-"'
    DECIMAL_FLEX: str = r'#,##0.00##;[Red]\-#,##0.00##;0'
    MONEY_RUB: str = r'#,##0 "₽";[Red]\-#,##0 "₽";"-"'
    MONEY_RUB_SIMPLE: str = '#,##0 "₽"'
    PERCENT_1: str = "0.0%"
    PERCENT_2: str = "0.00%"
    PERCENT_FLEX: str = "0.0#%"
    FX_INTEGER: str = "#,##0"
    FX_FLEX: str = "#,##0.0###"


FORMATS = ExcelFormats()


class ExcelNumberFormatError(RuntimeError):
    """Raised when Excel does not retain a requested number format."""


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
    """Convert legacy Russian-local masks to Excel's invariant syntax.

    The conversion intentionally targets only tokens used by this project.  A
    format that is already invariant is returned unchanged.
    """
    if format_code is None:
        return None

    code = str(format_code).strip()
    if not code:
        return code

    code = (
        code.replace("ДД", "dd")
        .replace("ММ", "mm")
        .replace("ГГ", "yy")
        .replace("\u00a0", " ")
        .replace(",0000", ".0000")
        .replace(",00##", ".00##")
        .replace(",00", ".00")
        .replace(",0###", ".0###")
        .replace(",0#", ".0#")
        .replace(",0", ".0")
        .replace("# ##", "#,##")
        .replace("[Red]-", r"[Red]\-")
        .replace("'-'", '"-"')
    )

    # Invariant NumberFormat is more reliable when a currency symbol is a
    # quoted literal. Avoid touching an already quoted symbol.
    code = code.replace(' "₽"', " __RUB_QUOTED__")
    code = code.replace(" ₽", ' "₽"')
    return code.replace(" __RUB_QUOTED__", ' "₽"')


def to_local_number_format(
    format_code: str | None,
    *,
    decimal_separator: str = ",",
    thousands_separator: str = " ",
) -> str | None:
    """Build a localized Excel mask from an invariant project format."""
    invariant = to_invariant_number_format(format_code)
    if invariant is None:
        return None
    if invariant in {FORMATS.GENERAL, FORMATS.TEXT}:
        return invariant

    code = invariant
    if "dd" in code or "yy" in code:
        code = code.replace("dd", "ДД").replace("mm", "ММ").replace("yy", "ГГ")

    code = code.replace("#,##", "#__THOUSANDS__##")
    code = (
        code.replace(".0000", f"{decimal_separator}0000")
        .replace(".00##", f"{decimal_separator}00##")
        .replace(".00", f"{decimal_separator}00")
        .replace(".0###", f"{decimal_separator}0###")
        .replace(".0#", f"{decimal_separator}0#")
        .replace(".0", f"{decimal_separator}0")
        .replace("#__THOUSANDS__##", f"#{thousands_separator}##")
        .replace(' "₽"', " ₽")
    )
    return code


def _read_number_format(target: Any) -> str | None:
    value = target.NumberFormat
    if value is None:
        return None
    return str(value).strip()


def _excel_separators(target: Any) -> tuple[str, str] | None:
    try:
        app = target.Application
        # Excel constants: xlDecimalSeparator=3, xlThousandsSeparator=4.
        international = app.International
        try:
            decimal_separator = international(3)
            thousands_separator = international(4)
        except TypeError:
            # Dynamic pywin32 dispatch can expose this indexed COM property as
            # a zero-based tuple instead of a callable method. Excel's
            # xlDecimalSeparator=3 and xlThousandsSeparator=4 constants are
            # therefore found at tuple indexes 2 and 3.
            decimal_separator = international[2]
            thousands_separator = international[3]
        return str(decimal_separator), str(thousands_separator)
    except Exception:
        return None


def _is_general(value: str | None) -> bool:
    return value is None or value.casefold() == FORMATS.GENERAL.casefold()


def set_number_format_safe(
    target: Any,
    format_en: str = FORMATS.GENERAL,
    format_local: str | None = None,
    *,
    verify: bool = True,
) -> str:
    """Apply and verify a number format, preferring invariant syntax.

    ``format_en`` keeps the legacy call signature used by existing exporters.
    If callers pass a Russian-local mask as the first argument, it is converted
    automatically.  A non-General format that Excel does not retain raises an
    explicit error instead of silently degrading the workbook.
    """
    requested = format_local if format_local and format_en == FORMATS.GENERAL else format_en
    requested = requested or format_local or FORMATS.GENERAL
    invariant = to_invariant_number_format(requested) or FORMATS.GENERAL

    separators = _excel_separators(target)
    attempts: list[tuple[str, str]] = []
    if separators is not None:
        decimal_separator, thousands_separator = separators
        generated_local = to_local_number_format(
            invariant,
            decimal_separator=decimal_separator,
            thousands_separator=thousands_separator,
        )
        for local_candidate in (generated_local, format_local, invariant):
            if local_candidate and ("NumberFormatLocal", local_candidate) not in attempts:
                attempts.append(("NumberFormatLocal", local_candidate))

        # NumberFormat is safe as an additional fallback only when Excel uses
        # the separators expected by the invariant mask.
        if decimal_separator == "." and thousands_separator == ",":
            attempts.append(("NumberFormat", invariant))
    else:
        attempts.append(("NumberFormat", invariant))
        local_candidate = format_local or to_local_number_format(invariant)
        if local_candidate:
            attempts.append(("NumberFormatLocal", local_candidate))

    errors: list[str] = []
    expects_general = invariant.casefold() == FORMATS.GENERAL.casefold()

    for attribute, value in attempts:
        try:
            setattr(target, attribute, value)
            if not verify:
                return invariant

            applied = _read_number_format(target)
            if expects_general:
                if _is_general(applied):
                    return applied or FORMATS.GENERAL
            elif not _is_general(applied):
                return applied or invariant

            errors.append(f"{attribute} accepted {value!r}, but Excel retained {applied!r}")
        except Exception as exc:
            errors.append(f"{attribute}={value!r}: {exc}")

    details = "; ".join(errors) if errors else "no format candidates"
    raise ExcelNumberFormatError(
        f"Excel did not apply number format {invariant!r}. {details}"
    )
