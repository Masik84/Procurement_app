from __future__ import annotations

"""Project-wide Excel formatting rule for all uC3 columns.

Any visible Excel header containing uC3 (case-insensitive, including variants
with spaces, line breaks, suffixes such as _1/_2, "uC3 PY",
"uC3 3 mnth", "min uC3 stock", etc.) is written/formatted as RUB money:

    # ##0 ₽;[Red]-# ##0 ₽;"-"

The rule is deliberately installed centrally so new uC3 report columns inherit
the same format automatically.
"""

import importlib
import re
import sys
from typing import Any


_INSTALLED = False
_IMPORT_MODULE_BEFORE_UC3_PATCH = importlib.import_module


def _is_uc3_header(value: object) -> bool:
    text = str(value or "").casefold()
    # Be tolerant of spaces/newlines/punctuation and even accidental Cyrillic "с".
    text = text.replace("с", "c")
    compact = re.sub(r"[^a-z0-9]+", "", text)
    return "uc3" in compact


def _patch_excel_column_format(module) -> None:
    if getattr(module, "_UC3_MONEY_PATCHED", False):
        return
    module._UC3_MONEY_PATCHED = True

    original_number_format = module.number_format_for_header

    def number_format_for_header_uc3(header: str):
        if _is_uc3_header(header):
            return module.FORMATS.MONEY_RUB
        return original_number_format(header)

    module.number_format_for_header = number_format_for_header_uc3

    original_excel_value = module.excel_value_by_header

    def excel_value_by_header_uc3(header: str, value: object) -> Any:
        if _is_uc3_header(header):
            if value is None or value == "":
                return ""
            number = module.parse_excel_number(value)
            return "" if number is None else number
        return original_excel_value(header, value)

    module.excel_value_by_header = excel_value_by_header_uc3


def _patch_excel_export_format(module) -> None:
    if getattr(module, "_UC3_MONEY_PATCHED", False):
        return
    module._UC3_MONEY_PATCHED = True

    original_is_money_header = module.is_money_header

    def is_money_header_uc3(header: str) -> bool:
        if _is_uc3_header(header):
            return True
        return original_is_money_header(header)

    module.is_money_header = is_money_header_uc3


def _patch_output_headers(module) -> None:
    if getattr(module, "_UC3_MONEY_PATCHED", False):
        return
    module._UC3_MONEY_PATCHED = True

    original_excel_spec = module.excel_spec

    def excel_spec_uc3(header: object):
        spec = original_excel_spec(header)
        if not _is_uc3_header(header):
            return spec

        if spec is not None:
            fill, font, _number_format = spec
        else:
            fill, font = (205, 205, 205), None
        return (fill, font, module.FORMATS.MONEY_RUB)

    module.excel_spec = excel_spec_uc3


def _patch_supplier_price_exporter(module) -> None:
    """Override the report-specific DECIMAL_2 formatting applied to uC3 columns."""
    exporter_cls = getattr(module, "SupplierPriceExporter", None)
    if exporter_cls is None:
        return

    original = exporter_cls._apply_calculated_export_formats_by_header
    if getattr(original, "_uc3_money_patch", False):
        return

    def apply_calculated_export_formats_uc3(self, ws, headers):
        original(self, ws, headers)

        header_map = self._header_map(headers)
        for header in headers:
            if _is_uc3_header(header):
                self._set_format_by_header(
                    ws,
                    header_map,
                    header,
                    module.FORMATS.MONEY_RUB,
                )

    apply_calculated_export_formats_uc3._uc3_money_patch = True
    exporter_cls._apply_calculated_export_formats_by_header = (
        apply_calculated_export_formats_uc3
    )


def _patch_loaded_modules() -> None:
    module = sys.modules.get("app.exports.excel_column_format")
    if module is not None:
        _patch_excel_column_format(module)

    module = sys.modules.get("app.utils.excel_export_format")
    if module is not None:
        _patch_excel_export_format(module)

    module = sys.modules.get("app.utils.output_headers")
    if module is not None:
        _patch_output_headers(module)

    module = sys.modules.get("app.exports.supplier_price_exporter")
    if module is not None:
        _patch_supplier_price_exporter(module)


def _patched_import_module(name: str, package: str | None = None):
    module = _IMPORT_MODULE_BEFORE_UC3_PATCH(name, package)
    _patch_loaded_modules()
    return module


def install_uc3_money_format() -> None:
    """Install the uC3 Excel rule once for the entire application."""
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True

    # Patch the shared formatting modules immediately. They are lightweight and
    # doing this here also covers direct exporter use outside the GUI pages.
    for module_name in (
        "app.exports.excel_column_format",
        "app.utils.excel_export_format",
        "app.utils.output_headers",
    ):
        try:
            _IMPORT_MODULE_BEFORE_UC3_PATCH(module_name)
        except Exception:
            # The normal page import will give another chance to patch it.
            pass

    _patch_loaded_modules()

    # The application loads heavy pages/exporters lazily through importlib.
    # Wrap the already-installed no-IS import hook rather than replacing it.
    importlib.import_module = _patched_import_module
