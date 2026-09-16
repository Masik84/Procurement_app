from __future__ import annotations

"""Project-wide uC3 Excel rule + central Excel presentation hook.

The original purpose of this module is to guarantee integer display for every
uC3 header. It also provides the existing import hook used to attach the final
central Excel rules to legacy exporters that save workbooks directly instead
of going through save_workbook_xlsx().
"""

import builtins
import importlib
import re
import sys
from typing import Any


_INSTALLED = False
_IMPORT_MODULE_BEFORE_UC3_PATCH = importlib.import_module
_IMPORT_BEFORE_UC3_PATCH = builtins.__import__


def _is_uc3_header(value: object) -> bool:
    text = str(value or "").casefold().replace("с", "c")
    compact = re.sub(r"[^a-z0-9]+", "", text)
    return "uc3" in compact


def _patch_excel_column_format(module) -> None:
    if getattr(module, "_UC3_INTEGER_PATCHED", False):
        return
    module._UC3_INTEGER_PATCHED = True

    original_number_format = module.number_format_for_header

    def number_format_for_header_uc3(header: str):
        if _is_uc3_header(header):
            return module.FORMATS.INTEGER
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
    if getattr(module, "_UC3_INTEGER_PATCHED", False):
        return
    module._UC3_INTEGER_PATCHED = True

    original_column_format_local = module.column_format_local

    def column_format_local_uc3(header: str) -> str:
        if _is_uc3_header(header):
            return module.FORMATS.INTEGER
        return original_column_format_local(header)

    module.column_format_local = column_format_local_uc3

    original_openpyxl_number_format = module.openpyxl_number_format

    def openpyxl_number_format_uc3(header: str) -> str:
        if _is_uc3_header(header):
            return module.FORMATS.INTEGER
        return original_openpyxl_number_format(header)

    module.openpyxl_number_format = openpyxl_number_format_uc3


def _patch_output_headers(module) -> None:
    if getattr(module, "_UC3_INTEGER_PATCHED", False):
        return
    module._UC3_INTEGER_PATCHED = True

    original_excel_spec = module.excel_spec

    def excel_spec_uc3(header: object):
        spec = original_excel_spec(header)
        if not _is_uc3_header(header):
            return spec
        if spec is not None:
            fill, font, _number_format = spec
        else:
            fill, font = (205, 205, 205), None
        return (fill, font, module.FORMATS.INTEGER)

    module.excel_spec = excel_spec_uc3


def _postprocess_saved_workbook(path) -> None:
    if not path:
        return
    try:
        from pathlib import Path
        import pythoncom
        import win32com.client as win32
        from app.utils.excel_format_rules import apply_central_workbook_rules

        target = Path(path)
        if not target.exists():
            return
        pythoncom.CoInitialize()
        excel = wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(str(target.resolve()))
            apply_central_workbook_rules(wb)
            wb.Save()
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
    except Exception:
        # Formatting post-processing must never make an otherwise valid export
        # fail. The normal exporter result remains usable and the traceback is
        # visible in app.log/console.
        import logging
        logging.getLogger(__name__).exception("Не удалось применить финальные Excel-правила к %s", path)


def _patch_supplier_price_exporter(module) -> None:
    exporter_cls = getattr(module, "SupplierPriceExporter", None)
    if exporter_cls is None or getattr(exporter_cls, "_CENTRAL_EXCEL_RULES_PATCHED", False):
        return
    # Set the guard before any imports below. builtins.__import__ is itself
    # patched by this module, so an unguarded import here can re-enter
    # _patch_loaded_modules() recursively.
    exporter_cls._CENTRAL_EXCEL_RULES_PATCHED = True

    # Legacy exporter methods may still pass old literal widths/formats. Make
    # the central rule authoritative for every registered standard header.
    from app.utils.excel_format_rules import number_format_for_header, width_for_header

    set_width = getattr(exporter_cls, "_set_width_by_header", None)
    if set_width is not None and not getattr(set_width, "_central_excel_rules_patch", False):
        def set_width_central(self, ws, header_map, header, width):
            resolved = width_for_header(header, width)
            return set_width(self, ws, header_map, header, resolved)
        set_width_central._central_excel_rules_patch = True
        exporter_cls._set_width_by_header = set_width_central

    set_format = getattr(exporter_cls, "_set_format_by_header", None)
    if set_format is not None and not getattr(set_format, "_central_excel_rules_patch", False):
        def set_format_central(self, ws, header_map, header, format_local):
            resolved = number_format_for_header(header) or format_local
            return set_format(self, ws, header_map, header, resolved)
        set_format_central._central_excel_rules_patch = True
        exporter_cls._set_format_by_header = set_format_central

    original = exporter_cls._apply_calculated_export_formats_by_header
    if getattr(original, "_uc3_integer_patch", False):
        return

    def apply_calculated_export_formats_uc3(self, ws, headers):
        original(self, ws, headers)

        # Preserve the long-standing rule that every uC3 variant is displayed
        # as an integer, including dynamically suffixed headers.
        header_map = self._header_map(headers)
        for header in headers:
            if _is_uc3_header(header):
                self._set_format_by_header(
                    ws,
                    header_map,
                    header,
                    module.FORMATS.INTEGER,
                )

        # This exporter saves CostCalc_ directly via Workbook.SaveAs(), so it
        # must explicitly receive the same final central rules as exporters
        # that call save_workbook_xlsx().
        from app.utils.excel_format_rules import apply_central_worksheet_rules
        apply_central_worksheet_rules(ws)

    apply_calculated_export_formats_uc3._uc3_integer_patch = True
    exporter_cls._apply_calculated_export_formats_by_header = apply_calculated_export_formats_uc3

    # Template/warning exports bypass save_workbook_xlsx(). Reopen the small
    # generated workbook once so standard columns still receive central widths
    # and formats instead of legacy local constants.
    for method_name in ("export_template", "export_qty_in_box_warnings"):
        original_export = getattr(exporter_cls, method_name, None)
        if original_export is None or getattr(original_export, "_central_excel_postprocess", False):
            continue

        def make_export_wrapper(original_method):
            def wrapper(self, *args, **kwargs):
                result = original_method(self, *args, **kwargs)
                _postprocess_saved_workbook(result)
                return result
            wrapper._central_excel_postprocess = True
            return wrapper

        setattr(exporter_cls, method_name, make_export_wrapper(original_export))


def _patch_price_report_exporter(module) -> None:
    exporter_cls = getattr(module, "PriceReportExporter", None)
    if exporter_cls is None or getattr(exporter_cls, "_CENTRAL_EXCEL_RULES_PATCHED", False):
        return
    exporter_cls._CENTRAL_EXCEL_RULES_PATCHED = True

    from app.utils.excel_format_rules import (
        apply_central_worksheet_rules,
        number_format_for_header,
        width_for_header,
    )

    set_width = getattr(exporter_cls, "_set_width_by_header", None)
    if set_width is not None:
        def set_width_central(self, ws, header_map, header, width):
            return set_width(self, ws, header_map, header, width_for_header(header, width))
        exporter_cls._set_width_by_header = set_width_central

    format_columns = getattr(exporter_cls, "_format_columns_by_headers", None)
    if format_columns is not None:
        def format_columns_central(self, ws, header_map, headers, format_local):
            for header in headers:
                format_columns(
                    self, ws, header_map, [header],
                    number_format_for_header(header) or format_local,
                )
        exporter_cls._format_columns_by_headers = format_columns_central

    for method_name in ("_format_product_report", "_format_supplier_report"):
        original = getattr(exporter_cls, method_name, None)
        if original is None:
            continue

        def make_wrapper(original_method):
            def wrapper(self, ws, headers, rows_count):
                result = original_method(self, ws, headers, rows_count)
                apply_central_worksheet_rules(ws)
                return result
            wrapper._central_excel_rules_patch = True
            return wrapper

        setattr(exporter_cls, method_name, make_wrapper(original))


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

    module = sys.modules.get("app.exports.price_report_exporter")
    if module is not None:
        _patch_price_report_exporter(module)


def _patched_import_module(name: str, package: str | None = None):
    module = _IMPORT_MODULE_BEFORE_UC3_PATCH(name, package)
    _patch_loaded_modules()
    return module


def _patched_import(name, globals=None, locals=None, fromlist=(), level=0):
    module = _IMPORT_BEFORE_UC3_PATCH(name, globals, locals, fromlist, level)
    _patch_loaded_modules()
    return module


def install_uc3_integer_format() -> None:
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True

    for module_name in (
        "app.exports.excel_column_format",
        "app.utils.excel_export_format",
        "app.utils.output_headers",
    ):
        try:
            _IMPORT_MODULE_BEFORE_UC3_PATCH(module_name)
        except Exception:
            pass

    _patch_loaded_modules()

    importlib.import_module = _patched_import_module
    builtins.__import__ = _patched_import
