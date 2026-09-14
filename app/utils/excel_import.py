from __future__ import annotations

import logging
import re
import shutil
import tempfile
import warnings
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd

from app.utils.text import clean_multi_spaces


logger = logging.getLogger(__name__)

_INTEGER_FLOAT_TEXT_RE = re.compile(r"^[+]?([0-9]+)[.,]0+$")

_OPENPYXL_DATA_VALIDATION_WARNING = (
    r"Data Validation extension is not supported and will be removed"
)


def excel_text(value: Any, *, none_if_empty: bool = False) -> str | None:
    """
    Converts Excel cells that are used as text/IDs/articles to a stable string.

    Excel often stores article/code cells as numbers. pandas/openpyxl may then
    return 149610, 149610.0, Decimal('149610') or the string '149610.0'.
    For matching and saving these values must be treated as the same text key:
    '149610'.

    Notes:
    - If a code was stored in Excel as a real numeric value with leading zeroes,
      the leading zeroes are already lost by Excel before import. They can only
      be preserved when the Excel cell is stored as text.
    - This helper is for import-time text conversion, not fuzzy product matching.
    """
    if value is None:
        return None if none_if_empty else ""

    try:
        if pd.isna(value):
            return None if none_if_empty else ""
    except Exception:
        logger.exception("Подавленная ошибка (см. traceback выше)")

    if isinstance(value, bool):
        text = str(int(value))
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value.is_integer():
            text = str(int(value))
        else:
            text = clean_multi_spaces(value)
    elif isinstance(value, Decimal):
        if value == value.to_integral_value():
            text = str(int(value))
        else:
            text = clean_multi_spaces(value)
    else:
        text = clean_multi_spaces(value)

    if not text or text.lower() == "nan":
        return None if none_if_empty else ""

    match = _INTEGER_FLOAT_TEXT_RE.fullmatch(text)
    if match:
        text = match.group(1)

    text = clean_multi_spaces(text)
    if not text:
        return None if none_if_empty else ""
    return text


def _read_excel(path: Path, **kwargs) -> pd.DataFrame:
    """
    Read an Excel sheet without changing text/code cells more than necessary.

    The defaults below restore the original Procurement import behaviour:
    identifiers/articles stay as object values and blank cells are not
    automatically converted to NaN unless a caller explicitly requests it.
    """
    if not kwargs.get("converters"):
        kwargs.setdefault("dtype", object)
    kwargs.setdefault("keep_default_na", False)

    # Some vendor/customer workbooks contain Excel's extended Data Validation
    # records (x14/x15). openpyxl can still read the cell data, but emits a
    # UserWarning because it does not preserve that extension in its in-memory
    # workbook model. Procurement imports are read-only, so the source workbook
    # is never saved back through openpyxl here and its validations remain
    # untouched. Suppress only this one known warning; all other warnings stay
    # visible in the console/logs.
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=_OPENPYXL_DATA_VALIDATION_WARNING,
            category=UserWarning,
        )
        return pd.read_excel(path, **kwargs)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        (
            str(column).replace("\n", " ").strip()
            if str(column).strip()
            else f"Unnamed: {index}"
        )
        for index, column in enumerate(df.columns)
    ]
    duplicates = defaultdict(int)
    normalized = []
    for column in df.columns:
        count = duplicates[column]
        duplicates[column] += 1
        normalized.append(column if count == 0 else f"{column}.{count}")
    df.columns = normalized
    return df


def _xls_to_xlsx_with_excel(source: Path) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="procurement_xls_"))
    converted = temp_dir / f"{source.stem}.xlsx"

    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(source.resolve()))
        workbook.SaveAs(str(converted.resolve()), FileFormat=51)
        workbook.Close(False)
        workbook = None
        return converted
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def read_excel_raw(file_path: str | Path, **kwargs) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    suffix = path.suffix.lower()
    if suffix == ".xls":
        converted = _xls_to_xlsx_with_excel(path)
        try:
            return _normalize_columns(
                _read_excel(converted, engine="openpyxl", **kwargs)
            )
        finally:
            shutil.rmtree(converted.parent, ignore_errors=True)

    return _normalize_columns(_read_excel(path, **kwargs))
