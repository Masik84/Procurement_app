from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import shutil
import subprocess
import tempfile
import warnings

import pandas as pd
from openpyxl import load_workbook


_OPENPYXL_DATA_VALIDATION_WARNING = (
    r"Data Validation extension is not supported and will be removed"
)


def _read_excel(path: Path, **kwargs) -> pd.DataFrame:
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
        (str(column).replace("\n", " ").strip() if str(column).strip() else f"Unnamed: {index}")
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
            return _normalize_columns(_read_excel(converted, engine="openpyxl", **kwargs))
        finally:
            shutil.rmtree(converted.parent, ignore_errors=True)
    return _normalize_columns(_read_excel(path, **kwargs))
