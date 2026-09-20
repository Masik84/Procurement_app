from __future__ import annotations

"""Safe, application-owned Excel COM lifecycle helpers.

Procurement creates Excel with ``win32com.client.DispatchEx``.  That gives the
application its own Excel instance.  The guard below wraps only those
``DispatchEx('Excel.Application')`` objects; it never attaches to an Excel
instance opened by the user and never enumerates or kills EXCEL.EXE processes.

The main reason for the wrapper is that ``Excel.Application.Quit()`` may block
inside a worker thread after the workbook has already been saved.  Instead we
send WM_CLOSE to the exact window handle returned by the program-owned Excel
instance.  The exporter can then return normally and release its COM proxies;
Excel exits when those references are released.
"""

import ctypes
import logging
import os
import threading
from typing import Callable, Any

logger = logging.getLogger(__name__)

_WM_CLOSE = 0x0010
_tls = threading.local()
_installed = False
_original_dispatch_ex = None


def set_excel_progress_reporter(reporter: Callable[[str], None] | None) -> None:
    """Attach a progress sink to the current worker thread only."""
    _tls.reporter = reporter


def clear_excel_progress_reporter() -> None:
    try:
        del _tls.reporter
    except AttributeError:
        pass


def _report(message: str) -> None:
    reporter = getattr(_tls, "reporter", None)
    if reporter is None:
        return
    try:
        reporter(str(message))
    except Exception:
        logger.exception("Не удалось передать этап Excel в окно фоновых задач")


class _WorkbookProxy:
    """Transparent proxy for a newly created output workbook."""

    __slots__ = ("_target",)

    def __init__(self, target: Any) -> None:
        object.__setattr__(self, "_target", target)

    def __getattr__(self, name: str):
        return getattr(object.__getattribute__(self, "_target"), name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_target"), name, value)

    def SaveAs(self, *args, **kwargs):  # noqa: N802 - COM API
        result = object.__getattribute__(self, "_target").SaveAs(*args, **kwargs)
        _report("Excel: файл сохранён.")
        return result

    def Save(self, *args, **kwargs):  # noqa: N802 - COM API
        result = object.__getattribute__(self, "_target").Save(*args, **kwargs)
        _report("Excel: файл сохранён.")
        return result


class _WorkbooksProxy:
    """Wrap only workbooks created by the application; opened source files stay raw."""

    __slots__ = ("_target",)

    def __init__(self, target: Any) -> None:
        object.__setattr__(self, "_target", target)

    def __getattr__(self, name: str):
        return getattr(object.__getattribute__(self, "_target"), name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_target"), name, value)

    def Add(self, *args, **kwargs):  # noqa: N802 - COM API
        workbook = object.__getattribute__(self, "_target").Add(*args, **kwargs)
        return _WorkbookProxy(workbook)

    def Open(self, *args, **kwargs):  # noqa: N802 - COM API
        # Source workbooks are intentionally not proxied: some exporters close
        # them during the calculation, before the final output workbook is done.
        return object.__getattribute__(self, "_target").Open(*args, **kwargs)


class _ExcelApplicationProxy:
    """Proxy for one Excel instance created by Procurement via DispatchEx."""

    __slots__ = ("_target", "_hwnd")

    def __init__(self, target: Any) -> None:
        object.__setattr__(self, "_target", target)
        try:
            hwnd = int(target.Hwnd)
        except Exception:
            hwnd = 0
        object.__setattr__(self, "_hwnd", hwnd)

    def __getattr__(self, name: str):
        target = object.__getattribute__(self, "_target")
        if name == "Workbooks":
            return _WorkbooksProxy(target.Workbooks)
        return getattr(target, name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_target"), name, value)

    def Quit(self):  # noqa: N802 - COM API
        """Request closure of only the program-owned Excel window, without blocking."""
        target = object.__getattribute__(self, "_target")
        hwnd = object.__getattribute__(self, "_hwnd")
        _report("Excel: завершаю служебный процесс Excel...")
        try:
            target.DisplayAlerts = False
        except Exception:
            pass

        # Do not use taskkill/TerminateProcess and do not enumerate Excel.
        # WM_CLOSE is sent only to the HWND captured from this DispatchEx object.
        if os.name == "nt" and hwnd:
            try:
                posted = ctypes.windll.user32.PostMessageW(hwnd, _WM_CLOSE, 0, 0)
                if posted:
                    _report("Excel: служебному процессу отправлена команда завершения.")
                    return None
            except Exception:
                logger.exception("Не удалось отправить WM_CLOSE служебному Excel")

        # Fallback for an unusual environment where HWND is unavailable.
        # This still addresses only the DispatchEx-created instance.
        result = target.Quit()
        _report("Excel: служебный процесс завершён.")
        return result


def install_excel_dispatch_guard() -> bool:
    """Install the guard once for all current and future Excel exporters."""
    global _installed, _original_dispatch_ex
    if _installed:
        return True
    try:
        import win32com.client as win32
    except ImportError:
        # Allows non-Windows tooling/tests to import the package.
        return False

    original = win32.DispatchEx
    if getattr(original, "_procurement_excel_guard", False):
        _installed = True
        return True

    def guarded_dispatch_ex(prog_id, *args, **kwargs):
        obj = original(prog_id, *args, **kwargs)
        if str(prog_id).strip().lower() == "excel.application":
            return _ExcelApplicationProxy(obj)
        return obj

    guarded_dispatch_ex._procurement_excel_guard = True
    guarded_dispatch_ex._procurement_original = original
    _original_dispatch_ex = original
    win32.DispatchEx = guarded_dispatch_ex
    _installed = True
    return True
