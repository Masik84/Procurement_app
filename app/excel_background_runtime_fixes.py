from __future__ import annotations

"""Compatibility stub.

The previous revision installed a global Excel DispatchEx/WM_CLOSE proxy and
changed the task-center Hide button into Minimize. Both behaviours are disabled.
Excel exporters now use their normal lifecycle, and Product Articles uses a
non-COM export path from excel_export_worker.py.
"""

_installed = False


def install_excel_background_runtime_fixes() -> None:
    global _installed
    _installed = True
