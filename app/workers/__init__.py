"""Background-worker helpers for Procurement App.

Long-running Excel exports are started through
``app.workers.excel_export_worker.start_excel_export`` which now delegates to
the application-wide BackgroundTaskManager.

Do not import the removed legacy ``ExcelExportWorker`` here: importing any
``app.workers.*`` submodule executes this package initializer first, and the
stale re-export prevented pages such as Target uC3 from being imported.
"""

__all__: list[str] = []
