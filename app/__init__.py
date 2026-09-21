from __future__ import annotations

# Runtime compatibility layer for the current "no IS" application mode.
# Keep it installed before pages/services are imported.
from app.no_is_runtime import install_no_is_runtime

install_no_is_runtime()

# Compatibility fixes for the no-IS runtime adapters introduced by the
# systemic hotfix (ProductStock primary key + Order Planning signal handling).
from app.runtime_consistency_fixes import install_runtime_consistency_fixes

install_runtime_consistency_fixes()

# Keep background informational messages in the central task window and make
# the no-IS Order Planning export wrapper compatible with restored 3/5-month
# arguments.  This patch deliberately does not replace the task manager,
# QThread lifecycle or Excel COM objects.
try:
    from app.background_task_fixes import install_background_task_fixes
except ModuleNotFoundError as exc:
    # DB/Alembic tooling may import app without desktop GUI dependencies.
    if exc.name != "PySide6":
        raise
else:
    install_background_task_fixes()

# Supplier Price: remove row-by-row DB round-trips and keep CostCalc Excel as a
# separate background task after the DB save/calculation has completed.
try:
    from app.supplier_price_performance_fixes import install_supplier_price_performance_fixes
except ModuleNotFoundError as exc:
    # DB/Alembic tooling may import app without desktop GUI dependencies.
    if exc.name != "PySide6":
        raise
else:
    install_supplier_price_performance_fixes()

# Excel rule: every uC3 column, regardless of its concrete caption variant,
# must be exported as a numeric value with integer display format.
from app.uc3_excel_format import install_uc3_integer_format

install_uc3_integer_format()
