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
# arguments.
try:
    from app.background_task_fixes import install_background_task_fixes
except ModuleNotFoundError as exc:
    if exc.name != "PySide6":
        raise
else:
    install_background_task_fixes()

# Supplier Price: remove row-by-row DB round-trips and keep CostCalc Excel as a
# separate background task after the DB save/calculation has completed.
try:
    from app.supplier_price_performance_fixes import install_supplier_price_performance_fixes
except ModuleNotFoundError as exc:
    if exc.name != "PySide6":
        raise
else:
    install_supplier_price_performance_fixes()

# Supplier Price follow-up:
# - bulk import/matching into staging;
# - compatibility between optimized group CostCalc and the no-IS adapter.
try:
    from app.supplier_price_runtime_fixes import install_supplier_price_runtime_fixes
except ModuleNotFoundError as exc:
    if exc.name != "PySide6":
        raise
else:
    install_supplier_price_runtime_fixes()

# Product Stock save performance follow-up. Older working copies may not have
# this optional runtime module yet, so absence of the module itself is allowed.
try:
    from app.product_stock_performance_fixes import install_product_stock_performance_fixes
except ModuleNotFoundError as exc:
    if exc.name not in {"PySide6", "app.product_stock_performance_fixes"}:
        raise
else:
    install_product_stock_performance_fixes()

# Customer Costs: make calculation Excel independent of native Excel COM and
# keep background completion state initialized for every page instance.
try:
    from app.customer_cost_background_fixes import install_customer_cost_background_fixes
except ModuleNotFoundError as exc:
    if exc.name not in {"PySide6", "openpyxl", "app.customer_cost_background_fixes"}:
        raise
else:
    install_customer_cost_background_fixes()

# Target Price: calculated workbook is also the source-selection workbook.
try:
    from app.target_price_selection_workflow import install_target_price_selection_workflow
except ModuleNotFoundError as exc:
    if exc.name != "PySide6":
        raise
else:
    install_target_price_selection_workflow()

# Excel rule: every uC3 column, regardless of its concrete caption variant,
# must be exported as a numeric value with integer display format.
from app.uc3_excel_format import install_uc3_integer_format

install_uc3_integer_format()
