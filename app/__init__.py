from __future__ import annotations

# Runtime compatibility layer for the current "no IS" application mode.
# It keeps the legacy DB schema intact while removing active IS behaviour
# from supplier orders, order planning, UI and Excel exports.
from app.no_is_runtime import install_no_is_runtime

install_no_is_runtime()

# Excel rule: every uC3 column, regardless of its concrete caption variant,
# must be exported as a numeric value with integer display format.
from app.uc3_excel_format import install_uc3_integer_format

install_uc3_integer_format()
