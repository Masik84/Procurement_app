from __future__ import annotations

# Runtime compatibility layer for the current "no IS" application mode.
# Keep it installed before pages/services are imported.
from app.no_is_runtime import install_no_is_runtime

install_no_is_runtime()

# Compatibility fixes for the no-IS runtime adapters introduced by the
# systemic hotfix (ProductStock primary key + Order Planning signal handling).
from app.runtime_consistency_fixes import install_runtime_consistency_fixes

install_runtime_consistency_fixes()

# Excel rule: every uC3 column, regardless of its concrete caption variant,
# must be exported as a numeric value with integer display format.
from app.uc3_excel_format import install_uc3_integer_format

install_uc3_integer_format()
