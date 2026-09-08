from __future__ import annotations

# Runtime compatibility layer for the current "no IS" application mode.
# It keeps the legacy DB schema intact while removing active IS behaviour
# from supplier orders, order planning, UI and Excel exports.
from app.no_is_runtime import install_no_is_runtime

install_no_is_runtime()
