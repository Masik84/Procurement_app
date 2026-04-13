from __future__ import annotations

import getpass
import uuid


def generate_import_batch_id() -> str:
    """
    Returns a UUID string suitable for batch_id fields
    used in temp/staging tables.
    """
    return str(uuid.uuid4())


def get_current_username(default: str = "system") -> str:
    """
    Returns the current OS username.
    Falls back to `default` if unavailable.
    """
    try:
        username = getpass.getuser().strip()
        return username or default
    except Exception:
        return default