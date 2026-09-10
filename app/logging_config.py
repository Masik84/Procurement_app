from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

_CONFIGURED = False


def setup_logging(base_dir: Path, *, level: int = logging.INFO) -> Path:
    """Configure logging for the whole application.

    Call this once, as early as possible (before any other app.* module logs
    anything) - ideally at the very top of main.py, before importing pages.

    Every module in the project uses ``logger = logging.getLogger(__name__)``,
    so the log line always shows which module logged it (e.g.
    ``app.exports.supplier_price_exporter``), and ``logger.exception(...)``
    additionally records the full traceback (exact file/line/function where
    the error actually happened), not just the place that caught it.

    Writes to <base_dir>/logs/app.log (rotated at 5 MB, 5 backups kept) and
    also echoes to stderr so it is visible during development.
    """
    global _CONFIGURED
    logs_dir = base_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / "app.log"

    if _CONFIGURED:
        return log_path

    formatter = logging.Formatter(
        fmt="%(asctime)s %(levelname)-8s %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        log_path, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    # Third-party libraries are noisy on INFO; keep them quieter unless
    # something goes wrong.
    for noisy in ("PIL", "urllib3", "matplotlib"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    _CONFIGURED = True
    logging.getLogger(__name__).info("Логирование инициализировано, файл: %s", log_path)
    return log_path
