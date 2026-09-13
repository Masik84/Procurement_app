from __future__ import annotations

import faulthandler
import logging
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

_CONFIGURED = False
_FAULT_STREAM = None


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
    writes routine diagnostics to the log file; only ERROR/CRITICAL messages
    are echoed to stderr during development.
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
    # Keep routine diagnostics in app.log only.  The console is useful during
    # development for actual failures, but startup INFO/WARNING chatter should
    # not be shown to users (and packaged GUI .exe builds may have no console).
    console_handler.setLevel(logging.ERROR)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    # Third-party libraries are noisy on INFO; keep them quieter unless
    # something goes wrong.
    for noisy in ("PIL", "urllib3", "matplotlib"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    # Native crashes (for example Windows 0xC0000005 from Qt/PySide) bypass
    # Python exception handlers completely.  Route faulthandler to the same
    # app.log so a fatal native traceback is not lost in the console window.
    # Keep the stream alive for the lifetime of the process: faulthandler
    # writes directly to its file descriptor when the interpreter is crashing.
    global _FAULT_STREAM
    try:
        _FAULT_STREAM = log_path.open("a", encoding="utf-8", buffering=1)
        _FAULT_STREAM.write("\n--- native crash capture enabled ---\n")
        _FAULT_STREAM.flush()
        faulthandler.enable(file=_FAULT_STREAM, all_threads=True)
    except Exception:
        # File logging itself is already available through RotatingFileHandler.
        # Never make the application fail merely because native crash capture
        # could not be enabled.
        logging.getLogger(__name__).exception(
            "Не удалось включить запись native crash dump в app.log"
        )

    _CONFIGURED = True
    log = logging.getLogger(__name__)
    log.info("Логирование инициализировано, файл: %s", log_path)
    log.info("Native crash capture: %s", "app.log" if _FAULT_STREAM is not None else "недоступен")
    flush_logs()
    return log_path


def flush_logs() -> None:
    """Flush all configured logging streams immediately.

    Startup diagnostics must reach disk before a possible native Qt crash,
    because an access violation gives Python no chance to flush buffered data.
    """
    root_logger = logging.getLogger()
    for handler in list(root_logger.handlers):
        try:
            handler.flush()
        except Exception:
            pass
    if _FAULT_STREAM is not None:
        try:
            _FAULT_STREAM.flush()
        except Exception:
            pass


def log_startup_stage(logger: logging.Logger, stage: str) -> None:
    """Write and flush one startup checkpoint used to localise native crashes."""
    logger.info("STARTUP | %s", stage)
    flush_logs()


def install_thread_exception_logging(logger: logging.Logger) -> None:
    """Record uncaught exceptions from Python worker threads as well."""
    def _thread_hook(args):
        logger.critical(
            "Необработанное исключение в потоке %s",
            getattr(args.thread, "name", "<unknown>"),
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )
        flush_logs()

    threading.excepthook = _thread_hook
