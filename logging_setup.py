"""
Centralized logging.

The original tool's only diagnostic output was print() into a console
window. Package this as a windowed executable (which you want, so no
console flashes open on launch) and every one of those messages goes
into a void — there's no way to debug a field issue.

configure_logging() attaches a RotatingFileHandler to the "balboa" logger
once, at startup. Every other module gets a child logger via get_logger(),
which propagates up to "balboa" automatically — no per-module handler
wiring needed.

log_bridge() is the adapter used throughout the service layer: it always
writes to the file logger, and optionally also forwards to a UI callback
(a Qt signal's .emit, or a QPlainTextEdit append) for live display. That
means a log message is never silently discarded just because a dialog
wasn't open to receive it — which was happening in a couple of call sites
in the original UI wiring (log=lambda m: None).
"""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Callable, Optional

LOGGER_ROOT = "balboa"


def configure_logging(log_dir: Path, level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(LOGGER_ROOT)
    if logger.handlers:
        return logger  # already configured, e.g. re-entrant call

    log_dir.mkdir(parents=True, exist_ok=True)
    logger.setLevel(level)

    file_handler = RotatingFileHandler(
        log_dir / "balboa.log", maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-7s  %(name)-24s  %(message)s"))
    logger.addHandler(file_handler)

    # Only echo to console in a dev run -- a frozen/windowed build has no
    # console to write to, and attaching one causes noisy errors on Windows.
    if not getattr(sys, "frozen", False):
        console = logging.StreamHandler(sys.stdout)
        console.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(console)

    def _log_uncaught(exc_type, exc_value, exc_tb):
        logger.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = _log_uncaught

    logger.info("Logging initialized -> %s", log_dir / "balboa.log")
    return logger


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(f"{LOGGER_ROOT}.{name}")


def log_bridge(ui_callback: Optional[Callable[[str], None]] = None,
               logger: Optional[logging.Logger] = None) -> Callable[[str], None]:
    """Returns a callable(msg) that always writes to the file logger and
    optionally forwards to a UI callback for live display."""
    logger = logger or logging.getLogger(LOGGER_ROOT)

    def _log(msg: str) -> None:
        logger.info(msg)
        if ui_callback:
            try:
                ui_callback(msg)
            except Exception:
                logger.exception("UI log callback raised")

    return _log
