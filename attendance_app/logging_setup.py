"""Application logging configuration."""

from __future__ import annotations

import logging
from pathlib import Path


LOGGER_NAME = "attendance_app"


class _PortableFileHandler(logging.FileHandler):
    """Release the file between writes so temporary app roots can be removed."""

    def __init__(self, filename):
        super().__init__(filename, encoding="utf-8", delay=True)

    def emit(self, record):
        try:
            super().emit(record)
        finally:
            self.close()


def configure_logging(config) -> logging.Logger:
    """Configure the application's stable file logger and return it.

    Configuration is idempotent for a given log path, which is important when
    repositories and services are constructed directly by callers or tests.
    """
    log_path = Path(config.log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    resolved = str(log_path.resolve())
    for handler in list(logger.handlers):
        if getattr(handler, "_attendance_log_path", None) == resolved:
            return logger
        logger.removeHandler(handler)
        handler.close()
    log_path.touch(exist_ok=True)
    handler = _PortableFileHandler(log_path)
    handler._attendance_log_path = resolved  # type: ignore[attr-defined]
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    return logger


# Short alias for callers that prefer the conventional setup name.
setup_logging = configure_logging


def get_logger() -> logging.Logger:
    return logging.getLogger(LOGGER_NAME)
