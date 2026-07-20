"""Application logging setup."""

from __future__ import annotations

import logging
from pathlib import Path


def setup_logger(log_file: str | Path) -> logging.Logger:
    """Create a file logger for extraction errors and processing details."""
    logger = logging.getLogger("tcr_extractor")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger
