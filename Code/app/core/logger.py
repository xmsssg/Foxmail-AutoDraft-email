from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from app.core.config_loader import expand_path


def setup_logger(settings: dict) -> logging.Logger:
    log_settings = settings["logging"]
    log_path = Path(expand_path(log_settings["path"]))
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("autoemail")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    max_bytes = int(log_settings.get("max_size_mb", 20)) * 1024 * 1024
    backup_count = int(log_settings.get("backup_count", 5))

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    console_handler = logging.StreamHandler()

    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger
