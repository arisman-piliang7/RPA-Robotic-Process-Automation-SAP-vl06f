"""
utils/logger.py
---------------
Centralised logging configuration for the RPA application.
"""

import logging
import os
from logging.handlers import TimedRotatingFileHandler


def setup_logger(log_dir: str = "logs",
                 log_level: str = "INFO",
                 backup_count: int = 7) -> logging.Logger:
    """
    Configure and return the root application logger.

    Creates a ``logs/`` directory if it does not exist and registers two
    handlers:

    * **StreamHandler** – writes INFO+ messages to the console.
    * **TimedRotatingFileHandler** – writes all messages to a daily log file
      inside *log_dir*, keeping *backup_count* rotated copies.

    Parameters
    ----------
    log_dir:
        Directory where log files will be stored.
    log_level:
        Minimum log level string, e.g. ``"DEBUG"`` or ``"INFO"``.
    backup_count:
        How many daily log files to keep before rotation removes old ones.

    Returns
    -------
    logging.Logger
        Configured root logger.
    """
    os.makedirs(log_dir, exist_ok=True)

    numeric_level = getattr(logging, log_level.upper(), logging.INFO)

    logger = logging.getLogger("rpa_vl06f")
    logger.setLevel(numeric_level)

    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-8s] %(name)s – %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(numeric_level)
    console_handler.setFormatter(formatter)

    log_file = os.path.join(log_dir, "rpa_vl06f.log")
    file_handler = TimedRotatingFileHandler(
        filename=log_file,
        when="midnight",
        interval=1,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(numeric_level)
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger
