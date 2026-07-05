"""Logger estruturado, prefixado por módulo."""
import logging
import sys
from typing import Final

from config.settings import APP_CONFIG

_FMT: Final = "%(asctime)s | %(levelname)-7s | %(name)-30s | %(message)s"
_DATE_FMT: Final = "%Y-%m-%d %H:%M:%S"


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(APP_CONFIG.log_level.upper())
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter(_FMT, _DATE_FMT))
    logger.addHandler(handler)
    logger.propagate = False
    return logger
