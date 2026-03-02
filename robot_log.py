"""Логирование работы робота в robot.log."""
import logging
from pathlib import Path

LOG_FILE = Path(__file__).resolve().parent / "robot.log"
_logger = None


def get_logger():
    global _logger
    if _logger is not None:
        return _logger
    _logger = logging.getLogger("robot")
    _logger.setLevel(logging.INFO)
    if not _logger.handlers:
        fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        _logger.addHandler(fh)
    return _logger
