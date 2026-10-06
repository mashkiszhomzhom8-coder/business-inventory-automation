import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "app.log"
APP_LOGGER_NAME = "inventory_app"


def setup_logging():
    """Настраивает и возвращает файловый логгер приложения.

    Returns:
        Логгер с ротацией файла журнала.
    """
    LOG_DIR.mkdir(exist_ok=True)
    logger = logging.getLogger(APP_LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not any(
        isinstance(handler, RotatingFileHandler) and Path(handler.baseFilename) == LOG_FILE
        for handler in logger.handlers
    ):
        handler = RotatingFileHandler(LOG_FILE, maxBytes=1024 * 1024, backupCount=5, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        logger.addHandler(handler)

    return logger
