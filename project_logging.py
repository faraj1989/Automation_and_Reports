import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

from project_config import env_path_str, load_env_file

load_env_file()

DEFAULT_LOG_DIR = Path(env_path_str("PROJECT_LOG_DIR", str(Path(__file__).resolve().parent / "logs")))


def ensure_log_dir(log_dir: Path | str = None) -> Path:
    path = Path(log_dir) if log_dir else DEFAULT_LOG_DIR
    path = path.expanduser()
    if not path.is_absolute():
        path = DEFAULT_LOG_DIR / path
    path.mkdir(parents=True, exist_ok=True)
    return path


def setup_logger(name: str, log_file: str | Path = None, level: int = logging.INFO, console: bool = True) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(level)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    if console:
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        logger.addHandler(stream_handler)

    file_path = Path(log_file) if log_file else ensure_log_dir() / f"{name}.log"
    file_path = file_path.expanduser()
    if not file_path.is_absolute():
        file_path = ensure_log_dir() / file_path
    file_path.parent.mkdir(parents=True, exist_ok=True)

    file_handler = TimedRotatingFileHandler(str(file_path), when="midnight", backupCount=14, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.debug("Logger initialized. Writing logs to %s", file_path)
    return logger
