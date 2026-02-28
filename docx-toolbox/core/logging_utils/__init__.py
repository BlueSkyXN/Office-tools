"""统一日志模块 — 遵循 CORE-INTERFACE.md §5"""

import logging
import os
import threading
from datetime import datetime
from pathlib import Path


LOG_FORMAT = "%(asctime)s | %(levelname)-5s | %(module)s | %(message)s"
LOG_DATEFMT = "%Y-%m-%dT%H:%M:%S%z"

_logger: logging.Logger | None = None
_logger_lock = threading.Lock()


def get_logger(name: str = "docx_toolbox") -> logging.Logger:
    """获取共享 logger，首次调用时自动初始化"""
    global _logger
    if _logger is not None:
        return _logger
    with _logger_lock:
        if _logger is not None:
            return _logger
        _logger = logging.getLogger(name)
        _logger.setLevel(logging.DEBUG)

        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        console.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=LOG_DATEFMT))
        _logger.addHandler(console)
    return _logger


def setup_file_logging(log_dir: str | Path | None = None, task_id: str = "") -> Path:
    """为指定任务创建文件日志处理器，返回日志文件路径"""
    logger = get_logger()

    if log_dir is None:
        log_dir = Path("logs") / datetime.now().strftime("%Y-%m-%d")
    else:
        log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    filename = f"{task_id}.log" if task_id else "session.log"
    log_path = log_dir / filename

    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=LOG_DATEFMT))
    logger.addHandler(handler)
    return log_path
