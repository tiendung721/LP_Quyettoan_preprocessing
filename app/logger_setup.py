"""Thiết lập logging cho ứng dụng.

Ghi log ra cả file (theo ngày) và console. Tên logger dùng chung cho toàn app
để các module khác chỉ cần ``logging.getLogger("QuyetToan")``.
"""

from __future__ import annotations

import logging
import os
from datetime import datetime

LOGGER_NAME = "QuyetToan"

# Định dạng log dùng giờ local của máy (không dùng UTC).
_FILE_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def setup_logger(logs_dir: str, level: int = logging.INFO) -> logging.Logger:
    """Khởi tạo và trả về logger chính của ứng dụng.

    Args:
        logs_dir: Thư mục chứa file log (ví dụ ``D:\\RPA_QuyetToan\\Logs``).
        level: Mức log tối thiểu.
    """
    os.makedirs(logs_dir, exist_ok=True)

    log_file = os.path.join(logs_dir, f"app_{datetime.now():%Y%m%d}.log")

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)

    # Tránh thêm handler trùng lặp nếu setup_logger được gọi nhiều lần.
    if logger.handlers:
        return logger

    formatter = logging.Formatter(_FILE_FORMAT, _DATE_FORMAT)

    # Handler ghi ra file (mã hóa utf-8 để hiển thị tiếng Việt).
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler ghi ra console để tiện debug khi chạy python main.py.
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.propagate = False
    return logger


def get_logger() -> logging.Logger:
    """Trả về logger đã cấu hình (dùng ở các module khác)."""
    return logging.getLogger(LOGGER_NAME)
