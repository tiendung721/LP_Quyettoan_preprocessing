"""Điểm khởi chạy ứng dụng Trợ Lý Quyết Toán RPA.

Chạy bằng:  python main.py
"""

from __future__ import annotations

import sys
import traceback


def main() -> int:
    # Import bên trong hàm để có thể hiển thị thông báo thân thiện nếu
    # thiếu thư viện (PySide6/watchdog/openpyxl).
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
    except ModuleNotFoundError:
        print(
            "Thiếu thư viện PySide6. Vui lòng cài đặt bằng lệnh:\n"
            "    pip install -r requirements.txt"
        )
        return 1

    from app import theme
    from app.config import AppConfig
    from app.database import Database
    from app.logger_setup import setup_logger
    from app.main_window import MainWindow

    app = QApplication(sys.argv)
    app.setApplicationName("Trợ Lý Quyết Toán RPA")
    theme.apply_theme(app)

    try:
        # 1) Nạp cấu hình (tự tạo settings.json nếu chưa có).
        config = AppConfig.load()
        # 2) Tạo các thư mục cần thiết.
        config.ensure_folders()
        # 3) Thiết lập logging.
        logger = setup_logger(config.logs_dir)
        logger.info("==== Khởi động ứng dụng ====")
        logger.info("Đã nạp cấu hình từ: %s", config.path)
        # 4) Khởi tạo database.
        database = Database(config.database_path)
        database.init_db()
        logger.info("Đã khởi tạo database: %s", config.database_path)
        # 5) Mở cửa sổ chính.
        window = MainWindow(config, database, logger)
        window.show()
    except Exception as exc:  # noqa: BLE001 - hiển thị lỗi thân thiện thay vì crash
        traceback.print_exc()
        QMessageBox.critical(
            None,
            "Lỗi khởi động",
            "Ứng dụng gặp lỗi khi khởi động.\n"
            f"Chi tiết: {exc}",
        )
        return 1

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
