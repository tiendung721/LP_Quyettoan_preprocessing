"""Theo dõi thư mục Downloads bằng watchdog.

Khi phát hiện file output hợp lệ, xử lý trong luồng riêng (không block UI):
chờ file tải xong -> sao lưu + chuyển vào Outputs -> ghi database ->
phát tín hiệu (signal) về MainWindow.
"""

from __future__ import annotations

import os
import threading
from typing import TYPE_CHECKING

from PySide6.QtCore import QObject, Signal
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from . import file_utils

if TYPE_CHECKING:  # tránh import vòng khi chạy thật
    import logging

    from .config import AppConfig
    from .database import Database


class WatcherSignals(QObject):
    """Các tín hiệu phát về luồng UI (kết nối kiểu queued tự động)."""

    output_ready = Signal(dict)   # có file output mới đã sẵn sàng kiểm tra
    file_error = Signal(str)      # lỗi khi xử lý một file (không crash app)
    log_message = Signal(str)     # dòng log để hiển thị trên UI


class _DownloadEventHandler(FileSystemEventHandler):
    def __init__(self, watcher: "DownloadWatcher"):
        super().__init__()
        self.watcher = watcher

    def on_created(self, event):
        if not event.is_directory:
            self.watcher.handle_candidate(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self.watcher.handle_candidate(event.src_path)

    def on_moved(self, event):
        # Chrome đôi khi đổi tên .crdownload -> tên thật khi tải xong.
        dest = getattr(event, "dest_path", None)
        if dest and not event.is_directory:
            self.watcher.handle_candidate(dest)


class DownloadWatcher:
    def __init__(self, config: "AppConfig", database: "Database", logger: "logging.Logger"):
        self.config = config
        self.database = database
        self.logger = logger
        self.signals = WatcherSignals()

        self._observer = None
        self._processing = set()
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ #
    def start(self) -> None:
        """Bắt đầu theo dõi thư mục download."""
        folder = self.config.download_folder
        os.makedirs(folder, exist_ok=True)

        self._observer = Observer()
        self._observer.schedule(_DownloadEventHandler(self), folder, recursive=False)
        self._observer.start()

        msg = f"Bắt đầu theo dõi thư mục tải về: {folder}"
        self.logger.info(msg)
        self.signals.log_message.emit(msg)

    def stop(self) -> None:
        """Dừng theo dõi (gọi khi đóng app hoặc đổi cấu hình)."""
        if self._observer is not None:
            try:
                self._observer.stop()
                self._observer.join(timeout=5)
            except Exception:  # noqa: BLE001 - không để việc dừng gây crash
                pass
            self._observer = None
            self.logger.info("Đã dừng theo dõi thư mục tải về.")

    # ------------------------------------------------------------------ #
    def handle_candidate(self, path: str) -> None:
        """Lọc nhanh và khởi động luồng xử lý cho một file ứng viên."""
        try:
            # Bỏ qua file tạm khi đang tải.
            if file_utils.is_temp_download_file(path):
                return
            # Bỏ qua nếu không đúng đuôi / không khớp mẫu tên.
            if not file_utils.is_allowed_output_file(
                path,
                self.config.allowed_extensions,
                self.config.output_file_patterns,
            ):
                return

            with self._lock:
                if path in self._processing:
                    return
                self._processing.add(path)

            worker = threading.Thread(
                target=self._process_file, args=(path,), daemon=True
            )
            worker.start()
        except Exception:  # noqa: BLE001
            self.logger.exception("Lỗi khi lọc file ứng viên: %s", path)

    # ------------------------------------------------------------------ #
    def _process_file(self, path: str) -> None:
        """Chờ file ổn định, sao lưu, chuyển vào output và ghi database."""
        name = os.path.basename(path)
        try:
            self._emit_log(f"Phát hiện file mới: {name}")

            stable = file_utils.wait_until_file_stable(
                path,
                stable_seconds=self.config.download_stable_seconds,
                timeout=60,
            )
            if not stable:
                # File có thể đã bị di chuyển bởi lần xử lý khác, hoặc tải lỗi.
                if not os.path.exists(path):
                    return
                warn = (
                    f"File '{name}' tải chưa ổn định sau 60 giây, tạm bỏ qua. "
                    "Vui lòng kiểm tra lại việc tải file."
                )
                self.logger.warning(warn)
                self.signals.file_error.emit(warn)
                return

            self._emit_log(f"File đã tải xong (ổn định): {name}")

            result = file_utils.safe_move_download_to_output(
                path, self.config.output_folder, self.config.backup_folder
            )
            self._emit_log(
                "Đã sao lưu bản gốc và chuyển vào thư mục output: "
                f"{result['file_name']}"
            )

            record_id = self.database.insert_processed_file(
                working_path=result["working_path"],
                status="READY_FOR_REVIEW",
                original_download_path=result["original_download_path"],
                backup_path=result["backup_path"],
                file_name=result["file_name"],
                file_size=result["file_size"],
                file_hash=result["file_hash"],
                note="Trợ lý đã bóc tách dữ liệu xong.",
            )
            result["id"] = record_id
            result["status"] = "READY_FOR_REVIEW"

            self.logger.info(
                "Đã ghi database bản ghi #%s cho file %s",
                record_id,
                result["file_name"],
            )
            self.signals.output_ready.emit(result)
        except Exception as exc:  # noqa: BLE001 - lỗi 1 file không được làm sập app
            self.logger.exception("Lỗi khi xử lý file tải về: %s", path)
            self.signals.file_error.emit(
                f"Không xử lý được file '{name}': {exc}"
            )
        finally:
            with self._lock:
                self._processing.discard(path)

    # ------------------------------------------------------------------ #
    def _emit_log(self, message: str) -> None:
        self.logger.info(message)
        self.signals.log_message.emit(message)
