"""Giao diện chính (PySide6) của Trợ Lý Quyết Toán RPA.

Gồm 5 vùng: cấu hình, mở trợ lý, file output hiện tại, duyệt dữ liệu, lịch sử/log.
Xử lý toàn bộ button handler, nhận tín hiệu từ watcher và cập nhật giao diện.
"""

from __future__ import annotations

import logging
import os
import subprocess
from typing import Any, Dict, Optional

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import excel_preview, file_utils
from .config import AppConfig
from .database import Database
from .watcher import DownloadWatcher

# Nhãn tiếng Việt cho từng trạng thái.
STATUS_LABELS = {
    "WAITING_FOR_DOWNLOAD": "Đang chờ tải file output",
    "DOWNLOADED": "Đã tải về",
    "READY_FOR_REVIEW": "Sẵn sàng để kiểm tra",
    "OPENED_FOR_REVIEW": "Đang mở để kiểm tra",
    "REVIEW_SAVED": "Đã lưu sau khi chỉnh sửa",
    "REVIEW_CONFIRMED": "Đã xác nhận dùng file này",
    "READY_TO_PREVIEW": "Sẵn sàng duyệt dữ liệu",
    "PREVIEWED": "Đã duyệt dữ liệu",
    "ERROR": "Lỗi",
}

# Các trạng thái coi như đã được người dùng xác nhận (cho phép duyệt dữ liệu).
CONFIRMED_STATUSES = {"REVIEW_CONFIRMED", "READY_TO_PREVIEW", "PREVIEWED"}


def status_text(status: Optional[str]) -> str:
    if not status:
        return "—"
    return f"{STATUS_LABELS.get(status, status)} ({status})"


# ---------------------------------------------------------------------------
# Cầu nối logging -> QPlainTextEdit
# ---------------------------------------------------------------------------
class _LogEmitter(QObject):
    message = Signal(str)


class _QtLogHandler(logging.Handler):
    """Handler đẩy log runtime lên ô log của giao diện."""

    def __init__(self, emitter: _LogEmitter):
        super().__init__()
        self.emitter = emitter
        self.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        )

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.emitter.message.emit(self.format(record))
        except Exception:  # noqa: BLE001 - lỗi hiển thị log không được làm sập app
            pass


# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, config: AppConfig, database: Database, logger: logging.Logger):
        super().__init__()
        self.config = config
        self.database = database
        self.logger = logger

        # Trạng thái vận hành.
        self.overall_status: str = "WAITING_FOR_DOWNLOAD"
        self.current_working_file: Optional[Dict[str, Any]] = None
        self.current_reviewed_file: Optional[str] = None

        self.setWindowTitle("Trợ Lý Quyết Toán RPA")
        self.resize(1080, 820)

        self._build_ui()
        self._setup_log_bridge()

        # Watcher theo dõi thư mục download.
        self.watcher = DownloadWatcher(self.config, self.database, self.logger)
        self.watcher.signals.output_ready.connect(self.on_output_ready)
        self.watcher.signals.file_error.connect(self.on_file_error)
        self.watcher.signals.log_message.connect(self.append_log)

        self._load_last_record()
        self._start_watcher()

    # ================================================================== #
    # Dựng giao diện
    # ================================================================== #
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        outer.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setSpacing(10)

        layout.addWidget(self._build_config_group())
        layout.addWidget(self._build_open_assistant_group())
        layout.addWidget(self._build_current_file_group())
        layout.addWidget(self._build_preview_group())
        layout.addWidget(self._build_history_group(), stretch=1)

    # ---- Vùng 1: Cấu hình ------------------------------------------- #
    def _build_config_group(self) -> QGroupBox:
        box = QGroupBox("1. Cấu hình")
        grid = QGridLayout(box)

        self.edit_bat = QLineEdit(self.config.bat_path)
        self.edit_download = QLineEdit(self.config.download_folder)
        self.edit_output = QLineEdit(self.config.output_folder)
        self.edit_backup = QLineEdit(self.config.backup_folder)
        self.edit_daily = QLineEdit(self.config.daily_tracking_file)

        rows = [
            ("Đường dẫn file .bat:", self.edit_bat, self._browse_bat),
            ("Thư mục download:", self.edit_download, self._browse_download),
            ("Thư mục output (file làm việc):", self.edit_output, self._browse_output),
            ("Thư mục backup (bản gốc):", self.edit_backup, self._browse_backup),
            ("File theo dõi hàng ngày:", self.edit_daily, self._browse_daily),
        ]
        for r, (label, edit, handler) in enumerate(rows):
            grid.addWidget(QLabel(label), r, 0)
            grid.addWidget(edit, r, 1)
            btn = QPushButton("Chọn...")
            btn.clicked.connect(handler)
            grid.addWidget(btn, r, 2)

        btn_row = QHBoxLayout()
        self.btn_save_config = QPushButton("Lưu cấu hình")
        self.btn_save_config.clicked.connect(self.on_save_config)
        self.btn_check_config = QPushButton("Kiểm tra cấu hình")
        self.btn_check_config.clicked.connect(self.on_check_config)
        self.btn_open_output = QPushButton("Mở thư mục output")
        self.btn_open_output.clicked.connect(self.on_open_output_folder)
        btn_row.addWidget(self.btn_save_config)
        btn_row.addWidget(self.btn_check_config)
        btn_row.addWidget(self.btn_open_output)
        btn_row.addStretch(1)
        grid.addLayout(btn_row, len(rows), 0, 1, 3)

        return box

    # ---- Vùng 2: Mở trợ lý ------------------------------------------ #
    def _build_open_assistant_group(self) -> QGroupBox:
        box = QGroupBox("2. Mở trợ lý quyết toán")
        v = QVBoxLayout(box)

        self.btn_open_assistant = QPushButton("Mở trợ lý quyết toán")
        self.btn_open_assistant.setMinimumHeight(56)
        big = QFont()
        big.setPointSize(13)
        big.setBold(True)
        self.btn_open_assistant.setFont(big)
        self.btn_open_assistant.clicked.connect(self.on_open_assistant)
        v.addWidget(self.btn_open_assistant)

        self.lbl_overall_status = QLabel("Trạng thái: Sẵn sàng.")
        self.lbl_overall_status.setStyleSheet("color: #0b5394; font-weight: bold;")
        v.addWidget(self.lbl_overall_status)
        return box

    # ---- Vùng 3: File output hiện tại ------------------------------- #
    def _build_current_file_group(self) -> QGroupBox:
        box = QGroupBox("3. File output hiện tại")
        grid = QGridLayout(box)

        self.lbl_file_path = QLabel("—")
        self.lbl_file_path.setWordWrap(True)
        self.lbl_file_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_file_status = QLabel("—")
        self.lbl_file_detected = QLabel("—")
        self.lbl_file_note = QLabel("—")
        self.lbl_file_note.setWordWrap(True)
        self.lbl_file_note.setStyleSheet("color: #a94442;")

        grid.addWidget(QLabel("Đường dẫn file output:"), 0, 0)
        grid.addWidget(self.lbl_file_path, 0, 1)
        grid.addWidget(QLabel("Trạng thái file:"), 1, 0)
        grid.addWidget(self.lbl_file_status, 1, 1)
        grid.addWidget(QLabel("Thời gian phát hiện:"), 2, 0)
        grid.addWidget(self.lbl_file_detected, 2, 1)
        grid.addWidget(QLabel("Ghi chú / lỗi:"), 3, 0)
        grid.addWidget(self.lbl_file_note, 3, 1)

        btn_row = QHBoxLayout()
        self.btn_open_result = QPushButton("Mở file kết quả")
        self.btn_open_result.clicked.connect(self.on_open_result_file)
        self.btn_confirm_review = QPushButton("Đã kiểm tra và dùng file này")
        self.btn_confirm_review.clicked.connect(self.on_confirm_reviewed)
        self.btn_open_containing = QPushButton("Mở thư mục chứa file")
        self.btn_open_containing.clicked.connect(self.on_open_containing_folder)
        btn_row.addWidget(self.btn_open_result)
        btn_row.addWidget(self.btn_confirm_review)
        btn_row.addWidget(self.btn_open_containing)
        btn_row.addStretch(1)
        grid.addLayout(btn_row, 4, 0, 1, 2)

        return box

    # ---- Vùng 4: Duyệt dữ liệu -------------------------------------- #
    def _build_preview_group(self) -> QGroupBox:
        box = QGroupBox("4. Duyệt dữ liệu từ file output")
        v = QVBoxLayout(box)

        self.btn_preview = QPushButton("Duyệt dữ liệu từ file output")
        self.btn_preview.setMinimumHeight(40)
        self.btn_preview.setEnabled(False)  # chỉ bật sau khi xác nhận file
        self.btn_preview.clicked.connect(self.on_preview_data)
        v.addWidget(self.btn_preview)

        self.lbl_preview_info = QLabel("Chưa duyệt dữ liệu.")
        v.addWidget(self.lbl_preview_info)

        self.table_preview = QTableWidget(0, 0)
        self.table_preview.setMinimumHeight(200)
        self.table_preview.horizontalHeader().setStretchLastSection(True)
        v.addWidget(self.table_preview)

        return box

    # ---- Vùng 5: Lịch sử / Log -------------------------------------- #
    def _build_history_group(self) -> QGroupBox:
        box = QGroupBox("5. Lịch sử & Nhật ký")
        v = QVBoxLayout(box)

        splitter = QSplitter(Qt.Vertical)

        self.table_history = QTableWidget(0, 6)
        self.table_history.setHorizontalHeaderLabels(
            ["ID", "Thời gian", "Tên file", "Đường dẫn working", "Trạng thái", "Ghi chú"]
        )
        self.table_history.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self.table_history.horizontalHeader().setStretchLastSection(True)
        self.table_history.setEditTriggers(QTableWidget.NoEditTriggers)
        splitter.addWidget(self.table_history)

        log_wrap = QWidget()
        log_layout = QVBoxLayout(log_wrap)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.addWidget(QLabel("Nhật ký hoạt động:"))
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumBlockCount(2000)
        log_layout.addWidget(self.txt_log)
        splitter.addWidget(log_wrap)

        splitter.setSizes([260, 180])
        v.addWidget(splitter)
        return box

    # ================================================================== #
    # Khởi tạo phụ trợ
    # ================================================================== #
    def _setup_log_bridge(self) -> None:
        self._log_emitter = _LogEmitter()
        self._log_emitter.message.connect(self.append_log)
        handler = _QtLogHandler(self._log_emitter)
        self.logger.addHandler(handler)

    def _start_watcher(self) -> None:
        try:
            self.watcher.start()
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không khởi động được watcher.")
            self.show_error(
                "Không theo dõi được thư mục download",
                f"Chi tiết: {exc}",
            )

    def _load_last_record(self) -> None:
        """Nạp lại bản ghi mới nhất khi mở app (nếu file working còn tồn tại)."""
        try:
            record = self.database.get_latest_file()
        except Exception:  # noqa: BLE001
            self.logger.exception("Không đọc được bản ghi mới nhất.")
            record = None

        self.refresh_history()

        if not record:
            return

        working_path = record.get("working_path")
        if working_path and os.path.exists(working_path):
            self.current_working_file = {
                "id": record.get("id"),
                "working_path": working_path,
                "backup_path": record.get("backup_path"),
                "original_download_path": record.get("original_download_path"),
                "file_name": record.get("file_name"),
                "file_size": record.get("file_size"),
                "file_hash": record.get("file_hash"),
                "status": record.get("status"),
                "detected_at": record.get("created_at"),
            }
            self._update_current_file_labels(
                note=record.get("note") or "Đã khôi phục từ phiên làm việc trước."
            )
            # Nếu trước đó người dùng đã xác nhận thì cho phép duyệt lại.
            if record.get("status") in CONFIRMED_STATUSES:
                self.current_reviewed_file = working_path
                self.btn_preview.setEnabled(True)
            self.append_log(
                f"Khôi phục file gần nhất: {record.get('file_name')}"
            )
        else:
            self.append_log(
                "Bản ghi gần nhất không còn file working trên ổ đĩa."
            )

    # ================================================================== #
    # Tiện ích UI
    # ================================================================== #
    def append_log(self, text: str) -> None:
        self.txt_log.appendPlainText(text)

    def show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def show_info(self, title: str, message: str) -> None:
        QMessageBox.information(self, title, message)

    def show_warning(self, title: str, message: str) -> None:
        QMessageBox.warning(self, title, message)

    def set_overall_status(self, status: str, message: str) -> None:
        self.overall_status = status
        self.lbl_overall_status.setText(f"Trạng thái: {message}")

    def _update_current_file_labels(self, note: Optional[str] = None) -> None:
        cw = self.current_working_file
        if not cw:
            self.lbl_file_path.setText("—")
            self.lbl_file_status.setText("—")
            self.lbl_file_detected.setText("—")
            self.lbl_file_note.setText("—")
            return
        self.lbl_file_path.setText(cw.get("working_path") or "—")
        self.lbl_file_status.setText(status_text(cw.get("status")))
        self.lbl_file_detected.setText(cw.get("detected_at") or "—")
        if note is not None:
            self.lbl_file_note.setText(note)

    def refresh_history(self) -> None:
        try:
            records = self.database.get_all_processed_files(limit=200)
        except Exception:  # noqa: BLE001
            self.logger.exception("Không đọc được lịch sử file.")
            return

        self.table_history.setRowCount(len(records))
        for r, rec in enumerate(records):
            values = [
                str(rec.get("id", "")),
                rec.get("created_at", "") or "",
                rec.get("file_name", "") or "",
                rec.get("working_path", "") or "",
                status_text(rec.get("status")),
                rec.get("note", "") or "",
            ]
            for c, val in enumerate(values):
                self.table_history.setItem(r, c, QTableWidgetItem(val))
        self.table_history.resizeColumnsToContents()

    def _update_record_status(
        self,
        status: str,
        note: Optional[str] = None,
        mark_reviewed: bool = False,
        mark_previewed: bool = False,
    ) -> None:
        """Cập nhật DB + nhãn cho file hiện tại."""
        if not self.current_working_file:
            return
        record_id = self.current_working_file.get("id")
        if record_id is not None:
            try:
                self.database.update_status(
                    record_id,
                    status,
                    note=note,
                    mark_reviewed=mark_reviewed,
                    mark_previewed=mark_previewed,
                )
            except Exception:  # noqa: BLE001
                self.logger.exception("Không cập nhật được trạng thái DB.")
        self.current_working_file["status"] = status
        self._update_current_file_labels(note=note)
        self.refresh_history()

    # ================================================================== #
    # Vùng 1: handler cấu hình
    # ================================================================== #
    def _browse_bat(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file .bat", self.edit_bat.text(), "Batch (*.bat);;Tất cả (*.*)"
        )
        if path:
            self.edit_bat.setText(os.path.normpath(path))

    def _browse_download(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục download", self.edit_download.text()
        )
        if path:
            self.edit_download.setText(os.path.normpath(path))

    def _browse_output(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục output", self.edit_output.text()
        )
        if path:
            self.edit_output.setText(os.path.normpath(path))

    def _browse_backup(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục backup", self.edit_backup.text()
        )
        if path:
            self.edit_backup.setText(os.path.normpath(path))

    def _browse_daily(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file theo dõi hàng ngày",
            self.edit_daily.text(),
            "Excel (*.xlsx *.xlsm);;Tất cả (*.*)",
        )
        if path:
            self.edit_daily.setText(os.path.normpath(path))

    def on_save_config(self) -> None:
        try:
            self.config.set("bat_path", self.edit_bat.text().strip())
            self.config.set("download_folder", self.edit_download.text().strip())
            self.config.set("output_folder", self.edit_output.text().strip())
            self.config.set("backup_folder", self.edit_backup.text().strip())
            self.config.set("daily_tracking_file", self.edit_daily.text().strip())
            self.config.ensure_folders()
            self.config.save()
            self.logger.info("Đã lưu cấu hình vào %s", self.config.path)

            # Khởi động lại watcher để áp dụng thư mục download mới.
            self.watcher.stop()
            self.watcher = DownloadWatcher(self.config, self.database, self.logger)
            self.watcher.signals.output_ready.connect(self.on_output_ready)
            self.watcher.signals.file_error.connect(self.on_file_error)
            self.watcher.signals.log_message.connect(self.append_log)
            self._start_watcher()

            self.show_info("Lưu cấu hình", "Đã lưu cấu hình thành công.")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi lưu cấu hình.")
            self.show_error("Lỗi lưu cấu hình", f"Chi tiết: {exc}")

    def on_check_config(self) -> None:
        lines = []

        bat = self.edit_bat.text().strip()
        lines.append(
            f"• File .bat: {'OK' if os.path.isfile(bat) else 'KHÔNG TỒN TẠI'}\n  {bat}"
        )
        dl = self.edit_download.text().strip()
        lines.append(
            f"• Thư mục download: {'OK' if os.path.isdir(dl) else 'KHÔNG TỒN TẠI'}\n  {dl}"
        )
        out = self.edit_output.text().strip()
        lines.append(
            f"• Thư mục output: {'OK' if os.path.isdir(out) else 'KHÔNG TỒN TẠI'}\n  {out}"
        )
        bak = self.edit_backup.text().strip()
        lines.append(
            f"• Thư mục backup: {'OK' if os.path.isdir(bak) else 'KHÔNG TỒN TẠI'}\n  {bak}"
        )
        daily = self.edit_daily.text().strip()
        daily_dir = os.path.dirname(daily)
        daily_ok = "OK" if os.path.isfile(daily) else (
            "CHƯA CÓ FILE (thư mục OK)" if os.path.isdir(daily_dir) else "THIẾU THƯ MỤC"
        )
        lines.append(f"• File theo dõi hàng ngày: {daily_ok}\n  {daily}")

        self.logger.info("Người dùng kiểm tra cấu hình.")
        self.show_info("Kiểm tra cấu hình", "\n".join(lines))

    def on_open_output_folder(self) -> None:
        folder = self.edit_output.text().strip()
        try:
            os.makedirs(folder, exist_ok=True)
            file_utils.open_folder(folder)
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được thư mục output.")
            self.show_error("Lỗi", f"Không mở được thư mục output.\nChi tiết: {exc}")

    # ================================================================== #
    # Vùng 2: mở trợ lý
    # ================================================================== #
    def on_open_assistant(self) -> None:
        bat_path = self.config.bat_path
        if not bat_path or not os.path.isfile(bat_path):
            self.logger.error("Không tìm thấy file .bat: %s", bat_path)
            self.show_error(
                "Không tìm thấy file .bat",
                "Đường dẫn file .bat không tồn tại. Vui lòng kiểm tra lại cấu hình.\n"
                f"Đường dẫn hiện tại: {bat_path}",
            )
            return
        try:
            subprocess.Popen(
                f'"{bat_path}"',
                shell=True,
                cwd=os.path.dirname(bat_path) or None,
            )
            self.set_overall_status(
                "WAITING_FOR_DOWNLOAD",
                "Đã mở trợ lý quyết toán. Đang chờ tải file output...",
            )
            self.logger.info("Đã mở trợ lý quyết toán qua file .bat: %s", bat_path)
            self.append_log("Đã mở trợ lý quyết toán. Đang chờ tải file output...")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi mở trợ lý quyết toán.")
            self.show_error(
                "Lỗi mở trợ lý",
                f"Không chạy được file .bat.\nChi tiết: {exc}",
            )

    # ================================================================== #
    # Vùng 3: file output hiện tại
    # ================================================================== #
    def on_output_ready(self, result: Dict[str, Any]) -> None:
        """Nhận tín hiệu từ watcher khi có file output mới."""
        self.current_working_file = result
        # File mới -> phải xác nhận lại trước khi duyệt.
        self.current_reviewed_file = None
        self.btn_preview.setEnabled(False)

        self._update_current_file_labels(
            note="Trợ lý đã bóc tách dữ liệu xong. Vui lòng mở file để kiểm tra."
        )
        self.set_overall_status(
            "READY_FOR_REVIEW",
            "Trợ lý đã bóc tách dữ liệu xong. Vui lòng mở file để kiểm tra.",
        )
        self.refresh_history()
        self.append_log(
            f"File output mới sẵn sàng: {result.get('file_name')}"
        )

        if self.config.auto_open_after_download:
            try:
                file_utils.open_file(result["working_path"])
                self._update_record_status(
                    "OPENED_FOR_REVIEW",
                    note="Đã tự động mở file để kiểm tra.",
                )
            except Exception:  # noqa: BLE001
                self.logger.exception("Không tự mở được file output.")

    def on_file_error(self, message: str) -> None:
        """Nhận tín hiệu lỗi từ watcher (không dùng hộp thoại chặn)."""
        self.lbl_file_note.setText(message)
        self.set_overall_status("ERROR", message)
        self.append_log("LỖI: " + message)

    def on_open_result_file(self) -> None:
        if not self.current_working_file:
            self.show_warning("Chưa có file", "Hiện chưa có file output nào để mở.")
            return
        path = self.current_working_file.get("working_path")
        if not path or not os.path.exists(path):
            self.show_error(
                "File không tồn tại",
                "File output không còn tồn tại trên ổ đĩa.",
            )
            return
        try:
            file_utils.open_file(path)
            self._update_record_status(
                "OPENED_FOR_REVIEW",
                note="Đã mở file để kiểm tra/chỉnh sửa.",
            )
            self.logger.info("Đã mở file kết quả: %s", path)
            self.append_log(f"Đã mở file kết quả: {os.path.basename(path)}")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được file kết quả.")
            self.show_error("Lỗi mở file", f"Không mở được file.\nChi tiết: {exc}")

    def on_confirm_reviewed(self) -> None:
        if not self.current_working_file:
            self.show_warning(
                "Chưa có file", "Chưa có file output hiện tại để xác nhận."
            )
            return
        path = self.current_working_file.get("working_path")
        if not path or not os.path.exists(path):
            self.show_error(
                "File không tồn tại", "File output không còn tồn tại trên ổ đĩa."
            )
            return

        # Kiểm tra file có đang bị Excel khóa / còn file tạm ~$ hay không.
        try:
            locked = file_utils.is_file_locked(path)
        except FileNotFoundError:
            self.show_error(
                "File không tồn tại", "File output không còn tồn tại trên ổ đĩa."
            )
            return
        if locked:
            self.show_warning(
                "File đang mở",
                "File vẫn đang mở trong Excel, vui lòng lưu và đóng file trước.",
            )
            return

        self.current_reviewed_file = path
        self.btn_preview.setEnabled(True)
        self._update_record_status(
            "REVIEW_CONFIRMED",
            note="Người dùng đã kiểm tra và chọn dùng file này.",
            mark_reviewed=True,
        )
        self.set_overall_status(
            "REVIEW_CONFIRMED", "Đã xác nhận file. Có thể duyệt dữ liệu."
        )
        self.logger.info("Xác nhận dùng file: %s", path)
        self.append_log(f"Đã xác nhận dùng file: {os.path.basename(path)}")
        self.show_info(
            "Đã xác nhận",
            "Đã ghi nhận file này. Bạn có thể bấm 'Duyệt dữ liệu từ file output'.",
        )

    def on_open_containing_folder(self) -> None:
        if not self.current_working_file:
            self.show_warning("Chưa có file", "Hiện chưa có file output nào.")
            return
        path = self.current_working_file.get("working_path")
        try:
            file_utils.open_folder(path)
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được thư mục chứa file.")
            self.show_error("Lỗi", f"Không mở được thư mục.\nChi tiết: {exc}")

    # ================================================================== #
    # Vùng 4: duyệt dữ liệu
    # ================================================================== #
    def on_preview_data(self) -> None:
        if not self.current_reviewed_file:
            self.show_warning(
                "Chưa xác nhận file",
                "Bạn cần bấm 'Đã kiểm tra và dùng file này' trước khi duyệt dữ liệu.",
            )
            return
        path = self.current_reviewed_file
        if not os.path.exists(path):
            self.show_error(
                "File không tồn tại", "File đã chọn không còn tồn tại trên ổ đĩa."
            )
            return

        try:
            locked = file_utils.is_file_locked(path)
        except FileNotFoundError:
            self.show_error(
                "File không tồn tại", "File đã chọn không còn tồn tại trên ổ đĩa."
            )
            return
        if locked:
            self.show_warning(
                "File đang mở",
                "File vẫn đang mở trong Excel, vui lòng lưu và đóng file trước.",
            )
            return

        try:
            data = excel_preview.preview_file(path, max_rows=20)
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi đọc dữ liệu preview.")
            self._update_record_status("ERROR", note=f"Lỗi đọc file: {exc}")
            self.show_error("Lỗi đọc dữ liệu", f"Không đọc được file.\nChi tiết: {exc}")
            return

        self._populate_preview_table(data)
        sheets = ", ".join(data.get("sheets") or []) or "—"
        self.lbl_preview_info.setText(
            f"Sheet: [{sheets}]  |  Đang xem: {data.get('sheet_name')}  |  "
            f"Số dòng: {data.get('row_count')}  |  Số cột: {data.get('column_count')}"
        )
        self._update_record_status(
            "PREVIEWED",
            note="Đã duyệt (xem trước) dữ liệu từ file output.",
            mark_previewed=True,
        )
        self.set_overall_status("PREVIEWED", "Đã duyệt dữ liệu file output.")
        self.logger.info(
            "Đã duyệt dữ liệu file %s (%s dòng, %s cột).",
            os.path.basename(path),
            data.get("row_count"),
            data.get("column_count"),
        )
        self.append_log(f"Đã duyệt dữ liệu: {os.path.basename(path)}")

    def _populate_preview_table(self, data: Dict[str, Any]) -> None:
        rows = data.get("rows") or []
        col_count = data.get("column_count") or (max((len(r) for r in rows), default=0))
        col_count = max(col_count, max((len(r) for r in rows), default=0))

        self.table_preview.clear()
        self.table_preview.setRowCount(len(rows))
        self.table_preview.setColumnCount(col_count)
        self.table_preview.setHorizontalHeaderLabels(
            [f"Cột {i + 1}" for i in range(col_count)]
        )
        for r, row in enumerate(rows):
            for c in range(col_count):
                val = row[c] if c < len(row) else ""
                self.table_preview.setItem(r, c, QTableWidgetItem(str(val)))
        self.table_preview.resizeColumnsToContents()

    # ================================================================== #
    # Đóng cửa sổ
    # ================================================================== #
    def closeEvent(self, event) -> None:  # noqa: N802 - override Qt
        try:
            self.watcher.stop()
        except Exception:  # noqa: BLE001
            pass
        self.logger.info("==== Đóng ứng dụng ====")
        super().closeEvent(event)
