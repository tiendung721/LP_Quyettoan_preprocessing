"""Giao diện chính (PySide6) của Trợ Lý Quyết Toán RPA.

Bố cục 2 mục điều hướng bên trái:
    - "Chức năng": luồng làm việc 2 bước cho người dùng phổ thông (mở trợ lý ->
      file tự mở trong Excel kèm popup hướng dẫn -> đánh dấu đã kiểm tra xong).
    - "Cài đặt": đường dẫn/thư mục, lịch sử xử lý và nhật ký.

Toàn bộ logic nghiệp vụ (watcher, database, đọc excel, thao tác file) giữ
nguyên; file này chỉ lo phần hiển thị và nối các handler với giao diện.
"""

from __future__ import annotations

import logging
import os
import subprocess
from typing import Any, Dict, Optional

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import file_utils
from .config import AppConfig
from .database import Database
from .watcher import DownloadWatcher

# Nhãn tiếng Việt cho từng trạng thái (không kèm mã kỹ thuật cho dễ hiểu).
STATUS_LABELS = {
    "WAITING_FOR_DOWNLOAD": "Đang chờ tải file output",
    "DOWNLOADED": "Đã tải về",
    "READY_FOR_REVIEW": "Sẵn sàng để kiểm tra",
    "OPENED_FOR_REVIEW": "Đang mở để kiểm tra",
    "REVIEW_SAVED": "Đã lưu sau khi chỉnh sửa",
    "COMPLETED": "Đã kiểm tra & hoàn tất",
    # Các trạng thái cũ (giữ để hiển thị lại lịch sử từ phiên bản trước).
    "REVIEW_CONFIRMED": "Đã xác nhận dùng file này",
    "READY_TO_PREVIEW": "Sẵn sàng duyệt dữ liệu",
    "PREVIEWED": "Đã duyệt dữ liệu",
    "ERROR": "Lỗi",
}


def friendly_status(status: Optional[str]) -> str:
    """Trả về nhãn tiếng Việt thuần cho trạng thái (không hiện mã)."""
    if not status:
        return "—"
    return STATUS_LABELS.get(status, status)


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

        self.setWindowTitle("Trợ Lý Quyết Toán RPA")
        self.resize(1120, 800)
        self.setMinimumSize(960, 640)

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
        central.setObjectName("appRoot")
        self.setCentralWidget(central)

        row = QHBoxLayout(central)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(0)

        self.stack = QStackedWidget()
        self.stack.setObjectName("content")

        sidebar = self._build_sidebar()
        self.stack.addWidget(self._build_function_page())
        self.stack.addWidget(self._build_settings_page())

        row.addWidget(sidebar)
        row.addWidget(self.stack, stretch=1)

        self.nav.setCurrentRow(0)
        self._update_current_file_labels()

    # ---- Thanh điều hướng bên trái ---------------------------------- #
    def _build_sidebar(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("sidebar")
        panel.setFixedWidth(214)

        v = QVBoxLayout(panel)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        title = QLabel("Trợ Lý Quyết Toán")
        title.setObjectName("appTitle")
        title.setWordWrap(True)
        subtitle = QLabel("RPA hỗ trợ quyết toán")
        subtitle.setObjectName("appSubtitle")
        v.addWidget(title)
        v.addWidget(subtitle)

        self.nav = QListWidget()
        self.nav.setObjectName("navList")
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        QListWidgetItem("Chức năng", self.nav)
        QListWidgetItem("Cài đặt", self.nav)
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        v.addWidget(self.nav)

        v.addStretch(1)
        footer = QLabel("Phiên bản 1.0")
        footer.setObjectName("navFooter")
        v.addWidget(footer)

        return panel

    # ---- Tiện ích dựng UI ------------------------------------------- #
    def _card(self, title: str) -> QFrame:
        frame = QFrame()
        frame.setObjectName("card")
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(18, 16, 18, 18)
        lay.setSpacing(10)
        lbl = QLabel(title)
        lbl.setObjectName("cardTitle")
        lay.addWidget(lbl)
        return frame

    def _make_button(
        self,
        text: str,
        handler,
        variant: Optional[str] = None,
        big: bool = False,
    ) -> QPushButton:
        btn = QPushButton(text)
        btn.setCursor(Qt.PointingHandCursor)
        if handler is not None:
            btn.clicked.connect(handler)
        if variant:
            btn.setProperty("variant", variant)
        if big:
            btn.setObjectName("bigButton")
        return btn

    def _scroll_page(self) -> tuple[QScrollArea, QVBoxLayout]:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        page = QWidget()
        page.setObjectName("page")
        scroll.setWidget(page)
        v = QVBoxLayout(page)
        v.setContentsMargins(24, 24, 24, 24)
        v.setSpacing(16)
        return scroll, v

    # ---- Trang "Chức năng" ------------------------------------------ #
    def _build_function_page(self) -> QScrollArea:
        scroll, v = self._scroll_page()

        # --- Bước 1: mở trợ lý ---
        card1 = self._card("1  ·  Mở trợ lý quyết toán")
        hint1 = QLabel(
            "Bấm nút bên dưới để khởi động trợ lý, gửi dữ liệu lên GPT rồi tải "
            "file output về. Phần mềm sẽ tự phát hiện file mới."
        )
        hint1.setObjectName("cardHint")
        hint1.setWordWrap(True)
        card1.layout().addWidget(hint1)

        self.btn_open_assistant = self._make_button(
            "Mở trợ lý quyết toán",
            self.on_open_assistant,
            variant="primary",
            big=True,
        )
        card1.layout().addWidget(self.btn_open_assistant)

        self.lbl_overall_status = QLabel("Sẵn sàng.")
        self.lbl_overall_status.setObjectName("statusText")
        self.lbl_overall_status.setWordWrap(True)
        card1.layout().addWidget(self.lbl_overall_status)
        v.addWidget(card1)

        # --- Bước 2: kiểm tra & hoàn tất ---
        card2 = self._card("2  ·  Kiểm tra & hoàn tất file kết quả")
        hint2 = QLabel(
            "Bạn có thể tải file mới từ trợ lý hoặc chọn một file có sẵn mà không "
            "cần chạy Bước 1. Hãy kiểm tra, lưu và đóng file, rồi bấm “Đã kiểm tra "
            "xong”. Phần mềm sẽ sao chép một bản vào Output và giữ nguyên file gốc."
        )
        hint2.setObjectName("cardHint")
        hint2.setWordWrap(True)
        card2.layout().addWidget(hint2)
        self.lbl_file_name = QLabel("Chưa có file kết quả")
        self.lbl_file_name.setObjectName("fileName")
        self.lbl_file_name.setWordWrap(True)
        self.lbl_file_name.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_file_status = QLabel("—")
        self.lbl_file_status.setObjectName("metaText")
        self.lbl_file_status.setWordWrap(True)
        self.lbl_file_detected = QLabel("")
        self.lbl_file_detected.setObjectName("metaText")
        self.lbl_file_note = QLabel("")
        self.lbl_file_note.setObjectName("noteText")
        self.lbl_file_note.setWordWrap(True)

        card2.layout().addWidget(self.lbl_file_name)
        card2.layout().addWidget(self.lbl_file_status)
        card2.layout().addWidget(self.lbl_file_detected)
        card2.layout().addWidget(self.lbl_file_note)

        row2 = QHBoxLayout()
        self.btn_select_existing = self._make_button(
            "Chọn file có sẵn",
            self.on_select_existing_file,
        )
        self.btn_open_result = self._make_button(
            "Mở lại file",
            self.on_open_result_file,
        )
        self.btn_open_containing = self._make_button(
            "Mở thư mục chứa file",
            self.on_open_containing_folder,
        )
        self.btn_mark_done = self._make_button(
            "Đã kiểm tra xong",
            self.on_mark_done,
            variant="success",
        )
        # Nút phụ ở bên trái, nút "Đã kiểm tra xong" đẩy ra ngoài cùng bên phải.
        row2.addWidget(self.btn_select_existing)
        row2.addWidget(self.btn_open_result)
        row2.addWidget(self.btn_open_containing)
        row2.addStretch(1)
        row2.addWidget(self.btn_mark_done)
        card2.layout().addLayout(row2)
        v.addWidget(card2)

        v.addStretch(1)
        return scroll

    # ---- Trang "Cài đặt" -------------------------------------------- #
    def _build_settings_page(self) -> QScrollArea:
        scroll, v = self._scroll_page()

        # --- Đường dẫn & thư mục ---
        cfg = self._card("Đường dẫn & thư mục")
        cfg_hint = QLabel(
            "Cài đặt một lần khi bắt đầu dùng. Sau khi sửa, bấm “Lưu cấu hình”."
        )
        cfg_hint.setObjectName("cardHint")
        cfg_hint.setWordWrap(True)
        cfg.layout().addWidget(cfg_hint)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.edit_bat = QLineEdit(self.config.bat_path)
        self.edit_download = QLineEdit(self.config.download_folder)
        self.edit_output = QLineEdit(self.config.output_folder)
        self.edit_daily = QLineEdit(self.config.daily_tracking_file)

        rows = [
            ("File .bat mở trợ lý:", self.edit_bat, self._browse_bat),
            ("Thư mục download (file tải về):", self.edit_download, self._browse_download),
            ("Thư mục output (bản đã kiểm tra):", self.edit_output, self._browse_output),
            ("File theo dõi hàng ngày:", self.edit_daily, self._browse_daily),
        ]
        for r, (label, edit, handler) in enumerate(rows):
            lbl = QLabel(label)
            lbl.setObjectName("formLabel")
            grid.addWidget(lbl, r, 0)
            grid.addWidget(edit, r, 1)
            btn = self._make_button("Chọn…", handler)
            btn.setFixedWidth(96)
            grid.addWidget(btn, r, 2)
        grid.setColumnStretch(1, 1)
        cfg.layout().addLayout(grid)

        btn_row = QHBoxLayout()
        self.btn_save_config = self._make_button(
            "Lưu cấu hình",
            self.on_save_config,
            variant="primary",
        )
        self.btn_check_config = self._make_button(
            "Kiểm tra cấu hình",
            self.on_check_config,
        )
        self.btn_open_output = self._make_button(
            "Mở thư mục output",
            self.on_open_output_folder,
        )
        btn_row.addWidget(self.btn_save_config)
        btn_row.addWidget(self.btn_check_config)
        btn_row.addWidget(self.btn_open_output)
        btn_row.addStretch(1)
        cfg.layout().addLayout(btn_row)
        v.addWidget(cfg)

        # --- Lịch sử xử lý ---
        hist = self._card("Lịch sử xử lý")
        self.table_history = QTableWidget(0, 4)
        self.table_history.setHorizontalHeaderLabels(
            ["Thời gian", "Tên file", "Trạng thái", "Ghi chú"]
        )
        self.table_history.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self.table_history.horizontalHeader().setStretchLastSection(True)
        self.table_history.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_history.setAlternatingRowColors(True)
        self.table_history.setMinimumHeight(200)
        hist.layout().addWidget(self.table_history)
        v.addWidget(hist)

        # --- Nhật ký hoạt động ---
        logc = self._card("Nhật ký hoạt động")
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumBlockCount(2000)
        self.txt_log.setMinimumHeight(160)
        logc.layout().addWidget(self.txt_log)
        v.addWidget(logc)

        v.addStretch(0)
        return scroll

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
                "output_path": record.get("output_path"),
            }
            self._update_current_file_labels(
                note=record.get("note") or "Đã khôi phục từ phiên làm việc trước."
            )
            self.overall_status = record.get("status") or self.overall_status
            self._update_button_states()
            self.append_log(
                f"Khôi phục file gần nhất: {record.get('file_name')}"
            )
        else:
            self.append_log(
                "Bản ghi gần nhất không còn file trên ổ đĩa."
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
        self.lbl_overall_status.setText(message)

    def _update_button_states(self) -> None:
        """Bật/tắt các nút theo việc đã có file hay chưa."""
        cw = self.current_working_file
        has_file = bool(
            cw and os.path.exists(cw.get("working_path") or "")
        )
        self.btn_open_result.setEnabled(has_file)
        self.btn_mark_done.setEnabled(has_file)
        self.btn_open_containing.setEnabled(has_file)

    def _update_current_file_labels(self, note: Optional[str] = None) -> None:
        cw = self.current_working_file
        if not cw:
            self.lbl_file_name.setText("Chưa có file kết quả")
            self.lbl_file_name.setToolTip("")
            self.lbl_file_status.setText(
                "Hãy tải file từ trợ lý hoặc bấm “Chọn file có sẵn”."
            )
            self.lbl_file_detected.setText("")
            self.lbl_file_note.setText("")
            self._update_button_states()
            return

        working_path = cw.get("working_path") or ""
        self.lbl_file_name.setText(
            cw.get("file_name") or os.path.basename(working_path) or "—"
        )
        self.lbl_file_name.setToolTip(working_path)
        self.lbl_file_status.setText("Trạng thái: " + friendly_status(cw.get("status")))
        detected = cw.get("detected_at") or ""
        self.lbl_file_detected.setText(
            ("Phát hiện lúc: " + detected) if detected else ""
        )
        if note is not None:
            self.lbl_file_note.setText(note)
        self._update_button_states()

    def refresh_history(self) -> None:
        try:
            records = self.database.get_all_processed_files(limit=200)
        except Exception:  # noqa: BLE001
            self.logger.exception("Không đọc được lịch sử file.")
            return

        self.table_history.setRowCount(len(records))
        for r, rec in enumerate(records):
            values = [
                rec.get("created_at", "") or "",
                rec.get("file_name", "") or "",
                friendly_status(rec.get("status")),
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
        output_path: Optional[str] = None,
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
                    output_path=output_path,
                )
            except Exception:  # noqa: BLE001
                self.logger.exception("Không cập nhật được trạng thái DB.")
        self.current_working_file["status"] = status
        self._update_current_file_labels(note=note)
        self.refresh_history()

    # ================================================================== #
    # Cài đặt: handler cấu hình
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
    # Chức năng: mở trợ lý
    # ================================================================== #
    def on_open_assistant(self) -> None:
        bat_path = self.config.bat_path
        if not bat_path or not os.path.isfile(bat_path):
            self.logger.error("Không tìm thấy file .bat: %s", bat_path)
            self.show_error(
                "Không tìm thấy file .bat",
                "Đường dẫn file .bat không tồn tại. Vui lòng kiểm tra lại trong tab "
                "Cài đặt.\n"
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
    # Chức năng: file output hiện tại
    # ================================================================== #
    def on_select_existing_file(self) -> None:
        """Chọn file cũ để thực hiện trực tiếp Bước 2, không cần chạy Bước 1."""
        extensions = []
        for configured_ext in self.config.allowed_extensions or []:
            ext = str(configured_ext).strip().lower()
            if ext:
                extensions.append(ext if ext.startswith(".") else f".{ext}")
        extension_filter = " ".join(f"*{ext}" for ext in extensions)
        file_filter = (
            f"File kết quả ({extension_filter});;Tất cả (*.*)"
            if extension_filter
            else "Tất cả (*.*)"
        )
        start_dir = (
            self.config.download_folder
            if os.path.isdir(self.config.download_folder)
            else ""
        )
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file có sẵn để kiểm tra",
            start_dir,
            file_filter,
        )
        if not path:
            return

        path = os.path.normpath(path)
        ext = os.path.splitext(path)[1].lower()
        allowed = {item.lower() for item in extensions}
        if allowed and ext not in allowed:
            self.show_warning(
                "File không hợp lệ",
                "Định dạng file chưa được hỗ trợ. Các định dạng cho phép: "
                + ", ".join(extensions),
            )
            return
        if file_utils.is_temp_download_file(path):
            self.show_warning(
                "File chưa sẵn sàng",
                "Không thể chọn file tạm hoặc file đang tải dở.",
            )
            return

        try:
            result = file_utils.download_file_info(path)
            record_id = self.database.insert_processed_file(
                working_path=result["working_path"],
                status="READY_FOR_REVIEW",
                original_download_path=result["original_download_path"],
                backup_path=result["backup_path"],
                file_name=result["file_name"],
                file_size=result["file_size"],
                file_hash=result["file_hash"],
                note="Người dùng đã chọn file có sẵn để kiểm tra.",
            )
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không tiếp nhận được file có sẵn: %s", path)
            self.show_error(
                "Không chọn được file",
                f"Không thể tiếp nhận file đã chọn.\nChi tiết: {exc}",
            )
            return

        result["id"] = record_id
        result["status"] = "READY_FOR_REVIEW"
        self.current_working_file = result
        note = "Đã chọn file có sẵn. Hãy kiểm tra file rồi xác nhận hoàn tất."
        self._update_current_file_labels(note=note)
        self.set_overall_status("READY_FOR_REVIEW", note)
        self.refresh_history()
        self.logger.info("Đã chọn file có sẵn cho Bước 2: %s", path)
        self.append_log(f"Đã chọn file có sẵn: {os.path.basename(path)}")

    def on_output_ready(self, result: Dict[str, Any]) -> None:
        """Nhận tín hiệu từ watcher khi có file output mới.

        Luồng: file tải về giữ nguyên ở Downloads và tự mở trong Excel để người
        dùng kiểm tra. Khi bấm "Đã kiểm tra xong" mới sao chép sang Output.
        """
        self.current_working_file = result
        self._update_current_file_labels(
            note="Trợ lý đã bóc tách dữ liệu xong. Đang mở file để bạn kiểm tra..."
        )
        self.set_overall_status(
            "READY_FOR_REVIEW",
            "Trợ lý đã bóc tách dữ liệu xong. Đang mở file để bạn kiểm tra...",
        )
        self.refresh_history()
        self.append_log(f"File mới sẵn sàng: {result.get('file_name')}")
        self._auto_open_file()

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
            self.set_overall_status(
                "OPENED_FOR_REVIEW", "Đang mở file để kiểm tra/chỉnh sửa."
            )
            self.logger.info("Đã mở file kết quả: %s", path)
            self.append_log(f"Đã mở file kết quả: {os.path.basename(path)}")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được file kết quả.")
            self.show_error("Lỗi mở file", f"Không mở được file.\nChi tiết: {exc}")

    def _auto_open_file(self) -> None:
        """Tự mở file (trong Downloads) bằng Excel để người dùng kiểm tra."""
        cw = self.current_working_file
        if not cw:
            return
        path = cw.get("working_path") or ""
        try:
            file_utils.open_file(path)
            self._update_record_status(
                "OPENED_FOR_REVIEW",
                note="Đã tự động mở file để kiểm tra.",
            )
            self.set_overall_status(
                "OPENED_FOR_REVIEW",
                "Đã mở file để bạn kiểm tra. Chỉnh xong hãy lưu, đóng file rồi "
                "bấm “Đã kiểm tra xong”.",
            )
            self.logger.info("Đã tự mở file: %s", path)
            self.append_log(f"Đã tự mở file: {os.path.basename(path)}")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không tự mở được file.")
            self.lbl_file_note.setText(
                "Không tự mở được file. Hãy bấm 'Mở lại file' để mở thủ công. "
                f"Chi tiết: {exc}"
            )
            self.set_overall_status(
                "READY_FOR_REVIEW",
                "Không tự mở được file. Hãy bấm 'Mở lại file' để mở thủ công.",
            )

    def _mark_review_done(self) -> bool:
        """Sao chép file sang Output rồi đánh dấu đã kiểm tra xong.

        Trả về True nếu thành công.
        """
        cw = self.current_working_file
        if not cw:
            return False
        path = cw.get("working_path") or ""
        if not path or not os.path.exists(path):
            self.show_error(
                "File không tồn tại", "File tải về không còn tồn tại trên ổ đĩa."
            )
            return False

        # File phải được lưu & đóng thì bản sao sang Output mới đúng nội dung.
        try:
            locked = file_utils.is_file_locked(path)
        except FileNotFoundError:
            self.show_error(
                "File không tồn tại", "File tải về không còn tồn tại trên ổ đĩa."
            )
            return False
        if locked:
            self.show_warning(
                "File đang mở",
                "File vẫn đang mở trong Excel. Vui lòng lưu và đóng file trước "
                "khi bấm “Đã kiểm tra xong”.",
            )
            return False

        # Bản Output đã tạo trước đó cho CHÍNH file này (nếu có) -> sẽ thay thế,
        # không sinh thêm file mới.
        prev_output = cw.get("output_path")

        try:
            output_path = file_utils.copy_download_to_output(
                path, self.config.output_folder
            )
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không tạo được bản trong Output.")
            self.show_error(
                "Lỗi lưu Output",
                f"Không tạo được bản trong thư mục Output.\nChi tiết: {exc}",
            )
            return False

        # Xóa bản Output cũ của file này để chỉ giữ 1 bản (tên theo timestamp mới).
        replaced = False
        if (
            prev_output
            and os.path.exists(prev_output)
            and os.path.abspath(prev_output) != os.path.abspath(output_path)
            and os.path.abspath(prev_output) != os.path.abspath(path)
            and file_utils.is_path_within(prev_output, self.config.output_folder)
        ):
            try:
                os.remove(prev_output)
                replaced = True
            except OSError:
                self.logger.warning(
                    "Không xóa được bản Output cũ: %s", prev_output
                )

        cw["output_path"] = output_path
        base = os.path.basename(output_path)
        note = (
            f"Đã kiểm tra xong. {'Đã cập nhật' if replaced else 'Đã tạo'} "
            f"bản Output: {base}"
        )
        self._update_record_status(
            "COMPLETED",
            note=note,
            mark_reviewed=True,
            output_path=output_path,
        )
        self.set_overall_status(
            "COMPLETED",
            "Đã kiểm tra & hoàn tất. Đã lưu bản vào Output và giữ nguyên file gốc.",
        )
        self.logger.info(
            "%s bản Output: %s",
            "Đã cập nhật" if replaced else "Đã tạo",
            output_path,
        )
        self.append_log(
            f"{'Đã cập nhật' if replaced else 'Đã tạo'} bản Output: {base}"
        )
        return True

    def on_mark_done(self) -> None:
        if not self.current_working_file:
            self.show_warning(
                "Chưa có file", "Hiện chưa có file nào để đánh dấu hoàn tất."
            )
            return
        if self._mark_review_done():
            self.show_info(
                "Hoàn tất",
                "Đã lưu bản đã kiểm tra trong thư mục Output. File gốc vẫn được "
                "giữ nguyên tại thư mục ban đầu.",
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
    # Đóng cửa sổ
    # ================================================================== #
    def closeEvent(self, event) -> None:  # noqa: N802 - override Qt
        try:
            self.watcher.stop()
        except Exception:  # noqa: BLE001
            pass
        self.logger.info("==== Đóng ứng dụng ====")
        super().closeEvent(event)
