"""Giao diện chính (PySide6) của Trợ Lý Quyết Toán RPA.

Bố cục 2 mục điều hướng bên trái:
    - "Chức năng": luồng làm việc 4 bước, mỗi bước đúng một nút bấm (mở trợ lý ->
      xem file bóc tách -> nhập lên file hàng ngày -> nhập lên phần mềm quyết
      toán bằng luồng PAD RPA).
    - "Cài đặt": đường dẫn/thư mục, lịch sử xử lý và nhật ký.

Mọi thao tác phụ (chọn file, mở thư mục, xem dữ liệu chờ) đều được phần mềm tự
xử lý để người dùng chỉ phải bấm đúng các nút của luồng chính.
"""

from __future__ import annotations

import logging
import os
import subprocess
from datetime import datetime
from typing import Any, Dict, Optional

from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
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
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import file_utils
from .config import AppConfig
from .daily_import import EXPENSE_SHEET, INFO_SHEET, DailyImportError, ImportAnalysis, ImportSummary
from .daily_import import DailyImportService, remove_extract_rows
from .daily_import_ui import (
    ConflictDialog,
    FunctionWorker,
    JsonExtractDialog,
    UnmatchedRowsDialog,
)
from .database import Database
from .dialogs.select_sqt_dialog import SelectSqtDialog
from .services.rpa_launcher import RpaLaunchError, launch_bat, validate_bat_path
from .services.sqt_selection_service import (
    DailyFileLockedError,
    DailyFileNotFoundError,
    EmptySqtListError,
    EXPENSE_SELECTION_OPERATION,
    MissingColumnError,
    MissingSheetError,
    SelectionJsonWriteError,
    SELECTION_OPERATION,
    SqtSelectionError,
    read_sqt_items,
    write_selection_json,
)
from .watcher import OutputWatcher

# Nhãn tiếng Việt cho từng trạng thái (không kèm mã kỹ thuật cho dễ hiểu).
STATUS_LABELS = {
    "WAITING_FOR_DOWNLOAD": "Đang chờ file output",
    "DOWNLOADED": "Đã nhận file output",
    "READY_FOR_REVIEW": "Sẵn sàng để kiểm tra",
    "OPENED_FOR_REVIEW": "Đang mở để kiểm tra",
    "REVIEW_SAVED": "Đã lưu sau khi chỉnh sửa",
    "COMPLETED": "Đã kiểm tra & hoàn tất",
    "DAILY_IMPORT_PENDING": "Đang chờ nhập file theo dõi",
    "DAILY_IMPORT_PARTIAL": "Đã nhập một phần vào file theo dõi",
    "DAILY_IMPORTED": "Đã nhập vào file theo dõi",
    "DAILY_IMPORT_ERROR": "Lỗi nhập file theo dõi",
    "RPA_LAUNCHED": "Đã khởi chạy RPA quyết toán",
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
        self.daily_import_service = DailyImportService(
            self.database,
            self.logger,
            cargo_mappings=self.config.cargo_name_mappings,
        )
        self._daily_import_running = False
        self._daily_thread: Optional[QThread] = None
        self._daily_worker: Optional[FunctionWorker] = None
        self._daily_worker_result = None
        self._daily_worker_error: Optional[str] = None
        self._daily_worker_callback = None
        self._rpa_cooldown = False
        self._select_sqt_dialog_open = False
        self._rpa_launched_at: Optional[str] = None
        self._last_saved_shown = ""

        self.setWindowTitle("Trợ Lý Quyết Toán RPA")
        self.resize(1120, 800)
        self.setMinimumSize(960, 640)

        self._build_ui()
        self._setup_log_bridge()

        # Watcher theo dõi thư mục output.
        self.watcher = OutputWatcher(self.config, self.database, self.logger)
        self.watcher.signals.output_ready.connect(self.on_output_ready)
        self.watcher.signals.file_error.connect(self.on_file_error)
        self.watcher.signals.log_message.connect(self.append_log)

        self._load_last_record()
        self._start_watcher()

        # Người dùng bấm Lưu trong JSON editor nên soi lại thời điểm lưu của file
        # bóc tách để nhãn ở Bước 2 luôn đúng.
        self._saved_timer = QTimer(self)
        self._saved_timer.setInterval(2000)
        self._saved_timer.timeout.connect(self._refresh_saved_label)
        self._saved_timer.start()

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
        action: bool = False,
    ) -> QPushButton:
        """Tạo nút; ``action=True`` cho nút hành động chính của một bước."""
        btn = QPushButton(text)
        btn.setCursor(Qt.PointingHandCursor)
        if handler is not None:
            btn.clicked.connect(handler)
        if variant:
            btn.setProperty("variant", variant)
        if action:
            btn.setObjectName("actionButton")
            # Không kéo giãn hết chiều ngang thẻ: nút chỉ rộng vừa đủ nội dung và
            # không bị bóp lại khi nhãn bên trái dài.
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        return btn

    def _action_row(self, button: QPushButton, *left_widgets: QWidget) -> QHBoxLayout:
        """Hàng cuối thẻ: thông tin phụ bên trái, nút hành động dồn về bên phải.

        Các nhãn bên trái chỉ hiện thông tin ngắn (tên file, thời điểm lưu) nên
        không đẩy nút tràn khỏi thẻ; đường dẫn đầy đủ nằm ở tooltip.
        """
        row = QHBoxLayout()
        row.setSpacing(10)
        for widget in left_widgets:
            row.addWidget(widget, 0, Qt.AlignVCenter)
        row.addStretch(1)
        row.addWidget(button, 0, Qt.AlignVCenter)
        return row

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

    # ---- Chip trạng thái + thẻ theo bước + hướng dẫn thu gọn -------- #
    @staticmethod
    def _repolish(widget: QWidget) -> None:
        style = widget.style()
        style.unpolish(widget)
        style.polish(widget)
        widget.update()

    def _set_chip(self, chip: QLabel, text: str, tone: str) -> None:
        chip.setText(text)
        if chip.property("tone") != tone:
            chip.setProperty("tone", tone)
            self._repolish(chip)

    def _set_step(self, index: int, text: str, tone: str) -> None:
        """Đặt chip trạng thái + màu huy hiệu số cho một bước."""
        self._set_chip(getattr(self, f"chip{index}"), text, tone)
        badge = getattr(self, f"badge{index}")
        if badge.property("tone") != tone:
            badge.setProperty("tone", tone)
            self._repolish(badge)

    @staticmethod
    def _guide_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("guideText")
        lbl.setWordWrap(True)
        return lbl

    def _step_card(self, number: str, title: str):
        """Tạo thẻ 1 bước: huy hiệu số + tiêu đề + chip trạng thái + nút hướng dẫn.

        Trả về (frame, badge, chip, guide_layout, body_layout).
        """
        frame = QFrame()
        frame.setObjectName("card")
        frame.setProperty("active", "false")
        outer = QVBoxLayout(frame)
        outer.setContentsMargins(18, 14, 18, 16)
        outer.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(10)
        badge = QLabel(number)
        badge.setObjectName("stepBadge")
        badge.setProperty("tone", "wait")
        badge.setFixedSize(26, 26)
        badge.setAlignment(Qt.AlignCenter)
        header.addWidget(badge, 0, Qt.AlignVCenter)
        lbl = QLabel(title)
        lbl.setObjectName("cardTitle")
        lbl.setWordWrap(False)  # tiêu đề bước luôn nằm 1 dòng
        header.addWidget(lbl, 0, Qt.AlignVCenter)
        header.addStretch(1)
        chip = QLabel("")
        chip.setObjectName("statusChip")
        chip.setProperty("tone", "wait")
        header.addWidget(chip, 0, Qt.AlignVCenter)
        toggle = QPushButton("Hướng dẫn  ⌄")
        toggle.setObjectName("guideToggle")
        toggle.setCheckable(True)
        toggle.setCursor(Qt.PointingHandCursor)
        header.addWidget(toggle, 0, Qt.AlignVCenter)
        outer.addLayout(header)

        guide = QFrame()
        guide.setObjectName("guideBox")
        guide.setVisible(False)
        guide_layout = QVBoxLayout(guide)
        guide_layout.setContentsMargins(12, 10, 12, 10)
        guide_layout.setSpacing(4)
        outer.addWidget(guide)

        body = QVBoxLayout()
        body.setSpacing(8)
        outer.addLayout(body)

        toggle.toggled.connect(
            lambda checked, g=guide, t=toggle: self._on_guide_toggle(g, t, checked)
        )
        self._guide_toggles.append(toggle)
        return frame, badge, chip, guide_layout, body

    @staticmethod
    def _on_guide_toggle(guide: QFrame, toggle: QPushButton, checked: bool) -> None:
        guide.setVisible(checked)
        toggle.setText("Hướng dẫn  ⌃" if checked else "Hướng dẫn  ⌄")

    def _on_toggle_all_guides(self, checked: bool) -> None:
        for toggle in self._guide_toggles:
            if toggle.isChecked() != checked:
                toggle.setChecked(checked)

    def _refresh_step_states(self) -> None:
        """Cập nhật chip trạng thái + viền thẻ + dòng hint theo tiến độ."""
        cw = self.current_working_file or {}
        status = cw.get("status")
        has_file = bool(cw and os.path.exists(cw.get("working_path") or ""))
        daily_path = self.config.daily_tracking_file
        has_daily = bool(daily_path and os.path.isfile(daily_path))
        imported = status in {"DAILY_IMPORTED", "DAILY_IMPORT_PARTIAL", "RPA_LAUNCHED"}
        is_error = self.overall_status == "ERROR"

        # Bước 1 — mở trợ lý
        if has_file:
            self._set_step(1, "Đã có file output", "done")
        else:
            self._set_step(1, "Bắt đầu tại đây", "active")

        # Bước 2 — file bóc tách
        if not has_file:
            self._set_step(2, "Chờ file output", "wait")
        elif is_error:
            self._set_step(2, "Lỗi", "error")
        else:
            self._set_step(2, "Đã có file", "done")

        # Bước 3 — nhập lên file hàng ngày
        if not has_file:
            self._set_step(3, "Chờ file bóc tách", "wait")
        elif not has_daily:
            self._set_step(3, "Thiếu file theo dõi", "warn")
        elif status == "DAILY_IMPORT_ERROR":
            self._set_step(3, "Lỗi", "error")
        elif status == "DAILY_IMPORT_PARTIAL":
            self._set_step(3, "Còn dữ liệu chờ", "warn")
        elif imported:
            self._set_step(3, "Đã nhập", "done")
        else:
            self._set_step(3, "Sẵn sàng", "active")

        # Bước 4 — chạy RPA lên phần mềm quyết toán (luôn dùng được)
        if self._rpa_launched_at:
            self._set_step(4, "Đã khởi chạy", "done")
        else:
            self._set_step(4, "Sẵn sàng", "active")

        # Viền nổi bước đang cần thao tác
        if not has_file:
            active = 1
        elif not imported:
            active = 3
        elif not self._rpa_launched_at:
            active = 4
        else:
            active = 0
        for idx, card in (
            (1, self.card_step1),
            (2, self.card_step2),
            (3, self.card_step3),
            (4, self.card_step4),
        ):
            want = "true" if idx == active else "false"
            if card.property("active") != want:
                card.setProperty("active", want)
                self._repolish(card)

        # Dòng hint Bước 3: chỉ đặt khi chưa có file hoặc sẵn-sàng-chưa-nhập
        # (không đè khi đang chạy hoặc đã có kết quả nhập).
        if not self._daily_import_running:
            if not has_file:
                self.lbl_daily_status.setText(
                    "Chờ file bóc tách từ trợ lý rồi mới nhập được."
                )
            elif not has_daily:
                self.lbl_daily_status.setText(
                    "Chưa tìm thấy file theo dõi hàng ngày — hãy trỏ đúng file trong "
                    "tab Cài đặt."
                )
            elif status not in (
                "DAILY_IMPORTED",
                "DAILY_IMPORT_PARTIAL",
                "DAILY_IMPORT_ERROR",
            ):
                self.lbl_daily_status.setText(
                    "Sẵn sàng: phần mềm sẽ lấy bản lưu mới nhất của file bóc tách."
                )

    # ---- Trang "Chức năng" ------------------------------------------ #
    def _build_function_page(self) -> QScrollArea:
        scroll, v = self._scroll_page()
        self._guide_toggles = []

        # Công tắc hiện/ẩn hướng dẫn cho toàn bộ các bước.
        top = QHBoxLayout()
        top.addStretch(1)
        self.chk_show_guide = QCheckBox("📖  Hiện hướng dẫn")
        self.chk_show_guide.setCursor(Qt.PointingHandCursor)
        self.chk_show_guide.toggled.connect(self._on_toggle_all_guides)
        top.addWidget(self.chk_show_guide)
        v.addLayout(top)

        # --- Bước 1: mở trợ lý ---
        self.card_step1, self.badge1, self.chip1, guide1, body1 = self._step_card(
            "1", "Mở trợ lý quyết toán"
        )
        guide1.addWidget(self._guide_label(
            "Bấm nút để khởi động trợ lý, gửi dữ liệu lên GPT rồi tải file bóc tách "
            "về. Phần mềm sẽ tự phát hiện file mới."
        ))
        self.lbl_overall_status = QLabel("Sẵn sàng.")
        self.lbl_overall_status.setObjectName("hintText")
        self.lbl_overall_status.setWordWrap(True)
        body1.addWidget(self.lbl_overall_status)
        self.btn_open_assistant = self._make_button(
            "Mở trợ lý",
            self.on_open_assistant,
            variant="primary",
            action=True,
        )
        body1.addLayout(self._action_row(self.btn_open_assistant))
        v.addWidget(self.card_step1)

        # --- Bước 2: xem file bóc tách dữ liệu ---
        self.card_step2, self.badge2, self.chip2, guide2, body2 = self._step_card(
            "2", "Xem file bóc tách dữ liệu"
        )
        guide2.addWidget(self._guide_label(
            "Khi trợ lý tải file JSON về, phần mềm tự mở bảng tổng quan cả lô bóc "
            "tách. Chọn một dòng rồi bấm “Sửa dòng” hoặc “Xóa dòng”; bấm “Lưu” là "
            "màn hình tự đóng và dòng “Lưu thành công lần cuối” bên dưới cập nhật "
            "theo. Dòng nào Bước 3 chưa nhập được sẽ được tô đỏ kèm lý do."
        ))
        self.lbl_file_name = QLabel("Chưa có file bóc tách")
        self.lbl_file_name.setObjectName("fileName")
        self.lbl_file_name.setWordWrap(True)
        self.lbl_file_name.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_file_note = QLabel("")
        self.lbl_file_note.setObjectName("hintText")
        self.lbl_file_note.setWordWrap(True)
        body2.addWidget(self.lbl_file_name)
        body2.addWidget(self.lbl_file_note)

        self.lbl_file_saved = QLabel("Lưu thành công lần cuối: —")
        self.lbl_file_saved.setObjectName("metaText")
        self.btn_view_extract = self._make_button(
            "Xem file bóc tách",
            self.on_view_extract_file,
            variant="primary",
            action=True,
        )
        body2.addLayout(self._action_row(self.btn_view_extract, self.lbl_file_saved))
        v.addWidget(self.card_step2)

        # --- Bước 3: nhập lên file hàng ngày ---
        self.card_step3, self.badge3, self.chip3, guide3, body3 = self._step_card(
            "3", "Nhập lên file hàng ngày"
        )
        guide3.addWidget(self._guide_label(
            "Phần mềm đọc thẳng file JSON bạn vừa xem ở Bước 2, ghép Phiếu cân, Bill "
            "và khoản chi rồi cập nhật file theo dõi hàng ngày. Dòng nào không ghép "
            "được sẽ không nhập: bạn chọn quay lại Bước 2 để sửa, hoặc bỏ các dòng "
            "đó và vẫn nhập phần còn lại."
        ))
        self.lbl_daily_status = QLabel("Chưa nhập dữ liệu vào file theo dõi.")
        self.lbl_daily_status.setObjectName("hintText")
        self.lbl_daily_status.setWordWrap(True)
        self.progress_daily = QProgressBar()
        self.progress_daily.setRange(0, 0)
        self.progress_daily.setTextVisible(False)
        self.progress_daily.setVisible(False)
        body3.addWidget(self.lbl_daily_status)
        body3.addWidget(self.progress_daily)

        self.lbl_daily_path = QLabel("")
        self.lbl_daily_path.setObjectName("metaText")
        self.lbl_daily_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.btn_import_daily = self._make_button(
            "Nhập lên file hàng ngày",
            self.on_import_daily,
            variant="primary",
            action=True,
        )
        body3.addLayout(self._action_row(self.btn_import_daily, self.lbl_daily_path))
        v.addWidget(self.card_step3)

        # --- Bước 4: nhập lên phần mềm quyết toán bằng RPA ---
        self.card_step4, self.badge4, self.chip4, guide4, body4 = self._step_card(
            "4", "Nhập lên phần mềm quyết toán"
        )
        guide4.addWidget(self._guide_label(
            "Tạo Số QT mới sẽ gọi riêng flow PAD tạo mới và không đọc Excel. Nhập thông tin sẽ đọc "
            "sheet Thong_Tin_Quyet_Toan; Nhập khoản chi sẽ đọc sheet Khoan_Chi. Cả hai đều cho "
            "chọn một hoặc nhiều SQT rồi ghi JSON tạm cho PAD."
        ))
        self.lbl_rpa_status = QLabel("Chưa chạy luồng RPA trong phiên này.")
        self.lbl_rpa_status.setObjectName("hintText")
        self.lbl_rpa_status.setWordWrap(True)
        body4.addWidget(self.lbl_rpa_status)

        self.lbl_create_new_path = QLabel("")
        self.lbl_create_new_path.setObjectName("metaText")
        self.btn_create_new_sqt = self._make_button(
            "Tạo 1 Số QT mới trên PM",
            self.on_create_new_sqt,
            variant="success",
            action=True,
        )
        body4.addLayout(self._action_row(self.btn_create_new_sqt, self.lbl_create_new_path))

        self.lbl_input_information_path = QLabel("")
        self.lbl_input_information_path.setObjectName("metaText")
        self.btn_input_information = self._make_button(
            "Nhập thông tin",
            self.on_input_information,
            variant="primary",
            action=True,
        )
        body4.addLayout(self._action_row(self.btn_input_information, self.lbl_input_information_path))

        self.lbl_input_expense_path = QLabel("")
        self.lbl_input_expense_path.setObjectName("metaText")
        self.btn_input_expense = self._make_button(
            "Nhập khoản chi",
            self.on_input_expense,
            variant="primary",
            action=True,
        )
        body4.addLayout(self._action_row(self.btn_input_expense, self.lbl_input_expense_path))
        v.addWidget(self.card_step4)

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
        self.edit_create_new_bat = QLineEdit(self.config.pad_create_new_bat_path)
        self.edit_input_information_bat = QLineEdit(
            self.config.pad_input_information_bat_path
        )
        self.edit_input_expense_bat = QLineEdit(self.config.pad_input_expense_bat_path)
        self.edit_output = QLineEdit(self.config.output_folder)
        self.edit_daily = QLineEdit(self.config.daily_tracking_file)
        self.edit_daily.setReadOnly(True)

        rows = [
            ("File .bat mở trợ lý:", self.edit_bat, self._browse_bat),
            (
                "File .bat tạo 1 Số QT mới:",
                self.edit_create_new_bat,
                self._browse_create_new_bat,
            ),
            (
                "File .bat nhập thông tin:",
                self.edit_input_information_bat,
                self._browse_input_information_bat,
            ),
            (
                "File .bat nhập khoản chi:",
                self.edit_input_expense_bat,
                self._browse_input_expense_bat,
            ),
            ("Thư mục output:", self.edit_output, self._browse_output),
            ("File theo dõi hàng ngày trong output:", self.edit_daily, None),
        ]
        for r, (label, edit, handler) in enumerate(rows):
            lbl = QLabel(label)
            lbl.setObjectName("formLabel")
            grid.addWidget(lbl, r, 0)
            grid.addWidget(edit, r, 1)
            if handler is not None:
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
                "Không theo dõi được thư mục output",
                f"Chi tiết: {exc}",
            )

    def _load_last_record(self) -> None:
        """Nạp lại file bóc tách đang dùng khi mở app.

        Ưu tiên bản ghi gần nhất trong database; nếu file đó không còn trên ổ
        đĩa thì tự nhận file bóc tách mới nhất trong thư mục output (người dùng
        không phải chọn tay).
        """
        try:
            record = self.database.get_latest_file()
        except Exception:  # noqa: BLE001
            self.logger.exception("Không đọc được bản ghi mới nhất.")
            record = None

        self.refresh_history()

        working_path = (record or {}).get("working_path")
        if record and working_path and os.path.exists(working_path):
            self._apply_record(
                record,
                note=record.get("note") or "Đã khôi phục từ phiên làm việc trước.",
            )
            self.append_log(f"Khôi phục file gần nhất: {record.get('file_name')}")
            return

        self.append_log("Chưa có file bóc tách nào đang mở dở. Đang tìm trong thư mục output...")
        self._adopt_latest_output()

    def _apply_record(self, record: Dict[str, Any], note: Optional[str] = None) -> None:
        """Đưa một bản ghi database thành file bóc tách đang làm việc."""
        self.current_working_file = {
            "id": record.get("id"),
            "working_path": record.get("working_path"),
            "backup_path": record.get("backup_path"),
            "original_download_path": record.get("original_download_path"),
            "file_name": record.get("file_name"),
            "file_size": record.get("file_size"),
            "file_hash": record.get("file_hash"),
            "status": record.get("status"),
            "detected_at": record.get("created_at"),
            "output_path": record.get("output_path"),
        }
        self.overall_status = record.get("status") or self.overall_status
        self._mark_current_file_handled()
        self._update_current_file_labels(note=note)

    def _adopt_latest_output(self) -> None:
        """Tự nhận file bóc tách mới nhất trong thư mục output (nếu có)."""
        path = file_utils.find_latest_output_file(
            self.config.output_folder,
            self.config.allowed_extensions,
            self.config.output_file_patterns,
        )
        if not path:
            self._update_current_file_labels()
            return

        record = self.database.get_file_by_working_path(path)
        if record:
            self._apply_record(
                record, note="Đã nhận lại file bóc tách gần nhất trong thư mục output."
            )
            self.append_log(f"Nhận lại file bóc tách: {os.path.basename(path)}")
            return

        try:
            result = file_utils.output_file_info(path)
            record_id = self.database.insert_processed_file(
                working_path=result["working_path"],
                status="READY_FOR_REVIEW",
                original_download_path=result["original_download_path"],
                backup_path=result["backup_path"],
                file_name=result["file_name"],
                file_size=result["file_size"],
                file_hash=result["file_hash"],
                note="Phần mềm tự nhận file bóc tách mới nhất trong thư mục output.",
            )
        except Exception:  # noqa: BLE001 - không có file cũng không được sập app
            self.logger.exception("Không tiếp nhận được file bóc tách: %s", path)
            self._update_current_file_labels()
            return

        result["id"] = record_id
        result["status"] = "READY_FOR_REVIEW"
        self.current_working_file = result
        self.overall_status = "READY_FOR_REVIEW"
        self._mark_current_file_handled()
        self._update_current_file_labels(
            note="Phần mềm tự nhận file bóc tách mới nhất trong thư mục output."
        )
        self.refresh_history()
        self.logger.info("Tự nhận file bóc tách: %s", path)
        self.append_log(f"Tự nhận file bóc tách: {os.path.basename(path)}")

    def _mark_current_file_handled(self) -> None:
        """Báo watcher biết file đang dùng không phải file mới.

        Người dùng chỉnh và lưu file ngay trong thư mục output, nên nếu không
        đánh dấu thì watchdog sẽ coi mỗi lần lưu là một file mới: ghi thêm bản
        ghi trùng và tự mở lại file đang chỉnh.
        """
        path = (self.current_working_file or {}).get("working_path")
        if path:
            self.watcher.mark_handled(path)

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
        self._refresh_step_states()

    def _update_button_states(self) -> None:
        """Bật/tắt các nút theo việc đã có file hay chưa.

        Không kiểm tra khóa file ở đây; việc kiểm tra file theo dõi đang mở được
        làm ngay lúc bấm nút, kèm thông báo hướng dẫn.
        """
        cw = self.current_working_file
        has_file = bool(cw and os.path.exists(cw.get("working_path") or ""))
        daily_path = self.config.daily_tracking_file

        self.btn_view_extract.setEnabled(has_file)
        self.btn_import_daily.setEnabled(
            has_file
            and bool(daily_path and os.path.isfile(daily_path))
            and not self._daily_import_running
        )
        # Bước 4 luôn dùng được (trừ vài giây chống bấm trùng); nếu chưa cấu
        # hình file .bat thì báo lỗi hướng dẫn ngay lúc bấm.
        self.btn_create_new_sqt.setEnabled(not self._rpa_cooldown)
        self.btn_input_information.setEnabled(
            not self._rpa_cooldown and not self._select_sqt_dialog_open
        )
        self.btn_input_expense.setEnabled(
            not self._rpa_cooldown and not self._select_sqt_dialog_open
        )

        self._refresh_daily_labels()
        self._refresh_step_states()

    def _refresh_daily_labels(self) -> None:
        """Cập nhật phần mô tả Bước 3 và Bước 4 mà không thay đổi dữ liệu."""
        daily_path = self.config.daily_tracking_file
        if daily_path:
            self.lbl_daily_path.setText("Ghi vào: " + os.path.basename(daily_path))
            self.lbl_daily_path.setToolTip(daily_path)
        else:
            self.lbl_daily_path.setText("Ghi vào: (chưa cấu hình file theo dõi)")
            self.lbl_daily_path.setToolTip("")

        create_bat = self.config.pad_create_new_bat_path
        if create_bat and os.path.isfile(create_bat):
            self.lbl_create_new_path.setText("Tạo mới: " + os.path.basename(create_bat))
            self.lbl_create_new_path.setToolTip(create_bat)
        else:
            self.lbl_create_new_path.setText("Chưa cấu hình BAT tạo mới (xem tab Cài đặt)")
            self.lbl_create_new_path.setToolTip("")

        input_bat = self.config.pad_input_information_bat_path
        if input_bat and os.path.isfile(input_bat):
            self.lbl_input_information_path.setText(
                "Nhập thông tin: " + os.path.basename(input_bat)
            )
            self.lbl_input_information_path.setToolTip(input_bat)
        else:
            self.lbl_input_information_path.setText(
                "Chưa cấu hình BAT nhập thông tin (xem tab Cài đặt)"
            )
            self.lbl_input_information_path.setToolTip("")

        expense_bat = self.config.pad_input_expense_bat_path
        if expense_bat and os.path.isfile(expense_bat):
            self.lbl_input_expense_path.setText(
                "Nhập khoản chi: " + os.path.basename(expense_bat)
            )
            self.lbl_input_expense_path.setToolTip(expense_bat)
        else:
            self.lbl_input_expense_path.setText(
                "Chưa cấu hình BAT nhập khoản chi (xem tab Cài đặt)"
            )
            self.lbl_input_expense_path.setToolTip("")

    def _refresh_saved_label(self) -> None:
        """Đồng bộ dòng “Lưu thành công lần cuối” với thời điểm lưu thật của file."""
        path = (self.current_working_file or {}).get("working_path") or ""
        text = file_utils.last_saved_text(path) if path else ""
        if text == self._last_saved_shown:
            return
        self._last_saved_shown = text
        self.lbl_file_saved.setText("Lưu thành công lần cuối: " + (text or "—"))
        self._update_button_states()

    def _set_daily_running(self, running: bool, message: str = "") -> None:
        self._daily_import_running = running
        self.progress_daily.setVisible(running)
        if message:
            self.lbl_daily_status.setText(message)
        self._update_button_states()

    def _update_current_file_labels(self, note: Optional[str] = None) -> None:
        cw = self.current_working_file
        if not cw:
            self.lbl_file_name.setText("Chưa có file bóc tách")
            self.lbl_file_name.setToolTip("")
            self.lbl_file_note.setText(
                "Hãy bấm “Mở trợ lý quyết toán” ở Bước 1 rồi tải file bóc tách về."
            )
            self._refresh_saved_label()
            self._update_button_states()
            return

        working_path = cw.get("working_path") or ""
        self.lbl_file_name.setText(
            cw.get("file_name") or os.path.basename(working_path) or "—"
        )
        self.lbl_file_name.setToolTip(working_path)
        if note is not None:
            self.lbl_file_note.setText(note)
        self._refresh_saved_label()
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

    def _browse_rpa_bat(self) -> None:
        self._browse_input_information_bat()

    def _browse_create_new_bat(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file .bat tạo 1 Số QT mới",
            self.edit_create_new_bat.text(),
            "Batch (*.bat);;Tất cả (*.*)",
        )
        if path:
            self.edit_create_new_bat.setText(os.path.normpath(path))

    def _browse_input_information_bat(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file .bat nhập thông tin",
            self.edit_input_information_bat.text(),
            "Batch (*.bat);;Tất cả (*.*)",
        )
        if path:
            self.edit_input_information_bat.setText(os.path.normpath(path))

    def _browse_input_expense_bat(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file .bat nhập khoản chi",
            self.edit_input_expense_bat.text(),
            "Batch (*.bat);;Tất cả (*.*)",
        )
        if path:
            self.edit_input_expense_bat.setText(os.path.normpath(path))

    def _browse_output(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục output", self.edit_output.text()
        )
        if path:
            output_folder = os.path.normpath(path)
            self.edit_output.setText(output_folder)
            self.edit_daily.setText(
                os.path.join(output_folder, os.path.basename(self.config.daily_tracking_file))
            )

    def on_save_config(self) -> None:
        try:
            self.config.set("bat_path", self.edit_bat.text().strip())
            self.config.set(
                "pad_create_new_bat_path",
                self.edit_create_new_bat.text().strip(),
            )
            self.config.set(
                "pad_input_information_bat_path",
                self.edit_input_information_bat.text().strip(),
            )
            self.config.set(
                "pad_input_expense_bat_path",
                self.edit_input_expense_bat.text().strip(),
            )
            self.config.set_output_folder(self.edit_output.text().strip())
            self.edit_output.setText(self.config.output_folder)
            self.edit_daily.setText(self.config.daily_tracking_file)
            self.config.ensure_folders()
            self.config.save()
            self.logger.info("Đã lưu cấu hình vào %s", self.config.path)

            # Khởi động lại watcher để áp dụng thư mục output mới.
            self.watcher.stop()
            self.watcher = OutputWatcher(self.config, self.database, self.logger)
            self.watcher.signals.output_ready.connect(self.on_output_ready)
            self.watcher.signals.file_error.connect(self.on_file_error)
            self.watcher.signals.log_message.connect(self.append_log)
            self._mark_current_file_handled()
            self._start_watcher()

            self._update_button_states()
            self.show_info("Lưu cấu hình", "Đã lưu cấu hình thành công.")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi lưu cấu hình.")
            self.show_error("Lỗi lưu cấu hình", f"Chi tiết: {exc}")

    def on_check_config(self) -> None:
        lines = []

        bat = self.edit_bat.text().strip()
        lines.append(
            f"• File .bat mở trợ lý: {'OK' if os.path.isfile(bat) else 'KHÔNG TỒN TẠI'}"
            f"\n  {bat}"
        )
        create_new_bat = self.edit_create_new_bat.text().strip()
        lines.append(
            f"• File .bat tạo 1 Số QT mới: "
            f"{'OK' if os.path.isfile(create_new_bat) else 'KHÔNG TỒN TẠI'}\n  {create_new_bat}"
        )
        input_information_bat = self.edit_input_information_bat.text().strip()
        lines.append(
            f"• File .bat nhập thông tin: "
            f"{'OK' if os.path.isfile(input_information_bat) else 'KHÔNG TỒN TẠI'}\n  {input_information_bat}"
        )
        input_expense_bat = self.edit_input_expense_bat.text().strip()
        lines.append(
            f"• File .bat nhập khoản chi: "
            f"{'OK' if os.path.isfile(input_expense_bat) else 'KHÔNG TỒN TẠI'}\n  {input_expense_bat}"
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
                "Đã mở trợ lý quyết toán. Đang chờ file output trong thư mục output...",
            )
            self.logger.info("Đã mở trợ lý quyết toán qua file .bat: %s", bat_path)
            self.append_log("Đã mở trợ lý quyết toán. Đang chờ file output trong thư mục output...")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi mở trợ lý quyết toán.")
            self.show_error(
                "Lỗi mở trợ lý",
                f"Không chạy được file .bat.\nChi tiết: {exc}",
            )

    # ================================================================== #
    # Chức năng: file output hiện tại
    # ================================================================== #
    def on_output_ready(self, result: Dict[str, Any]) -> None:
        """Nhận tín hiệu từ watcher khi có file bóc tách mới.

        File JSON bóc tách là bộ nhớ tạm của lô đang làm: nó giữ nguyên trong
        output để người dùng kiểm tra ở Bước 2 và Bước 3 đọc trực tiếp. Lô
        mới thay thế lô cũ nên bản JSON tạm trước đó được dọn đi.
        """
        self._discard_temp_json(keep_path=result.get("working_path"))
        self.current_working_file = result
        self._update_current_file_labels(
            note="Trợ lý đã bóc tách dữ liệu xong. Đang mở màn hình kiểm tra JSON..."
        )
        self.set_overall_status(
            "READY_FOR_REVIEW",
            "Trợ lý đã bóc tách dữ liệu xong. Đang mở màn hình kiểm tra JSON...",
        )
        self.refresh_history()
        self.append_log(f"File mới sẵn sàng: {result.get('file_name')}")
        self._open_extract_editor(auto=True)

    def _discard_temp_json(self, keep_path: Optional[str] = None) -> None:
        """Xóa file JSON tạm đang giữ (nếu có) khi nó không còn cần nữa.

        Chỉ xóa file nằm trong thư mục output đã cấu hình, để không bao giờ
        đụng vào file người dùng để ở nơi khác.
        """
        path = (self.current_working_file or {}).get("working_path") or ""
        if not path or (keep_path and os.path.normcase(path) == os.path.normcase(keep_path)):
            return
        if not file_utils.is_in_folder(path, self.config.output_folder):
            return
        if file_utils.remove_file(path):
            self.watcher.forget(path)
            self.logger.info("Đã dọn file JSON tạm cũ: %s", path)
            self.append_log(f"Đã dọn file JSON tạm cũ: {os.path.basename(path)}")

    def on_file_error(self, message: str) -> None:
        """Nhận tín hiệu lỗi từ watcher (không dùng hộp thoại chặn)."""
        self.lbl_file_note.setText(message)
        self.set_overall_status("ERROR", message)
        self.append_log("LỖI: " + message)

    def on_view_extract_file(self) -> None:
        """Bước 2: mở lại file bóc tách để xem/chỉnh sửa."""
        if not self.current_working_file:
            self.show_warning(
                "Chưa có file",
                "Chưa có file bóc tách nào. Hãy mở trợ lý ở Bước 1 và lưu file vào thư mục output.",
            )
            return
        path = self.current_working_file.get("working_path")
        if not path or not os.path.exists(path):
            self.show_error(
                "File không tồn tại",
                "File bóc tách không còn tồn tại trên ổ đĩa. Hãy lưu lại file vào thư mục output.",
            )
            self._update_button_states()
            return
        self._open_extract_editor(auto=False)

    def _open_extract_editor(
        self,
        auto: bool = False,
        error_rows: Optional[Dict[int, str]] = None,
    ) -> None:
        """Mở UI kiểm tra/sửa file bóc tách JSON.

        ``error_rows`` là các dòng Bước 3 không nhập được ({vị trí trong
        du_lieu_boc_tach: lý do}); JsonExtractDialog tô đỏ và ghi lý do cho các
        dòng này để người dùng sửa hoặc xóa.
        """
        cw = self.current_working_file
        if not cw:
            return
        path = cw.get("working_path") or ""
        try:
            dialog = JsonExtractDialog(path, self, error_rows=error_rows)
            dialog.exec()
            if dialog.saved:
                cw["file_size"] = os.path.getsize(path)
                cw["file_hash"] = file_utils.sha256_file(path)
                self._update_record_status(
                    "REVIEW_SAVED",
                    note="Đã lưu dữ liệu bóc tách JSON sau khi kiểm tra.",
                    mark_reviewed=True,
                )
                status = "REVIEW_SAVED"
                message = (
                    "Đã lưu dữ liệu bóc tách JSON. Có thể bấm “Nhập lên file hàng ngày”."
                )
            else:
                self._update_record_status(
                    "OPENED_FOR_REVIEW",
                    note="Đã mở dữ liệu bóc tách JSON để kiểm tra.",
                )
                status = "OPENED_FOR_REVIEW"
                message = (
                    "Đã mở dữ liệu bóc tách JSON. Kiểm tra, bấm “Lưu” nếu có chỉnh sửa, "
                    "rồi bấm “Nhập lên file hàng ngày”."
                )
            self.set_overall_status(status, message)
            action = "Tự mở" if auto else "Mở"
            self.logger.info("%s dữ liệu bóc tách JSON: %s", action, path)
            self.append_log(f"{action} dữ liệu bóc tách JSON: {os.path.basename(path)}")
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được dữ liệu bóc tách JSON.")
            self.lbl_file_note.setText(
                "Không mở được dữ liệu bóc tách JSON. Hãy kiểm tra lại file trong output. "
                f"Chi tiết: {exc}"
            )
            self.set_overall_status(
                "READY_FOR_REVIEW",
                "Không mở được dữ liệu bóc tách JSON.",
            )

    # ================================================================== #
    # Bước 3: nhập dữ liệu vào file theo dõi hàng ngày
    # ================================================================== #
    def _start_daily_worker(self, function, callback) -> None:
        """Chạy một tác vụ nặng, sau đó giao kết quả về UI thread."""
        if self._daily_thread is not None:
            return
        self._daily_worker_result = None
        self._daily_worker_error = None
        self._daily_worker_callback = callback
        self._daily_thread = QThread(self)
        self._daily_worker = FunctionWorker(function)
        self._daily_worker.moveToThread(self._daily_thread)
        self._daily_thread.started.connect(self._daily_worker.run)
        self._daily_worker.finished.connect(self._store_daily_worker_result)
        self._daily_worker.failed.connect(self._store_daily_worker_error)
        self._daily_worker.finished.connect(self._daily_thread.quit)
        self._daily_worker.failed.connect(self._daily_thread.quit)
        self._daily_thread.finished.connect(self._finish_daily_worker)
        self._daily_thread.start()

    def _store_daily_worker_result(self, result) -> None:
        self._daily_worker_result = result

    def _store_daily_worker_error(self, message: str) -> None:
        self._daily_worker_error = message

    def _finish_daily_worker(self) -> None:
        thread = self._daily_thread
        worker = self._daily_worker
        callback = self._daily_worker_callback
        result = self._daily_worker_result
        error = self._daily_worker_error
        self._daily_thread = None
        self._daily_worker = None
        self._daily_worker_callback = None
        if worker is not None:
            worker.deleteLater()
        if thread is not None:
            thread.deleteLater()
        if error:
            self._set_daily_running(False, "Không thể nhập dữ liệu vào file theo dõi.")
            self.logger.error("Lỗi Bước 3: %s", error)
            self.show_error(
                "Không nhập được dữ liệu",
                error + "\n\nFile theo dõi vẫn được giữ nguyên.",
            )
            self._refresh_daily_labels()
            return
        if callback is not None:
            callback(result)

    def on_import_daily(self) -> None:
        """Bước 3: lấy bản lưu mới nhất của file bóc tách và nhập lên file theo dõi."""
        if self._daily_import_running:
            return
        cw = self.current_working_file or {}
        source_path = cw.get("working_path") or ""
        daily_path = self.config.daily_tracking_file

        if not source_path or not os.path.isfile(source_path):
            self.show_warning(
                "Chưa có file bóc tách",
                "Chưa có file bóc tách nào để nhập. Hãy mở trợ lý ở Bước 1 và tải "
                "file về.",
            )
            self._update_button_states()
            return
        if not daily_path or not os.path.isfile(daily_path):
            self.show_warning(
                "Chưa có file theo dõi",
                "Không tìm thấy file theo dõi hàng ngày. Hãy chọn đúng file trong "
                "tab Cài đặt.",
            )
            return

        try:
            if file_utils.is_file_locked(daily_path):
                self.show_warning(
                    "File theo dõi đang mở",
                    "File theo dõi hàng ngày đang mở trong Excel. Hãy lưu và đóng "
                    "file rồi bấm lại.",
                )
                return
        except FileNotFoundError:
            self.show_error(
                "File không tồn tại",
                "File theo dõi không còn trên ổ đĩa.",
            )
            self._update_button_states()
            return

        # Bước 3 đọc thẳng file JSON tạm mà người dùng vừa xem/sửa ở Bước 2:
        # không tạo thêm bản sao trong thư mục Output nữa.
        self._update_record_status(
            "DAILY_IMPORT_PENDING",
            note="Đang nhập dữ liệu lên file theo dõi hàng ngày.",
            mark_reviewed=True,
        )

        self._set_daily_running(True, "Đang phân tích Phiếu cân, Bill và khoản chi...")
        record_id = cw.get("id")
        self._start_daily_worker(
            lambda: self.daily_import_service.analyze(
                source_path,
                daily_path,
                processed_file_id=record_id,
            ),
            self._on_daily_analysis_ready,
        )

    def _on_daily_analysis_ready(self, analysis: ImportAnalysis) -> None:
        if analysis.unmatched_rows:
            self._handle_unmatched_rows(analysis)
            return

        if not analysis.has_changes:
            self._set_daily_running(
                False, "Không có dòng mới nào để nhập vào file theo dõi."
            )
            self.show_info(
                "Không có dữ liệu mới",
                "Mọi chứng từ trong file bóc tách đều đã có sẵn trong file theo dõi "
                "(phần mềm đối chiếu bằng cột MD5). File theo dõi giữ nguyên.\n\n"
                "Muốn nhập lại thì xóa các dòng tương ứng khỏi file theo dõi rồi bấm "
                "lại nút này.",
            )
            return

        conflict_decisions: Dict[str, bool] = {}
        if analysis.conflicts:
            self.progress_daily.setVisible(False)
            dialog = ConflictDialog(analysis.conflicts, self)
            if dialog.exec() != QDialog.Accepted:
                self._set_daily_running(False, "Đã hủy lần nhập; chưa có dữ liệu nào được ghi.")
                return
            conflict_decisions = dialog.decisions()

        summary_text = (
            "Phần mềm đã phân tích xong:\n\n"
            f"• Dòng quyết toán mới: {analysis.new_info_count}\n"
            f"• Dòng quyết toán được cập nhật: {analysis.updated_info_count}\n"
            f"• Dòng khoản chi sẽ ghi: {len(analysis.expense_changes)}\n"
            f"• Chứng từ đã có sẵn trong file theo dõi (bỏ qua): "
            f"{analysis.duplicate_documents}\n\n"
            "Tiếp tục cập nhật file theo dõi hàng ngày?"
        )
        answer = QMessageBox.question(
            self,
            "Xác nhận nhập dữ liệu",
            summary_text,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            self._set_daily_running(False, "Đã hủy trước khi ghi file theo dõi.")
            return
        self.progress_daily.setVisible(True)
        self.lbl_daily_status.setText("Đang cập nhật file theo dõi...")
        self._start_daily_worker(
            lambda: self.daily_import_service.commit(
                analysis, conflict_decisions=conflict_decisions
            ),
            self._on_daily_commit_finished,
        )

    def _handle_unmatched_rows(self, analysis: ImportAnalysis) -> None:
        """Hỏi người dùng đúng hai lựa chọn cho các dòng không ghép được.

        Quay lại Bước 2: không ghi gì thêm, mở lại màn hình kiểm tra với các
        dòng lỗi được tô đỏ. Hủy dòng lỗi: xóa hẳn các dòng đó khỏi JSON tạm rồi
        phân tích lại để phần còn lại vẫn được nhập trong cùng đợt này.
        """
        self.progress_daily.setVisible(False)
        self.lbl_daily_status.setText(
            f"{len(analysis.unmatched_rows)} dòng chưa ghép được với dữ liệu quyết toán."
        )
        dialog = UnmatchedRowsDialog(analysis.unmatched_rows, self)
        dialog.exec()

        if dialog.choice != UnmatchedRowsDialog.DROP_ERRORS:
            self._set_daily_running(
                False,
                "Chưa nhập dữ liệu. Hãy sửa hoặc xóa các dòng lỗi ở Bước 2 rồi nhập lại.",
            )
            self._open_extract_editor(
                error_rows={
                    row.json_index: row.reason
                    for row in analysis.unmatched_rows
                    if row.json_index >= 0
                }
            )
            return

        try:
            removed = remove_extract_rows(
                analysis.output_path,
                [row.json_index for row in analysis.unmatched_rows],
            )
        except DailyImportError as exc:
            self._set_daily_running(False, "Không xóa được các dòng lỗi khỏi file JSON.")
            self.logger.error("Không xóa được dòng lỗi khỏi JSON tạm: %s", exc)
            self.show_error("Không xóa được dòng lỗi", str(exc))
            return

        self.logger.info("Đã xóa %s dòng lỗi khỏi file JSON tạm.", removed)
        self.append_log(f"Đã xóa {removed} dòng lỗi khỏi file JSON tạm.")
        self._set_daily_running(True, "Đang nhập lại các dòng còn lại...")
        self._start_daily_worker(
            lambda: self.daily_import_service.analyze(
                analysis.output_path,
                analysis.daily_path,
                processed_file_id=analysis.processed_file_id,
            ),
            self._on_daily_analysis_ready,
        )

    def _on_daily_commit_finished(self, summary: ImportSummary) -> None:
        self._set_daily_running(False)
        cw = self.current_working_file or {}
        cw["status"] = summary.status
        self.lbl_daily_status.setText("Hoàn tất.")
        self._update_current_file_labels(
            note=(
                f"Đã cập nhật file theo dõi: {summary.new_info} dòng mới, "
                f"{summary.updated_info} dòng cập nhật, {summary.new_expenses} khoản chi mới."
            )
        )
        self.refresh_history()
        self._refresh_daily_labels()
        self.append_log(
            f"Bước 3 hoàn tất: {summary.new_info} dòng mới, "
            f"{summary.updated_info} dòng cập nhật, {summary.new_expenses} khoản chi."
        )
        message = (
            "Đã cập nhật file theo dõi hàng ngày.\n\n"
            f"Dòng quyết toán mới: {summary.new_info}\n"
            f"Dòng quyết toán cập nhật: {summary.updated_info}\n"
            f"Khoản chi mới: {summary.new_expenses}\n"
            f"Khoản chi cập nhật: {summary.updated_expenses}\n"
            f"Chứng từ đã có sẵn nên bỏ qua: {summary.duplicate_documents}"
        )
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle("Nhập dữ liệu hoàn tất")
        box.setText(message)
        open_button = box.addButton("Mở file theo dõi", QMessageBox.ActionRole)
        box.addButton("Đóng", QMessageBox.AcceptRole)
        box.exec()
        if box.clickedButton() is open_button:
            self._open_daily_file()

        # Lô này đã nhập xong và không còn dòng lỗi -> dọn bộ nhớ tạm.
        self._finish_temp_json()

    def _finish_temp_json(self) -> None:
        """Dọn file JSON tạm sau khi đã nhập xong toàn bộ lô."""
        self._discard_temp_json()
        self.current_working_file = None
        self.overall_status = "DAILY_IMPORTED"
        # Không tự tay reset _last_saved_shown ở đây: để _refresh_saved_label thấy
        # giá trị cũ khác giá trị mới ("") mà xóa dòng “Lưu thành công lần cuối”,
        # nếu không nhãn sẽ kẹt lại mốc giờ của file vừa bị dọn.
        self._update_current_file_labels()
        self.lbl_file_note.setText(
            "Đã nhập xong lô này; file JSON tạm đã được dọn. Bấm “Mở trợ lý” ở Bước 1 "
            "khi có lô chứng từ mới."
        )

    def _open_daily_file(self) -> None:
        path = self.config.daily_tracking_file
        if not path or not os.path.isfile(path):
            self.show_warning(
                "Chưa có file theo dõi",
                "Không tìm thấy file theo dõi hàng ngày. Hãy kiểm tra Cài đặt.",
            )
            return
        try:
            file_utils.open_file(path)
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Không mở được file theo dõi hàng ngày.")
            self.show_error("Không mở được file", str(exc))

    # ================================================================== #
    # Bước 4: chạy luồng PAD RPA lên phần mềm quyết toán
    # ================================================================== #
    def _begin_rpa_cooldown(self) -> None:
        self._rpa_cooldown = True
        self._update_button_states()
        QTimer.singleShot(5000, self._end_rpa_cooldown)

    def _mark_rpa_launched(self, message: str) -> None:
        launched_at = datetime.now().strftime("%H:%M:%S ngày %d/%m/%Y")
        self._rpa_launched_at = launched_at
        self.lbl_rpa_status.setText(f"{message} lúc {launched_at}.")
        self.append_log(f"{message} lúc {launched_at}.")
        self._begin_rpa_cooldown()

    def on_create_new_sqt(self) -> None:
        """Gọi flow PAD tạo đúng một Số QT mới, không đọc Excel."""
        bat_path = self.config.pad_create_new_bat_path
        missing = (
            "Không tìm thấy file chạy RPA tạo mới.\n"
            "Vui lòng kiểm tra cấu hình đường dẫn BAT."
        )
        try:
            launch_bat(
                bat_path,
                logger=self.logger,
                description="tạo 1 Số QT mới",
                missing_message=missing,
            )
        except RpaLaunchError as exc:
            self.logger.error("Không khởi chạy được RPA tạo mới: %s", exc)
            self.show_error(
                "Không chạy được RPA tạo mới",
                f"{exc}\n\nĐường dẫn hiện tại: {bat_path}",
            )
            return

        self.logger.info("Đã gọi BAT tạo 1 Số QT mới: %s", bat_path)
        self._mark_rpa_launched("Đã khởi chạy RPA tạo 1 Số QT mới")

    def _run_selected_sqt_rpa(
        self,
        *,
        sheet_name: str,
        operation: str,
        bat_path: str,
        action_label: str,
        missing_bat_title: str,
        missing_bat_message: str,
        launched_message: str,
    ) -> None:
        """Đọc SQT từ Excel, ghi JSON lựa chọn rồi gọi một flow PAD có chọn SQT."""
        if self._select_sqt_dialog_open:
            self.show_warning(
                "Đang chọn SQT",
                "Có một cửa sổ chọn Số quyết toán đang mở. Vui lòng hoàn tất cửa sổ đó trước.",
            )
            return

        daily_path = self.config.daily_tracking_file
        try:
            validate_bat_path(bat_path, missing_bat_message)
        except RpaLaunchError as exc:
            self.logger.error("Không tìm thấy BAT %s: %s", action_label, bat_path)
            self.show_error(
                missing_bat_title,
                f"{exc}\n\nĐường dẫn hiện tại: {bat_path}",
            )
            return

        self.logger.info(
            "Người dùng mở popup chọn SQT cho %s từ file: %s (sheet %s)",
            action_label,
            daily_path,
            sheet_name,
        )
        try:
            sqt_items = read_sqt_items(daily_path, sheet_name=sheet_name)
        except DailyFileNotFoundError as exc:
            self.logger.error("Không tìm thấy file Excel hằng ngày: %s", daily_path)
            self.show_error(
                "Không tìm thấy file Excel hằng ngày",
                f"{exc}\n\nĐường dẫn hiện tại: {daily_path}",
            )
            return
        except DailyFileLockedError as exc:
            self.logger.error("File Excel hằng ngày đang bị khóa: %s", daily_path)
            self.show_warning("File Excel đang bị khóa", str(exc))
            return
        except (MissingSheetError, MissingColumnError, EmptySqtListError) as exc:
            self.logger.error("Dữ liệu SQT trong Excel chưa hợp lệ: %s", exc)
            self.show_error("Dữ liệu Excel chưa hợp lệ", str(exc))
            return
        except SqtSelectionError as exc:
            self.logger.exception("Không đọc được danh sách SQT.")
            self.show_error("Không đọc được file Excel", str(exc))
            return

        self.logger.info("Đọc được %s SQT từ sheet %s.", len(sqt_items), sheet_name)
        dialog = SelectSqtDialog(sqt_items, self)
        self._select_sqt_dialog_open = True
        self._update_button_states()
        try:
            result = dialog.exec()
        finally:
            self._select_sqt_dialog_open = False
            self._update_button_states()

        if result != QDialog.Accepted:
            self.logger.info(
                "Người dùng hủy popup chọn SQT cho %s; không ghi JSON và không gọi BAT.",
                action_label,
            )
            return

        selected_sqt = dialog.selected_values()
        if not selected_sqt:
            self.show_warning(
                "Chưa chọn SQT",
                "Vui lòng chọn ít nhất một Số quyết toán.",
            )
            return

        try:
            json_path = write_selection_json(
                self.config.rpa_input_selection_path,
                daily_path,
                selected_sqt,
                sheet_name=sheet_name,
                operation=operation,
            )
        except SelectionJsonWriteError as exc:
            self.logger.exception("Không ghi được JSON lựa chọn SQT.")
            self.show_error("Không ghi được JSON", str(exc))
            return

        self.logger.info("Danh sách SQT đã chọn: %s", selected_sqt)
        self.logger.info("Đã tạo JSON lựa chọn RPA: %s", json_path)
        self.logger.info("File Excel hằng ngày dùng cho PAD: %s", os.path.abspath(daily_path))

        try:
            launch_bat(
                bat_path,
                logger=self.logger,
                description=action_label,
                missing_message=missing_bat_message,
            )
        except RpaLaunchError as exc:
            self.logger.error("Không khởi chạy được RPA %s: %s", action_label, exc)
            self.show_error(f"Không chạy được RPA {action_label}", str(exc))
            return

        self.logger.info("Đã gọi BAT %s: %s", action_label, bat_path)
        self._mark_rpa_launched(launched_message)

    def on_input_information(self) -> None:
        """Đọc SQT từ sheet thông tin, ghi JSON lựa chọn rồi gọi flow PAD."""
        self._run_selected_sqt_rpa(
            sheet_name=INFO_SHEET,
            operation=SELECTION_OPERATION,
            bat_path=self.config.pad_input_information_bat_path,
            action_label="nhập thông tin",
            missing_bat_title="Không tìm thấy BAT nhập thông tin",
            missing_bat_message=(
                "Không tìm thấy file chạy RPA nhập thông tin.\n"
                "Vui lòng kiểm tra cấu hình đường dẫn BAT."
            ),
            launched_message="Đã khởi chạy RPA nhập thông tin",
        )

    def on_input_expense(self) -> None:
        """Đọc SQT từ sheet khoản chi, ghi JSON lựa chọn rồi gọi flow PAD."""
        self._run_selected_sqt_rpa(
            sheet_name=EXPENSE_SHEET,
            operation=EXPENSE_SELECTION_OPERATION,
            bat_path=self.config.pad_input_expense_bat_path,
            action_label="nhập khoản chi",
            missing_bat_title="Không tìm thấy BAT nhập khoản chi",
            missing_bat_message=(
                "Không tìm thấy file chạy RPA nhập khoản chi.\n"
                "Vui lòng kiểm tra cấu hình đường dẫn BAT."
            ),
            launched_message="Đã khởi chạy RPA nhập khoản chi",
        )

    def on_run_rpa(self) -> None:
        rpa_bat = self.config.pad_bat_path
        if not rpa_bat or not os.path.isfile(rpa_bat):
            self.logger.error("Không tìm thấy file .bat chạy RPA: %s", rpa_bat)
            self.show_error(
                "Chưa có file .bat chạy RPA",
                "Chưa tìm thấy file .bat chạy luồng PAD RPA. Hãy vào tab Cài đặt, "
                "trỏ tới đúng file rồi bấm “Lưu cấu hình”.\n"
                f"Đường dẫn hiện tại: {rpa_bat}",
            )
            return

        answer = QMessageBox.question(
            self,
            "Xác nhận chạy RPA",
            "Phần mềm sẽ chạy luồng RPA để nhập dữ liệu mới lên phần mềm quyết toán.\n\n"
            "Hãy chắc chắn dữ liệu đã được nhập vào file theo dõi hàng ngày và "
            "không dùng chuột/bàn phím trong lúc RPA chạy.\n\n"
            "Tiếp tục?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        try:
            subprocess.Popen(
                f'"{rpa_bat}"',
                shell=True,
                cwd=os.path.dirname(rpa_bat) or None,
            )
        except Exception as exc:  # noqa: BLE001
            self.logger.exception("Lỗi khi chạy luồng RPA.")
            self.show_error(
                "Lỗi chạy RPA",
                f"Không chạy được file .bat.\nChi tiết: {exc}",
            )
            return

        launched_at = datetime.now().strftime("%H:%M:%S ngày %d/%m/%Y")
        self._rpa_launched_at = launched_at
        self.lbl_rpa_status.setText(
            f"Đã khởi chạy luồng RPA lúc {launched_at}. Theo dõi tiến trình trong "
            "cửa sổ Power Automate."
        )
        self.logger.info("Đã khởi chạy luồng RPA qua file .bat: %s", rpa_bat)
        self.append_log(f"Đã khởi chạy luồng RPA lúc {launched_at}.")
        if self.current_working_file:
            self._update_record_status(
                "RPA_LAUNCHED",
                note=f"Đã khởi chạy luồng RPA quyết toán lúc {launched_at}.",
            )

        # Chống bấm hai lần liên tiếp làm chạy trùng luồng RPA.
        self._rpa_cooldown = True
        self._update_button_states()
        QTimer.singleShot(5000, self._end_rpa_cooldown)

    def _end_rpa_cooldown(self) -> None:
        self._rpa_cooldown = False
        self._update_button_states()

    # ================================================================== #
    # Đóng cửa sổ
    # ================================================================== #
    def closeEvent(self, event) -> None:  # noqa: N802 - override Qt
        self._saved_timer.stop()
        if self._daily_thread is not None and self._daily_thread.isRunning():
            self._daily_thread.quit()
            self._daily_thread.wait(3000)
        try:
            self.watcher.stop()
        except Exception:  # noqa: BLE001
            pass
        self.logger.info("==== Đóng ứng dụng ====")
        super().closeEvent(event)
