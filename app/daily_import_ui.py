"""Các hộp thoại thân thiện cho luồng nhập file theo dõi hàng ngày."""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import QObject, Qt, Signal, Slot
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .daily_import import (
    DOC_BILL,
    DOC_EXPENSE,
    DOC_SCALE,
    DailyImportError,
    FieldConflict,
    UnmatchedRow,
    classify_doc_type,
    extract_summary,
    load_extract_payload,
    match_date_of,
    parse_date,
    parse_number,
    save_extract_payload,
)

# Dòng lỗi (chưa nhập được) và dòng cần soát lại được tô nền để nhìn ra ngay.
ERROR_BG = QColor("#FEE2E2")
ERROR_FG = QColor("#B91C1C")
WARN_BG = QColor("#FEF3C7")


def _display_date(value: Any) -> str:
    parsed = parse_date(value)
    if not parsed:
        return str(value or "")
    try:
        return datetime.strptime(parsed, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(value or "")


def _display_number(value: Any, decimals: int = 0) -> str:
    """Số theo cách viết Việt Nam: 49.572.000 (tiền) hoặc 27,83 (tấn)."""
    number = parse_number(value)
    if number is None:
        return ""
    text = f"{number:,.{decimals}f}"
    return text.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def _display_value(value: Any) -> str:
    """Giá trị thô của một trường JSON để đưa vào ô nhập."""
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, indent=2)
    return str(value)


def _main_info(doc_type: str, row: Dict[str, Any]) -> str:
    """Thông tin nhận dạng chính của một dòng; thiếu thì để trống."""
    parts: List[str] = []
    if doc_type == DOC_EXPENSE:
        if row.get("so_hd"):
            parts.append(f"HĐ {row['so_hd']}")
        unit_price = _display_number(row.get("don_gia"))
        if unit_price:
            parts.append(f"ĐG {unit_price}")
    elif doc_type == DOC_SCALE:
        if row.get("so_chi_seal"):
            parts.append(f"Seal {row['so_chi_seal']}")
        elif row.get("bien_so_xe"):
            parts.append(f"Xe {row['bien_so_xe']}")
    elif doc_type == DOC_BILL:
        if row.get("so_bill"):
            parts.append(f"Bill {row['so_bill']}")
        elif row.get("ten_tau"):
            parts.append(str(row["ten_tau"]))
    return "  •  ".join(parts)


def _amount_info(doc_type: str, row: Dict[str, Any]) -> str:
    """Cột “Số tấn / Tổng tiền” tùy theo loại chứng từ."""
    tons = _display_number(row.get("so_tan"), 2)
    if doc_type == DOC_SCALE:
        return f"{tons} tấn" if tons else ""
    if doc_type == DOC_EXPENSE:
        return _display_number(row.get("tong_tien"))
    if doc_type == DOC_BILL:
        return f"{tons} tấn" if tons else _display_number(row.get("tong_tien"))
    return ""


def _warning_count(value: Any) -> int:
    if isinstance(value, list):
        return len(value)
    return 0 if value in (None, "") else 1


class FunctionWorker(QObject):
    """Chạy một callable ở QThread và trả kết quả về UI thread."""

    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, function: Callable[[], Any]):
        super().__init__()
        self.function = function

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(self.function())
        except Exception as exc:  # noqa: BLE001 - chuyển lỗi sang UI
            self.failed.emit(str(exc))


class _EditExtractRowDialog(QDialog):
    """Form chi tiết một dòng bóc tách: sửa đúng các trường của schema JSON."""

    FIELDS = [
        ("stt_hien_thi", "STT hiển thị"),
        ("file_nguon", "File nguồn"),
        ("ma_md5_file", "Mã MD5 file"),
        ("loai_chung_tu", "Loại chứng từ"),
        ("trang_thai", "Trạng thái"),
        ("ngay_dong", "Ngày Đóng"),
        ("so_container", "Số Container"),
        ("bien_so_xe", "Biển số xe"),
        ("so_tan", "Số tấn"),
        ("loai_hang", "Loại hàng"),
        ("noi_dong", "Nơi đóng"),
        ("nguoi_nhan", "Người nhận"),
        ("ten_tau", "Tên tàu"),
        ("ngay_chay", "Ngày chạy"),
        ("vt_bien", "VT biển"),
        ("gia_vat_lieu", "Giá vật liệu"),
        ("so_hd", "Số HĐ"),
        ("so_bill", "Số Bill"),
        ("so_chi_seal", "Số chì/Seal"),
        ("don_gia", "Đơn giá"),
        ("thanh_tien", "Thành tiền"),
        ("vat", "VAT"),
        ("tong_tien", "Tổng tiền"),
        ("truong_khac", "Trường khác"),
        ("do_tin_cay", "Độ tin cậy"),
        ("canh_bao", "Cảnh báo"),
    ]
    INT_FIELDS = {"stt_hien_thi"}
    NUMBER_FIELDS = {
        "so_tan",
        "gia_vat_lieu",
        "don_gia",
        "thanh_tien",
        "vat",
        "tong_tien",
    }
    STRUCTURED_DEFAULTS = {"truong_khac": {}, "canh_bao": []}

    def __init__(self, row: Dict[str, Any], reason: str = "", parent=None):
        super().__init__(parent)
        self.row = row
        self.editors: Dict[str, QWidget] = {}
        self._updated: Dict[str, Any] = dict(row)

        self.setWindowTitle("Sửa dòng bóc tách")
        self.resize(620, 700)
        self.setMinimumSize(520, 420)

        layout = QVBoxLayout(self)

        if reason:
            note = QLabel(f"Dòng này chưa nhập được: {reason}")
            note.setObjectName("noteText")
            note.setWordWrap(True)
            layout.addWidget(note)

        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        form.setLabelAlignment(Qt.AlignRight)
        form.setContentsMargins(4, 4, 12, 4)
        form.setSpacing(8)
        for key, label in self.FIELDS:
            value = _display_value(row.get(key))
            if key in self.STRUCTURED_DEFAULTS:
                editor: QWidget = QPlainTextEdit(value)
                editor.setMinimumHeight(90)
            else:
                editor = QLineEdit(value)
            form.addRow(label + ":", editor)
            self.editors[key] = editor

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(form_host)
        layout.addWidget(scroll, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Xong")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @Slot()
    def _on_accept(self) -> None:
        updated = dict(self.row)
        try:
            for key, editor in self.editors.items():
                text = (
                    editor.toPlainText()
                    if isinstance(editor, QPlainTextEdit)
                    else editor.text()
                )
                updated[key] = self._parse_value(key, text)
        except ValueError as exc:
            QMessageBox.warning(self, "Dữ liệu chưa hợp lệ", str(exc))
            return
        self._updated = updated
        self.accept()

    def _parse_value(self, key: str, text: str) -> Any:
        value = text.strip()
        if key in self.STRUCTURED_DEFAULTS:
            if not value:
                default = self.STRUCTURED_DEFAULTS[key]
                return list(default) if isinstance(default, list) else dict(default)
            try:
                return json.loads(value)
            except json.JSONDecodeError as exc:
                label = dict(self.FIELDS).get(key, key)
                raise ValueError(f"Trường '{label}' phải là JSON hợp lệ: {exc}") from exc
        if not value:
            return None
        if key in self.INT_FIELDS:
            parsed = parse_number(value)
            if parsed is None or int(parsed) != parsed:
                label = dict(self.FIELDS).get(key, key)
                raise ValueError(f"Trường '{label}' phải là số nguyên.")
            return int(parsed)
        if key in self.NUMBER_FIELDS:
            parsed = parse_number(value)
            if parsed is None:
                label = dict(self.FIELDS).get(key, key)
                raise ValueError(f"Trường '{label}' phải là số.")
            return parsed
        return value

    def updated_row(self) -> Dict[str, Any]:
        return self._updated


class JsonExtractDialog(QDialog):
    """Bước 2: xem cả lô bóc tách dạng bảng, sửa/xóa từng dòng rồi lưu lại.

    ``error_rows`` là các dòng Bước 3 không ghép được ({vị trí trong
    du_lieu_boc_tach: lý do}); các dòng này được tô đỏ kèm lý do để người dùng
    sửa hoặc xóa ngay tại đây.
    """

    COLUMNS = [
        "STT",
        "Loại chứng từ",
        "File nguồn",
        "Container",
        "Ngày dùng để ghép",
        "Thông tin chính",
        "Số tấn / Tổng tiền",
        "Trạng thái",
    ]

    def __init__(
        self,
        path: str,
        parent=None,
        error_rows: Optional[Dict[int, str]] = None,
    ):
        super().__init__(parent)
        self.path = path
        self.payload: Dict[str, Any] = load_extract_payload(path)
        self.rows: List[Dict[str, Any]] = [
            dict(row) for row in self.payload.get("du_lieu_boc_tach") or []
        ]
        # Lý do lỗi đi kèm từng dòng, giữ song song với self.rows để khi xóa dòng
        # thì lý do không bị lệch vị trí.
        self.row_errors: List[str] = [
            str((error_rows or {}).get(index, "")) for index in range(len(self.rows))
        ]
        self.saved = False
        self._dirty = False

        self.setWindowTitle("Kiểm tra dữ liệu bóc tách")
        self.resize(1160, 660)
        self.setMinimumSize(900, 480)

        layout = QVBoxLayout(self)

        title = QLabel(path)
        title.setObjectName("metaText")
        title.setWordWrap(True)
        title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(title)

        self.lbl_error_hint = QLabel()
        self.lbl_error_hint.setObjectName("noteText")
        self.lbl_error_hint.setWordWrap(True)
        self.lbl_error_hint.setVisible(False)
        layout.addWidget(self.lbl_error_hint)

        summary = QFrame()
        summary.setObjectName("guideBox")
        summary_layout = QVBoxLayout(summary)
        summary_layout.setContentsMargins(12, 10, 12, 10)
        summary_layout.setSpacing(4)
        self.lbl_summary = QLabel()
        self.lbl_summary.setWordWrap(True)
        self.lbl_types = QLabel()
        self.lbl_types.setObjectName("metaText")
        self.lbl_types.setWordWrap(True)
        summary_layout.addWidget(self.lbl_summary)
        summary_layout.addWidget(self.lbl_types)
        layout.addWidget(summary)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.doubleClicked.connect(self.edit_row)
        layout.addWidget(self.table, stretch=1)

        buttons = QHBoxLayout()
        btn_edit = QPushButton("Sửa dòng")
        btn_delete = QPushButton("Xóa dòng")
        btn_save = QPushButton("Lưu")
        btn_save.setProperty("variant", "primary")
        btn_close = QPushButton("Đóng")
        btn_edit.clicked.connect(self.edit_row)
        btn_delete.clicked.connect(self.delete_row)
        btn_save.clicked.connect(self.save)
        btn_close.clicked.connect(self.reject)
        buttons.addWidget(btn_edit)
        buttons.addWidget(btn_delete)
        buttons.addStretch(1)
        buttons.addWidget(btn_save)
        buttons.addWidget(btn_close)
        layout.addLayout(buttons)

        self._refresh()
        if self.rows:
            self.table.selectRow(self._first_error_row())

    # ------------------------------------------------------------------ #
    # Hiển thị
    # ------------------------------------------------------------------ #
    def _refresh(self) -> None:
        error_indexes = {
            index for index, reason in enumerate(self.row_errors) if reason
        }
        counts = extract_summary(
            self.rows, self.payload.get("canh_bao"), error_indexes
        )
        self.lbl_summary.setText(
            f"Tổng số file: {counts['files']}     •     "
            f"Tổng số dòng bóc tách: {counts['total']}     •     "
            f"Dòng OK: {counts['ok']}     •     "
            f"Cần kiểm tra: {counts['need_check']}     •     "
            f"Cảnh báo: {counts['warnings']}"
        )
        self.lbl_types.setText(
            f"Khoản chi: {counts[DOC_EXPENSE]}     •     "
            f"Phiếu cân: {counts[DOC_SCALE]}     •     "
            f"Bill: {counts[DOC_BILL]}     •     "
            f"Loại khác: {counts['other']}"
        )
        self.lbl_error_hint.setVisible(bool(error_indexes))
        if error_indexes:
            self.lbl_error_hint.setText(
                f"{len(error_indexes)} dòng chưa nhập được (tô đỏ bên dưới). Hãy sửa "
                "lại hoặc xóa dòng đó, bấm “Lưu” rồi nhập lại ở Bước 3."
            )

        self.table.setRowCount(len(self.rows))
        for index, row in enumerate(self.rows):
            reason = self.row_errors[index]
            doc_type = classify_doc_type(row.get("loai_chung_tu"))
            warnings = _warning_count(row.get("canh_bao"))
            needs_check = bool(row.get("trang_thai")) and str(
                row.get("trang_thai")
            ).strip().upper() != "OK"
            for col, text in enumerate(self._row_cells(index, row, doc_type)):
                item = QTableWidgetItem(text)
                item.setToolTip(self._row_tooltip(row, reason))
                if reason:
                    item.setBackground(ERROR_BG)
                    item.setForeground(ERROR_FG)
                elif needs_check or warnings:
                    item.setBackground(WARN_BG)
                self.table.setItem(index, col, item)

    def _row_cells(
        self, index: int, row: Dict[str, Any], doc_type: str
    ) -> List[str]:
        stt = row.get("stt_hien_thi")
        return [
            str(stt if stt not in (None, "") else index + 1),
            str(row.get("loai_chung_tu") or "—"),
            str(row.get("file_nguon") or ""),
            str(row.get("so_container") or ""),
            _display_date(match_date_of(row)),
            _main_info(doc_type, row),
            _amount_info(doc_type, row),
            self._status_text(index, row),
        ]

    def _status_text(self, index: int, row: Dict[str, Any]) -> str:
        if self.row_errors[index]:
            return self.row_errors[index]
        status = str(row.get("trang_thai") or "").strip() or "OK"
        warnings = _warning_count(row.get("canh_bao"))
        return f"{status} • {warnings} cảnh báo" if warnings else status

    @staticmethod
    def _row_tooltip(row: Dict[str, Any], reason: str) -> str:
        lines = []
        if reason:
            lines.append(f"Chưa nhập được: {reason}")
        lines.append(f"File nguồn: {row.get('file_nguon') or '—'}")
        warnings = row.get("canh_bao")
        if isinstance(warnings, list) and warnings:
            lines.append("Cảnh báo: " + json.dumps(warnings, ensure_ascii=False))
        return "\n".join(lines)

    def _first_error_row(self) -> int:
        for index, reason in enumerate(self.row_errors):
            if reason:
                return index
        return 0

    def _selected_index(self) -> Optional[int]:
        index = self.table.currentRow()
        if index < 0 or index >= len(self.rows):
            QMessageBox.information(
                self, "Chưa chọn dòng", "Hãy chọn một dòng trong bảng trước."
            )
            return None
        return index

    # ------------------------------------------------------------------ #
    # Thao tác
    # ------------------------------------------------------------------ #
    @Slot()
    def edit_row(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        dialog = _EditExtractRowDialog(self.rows[index], self.row_errors[index], self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.rows[index] = dialog.updated_row()
        # Dòng vừa sửa coi như đã xử lý; Bước 3 sẽ chấm lại khi nhập.
        self.row_errors[index] = ""
        self._dirty = True
        self._refresh()
        self.table.selectRow(index)

    @Slot()
    def delete_row(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        row = self.rows[index]
        answer = QMessageBox.question(
            self,
            "Xóa dòng",
            "Xóa dòng này khỏi dữ liệu bóc tách?\n\n"
            f"{row.get('loai_chung_tu') or 'Chứng từ'} • "
            f"{row.get('so_container') or '—'} • {row.get('file_nguon') or '—'}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.rows.pop(index)
        self.row_errors.pop(index)
        self._dirty = True
        self._refresh()
        if self.rows:
            self.table.selectRow(min(index, len(self.rows) - 1))

    @Slot()
    def save(self) -> None:
        """Lưu xong là đóng luôn.

        Không hiện thêm hộp thoại “đã lưu” bắt người dùng tắt: dòng “Lưu thành công
        lần cuối” ở Bước 2 tự cập nhật, thế là đủ báo thành công.
        """
        self.payload["du_lieu_boc_tach"] = self.rows
        try:
            save_extract_payload(self.path, self.payload)
        except DailyImportError as exc:
            QMessageBox.critical(self, "Không lưu được JSON", str(exc))
            return
        self.saved = True
        self._dirty = False
        self.accept()

    def reject(self) -> None:  # noqa: N802 - override Qt
        if self._dirty:
            answer = QMessageBox.question(
                self,
                "Chưa lưu thay đổi",
                "Bạn đã sửa dữ liệu nhưng chưa bấm “Lưu”. Đóng và bỏ các thay đổi?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
        super().reject()


class UnmatchedRowsDialog(QDialog):
    """Popup khi có dòng không ghép được: sửa lại ở Bước 2, hoặc bỏ dòng lỗi."""

    BACK_TO_REVIEW = "REVIEW"
    DROP_ERRORS = "DROP"

    def __init__(self, rows: List[UnmatchedRow], parent=None):
        super().__init__(parent)
        # Đóng cửa sổ bằng dấu X = không nhập gì thêm, quay lại kiểm tra.
        self.choice = self.BACK_TO_REVIEW
        self.setWindowTitle("Có dòng chưa nhập được")
        self.resize(880, 420)

        layout = QVBoxLayout(self)
        hint = QLabel(
            f"{len(rows)} dòng không ghép được với dữ liệu quyết toán nên chưa được "
            "nhập. Các dòng còn lại vẫn nhập bình thường."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(len(rows), 5)
        table.setHorizontalHeaderLabels(
            [
                "Loại chứng từ",
                "Container",
                "Ngày dùng để ghép",
                "Số HĐ/Bill/Seal",
                "Lý do",
            ]
        )
        table.setAlternatingRowColors(True)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        for index, row in enumerate(rows):
            values = [
                row.doc_label,
                row.container,
                _display_date(row.match_date),
                row.reference,
                row.reason,
            ]
            for col, value in enumerate(values):
                table.setItem(index, col, QTableWidgetItem(str(value)))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table, stretch=1)

        buttons = QHBoxLayout()
        btn_drop = QPushButton("Hủy các dòng lỗi và nhập các dòng còn lại")
        btn_review = QPushButton("Quay lại bước 2 để kiểm tra")
        btn_review.setProperty("variant", "primary")
        btn_review.setDefault(True)
        btn_drop.clicked.connect(self._choose_drop)
        btn_review.clicked.connect(self._choose_review)
        buttons.addStretch(1)
        buttons.addWidget(btn_drop)
        buttons.addWidget(btn_review)
        layout.addLayout(buttons)

    @Slot()
    def _choose_review(self) -> None:
        self.choice = self.BACK_TO_REVIEW
        self.accept()

    @Slot()
    def _choose_drop(self) -> None:
        self.choice = self.DROP_ERRORS
        self.accept()


class ConflictDialog(QDialog):
    """Cho người dùng quyết định khi dữ liệu mới khác ô đã nhập."""

    def __init__(self, conflicts: List[FieldConflict], parent=None):
        super().__init__(parent)
        self.conflicts = conflicts
        self.combos: Dict[str, QComboBox] = {}
        self.setWindowTitle("Kiểm tra dữ liệu khác nhau")
        self.resize(900, 440)

        layout = QVBoxLayout(self)
        hint = QLabel(
            "Dữ liệu mới khác thông tin đang có. Mặc định phần mềm giữ dữ liệu "
            "hiện tại; bạn có thể chọn dùng dữ liệu mới cho từng trường."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(len(conflicts), 5)
        table.setHorizontalHeaderLabels(
            ["Dòng Excel", "Trường", "Hiện tại", "Dữ liệu mới", "Lựa chọn"]
        )
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        for row_index, conflict in enumerate(conflicts):
            values = [
                str(conflict.target_row),
                conflict.field_name,
                str(conflict.current_value or ""),
                str(conflict.new_value or ""),
            ]
            for col, value in enumerate(values):
                table.setItem(row_index, col, QTableWidgetItem(value))
            combo = QComboBox()
            combo.addItem("Giữ dữ liệu hiện tại", False)
            combo.addItem("Dùng dữ liệu mới", True)
            table.setCellWidget(row_index, 4, combo)
            self.combos[conflict.conflict_id] = combo
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Xác nhận lựa chọn")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy lần nhập")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def decisions(self) -> Dict[str, bool]:
        return {
            key: bool(combo.currentData()) for key, combo in self.combos.items()
        }
