"""Các hộp thoại thân thiện cho luồng nhập file theo dõi hàng ngày."""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import QObject, Qt, Signal, Slot
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
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .daily_import import (
    BillChoiceRequest,
    FieldConflict,
    STATE_IGNORED,
    STATE_LABELS,
    normalize_cargo,
    normalize_container,
    parse_date,
    parse_number,
)
from .database import Database


def _display_date(value: Any) -> str:
    parsed = parse_date(value)
    if not parsed:
        return str(value or "")
    try:
        return datetime.strptime(parsed, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(value or "")


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


class JsonExtractDialog(QDialog):
    """Dialog kiểm tra/sửa dữ liệu bóc tách JSON trước khi nhập file theo dõi."""

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

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path
        self.payload: Dict[str, Any] = {}
        self.rows: List[Dict[str, Any]] = []
        self.current_index = -1
        self.saved = False
        self._dirty = False
        self._loading = False

        self._load_file()

        self.setWindowTitle("Kiểm tra dữ liệu bóc tách")
        self.resize(860, 640)
        self.setMinimumSize(680, 460)

        layout = QVBoxLayout(self)

        title = QLabel(path)
        title.setObjectName("metaText")
        title.setWordWrap(True)
        title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(title)

        row_picker = QHBoxLayout()
        row_picker.addWidget(QLabel("Dòng dữ liệu:"))
        self.row_combo = QComboBox()
        self.row_combo.currentIndexChanged.connect(self._on_row_changed)
        row_picker.addWidget(self.row_combo, stretch=1)
        layout.addLayout(row_picker)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Tên trường", "Dữ liệu"])
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(lambda _item: self._mark_dirty())
        layout.addWidget(self.table, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Close)
        buttons.button(QDialogButtonBox.Save).setText("Lưu")
        buttons.button(QDialogButtonBox.Close).setText("Đóng")
        buttons.button(QDialogButtonBox.Save).clicked.connect(self.save)
        buttons.rejected.connect(self.accept)
        layout.addWidget(buttons)

        self._populate_rows()

    def _load_file(self) -> None:
        try:
            with open(self.path, "r", encoding="utf-8-sig") as f:
                payload = json.load(f)
        except json.JSONDecodeError as exc:
            raise ValueError(f"File JSON không hợp lệ: {exc}") from exc
        except UnicodeDecodeError as exc:
            raise ValueError("File JSON không đọc được bằng UTF-8.") from exc
        except OSError as exc:
            raise ValueError(f"Không đọc được file JSON: {exc}") from exc

        if not isinstance(payload, dict):
            raise ValueError("File JSON phải có object gốc.")
        rows = payload.get("du_lieu_boc_tach")
        if not isinstance(rows, list):
            raise ValueError("File JSON thiếu mảng 'du_lieu_boc_tach'.")
        if any(not isinstance(row, dict) for row in rows):
            raise ValueError("Mỗi dòng trong 'du_lieu_boc_tach' phải là object JSON.")

        self.payload = payload
        self.rows = [dict(row) for row in rows]

    def _populate_rows(self) -> None:
        self._loading = True
        self.row_combo.clear()
        for index, row in enumerate(self.rows, start=1):
            self.row_combo.addItem(self._row_label(index, row), index - 1)
        self.row_combo.setEnabled(bool(self.rows))
        self._loading = False
        if self.rows:
            self.current_index = 0
            self._load_row(0)
        else:
            self.current_index = -1
            self.table.setRowCount(0)

    @staticmethod
    def _row_label(index: int, row: Dict[str, Any]) -> str:
        parts = [
            str(row.get("loai_chung_tu") or "Chứng từ"),
            str(row.get("so_container") or "").strip(),
            str(row.get("file_nguon") or "").strip(),
        ]
        suffix = " | ".join(part for part in parts if part)
        return f"{index}. {suffix}" if suffix else f"{index}. Dòng bóc tách"

    def _load_row(self, row_index: int) -> None:
        self._loading = True
        row = self.rows[row_index]
        self.table.setRowCount(len(self.FIELDS))
        for table_row, (key, label) in enumerate(self.FIELDS):
            name_item = QTableWidgetItem(label)
            name_item.setData(Qt.UserRole, key)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            value_item = QTableWidgetItem(self._display_value(row.get(key)))
            self.table.setItem(table_row, 0, name_item)
            self.table.setItem(table_row, 1, value_item)
        self._loading = False
        self._dirty = False

    @staticmethod
    def _display_value(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False, indent=2)
        return str(value)

    def _on_row_changed(self, index: int) -> None:
        if self._loading or index < 0:
            return
        target = int(self.row_combo.itemData(index))
        previous = self.current_index
        if previous >= 0 and not self._store_current_row():
            self.row_combo.blockSignals(True)
            self.row_combo.setCurrentIndex(previous)
            self.row_combo.blockSignals(False)
            return
        self.current_index = target
        self._load_row(target)

    def _store_current_row(self) -> bool:
        if self.current_index < 0 or self.current_index >= len(self.rows):
            return True
        updated = dict(self.rows[self.current_index])
        try:
            for table_row in range(self.table.rowCount()):
                name_item = self.table.item(table_row, 0)
                value_item = self.table.item(table_row, 1)
                if name_item is None:
                    continue
                key = str(name_item.data(Qt.UserRole) or "")
                text = value_item.text() if value_item else ""
                updated[key] = self._parse_value(key, text)
        except ValueError as exc:
            QMessageBox.warning(self, "Dữ liệu chưa hợp lệ", str(exc))
            return False
        self.rows[self.current_index] = updated
        return True

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

    def _mark_dirty(self) -> None:
        if not self._loading:
            self._dirty = True

    @Slot()
    def save(self) -> None:
        if not self._store_current_row():
            return
        self.payload["du_lieu_boc_tach"] = self.rows
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.payload, f, ensure_ascii=False, indent=2)
                f.write("\n")
        except OSError as exc:
            QMessageBox.critical(self, "Không lưu được JSON", str(exc))
            return
        self.saved = True
        self._dirty = False
        QMessageBox.information(self, "Đã lưu", "Đã lưu dữ liệu bóc tách JSON.")


class BillSelectionDialog(QDialog):
    """Popup chọn Bill và/hoặc dòng quyết toán khi dữ liệu bị mơ hồ."""

    def __init__(self, requests: List[BillChoiceRequest], parent=None):
        super().__init__(parent)
        self.requests = requests
        self.combos: Dict[str, QComboBox] = {}
        self.target_combos: Dict[str, QComboBox] = {}
        self.setWindowTitle("Chọn dữ liệu phù hợp")
        self.resize(1180, 460)

        layout = QVBoxLayout(self)
        hint = QLabel(
            "Một số container khớp nhiều Bill hoặc nhiều dòng quyết toán. "
            "Hãy chọn Bill và SQT đúng; nếu chưa chắc, chọn “Để xử lý sau”."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(len(requests), 6)
        table.setHorizontalHeaderLabels(
            [
                "Container",
                "Ngày đóng",
                "Số tấn",
                "Loại hàng",
                "Bill được chọn",
                "Dòng quyết toán",
            ]
        )
        table.setAlternatingRowColors(True)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        for row_index, request in enumerate(requests):
            table.setItem(row_index, 0, QTableWidgetItem(request.container))
            table.setItem(row_index, 1, QTableWidgetItem(_display_date(request.close_date)))
            table.setItem(
                row_index,
                2,
                QTableWidgetItem("" if request.tons is None else f"{request.tons:g}"),
            )
            table.setItem(row_index, 3, QTableWidgetItem(request.cargo))
            combo = QComboBox()
            if request.candidates:
                if len(request.candidates) > 1:
                    combo.addItem("Để xử lý sau", "__SKIP__")
                    combo.addItem("Bỏ qua container này", "__IGNORE__")
                for candidate in request.candidates:
                    label = " | ".join(
                        part
                        for part in (
                            candidate.bill_no or "Không có số Bill",
                            candidate.vessel,
                            _display_date(candidate.sail_date),
                            candidate.carrier,
                        )
                        if part
                    )
                    combo.addItem(label, candidate.md5)
                    combo.setItemData(
                        combo.count() - 1,
                        f"File: {candidate.source_name}\nMD5: {candidate.md5 or '—'}\n"
                        f"Seal: {candidate.seal or '—'}",
                        Qt.ToolTipRole,
                    )
            else:
                combo.addItem("Không có Bill để chọn", "__NO_BILL__")
                combo.setEnabled(False)
            table.setCellWidget(row_index, 4, combo)
            self.combos[request.subject_key] = combo

            target_combo = QComboBox()
            if request.target_candidates:
                if len(request.target_candidates) > 1:
                    target_combo.addItem("Để xử lý sau", "__SKIP__")
                for candidate in request.target_candidates:
                    label = " | ".join(
                        part
                        for part in (
                            f"SQT {candidate.sqt}",
                            f"Dòng {candidate.row_number}" if candidate.row_number else "",
                            _display_date(candidate.close_date),
                            candidate.cargo,
                            candidate.vessel,
                        )
                        if part
                    )
                    target_combo.addItem(label, f"sqt:{candidate.sqt}")
                if request.allow_new_target:
                    target_combo.addItem("Tạo dòng quyết toán mới", "__NEW__")
            else:
                target_combo.addItem("Tạo dòng quyết toán mới", "__NEW__")
                target_combo.setEnabled(False)
            table.setCellWidget(row_index, 5, target_combo)
            self.target_combos[request.target_subject_key] = target_combo
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            orientation=Qt.Horizontal,
        )
        buttons.button(QDialogButtonBox.Ok).setText("Tiếp tục")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy toàn bộ")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def decisions(self) -> Dict[str, str]:
        result = {
            key: str(combo.currentData() or "__SKIP__")
            for key, combo in self.combos.items()
        }
        result.update(
            {
                key: str(combo.currentData() or "__SKIP__")
                for key, combo in self.target_combos.items()
            }
        )
        return result


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


class _EditPendingDialog(QDialog):
    FIELDS = [
        ("close_date", "Ngày đóng / Ngày tháng"),
        ("container", "Số Container"),
        ("tons", "Số tấn"),
        ("cargo", "Loại hàng"),
        ("place", "Nơi đóng"),
        ("seal", "Số chì/Seal"),
        ("vessel", "Tên tàu"),
        ("sail_date", "Ngày chạy"),
        ("carrier", "VT biển"),
        ("invoice_no", "Số HĐ"),
        ("material_price", "Giá vật liệu"),
        ("unit_price", "Đơn giá"),
        ("amount", "Thành tiền"),
        ("vat", "VAT"),
        ("total", "Tổng tiền"),
    ]

    def __init__(self, item: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.item = item
        self.edits: Dict[str, QLineEdit] = {}
        self.setWindowTitle("Sửa dữ liệu chờ")
        self.resize(560, 620)
        self.setMinimumSize(460, 400)
        layout = QVBoxLayout(self)

        # Form đặt trong vùng cuộn để không bị chật/che khi có nhiều trường.
        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        form.setLabelAlignment(Qt.AlignRight)
        form.setContentsMargins(4, 4, 12, 4)
        form.setSpacing(8)
        data = item.get("data") or {}
        for key, label in self.FIELDS:
            value = data.get(key)
            if key in ("close_date", "sail_date"):
                value = _display_date(value)
            edit = QLineEdit("" if value is None else str(value))
            form.addRow(label + ":", edit)
            self.edits[key] = edit
        self.edit_sqt = QLineEdit(
            "" if item.get("matched_sqt") is None else str(item.get("matched_sqt"))
        )
        form.addRow("Ghép vào SQT PM:", self.edit_sqt)
        self.new_sqt_combo = QComboBox()
        self.new_sqt_combo.addItem("Tự động tìm dòng phù hợp", False)
        self.new_sqt_combo.addItem("Yêu cầu tạo SQT mới", True)
        if data.get("force_new_sqt"):
            self.new_sqt_combo.setCurrentIndex(1)
        form.addRow("Cách xử lý:", self.new_sqt_combo)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(form_host)
        layout.addWidget(scroll, stretch=1)

        advanced = QLabel(
            f"Loại dữ liệu: {item.get('document_type') or '—'}\n"
            f"MD5: {item.get('document_md5') or '—'}"
        )
        advanced.setWordWrap(True)
        advanced.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(advanced)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Lưu và kiểm tra lại")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def updated_data(self) -> Dict[str, Any]:
        data = dict(self.item.get("data") or {})
        for key, edit in self.edits.items():
            value = edit.text().strip()
            if key == "container":
                data[key] = normalize_container(value)
            elif key in ("close_date", "sail_date"):
                data[key] = parse_date(value)
            elif key in (
                "tons",
                "material_price",
                "unit_price",
                "amount",
                "vat",
                "total",
            ):
                data[key] = parse_number(value)
            elif key == "cargo":
                cargo, recognized = normalize_cargo(value)
                data[key] = cargo
                data["cargo_recognized"] = recognized
            else:
                data[key] = value
        data["preferred_sqt"] = (
            int(self.edit_sqt.text().strip())
            if self.edit_sqt.text().strip().isdigit()
            else None
        )
        data["force_new_sqt"] = bool(self.new_sqt_combo.currentData())
        return data

    def preferred_sqt(self) -> Optional[int]:
        text = self.edit_sqt.text().strip()
        return int(text) if text.isdigit() else None


class PendingDataDialog(QDialog):
    """Danh sách dữ liệu tạm; cho sửa, bỏ qua, khôi phục và match lại."""

    def __init__(self, database: Database, parent=None):
        super().__init__(parent)
        self.database = database
        self.retry_requested = False
        self.items: List[Dict[str, Any]] = []
        self.setWindowTitle("Dữ liệu chờ xử lý")
        self.resize(1120, 600)

        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Hiển thị:"))
        self.filter_combo = QComboBox()
        self.filter_combo.addItem("Tất cả dữ liệu chờ", "ACTIVE")
        for state, label in STATE_LABELS.items():
            self.filter_combo.addItem(label, state)
        self.filter_combo.currentIndexChanged.connect(self.refresh)
        top.addWidget(self.filter_combo)
        top.addStretch(1)
        layout.addLayout(top)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(
            [
                "Trạng thái",
                "Loại",
                "Ngày",
                "Container",
                "Số tấn",
                "Loại hàng",
                "Số Bill/HĐ",
                "SQT chọn",
                "Tàu",
                "Ghi chú",
            ]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        row = QHBoxLayout()
        btn_edit = QPushButton("Sửa dữ liệu")
        btn_ignore = QPushButton("Bỏ qua")
        btn_restore = QPushButton("Khôi phục")
        btn_retry = QPushButton("Chạy match lại")
        btn_close = QPushButton("Đóng")
        btn_edit.clicked.connect(self.edit_selected)
        btn_ignore.clicked.connect(self.ignore_selected)
        btn_restore.clicked.connect(self.restore_selected)
        btn_retry.clicked.connect(self.retry)
        btn_close.clicked.connect(self.accept)
        for button in (btn_edit, btn_ignore, btn_restore, btn_retry):
            row.addWidget(button)
        row.addStretch(1)
        row.addWidget(btn_close)
        layout.addLayout(row)
        self.refresh()

    def _selected(self) -> Optional[Dict[str, Any]]:
        row = self.table.currentRow()
        if row < 0 or row >= len(self.items):
            QMessageBox.information(self, "Chưa chọn dữ liệu", "Hãy chọn một dòng trước.")
            return None
        return self.items[row]

    @Slot()
    def refresh(self) -> None:
        selected_filter = str(self.filter_combo.currentData() or "ACTIVE")
        include_completed = selected_filter not in ("ACTIVE",)
        records = self.database.list_staged_rows(include_completed=include_completed)
        if selected_filter == "ACTIVE":
            records = [
                item
                for item in records
                if item.get("state") not in ("COMPLETED", "IGNORED")
            ]
        elif selected_filter:
            records = [item for item in records if item.get("state") == selected_filter]
        self.items = records
        self.table.setRowCount(len(records))
        for row_index, item in enumerate(records):
            data = item.get("data") or {}
            values = [
                STATE_LABELS.get(item.get("state"), item.get("state") or ""),
                item.get("document_type") or "",
                _display_date(data.get("close_date")),
                data.get("container") or "",
                "" if data.get("tons") is None else str(data.get("tons")),
                data.get("cargo") or "",
                data.get("bill_no") or data.get("invoice_no") or "",
                item.get("matched_sqt") or data.get("preferred_sqt") or "",
                data.get("vessel") or "",
                item.get("note") or "",
            ]
            for col, value in enumerate(values):
                table_item = QTableWidgetItem(str(value))
                if item.get("state") in ("MISSING_DATA", "CONFLICT", "NEEDS_BILL_SELECTION"):
                    table_item.setBackground(Qt.GlobalColor.yellow)
                self.table.setItem(row_index, col, table_item)

    @Slot()
    def edit_selected(self) -> None:
        item = self._selected()
        if not item:
            return
        dialog = _EditPendingDialog(item, self)
        if dialog.exec() == QDialog.Accepted:
            self.database.update_staged_row(
                int(item["id"]),
                state="PENDING",
                data=dialog.updated_data(),
                matched_sqt=dialog.preferred_sqt() or 0,
                note="Người dùng đã sửa; chờ chạy match lại.",
            )
            self.database.refresh_document_status(item.get("document_md5") or "")
            self.refresh()

    @Slot()
    def ignore_selected(self) -> None:
        item = self._selected()
        if not item:
            return
        answer = QMessageBox.question(
            self,
            "Bỏ qua dữ liệu",
            "Dữ liệu này sẽ không tự xuất hiện lại ở các lần nhập sau. Bạn vẫn có "
            "thể khôi phục trong mục “Đã bỏ qua”. Tiếp tục?",
        )
        if answer == QMessageBox.Yes:
            self.database.update_staged_row(
                int(item["id"]), state=STATE_IGNORED, note="Người dùng đã bỏ qua."
            )
            self.database.refresh_document_status(item.get("document_md5") or "")
            self.refresh()

    @Slot()
    def restore_selected(self) -> None:
        item = self._selected()
        if not item:
            return
        self.database.update_staged_row(
            int(item["id"]), state="PENDING", note="Đã khôi phục để xử lý lại."
        )
        self.database.refresh_document_status(item.get("document_md5") or "")
        self.refresh()

    @Slot()
    def retry(self) -> None:
        self.retry_requested = True
        self.accept()
