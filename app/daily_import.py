"""Phân tích và nhập output GPT vào file quyết toán hằng ngày.

Module này không phụ thuộc giao diện Qt. Giao diện chỉ gọi ``analyze`` để lấy
kế hoạch, thu thập lựa chọn của người dùng rồi gọi ``commit`` để ghi an toàn.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
from copy import copy
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import file_utils
from .database import Database


INFO_SHEET = "Thong_Tin_Quyet_Toan"
EXPENSE_SHEET = "Khoan_Chi"

INFO_HEADERS = [
    "SQT PM",
    "Ngày Đóng",
    "Số Container",
    "Biển số xe",
    "Số chì/Seal",
    "Lô SX",
    "Số tấn",
    "Loại hàng",
    "Nơi đóng",
    "Tên tàu",
    "Ngày chạy",
    "Dự kiến giao",
    "Người nhận",
    "VT biển",
    "Vận chuyển",
    "HD HP",
    "HD HCM",
    "Ngày Giao",
    "Hóa Đơn quyết toán",
    "Ghi chú",
    "Trạng thái nhập",
    "Ngày cập nhật",
    "MD5",
]

EXPENSE_HEADERS = [
    "SQT PM",
    "Ngày tháng",
    "Số Container",
    "Số HĐ",
    "Giá vật liệu",
    "Đơn giá",
    "Thành tiền",
    "VAT",
    "Tổng tiền",
    "Trạng thái nhập",
    "Ngày cập nhật",
    "MD5",
]

MD5_HEADER_ALIASES = ("MD5", "Mã MD5")

DOC_SCALE = "PHIEU_CAN"
DOC_BILL = "BILL"
DOC_EXPENSE = "KHOAN_CHI"

STATE_WAITING_BILL = "WAITING_BILL"
STATE_WAITING_SCALE = "WAITING_SCALE"
STATE_NEEDS_BILL = "NEEDS_BILL_SELECTION"
STATE_MISSING = "MISSING_DATA"
STATE_CONFLICT = "CONFLICT"
STATE_READY = "READY"
STATE_COMPLETED = "COMPLETED"
STATE_IGNORED = "IGNORED"

STATE_LABELS = {
    STATE_WAITING_BILL: "Chờ Bill",
    STATE_WAITING_SCALE: "Chờ Phiếu cân",
    STATE_NEEDS_BILL: "Cần chọn Bill",
    STATE_MISSING: "Thiếu dữ liệu",
    STATE_CONFLICT: "Xung đột dữ liệu",
    STATE_READY: "Sẵn sàng nhập",
    STATE_COMPLETED: "Đã hoàn tất",
    STATE_IGNORED: "Đã bỏ qua",
    "PENDING": "Đang chờ xử lý",
    "FAILED": "Xử lý lỗi",
}

_ACCENTED_VIETNAMESE = re.compile(
    r"[ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊ"
    r"ÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]",
    re.IGNORECASE,
)

DEFAULT_CARGO_MAPPINGS = {
    "VOI": "Vôi",
    "VOI ROI": "Vôi rời",
    "VOI BOT": "Vôi bột",
    "VOI BOT NONG NGHIEP": "Vôi bột nông nghiệp",
    "BOT NN BAO 25 KG": "Bột NN bao 25 kg",
}


class DailyImportError(RuntimeError):
    """Lỗi nghiệp vụ có nội dung phù hợp để hiển thị cho người dùng."""


@dataclass
class ExtractedRow:
    staged_id: int = 0
    source_row: int = 0
    row_key: str = ""
    source_name: str = ""
    md5: str = ""
    doc_type: str = ""
    source_status: str = ""
    close_date: Optional[str] = None
    container: str = ""
    tons: Optional[float] = None
    cargo: str = ""
    cargo_recognized: bool = True
    place: str = ""
    recipient: str = ""
    truck_no: str = ""
    vessel: str = ""
    sail_date: Optional[str] = None
    carrier: str = ""
    material_price: Optional[float] = None
    invoice_no: str = ""
    bill_no: str = ""
    seal: str = ""
    unit_price: Optional[float] = None
    amount: Optional[float] = None
    vat: Optional[float] = None
    total: Optional[float] = None
    other: str = ""
    confidence: str = ""
    warning: str = ""
    preferred_sqt: Optional[int] = None
    force_new_sqt: bool = False
    md5_is_synthetic: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any], staged_id: int = 0) -> "ExtractedRow":
        allowed = cls.__dataclass_fields__.keys()
        values = {key: data.get(key) for key in allowed if key in data}
        values["staged_id"] = staged_id or int(values.get("staged_id") or 0)
        return cls(**values)


@dataclass
class BillCandidate:
    staged_id: int
    md5: str
    source_name: str
    bill_no: str
    container: str
    seal: str
    vessel: str
    sail_date: Optional[str]
    carrier: str


@dataclass
class TargetCandidate:
    row_number: int
    sqt: int
    close_date: Optional[str]
    container: str
    cargo: str
    vessel: str = ""


@dataclass
class BillChoiceRequest:
    subject_key: str
    container: str
    close_date: Optional[str]
    tons: Optional[float]
    cargo: str
    candidates: List[BillCandidate]
    target_candidates: List[TargetCandidate] = field(default_factory=list)
    allow_new_target: bool = True

    @property
    def target_subject_key(self) -> str:
        return f"{self.subject_key}:target"


@dataclass
class FieldConflict:
    conflict_id: str
    target_row: int
    field_name: str
    current_value: Any
    new_value: Any


@dataclass
class InfoChange:
    action: str
    target_row: Optional[int]
    sqt: int
    values: Dict[str, Any]
    staged_ids: List[int]
    document_md5s: List[str]
    check_status: str = "OK"
    conflicts: List[FieldConflict] = field(default_factory=list)


@dataclass
class ExpenseChange:
    action: str
    target_row: Optional[int]
    values: Dict[str, Any]
    staged_ids: List[int]
    document_md5s: List[str]
    check_status: str = "OK"


@dataclass
class ImportAnalysis:
    output_path: str
    daily_path: str
    processed_file_id: Optional[int]
    extracted_rows: int = 0
    duplicate_documents: int = 0
    info_changes: List[InfoChange] = field(default_factory=list)
    expense_changes: List[ExpenseChange] = field(default_factory=list)
    bill_choices: List[BillChoiceRequest] = field(default_factory=list)
    conflicts: List[FieldConflict] = field(default_factory=list)
    pending_count: int = 0
    pending_states: Dict[int, Tuple[str, str]] = field(default_factory=dict)
    completed_staged_ids: List[int] = field(default_factory=list)

    @property
    def new_info_count(self) -> int:
        return sum(change.action == "CREATE" for change in self.info_changes)

    @property
    def updated_info_count(self) -> int:
        return sum(change.action == "UPDATE" for change in self.info_changes)


@dataclass
class ImportSummary:
    new_info: int = 0
    updated_info: int = 0
    new_expenses: int = 0
    updated_expenses: int = 0
    duplicate_documents: int = 0
    pending_count: int = 0
    backup_path: str = ""
    status: str = "DAILY_IMPORTED"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class _TargetInfoRow:
    row_number: int
    sqt: int
    close_date: Optional[str]
    container: str
    cargo: str
    values: Dict[str, Any]


@dataclass
class _WorkbookSnapshot:
    info_rows: List[_TargetInfoRow]
    expense_rows: List[Tuple[int, Dict[str, Any]]]
    target_md5: set[str]
    max_sqt: int


def _strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFD", value or "")
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def _key_text(value: Any) -> str:
    text = _strip_accents(str(value or "")).upper().replace("Đ", "D")
    return re.sub(r"\s+", " ", text).strip()


def normalize_container(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_valid_container(value: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{4}\d{7}", value or ""))


def normalize_md5(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").lower())


def parse_date(value: Any) -> Optional[str]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def excel_date(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        parsed = date.fromisoformat(value)
    except ValueError:
        return None
    return datetime(parsed.year, parsed.month, parsed.day)


def parse_number(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(" ", "")
    text = re.sub(r"[^0-9,.-]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        text = "".join(parts) if len(parts[-1]) == 3 else text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def normalize_cargo(
    value: Any, mappings: Optional[Dict[str, str]] = None
) -> Tuple[str, bool]:
    original = re.sub(r"\s+", " ", str(value or "")).strip()
    if not original:
        return "", False
    mapping = dict(DEFAULT_CARGO_MAPPINGS)
    mapping.update(mappings or {})
    key = _key_text(original)
    normalized_mapping = {_key_text(k): v for k, v in mapping.items()}
    if key in normalized_mapping:
        return normalized_mapping[key], True
    if _ACCENTED_VIETNAMESE.search(original):
        return original, True
    replacements = [
        (r"^VOI BOT NONG NGHIEP\b", "Vôi bột nông nghiệp"),
        (r"^VOI CUC\b", "Vôi cục"),
        (r"^VOI CANXI\b", "Vôi canxi"),
        (r"^VOI ROI\b", "Vôi rời"),
        (r"^VOI BOT\b", "Vôi bột"),
        (r"^VOI\b", "Vôi"),
    ]
    for pattern, replacement in replacements:
        if re.search(pattern, key):
            tail = re.sub(pattern, "", key).strip()
            tail = tail.replace(" XA ROI", " xá rời").replace(" CO CHAN", " có chặn")
            return (replacement + (" " + tail.lower() if tail else "")).strip(), not bool(tail)
    return original, False


def _doc_type(value: Any) -> str:
    key = _key_text(value)
    if "PHIEU CAN" in key:
        return DOC_SCALE
    if key == "BILL" or "BILL OF LADING" in key:
        return DOC_BILL
    if "KHOAN CHI" in key or "CHI PHI" in key:
        return DOC_EXPENSE
    return ""


def _join_md5(*values: Any) -> str:
    result: List[str] = []
    for value in values:
        for part in re.split(r"[;,\n]+", str(value or "")):
            md5 = normalize_md5(part)
            if md5 and md5 not in result:
                result.append(md5)
    return ";".join(result)


def _md5_cell_value(values: Dict[str, Any]) -> Any:
    for header in MD5_HEADER_ALIASES:
        if values.get(header) not in (None, ""):
            return values.get(header)
    return ""


def _normalize_import_status(value: Any, default: str = "Chưa nhập") -> str:
    key = _key_text(value)
    if key in ("DA NHAP", "HOAN THANH"):
        return "Đã nhập"
    if key == "CHUA NHAP":
        return "Chưa nhập"
    return default


def _parse_target_decision(value: Any) -> Optional[int]:
    text = str(value or "").strip()
    if text.startswith("sqt:"):
        text = text[4:]
    return int(text) if text.isdigit() else None


def _stable_row_key(row: ExtractedRow, occurrence: int = 0) -> str:
    identity = {
        "type": row.doc_type,
        "container": row.container,
        "close_date": row.close_date,
        "sail_date": row.sail_date,
        "bill_no": row.bill_no,
        "invoice_no": row.invoice_no,
        "seal": row.seal,
        "cargo": row.cargo,
        "tons": row.tons,
        "amount": row.amount,
        "total": row.total,
        "occurrence": occurrence,
    }
    raw = json.dumps(identity, ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


class DailyImportService:
    def __init__(
        self,
        database: Database,
        logger,
        cargo_mappings: Optional[Dict[str, str]] = None,
    ):
        self.database = database
        self.logger = logger
        self.cargo_mappings = cargo_mappings or {}

    # ------------------------------------------------------------------ #
    # Phân tích
    # ------------------------------------------------------------------ #
    def analyze(
        self,
        output_path: str,
        daily_path: str,
        processed_file_id: Optional[int] = None,
        bill_decisions: Optional[Dict[str, str]] = None,
    ) -> ImportAnalysis:
        if not os.path.isfile(output_path):
            raise DailyImportError("Không tìm thấy bản Output đã kiểm tra.")
        if not os.path.isfile(daily_path):
            raise DailyImportError(
                "Không tìm thấy file theo dõi hàng ngày. Vui lòng kiểm tra Cài đặt."
            )
        if file_utils.is_file_locked(daily_path):
            raise DailyImportError(
                "File theo dõi hàng ngày đang mở. Vui lòng lưu và đóng Excel rồi thử lại."
            )

        parsed = self._read_output(output_path)
        snapshot = self._read_daily_snapshot(daily_path)
        analysis = ImportAnalysis(
            output_path=output_path,
            daily_path=daily_path,
            processed_file_id=processed_file_id,
            extracted_rows=len(parsed),
        )

        statuses = self.database.get_document_statuses(
            sorted({row.md5 for row in parsed if row.md5})
        )
        occurrences: Dict[Tuple[str, str], int] = {}
        for row in parsed:
            if not row.md5:
                continue
            key = (row.md5, _stable_row_key(row))
            occurrences[key] = occurrences.get(key, 0) + 1
            row.row_key = _stable_row_key(row, occurrences[key] - 1)
            self.database.upsert_source_document(
                row.md5,
                row.doc_type,
                row.source_name,
                processed_file_id,
            )
            row.staged_id = self.database.upsert_staged_row(
                row.md5,
                row.row_key,
                row.source_row,
                row.doc_type,
                row.container,
                row.to_dict(),
            )

        duplicate_md5 = {
            row.md5
            for row in parsed
            if row.md5
            and not row.md5_is_synthetic
            and (
                statuses.get(row.md5) in ("COMPLETED", "IGNORED")
                or (row.md5 in snapshot.target_md5 and row.md5 not in statuses)
            )
        }
        analysis.duplicate_documents = len(duplicate_md5)

        active_records = self.database.list_staged_rows(include_completed=False)
        active = [
            ExtractedRow.from_dict(item.get("data") or {}, int(item["id"]))
            for item in active_records
        ]
        for row, item in zip(active, active_records):
            row.row_key = item.get("row_key") or row.row_key
            row.md5 = item.get("document_md5") or row.md5
            row.preferred_sqt = (
                int(item.get("matched_sqt"))
                if item.get("matched_sqt") is not None
                else row.preferred_sqt
            )

        saved_decisions = self.database.get_match_decisions()
        decisions = {**saved_decisions, **(bill_decisions or {})}
        bills_by_container: Dict[str, List[ExtractedRow]] = {}
        scales: List[ExtractedRow] = []
        expenses: List[ExtractedRow] = []
        for row in active:
            if row.doc_type == DOC_BILL:
                bills_by_container.setdefault(row.container, []).append(row)
            elif row.doc_type == DOC_SCALE:
                scales.append(row)
            elif row.doc_type == DOC_EXPENSE:
                expenses.append(row)

        next_sqt = snapshot.max_sqt + 1
        used_bills: set[int] = set()
        resolved_bill_containers: set[str] = set()
        choice_pending_containers: set[str] = set()
        virtual_rows = list(snapshot.info_rows)

        for scale in scales:
            subject_key = f"scale:{scale.row_key}"
            candidates = self._distinct_bills(
                row
                for row in bills_by_container.get(scale.container, [])
            )
            target_options = self._scale_target_options(virtual_rows, scale)
            target, target_waiting = self._resolve_target_choice(
                subject_key,
                scale,
                target_options,
                decisions,
                analysis,
                bill_candidates=candidates,
                allow_new=True,
            )
            if target_waiting:
                choice_pending_containers.add(scale.container)
                continue

            selected_bill: Optional[ExtractedRow] = None
            if len(candidates) == 1:
                selected_bill = candidates[0]
            elif len(candidates) > 1:
                selected_md5 = decisions.get(subject_key)
                if selected_md5 == "__SKIP__":
                    analysis.pending_states[scale.staged_id] = (
                        STATE_NEEDS_BILL,
                        "Người dùng chọn để xử lý sau.",
                    )
                    continue
                if selected_md5 == "__IGNORE__":
                    analysis.pending_states[scale.staged_id] = (
                        STATE_IGNORED,
                        "Người dùng bỏ qua container này.",
                    )
                    continue
                selected_bill = next(
                    (item for item in candidates if item.md5 == selected_md5), None
                )
                if selected_bill is None:
                    analysis.bill_choices.append(
                        self._bill_choice(
                            subject_key,
                            scale,
                            candidates,
                            target_options,
                        )
                    )
                    choice_pending_containers.add(scale.container)
                    analysis.pending_states[scale.staged_id] = (
                        STATE_NEEDS_BILL,
                        "Cần chọn Bill hoặc dòng SQT phù hợp.",
                    )
                    for bill in candidates:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_NEEDS_BILL,
                            "Chờ người dùng chọn Bill phù hợp.",
                        )
                    continue

            fallback = target.values if target else (
                target_options[0].values if len(target_options) == 1 else {}
            )
            sqt = target.sqt if target else next_sqt
            action = "UPDATE" if target else "CREATE"
            if not target:
                next_sqt += 1
            previous_same_container = any(
                item.container == scale.container
                and item.close_date != scale.close_date
                for item in snapshot.info_rows
            )
            check_status = (
                "Cần kiểm tra"
                if previous_same_container
                or scale.force_new_sqt
                or not scale.cargo_recognized
                # Tạo QT mới mà không có Bill (thông tin tàu nhập tay) -> đánh dấu
                # để người dùng soát lại cho an toàn.
                or (not selected_bill and not target)
                or self._source_needs_review(scale, selected_bill)
                else "OK"
            )
            values = self._info_values(
                scale,
                selected_bill,
                sqt,
                fallback,
                check_status,
            )
            conflicts = self._find_conflicts(target, values) if target else []
            if conflicts:
                check_status = "Cần kiểm tra"
                analysis.conflicts.extend(conflicts)
            staged_ids = [scale.staged_id]
            md5_values = [self._real_md5(scale)]
            if selected_bill:
                staged_ids.append(selected_bill.staged_id)
                md5_values.append(self._real_md5(selected_bill))
                used_bills.add(selected_bill.staged_id)
                resolved_bill_containers.add(scale.container)
                analysis.pending_states[selected_bill.staged_id] = (STATE_READY, "")
            analysis.info_changes.append(
                InfoChange(
                    action=action,
                    target_row=target.row_number if target else None,
                    sqt=sqt,
                    values=values,
                    staged_ids=staged_ids,
                    document_md5s=md5_values,
                    check_status=check_status,
                    conflicts=conflicts,
                )
            )
            analysis.pending_states[scale.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.extend(staged_ids)
            if action == "CREATE":
                virtual_rows.append(
                    _TargetInfoRow(
                        row_number=0,
                        sqt=sqt,
                        close_date=scale.close_date,
                        container=scale.container,
                        cargo=scale.cargo,
                        values=values,
                    )
                )

        # Bill có thể bổ sung trực tiếp cho một dòng quyết toán đã tồn tại.
        for container, raw_candidates in bills_by_container.items():
            if container in choice_pending_containers:
                continue
            remaining = [
                row
                for row in raw_candidates
                if row.staged_id not in used_bills
            ]
            if not remaining:
                continue
            if container in resolved_bill_containers:
                for bill in remaining:
                    analysis.pending_states[bill.staged_id] = (
                        STATE_NEEDS_BILL,
                        "Bill khác chưa được chọn cho Container này.",
                    )
                continue
            target_rows = [row for row in virtual_rows if row.container == container]
            preferred_sqt = next(
                (row.preferred_sqt for row in remaining if row.preferred_sqt), None
            )
            if preferred_sqt:
                target_rows = [row for row in target_rows if row.sqt == preferred_sqt]
            candidates = self._distinct_bills(remaining)
            subject_key = f"target:{container}:{preferred_sqt or 'auto'}"
            target_dates = {row.close_date for row in target_rows if row.close_date}
            container_decision = None
            if len(target_dates) == 1:
                container_decision = decisions.get(
                    f"container:{container}:{next(iter(target_dates))}"
                )
            if (
                container_decision
                and container_decision not in ("__NEW__", "__SKIP__", "__IGNORE__")
                and not str(container_decision).startswith("sqt:")
                and all(
                    item.md5 != container_decision for item in candidates
                )
            ):
                for bill in remaining:
                    analysis.pending_states[bill.staged_id] = (
                        STATE_WAITING_SCALE,
                        "Bill này không được chọn cho dòng hiện có; chờ Phiếu cân khác.",
                    )
                continue
            selected_bill = next(
                (item for item in candidates if item.md5 == container_decision),
                candidates[0] if len(candidates) == 1 else None,
            )
            if len(candidates) > 1:
                selected_md5 = decisions.get(subject_key)
                if selected_md5 == "__SKIP__":
                    for bill in remaining:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_NEEDS_BILL,
                            "Người dùng chọn để xử lý sau.",
                        )
                    continue
                if selected_md5 == "__IGNORE__":
                    for bill in remaining:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_IGNORED,
                            "Người dùng bỏ qua container này.",
                        )
                    continue
                selected_bill = next(
                    (item for item in candidates if item.md5 == selected_md5), None
                )
                if selected_bill is None:
                    seed_close_date = target_rows[0].close_date if target_rows else None
                    seed_tons = (
                        parse_number(target_rows[0].values.get("Số tấn"))
                        if target_rows
                        else None
                    )
                    seed_cargo = (
                        str(target_rows[0].values.get("Loại hàng") or "")
                        if target_rows
                        else ""
                    )
                    seed = ExtractedRow(
                        container=container,
                        close_date=seed_close_date,
                        tons=seed_tons,
                        cargo=seed_cargo,
                    )
                    analysis.bill_choices.append(
                        self._bill_choice(subject_key, seed, candidates, target_rows)
                    )
                    for bill in remaining:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_NEEDS_BILL,
                            "Chờ người dùng chọn Bill phù hợp.",
                        )
                    continue
            if selected_bill is None:
                continue

            target, target_waiting = self._resolve_target_choice(
                subject_key,
                selected_bill,
                target_rows,
                decisions,
                analysis,
                bill_candidates=candidates,
                allow_new=True,
            )
            if target_waiting:
                continue
            action = "UPDATE" if target else "CREATE"
            sqt = target.sqt if target else next_sqt
            if not target:
                next_sqt += 1
            fallback = target.values if target else {}
            values = self._bill_info_values(selected_bill, sqt, fallback)
            conflicts = self._find_conflicts(target, values) if target else []
            analysis.conflicts.extend(conflicts)
            analysis.info_changes.append(
                InfoChange(
                    action=action,
                    target_row=target.row_number if target else None,
                    sqt=sqt,
                    values=values,
                    staged_ids=[selected_bill.staged_id],
                    document_md5s=[self._real_md5(selected_bill)],
                    check_status="OK",
                    conflicts=conflicts,
                )
            )
            if action == "CREATE":
                virtual_rows.append(
                    _TargetInfoRow(
                        row_number=0,
                        sqt=sqt,
                        close_date=selected_bill.close_date,
                        container=selected_bill.container,
                        cargo=selected_bill.cargo,
                        values=values,
                    )
                )
            used_bills.add(selected_bill.staged_id)
            analysis.pending_states[selected_bill.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.append(selected_bill.staged_id)
            for bill in remaining:
                if bill.staged_id != selected_bill.staged_id:
                    analysis.pending_states[bill.staged_id] = (
                        STATE_NEEDS_BILL,
                        "Bill không được chọn cho dòng quyết toán này.",
                    )

        self._analyze_expenses(expenses, virtual_rows, snapshot, analysis)

        for item in active_records:
            row_id = int(item["id"])
            state, note = analysis.pending_states.get(
                row_id, (item.get("state") or "PENDING", item.get("note") or "")
            )
            self.database.update_staged_row(row_id, state=state, note=note)
        for md5 in {item.get("document_md5") or "" for item in active_records}:
            if md5:
                self.database.refresh_document_status(md5)
        analysis.pending_count = sum(
            item.get("state") not in (STATE_READY, STATE_COMPLETED, STATE_IGNORED)
            for item in self.database.list_staged_rows(include_completed=False)
        )
        return analysis

    # ------------------------------------------------------------------ #
    # Ghi workbook
    # ------------------------------------------------------------------ #
    def commit(
        self,
        analysis: ImportAnalysis,
        conflict_decisions: Optional[Dict[str, bool]] = None,
    ) -> ImportSummary:
        conflict_decisions = conflict_decisions or {}
        daily_path = analysis.daily_path
        if file_utils.is_file_locked(daily_path):
            raise DailyImportError(
                "File theo dõi hàng ngày đang mở. Vui lòng đóng Excel rồi thử lại."
            )
        run_id = self.database.create_import_run(
            analysis.processed_file_id, analysis.output_path, daily_path
        )
        backup_path = ""
        temp_path = ""
        try:
            wb = load_workbook(daily_path)
            self._upgrade_workbook(wb)
            info_ws = wb[INFO_SHEET]
            expense_ws = wb[EXPENSE_SHEET]
            info_map = self._header_map(info_ws)
            expense_map = self._header_map(expense_ws)

            for change in analysis.info_changes:
                row_number = (
                    change.target_row
                    if change.action == "UPDATE" and change.target_row
                    else info_ws.max_row + 1
                )
                if change.action == "CREATE":
                    self._copy_row_style(info_ws, max(2, info_ws.max_row), row_number)
                conflict_by_field = {
                    conflict.field_name: conflict for conflict in change.conflicts
                }
                for name, value in change.values.items():
                    if name not in info_map:
                        continue
                    cell = info_ws.cell(row_number, info_map[name])
                    conflict = conflict_by_field.get(name)
                    if conflict:
                        if conflict_decisions.get(conflict.conflict_id, False):
                            cell.value = value
                    elif change.action == "CREATE" or cell.value in (None, "") or name in (
                        "MD5",
                        "Ngày cập nhật",
                    ):
                        cell.value = value

            for change in analysis.expense_changes:
                row_number = (
                    change.target_row
                    if change.action == "UPDATE" and change.target_row
                    else expense_ws.max_row + 1
                )
                if change.action == "CREATE":
                    self._copy_row_style(expense_ws, max(2, expense_ws.max_row), row_number)
                for name, value in change.values.items():
                    if name not in expense_map:
                        continue
                    cell = expense_ws.cell(row_number, expense_map[name])
                    if change.action == "CREATE" or cell.value in (None, "") or name in (
                        "MD5",
                        "Ngày cập nhật",
                    ):
                        cell.value = value

            self._format_workbook(wb)
            suffix = Path(daily_path).suffix or ".xlsx"
            temp_path = str(Path(daily_path).with_name(f".__daily_import_{os.getpid()}{suffix}"))
            wb.save(temp_path)
            wb.close()
            self._validate_saved_workbook(temp_path)
            os.replace(temp_path, daily_path)

            completed_ids = sorted(set(analysis.completed_staged_ids))
            affected_md5: set[str] = set()
            for row_id in completed_ids:
                item = self.database.get_staged_row(row_id)
                if not item:
                    continue
                affected_md5.add(item.get("document_md5") or "")
                self.database.update_staged_row(row_id, state=STATE_COMPLETED)
            for md5 in affected_md5:
                if md5:
                    self.database.refresh_document_status(md5)

            pending = self.database.count_pending_rows()
            status = "DAILY_IMPORT_PARTIAL" if pending else "DAILY_IMPORTED"
            summary = ImportSummary(
                new_info=analysis.new_info_count,
                updated_info=analysis.updated_info_count,
                new_expenses=sum(c.action == "CREATE" for c in analysis.expense_changes),
                updated_expenses=sum(c.action == "UPDATE" for c in analysis.expense_changes),
                duplicate_documents=analysis.duplicate_documents,
                pending_count=pending,
                backup_path=backup_path,
                status=status,
            )
            self.database.finish_import_run(
                run_id, "COMPLETED", backup_path=backup_path, summary=summary.to_dict()
            )
            if analysis.processed_file_id is not None:
                self.database.update_status(
                    analysis.processed_file_id,
                    status,
                    note=self._summary_note(summary),
                )
            return summary
        except Exception as exc:
            # Ghi theo cơ chế file tạm + os.replace (nguyên tử) nên file theo dõi
            # gốc không bị hỏng dở; chỉ cần dọn file tạm nếu lỗi.
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            self.database.finish_import_run(
                run_id,
                "FAILED",
                backup_path=backup_path or None,
                error_message=str(exc),
            )
            if analysis.processed_file_id is not None:
                self.database.update_status(
                    analysis.processed_file_id,
                    "DAILY_IMPORT_ERROR",
                    note=f"Lỗi nhập file theo dõi: {exc}",
                )
            if isinstance(exc, DailyImportError):
                raise
            raise DailyImportError(f"Không cập nhật được file theo dõi: {exc}") from exc

    # ------------------------------------------------------------------ #
    # Đọc dữ liệu
    # ------------------------------------------------------------------ #
    def _read_output(self, path: str) -> List[ExtractedRow]:
        if Path(path).suffix.lower() != ".json":
            raise DailyImportError("File bóc tách phải là định dạng JSON (.json).")

        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                payload = json.load(f)
        except json.JSONDecodeError as exc:
            raise DailyImportError(f"File JSON bóc tách không hợp lệ: {exc}") from exc
        except UnicodeDecodeError as exc:
            raise DailyImportError(
                "File JSON bóc tách không đọc được bằng UTF-8. Vui lòng xuất/tải lại file."
            ) from exc
        except OSError as exc:
            raise DailyImportError(f"Không đọc được file JSON bóc tách: {exc}") from exc

        if not isinstance(payload, dict):
            raise DailyImportError("File JSON bóc tách phải có object gốc.")
        rows = payload.get("du_lieu_boc_tach")
        if not isinstance(rows, list):
            raise DailyImportError("File JSON thiếu mảng 'du_lieu_boc_tach'.")

        result: List[ExtractedRow] = []
        for index, item in enumerate(rows, start=1):
            if not isinstance(item, dict):
                raise DailyImportError(
                    f"Dòng du_lieu_boc_tach #{index} phải là object JSON."
                )
            if not any(value not in (None, "", [], {}) for value in item.values()):
                continue

            doc_type = _doc_type(item.get("loai_chung_tu"))
            if not doc_type:
                continue

            cargo, recognized = normalize_cargo(
                item.get("loai_hang"),
                self.cargo_mappings,
            )
            close_date = parse_date(item.get("ngay_dong"))
            sail_date = parse_date(item.get("ngay_chay"))
            source_row_number = parse_number(item.get("stt_hien_thi"))
            row = ExtractedRow(
                source_row=int(source_row_number or index),
                source_name=str(item.get("file_nguon") or ""),
                md5=normalize_md5(item.get("ma_md5_file")),
                doc_type=doc_type,
                source_status=str(item.get("trang_thai") or ""),
                close_date=close_date,
                container=normalize_container(item.get("so_container")),
                tons=parse_number(item.get("so_tan")),
                cargo=cargo,
                cargo_recognized=recognized,
                place=str(item.get("noi_dong") or "").strip(),
                recipient=str(item.get("nguoi_nhan") or "").strip(),
                truck_no=str(item.get("bien_so_xe") or "").strip(),
                vessel=str(item.get("ten_tau") or "").strip(),
                sail_date=sail_date,
                carrier=str(item.get("vt_bien") or "").strip(),
                material_price=parse_number(item.get("gia_vat_lieu")),
                invoice_no=str(item.get("so_hd") or "").strip(),
                bill_no=str(item.get("so_bill") or "").strip(),
                seal=str(item.get("so_chi_seal") or "").strip(),
                unit_price=parse_number(item.get("don_gia")),
                amount=parse_number(item.get("thanh_tien")),
                vat=parse_number(item.get("vat")),
                total=parse_number(item.get("tong_tien")),
                other=self._json_cell_text(item.get("truong_khac")),
                confidence=str(item.get("do_tin_cay") or "").strip(),
                warning=self._json_warning_text(item.get("canh_bao")),
            )
            if doc_type == DOC_EXPENSE:
                row.close_date = sail_date or close_date
            if not row.md5:
                # Không có MD5 thì vẫn theo dõi được trong danh sách chờ, nhưng
                # dùng khóa cục bộ để không làm mất dòng.
                row.md5 = "missing-" + hashlib.sha1(
                    f"{path}|{index}|{row.container}".encode("utf-8")
                ).hexdigest()
                row.md5_is_synthetic = True
            result.append(row)
        return result

    @staticmethod
    def _json_cell_text(value: Any) -> str:
        if value in (None, ""):
            return ""
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False, sort_keys=True)
        return str(value).strip()

    @classmethod
    def _json_warning_text(cls, value: Any) -> str:
        if value in (None, ""):
            return ""
        if isinstance(value, list):
            return "; ".join(
                cls._json_cell_text(item) for item in value if item not in (None, "")
            )
        return cls._json_cell_text(value)

    def _read_daily_snapshot(self, path: str) -> _WorkbookSnapshot:
        try:
            wb = load_workbook(path, read_only=True, data_only=False)
        except Exception as exc:
            raise DailyImportError(f"Không đọc được file theo dõi hàng ngày: {exc}") from exc
        try:
            for sheet_name in (INFO_SHEET, EXPENSE_SHEET):
                if sheet_name not in wb.sheetnames:
                    raise DailyImportError(
                        f"File theo dõi thiếu sheet '{sheet_name}'."
                    )
            info_ws = wb[INFO_SHEET]
            expense_ws = wb[EXPENSE_SHEET]
            info_map = self._validate_daily_headers(info_ws, INFO_HEADERS)
            expense_map = self._validate_daily_headers(expense_ws, EXPENSE_HEADERS)
            info_rows: List[_TargetInfoRow] = []
            target_md5: set[str] = set()
            max_sqt = 0
            for row_number in range(2, info_ws.max_row + 1):
                values = {
                    name: info_ws.cell(row_number, col).value
                    for name, col in info_map.items()
                }
                if not any(value not in (None, "") for value in values.values()):
                    continue
                sqt_value = parse_number(values.get("SQT PM"))
                sqt = int(sqt_value or 0)
                max_sqt = max(max_sqt, sqt)
                container = normalize_container(values.get("Số Container"))
                cargo, _ = normalize_cargo(values.get("Loại hàng"), self.cargo_mappings)
                info_rows.append(
                    _TargetInfoRow(
                        row_number=row_number,
                        sqt=sqt,
                        close_date=parse_date(values.get("Ngày Đóng")),
                        container=container,
                        cargo=cargo,
                        values=values,
                    )
                )
                target_md5.update(
                    normalize_md5(part)
                    for part in re.split(r"[;,\n]+", str(_md5_cell_value(values) or ""))
                    if normalize_md5(part)
                )
            expense_rows: List[Tuple[int, Dict[str, Any]]] = []
            for row_number in range(2, expense_ws.max_row + 1):
                values = {
                    name: expense_ws.cell(row_number, col).value
                    for name, col in expense_map.items()
                }
                if not any(value not in (None, "") for value in values.values()):
                    continue
                expense_rows.append((row_number, values))
                target_md5.update(
                    normalize_md5(part)
                    for part in re.split(r"[;,\n]+", str(_md5_cell_value(values) or ""))
                    if normalize_md5(part)
                )
            return _WorkbookSnapshot(info_rows, expense_rows, target_md5, max_sqt)
        finally:
            wb.close()

    # ------------------------------------------------------------------ #
    # Matching helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _target_group(
        rows: Sequence[_TargetInfoRow], close_date: Optional[str], container: str
    ) -> List[_TargetInfoRow]:
        return [
            row
            for row in rows
            if row.close_date == close_date and row.container == container
        ]

    @staticmethod
    def _real_md5(row: ExtractedRow) -> str:
        return "" if row.md5_is_synthetic or row.md5.startswith("missing-") else row.md5

    def _scale_target_options(
        self, rows: Sequence[_TargetInfoRow], scale: ExtractedRow
    ) -> List[_TargetInfoRow]:
        if scale.force_new_sqt:
            return []
        if scale.preferred_sqt:
            preferred = [
                row
                for row in rows
                if row.sqt == scale.preferred_sqt
                and (not scale.container or row.container == scale.container)
            ]
            if preferred:
                return preferred
        group = self._target_group(rows, scale.close_date, scale.container)
        if not group:
            return []
        if scale.cargo and any(row.cargo for row in group):
            exact = [
                row for row in group if _key_text(row.cargo) == _key_text(scale.cargo)
            ]
            return exact
        return group

    def _resolve_target_choice(
        self,
        subject_key: str,
        subject: ExtractedRow,
        candidates: Sequence[_TargetInfoRow],
        decisions: Dict[str, str],
        analysis: ImportAnalysis,
        *,
        bill_candidates: Sequence[ExtractedRow] = (),
        allow_new: bool = True,
    ) -> Tuple[Optional[_TargetInfoRow], bool]:
        if not candidates:
            return None, False
        if len(candidates) == 1:
            return candidates[0], False

        target_key = f"{subject_key}:target"
        selected = decisions.get(target_key)
        if selected == "__SKIP__":
            analysis.pending_states[subject.staged_id] = (
                STATE_NEEDS_BILL,
                "Người dùng chọn để xử lý sau.",
            )
            return None, True
        if selected == "__NEW__" and allow_new:
            return None, False
        selected_sqt = _parse_target_decision(selected)
        if selected_sqt is not None:
            target = next((row for row in candidates if row.sqt == selected_sqt), None)
            if target:
                return target, False

        analysis.bill_choices.append(
            self._bill_choice(subject_key, subject, bill_candidates, candidates, allow_new)
        )
        analysis.pending_states[subject.staged_id] = (
            STATE_NEEDS_BILL,
            "Container khớp nhiều dòng quyết toán; cần chọn SQT.",
        )
        return None, True

    @staticmethod
    def _distinct_bills(rows: Iterable[ExtractedRow]) -> List[ExtractedRow]:
        result: List[ExtractedRow] = []
        seen = set()
        for row in rows:
            if row.md5 not in seen:
                seen.add(row.md5)
                result.append(row)
        return result

    def _bill_choice(
        self,
        subject_key: str,
        subject: ExtractedRow,
        candidates: Sequence[ExtractedRow],
        target_candidates: Sequence[_TargetInfoRow] = (),
        allow_new_target: bool = True,
    ) -> BillChoiceRequest:
        return BillChoiceRequest(
            subject_key=subject_key,
            container=subject.container,
            close_date=subject.close_date,
            tons=subject.tons,
            cargo=subject.cargo,
            candidates=[
                BillCandidate(
                    staged_id=row.staged_id,
                    md5=row.md5,
                    source_name=row.source_name,
                    bill_no=row.bill_no,
                    container=row.container,
                    seal=row.seal,
                    vessel=row.vessel,
                    sail_date=row.sail_date,
                    carrier=row.carrier,
                )
                for row in candidates
            ],
            target_candidates=[
                TargetCandidate(
                    row_number=row.row_number,
                    sqt=row.sqt,
                    close_date=row.close_date,
                    container=row.container,
                    cargo=row.cargo,
                    vessel=str(row.values.get("Tên tàu") or ""),
                )
                for row in target_candidates
            ],
            allow_new_target=allow_new_target,
        )

    @staticmethod
    def _source_needs_review(*rows: Optional[ExtractedRow]) -> bool:
        for row in rows:
            if row is None:
                continue
            if row.warning:
                return True
            if _key_text(row.source_status) not in ("", "OK"):
                return True
            if _key_text(row.confidence) == "THAP":
                return True
        return False

    def _info_values(
        self,
        scale: ExtractedRow,
        bill: Optional[ExtractedRow],
        sqt: int,
        fallback: Dict[str, Any],
        check_status: str,
    ) -> Dict[str, Any]:
        vessel = (
            (bill.vessel if bill else "")
            or str(fallback.get("Tên tàu") or "")
            or (scale.vessel or "")
        )
        sail_date = (
            (bill.sail_date if bill else None)
            or parse_date(fallback.get("Ngày chạy"))
            or scale.sail_date
        )
        carrier = (
            (bill.carrier if bill else "")
            or str(fallback.get("VT biển") or "")
            or (scale.carrier or "")
        )
        seal = scale.seal or (bill.seal if bill else "") or fallback.get("Số chì/Seal")
        return {
            "SQT PM": sqt,
            "Ngày Đóng": excel_date(scale.close_date),
            "Số Container": scale.container,
            "Biển số xe": scale.truck_no or fallback.get("Biển số xe"),
            "Số chì/Seal": seal,
            "Số tấn": scale.tons,
            "Loại hàng": scale.cargo,
            "Nơi đóng": scale.place or fallback.get("Nơi đóng"),
            "Tên tàu": vessel,
            "Ngày chạy": excel_date(sail_date),
            "Người nhận": scale.recipient or fallback.get("Người nhận"),
            "VT biển": carrier,
            "MD5": _join_md5(
                _md5_cell_value(fallback),
                self._real_md5(scale),
                self._real_md5(bill) if bill else "",
            ),
            "Trạng thái nhập": _normalize_import_status(
                fallback.get("Trạng thái nhập"), "Chưa nhập"
            ),
            "Ngày cập nhật": datetime.now(),
        }

    def _bill_info_values(
        self, bill: ExtractedRow, sqt: int, fallback: Dict[str, Any]
    ) -> Dict[str, Any]:
        return {
            "SQT PM": sqt,
            "Ngày Đóng": excel_date(bill.close_date) or fallback.get("Ngày Đóng"),
            "Số Container": bill.container or fallback.get("Số Container"),
            "Biển số xe": bill.truck_no or fallback.get("Biển số xe"),
            "Số chì/Seal": bill.seal or fallback.get("Số chì/Seal"),
            "Số tấn": bill.tons if bill.tons is not None else fallback.get("Số tấn"),
            "Loại hàng": bill.cargo or fallback.get("Loại hàng"),
            "Nơi đóng": bill.place or fallback.get("Nơi đóng"),
            "Tên tàu": bill.vessel or fallback.get("Tên tàu"),
            "Ngày chạy": excel_date(bill.sail_date) or fallback.get("Ngày chạy"),
            "Người nhận": bill.recipient or fallback.get("Người nhận"),
            "VT biển": bill.carrier or fallback.get("VT biển"),
            "MD5": _join_md5(_md5_cell_value(fallback), self._real_md5(bill)),
            "Trạng thái nhập": _normalize_import_status(
                fallback.get("Trạng thái nhập"), "Chưa nhập"
            ),
            "Ngày cập nhật": datetime.now(),
        }

    @staticmethod
    def _find_conflicts(
        target: Optional[_TargetInfoRow], values: Dict[str, Any]
    ) -> List[FieldConflict]:
        if target is None:
            return []
        ignored = {
            "MD5",
            "Trạng thái nhập",
            "Ngày cập nhật",
        }
        conflicts = []
        for name, new_value in values.items():
            current = target.values.get(name)
            if name in ignored or current in (None, "") or new_value in (None, ""):
                continue
            current_compare = parse_date(current) or _key_text(current)
            new_compare = parse_date(new_value) or _key_text(new_value)
            if current_compare != new_compare:
                raw_id = f"{target.row_number}|{name}|{current}|{new_value}"
                conflicts.append(
                    FieldConflict(
                        conflict_id=hashlib.sha1(raw_id.encode("utf-8")).hexdigest(),
                        target_row=target.row_number,
                        field_name=name,
                        current_value=current,
                        new_value=new_value,
                    )
                )
        return conflicts

    def _analyze_expenses(
        self,
        expenses: Sequence[ExtractedRow],
        virtual_rows: Sequence[_TargetInfoRow],
        snapshot: _WorkbookSnapshot,
        analysis: ImportAnalysis,
    ) -> None:
        for expense in expenses:
            matches = [
                row
                for row in virtual_rows
                if row.close_date == expense.close_date
                and row.container == expense.container
            ]
            manual_match = False
            if expense.preferred_sqt:
                preferred = [row for row in matches if row.sqt == expense.preferred_sqt]
                if not preferred:
                    preferred = [
                        row
                        for row in virtual_rows
                        if row.sqt == expense.preferred_sqt
                        and row.container == expense.container
                    ]
                    manual_match = bool(preferred)
                if preferred:
                    matches = preferred
            sqt_values = {row.sqt for row in matches if row.sqt}
            sqt = next(iter(sqt_values)) if len(sqt_values) == 1 else None
            check_status = (
                "Cần kiểm tra" if self._source_needs_review(expense) else "OK"
            )
            if manual_match:
                check_status = "Cần kiểm tra"
            if (
                expense.total is not None
                and expense.amount is not None
                and expense.vat is not None
                and abs(float(expense.total) - float(expense.amount) - float(expense.vat)) > 1
            ):
                check_status = "Cần kiểm tra"
            existing = next(
                (
                    (row_number, values)
                    for row_number, values in snapshot.expense_rows
                    if parse_date(values.get("Ngày tháng")) == expense.close_date
                    and normalize_container(values.get("Số Container")) == expense.container
                ),
                None,
            )
            values = {
                "SQT PM": sqt,
                "Ngày tháng": excel_date(expense.close_date),
                "Số Container": expense.container,
                "Số HĐ": expense.invoice_no,
                "Giá vật liệu": expense.material_price,
                "Đơn giá": expense.unit_price,
                "Thành tiền": expense.amount,
                "VAT": expense.vat,
                "Tổng tiền": expense.total,
                "MD5": _join_md5(
                    _md5_cell_value(existing[1]) if existing else "",
                    self._real_md5(expense),
                ),
                "Trạng thái nhập": _normalize_import_status(
                    existing[1].get("Trạng thái nhập") if existing else None,
                    "Chưa nhập",
                ),
                "Ngày cập nhật": datetime.now(),
            }
            analysis.expense_changes.append(
                ExpenseChange(
                    action="UPDATE" if existing else "CREATE",
                    target_row=existing[0] if existing else None,
                    values=values,
                    staged_ids=[expense.staged_id],
                    document_md5s=[self._real_md5(expense)],
                    check_status=check_status,
                )
            )
            analysis.pending_states[expense.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.append(expense.staged_id)

    # ------------------------------------------------------------------ #
    # Workbook helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _header_map(ws) -> Dict[str, int]:
        mapping = {
            str(ws.cell(1, col).value): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(1, col).value not in (None, "")
        }
        if "MD5" not in mapping and "Mã MD5" in mapping:
            mapping["MD5"] = mapping["Mã MD5"]
        if "Mã MD5" not in mapping and "MD5" in mapping:
            mapping["Mã MD5"] = mapping["MD5"]
        return mapping

    def _validate_daily_headers(self, ws, expected: Sequence[str]) -> Dict[str, int]:
        mapping = self._header_map(ws)
        auto_addable = {"Biển số xe", "Ngày cập nhật", "MD5"}
        missing = [
            name
            for name in expected
            if name not in mapping and name not in auto_addable
        ]
        if missing:
            raise DailyImportError(
                f"Sheet '{ws.title}' thiếu các cột: " + ", ".join(missing)
            )
        return mapping

    def _upgrade_workbook(self, wb) -> None:
        for sheet_name, expected in (
            (INFO_SHEET, INFO_HEADERS),
            (EXPENSE_SHEET, EXPENSE_HEADERS),
        ):
            if sheet_name not in wb.sheetnames:
                raise DailyImportError(f"File theo dõi thiếu sheet '{sheet_name}'.")
            ws = wb[sheet_name]
            self._validate_daily_headers(ws, expected)
            if sheet_name == INFO_SHEET:
                self._ensure_column(ws, "Biển số xe", before_header="Số chì/Seal")
            self._ensure_column(ws, "Trạng thái nhập")
            self._ensure_column(ws, "Ngày cập nhật", before_header="MD5")
            if "MD5" not in self._header_map(ws):
                self._ensure_column(ws, "MD5")
            mapping = self._header_map(ws)
            missing_after = [name for name in expected if name not in mapping]
            if missing_after:
                raise DailyImportError(
                    f"Không nâng cấp được sheet '{sheet_name}': "
                    + ", ".join(missing_after)
                )

    def _ensure_column(
        self,
        ws,
        header: str,
        *,
        before_header: Optional[str] = None,
    ) -> None:
        mapping = self._header_map(ws)
        if header in mapping:
            return
        insert_col = ws.max_column + 1
        if before_header and before_header in mapping:
            insert_col = mapping[before_header]
            ws.insert_cols(insert_col, 1)
        ws.cell(1, insert_col).value = header
        style_source_col = min(insert_col + 1, ws.max_column)
        if style_source_col != insert_col:
            ws.cell(1, insert_col)._style = copy(ws.cell(1, style_source_col)._style)
        for row_number in range(2, ws.max_row + 1):
            if style_source_col != insert_col:
                ws.cell(row_number, insert_col)._style = copy(
                    ws.cell(row_number, style_source_col)._style
                )

    @staticmethod
    def _copy_row_style(ws, source_row: int, target_row: int) -> None:
        if source_row < 2 or source_row > ws.max_row:
            return
        for col in range(1, ws.max_column + 1):
            source = ws.cell(source_row, col)
            target = ws.cell(target_row, col)
            if source.has_style:
                target._style = copy(source._style)
            target.number_format = source.number_format
            target.alignment = copy(source.alignment)

    def _format_workbook(self, wb) -> None:
        for sheet_name in (INFO_SHEET, EXPENSE_SHEET):
            ws = wb[sheet_name]
            mapping = self._header_map(ws)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1, ws.max_row)}"
            # Form mới chỉ còn hai trạng thái nhập phục vụ PAD.
            ws.data_validations.dataValidation = []
            import_col = get_column_letter(mapping["Trạng thái nhập"])
            import_validation = DataValidation(
                type="list", formula1='"Chưa nhập,Đã nhập"', allow_blank=False
            )
            ws.add_data_validation(import_validation)
            import_validation.add(f"{import_col}2:{import_col}10000")
            status_col = mapping["Trạng thái nhập"]
            for row_number in range(2, ws.max_row + 1):
                if any(
                    ws.cell(row_number, col).value not in (None, "")
                    for col in range(1, ws.max_column + 1)
                    if col != status_col
                ):
                    ws.cell(row_number, status_col).value = _normalize_import_status(
                        ws.cell(row_number, status_col).value,
                        "Chưa nhập",
                    )
            for header in ("Ngày Đóng", "Ngày chạy", "Ngày Giao", "Ngày tháng"):
                if header in mapping:
                    for row_number in range(2, ws.max_row + 1):
                        ws.cell(row_number, mapping[header]).number_format = "dd/mm/yyyy"
            if "Ngày cập nhật" in mapping:
                for row_number in range(2, ws.max_row + 1):
                    ws.cell(row_number, mapping["Ngày cập nhật"]).number_format = (
                        "dd/mm/yyyy hh:mm:ss"
                    )
            for header in ("Giá vật liệu", "Đơn giá", "Thành tiền", "VAT", "Tổng tiền"):
                if header in mapping:
                    for row_number in range(2, ws.max_row + 1):
                        ws.cell(row_number, mapping[header]).number_format = "#,##0"
            if "Số tấn" in mapping:
                for row_number in range(2, ws.max_row + 1):
                    ws.cell(row_number, mapping["Số tấn"]).number_format = "0.00"
            # Cập nhật bảng Excel nếu workbook cũ có Table.
            for table in ws.tables.values():
                table.ref = f"A1:{get_column_letter(ws.max_column)}{max(2, ws.max_row)}"

    def _validate_saved_workbook(self, path: str) -> None:
        wb = load_workbook(path, read_only=True, data_only=False)
        try:
            for sheet_name, expected in (
                (INFO_SHEET, INFO_HEADERS),
                (EXPENSE_SHEET, EXPENSE_HEADERS),
            ):
                if sheet_name not in wb.sheetnames:
                    raise DailyImportError(f"File tạm thiếu sheet '{sheet_name}'.")
                mapping = self._header_map(wb[sheet_name])
                missing = [name for name in expected if name not in mapping]
                if missing:
                    raise DailyImportError(
                        f"File tạm thiếu cột ở sheet '{sheet_name}': "
                        + ", ".join(missing)
                    )
        finally:
            wb.close()

    @staticmethod
    def _summary_note(summary: ImportSummary) -> str:
        return (
            f"Đã nhập file theo dõi: {summary.new_info} dòng quyết toán mới, "
            f"{summary.updated_info} dòng cập nhật, {summary.new_expenses} khoản chi mới; "
            f"còn {summary.pending_count} dữ liệu chờ."
        )
