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
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import file_utils
from .database import Database


OUTPUT_SHEET = "Du_Lieu_Boc_Tach"
INFO_SHEET = "Thong_Tin_Quyet_Toan"
EXPENSE_SHEET = "Khoan_Chi"

OUTPUT_HEADERS = [
    "STT hiển thị",
    "File nguồn",
    "Mã MD5 file",
    "Loại chứng từ",
    "Trạng thái",
    "Ngày Đóng",
    "Số Container",
    "Số tấn",
    "Loại hàng",
    "Nơi đóng",
    "Tên tàu",
    "Ngày chạy",
    "VT biển",
    "Giá vật liệu",
    "Số HĐ",
    "Số Bill",
    "Số chì/Seal",
    "Đơn giá",
    "Thành tiền",
    "VAT",
    "Tổng tiền",
    "Trường khác",
    "Độ tin cậy",
    "Cảnh báo",
]

INFO_HEADERS = [
    "SQT PM",
    "Ngày Đóng",
    "Số Container",
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
    "Mã MD5",
    "Độ tin cậy",
    "Trạng thái kiểm tra",
    "Trạng thái nhập",
    "Ngày cập nhật",
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
    "Mã MD5",
    "Trạng thái kiểm tra",
    "Trạng thái nhập",
    "Ngày cập nhật",
]

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
class BillChoiceRequest:
    subject_key: str
    container: str
    close_date: Optional[str]
    tons: Optional[float]
    cargo: str
    candidates: List[BillCandidate]


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


def _confidence_min(*values: str) -> str:
    rank = {"THAP": 0, "TRUNG BINH": 1, "CAO": 2}
    present = [value for value in values if value]
    if not present:
        return ""
    return min(present, key=lambda item: rank.get(_key_text(item), 0))


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
        if file_utils.is_file_locked(output_path):
            raise DailyImportError("File Output đang mở. Vui lòng đóng file rồi thử lại.")
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
            if statuses.get(row.md5) in ("COMPLETED", "IGNORED"):
                continue
            if row.md5 in snapshot.target_md5 and row.md5 not in statuses:
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
        virtual_rows = list(snapshot.info_rows)

        for scale in scales:
            if scale.md5.startswith("missing-"):
                analysis.pending_states[scale.staged_id] = (
                    STATE_MISSING,
                    "Chứng từ không có Mã MD5 nên chưa thể nhập an toàn.",
                )
                continue
            target_group = (
                []
                if scale.force_new_sqt
                else self._target_group(virtual_rows, scale.close_date, scale.container)
            )
            if scale.preferred_sqt:
                preferred = [
                    item for item in target_group if item.sqt == scale.preferred_sqt
                ]
                if not preferred:
                    preferred = [
                        item
                        for item in virtual_rows
                        if item.sqt == scale.preferred_sqt
                        and item.container == scale.container
                    ]
                if preferred:
                    target_group = preferred
            target, split_sqt, group_conflict = self._choose_target(target_group, scale)
            if group_conflict:
                analysis.pending_states[scale.staged_id] = (
                    STATE_CONFLICT,
                    "Có nhiều SQT khác nhau cùng Ngày Đóng và Container.",
                )
                continue

            candidates = self._distinct_bills(
                row
                for row in bills_by_container.get(scale.container, [])
                if not row.md5.startswith("missing-")
            )
            subject_key = f"scale:{scale.row_key}"
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
                        self._bill_choice(subject_key, scale, candidates)
                    )
                    analysis.pending_states[scale.staged_id] = (
                        STATE_NEEDS_BILL,
                        "Container xuất hiện trên nhiều Bill.",
                    )
                    for bill in candidates:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_NEEDS_BILL,
                            "Chờ người dùng chọn Bill phù hợp.",
                        )
                    continue

            fallback = target.values if target else (
                target_group[0].values if target_group else {}
            )
            # Ưu tiên Bill -> dòng đã có trong file theo dõi -> thông tin người
            # dùng nhập tay ngay trên dòng phiếu cân (cho phép hoàn tất khi
            # không có Bill).
            vessel = (
                (selected_bill.vessel if selected_bill else "")
                or str(fallback.get("Tên tàu") or "")
                or (scale.vessel or "")
            )
            sail_date = (
                (selected_bill.sail_date if selected_bill else None)
                or parse_date(fallback.get("Ngày chạy"))
                or scale.sail_date
            )
            carrier = (
                (selected_bill.carrier if selected_bill else "")
                or str(fallback.get("VT biển") or "")
                or (scale.carrier or "")
            )
            missing = self._missing_info_fields(scale, vessel, sail_date, carrier)
            if missing:
                state = STATE_WAITING_BILL if not selected_bill else STATE_MISSING
                analysis.pending_states[scale.staged_id] = (
                    state,
                    "Còn thiếu: " + ", ".join(missing),
                )
                continue

            sqt = target.sqt if target else (split_sqt or next_sqt)
            action = "UPDATE" if target else "CREATE"
            if not target and not split_sqt:
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
                values["Trạng thái kiểm tra"] = check_status
                analysis.conflicts.extend(conflicts)
            staged_ids = [scale.staged_id]
            md5_values = [scale.md5]
            if selected_bill:
                staged_ids.append(selected_bill.staged_id)
                md5_values.append(selected_bill.md5)
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
            remaining = [
                row
                for row in raw_candidates
                if row.staged_id not in used_bills and not row.md5.startswith("missing-")
            ]
            for row in raw_candidates:
                if row.md5.startswith("missing-"):
                    analysis.pending_states[row.staged_id] = (
                        STATE_MISSING,
                        "Chứng từ không có Mã MD5 nên chưa thể nhập an toàn.",
                    )
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
            sqt_values = {row.sqt for row in target_rows if row.sqt}
            candidates = self._distinct_bills(remaining)
            if not target_rows:
                for bill in remaining:
                    analysis.pending_states[bill.staged_id] = (
                        STATE_WAITING_SCALE,
                        "Chưa tìm thấy Phiếu cân hoặc dòng quyết toán cùng Container.",
                    )
                continue
            if len(sqt_values) != 1:
                for bill in remaining:
                    analysis.pending_states[bill.staged_id] = (
                        STATE_CONFLICT,
                        "Container đang thuộc nhiều SQT khác nhau.",
                    )
                continue
            subject_key = f"target:{container}:{next(iter(sqt_values))}"
            target_dates = {row.close_date for row in target_rows if row.close_date}
            container_decision = None
            if len(target_dates) == 1:
                container_decision = decisions.get(
                    f"container:{container}:{next(iter(target_dates))}"
                )
            if container_decision and all(
                item.md5 != container_decision for item in candidates
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
                    seed = ExtractedRow(
                        container=container,
                        close_date=target_rows[0].close_date,
                        tons=parse_number(target_rows[0].values.get("Số tấn")),
                        cargo=str(target_rows[0].values.get("Loại hàng") or ""),
                    )
                    analysis.bill_choices.append(
                        self._bill_choice(subject_key, seed, candidates)
                    )
                    for bill in remaining:
                        analysis.pending_states[bill.staged_id] = (
                            STATE_NEEDS_BILL,
                            "Chờ người dùng chọn Bill phù hợp.",
                        )
                    continue
            if selected_bill is None:
                continue
            for target in target_rows:
                values = dict(target.values)
                values.update(
                    {
                        "Số chì/Seal": selected_bill.seal or values.get("Số chì/Seal"),
                        "Tên tàu": selected_bill.vessel or values.get("Tên tàu"),
                        "Ngày chạy": excel_date(selected_bill.sail_date)
                        or values.get("Ngày chạy"),
                        "VT biển": selected_bill.carrier or values.get("VT biển"),
                        "Mã MD5": _join_md5(values.get("Mã MD5"), selected_bill.md5),
                        "Độ tin cậy": _confidence_min(
                            str(values.get("Độ tin cậy") or ""), selected_bill.confidence
                        ),
                        "Ngày cập nhật": datetime.now(),
                    }
                )
                complete = all(
                    values.get(name) not in (None, "")
                    for name in (
                        "Ngày Đóng",
                        "Số Container",
                        "Số tấn",
                        "Loại hàng",
                        "Tên tàu",
                        "Ngày chạy",
                        "VT biển",
                    )
                )
                values["Trạng thái kiểm tra"] = (
                    "OK"
                    if complete and not self._source_needs_review(selected_bill)
                    else "Cần kiểm tra"
                )
                conflicts = self._find_conflicts(target, values)
                analysis.conflicts.extend(conflicts)
                analysis.info_changes.append(
                    InfoChange(
                        action="UPDATE",
                        target_row=target.row_number,
                        sqt=target.sqt,
                        values=values,
                        staged_ids=[selected_bill.staged_id],
                        document_md5s=[selected_bill.md5],
                        check_status=values["Trạng thái kiểm tra"],
                        conflicts=conflicts,
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
                        "Mã MD5",
                        "Độ tin cậy",
                        "Trạng thái kiểm tra",
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
                        "Mã MD5",
                        "Trạng thái kiểm tra",
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
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            raise DailyImportError(f"Không đọc được file Output: {exc}") from exc
        try:
            if OUTPUT_SHEET not in wb.sheetnames:
                raise DailyImportError(
                    f"File Output thiếu sheet '{OUTPUT_SHEET}'."
                )
            ws = wb[OUTPUT_SHEET]
            headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
            missing = [header for header in OUTPUT_HEADERS if header not in headers]
            if missing:
                raise DailyImportError(
                    "File Output thiếu các cột: " + ", ".join(missing)
                )
            index = {name: headers.index(name) + 1 for name in OUTPUT_HEADERS}
            result: List[ExtractedRow] = []
            for row_number in range(2, ws.max_row + 1):
                values = [ws.cell(row_number, col).value for col in range(1, ws.max_column + 1)]
                if not any(value not in (None, "") for value in values):
                    continue
                doc_type = _doc_type(ws.cell(row_number, index["Loại chứng từ"]).value)
                if not doc_type:
                    continue
                cargo, recognized = normalize_cargo(
                    ws.cell(row_number, index["Loại hàng"]).value,
                    self.cargo_mappings,
                )
                close_date = parse_date(ws.cell(row_number, index["Ngày Đóng"]).value)
                sail_date = parse_date(ws.cell(row_number, index["Ngày chạy"]).value)
                row = ExtractedRow(
                    source_row=row_number,
                    source_name=str(ws.cell(row_number, index["File nguồn"]).value or ""),
                    md5=normalize_md5(ws.cell(row_number, index["Mã MD5 file"]).value),
                    doc_type=doc_type,
                    source_status=str(ws.cell(row_number, index["Trạng thái"]).value or ""),
                    close_date=close_date,
                    container=normalize_container(
                        ws.cell(row_number, index["Số Container"]).value
                    ),
                    tons=parse_number(ws.cell(row_number, index["Số tấn"]).value),
                    cargo=cargo,
                    cargo_recognized=recognized,
                    place=str(ws.cell(row_number, index["Nơi đóng"]).value or "").strip(),
                    vessel=str(ws.cell(row_number, index["Tên tàu"]).value or "").strip(),
                    sail_date=sail_date,
                    carrier=str(ws.cell(row_number, index["VT biển"]).value or "").strip(),
                    material_price=parse_number(
                        ws.cell(row_number, index["Giá vật liệu"]).value
                    ),
                    invoice_no=str(ws.cell(row_number, index["Số HĐ"]).value or "").strip(),
                    bill_no=str(ws.cell(row_number, index["Số Bill"]).value or "").strip(),
                    seal=str(ws.cell(row_number, index["Số chì/Seal"]).value or "").strip(),
                    unit_price=parse_number(ws.cell(row_number, index["Đơn giá"]).value),
                    amount=parse_number(ws.cell(row_number, index["Thành tiền"]).value),
                    vat=parse_number(ws.cell(row_number, index["VAT"]).value),
                    total=parse_number(ws.cell(row_number, index["Tổng tiền"]).value),
                    other=str(ws.cell(row_number, index["Trường khác"]).value or "").strip(),
                    confidence=str(
                        ws.cell(row_number, index["Độ tin cậy"]).value or ""
                    ).strip(),
                    warning=str(ws.cell(row_number, index["Cảnh báo"]).value or "").strip(),
                )
                if doc_type == DOC_EXPENSE:
                    row.close_date = sail_date or close_date
                if not row.md5:
                    # Không có MD5 thì vẫn theo dõi được trong danh sách chờ, nhưng
                    # dùng khóa cục bộ để không làm mất dòng.
                    row.md5 = "missing-" + hashlib.sha1(
                        f"{path}|{row_number}|{row.container}".encode("utf-8")
                    ).hexdigest()
                    row.warning = (row.warning + "; Thiếu MD5 nguồn").strip("; ")
                result.append(row)
            return result
        finally:
            wb.close()

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
                    for part in re.split(r"[;,\n]+", str(values.get("Mã MD5") or ""))
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
                    for part in re.split(r"[;,\n]+", str(values.get("Mã MD5") or ""))
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
    def _choose_target(
        group: Sequence[_TargetInfoRow], scale: ExtractedRow
    ) -> Tuple[Optional[_TargetInfoRow], Optional[int], bool]:
        if not group:
            return None, None, False
        sqt_values = {row.sqt for row in group if row.sqt}
        if len(sqt_values) != 1:
            return None, None, True
        exact = [row for row in group if _key_text(row.cargo) == _key_text(scale.cargo)]
        if len(exact) > 1:
            return None, None, True
        if exact:
            return exact[0], None, False
        return None, next(iter(sqt_values)), False

    @staticmethod
    def _distinct_bills(rows: Iterable[ExtractedRow]) -> List[ExtractedRow]:
        result: List[ExtractedRow] = []
        seen = set()
        for row in rows:
            if row.md5 not in seen:
                seen.add(row.md5)
                result.append(row)
        return result

    @staticmethod
    def _bill_choice(
        subject_key: str,
        subject: ExtractedRow,
        candidates: Sequence[ExtractedRow],
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
        )

    @staticmethod
    def _missing_info_fields(
        scale: ExtractedRow,
        vessel: str,
        sail_date: Optional[str],
        carrier: str,
    ) -> List[str]:
        checks = [
            ("Ngày Đóng", scale.close_date),
            ("Số Container", scale.container if is_valid_container(scale.container) else None),
            ("Số tấn", scale.tons if scale.tons and scale.tons > 0 else None),
            ("Loại hàng", scale.cargo),
            ("Tên tàu", vessel),
            ("Ngày chạy", sail_date),
            ("VT biển", carrier),
        ]
        return [name for name, value in checks if value in (None, "")]

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
            "Số chì/Seal": seal,
            "Số tấn": scale.tons,
            "Loại hàng": scale.cargo,
            "Nơi đóng": scale.place or fallback.get("Nơi đóng"),
            "Tên tàu": vessel,
            "Ngày chạy": excel_date(sail_date),
            "VT biển": carrier,
            "Mã MD5": _join_md5(
                fallback.get("Mã MD5"), scale.md5, bill.md5 if bill else ""
            ),
            "Độ tin cậy": _confidence_min(
                str(fallback.get("Độ tin cậy") or ""),
                scale.confidence,
                bill.confidence if bill else "",
            ),
            "Trạng thái kiểm tra": check_status,
            "Trạng thái nhập": fallback.get("Trạng thái nhập") or "Chưa nhập",
            "Ngày cập nhật": datetime.now(),
        }

    @staticmethod
    def _find_conflicts(
        target: Optional[_TargetInfoRow], values: Dict[str, Any]
    ) -> List[FieldConflict]:
        if target is None:
            return []
        ignored = {
            "Mã MD5",
            "Độ tin cậy",
            "Trạng thái kiểm tra",
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
            if expense.md5.startswith("missing-"):
                analysis.pending_states[expense.staged_id] = (
                    STATE_MISSING,
                    "Chứng từ không có Mã MD5 nên chưa thể nhập an toàn.",
                )
                continue
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
            if len(sqt_values) != 1:
                note = (
                    "Không tìm thấy SQT theo Ngày tháng và Container."
                    if not sqt_values
                    else "Tìm thấy nhiều SQT khác nhau."
                )
                analysis.pending_states[expense.staged_id] = (STATE_CONFLICT, note)
                continue
            sqt = next(iter(sqt_values))
            required = {
                "Ngày tháng": expense.close_date,
                "Số Container": (
                    expense.container if is_valid_container(expense.container) else None
                ),
                "Số HĐ": expense.invoice_no,
                "Đơn giá": expense.unit_price,
                "Thành tiền": expense.amount,
                "VAT": expense.vat,
                "Tổng tiền": expense.total,
            }
            missing = [name for name, value in required.items() if value in (None, "")]
            if missing:
                analysis.pending_states[expense.staged_id] = (
                    STATE_MISSING,
                    "Còn thiếu: " + ", ".join(missing),
                )
                continue
            check_status = (
                "Cần kiểm tra" if self._source_needs_review(expense) else "OK"
            )
            if manual_match:
                check_status = "Cần kiểm tra"
            if abs(float(expense.total) - float(expense.amount) - float(expense.vat)) > 1:
                check_status = "Cần kiểm tra"
            existing = next(
                (
                    (row_number, values)
                    for row_number, values in snapshot.expense_rows
                    if parse_date(values.get("Ngày tháng")) == expense.close_date
                    and normalize_container(values.get("Số Container")) == expense.container
                    and _key_text(values.get("Số HĐ")) == _key_text(expense.invoice_no)
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
                "Mã MD5": _join_md5(
                    existing[1].get("Mã MD5") if existing else "", expense.md5
                ),
                "Trạng thái kiểm tra": check_status,
                "Trạng thái nhập": (
                    existing[1].get("Trạng thái nhập") if existing else "Chưa nhập"
                )
                or "Chưa nhập",
                "Ngày cập nhật": datetime.now(),
            }
            analysis.expense_changes.append(
                ExpenseChange(
                    action="UPDATE" if existing else "CREATE",
                    target_row=existing[0] if existing else None,
                    values=values,
                    staged_ids=[expense.staged_id],
                    document_md5s=[expense.md5],
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
        return {
            str(ws.cell(1, col).value): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(1, col).value not in (None, "")
        }

    def _validate_daily_headers(self, ws, expected: Sequence[str]) -> Dict[str, int]:
        mapping = self._header_map(ws)
        # Hai cột trạng thái kiểm tra có thể chưa tồn tại; commit sẽ tự nâng cấp.
        missing = [name for name in expected if name not in mapping and name != "Trạng thái kiểm tra"]
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
            mapping = self._validate_daily_headers(ws, expected)
            if "Trạng thái kiểm tra" not in mapping:
                import_col = mapping["Trạng thái nhập"]
                ws.insert_cols(import_col, 1)
                ws.cell(1, import_col).value = "Trạng thái kiểm tra"
                if import_col + 1 <= ws.max_column:
                    ws.cell(1, import_col)._style = copy(ws.cell(1, import_col + 1)._style)
                for row_number in range(2, ws.max_row + 1):
                    if any(
                        ws.cell(row_number, col).value not in (None, "")
                        for col in range(1, ws.max_column + 1)
                        if col != import_col
                    ):
                        ws.cell(row_number, import_col).value = "OK"
                        if import_col + 1 <= ws.max_column:
                            ws.cell(row_number, import_col)._style = copy(
                                ws.cell(row_number, import_col + 1)._style
                            )
            mapping = self._header_map(ws)
            missing_after = [name for name in expected if name not in mapping]
            if missing_after:
                raise DailyImportError(
                    f"Không nâng cấp được sheet '{sheet_name}': "
                    + ", ".join(missing_after)
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
        green = PatternFill("solid", fgColor="E2F0D9")
        yellow = PatternFill("solid", fgColor="FFF2CC")
        for sheet_name in (INFO_SHEET, EXPENSE_SHEET):
            ws = wb[sheet_name]
            mapping = self._header_map(ws)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1, ws.max_row)}"
            # Xóa validation cũ của hai cột trạng thái rồi dựng lại phạm vi rộng.
            ws.data_validations.dataValidation = []
            check_col = get_column_letter(mapping["Trạng thái kiểm tra"])
            import_col = get_column_letter(mapping["Trạng thái nhập"])
            check_validation = DataValidation(
                type="list", formula1='"OK,Cần kiểm tra"', allow_blank=False
            )
            import_validation = DataValidation(
                type="list", formula1='"Chưa nhập,Đã nhập"', allow_blank=False
            )
            ws.add_data_validation(check_validation)
            ws.add_data_validation(import_validation)
            check_validation.add(f"{check_col}2:{check_col}10000")
            import_validation.add(f"{import_col}2:{import_col}10000")
            ws.conditional_formatting.add(
                f"{check_col}2:{check_col}10000",
                FormulaRule(formula=[f'${check_col}2="OK"'], fill=green),
            )
            ws.conditional_formatting.add(
                f"{check_col}2:{check_col}10000",
                FormulaRule(formula=[f'${check_col}2="Cần kiểm tra"'], fill=yellow),
            )
            for header in ("Ngày Đóng", "Ngày chạy", "Ngày Giao", "Ngày tháng", "Ngày cập nhật"):
                if header in mapping:
                    for row_number in range(2, ws.max_row + 1):
                        ws.cell(row_number, mapping[header]).number_format = "dd/mm/yyyy"
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
