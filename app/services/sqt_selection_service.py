"""Read SQT choices from the daily workbook and write PAD selection JSON."""

from __future__ import annotations

import json
import math
import os
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, List

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .. import file_utils
from ..daily_import import INFO_SHEET

SQT_COLUMN = "SQT PM"
SELECTION_OPERATION = "NHAP_THONG_TIN"


class SqtSelectionError(RuntimeError):
    """Business error safe to show directly to the user."""


class DailyFileNotFoundError(SqtSelectionError):
    pass


class DailyFileLockedError(SqtSelectionError):
    pass


class MissingSheetError(SqtSelectionError):
    pass


class MissingColumnError(SqtSelectionError):
    pass


class EmptySqtListError(SqtSelectionError):
    pass


class SelectionJsonWriteError(SqtSelectionError):
    pass


@dataclass(frozen=True)
class SqtSelectionItem:
    value: str
    row_count: int

    @property
    def display_text(self) -> str:
        return f"{self.value} — {self.row_count} dòng"


def normalize_sqt_value(value: Any) -> str:
    """Return the stable string sent to PAD for an Excel SQT cell value."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).strip()
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, "g").strip()

    text = str(value).strip()
    if not text:
        return ""
    if text.endswith(".0"):
        head = text[:-2]
        if head and head.lstrip("+-").isdigit():
            return str(int(head))
    return text


def _header_map(ws) -> dict[str, int]:
    return {
        str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value not in (None, "")
    }


def read_sqt_items(daily_file: str, sheet_name: str = INFO_SHEET) -> List[SqtSelectionItem]:
    """Read unique SQT values in first-seen order and count source rows."""
    if not daily_file or not os.path.isfile(daily_file):
        raise DailyFileNotFoundError("Không tìm thấy file Excel hằng ngày.")

    try:
        if file_utils.is_file_locked(daily_file):
            raise DailyFileLockedError(
                "File Excel hằng ngày đang bị khóa. Vui lòng đóng Excel và thử lại."
            )
    except FileNotFoundError as exc:
        raise DailyFileNotFoundError("Không tìm thấy file Excel hằng ngày.") from exc

    wb = None
    try:
        wb = load_workbook(daily_file, read_only=True, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise MissingSheetError(f"Không tìm thấy sheet “{sheet_name}”.")

        ws = wb[sheet_name]
        headers = _header_map(ws)
        if SQT_COLUMN not in headers:
            raise MissingColumnError(
                f"Không tìm thấy cột “{SQT_COLUMN}” trong sheet “{sheet_name}”."
            )

        sqt_col = headers[SQT_COLUMN]
        counts: "OrderedDict[str, int]" = OrderedDict()
        for row in ws.iter_rows(
            min_row=2,
            min_col=sqt_col,
            max_col=sqt_col,
            values_only=True,
        ):
            sqt = normalize_sqt_value(row[0] if row else None)
            if not sqt:
                continue
            counts[sqt] = counts.get(sqt, 0) + 1

        if not counts:
            raise EmptySqtListError("Sheet không có SQT hợp lệ.")
        return [SqtSelectionItem(value=sqt, row_count=count) for sqt, count in counts.items()]
    except SqtSelectionError:
        raise
    except PermissionError as exc:
        raise DailyFileLockedError(
            "File Excel hằng ngày đang bị khóa. Vui lòng đóng Excel và thử lại."
        ) from exc
    except (InvalidFileException, OSError, ValueError) as exc:
        raise SqtSelectionError(
            f"Không thể đọc file {os.path.basename(daily_file)}. "
            "Vui lòng đóng Excel và thử lại."
        ) from exc
    finally:
        if wb is not None:
            wb.close()


def write_selection_json(
    target_path: str | os.PathLike[str],
    daily_file: str,
    selected_sqt: Iterable[str],
    sheet_name: str = INFO_SHEET,
) -> str:
    """Write the selected SQT list atomically for PAD to consume."""
    selected = [str(value) for value in selected_sqt if str(value).strip()]
    if not selected:
        raise SelectionJsonWriteError("Vui lòng chọn ít nhất một Số quyết toán.")

    target = Path(target_path)
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise SelectionJsonWriteError("Không tạo được thư mục runtime.") from exc

    payload = {
        "operation": SELECTION_OPERATION,
        "daily_file": str(Path(daily_file).resolve()),
        "sheet_name": sheet_name,
        "selected_sqt": selected,
    }

    temp_path = target.with_name(target.name + ".tmp")
    try:
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(temp_path, target)
    except OSError as exc:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass
        raise SelectionJsonWriteError("Không ghi được JSON lựa chọn RPA.") from exc

    return str(target.resolve())

