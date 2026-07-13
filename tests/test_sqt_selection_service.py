from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from app.daily_import import INFO_SHEET
from app.services.sqt_selection_service import (
    EmptySqtListError,
    MissingColumnError,
    MissingSheetError,
    normalize_sqt_value,
    read_sqt_items,
    write_selection_json,
)


class SqtSelectionServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.daily = self.root / "Quyết toán hằng ngày.xlsx"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _make_daily(self, values, *, sheet_name: str = INFO_SHEET, header: str = "SQT PM") -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append([header, "Số Container"])
        for index, value in enumerate(values, start=1):
            ws.append([value, f"CONT{index:02d}"])
        wb.save(self.daily)
        wb.close()

    def test_normalize_sqt_value(self) -> None:
        self.assertEqual(normalize_sqt_value(596), "596")
        self.assertEqual(normalize_sqt_value(596.0), "596")
        self.assertEqual(normalize_sqt_value("596"), "596")
        self.assertEqual(normalize_sqt_value("596.0"), "596")
        self.assertEqual(normalize_sqt_value(""), "")
        self.assertEqual(normalize_sqt_value(None), "")

    def test_read_sqt_items_keeps_first_seen_order_and_counts_rows(self) -> None:
        self._make_daily([596, 596.0, "597", "", None, "596", 598.0])

        items = read_sqt_items(str(self.daily))

        self.assertEqual(
            [(item.value, item.row_count, item.display_text) for item in items],
            [
                ("596", 3, "596 — 3 dòng"),
                ("597", 1, "597 — 1 dòng"),
                ("598", 1, "598 — 1 dòng"),
            ],
        )

    def test_missing_sheet_is_reported(self) -> None:
        self._make_daily([596], sheet_name="SheetKhac")

        with self.assertRaises(MissingSheetError):
            read_sqt_items(str(self.daily))

    def test_missing_sqt_column_is_reported(self) -> None:
        self._make_daily([596], header="SQT Khac")

        with self.assertRaises(MissingColumnError):
            read_sqt_items(str(self.daily))

    def test_empty_sqt_list_is_reported(self) -> None:
        self._make_daily(["", None])

        with self.assertRaises(EmptySqtListError):
            read_sqt_items(str(self.daily))

    def test_write_selection_json_schema_and_overwrite(self) -> None:
        target = self.root / "runtime" / "rpa_input_selection.json"

        first = write_selection_json(target, str(self.daily), ["596"])
        second = write_selection_json(target, str(self.daily), ["597", "598"])

        self.assertEqual(first, second)
        with open(target, encoding="utf-8") as f:
            payload = json.load(f)

        self.assertEqual(payload["operation"], "NHAP_THONG_TIN")
        self.assertEqual(Path(payload["daily_file"]), self.daily.resolve())
        self.assertEqual(payload["sheet_name"], INFO_SHEET)
        self.assertEqual(payload["selected_sqt"], ["597", "598"])
        self.assertNotIn("RPA_Queue", json.dumps(payload, ensure_ascii=False))
        self.assertNotIn("CREATE_NEW", json.dumps(payload, ensure_ascii=False))
        self.assertFalse(target.with_name(target.name + ".tmp").exists())


if __name__ == "__main__":
    unittest.main()

