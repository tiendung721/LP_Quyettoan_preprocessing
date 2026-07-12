from __future__ import annotations

import logging
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.daily_import import (
    EXPENSE_HEADERS,
    EXPENSE_SHEET,
    INFO_HEADERS,
    INFO_SHEET,
    OUTPUT_HEADERS,
    OUTPUT_SHEET,
    DailyImportService,
    normalize_cargo,
)
from app.database import Database


LEGACY_INFO_HEADERS = [h for h in INFO_HEADERS if h != "Trạng thái kiểm tra"]
LEGACY_EXPENSE_HEADERS = [h for h in EXPENSE_HEADERS if h != "Trạng thái kiểm tra"]


class DailyImportServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = Database(str(self.root / "state.db"))
        self.db.init_db()
        self.service = DailyImportService(self.db, logging.getLogger("daily-test"))
        self.daily = self.root / "daily.xlsx"
        self.output = self.root / "output.xlsx"
        self._make_daily()

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _make_daily(self) -> None:
        wb = Workbook()
        info = wb.active
        info.title = INFO_SHEET
        info.append(LEGACY_INFO_HEADERS)
        values = {name: None for name in LEGACY_INFO_HEADERS}
        values.update(
            {
                "SQT PM": 100,
                "Ngày Đóng": datetime(2026, 6, 1),
                "Số Container": "ABCD1234567",
                "Số tấn": 28.0,
                "Loại hàng": "Vôi rời",
                "Nơi đóng": "Kho A",
                "Tên tàu": "TÀU CŨ 01S",
                "Ngày chạy": datetime(2026, 6, 2),
                "VT biển": "Hãng A",
                "Trạng thái nhập": "Đã nhập",
            }
        )
        info.append([values[name] for name in LEGACY_INFO_HEADERS])
        values.update(
            {
                "SQT PM": 101,
                "Ngày Đóng": datetime(2026, 6, 3),
                "Số Container": "EFGH1234567",
                "Số tấn": 27.0,
                "Loại hàng": "Vôi",
                "Tên tàu": "TÀU B 02S",
                "Ngày chạy": datetime(2026, 6, 4),
                "VT biển": "Hãng B",
            }
        )
        info.append([values[name] for name in LEGACY_INFO_HEADERS])
        expense = wb.create_sheet(EXPENSE_SHEET)
        expense.append(LEGACY_EXPENSE_HEADERS)
        wb.save(self.daily)
        wb.close()

    def _make_output(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = OUTPUT_SHEET
        ws.append(OUTPUT_HEADERS)

        def add(**kwargs):
            row = {name: None for name in OUTPUT_HEADERS}
            row.update(kwargs)
            ws.append([row[name] for name in OUTPUT_HEADERS])

        add(
            **{
                "File nguồn": "scale.jpg",
                "Mã MD5 file": "scale-md5",
                "Loại chứng từ": "Phiếu cân",
                "Trạng thái": "OK",
                "Ngày Đóng": "05/06/2026",
                "Số Container": "IJKL 1234567",
                "Số tấn": 29.5,
                "Loại hàng": "VOI ROI",
                "Nơi đóng": "Kho C",
                "Độ tin cậy": "Cao",
            }
        )
        add(
            **{
                "File nguồn": "bill.pdf",
                "Mã MD5 file": "bill-md5",
                "Loại chứng từ": "Bill",
                "Trạng thái": "OK",
                "Số Container": "IJKL1234567",
                "Tên tàu": "TÀU C 03S",
                "Ngày chạy": "06/06/2026",
                "VT biển": "Hãng C",
                "Số Bill": "BILL-03",
                "Số chì/Seal": "SEAL03",
                "Độ tin cậy": "Cao",
            }
        )
        add(
            **{
                "File nguồn": "cost.jpg",
                "Mã MD5 file": "cost-md5",
                "Loại chứng từ": "Khoản chi",
                "Trạng thái": "OK",
                "Ngày chạy": "03/06/2026",
                "Số Container": "EFGH1234567",
                "Số HĐ": "51",
                "Đơn giá": 1_700_000,
                "Thành tiền": 45_900_000,
                "VAT": 3_672_000,
                "Tổng tiền": 49_572_000,
                "Độ tin cậy": "Cao",
            }
        )
        for _ in range(5):
            ws.append([None] * len(OUTPUT_HEADERS))
        wb.save(self.output)
        wb.close()

    def test_normalize_cargo_with_vietnamese_accents(self) -> None:
        self.assertEqual(normalize_cargo("VOI")[0], "Vôi")
        self.assertEqual(normalize_cargo("VOI ROI")[0], "Vôi rời")
        self.assertEqual(normalize_cargo("Vôi cục C2 xá rời")[0], "Vôi cục C2 xá rời")

    def test_end_to_end_upgrade_match_and_deduplicate(self) -> None:
        self._make_output()
        analysis = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(analysis.extracted_rows, 3)
        self.assertEqual(analysis.new_info_count, 1)
        self.assertEqual(len(analysis.expense_changes), 1)
        self.assertEqual(analysis.expense_changes[0].values["SQT PM"], 101)

        summary = self.service.commit(analysis)
        self.assertEqual(summary.new_info, 1)
        self.assertEqual(summary.new_expenses, 1)
        # Không còn sao lưu file theo dõi: không tạo thư mục Backups.
        self.assertEqual(summary.backup_path, "")
        self.assertFalse((Path(self.daily).parent / "Backups").exists())

        wb = load_workbook(self.daily, data_only=False)
        info = wb[INFO_SHEET]
        expense = wb[EXPENSE_SHEET]
        self.assertIn("Trạng thái kiểm tra", [cell.value for cell in info[1]])
        self.assertIn("Trạng thái kiểm tra", [cell.value for cell in expense[1]])
        self.assertEqual(info.max_row, 4)
        self.assertEqual(expense.max_row, 2)
        headers = {cell.value: cell.column for cell in info[1]}
        self.assertEqual(info.cell(2, headers["Trạng thái nhập"]).value, "Đã nhập")
        self.assertEqual(info.cell(4, headers["SQT PM"]).value, 102)
        self.assertEqual(info.cell(4, headers["Loại hàng"]).value, "Vôi rời")
        wb.close()

        again = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(again.new_info_count, 0)
        self.assertEqual(len(again.expense_changes), 0)
        self.assertEqual(again.duplicate_documents, 3)

    def test_bill_without_scale_is_persisted(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = OUTPUT_SHEET
        ws.append(OUTPUT_HEADERS)
        data = {name: None for name in OUTPUT_HEADERS}
        data.update(
            {
                "File nguồn": "waiting.pdf",
                "Mã MD5 file": "waiting-bill",
                "Loại chứng từ": "Bill",
                "Số Container": "MNOP1234567",
                "Tên tàu": "TÀU CHỜ 01S",
                "Ngày chạy": "10/06/2026",
                "VT biển": "Hãng chờ",
            }
        )
        ws.append([data[name] for name in OUTPUT_HEADERS])
        wb.save(self.output)
        wb.close()

        analysis = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(analysis.new_info_count, 0)
        pending = self.db.list_staged_rows()
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0]["state"], "WAITING_SCALE")

    def test_multiple_bills_requires_and_remembers_user_choice(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = OUTPUT_SHEET
        ws.append(OUTPUT_HEADERS)

        def add(**kwargs):
            data = {name: None for name in OUTPUT_HEADERS}
            data.update(kwargs)
            ws.append([data[name] for name in OUTPUT_HEADERS])

        add(
            **{
                "File nguồn": "scale.jpg",
                "Mã MD5 file": "scale-choice",
                "Loại chứng từ": "Phiếu cân",
                "Trạng thái": "OK",
                "Ngày Đóng": "08/06/2026",
                "Số Container": "QRST1234567",
                "Số tấn": 28,
                "Loại hàng": "VOI",
            }
        )
        for md5, bill_no, vessel in (
            ("bill-choice-a", "BILL-A", "TÀU A 01S"),
            ("bill-choice-b", "BILL-B", "TÀU B 02S"),
        ):
            add(
                **{
                    "File nguồn": bill_no + ".pdf",
                    "Mã MD5 file": md5,
                    "Loại chứng từ": "Bill",
                    "Trạng thái": "OK",
                    "Số Container": "QRST1234567",
                    "Số Bill": bill_no,
                    "Tên tàu": vessel,
                    "Ngày chạy": "09/06/2026",
                    "VT biển": "Hãng tàu",
                }
            )
        wb.save(self.output)
        wb.close()

        first = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(len(first.bill_choices), 1)
        request = first.bill_choices[0]
        selected = "bill-choice-b"
        self.db.save_match_decision(request.subject_key, request.container, selected)
        self.db.save_match_decision(
            f"container:{request.container}:{request.close_date}",
            request.container,
            selected,
        )
        second = self.service.analyze(
            str(self.output),
            str(self.daily),
            1,
            bill_decisions={request.subject_key: selected},
        )
        self.assertEqual(len(second.bill_choices), 0)
        self.assertEqual(second.new_info_count, 1)
        self.assertEqual(second.info_changes[0].values["Tên tàu"], "TÀU B 02S")
        self.service.commit(second)

        third = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(third.new_info_count, 0)
        self.assertEqual(third.updated_info_count, 0)
        remaining = [
            item
            for item in self.db.list_staged_rows()
            if item.get("document_md5") == "bill-choice-a"
        ]
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0]["state"], "WAITING_SCALE")


if __name__ == "__main__":
    unittest.main()
