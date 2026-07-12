from __future__ import annotations

import json
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
    DailyImportError,
    DailyImportService,
    normalize_cargo,
)
from app.database import Database


LEGACY_INFO_HEADERS = list(INFO_HEADERS)
LEGACY_EXPENSE_HEADERS = list(EXPENSE_HEADERS)


class DailyImportServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = Database(str(self.root / "state.db"))
        self.db.init_db()
        self.service = DailyImportService(self.db, logging.getLogger("daily-test"))
        self.daily = self.root / "daily.xlsx"
        self.output = self.root / "output.json"
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
        self._write_output_rows(
            {
                "file_nguon": "scale.jpg",
                "ma_md5_file": "scale-md5",
                "loai_chung_tu": "Phiếu cân",
                "trang_thai": "OK",
                "ngay_dong": "05/06/2026",
                "so_container": "IJKL 1234567",
                "so_tan": 29.5,
                "loai_hang": "VOI ROI",
                "noi_dong": "Kho C",
                "nguoi_nhan": "LÊ PHẠM",
                "truong_khac": {"so_phieu": "0000733"},
                "do_tin_cay": "Cao",
                "canh_bao": [],
            },
            {
                "file_nguon": "bill.pdf",
                "ma_md5_file": "bill-md5",
                "loai_chung_tu": "Bill",
                "trang_thai": "OK",
                "so_container": "IJKL1234567",
                "ten_tau": "TÀU C 03S",
                "ngay_chay": "06/06/2026",
                "vt_bien": "Hãng C",
                "so_bill": "BILL-03",
                "so_chi_seal": "SEAL03",
                "do_tin_cay": "Cao",
            },
            {
                "file_nguon": "cost.jpg",
                "ma_md5_file": "cost-md5",
                "loai_chung_tu": "Khoản chi",
                "trang_thai": "OK",
                "ngay_chay": "03/06/2026",
                "so_container": "EFGH1234567",
                "so_hd": "51",
                "don_gia": 1_700_000,
                "thanh_tien": 45_900_000,
                "vat": 3_672_000,
                "tong_tien": 49_572_000,
                "do_tin_cay": "Cao",
                "canh_bao": [{"noi_dung": "Kiểm tra VAT"}],
            },
        )

    def _write_output_rows(self, *rows: dict) -> None:
        payload = {
            "metadata": {
                "phien_ban_schema": "1.0",
                "tong_so_dong_boc_tach": len(rows),
            },
            "du_lieu_boc_tach": list(rows),
            "canh_bao": [],
            "raw_data": [],
        }
        with open(self.output, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

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
        self.assertIn("MD5", [cell.value for cell in info[1]])
        self.assertIn("MD5", [cell.value for cell in expense[1]])
        self.assertEqual(info.max_row, 4)
        self.assertEqual(expense.max_row, 2)
        headers = {cell.value: cell.column for cell in info[1]}
        self.assertEqual(info.cell(2, headers["Trạng thái nhập"]).value, "Đã nhập")
        self.assertEqual(info.cell(4, headers["SQT PM"]).value, 102)
        self.assertEqual(info.cell(4, headers["Loại hàng"]).value, "Vôi rời")
        self.assertEqual(info.cell(4, headers["Người nhận"]).value, "LÊ PHẠM")
        self.assertEqual(info.cell(4, headers["Trạng thái nhập"]).value, "Chưa nhập")
        wb.close()

        again = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(again.new_info_count, 0)
        self.assertEqual(len(again.expense_changes), 0)
        self.assertEqual(again.duplicate_documents, 3)

    def test_bill_without_scale_is_persisted(self) -> None:
        self._write_output_rows(
            {
                "file_nguon": "waiting.pdf",
                "ma_md5_file": "waiting-bill",
                "loai_chung_tu": "Bill",
                "so_container": "MNOP1234567",
                "ten_tau": "TÀU CHỜ 01S",
                "ngay_chay": "10/06/2026",
                "vt_bien": "Hãng chờ",
            }
        )

        analysis = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(analysis.new_info_count, 1)
        self.assertEqual(analysis.info_changes[0].values["Số Container"], "MNOP1234567")
        self.assertEqual(analysis.info_changes[0].values["Tên tàu"], "TÀU CHỜ 01S")

        summary = self.service.commit(analysis)
        self.assertEqual(summary.new_info, 1)
        self.assertEqual(summary.pending_count, 0)

    def test_invalid_json_schema_is_reported(self) -> None:
        with open(self.output, "w", encoding="utf-8") as f:
            json.dump({"metadata": {}}, f)

        with self.assertRaises(DailyImportError) as ctx:
            self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertIn("du_lieu_boc_tach", str(ctx.exception))

    def test_multiple_bills_requires_and_remembers_user_choice(self) -> None:
        rows = [
            {
                "file_nguon": "scale.jpg",
                "ma_md5_file": "scale-choice",
                "loai_chung_tu": "Phiếu cân",
                "trang_thai": "OK",
                "ngay_dong": "08/06/2026",
                "so_container": "QRST1234567",
                "so_tan": 28,
                "loai_hang": "VOI",
            }
        ]
        for md5, bill_no, vessel in (
            ("bill-choice-a", "BILL-A", "TÀU A 01S"),
            ("bill-choice-b", "BILL-B", "TÀU B 02S"),
        ):
            rows.append(
                {
                    "file_nguon": bill_no + ".pdf",
                    "ma_md5_file": md5,
                    "loai_chung_tu": "Bill",
                    "trang_thai": "OK",
                    "so_container": "QRST1234567",
                    "so_bill": bill_no,
                    "ten_tau": vessel,
                    "ngay_chay": "09/06/2026",
                    "vt_bien": "Hãng tàu",
                }
            )
        self._write_output_rows(*rows)

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

    def test_expense_matches_by_date_and_container_without_invoice(self) -> None:
        self._write_output_rows(
            {
                "file_nguon": "cost-1.jpg",
                "ma_md5_file": "cost-one",
                "loai_chung_tu": "Khoản chi",
                "ngay_chay": "03/06/2026",
                "so_container": "EFGH1234567",
                "so_hd": "OLD",
                "don_gia": 100,
            }
        )
        first = self.service.analyze(str(self.output), str(self.daily), 1)
        self.service.commit(first)

        self._write_output_rows(
            {
                "file_nguon": "cost-2.jpg",
                "ma_md5_file": "cost-two",
                "loai_chung_tu": "Khoản chi",
                "ngay_chay": "03/06/2026",
                "so_container": "EFGH1234567",
                "so_hd": "NEW",
                "vat": 10,
            }
        )
        second = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(len(second.expense_changes), 1)
        self.assertEqual(second.expense_changes[0].action, "UPDATE")

    def test_missing_md5_still_imports(self) -> None:
        self._write_output_rows(
            {
                "file_nguon": "scale-no-md5.jpg",
                "loai_chung_tu": "Phiếu cân",
                "ngay_dong": "11/06/2026",
                "so_container": "ZZZZ1234567",
                "bien_so_xe": "15H 22404",
                "so_tan": 25,
                "loai_hang": "VOI",
            }
        )

        analysis = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(analysis.new_info_count, 1)
        self.assertEqual(analysis.info_changes[0].values["Biển số xe"], "15H 22404")
        self.assertEqual(analysis.info_changes[0].values["MD5"], "")

    def test_multiple_sqt_requires_target_choice(self) -> None:
        wb = load_workbook(self.daily)
        ws = wb[INFO_SHEET]
        values = {name: None for name in LEGACY_INFO_HEADERS}
        values.update(
            {
                "SQT PM": 200,
                "Ngày Đóng": datetime(2026, 6, 12),
                "Số Container": "MULT1234567",
                "Số tấn": 20,
            }
        )
        ws.append([values[name] for name in LEGACY_INFO_HEADERS])
        values["SQT PM"] = 201
        values["Số tấn"] = 21
        ws.append([values[name] for name in LEGACY_INFO_HEADERS])
        wb.save(self.daily)
        wb.close()

        self._write_output_rows(
            {
                "file_nguon": "scale-multi.jpg",
                "ma_md5_file": "scale-multi",
                "loai_chung_tu": "Phiếu cân",
                "ngay_dong": "12/06/2026",
                "so_container": "MULT1234567",
                "so_tan": 22,
            }
        )
        first = self.service.analyze(str(self.output), str(self.daily), 1)
        self.assertEqual(len(first.bill_choices), 1)
        request = first.bill_choices[0]
        self.assertEqual(len(request.target_candidates), 2)

        second = self.service.analyze(
            str(self.output),
            str(self.daily),
            1,
            bill_decisions={request.target_subject_key: "sqt:201"},
        )
        self.assertEqual(len(second.bill_choices), 0)
        self.assertEqual(second.info_changes[0].sqt, 201)


if __name__ == "__main__":
    unittest.main()
