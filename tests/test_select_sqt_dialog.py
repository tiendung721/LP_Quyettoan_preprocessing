from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication, QDialog, QMessageBox  # noqa: E402
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.dialogs.select_sqt_dialog import SelectSqtDialog  # noqa: E402
from app.services.sqt_selection_service import SqtSelectionItem  # noqa: E402


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


class SelectSqtDialogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        self.items = [
            SqtSelectionItem("596", 2),
            SqtSelectionItem("597", 1),
            SqtSelectionItem("598", 5),
        ]

    def test_select_one_or_many_preserves_original_order(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.checkboxes[2].setChecked(True)
        dialog.checkboxes[0].setChecked(True)

        self.assertEqual(dialog.selected_values(), ["596", "598"])

    def test_select_all_and_clear_all(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.select_all()
        self.assertEqual(dialog.selected_values(), ["596", "597", "598"])

        dialog.clear_all()
        self.assertEqual(dialog.selected_values(), [])

    def test_run_without_selection_warns_and_keeps_dialog_open(self) -> None:
        dialog = SelectSqtDialog(self.items)

        with patch.object(QMessageBox, "warning") as warning:
            dialog._on_run()

        warning.assert_called_once()
        self.assertEqual(dialog.result(), QDialog.Rejected)

    def test_run_with_selection_accepts(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.checkboxes[1].setChecked(True)

        dialog._on_run()

        self.assertEqual(dialog.result(), QDialog.Accepted)
        self.assertEqual(dialog.selected_values(), ["597"])


if __name__ == "__main__":
    unittest.main()
