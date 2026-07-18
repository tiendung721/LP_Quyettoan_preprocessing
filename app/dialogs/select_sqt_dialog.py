"""Dialog for selecting one or more SQT values."""

from __future__ import annotations

from typing import Iterable, List

from PySide6.QtCore import Qt, Slot
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from ..services.sqt_selection_service import SqtSelectionItem


class SelectSqtDialog(QDialog):
    """Single modal popup for one-or-many SQT selection."""

    def __init__(self, items: Iterable[SqtSelectionItem], parent=None):
        super().__init__(parent)
        self.items = list(items)
        self.checkboxes: List[QCheckBox] = []
        self._selected_values: List[str] = []
        self._selected_items: List[SqtSelectionItem] = []

        self.setWindowTitle("Chọn Số quyết toán cần nhập")
        self.setModal(True)
        self.resize(520, 560)
        self.setMinimumSize(420, 360)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        title = QLabel("CHỌN SỐ QUYẾT TOÁN CẦN NHẬP")
        title.setObjectName("cardTitle")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        host = QWidget()
        list_layout = QVBoxLayout(host)
        list_layout.setContentsMargins(4, 4, 8, 4)
        list_layout.setSpacing(8)
        for item in self.items:
            checkbox = QCheckBox(item.display_text)
            checkbox.setProperty("sqtValue", item.value)
            checkbox.setToolTip(item.value)
            self.checkboxes.append(checkbox)
            list_layout.addWidget(checkbox)
        list_layout.addStretch(1)
        scroll.setWidget(host)
        layout.addWidget(scroll, stretch=1)

        bulk = QHBoxLayout()
        btn_select_all = QPushButton("Chọn tất cả")
        btn_clear_all = QPushButton("Bỏ chọn tất cả")
        btn_select_all.clicked.connect(self.select_all)
        btn_clear_all.clicked.connect(self.clear_all)
        bulk.addWidget(btn_select_all)
        bulk.addWidget(btn_clear_all)
        bulk.addStretch(1)
        layout.addLayout(bulk)

        actions = QHBoxLayout()
        btn_cancel = QPushButton("Hủy")
        self.btn_run = QPushButton("Chạy nhập thông tin")
        self.btn_run.setProperty("variant", "primary")
        btn_cancel.clicked.connect(self.reject)
        self.btn_run.clicked.connect(self._on_run)
        actions.addWidget(btn_cancel)
        actions.addStretch(1)
        actions.addWidget(self.btn_run)
        layout.addLayout(actions)

    @Slot()
    def select_all(self) -> None:
        for checkbox in self.checkboxes:
            checkbox.setChecked(True)

    @Slot()
    def clear_all(self) -> None:
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)

    def selected_values(self) -> List[str]:
        if self._selected_values:
            return list(self._selected_values)
        return [
            item.value
            for item, checkbox in zip(self.items, self.checkboxes)
            if checkbox.isChecked()
        ]

    def selected_items(self) -> List[SqtSelectionItem]:
        if self._selected_items:
            return list(self._selected_items)
        return [
            item
            for item, checkbox in zip(self.items, self.checkboxes)
            if checkbox.isChecked()
        ]

    @Slot()
    def _on_run(self) -> None:
        selected_items = self.selected_items()
        if not selected_items:
            QMessageBox.warning(
                self,
                "Chưa chọn SQT",
                "Vui lòng chọn ít nhất một Số quyết toán.",
            )
            return
        self._selected_items = selected_items
        self._selected_values = [item.value for item in selected_items]
        self.accept()
