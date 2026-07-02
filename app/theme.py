"""Bộ giao diện (QSS) tông sáng, hiện đại cho Trợ Lý Quyết Toán RPA.

Tách phần trang trí (màu sắc, bo góc, khoảng cách) khỏi logic giao diện để dễ
chỉnh sửa. Chỉ cần sửa ``STYLESHEET`` bên dưới là đổi được diện mạo toàn app.
"""

from __future__ import annotations

from PySide6.QtWidgets import QApplication, QWidget

# Bảng màu chủ đạo (dùng lại ở nơi cần đặt màu bằng code).
COLORS = {
    "bg": "#F4F6F9",
    "card": "#FFFFFF",
    "border": "#E5E7EB",
    "text": "#1F2937",
    "muted": "#6B7280",
    "primary": "#2563EB",
    "success": "#16A34A",
    "danger": "#DC2626",
}

STYLESHEET = """
* {
    font-family: "Segoe UI", "Segoe UI Variable", Arial, sans-serif;
    font-size: 10.5pt;
    color: #1F2937;
}

QWidget#appRoot { background: #F4F6F9; }
QWidget#page { background: #F4F6F9; }
QStackedWidget#content { background: #F4F6F9; }

/* ---------- Thanh điều hướng bên trái ---------- */
QWidget#sidebar { background: #FFFFFF; border-right: 1px solid #E5E7EB; }
QLabel#appTitle { font-size: 15pt; font-weight: 700; color: #111827; padding: 20px 16px 2px 16px; }
QLabel#appSubtitle { font-size: 9pt; color: #6B7280; padding: 0 16px 14px 16px; }
QListWidget#navList { background: transparent; border: none; outline: 0; }
QListWidget#navList::item { padding: 11px 14px; margin: 3px 10px; border-radius: 8px; color: #374151; }
QListWidget#navList::item:hover { background: #F3F4F6; }
QListWidget#navList::item:selected { background: #EFF6FF; color: #1D4ED8; font-weight: 600; }
QLabel#navFooter { color: #9CA3AF; font-size: 8.5pt; padding: 10px 16px 14px 16px; }

/* ---------- Thẻ (card) ---------- */
QFrame#card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; }
QLabel#cardTitle { font-size: 12pt; font-weight: 700; color: #111827; }
QLabel#cardHint { color: #6B7280; }

/* ---------- Nhãn trợ giúp ---------- */
QLabel#statusText { color: #1D4ED8; font-weight: 600; }
QLabel#fileName { font-size: 12.5pt; font-weight: 700; color: #111827; }
QLabel#metaText { color: #6B7280; }
QLabel#noteText { color: #B91C1C; }
QLabel#formLabel { color: #374151; }

/* ---------- Nút bấm ---------- */
QPushButton {
    background: #FFFFFF; color: #374151; border: 1px solid #D1D5DB;
    border-radius: 8px; padding: 8px 14px;
}
QPushButton:hover { background: #F3F4F6; }
QPushButton:pressed { background: #E5E7EB; }
QPushButton:disabled { color: #9CA3AF; background: #F3F4F6; border-color: #E5E7EB; }

QPushButton[variant="primary"] { background: #2563EB; color: #FFFFFF; border: none; }
QPushButton[variant="primary"]:hover { background: #1D4ED8; }
QPushButton[variant="primary"]:pressed { background: #1E40AF; }
QPushButton[variant="primary"]:disabled { background: #93C5FD; color: #EFF6FF; }

QPushButton[variant="success"] { background: #16A34A; color: #FFFFFF; border: none; }
QPushButton[variant="success"]:hover { background: #15803D; }
QPushButton[variant="success"]:pressed { background: #166534; }
QPushButton[variant="success"]:disabled { background: #86EFAC; color: #F0FDF4; }

QPushButton#bigButton { min-height: 30px; font-size: 13pt; font-weight: 700; padding: 12px 18px; }

/* ---------- Ô nhập ---------- */
QLineEdit { background: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 8px; padding: 7px 10px; }
QLineEdit:focus { border: 1px solid #2563EB; }
QCheckBox { spacing: 8px; color: #374151; }

/* ---------- Thanh tiến trình các bước ---------- */
QLabel#stepCircle { border-radius: 16px; font-weight: 700; font-size: 11pt; }
QLabel#stepCircle[state="pending"] { background: #E5E7EB; color: #9CA3AF; }
QLabel#stepCircle[state="active"] { background: #2563EB; color: #FFFFFF; }
QLabel#stepCircle[state="done"] { background: #16A34A; color: #FFFFFF; }
QLabel#stepLabel { color: #9CA3AF; font-size: 9.5pt; }
QLabel#stepLabel[state="active"] { color: #1D4ED8; font-weight: 600; }
QLabel#stepLabel[state="done"] { color: #16A34A; }
QFrame#stepLine { background: #E5E7EB; border: none; }

/* ---------- Bảng ---------- */
QTableWidget {
    background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px;
    gridline-color: #F1F3F5; alternate-background-color: #FAFBFC;
}
QTableWidget::item { padding: 4px 6px; }
QTableWidget::item:selected { background: #EFF6FF; color: #1F2937; }
QHeaderView::section {
    background: #F9FAFB; color: #374151; padding: 8px; border: none;
    border-bottom: 1px solid #E5E7EB; font-weight: 600;
}
QTableCornerButton::section { background: #F9FAFB; border: none; }

/* ---------- Nhật ký ---------- */
QPlainTextEdit {
    background: #F9FAFB; color: #374151; border: 1px solid #E5E7EB;
    border-radius: 8px; font-family: "Consolas", "Courier New", monospace; font-size: 9.5pt;
}

/* ---------- Vùng cuộn ---------- */
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 12px; margin: 2px; }
QScrollBar::handle:vertical { background: #CBD1DA; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #AEB6C2; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
"""


def apply_theme(app: QApplication) -> None:
    """Áp bộ style tông sáng cho toàn ứng dụng."""
    app.setStyleSheet(STYLESHEET)


def repolish(widget: QWidget) -> None:
    """Ép Qt vẽ lại widget sau khi đổi thuộc tính động (dynamic property)."""
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()
