"""Quản lý cấu hình ứng dụng (settings.json).

- Nếu file settings.json chưa tồn tại thì tự tạo với giá trị mặc định.
- Tự tạo các thư mục cần thiết nếu còn thiếu.
- Cung cấp các đường dẫn dẫn xuất (logs, database, ...).
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict

# Thư mục gốc mặc định của toàn hệ thống.
DEFAULT_APP_ROOT = r"D:\RPA_QuyetToan"

# Vị trí mặc định của file cấu hình.
DEFAULT_CONFIG_PATH = os.path.join(DEFAULT_APP_ROOT, "Config", "settings.json")

# Các thư mục con chuẩn theo cấu trúc dự án.
SUBFOLDERS = [
    "App",
    "Config",
    "Launcher",
    "Downloads",
    "Outputs",
    "Daily",
    "Database",
    "Logs",
]


def get_default_settings(app_root: str = DEFAULT_APP_ROOT) -> Dict[str, Any]:
    """Trả về dict cấu hình mặc định."""
    return {
        "app_root": app_root,
        "bat_path": os.path.join(app_root, "Launcher", "Mo_Tro_Ly_Quyet_Toan.bat"),
        "download_folder": os.path.join(app_root, "Downloads"),
        "output_folder": os.path.join(app_root, "Outputs"),
        "daily_tracking_file": os.path.join(
            app_root, "Daily", "file_theo_doi_hang_ngay.xlsx"
        ),
        "allowed_extensions": [".xlsx", ".xlsm", ".csv"],
        "output_file_patterns": [
            "input_quyet_toan*.xlsx",
            "output*.xlsx",
            "input_trip*.xlsx",
            "*.xlsx",
        ],
        "download_stable_seconds": 3,
        "auto_open_after_download": False,
    }


class AppConfig:
    """Bọc dữ liệu cấu hình và các thao tác liên quan."""

    def __init__(self, data: Dict[str, Any], path: str):
        self.data = data
        self.path = path

    # ------------------------------------------------------------------ #
    # Nạp / lưu
    # ------------------------------------------------------------------ #
    @classmethod
    def load(cls, config_path: str = DEFAULT_CONFIG_PATH) -> "AppConfig":
        """Nạp cấu hình từ file; nếu chưa có thì tạo mặc định."""
        defaults = get_default_settings()

        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, OSError):
                # File hỏng -> dùng mặc định nhưng không xóa file cũ.
                data = {}
            # Bổ sung các khóa còn thiếu bằng giá trị mặc định.
            merged = {**defaults, **(data or {})}
            config = cls(merged, config_path)
        else:
            config = cls(defaults, config_path)
            config.save()

        return config

    def save(self) -> None:
        """Ghi cấu hình xuống file settings.json."""
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    # ------------------------------------------------------------------ #
    # Truy cập tiện lợi
    # ------------------------------------------------------------------ #
    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value

    @property
    def app_root(self) -> str:
        return self.data.get("app_root", DEFAULT_APP_ROOT)

    @property
    def bat_path(self) -> str:
        return self.data.get("bat_path", "")

    @property
    def download_folder(self) -> str:
        return self.data.get("download_folder", "")

    @property
    def output_folder(self) -> str:
        return self.data.get("output_folder", "")

    @property
    def daily_tracking_file(self) -> str:
        return self.data.get("daily_tracking_file", "")

    @property
    def allowed_extensions(self):
        return self.data.get("allowed_extensions", [])

    @property
    def output_file_patterns(self):
        return self.data.get("output_file_patterns", [])

    @property
    def download_stable_seconds(self) -> int:
        try:
            return int(self.data.get("download_stable_seconds", 3))
        except (TypeError, ValueError):
            return 3

    @property
    def auto_open_after_download(self) -> bool:
        return bool(self.data.get("auto_open_after_download", False))

    # Đường dẫn dẫn xuất theo app_root.
    @property
    def logs_dir(self) -> str:
        return os.path.join(self.app_root, "Logs")

    @property
    def config_dir(self) -> str:
        return os.path.join(self.app_root, "Config")

    @property
    def database_path(self) -> str:
        return os.path.join(self.app_root, "Database", "app_state.db")

    # ------------------------------------------------------------------ #
    # Tạo thư mục
    # ------------------------------------------------------------------ #
    def ensure_folders(self) -> None:
        """Tạo toàn bộ thư mục cần thiết nếu còn thiếu."""
        # Thư mục gốc + các thư mục con chuẩn.
        os.makedirs(self.app_root, exist_ok=True)
        for sub in SUBFOLDERS:
            os.makedirs(os.path.join(self.app_root, sub), exist_ok=True)

        # Các đường dẫn tùy biến (có thể nằm ngoài app_root).
        for folder in (
            self.download_folder,
            self.output_folder,
            self.logs_dir,
        ):
            if folder:
                os.makedirs(folder, exist_ok=True)

        # Thư mục chứa database và file theo dõi hàng ngày.
        os.makedirs(os.path.dirname(self.database_path), exist_ok=True)
        if self.daily_tracking_file:
            os.makedirs(os.path.dirname(self.daily_tracking_file), exist_ok=True)
