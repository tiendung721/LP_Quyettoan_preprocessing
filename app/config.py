"""Quản lý cấu hình ứng dụng (settings.json).

- Nếu file settings.json chưa tồn tại thì tự tạo với giá trị mặc định.
- Tự tạo các thư mục cần thiết nếu còn thiếu.
- Cung cấp các đường dẫn dẫn xuất (logs, database, ...).
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict

# Thư mục gốc mặc định nằm ngay tại project đang chạy.
DEFAULT_APP_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

# Vị trí mặc định của file cấu hình.
DEFAULT_CONFIG_PATH = os.path.join(DEFAULT_APP_ROOT, "Config", "settings.json")

# Các thư mục con chuẩn theo cấu trúc dự án.
# Dữ liệu người dùng chỉ gom vào một thư mục duy nhất: output.
SUBFOLDERS = [
    "Config",
    "Launcher",
    "output",
    "Database",
    "Logs",
    "runtime",
]

OUTPUT_FOLDER_NAME = "output"
DAILY_TRACKING_FILENAME = "quyet_toan_hang_ngay.xlsx"


def default_output_folder(app_root: str) -> str:
    return os.path.join(app_root, OUTPUT_FOLDER_NAME)


def default_daily_tracking_file(app_root: str) -> str:
    return os.path.join(default_output_folder(app_root), DAILY_TRACKING_FILENAME)


def _is_legacy_split_folder(path: str, app_root: str) -> bool:
    if not path:
        return True
    try:
        resolved = os.path.normcase(os.path.abspath(path))
        root = os.path.normcase(os.path.abspath(app_root))
    except (OSError, ValueError):
        return True
    name = os.path.basename(resolved).lower()
    if name in {"download", "downloads", "outputs", "daily"}:
        return True
    return resolved == os.path.normcase(os.path.join(root, "download"))


def normalize_single_output_settings(data: Dict[str, Any]) -> Dict[str, Any]:
    """Ép các đường dẫn dữ liệu về một thư mục làm việc duy nhất."""
    app_root = data.get("app_root") or DEFAULT_APP_ROOT

    configured_output = data.get("output_folder") or ""
    configured_download = data.get("download_folder") or ""
    if configured_output and not _is_legacy_split_folder(configured_output, app_root):
        output_folder = configured_output
    elif configured_download and not _is_legacy_split_folder(configured_download, app_root):
        output_folder = configured_download
    else:
        output_folder = default_output_folder(app_root)
    output_folder = os.path.normpath(output_folder)

    daily_path = data.get("daily_tracking_file") or ""
    if daily_path:
        daily_name = os.path.basename(daily_path) or DAILY_TRACKING_FILENAME
    else:
        daily_name = DAILY_TRACKING_FILENAME
    daily_path = os.path.normpath(os.path.join(output_folder, daily_name))

    data["output_folder"] = output_folder
    data["daily_tracking_file"] = daily_path
    if "file_stable_seconds" not in data and "download_stable_seconds" in data:
        data["file_stable_seconds"] = data["download_stable_seconds"]
    data.pop("download_folder", None)
    data.pop("download_stable_seconds", None)
    return data

# Nội dung file .bat mẫu cho luồng PAD RPA. Người dùng chỉ cần mở file này và
# điền lệnh gọi flow của mình vào (viết không dấu để hiển thị đúng trong cmd).
PAD_BAT_TEMPLATE = """@echo off
REM ==========================================================================
REM  Chay flow PAD RPA de nhap du lieu moi len phan mem quyet toan.
REM
REM  >>> DIEN LENH GOI FLOW PAD CUA BAN VAO PHIA DUOI, ROI XOA PHAN CANH BAO. <<<
REM
REM  Vi du:
REM    "C:\\Program Files (x86)\\Power Automate\\PAD.Console.Host.exe" ^
REM        -flow "Nhap_Quyet_Toan" -run
REM ==========================================================================

echo [PAD RPA] File .bat nay chua duoc cau hinh.
echo Hay mo file sau va dien lenh chay flow PAD cua ban:
echo    %~f0
pause
exit /b 1
"""

PAD_NEW_SQT_BAT_TEMPLATE = """@echo off
REM Run PAD flow: create exactly one new settlement number.
setlocal enableextensions enabledelayedexpansion

set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"

if exist "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" (
  "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new
  exit /b !errorlevel!
)

where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new
  exit /b !errorlevel!
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new
  exit /b !errorlevel!
)

echo [ERROR] Python not found. Cannot launch PAD flow.
exit /b 2
"""

PAD_INPUT_INFORMATION_BAT_TEMPLATE = """@echo off
REM Run PAD flow: import information for selected SQT values.
setlocal enableextensions enabledelayedexpansion

set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"

if exist "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" (
  "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_information
  exit /b !errorlevel!
)

where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_information
  exit /b !errorlevel!
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_information
  exit /b !errorlevel!
)

echo [ERROR] Python not found. Cannot launch PAD flow.
exit /b 2
"""

PAD_INPUT_EXPENSE_BAT_TEMPLATE = """@echo off
REM Run PAD flow: import expense rows for selected SQT values.
setlocal enableextensions enabledelayedexpansion

set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"

if exist "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" (
  "%PROJECT_ROOT%\\.venv\\Scripts\\python.exe" "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%PROJECT_ROOT%\\scripts\\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

echo [ERROR] Python not found. Cannot launch PAD flow.
exit /b 2
"""


def get_default_settings(app_root: str = DEFAULT_APP_ROOT) -> Dict[str, Any]:
    """Trả về dict cấu hình mặc định."""
    output_folder = default_output_folder(app_root)
    return {
        "app_root": app_root,
        "bat_path": os.path.join(app_root, "Launcher", "Mo_Tro_Ly_Quyet_Toan.bat"),
        "pad_bat_path": os.path.join(app_root, "Launcher", "Chay_PAD_Quyet_Toan.bat"),
        "pad_create_new_bat_path": os.path.join(
            app_root, "Launcher", "run_create_new_quyet_toan.bat"
        ),
        "pad_input_information_bat_path": os.path.join(
            app_root, "Launcher", "run_input_information.bat"
        ),
        "pad_input_expense_bat_path": os.path.join(
            app_root, "Launcher", "run_input_expense.bat"
        ),
        "output_folder": output_folder,
        "daily_tracking_file": default_daily_tracking_file(app_root),
        "allowed_extensions": [".json"],
        "output_file_patterns": [
            "boc_tach*.json",
            "rpa_input*.json",
            "*.json",
        ],
        "file_stable_seconds": 3,
        "cargo_name_mappings": {
            "VOI": "Vôi",
            "VOI ROI": "Vôi rời",
            "VOI BOT": "Vôi bột",
            "VOI BOT NONG NGHIEP": "Vôi bột nông nghiệp",
            "BOT NN BAO 25 KG": "Bột NN bao 25 kg",
        },
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
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, OSError):
                # File hỏng -> dùng mặc định nhưng không xóa file cũ.
                data = {}
            defaults = get_default_settings((data or {}).get("app_root", DEFAULT_APP_ROOT))
            # Bổ sung các khóa còn thiếu bằng giá trị mặc định.
            merged = {**defaults, **(data or {})}
            if (
                isinstance(data, dict)
                and "download_stable_seconds" in data
                and "file_stable_seconds" not in data
            ):
                merged["file_stable_seconds"] = data["download_stable_seconds"]
            # Luồng mới chỉ nhận file bóc tách JSON, kể cả khi settings cũ còn
            # lưu cấu hình Excel từ phiên bản trước.
            merged["allowed_extensions"] = defaults["allowed_extensions"]
            merged["output_file_patterns"] = defaults["output_file_patterns"]
            before_normalize = dict(merged)
            normalize_single_output_settings(merged)
            config = cls(merged, config_path)
            if merged != before_normalize:
                config.save()
        else:
            config = cls(get_default_settings(), config_path)
            normalize_single_output_settings(config.data)
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
    def pad_bat_path(self) -> str:
        return self.data.get("pad_bat_path", "")

    @property
    def pad_create_new_bat_path(self) -> str:
        return self.data.get("pad_create_new_bat_path", "")

    @property
    def pad_input_information_bat_path(self) -> str:
        return self.data.get("pad_input_information_bat_path", "")

    @property
    def pad_input_expense_bat_path(self) -> str:
        return self.data.get("pad_input_expense_bat_path", "")

    @property
    def output_folder(self) -> str:
        return self.data.get("output_folder", default_output_folder(self.app_root))

    @property
    def daily_tracking_file(self) -> str:
        return self.data.get("daily_tracking_file", default_daily_tracking_file(self.app_root))

    @property
    def allowed_extensions(self):
        return self.data.get("allowed_extensions", [])

    @property
    def output_file_patterns(self):
        return self.data.get("output_file_patterns", [])

    @property
    def file_stable_seconds(self) -> int:
        try:
            return int(self.data.get("file_stable_seconds", 3))
        except (TypeError, ValueError):
            return 3

    @property
    def cargo_name_mappings(self) -> Dict[str, str]:
        value = self.data.get("cargo_name_mappings", {})
        return value if isinstance(value, dict) else {}

    # Đường dẫn dẫn xuất theo app_root.
    @property
    def logs_dir(self) -> str:
        return os.path.join(self.app_root, "Logs")

    @property
    def runtime_dir(self) -> str:
        return os.path.join(self.app_root, "runtime")

    @property
    def rpa_input_selection_path(self) -> str:
        return os.path.join(self.runtime_dir, "rpa_input_selection.json")

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
        normalize_single_output_settings(self.data)

        # Thư mục gốc + các thư mục con chuẩn.
        os.makedirs(self.app_root, exist_ok=True)
        for sub in SUBFOLDERS:
            os.makedirs(os.path.join(self.app_root, sub), exist_ok=True)

        # Các đường dẫn tùy biến (có thể nằm ngoài app_root).
        for folder in (
            self.output_folder,
            self.logs_dir,
            self.runtime_dir,
        ):
            if folder:
                os.makedirs(folder, exist_ok=True)

        # Thư mục chứa database và file theo dõi hàng ngày.
        os.makedirs(os.path.dirname(self.database_path), exist_ok=True)
        if self.daily_tracking_file:
            os.makedirs(os.path.dirname(self.daily_tracking_file), exist_ok=True)

        self.ensure_pad_bat_template()

    def set_output_folder(self, folder: str) -> None:
        """Đổi thư mục làm việc duy nhất và đồng bộ các khóa cấu hình liên quan."""
        folder = os.path.normpath(folder or default_output_folder(self.app_root))
        self.data["output_folder"] = folder
        self.data.pop("download_folder", None)
        self.data["daily_tracking_file"] = os.path.join(
            folder,
            os.path.basename(self.daily_tracking_file) or DAILY_TRACKING_FILENAME,
        )

    def ensure_pad_bat_template(self) -> bool:
        """Tạo sẵn file .bat mẫu cho luồng PAD RPA nếu chưa có.

        Chỉ tạo khi file nằm trong thư mục Launcher của app_root, để không bao
        giờ đè lên file .bat thật mà người dùng tự trỏ tới nơi khác.

        Trả về True nếu vừa tạo file mẫu.
        """
        created = False
        created |= self._ensure_launcher_bat(self.pad_bat_path, PAD_BAT_TEMPLATE)
        created |= self._ensure_launcher_bat(
            self.pad_create_new_bat_path,
            PAD_NEW_SQT_BAT_TEMPLATE,
        )
        created |= self._ensure_launcher_bat(
            self.pad_input_information_bat_path,
            PAD_INPUT_INFORMATION_BAT_TEMPLATE,
        )
        created |= self._ensure_launcher_bat(
            self.pad_input_expense_bat_path,
            PAD_INPUT_EXPENSE_BAT_TEMPLATE,
        )
        return created

    def _ensure_launcher_bat(self, path: str, template: str) -> bool:
        if not path or os.path.exists(path):
            return False

        launcher_dir = os.path.join(self.app_root, "Launcher")
        if os.path.normcase(os.path.dirname(os.path.abspath(path))) != os.path.normcase(
            os.path.abspath(launcher_dir)
        ):
            return False

        os.makedirs(launcher_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(template)
        return True
