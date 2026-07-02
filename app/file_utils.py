"""Các hàm tiện ích thao tác với file output.

Bao gồm: kiểm tra file bị khóa, nhận diện file tạm khi tải, kiểm tra file
hợp lệ theo cấu hình, chờ file tải xong (ổn định dung lượng), tính hash,
di chuyển file download vào thư mục output chuẩn, mở file / mở thư mục.
"""

from __future__ import annotations

import fnmatch
import hashlib
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, List


# ---------------------------------------------------------------------------
# Kiểm tra trạng thái file
# ---------------------------------------------------------------------------
def excel_temp_file_path(path: str) -> str:
    """Trả về đường dẫn file tạm dạng ~$<tên file> mà Excel tạo khi mở."""
    folder = os.path.dirname(path)
    name = os.path.basename(path)
    return os.path.join(folder, "~$" + name)


def is_file_locked(path: str) -> bool:
    """Trả về True nếu file đang bị khóa (không mở được để đọc/ghi).

    - Với file Excel, còn kiểm tra sự tồn tại của file tạm ``~$<tên file>``
      trong cùng thư mục (dấu hiệu file đang được mở trong Excel).
    - Nếu file không tồn tại thì raise FileNotFoundError để caller xử lý.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    # Dấu hiệu file Excel đang mở: tồn tại file tạm ~$tenfile
    if os.path.exists(excel_temp_file_path(path)):
        return True

    # Thử mở file ở chế độ đọc + ghi (không truncate). Nếu bị khóa độc quyền
    # (ví dụ Excel đang giữ) sẽ ném PermissionError/OSError.
    try:
        with open(path, "r+b"):
            pass
    except (PermissionError, OSError):
        return True
    return False


def is_temp_download_file(path: str) -> bool:
    """Trả về True nếu là file tạm không nên xử lý.

    Gồm: đuôi .crdownload, .tmp, .part hoặc tên bắt đầu bằng '~$'.
    """
    name = os.path.basename(path).lower()
    if name.startswith("~$"):
        return True
    return name.endswith((".crdownload", ".tmp", ".part"))


def is_allowed_output_file(
    path: str,
    allowed_extensions: List[str],
    patterns: List[str],
) -> bool:
    """Kiểm tra file có thuộc đuôi cho phép và khớp mẫu tên hay không."""
    name = os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()

    allowed = [e.lower() for e in (allowed_extensions or [])]
    if allowed and ext not in allowed:
        return False

    # Không cấu hình mẫu -> chấp nhận mọi tên (đã lọc theo đuôi).
    if not patterns:
        return True

    name_lower = name.lower()
    for pat in patterns:
        if fnmatch.fnmatch(name_lower, pat.lower()):
            return True
    return False


# ---------------------------------------------------------------------------
# Chờ file tải xong
# ---------------------------------------------------------------------------
def wait_until_file_stable(
    path: str,
    stable_seconds: int = 3,
    timeout: int = 60,
) -> bool:
    """Chờ tới khi file có dung lượng ổn định và không bị khóa.

    Cơ chế: đọc dung lượng mỗi giây; nếu dung lượng giữ nguyên trong
    ``stable_seconds`` lần liên tiếp (tối thiểu 3 giây) và file không bị
    khóa thì coi là đã tải xong.

    Trả về True nếu ổn định, False nếu quá ``timeout`` giây vẫn chưa ổn định.
    """
    stable_seconds = max(int(stable_seconds), 3)
    start = time.time()
    last_size = -1
    stable_count = 0

    while time.time() - start < timeout:
        try:
            size = os.path.getsize(path)
        except OSError:
            # File có thể chưa sẵn sàng, thử lại.
            time.sleep(1)
            continue

        if size == last_size and size > 0:
            stable_count += 1
        else:
            stable_count = 0
        last_size = size

        if stable_count >= stable_seconds:
            try:
                if not is_file_locked(path):
                    return True
            except FileNotFoundError:
                return False
            # Vẫn còn bị khóa -> tiếp tục chờ.
            stable_count = 0

        time.sleep(1)

    return False


# ---------------------------------------------------------------------------
# Hash
# ---------------------------------------------------------------------------
def sha256_file(path: str, chunk_size: int = 1024 * 1024) -> str:
    """Tính mã băm SHA256 của file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Di chuyển file download vào output chuẩn
# ---------------------------------------------------------------------------
def safe_move_download_to_output(
    download_path: str,
    output_folder: str,
    backup_folder: str,
) -> Dict:
    """Sao lưu bản gốc rồi chuyển file download vào thư mục output theo ngày.

    - File làm việc (working) lưu vào ``output_folder\\YYYY-MM-DD\\``.
    - File backup bản gốc lưu vào ``backup_folder\\YYYY-MM-DD\\`` (thư mục riêng).
    - Tên file working: quyet_toan_output_YYYYMMDD_HHMMSS<ext>
    - Tên file backup:  backup_original_YYYYMMDD_HHMMSS<ext>
    - Luôn tạo backup bản gốc trước, không xóa/không ghi đè file cũ.
    - Giữ nguyên phần mở rộng gốc để đảm bảo đọc lại đúng (ví dụ .csv, .xlsm).

    Trả về dict thông tin file working.
    """
    if not os.path.exists(download_path):
        raise FileNotFoundError(download_path)

    now = datetime.now()
    date_dir = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    ext = os.path.splitext(download_path)[1].lower() or ".xlsx"

    # File working -> thư mục output theo ngày.
    output_day = os.path.join(output_folder, date_dir)
    os.makedirs(output_day, exist_ok=True)
    working_path = os.path.join(output_day, f"quyet_toan_output_{timestamp}{ext}")

    # File backup -> thư mục backup riêng theo ngày.
    backup_day = os.path.join(backup_folder, date_dir)
    os.makedirs(backup_day, exist_ok=True)
    backup_path = os.path.join(backup_day, f"backup_original_{timestamp}{ext}")

    # Nếu vì lý do nào đó tên đã tồn tại thì thêm hậu tố để không ghi đè.
    working_path = _unique_path(working_path)
    backup_path = _unique_path(backup_path)

    # 1) Tạo backup bản gốc trước (copy2 giữ metadata).
    shutil.copy2(download_path, backup_path)
    # 2) Chuyển bản gốc thành file working.
    shutil.move(download_path, working_path)

    file_size = os.path.getsize(working_path)
    file_hash = sha256_file(working_path)

    return {
        "original_download_path": download_path,
        "backup_path": backup_path,
        "working_path": working_path,
        "file_name": os.path.basename(working_path),
        "file_size": file_size,
        "file_hash": file_hash,
        "detected_at": now.strftime("%Y-%m-%d %H:%M:%S"),
    }


def _unique_path(path: str) -> str:
    """Trả về đường dẫn không trùng bằng cách thêm hậu tố _1, _2, ... nếu cần."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        candidate = f"{base}_{i}{ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


# ---------------------------------------------------------------------------
# Mở file / mở thư mục
# ---------------------------------------------------------------------------
def open_file(path: str) -> None:
    """Mở file bằng ứng dụng mặc định của Windows."""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]  # Windows
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def open_folder(path: str) -> None:
    """Mở thư mục (hoặc thư mục chứa file) bằng trình quản lý file."""
    target = path if os.path.isdir(path) else os.path.dirname(path)
    if not target or not os.path.exists(target):
        raise FileNotFoundError(target or path)
    if hasattr(os, "startfile"):
        os.startfile(target)  # type: ignore[attr-defined]  # Windows
    elif sys.platform == "darwin":
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])
