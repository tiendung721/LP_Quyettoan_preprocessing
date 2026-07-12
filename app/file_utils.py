"""Các hàm tiện ích thao tác với file output.

Bao gồm: kiểm tra file bị khóa, nhận diện file tạm khi tải, kiểm tra file
hợp lệ theo cấu hình, chờ file tải xong (ổn định dung lượng), tính hash,
dọn file JSON tạm, mở file / mở thư mục.
"""

from __future__ import annotations

import fnmatch
import hashlib
import os
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional


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
# Thông tin file tải về (giữ nguyên tại thư mục Downloads)
# ---------------------------------------------------------------------------
def download_file_info(download_path: str) -> Dict:
    """Thu thập thông tin file vừa tải về mà KHÔNG di chuyển/sao lưu.

    File JSON tải về là bộ nhớ tạm của một lô bóc tách: nó nằm nguyên trong thư
    mục Downloads để người dùng xem/sửa ở Bước 2 và để Bước 3 đọc trực tiếp.

    Trả về dict thông tin file đang làm việc (chính là file trong Downloads).
    """
    if not os.path.exists(download_path):
        raise FileNotFoundError(download_path)

    return {
        "original_download_path": download_path,
        "backup_path": None,
        "working_path": download_path,
        "file_name": os.path.basename(download_path),
        "file_size": os.path.getsize(download_path),
        "file_hash": sha256_file(download_path),
        "detected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def is_in_folder(path: str, folder: str) -> bool:
    """True nếu ``path`` nằm ngay trong ``folder`` (so sánh không phân biệt hoa/thường)."""
    if not path or not folder:
        return False
    try:
        parent = os.path.abspath(os.path.dirname(path))
        target = os.path.abspath(folder)
    except (OSError, ValueError):
        return False
    return os.path.normcase(parent) == os.path.normcase(target)


def remove_file(path: str) -> bool:
    """Xóa một file nếu còn tồn tại; trả về True nếu đã xóa được.

    Dùng để dọn file JSON tạm (lô cũ đã nhập xong hoặc bị lô mới thay thế).
    Không ném lỗi: xóa không được thì chỉ báo False cho caller ghi log.
    """
    if not path or not os.path.exists(path):
        return False
    try:
        os.remove(path)
        return True
    except OSError:
        return False


def find_latest_output_file(
    folder: str,
    allowed_extensions: List[str],
    patterns: List[str],
) -> Optional[str]:
    """Tìm file bóc tách hợp lệ mới nhất (theo thời gian sửa) trong ``folder``.

    Dùng khi mở lại app mà bản ghi cũ không còn file: phần mềm tự nhận file mới
    nhất trong thư mục tải về thay vì bắt người dùng chọn tay.
    """
    if not folder or not os.path.isdir(folder):
        return None

    latest_path: Optional[str] = None
    latest_mtime = -1.0
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        if is_temp_download_file(path):
            continue
        if not is_allowed_output_file(path, allowed_extensions, patterns):
            continue
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            continue
        if mtime > latest_mtime:
            latest_mtime = mtime
            latest_path = path
    return latest_path


def last_saved_text(path: str) -> str:
    """Trả về thời điểm lưu gần nhất của file dạng 'HH:MM ngày dd/mm/yyyy'.

    Lấy từ mtime của file chứ không lưu trong database, vì người dùng bấm Lưu
    trong JSON editor nên mtime là nguồn duy nhất luôn đúng.
    Trả về chuỗi rỗng nếu file không còn tồn tại.
    """
    try:
        mtime = os.path.getmtime(path)
    except (OSError, TypeError):
        return ""
    return datetime.fromtimestamp(mtime).strftime("%H:%M ngày %d/%m/%Y")


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
