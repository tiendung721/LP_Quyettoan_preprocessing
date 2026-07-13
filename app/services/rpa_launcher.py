"""Launch Power Automate Desktop batch files."""

from __future__ import annotations

import os
import subprocess
import tempfile
from typing import Optional


class RpaLaunchError(RuntimeError):
    """Business error safe to show directly to the user."""


def _read_output_file(path: str) -> str:
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError:
        return ""

    for encoding in ("utf-8-sig", "mbcs", "cp1258", "cp437"):
        try:
            return raw.decode(encoding).strip()
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", errors="replace").strip()


def validate_bat_path(bat_path: str, missing_message: str) -> str:
    path = os.path.abspath(os.path.normpath(bat_path or ""))
    if not bat_path or not os.path.isfile(path):
        raise RpaLaunchError(missing_message)
    return path


def launch_bat(
    bat_path: str,
    *,
    logger,
    description: str,
    missing_message: str,
) -> subprocess.Popen:
    """Start a BAT file without waiting for the PAD flow to finish."""
    path = validate_bat_path(bat_path, missing_message)
    logger.info("Gọi BAT %s: %s", description, path)
    output_path: Optional[str] = None
    try:
        env = os.environ.copy()
        env.setdefault("PYTHONUTF8", "1")
        env.setdefault("PYTHONIOENCODING", "utf-8")
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix="rpa_launch_",
            suffix=".log",
            delete=False,
        ) as output_file:
            output_path = output_file.name
            process = subprocess.Popen(
                ["cmd.exe", "/d", "/c", path],
                cwd=os.path.dirname(path) or None,
                env=env,
                stdout=output_file,
                stderr=subprocess.STDOUT,
            )
    except Exception as exc:  # noqa: BLE001 - converted to a user-facing message
        if output_path:
            try:
                os.remove(output_path)
            except OSError:
                pass
        raise RpaLaunchError(f"Lỗi khi gọi tiến trình BAT: {exc}") from exc

    try:
        return_code = process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        logger.info("BAT %s vẫn đang chạy sau kiểm tra nhanh: %s", description, path)
        return process

    output = ""
    if output_path and os.path.isfile(output_path):
        output = _read_output_file(output_path)
        try:
            os.remove(output_path)
        except OSError:
            pass

    if return_code != 0:
        message = f"BAT kết thúc với mã lỗi {return_code}."
        if output:
            message = f"{message}\n\n{output}"
        raise RpaLaunchError(message)

    logger.info("BAT %s đã kết thúc nhanh với mã 0: %s", description, path)
    return process
