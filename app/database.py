"""Lớp truy cập SQLite lưu trạng thái các file đã xử lý.

Mỗi thao tác mở một kết nối riêng để an toàn khi gọi từ nhiều luồng
(watchdog chạy ở luồng nền, UI ở luồng chính).
"""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any, Dict, List, Optional

_CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS processed_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_download_path TEXT,
    backup_path TEXT,
    working_path TEXT NOT NULL,
    file_name TEXT,
    file_size INTEGER,
    file_hash TEXT,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    reviewed_at TEXT,
    previewed_at TEXT,
    note TEXT
);
"""


def _now() -> str:
    """Thời gian local dạng chuỗi (không dùng UTC)."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Database:
    def __init__(self, db_path: str):
        self.db_path = db_path

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    # ------------------------------------------------------------------ #
    def init_db(self) -> None:
        """Tạo bảng nếu chưa tồn tại."""
        with self._connect() as conn:
            conn.execute(_CREATE_TABLE_SQL)
            conn.commit()

    # ------------------------------------------------------------------ #
    def insert_processed_file(
        self,
        working_path: str,
        status: str,
        original_download_path: Optional[str] = None,
        backup_path: Optional[str] = None,
        file_name: Optional[str] = None,
        file_size: Optional[int] = None,
        file_hash: Optional[str] = None,
        note: Optional[str] = None,
    ) -> int:
        """Thêm một bản ghi file mới, trả về id vừa tạo."""
        now = _now()
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO processed_files (
                    original_download_path, backup_path, working_path,
                    file_name, file_size, file_hash, status,
                    created_at, updated_at, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    original_download_path,
                    backup_path,
                    working_path,
                    file_name,
                    file_size,
                    file_hash,
                    status,
                    now,
                    now,
                    note,
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    # ------------------------------------------------------------------ #
    def update_status(
        self,
        record_id: int,
        status: str,
        note: Optional[str] = None,
        mark_reviewed: bool = False,
        mark_previewed: bool = False,
    ) -> None:
        """Cập nhật trạng thái của một bản ghi."""
        now = _now()
        fields = ["status = ?", "updated_at = ?"]
        params: List[Any] = [status, now]

        if note is not None:
            fields.append("note = ?")
            params.append(note)
        if mark_reviewed:
            fields.append("reviewed_at = ?")
            params.append(now)
        if mark_previewed:
            fields.append("previewed_at = ?")
            params.append(now)

        params.append(record_id)
        with self._connect() as conn:
            conn.execute(
                f"UPDATE processed_files SET {', '.join(fields)} WHERE id = ?",
                params,
            )
            conn.commit()

    # ------------------------------------------------------------------ #
    def get_latest_file(self) -> Optional[Dict[str, Any]]:
        """Lấy bản ghi mới nhất (theo id lớn nhất)."""
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM processed_files ORDER BY id DESC LIMIT 1"
            ).fetchone()
            return dict(row) if row else None

    # ------------------------------------------------------------------ #
    def get_all_processed_files(self, limit: int = 200) -> List[Dict[str, Any]]:
        """Lấy danh sách bản ghi, mới nhất trước."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM processed_files ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
            return [dict(r) for r in rows]

    # ------------------------------------------------------------------ #
    def get_file_by_id(self, record_id: int) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM processed_files WHERE id = ?", (record_id,)
            ).fetchone()
            return dict(row) if row else None
