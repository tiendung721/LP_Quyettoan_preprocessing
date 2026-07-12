"""Lớp truy cập SQLite lưu trạng thái các file đã xử lý.

Mỗi thao tác mở một kết nối riêng để an toàn khi gọi từ nhiều luồng
(watchdog chạy ở luồng nền, UI ở luồng chính).
"""

from __future__ import annotations

import sqlite3
import json
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterator, List, Optional

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
    output_path TEXT,
    note TEXT
);
"""

_CREATE_DAILY_IMPORT_TABLES_SQL = """
CREATE TABLE IF NOT EXISTS source_documents (
    md5 TEXT PRIMARY KEY,
    document_type TEXT NOT NULL,
    source_name TEXT,
    status TEXT NOT NULL DEFAULT 'PENDING',
    last_output_record_id INTEGER,
    first_seen_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    completed_at TEXT
);

CREATE TABLE IF NOT EXISTS staged_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_md5 TEXT NOT NULL,
    row_key TEXT NOT NULL,
    source_row_number INTEGER,
    document_type TEXT NOT NULL,
    container_norm TEXT,
    state TEXT NOT NULL DEFAULT 'PENDING',
    data_json TEXT NOT NULL,
    matched_sqt INTEGER,
    selected_bill_md5 TEXT,
    note TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(document_md5, row_key),
    FOREIGN KEY(document_md5) REFERENCES source_documents(md5)
);

CREATE TABLE IF NOT EXISTS daily_import_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    processed_file_id INTEGER,
    output_path TEXT NOT NULL,
    daily_path TEXT NOT NULL,
    backup_path TEXT,
    status TEXT NOT NULL,
    summary_json TEXT,
    error_message TEXT,
    created_at TEXT NOT NULL,
    completed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_staged_rows_state ON staged_rows(state);
CREATE INDEX IF NOT EXISTS idx_staged_rows_container ON staged_rows(container_norm);
"""


def _now() -> str:
    """Thời gian local dạng chuỗi (không dùng UTC)."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Database:
    def __init__(self, db_path: str):
        self.db_path = db_path

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    # ------------------------------------------------------------------ #
    def init_db(self) -> None:
        """Tạo bảng nếu chưa tồn tại (và bổ sung cột mới cho DB cũ)."""
        with self._connect() as conn:
            conn.execute(_CREATE_TABLE_SQL)
            # Migration: DB từ phiên bản cũ có thể chưa có cột output_path.
            cols = {row["name"] for row in conn.execute(
                "PRAGMA table_info(processed_files)"
            )}
            if "output_path" not in cols:
                conn.execute(
                    "ALTER TABLE processed_files ADD COLUMN output_path TEXT"
                )
            conn.executescript(_CREATE_DAILY_IMPORT_TABLES_SQL)
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
        output_path: Optional[str] = None,
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
        if output_path is not None:
            fields.append("output_path = ?")
            params.append(output_path)

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
    def get_file_by_working_path(self, working_path: str) -> Optional[Dict[str, Any]]:
        """Lấy bản ghi mới nhất ứng với một đường dẫn file đang làm việc."""
        if not working_path:
            return None
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM processed_files
                WHERE working_path = ?
                ORDER BY id DESC LIMIT 1
                """,
                (working_path,),
            ).fetchone()
            return dict(row) if row else None

    # ------------------------------------------------------------------ #
    # Dữ liệu nguồn và hàng chờ của Bước 3
    # ------------------------------------------------------------------ #
    def upsert_source_document(
        self,
        md5: str,
        document_type: str,
        source_name: str = "",
        output_record_id: Optional[int] = None,
    ) -> None:
        now = _now()
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO source_documents (
                    md5, document_type, source_name, status,
                    last_output_record_id, first_seen_at, updated_at
                ) VALUES (?, ?, ?, 'PENDING', ?, ?, ?)
                ON CONFLICT(md5) DO UPDATE SET
                    document_type = excluded.document_type,
                    source_name = CASE
                        WHEN excluded.source_name <> '' THEN excluded.source_name
                        ELSE source_documents.source_name
                    END,
                    last_output_record_id = COALESCE(
                        excluded.last_output_record_id,
                        source_documents.last_output_record_id
                    ),
                    updated_at = excluded.updated_at
                """,
                (md5, document_type, source_name or "", output_record_id, now, now),
            )
            conn.commit()

    def upsert_staged_row(
        self,
        document_md5: str,
        row_key: str,
        source_row_number: int,
        document_type: str,
        container_norm: str,
        data: Dict[str, Any],
        state: str = "PENDING",
        note: Optional[str] = None,
    ) -> int:
        now = _now()
        payload = json.dumps(data, ensure_ascii=False, default=str)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO staged_rows (
                    document_md5, row_key, source_row_number, document_type,
                    container_norm, state, data_json, note, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(document_md5, row_key) DO UPDATE SET
                    source_row_number = excluded.source_row_number,
                    document_type = excluded.document_type,
                    container_norm = excluded.container_norm,
                    data_json = CASE
                        WHEN staged_rows.state IN ('COMPLETED', 'IGNORED')
                            THEN staged_rows.data_json
                        ELSE excluded.data_json
                    END,
                    state = CASE
                        WHEN staged_rows.state IN ('COMPLETED', 'IGNORED')
                            THEN staged_rows.state
                        ELSE excluded.state
                    END,
                    note = COALESCE(excluded.note, staged_rows.note),
                    updated_at = excluded.updated_at
                """,
                (
                    document_md5,
                    row_key,
                    source_row_number,
                    document_type,
                    container_norm,
                    state,
                    payload,
                    note,
                    now,
                    now,
                ),
            )
            row = conn.execute(
                "SELECT id FROM staged_rows WHERE document_md5 = ? AND row_key = ?",
                (document_md5, row_key),
            ).fetchone()
            conn.commit()
            return int(row["id"])

    def list_staged_rows(
        self,
        include_completed: bool = False,
        limit: int = 2000,
    ) -> List[Dict[str, Any]]:
        where = "" if include_completed else "WHERE state NOT IN ('COMPLETED', 'IGNORED')"
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT * FROM staged_rows
                {where}
                ORDER BY updated_at DESC, id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            try:
                item["data"] = json.loads(item.pop("data_json"))
            except (json.JSONDecodeError, TypeError):
                item["data"] = {}
                item.pop("data_json", None)
            result.append(item)
        return result

    def get_staged_row(self, row_id: int) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM staged_rows WHERE id = ?", (row_id,)
            ).fetchone()
        if not row:
            return None
        item = dict(row)
        try:
            item["data"] = json.loads(item.pop("data_json"))
        except (json.JSONDecodeError, TypeError):
            item["data"] = {}
            item.pop("data_json", None)
        return item

    def update_staged_row(
        self,
        row_id: int,
        *,
        state: Optional[str] = None,
        data: Optional[Dict[str, Any]] = None,
        matched_sqt: Optional[int] = None,
        selected_bill_md5: Optional[str] = None,
        note: Optional[str] = None,
    ) -> None:
        fields = ["updated_at = ?"]
        params: List[Any] = [_now()]
        if state is not None:
            fields.append("state = ?")
            params.append(state)
        if data is not None:
            fields.append("data_json = ?")
            params.append(json.dumps(data, ensure_ascii=False, default=str))
            fields.append("container_norm = ?")
            params.append(str(data.get("container") or ""))
        if matched_sqt is not None:
            fields.append("matched_sqt = ?")
            params.append(matched_sqt)
        if selected_bill_md5 is not None:
            fields.append("selected_bill_md5 = ?")
            params.append(selected_bill_md5)
        if note is not None:
            fields.append("note = ?")
            params.append(note)
        params.append(row_id)
        with self._connect() as conn:
            conn.execute(
                f"UPDATE staged_rows SET {', '.join(fields)} WHERE id = ?", params
            )
            conn.commit()

    def count_pending_rows(self) -> int:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT COUNT(*) AS count
                FROM staged_rows
                WHERE state NOT IN ('COMPLETED', 'IGNORED')
                """
            ).fetchone()
            return int(row["count"])

    def delete_staged_rows(self, row_ids: List[int]) -> int:
        """Xóa hẳn một số dòng tạm (dòng không ghép được, dòng bị người dùng bỏ)."""
        values = [int(row_id) for row_id in row_ids if row_id]
        if not values:
            return 0
        placeholders = ",".join("?" for _ in values)
        with self._connect() as conn:
            cur = conn.execute(
                f"DELETE FROM staged_rows WHERE id IN ({placeholders})", values
            )
            conn.commit()
            return int(cur.rowcount or 0)

    def clear_pending_staged_rows(self) -> int:
        """Dọn toàn bộ dòng tạm chưa hoàn tất.

        File JSON bóc tách là bộ nhớ tạm nên mỗi lần phân tích lại phải dựng lại
        hàng chờ từ đúng file đang dùng. Các dòng đã COMPLETED/IGNORED được giữ
        lại để còn nhận ra chứng từ đã nhập trước đó.
        """
        with self._connect() as conn:
            cur = conn.execute(
                "DELETE FROM staged_rows WHERE state NOT IN ('COMPLETED', 'IGNORED')"
            )
            conn.commit()
            return int(cur.rowcount or 0)

    def refresh_document_status(self, md5: str) -> str:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT state FROM staged_rows WHERE document_md5 = ?", (md5,)
            ).fetchall()
            states = [row["state"] for row in rows]
            if not states:
                status = "PENDING"
            elif all(state in ("COMPLETED", "IGNORED") for state in states):
                status = (
                    "IGNORED" if all(state == "IGNORED" for state in states)
                    else "COMPLETED"
                )
            elif any(state == "COMPLETED" for state in states):
                status = "PARTIAL"
            elif any(state == "FAILED" for state in states):
                status = "FAILED"
            else:
                status = "PENDING"
            completed_at = _now() if status in ("COMPLETED", "IGNORED") else None
            conn.execute(
                """
                UPDATE source_documents
                SET status = ?, updated_at = ?, completed_at = ?
                WHERE md5 = ?
                """,
                (status, _now(), completed_at, md5),
            )
            conn.commit()
            return status

    def create_import_run(
        self,
        processed_file_id: Optional[int],
        output_path: str,
        daily_path: str,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO daily_import_runs (
                    processed_file_id, output_path, daily_path, status, created_at
                ) VALUES (?, ?, ?, 'RUNNING', ?)
                """,
                (processed_file_id, output_path, daily_path, _now()),
            )
            conn.commit()
            return int(cur.lastrowid)

    def finish_import_run(
        self,
        run_id: int,
        status: str,
        *,
        backup_path: Optional[str] = None,
        summary: Optional[Dict[str, Any]] = None,
        error_message: Optional[str] = None,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE daily_import_runs
                SET status = ?, backup_path = ?, summary_json = ?,
                    error_message = ?, completed_at = ?
                WHERE id = ?
                """,
                (
                    status,
                    backup_path,
                    json.dumps(summary or {}, ensure_ascii=False),
                    error_message,
                    _now(),
                    run_id,
                ),
            )
            conn.commit()
