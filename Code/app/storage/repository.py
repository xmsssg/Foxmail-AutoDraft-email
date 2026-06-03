from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

from app.core.config_loader import expand_path


def now_text() -> str:
    return datetime.now().isoformat(timespec="seconds")


class DraftRepository:
    def __init__(self, db_path: str) -> None:
        self.db_path = Path(expand_path(db_path))
        self.db_path.parent.mkdir(parents=True, exist_ok=True)

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def init_schema(self) -> None:
        with self.connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS draft_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_id TEXT NOT NULL,
                    customer_name TEXT NOT NULL,

                    file_path TEXT NOT NULL,
                    file_name TEXT NOT NULL,
                    file_size INTEGER NOT NULL,
                    file_mtime TEXT NOT NULL,
                    file_hash TEXT NOT NULL,

                    subject TEXT,
                    to_recipients TEXT,
                    cc_recipients TEXT,
                    body TEXT,

                    mapi_xml_path TEXT,
                    eml_path TEXT,

                    import_status TEXT NOT NULL,
                    import_message TEXT,
                    imported_at TEXT,

                    foxmail_msg_id INTEGER,
                    foxmail_mail_path TEXT,

                    status TEXT NOT NULL,
                    error_message TEXT,

                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,

                    UNIQUE(file_path, file_hash)
                )
                """
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_draft_records_status "
                "ON draft_records(status)"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_draft_records_customer_id "
                "ON draft_records(customer_id)"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_draft_records_created_at "
                "ON draft_records(created_at)"
            )

    def find_by_file_hash(self, file_path: str, file_hash: str) -> sqlite3.Row | None:
        with self.connect() as conn:
            return conn.execute(
                """
                SELECT * FROM draft_records
                WHERE file_path = ? AND file_hash = ?
                """,
                (file_path, file_hash),
            ).fetchone()

    def create_pending(
        self,
        *,
        customer_id: str,
        customer_name: str,
        file_path: str,
        file_name: str,
        file_size: int,
        file_mtime: str,
        file_hash: str,
    ) -> int:
        timestamp = now_text()
        with self.connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO draft_records (
                    customer_id, customer_name,
                    file_path, file_name, file_size, file_mtime, file_hash,
                    import_status, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    customer_id,
                    customer_name,
                    file_path,
                    file_name,
                    file_size,
                    file_mtime,
                    file_hash,
                    "pending",
                    "pending",
                    timestamp,
                    timestamp,
                ),
            )
            return int(cursor.lastrowid)

    def update_record(self, record_id: int, **fields: object) -> None:
        if not fields:
            return
        fields["updated_at"] = now_text()
        names = list(fields.keys())
        values = [fields[name] for name in names]
        assignments = ", ".join(f"{name} = ?" for name in names)
        with self.connect() as conn:
            conn.execute(
                f"UPDATE draft_records SET {assignments} WHERE id = ?",
                [*values, record_id],
            )

    def list_recent(self, limit: int = 50) -> list[sqlite3.Row]:
        with self.connect() as conn:
            return list(
                conn.execute(
                    """
                    SELECT
                        id, customer_name, file_name, subject,
                        status, import_status, foxmail_msg_id,
                        error_message, updated_at
                    FROM draft_records
                    ORDER BY id DESC
                    LIMIT ?
                    """,
                    (limit,),
                )
            )
