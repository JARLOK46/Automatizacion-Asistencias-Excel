from __future__ import annotations

from dataclasses import dataclass
import sqlite3
from pathlib import Path
from threading import RLock
from typing import Iterable

from .validation import AttendanceRecordInput
from .config import AppConfig
from .logging_setup import configure_logging


logger = configure_logging(AppConfig(Path.cwd()))


class DuplicateDocumentError(ValueError):
    """Raised when a document is already registered in the campaign."""


@dataclass(frozen=True)
class StoredAttendanceRecord:
    id: int
    campaign: str
    arrival_seq: int
    document_number: str
    export_status: str
    submission_id: str
    exported_row: int | None = None
    failure_detail: str | None = None


SCHEMA = """
CREATE TABLE IF NOT EXISTS attendance_records (
    id INTEGER PRIMARY KEY,
    campaign TEXT NOT NULL,
    arrival_seq INTEGER NOT NULL,
    submission_id TEXT UNIQUE NOT NULL,
    full_name TEXT NOT NULL,
    document_type TEXT NOT NULL CHECK (document_type IN ('CC', 'TI', 'CE')),
    document_number TEXT NOT NULL,
    gender TEXT NOT NULL CHECK (gender IN ('F', 'M', 'Otro')),
    training_sheet_number TEXT NOT NULL,
    modality TEXT NOT NULL,
    program TEXT NOT NULL,
    schedule TEXT NOT NULL,
    education_level TEXT NOT NULL,
    training_center TEXT NOT NULL,
    email TEXT NOT NULL,
    phone TEXT NOT NULL,
    registered_at_utc TEXT NOT NULL,
    registered_at_local TEXT NOT NULL,
    timezone TEXT NOT NULL,
    export_status TEXT NOT NULL DEFAULT 'pending'
        CHECK (export_status IN ('pending', 'exported', 'failed')),
    exported_row INTEGER,
    failure_detail TEXT,
    created_at_utc TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now')),
    UNIQUE (campaign, document_number),
    UNIQUE (campaign, arrival_seq)
);
CREATE INDEX IF NOT EXISTS idx_attendance_campaign_arrival
    ON attendance_records (campaign, arrival_seq);
CREATE TABLE IF NOT EXISTS attendance_rejections (
    id INTEGER PRIMARY KEY,
    campaign TEXT NOT NULL,
    document_number TEXT NOT NULL,
    reason TEXT NOT NULL,
    rejected_at_utc TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now'))
);
"""


class AttendanceRepository:
    def __init__(self, path: str | Path):
        # The localhost HTTP server handles requests on a worker thread.
        self._lock = RLock()
        # Direct repository users still receive operational logging.
        global logger
        logger = configure_logging(AppConfig(Path(path).parent))
        logger.info("database initialization path=%s", path)
        self.connection = sqlite3.connect(path, check_same_thread=False)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.executescript(SCHEMA)
        self.connection.commit()

    def close(self) -> None:
        with self._lock:
            self.connection.close()

    def save(self, campaign: str, record: AttendanceRecordInput) -> StoredAttendanceRecord:
        with self._lock:
            try:
                with self.connection:
                    next_seq = self.connection.execute(
                        "SELECT COALESCE(MAX(arrival_seq), 0) + 1 FROM attendance_records WHERE campaign = ?",
                        (campaign,),
                    ).fetchone()[0]
                    cursor = self.connection.execute(
                        """INSERT INTO attendance_records
                        (campaign, arrival_seq, submission_id, full_name, document_type,
                         document_number, gender, training_sheet_number, modality, program,
                         schedule, education_level, training_center, email, phone,
                         registered_at_utc, registered_at_local, timezone)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (campaign, next_seq, record.submission_id, record.full_name,
                         record.document_type, record.document_number, record.gender,
                         record.training_sheet_number, record.modality, record.program,
                         record.schedule, record.education_level, record.training_center,
                         record.email, record.phone, record.registered_at_utc,
                         record.registered_at_local, record.timezone),
                    )
            except sqlite3.IntegrityError as error:
                if "attendance_records.campaign, attendance_records.document_number" in str(error):
                    self._audit_rejection(campaign, record.document_number, "duplicate document")
                    logger.warning("duplicate attendance rejected campaign=%s", campaign)
                    raise DuplicateDocumentError(
                        f"document {record.document_number} is already registered in {campaign}"
                    ) from error
                logger.exception("database integrity error saving attendance campaign=%s", campaign)
                raise
            except Exception:
                logger.exception("unexpected database error saving attendance campaign=%s", campaign)
                raise
            logger.info("attendance saved record_id=%s campaign=%s", cursor.lastrowid, campaign)
            return StoredAttendanceRecord(cursor.lastrowid, campaign, next_seq,
                                          record.document_number, "pending", record.submission_id)

    def _audit_rejection(self, campaign: str, document_number: str, reason: str) -> None:
        with self.connection:
            self.connection.execute(
                "INSERT INTO attendance_rejections (campaign, document_number, reason) VALUES (?, ?, ?)",
                (campaign, document_number, reason),
            )

    def list_records(self, campaign: str) -> list[StoredAttendanceRecord]:
        with self._lock:
            rows = self.connection.execute(
            "SELECT id, campaign, arrival_seq, document_number, export_status, submission_id, exported_row, failure_detail "
            "FROM attendance_records WHERE campaign = ? ORDER BY arrival_seq", (campaign,)
        )
            return [StoredAttendanceRecord(**dict(row)) for row in rows]

    def pending(self, campaign: str) -> list[StoredAttendanceRecord]:
        return [record for record in self.list_records(campaign) if record.export_status == "pending"]

    def mark_exported(self, record_id: int, row: int) -> None:
        with self._lock, self.connection:
            self.connection.execute(
                "UPDATE attendance_records SET export_status = 'exported', exported_row = ?, failure_detail = NULL WHERE id = ?",
                (row, record_id),
            )

    def mark_failed(self, record_id: int, detail: str) -> None:
        with self._lock, self.connection:
            self.connection.execute(
                "UPDATE attendance_records SET export_status = 'failed', failure_detail = ? WHERE id = ?",
                (detail, record_id),
            )

    def list_rejections(self, campaign: str) -> list[sqlite3.Row]:
        with self._lock:
            return list(self.connection.execute(
            "SELECT * FROM attendance_rejections WHERE campaign = ? ORDER BY id", (campaign,)
        ))
