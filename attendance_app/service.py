from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone

from .db import AttendanceRepository, StoredAttendanceRecord
from .config import AppConfig
from .validation import AttendanceRecordInput
from .logging_setup import configure_logging
import editar_excel


@dataclass(frozen=True)
class ProjectionRecovery:
    record_id: int
    campaign: str
    status: str
    detail: str
    retryable: bool
    recorded_at_utc: str


def projection_failed(record: StoredAttendanceRecord, detail: str) -> ProjectionRecovery:
    return ProjectionRecovery(record.id, record.campaign, "failed", detail, True,
                              datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"))


def projection_succeeded(record: StoredAttendanceRecord, row: int) -> ProjectionRecovery:
    return ProjectionRecovery(record.id, record.campaign, f"exported:{row}", "", False,
                              datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"))


def retryable_records(repository: AttendanceRepository, campaign: str) -> list[StoredAttendanceRecord]:
    return [record for record in repository.list_records(campaign)
            if record.export_status in {"pending", "failed"}]


def mark_projection_failed(repository: AttendanceRepository, record: StoredAttendanceRecord,
                           detail: str) -> ProjectionRecovery:
    repository.mark_failed(record.id, detail)
    return projection_failed(record, detail)


def mark_projection_succeeded(repository: AttendanceRepository, record: StoredAttendanceRecord,
                              row: int) -> ProjectionRecovery:
    repository.mark_exported(record.id, row)
    return projection_succeeded(record, row)


def _input_for(repository: AttendanceRepository, record: StoredAttendanceRecord) -> AttendanceRecordInput:
    row = repository.connection.execute(
        "SELECT full_name, document_type, document_number, gender, training_sheet_number, modality, program, schedule, education_level, training_center, email, phone, registered_at_utc, registered_at_local, timezone, submission_id FROM attendance_records WHERE id = ?",
        (record.id,),
    ).fetchone()
    return AttendanceRecordInput(*row)


def _export_values(record: StoredAttendanceRecord, input_record: AttendanceRecordInput) -> dict:
    return {
        "arrival_seq": record.arrival_seq,
        "full_name": input_record.full_name,
        "document_type": input_record.document_type,
        "document_number": input_record.document_number,
        "gender": input_record.gender,
        "training_sheet_number": input_record.training_sheet_number,
        "modality": input_record.modality,
        "program": input_record.program,
        "schedule": input_record.schedule,
        "education_level": input_record.education_level,
        "training_center": input_record.training_center,
        "email": input_record.email,
        "phone": input_record.phone,
        "registered_at_local": input_record.registered_at_local,
    }


def _failure_detail(error: Exception, workbook, backups) -> str:
    report = editar_excel.diagnostic_report(workbook, backups)
    message = str(error)
    stage = "worksheet XML validation" if "Formato" in message or "worksheet" in message else "XLSX package validation"
    report_text = str(report) if report else "unavailable"
    return (f"{type(error).__name__}: {message}; workbook={workbook}; "
            f"stage={stage}; diagnostic_report={report_text}")


def export_pending(repository: AttendanceRepository, config: AppConfig, campaign: str) -> list[ProjectionRecovery]:
    logger = configure_logging(config)
    records = retryable_records(repository, campaign)
    logger.info("export attempt campaign=%s record_count=%s record_ids=%s", campaign, len(records), [r.id for r in records])
    if not records:
        return []
    workbook = config.workbook_path
    try:
        inputs = [_input_for(repository, record) for record in records]
        rows, _backup = editar_excel.export_records(
            [_export_values(record, item) for record, item in zip(records, inputs)],
            workbook, config.backups_path,
        )
    except Exception as error:
        detail = _failure_detail(error, workbook, config.backups_path)
        logger.exception("export failed campaign=%s record_ids=%s", campaign, [r.id for r in records])
        return [mark_projection_failed(repository, record, detail) for record in records]
    logger.info("export succeeded campaign=%s rows=%s record_ids=%s", campaign, rows, [r.id for r in records])
    return [mark_projection_succeeded(repository, record, row) for record, row in zip(records, rows)]
