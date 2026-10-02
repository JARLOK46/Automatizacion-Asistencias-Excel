from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime, timezone
import re
from typing import Mapping
from zoneinfo import ZoneInfo

from .config import DISPLAY_DATETIME_FORMAT


class ValidationError(ValueError):
    """Raised when a submitted attendance field violates its contract."""


@dataclass(frozen=True)
class AttendanceRecordInput:
    full_name: str
    document_type: str
    document_number: str
    gender: str
    training_sheet_number: str
    modality: str
    program: str
    schedule: str
    education_level: str
    training_center: str
    email: str
    phone: str
    registered_at_utc: str
    registered_at_local: str
    timezone: str
    submission_id: str


FIELDS = (
    "full_name", "document_type", "document_number", "gender",
    "training_sheet_number", "modality", "program", "schedule",
    "education_level", "training_center", "email", "phone",
)
_EMAIL = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
_DIGITS = re.compile(r"^\d+$")


def _required(fields: Mapping[str, object], name: str) -> str:
    value = fields.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{name} is required and must be non-empty")
    return " ".join(value.strip().split())


def build_record(fields: Mapping[str, object], now: datetime | None = None) -> AttendanceRecordInput:
    values = {name: _required(fields, name) for name in FIELDS}
    values["document_type"] = values["document_type"].upper()
    values["gender"] = values["gender"].title()
    values["email"] = values["email"].lower()

    if values["document_type"] not in {"CC", "TI", "CE"}:
        raise ValidationError("document_type must be CC, TI, or CE")
    if values["gender"] not in {"F", "M", "Otro"}:
        raise ValidationError("gender must be F, M, or Otro")
    for name in ("training_sheet_number", "phone"):
        if not _DIGITS.fullmatch(values[name]):
            raise ValidationError(f"{name} must contain digits only")
    if not _EMAIL.fullmatch(values["email"]):
        raise ValidationError("email has invalid syntax")

    instant = now or datetime.now().astimezone()
    if instant.tzinfo is None:
        raise ValidationError("now must be timezone-aware")
    local = instant.astimezone()
    utc = instant.astimezone(timezone.utc).replace(microsecond=0)
    return AttendanceRecordInput(
        **values,
        registered_at_utc=utc.isoformat().replace("+00:00", "Z"),
        registered_at_local=local.strftime(DISPLAY_DATETIME_FORMAT),
        timezone=getattr(local.tzinfo, "key", None) or str(local.tzinfo),
        submission_id=f"{utc.isoformat()}:{values['document_type']}:{values['document_number']}",
    )
