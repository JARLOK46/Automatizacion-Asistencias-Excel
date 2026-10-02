from dataclasses import dataclass
from pathlib import Path


CAMPAIGN_NAME = "CAMPANA DE SENSIBILIZACION EN VALORES"
DISPLAY_DATETIME_FORMAT = "%d/%m/%Y %H:%M"


@dataclass(frozen=True)
class AppConfig:
    """Local paths and constants shared by future CLI/export layers."""

    root: Path
    campaign: str = CAMPAIGN_NAME

    @property
    def database_path(self) -> Path:
        return self.root / "attendance.db"

    @property
    def log_path(self) -> Path:
        return self.root / "attendance.log"

    @property
    def workbook_path(self) -> Path:
        """The project-root workbook is both the export source and destination."""
        return self.root / "Listado ejemplo.xlsx"

    @property
    def backups_path(self) -> Path:
        """Directory reserved for backup copies, never workbook discovery."""
        return self.root / "backups"
