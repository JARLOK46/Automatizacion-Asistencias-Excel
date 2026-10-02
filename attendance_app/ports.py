from typing import Protocol


class AttendanceStore(Protocol):
    """Persistence boundary for the Python-owned SQLite implementation."""

    def close(self) -> None: ...
