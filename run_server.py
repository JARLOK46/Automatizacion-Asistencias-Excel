from attendance_app.config import AppConfig
from attendance_app.db import AttendanceRepository
from pathlib import Path

from attendance_app.web import serve


if __name__ == "__main__":
    config = AppConfig(Path(__file__).resolve().parent)
    repository = AttendanceRepository(config.database_path)
    try:
        print("Attendance form running at http://127.0.0.1:8000")
        print("Press Ctrl+C to stop.")
        serve(repository, config.campaign)
    finally:
        repository.close()
