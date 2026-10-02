from __future__ import annotations

from html import escape
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs
from pathlib import Path

from .db import AttendanceRepository, DuplicateDocumentError
from .config import AppConfig
from .service import export_pending
from .validation import FIELDS, ValidationError, build_record
from .logging_setup import configure_logging
import logging


logger = logging.getLogger("attendance_app")


FIELD_LABELS = {
    "full_name": "Full name",
    "document_type": "Document type",
    "document_number": "Document number",
    "gender": "Gender",
    "training_sheet_number": "Training sheet number",
    "modality": "Training modality",
    "program": "Training program",
    "schedule": "Schedule",
    "education_level": "Education level",
    "training_center": "Training center",
    "email": "Email",
    "phone": "Phone",
}


def _page(title: str, body: str) -> bytes:
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>{escape(title)}</title>
<style>body{{font-family:system-ui;max-width:760px;margin:2rem auto;padding:0 1rem}}label{{display:block;margin-top:1rem;font-weight:600}}input,select{{box-sizing:border-box;width:100%;padding:.55rem;margin-top:.25rem}}button{{margin-top:1.5rem;padding:.65rem 1rem}}.error{{color:#a00}}.success{{color:#075}}ul{{padding-left:1.25rem}}</style></head>
<body><h1>{escape(title)}</h1>{body}</body></html>""".encode("utf-8")


def form_page(message: str = "", error: bool = False) -> bytes:
    fields = []
    for name in FIELDS:
        label = escape(FIELD_LABELS[name])
        if name == "document_type":
            control = '<select name="document_type" required><option value="">Select...</option><option>CC</option><option>TI</option><option>CE</option></select>'
        elif name == "gender":
            control = '<select name="gender" required><option value="">Select...</option><option>F</option><option>M</option><option>Otro</option></select>'
        else:
            control = f'<input name="{escape(name)}" required>'
        fields.append(f"<label>{label}{control}</label>")
    notice = f'<p class="{"error" if error else "success"}">{escape(message)}</p>' if message else ""
    return _page("Attendance registration", notice + "<form method=post action=/submit>" + "".join(fields) + '<button type=submit>Register attendance</button></form>')


class AttendanceHandler(BaseHTTPRequestHandler):
    repository: AttendanceRepository
    campaign: str
    config: AppConfig

    def _send(self, status: int, content: bytes) -> None:
        try:
            self.send_response(status)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            # Browsers may cancel a request while navigating or refreshing.
            return

    def do_GET(self) -> None:
        logger.info("request start method=GET path=%s", self.path)
        if self.path != "/":
            self._send(404, _page("Not found", "<p>Page not found.</p>"))
            return
        self._send(200, form_page())

    def do_POST(self) -> None:
        logger.info("request start method=POST path=%s campaign=%s", self.path, self.campaign)
        if self.path != "/submit":
            self._send(404, _page("Not found", "<p>Page not found.</p>"))
            return
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length)
        fields = {key: values[0] for key, values in parse_qs(raw.decode("utf-8"), keep_blank_values=True).items()}
        try:
            record = build_record(fields)
            saved = self.repository.save(self.campaign, record)
        except (ValidationError, DuplicateDocumentError) as error:
            logger.warning("request rejected campaign=%s reason=%s", self.campaign, type(error).__name__)
            self._send(400, form_page(str(error), error=True))
            return
        except Exception:
            logger.exception("unexpected request error campaign=%s", self.campaign)
            self._send(500, _page("Unexpected error", "<p class=error>Please try again later.</p>"))
            return
        logger.info("database save successful record_id=%s campaign=%s", saved.id, self.campaign)
        recoveries = export_pending(self.repository, self.config, self.campaign) if self.config.workbook_path.exists() else []
        failed = [item for item in recoveries if item.status == "failed"]
        if failed:
            logger.error("export failed campaign=%s record_id=%s", self.campaign, failed[0].record_id)
            detail = escape(failed[0].detail)
            self._send(503, _page("Attendance saved; export failed", f"<p class=error>Registration saved, but workbook export failed: {detail}</p><p>Please retry after fixing the workbook.</p>"))
            return
        logger.info("export successful campaign=%s record_id=%s", self.campaign, saved.id)
        self._send(200, _page("Attendance registered", f"<p class=success>Registration accepted and exported. Arrival number: {saved.arrival_seq}.</p><p><a href=/>Register another person</a></p>"))

    def log_message(self, format: str, *args: object) -> None:
        return


def make_server(repository: AttendanceRepository, campaign: str, host: str = "127.0.0.1", port: int = 0, config: AppConfig | None = None) -> HTTPServer:
    config = config or AppConfig(Path.cwd(), campaign=campaign)
    configure_logging(config)
    handler = type("ConfiguredAttendanceHandler", (AttendanceHandler,), {"repository": repository, "campaign": campaign, "config": config})
    return HTTPServer((host, port), handler)


def serve(repository: AttendanceRepository, campaign: str, host: str = "127.0.0.1", port: int = 8000) -> None:
    server = make_server(repository, campaign, host, port)
    try:
        server.serve_forever()
    finally:
        server.server_close()
