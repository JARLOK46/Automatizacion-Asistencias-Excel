"""Append attendance records to the root XLSX using only ZIP and XML operations."""
from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import zipfile
import re
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "Listado ejemplo.xlsx"
BACKUPS = ROOT / "backups"
SHEET_NAME = "Formato"
START_ROW = 17
FIELDS = (
    "arrival_seq", "full_name", "document_type", "document_number", "gender",
    "training_sheet_number", "modality", "program", "schedule", "education_level",
    "training_center", "email", "phone", "registered_at_local",
)
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MARKUP_COMPAT_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XR2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XR3_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"


class InputError(ValueError):
    pass


def _xml_safe(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    if any(ord(c) < 0x20 and ord(c) not in (9, 10, 13) for c in text):
        raise InputError("input contains an XML-invalid control character")
    return text


def load_records(path: Path) -> list[dict[str, Any]]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InputError(f"invalid JSON input: {exc}") from exc
    records = value if isinstance(value, list) else [value]
    if not records or any(not isinstance(record, dict) for record in records):
        raise InputError("input must be one object or a non-empty array of objects")
    result = []
    for number, record in enumerate(records, 1):
        missing = [field for field in FIELDS if field not in record or record[field] is None]
        if missing:
            raise InputError(f"record {number} missing required fields: {', '.join(missing)}")
        for field in FIELDS:
            if isinstance(record[field], (dict, list, bool)):
                raise InputError(f"record {number} field {field} must be scalar")
            _xml_safe(record[field])
        try:
            int(record["arrival_seq"])
        except (TypeError, ValueError) as exc:
            raise InputError(f"record {number} arrival_seq must be an integer") from exc
        result.append(record)
    return result


def _column_number(value: str) -> int:
    n = 0
    for char in value:
        if char.isalpha():
            n = n * 26 + ord(char.upper()) - 64
    return n


def _column_letter(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def resolve_sheet(entries: dict[str, bytes]) -> str:
    workbook = ET.fromstring(entries["xl/workbook.xml"])
    rels = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
    targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(f"{{{PKG_REL_NS}}}Relationship")}
    for sheet in workbook.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet"):
        if sheet.attrib.get("name") == SHEET_NAME:
            rid = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = targets.get(rid)
            if not target:
                raise InputError("Formato worksheet relationship is missing")
            target_path = (Path("xl") / target).resolve() if not target.startswith("/") else Path(target.lstrip("/"))
            # resolve() above is host-specific; normalize the ZIP path explicitly.
            import posixpath
            return posixpath.normpath(posixpath.join("xl", target)).lstrip("/")
    raise InputError("workbook has no Formato worksheet")


def _clear_cell(cell: ET.Element) -> None:
    cell.attrib.pop("t", None)
    for child in list(cell):
        if child.tag.rsplit("}", 1)[-1] in {"v", "is", "f"}:
            cell.remove(child)


def _set_cell(cell: ET.Element, value: Any) -> None:
    _clear_cell(cell)
    if value is None or value == "":
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        child = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
        child.text = str(value)
    else:
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
        text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
        text.text = _xml_safe(value)


def _worksheet_xml(root: ET.Element, original: bytes) -> bytes:
    """Serialize a worksheet with Excel's conventional namespace declarations."""
    namespaces = {
        "": MAIN_NS, "r": REL_NS, "mc": MARKUP_COMPAT_NS,
        "x14ac": X14AC_NS, "xr": XR_NS, "xr2": XR2_NS, "xr3": XR3_NS,
    }
    # ElementTree discards prefixes while parsing. Recover additional declarations
    # from the source so extensions used by a particular workbook remain named.
    for match in re.finditer(rb"xmlns(?::([A-Za-z_][\w.-]*))?=[\"']([^\"']+)[\"']", original):
        prefix = (match.group(1) or b"").decode("ascii")
        uri = match.group(2).decode("utf-8")
        if uri not in namespaces.values():
            namespaces[prefix] = uri
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    body = ET.tostring(root, encoding="utf-8")
    start = body.find(b"<worksheet")
    end = body.find(b">", start)
    opening = body[start:end]
    for prefix, uri in namespaces.items():
        declaration = (b' xmlns="' if not prefix else b' xmlns:' + prefix.encode("ascii") + b'="') + uri.encode("utf-8") + b'"'
        if declaration not in opening and (b' xmlns="' + uri.encode("utf-8") + b'"') not in opening:
            opening += declaration
    body = body[:start] + opening + body[end:]
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body


def append_to_bytes(entries: dict[str, bytes], records: list[dict[str, Any]]) -> None:
    part = resolve_sheet(entries)
    original_xml = entries[part]
    root = ET.fromstring(original_xml)
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        raise InputError("Formato worksheet has no sheetData")
    rows = {int(row.attrib["r"]): row for row in sheet_data.findall(f"{{{MAIN_NS}}}row")}
    end = START_ROW - 1
    for row_number, row in rows.items():
        if row_number < START_ROW:
            continue
        cells = row.findall(f"{{{MAIN_NS}}}c")
        # Attendance is identified by a nonempty column B:N, not template sequence A.
        if any(2 <= _column_number(''.join(c for c in cell.attrib.get('r', '') if c.isalpha())) <= 14 and
               any(child.tag.rsplit('}', 1)[-1] in {'v', 'is'} and child.text is not None for child in cell)
               for cell in cells):
            end = max(end, row_number)
    for offset, record in enumerate(records):
        row_number = end + offset + 1
        row = rows.get(row_number)
        if row is None:
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_number), "spans": "1:14"})
            sheet_data.append(row)
        cells = {''.join(c for c in cell.attrib.get('r', '') if c.isalpha()): cell
                 for cell in row.findall(f"{{{MAIN_NS}}}c")}
        # New cells inherit the style of the nearest existing attendance row where possible.
        template = rows.get(max(START_ROW, end))
        if template is None:
            template = next(iter(rows.values()), None)
        template_cells = {} if template is None else {''.join(c for c in cell.attrib.get('r', '') if c.isalpha()): cell for cell in template.findall(f"{{{MAIN_NS}}}c")}
        for column, field in enumerate(FIELDS, 1):
            letter = _column_letter(column)
            cell = cells.get(letter)
            if cell is None:
                cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": f"{letter}{row_number}"})
                if letter in template_cells and "s" in template_cells[letter].attrib:
                    cell.set("s", template_cells[letter].attrib["s"])
                row.append(cell)
            _set_cell(cell, record[field])
    entries[part] = _worksheet_xml(root, original_xml)
    _validate_worksheet_bytes(entries[part])


def _written_rows(entries: dict[str, bytes]) -> tuple[int, ...]:
    part = resolve_sheet(entries)
    root = ET.fromstring(entries[part])
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        raise InputError("Formato worksheet has no sheetData")
    result = []
    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        number = int(row.attrib.get("r", "0"))
        if number < START_ROW:
            continue
        if any(2 <= _column_number(''.join(c for c in cell.attrib.get('r', '') if c.isalpha())) <= 14 and
               any(
                   (child.tag.rsplit('}', 1)[-1] == 'v' and child.text not in (None, ''))
                   or (child.tag.rsplit('}', 1)[-1] == 'is' and any(
                       text.tag.rsplit('}', 1)[-1] == 't' and text.text not in (None, '')
                       for text in child.iter()
                   ))
                   for child in cell
               )
               for cell in row.findall(f"{{{MAIN_NS}}}c")):
            result.append(number)
    return tuple(sorted(result))


def export_records(records: list[dict[str, Any]], workbook: Path,
                   backups_dir: Path) -> tuple[tuple[int, ...], Path]:
    """Create and validate a candidate workbook before replacing the source."""
    if not workbook.is_file():
        raise InputError(f"workbook does not exist: {workbook}")
    with zipfile.ZipFile(workbook) as source:
        infos = source.infolist()
        if len({info.filename for info in infos}) != len(infos):
            raise InputError("workbook contains duplicate ZIP entries")
        original = {info.filename: source.read(info.filename) for info in infos}
    entries = dict(original)
    before = set(_written_rows(entries))
    append_to_bytes(entries, records)
    rows = tuple(row for row in _written_rows(entries) if row not in before)
    _validate_package_candidate(original, entries)
    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup = backups_dir / f"{workbook.stem}.{stamp}{workbook.suffix}"
    shutil.copy2(workbook, backup)
    fd, temporary = tempfile.mkstemp(prefix=f".{workbook.name}.", suffix=".tmp", dir=workbook.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, content in entries.items():
                archive.writestr(name, content)
        validate_workbook(Path(temporary))
        os.replace(temporary, workbook)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return rows, backup


def _validate_worksheet_bytes(content: bytes) -> None:
    root = ET.fromstring(content)
    if root.tag != f"{{{MAIN_NS}}}worksheet":
        raise InputError("Formato worksheet has an invalid root element")
    ignorable = root.attrib.get(f"{{{MARKUP_COMPAT_NS}}}Ignorable", "").split()
    start = content.find(b"<worksheet")
    end = content.find(b">", start)
    header = content[start:end + 1]
    declared = set(re.findall(rb"xmlns(?::([A-Za-z_][\w.-]*))?=", header))
    declared_names = {(item or b"").decode("ascii") for item in declared}
    if any(prefix not in declared_names for prefix in ignorable):
        raise InputError("Formato worksheet has an undeclared ignorable namespace")
    if root.find(f"{{{MAIN_NS}}}sheetData") is None:
        raise InputError("Formato worksheet has no sheetData")


def _validate_package_candidate(original: dict[str, bytes], candidate: dict[str, bytes]) -> None:
    if set(original) != set(candidate):
        raise InputError("candidate changed XLSX package entries")
    target = resolve_sheet(original)
    for name, content in original.items():
        if name != target and candidate[name] != content:
            raise InputError(f"candidate changed protected package part: {name}")
    _validate_worksheet_bytes(candidate[target])


def validate_workbook(workbook: Path) -> dict[str, Any]:
    """Validate the XLSX package and Formato XML without modifying it."""
    with zipfile.ZipFile(workbook) as archive:
        names = set(archive.namelist())
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            raise InputError("XLSX package is missing workbook metadata")
        entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    part = resolve_sheet(entries)
    root = ET.fromstring(entries[part])
    if root.find(f"{{{MAIN_NS}}}sheetData") is None:
        raise InputError("Formato worksheet has no sheetData")
    # A second parse catches malformed serialized XML independently of the first tree.
    ET.fromstring(entries[part])
    return {"workbook": str(workbook), "worksheet": part, "rows": len(root.findall(f"{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"))}


def diagnostic_report(workbook: Path, backups_dir: Path) -> Path | None:
    """Write a small actionable validation report, best effort."""
    try:
        try:
            result = validate_workbook(workbook)
            status = "valid"
            error = ""
        except Exception as exc:
            result = {}
            status = "invalid"
            error = f"{type(exc).__name__}: {exc}"
        backups_dir.mkdir(parents=True, exist_ok=True)
        report = backups_dir / f"{workbook.stem}.diagnostic.txt"
        report.write_text("XLSX diagnostic\\n" + json.dumps({"status": status, **result, "error": error}, indent=2), encoding="utf-8")
        return report
    except (OSError, ValueError):
        return None


def run(records: list[dict[str, Any]], *, workbook: Path = DEFAULT_WORKBOOK, dry_run: bool = False) -> Path | None:
    if dry_run:
        if not workbook.is_file():
            raise InputError(f"workbook does not exist: {workbook}")
        with zipfile.ZipFile(workbook) as archive:
            entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        append_to_bytes(entries, records)
        return None
    _, backup = export_records(records, workbook, BACKUPS)
    return backup


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        records = load_records(args.input)
        backup = run(records, dry_run=args.dry_run)
    except (InputError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        parser.error(str(exc))
    if args.dry_run:
        print(f"dry-run: validated {len(records)} record(s); workbook unchanged")
    else:
        print(f"appended {len(records)} record(s); backup: {backup}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
