"""Read the shelf-read workbook, Alma shelf-report exports, and the analytic CSV."""

from __future__ import annotations

import csv
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

ANALYTIC_REQUIRED = (
    "Normalized Call Number",
    "Permanent Call Number",
    "Barcode",
)


def classify_input(path: Path) -> str | None:
    """Return the role of an input file, or None if it should be ignored."""
    name = path.name.lower()
    if name.startswith("~$") or name.startswith("."):
        return None
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return "analytic"
    if suffix not in {".xlsx", ".xlsm"}:
        return None
    if "missing" in name:
        return "missing"
    if "not in place" in name or "not_in_place" in name:
        return "not_in_place"
    if "out of range" in name or "out_of_range" in name or "out of set" in name:
        return "out_of_range"
    unmatched_tokens = (
        "no matched",
        "no_matched",
        "unmatched",
        "not found",
        "not_found",
        "not in db",
        "not_in_db",
        "no item",
        "no_item",
    )
    if any(token in name for token in unmatched_tokens):
        return "unmatched"
    if "raw" in name or "shelf read" in name or "shelf_read" in name:
        return "raw"
    return "unknown"


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None:
            return []
        missing = [name for name in ANALYTIC_REQUIRED if name not in reader.fieldnames]
        if missing:
            joined = ", ".join(missing)
            raise ValueError(f"{path.name} is missing analytic columns: {joined}")
        rows = []
        for record in reader:
            cleaned = {key: (value or "").strip() for key, value in record.items() if key}
            if any(cleaned.values()):
                rows.append(cleaned)
        return rows


def read_xlsx_table(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    """Read the first worksheet as headers plus header-keyed rows.

    Alma's shelf-report Excel files store every value as an inline string and
    leave the worksheet dimension as A1, which makes openpyxl drop the rows.
    This reader uses the worksheet XML directly and also accepts ordinary
    workbooks that use shared strings.
    """
    with zipfile.ZipFile(path) as workbook:
        shared = _shared_strings(workbook)
        sheet_path = _first_sheet_path(workbook)
        root = ET.fromstring(workbook.read(sheet_path))

    rows: dict[int, dict[int, str]] = {}
    for cell in root.iter(f"{MAIN_NS}c"):
        ref = cell.attrib.get("r")
        if not ref:
            continue
        column, row_number = _column_row(ref)
        text = _cell_text(cell, shared).strip()
        if text:
            rows.setdefault(row_number, {})[column] = text
    if not rows:
        return [], []

    header_number = min(rows)
    header_names: dict[int, str] = {}
    used: set[str] = set()
    for column, name in sorted(rows[header_number].items()):
        base = name.strip() or f"column_{column}"
        unique = base
        while unique in used:
            unique = f"{base}_{column}"
        used.add(unique)
        header_names[column] = unique

    records = []
    for row_number in sorted(rows):
        if row_number == header_number:
            continue
        record = {
            header_names[column]: value
            for column, value in rows[row_number].items()
            if column in header_names
        }
        if record:
            records.append(record)
    headers = [header_names[column] for column in sorted(header_names)]
    return headers, records


def _shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(f"{MAIN_NS}t"))
        for item in root.findall(f"{MAIN_NS}si")
    ]


def _first_sheet_path(workbook: zipfile.ZipFile) -> str:
    workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
    first_sheet = workbook_xml.find(f"{MAIN_NS}sheets/{MAIN_NS}sheet")
    if first_sheet is None:
        return "xl/worksheets/sheet1.xml"
    relation_id = first_sheet.attrib.get(f"{REL_NS}id")
    relationships = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    for relationship in relationships:
        if relationship.attrib.get("Id") != relation_id:
            continue
        target = relationship.attrib.get("Target", "worksheets/sheet1.xml")
        if target.startswith("/"):
            return target.lstrip("/")
        return "xl/" + target.lstrip("/")
    return "xl/worksheets/sheet1.xml"


def _column_row(cell_ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", cell_ref)
    if match is None:
        raise ValueError(f"Unrecognized cell reference {cell_ref}")
    column = 0
    for character in match.group(1):
        column = column * 26 + (ord(character) - 64)
    return column, int(match.group(2))


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{MAIN_NS}t"))
    value = cell.find(f"{MAIN_NS}v")
    if cell_type == "s":
        if value is None or value.text is None:
            return ""
        return shared[int(value.text)]
    if value is None or value.text is None:
        return ""
    return value.text
