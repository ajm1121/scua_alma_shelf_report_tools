"""Find the deposited files and build both workbooks."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from shelf_report.config import ORDER_THRESHOLD, SEARCH_LOCATIONS
from shelf_report.logic import (
    ISSUE_NOT_IN_ANALYTIC,
    ShelfItem,
    build_shelf_items,
    item_label,
    normalize_barcode,
)
from shelf_report.readers import classify_input, read_csv_rows, read_xlsx_table
from shelf_report.workbooks import prepare_alma_records, write_staff_report, write_student_workbook


class InputError(Exception):
    """The input folder is missing something the report cannot be built without."""


def run(
    input_dir: Path,
    output_dir: Path,
    threshold: int = ORDER_THRESHOLD,
    search_locations: list[str] | None = None,
) -> dict[str, Path]:
    discovered = discover_inputs(input_dir)
    raw_path = _one_file(discovered["raw"], "raw shelf read", input_dir)
    analytic_paths = discovered["analytic"]
    if not analytic_paths:
        raise InputError(
            f"No analytic CSV was found in {input_dir}. Export the shelf read dump and save it there."
        )

    raw_headers, raw_records = read_xlsx_table(raw_path)
    raw_rows = _raw_rows(raw_path, raw_headers, raw_records)
    raw_barcodes = {normalize_barcode(row["barcode"]) for row in raw_rows}

    analytic_path, unused_analytics = _choose_analytic(analytic_paths, raw_barcodes)
    analytic_rows = read_csv_rows(analytic_path)

    missing_records = _load_alma(discovered["missing"])
    not_in_place_records = _load_alma(discovered["not_in_place"])
    out_of_range_records = _load_alma(discovered["out_of_range"])
    unmatched_records = _load_alma(discovered["unmatched"])

    items, warnings = build_shelf_items(
        raw_rows,
        analytic_rows,
        unmatched_barcodes=_barcodes(unmatched_records),
        out_of_range_barcodes=_barcodes(out_of_range_records),
        not_in_place_barcodes=_barcodes(not_in_place_records),
        threshold=threshold,
    )
    for path in unused_analytics:
        warnings.append(f"Left unused because another CSV matched the shelf read: {path.name}")
    for path in discovered["unknown"]:
        warnings.append(f"Not used. Rename it so the script can tell what it is: {path.name}")

    by_barcode = {}
    for item in items:
        by_barcode.setdefault(item.barcode, item)
    missing = prepare_alma_records(missing_records, by_barcode)
    not_in_place = prepare_alma_records(not_in_place_records, by_barcode)
    out_of_range = prepare_alma_records(out_of_range_records, by_barcode)

    project_name = _project_name(raw_path)
    summary_rows, detail_rows = _summary(
        project_name,
        threshold,
        raw_path,
        analytic_path,
        discovered,
        items,
        analytic_rows,
        missing,
        not_in_place,
        out_of_range,
        unmatched_records,
        warnings,
    )

    staff_path = output_dir / f"{project_name}_staff_report.xlsx"
    student_path = output_dir / f"{project_name}_student_workbook.xlsx"
    try:
        write_staff_report(
            staff_path,
            items,
            missing,
            not_in_place,
            out_of_range,
            summary_rows,
            detail_rows,
        )
        write_student_workbook(
            student_path,
            items,
            missing,
            search_locations if search_locations is not None else SEARCH_LOCATIONS,
            threshold,
            project_name,
        )
    except PermissionError as error:
        raise InputError(
            "Could not write the workbook. Close it in Excel if it is open, then run the script again."
        ) from error
    return {"staff": staff_path, "student": student_path, "summary": summary_rows, "details": detail_rows}


def discover_inputs(input_dir: Path) -> dict[str, list[Path]]:
    if not input_dir.is_dir():
        raise InputError(f"Input folder does not exist: {input_dir}")
    groups: dict[str, list[Path]] = defaultdict(list)
    for path in sorted(input_dir.iterdir()):
        if not path.is_file():
            continue
        kind = classify_input(path)
        if kind is None:
            continue
        if kind == "unknown":
            kind = _classify_unknown_workbook(path)
        groups[kind].append(path)
    for kind in ("raw", "analytic", "missing", "not_in_place", "out_of_range", "unmatched", "unknown"):
        groups.setdefault(kind, [])
    return groups


def _classify_unknown_workbook(path: Path) -> str:
    headers, _records = read_xlsx_table(path)
    header_set = {header.casefold() for header in headers}
    if {"index", "barcode"} <= header_set and "call number" not in header_set:
        return "raw"
    return "unknown"


def _one_file(paths: list[Path], label: str, input_dir: Path) -> Path:
    if not paths:
        raise InputError(
            f"No {label} file was found in {input_dir}. "
            "The raw shelf read should have Index and Barcode columns, and 'raw' in the file name."
        )
    if len(paths) > 1:
        names = ", ".join(path.name for path in paths)
        raise InputError(f"Found more than one {label} file: {names}. Keep one shelf read in the folder per run.")
    return paths[0]


def _choose_analytic(paths: list[Path], raw_barcodes: set[str]) -> tuple[Path, list[Path]]:
    if len(paths) == 1:
        return paths[0], []
    scored = []
    for path in paths:
        rows = read_csv_rows(path)
        overlap = len(raw_barcodes & {normalize_barcode(row.get("Barcode")) for row in rows})
        scored.append((overlap, path))
    scored.sort(key=lambda pair: (-pair[0], pair[1].name))
    chosen = scored[0][1]
    return chosen, [path for _overlap, path in scored if path != chosen]


def _raw_rows(path: Path, headers: list[str], records: list[dict[str, str]]) -> list[dict[str, object]]:
    lookup = {header.casefold(): header for header in headers}
    if "barcode" not in lookup:
        raise InputError(f"{path.name} needs a Barcode column.")
    barcode_key = lookup["barcode"]
    index_key = lookup.get("index")
    rows = []
    for position, record in enumerate(records, start=1):
        barcode = record.get(barcode_key, "")
        if not barcode:
            continue
        if index_key and record.get(index_key):
            index = float(record[index_key])
        else:
            index = float(position)
        rows.append({"index": index, "barcode": barcode})
    if not rows:
        raise InputError(f"{path.name} has no barcode rows.")
    return rows


def _load_alma(paths: list[Path]) -> list[dict[str, str]]:
    records = []
    for path in paths:
        _headers, rows = read_xlsx_table(path)
        if rows and "Barcode" not in rows[0] and not any("Barcode" in row for row in rows):
            raise InputError(f"{path.name} looks like an Alma export but has no Barcode column.")
        records.extend(rows)
    return records


def _barcodes(records: list[dict[str, str]]) -> set[str]:
    return {normalize_barcode(record.get("Barcode")) for record in records if record.get("Barcode")}


def _project_name(raw_path: Path) -> str:
    stem = raw_path.stem
    for suffix in ("_raw_report", "_raw", " raw report", " raw"):
        if stem.lower().endswith(suffix):
            stem = stem[: -len(suffix)]
            break
    cleaned = stem.strip(" _-")
    return cleaned or raw_path.stem


def _summary(
    project_name: str,
    threshold: int,
    raw_path: Path,
    analytic_path: Path,
    discovered: dict[str, list[Path]],
    items: list[ShelfItem],
    analytic_rows: list[dict[str, str]],
    missing: list[dict[str, str]],
    not_in_place: list[dict[str, str]],
    out_of_range: list[dict[str, str]],
    unmatched_records: list[dict[str, str]],
    warnings: list[str],
) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    shelf_barcodes = [item.barcode for item in items]
    analytic_barcodes = [normalize_barcode(row.get("Barcode")) for row in analytic_rows if row.get("Barcode")]
    shelf_set = set(shelf_barcodes)
    analytic_set = set(analytic_barcodes)
    missing_from_analytic = sorted(shelf_set - analytic_set - {item.barcode for item in items if item.barcode_not_found})
    extra_analytic = sorted(analytic_set - shelf_set)
    matched = [item for item in items if item.in_analytic]
    order_items = [item for item in items if item.order_check]
    labels = [item_label(item) for item in items if item_label(item)]

    file_bits = [f"Raw: {raw_path.name}", f"Analytic: {analytic_path.name}"]
    for label, key in (
        ("Missing", "missing"),
        ("Not in place", "not_in_place"),
        ("Out of range", "out_of_range"),
        ("No matched barcode", "unmatched"),
    ):
        names = [path.name for path in discovered[key]]
        file_bits.append(f"{label}: {', '.join(names) if names else 'not deposited'}")

    summary = [
        ("Project", project_name),
        ("Order-check threshold", str(threshold)),
        ("Files", " | ".join(file_bits)),
        ("Shelf rows", str(len(items))),
        ("Unique shelf barcodes", str(len(shelf_set))),
        ("Analytic rows", str(len(analytic_rows))),
        ("Matched to the analytic", str(len(matched))),
        ("Barcode not found", str(sum(item.barcode_not_found for item in items))),
        ("Out of range / set", str(len(out_of_range))),
        ("On shelf but marked not in place", str(len(not_in_place))),
        ("Missing items", str(len(missing))),
        ("Possibly out of order", str(len(order_items))),
        ("First call number on the shelf", labels[0] if labels else ""),
        ("Last call number on the shelf", labels[-1] if labels else ""),
    ]

    details: list[tuple[str, str]] = []
    if labels:
        details.append((
            "How the order check was built",
            "Analytic index follows the normalized call number. Within one call number, "
            "descriptions such as v.1, v.2, v.10 are put in numeric order. Other descriptions "
            "stay in the analytic's export order. A row is listed for students when its shelf "
            f"index and analytic index differ by more than {threshold}. Out-of-range items "
            "keep their call numbers, and they do not get a rearrange instruction.",
        ))
    for barcode in missing_from_analytic:
        details.append(("On the shelf, missing from the analytic", barcode))
    for barcode in extra_analytic:
        details.append(("In the analytic, not on this shelf read", barcode))
    for item in items:
        if item.duplicate:
            details.append(("Duplicate barcode on the shelf read", item.barcode))
    for record in missing:
        if record.get("_on_shelf_text") == "Yes":
            details.append(("Missing-item barcode was also on the shelf", record.get("Barcode", "")))
    for record in not_in_place:
        if record.get("_on_shelf_text") != "Yes":
            details.append(("Not-in-place barcode was not on this shelf read", record.get("Barcode", "")))
    for record in out_of_range:
        if record.get("_on_shelf_text") != "Yes":
            details.append(("Out-of-range barcode was not on this shelf read", record.get("Barcode", "")))
    unmatched = _barcodes(unmatched_records)
    for barcode in sorted(unmatched - shelf_set):
        details.append(("Unmatched barcode was not on this shelf read", barcode))
    for item in items:
        if ISSUE_NOT_IN_ANALYTIC in item.issues:
            details.append(("Shelf barcode had no analytic row", f"{item.barcode} at shelf index {item.shelf_index:g}"))
    for warning in warnings:
        details.append(("Warning", warning))
    return summary, details
