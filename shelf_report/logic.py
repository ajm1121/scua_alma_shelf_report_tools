"""Turn a shelf read plus Alma exports into located, order-checked items."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timedelta

NOTE_NOT_FOUND = "Barcode not found"
NOTE_OUT_OF_RANGE = "Item on shelf but out of range/set"
ISSUE_DUPLICATE = "Duplicate barcode on shelf read"
ISSUE_NOT_IN_ANALYTIC = "Not in shelf read analytic"
ISSUE_NOT_IN_PLACE = "Marked not in place"

# Volume, Bd., and similar enumerations. "no. 193" is left alone so newspaper
# descriptions stay in the order Analytics exported.
VOLUME_DESCRIPTION = re.compile(r"(?i)(?:^|\s)(?:v|vol|bd|band)\.?\s*\d")


@dataclass
class ShelfItem:
    original_index: float
    shelf_index: float
    barcode: str
    normalized_call_number: str = ""
    permanent_call_number: str = ""
    description: str = ""
    copy_id: str = ""
    title: str = ""
    author: str = ""
    location_name: str = ""
    analytic_barcode: str = ""
    analytic_row: int = 0
    analytic_index: int | None = None
    current_location: str = ""
    order_check: str = ""
    notes: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    in_analytic: bool = False
    barcode_not_found: bool = False
    out_of_range: bool = False
    not_in_place: bool = False
    duplicate: bool = False

    @property
    def notes_text(self) -> str:
        return "; ".join(self.notes)

    @property
    def issues_text(self) -> str:
        return "; ".join(self.issues)


def normalize_barcode(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def natural_key(text: str) -> tuple:
    """Sort v.2 before v.10 without changing genuine text."""
    parts = re.split(r"(\d+)", (text or "").strip().casefold())
    key = []
    for part in parts:
        if not part:
            continue
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part))
    return tuple(key)


def call_number_sort_key(call_number: str) -> tuple:
    text = (call_number or "").strip()
    match = re.match(r"^([A-Za-z]+)\s*(\d+(?:\.\d+)?)?(.*)$", text)
    if match is None:
        return (text.casefold(), 0.0, natural_key(""))
    letters, number, rest = match.groups()
    return (
        letters.upper(),
        float(number) if number else -1.0,
        natural_key(rest or ""),
    )


def parse_creator(type_creator: str) -> str:
    """Pull a creator out of Alma's combined Type / Creator / Imprint column."""
    match = re.search(r"\bBy\s+(.+?)(?:\s*\(|\.\s*$)", type_creator or "")
    if match is None:
        return ""
    return match.group(1).strip().rstrip("., ")


def excel_serial_to_text(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    try:
        number = float(text)
    except ValueError:
        return text
    if number < 1 or number > 100000:
        return text
    moment = datetime(1899, 12, 30) + timedelta(days=number)
    if abs(number - round(number)) < 1e-8:
        return moment.strftime("%Y-%m-%d")
    return moment.strftime("%Y-%m-%d %H:%M")


def item_label(item: ShelfItem) -> str:
    call_number = item.permanent_call_number.strip()
    description = item.description.strip()
    if call_number and description:
        return f"{call_number} {description}"
    return call_number or description


def build_shelf_items(
    raw_rows: list[dict[str, object]],
    analytic_rows: list[dict[str, str]],
    unmatched_barcodes: set[str],
    out_of_range_barcodes: set[str],
    not_in_place_barcodes: set[str],
    threshold: int,
) -> tuple[list[ShelfItem], list[str]]:
    """Return shelf items in shelf order, plus warnings about the analytic."""
    warnings: list[str] = []
    items: list[ShelfItem] = []
    for row in raw_rows:
        barcode = normalize_barcode(row.get("barcode"))
        if not barcode:
            continue
        index = float(row.get("index") or 0)
        items.append(ShelfItem(original_index=index, shelf_index=index, barcode=barcode))

    counts = Counter(item.barcode for item in items)
    for item in items:
        if counts[item.barcode] > 1:
            item.duplicate = True

    analytic_by_barcode: dict[str, tuple[int, dict[str, str]]] = {}
    for row_number, row in enumerate(analytic_rows, start=1):
        barcode = normalize_barcode(row.get("Barcode"))
        if not barcode:
            continue
        if barcode in analytic_by_barcode:
            warnings.append(f"Analytic barcode {barcode} appears more than once; the first row was used.")
            continue
        analytic_by_barcode[barcode] = (row_number, row)

    matched_once: set[str] = set()
    for item in items:
        if item.barcode in unmatched_barcodes:
            item.barcode_not_found = True
            if item.barcode in analytic_by_barcode:
                warnings.append(
                    f"Barcode {item.barcode} is on the unmatched list and also in the analytic; "
                    "it was treated as barcode not found."
                )
        elif item.barcode in analytic_by_barcode:
            row_number, row = analytic_by_barcode[item.barcode]
            _apply_analytic_row(item, row, row_number)
            if item.barcode in matched_once:
                item.in_analytic = False
            else:
                item.in_analytic = True
                matched_once.add(item.barcode)
        item.out_of_range = item.barcode in out_of_range_barcodes
        item.not_in_place = item.barcode in not_in_place_barcodes

    _adjust_shelf_indexes(items)
    _assign_analytic_indexes(items)
    _assign_current_locations(items)
    _assign_order_checks(items, threshold)
    _compose_messages(items)
    items.sort(key=lambda item: (item.shelf_index, item.barcode))
    return items, warnings


def _apply_analytic_row(item: ShelfItem, row: dict[str, str], row_number: int) -> None:
    item.analytic_row = row_number
    item.analytic_barcode = normalize_barcode(row.get("Barcode"))
    item.normalized_call_number = (row.get("Normalized Call Number") or "").strip()
    item.permanent_call_number = (row.get("Permanent Call Number") or "").strip()
    item.description = (row.get("Description") or "").strip()
    item.copy_id = (row.get("Item Copy Id") or "").strip()
    item.title = (row.get("Title") or "").strip()
    item.author = (row.get("Author") or "").strip()
    item.location_name = (row.get("Location Name") or "").strip()


def _adjust_shelf_indexes(items: list[ShelfItem]) -> None:
    """File unmatched barcodes between their neighbors, matching the template macro.

    A barcode Alma cannot match is given a half index, and every later shelf
    position moves up by one. Matched items then stay on the same 1..n scale
    as the analytic index.
    """
    items.sort(key=lambda item: (item.original_index, item.barcode))
    skipped = 0
    for item in items:
        if item.barcode_not_found:
            item.shelf_index = item.original_index - skipped - 0.5
            skipped += 1
        else:
            item.shelf_index = item.original_index - skipped


def _assign_analytic_indexes(items: list[ShelfItem]) -> None:
    matched = [item for item in items if item.in_analytic]
    volume_groups = {
        item.normalized_call_number
        for item in matched
        if VOLUME_DESCRIPTION.search(item.description or "")
    }
    matched.sort(
        key=lambda item: (
            item.normalized_call_number,
            natural_key(item.description) if item.normalized_call_number in volume_groups else (),
            0 if item.normalized_call_number in volume_groups else item.analytic_row,
            _copy_sort(item.copy_id),
            item.barcode,
        )
    )
    for index, item in enumerate(matched, start=1):
        item.analytic_index = index


def _copy_sort(copy_id: str) -> tuple:
    text = (copy_id or "").strip()
    if text.isdigit():
        return (0, int(text))
    return (1, text)


def _assign_current_locations(items: list[ShelfItem]) -> None:
    ordered = sorted(items, key=lambda item: (item.shelf_index, item.barcode))
    labels = [item_label(item) for item in ordered]
    for index, item in enumerate(ordered):
        previous = next((labels[pos] for pos in range(index - 1, -1, -1) if labels[pos]), "")
        following = next((labels[pos] for pos in range(index + 1, len(labels)) if labels[pos]), "")
        item.current_location = _between(previous, following)


def _assign_order_checks(items: list[ShelfItem], threshold: int) -> None:
    ordered = sorted(
        (item for item in items if item.analytic_index is not None),
        key=lambda item: item.analytic_index or 0,
    )
    for index, item in enumerate(ordered):
        item.order_check = ""
        if item.out_of_range or item.barcode_not_found:
            continue
        if abs(item.shelf_index - (item.analytic_index or 0)) <= threshold:
            continue
        previous = _destination_neighbor(ordered, item, index, -1, threshold)
        following = _destination_neighbor(ordered, item, index, 1, threshold)
        place = _between(previous, following)
        if place:
            item.order_check = f"Rearrange, should go {place[:1].lower()}{place[1:]}"


def _destination_neighbor(ordered: list[ShelfItem], item: ShelfItem, index: int, step: int, threshold: int) -> str:
    """Name a nearby item that is already in call-number order.

    Other misplaced items are skipped, so a volume is not told to move
    next to the volume that is sitting beside it in the wrong place.
    """
    pos = index + step
    while 0 <= pos < len(ordered):
        neighbor = ordered[pos]
        if _settled(neighbor, threshold) and item_label(neighbor):
            return _neighbor_label(neighbor, item)
        pos += step
    return ""


def _settled(item: ShelfItem, threshold: int) -> bool:
    if item.analytic_index is None or item.out_of_range or item.barcode_not_found:
        return False
    return abs(item.shelf_index - item.analytic_index) <= threshold


def _neighbor_label(neighbor: ShelfItem, item: ShelfItem) -> str:
    label = item_label(neighbor)
    if label != item_label(item):
        return label
    neighbor_copy = neighbor.copy_id.strip()
    if neighbor_copy and neighbor_copy != item.copy_id.strip():
        return f"{label} (copy {neighbor_copy})"
    if neighbor.current_location:
        location = neighbor.current_location
        return f"{label} (the one {location[:1].lower()}{location[1:]})"
    return label


def _between(previous: str, following: str) -> str:
    if previous and following:
        return f"Between {previous} and {following}"
    if following:
        return f"Before {following}"
    if previous:
        return f"After {previous}"
    return ""


def _compose_messages(items: list[ShelfItem]) -> None:
    for item in items:
        notes: list[str] = []
        issues: list[str] = []
        if item.barcode_not_found:
            notes.append(NOTE_NOT_FOUND)
        if item.out_of_range:
            notes.append(NOTE_OUT_OF_RANGE)
        if item.duplicate:
            issues.append(ISSUE_DUPLICATE)
        if not item.in_analytic and not item.barcode_not_found and not item.analytic_barcode:
            issues.append(ISSUE_NOT_IN_ANALYTIC)
        if item.not_in_place:
            issues.append(ISSUE_NOT_IN_PLACE)
        item.notes = notes
        item.issues = issues
