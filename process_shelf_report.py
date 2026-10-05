"""Build the staff shelf report and the student workbook.

Drop the raw shelf read, the Alma shelf-report lists, and the analytic export
into the input folder, then run:

    python process_shelf_report.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from shelf_report.config import ORDER_THRESHOLD
from shelf_report.pipeline import InputError, run


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the staff shelf report and the student workbook.")
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("input"),
        help="Folder containing the raw shelf read, Alma lists, and analytic CSV. Defaults to input.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output"),
        help="Folder for the two workbooks. Defaults to output.",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=ORDER_THRESHOLD,
        help=f"Flag an item when shelf position and call-number position differ by more than this. Defaults to {ORDER_THRESHOLD}.",
    )
    args = parser.parse_args(argv)
    if args.threshold < 0:
        parser.error("--threshold must be zero or greater.")
    try:
        result = run(args.input, args.output, threshold=args.threshold)
    except InputError as error:
        print(error, file=sys.stderr)
        return 1

    print(f"Staff report: {result['staff']}")
    print(f"Student workbook: {result['student']}")
    for label, value in result["summary"]:
        print(f"{label}: {value}")
    if result["details"]:
        print("Review:")
        for label, value in result["details"]:
            print(f"- {label}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
