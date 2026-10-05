# SCUA Alma shelf report tools

Builds a staff shelf-read report and a student workbook from one Alma shelf report run.

## What to put in `input`

One shelf read per run:

| File | Required | How the script recognizes it |
| --- | --- | --- |
| Raw shelf read, with Index and Barcode columns | Yes | `.xlsx` or `.xlsm` with `raw` in the file name |
| Shelf read dump analytic | Yes | the `.csv` export |
| Missing items | No | file name contains `missing` |
| Items on shelf but marked not in place | No | file name contains `not in place` |
| Items on shelf but out of range/set | No | file name contains `out of range` |
| Barcodes with no matched item | No | file name contains `no matched`, `unmatched`, or `not found` |

Excel lock files (`~$...`) are ignored. If Alma does not produce an unmatched-barcode list, leave that file out.

## Run

```
python process_shelf_report.py
```

Optional flags:

```
python process_shelf_report.py --input input --output output --threshold 20
```

`--threshold` is how far a shelf position can drift from call-number order before a student is asked to move it. The written procedure uses 20. The older template formula uses 10.

Results land in `output`:

- `<project>_staff_report.xlsx` — shelf read in walking order, plus missing, not-in-place, out-of-range, and a summary
- `<project>_student_workbook.xlsx` — search for missing items, pull unmatched barcodes for Austin, and review possible misshelving

The places students mark while searching for missing items are in `shelf_report/config.py`.

## What the script decides

- Unmatched barcodes are noted `Barcode not found`, given a half index, and sent to the cart sheet. Later shelf positions move up so they stay comparable to call-number order.
- Out-of-range items are noted and kept on the staff report. They do not get a rearrange instruction.
- Items on the shelf but marked not in place stay on the staff report, with the shelf position filled in. Their Alma status is not a student task.
- Call-number order follows Alma's normalized call number. Descriptions like `v.1`, `v.2`, `v.10`, and `Bd.1` are ordered numerically inside that call number. Other descriptions stay in the analytic's export order.
- The first and last items use "Before" and "After" instead of the old formula, which quoted the column headings.

The order check is still approximate. Read the out-of-order sheet before giving the workbook to students.
