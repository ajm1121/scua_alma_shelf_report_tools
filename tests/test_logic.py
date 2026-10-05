import unittest
from pathlib import Path

from shelf_report.logic import (
    ISSUE_NOT_IN_PLACE,
    NOTE_NOT_FOUND,
    NOTE_OUT_OF_RANGE,
    build_shelf_items,
    natural_key,
    parse_creator,
)
from shelf_report.readers import classify_input, read_xlsx_table


def analytic(barcode, normalized, call_number, description="", title=""):
    return {
        "Normalized Call Number": normalized,
        "Permanent Call Number": call_number,
        "Description": description,
        "Item Copy Id": "1",
        "Barcode": barcode,
        "Title": title,
        "Author": "",
        "Location Name": "Spec Coll",
    }


class LogicTests(unittest.TestCase):
    def test_volume_descriptions_sort_numerically(self):
        self.assertLess(natural_key("v.2"), natural_key("v.10"))
        self.assertLess(natural_key("Bd.2"), natural_key("Bd.10"))
        self.assertLess(natural_key("t.2 1667-71"), natural_key("t.10 1682"))
        self.assertLess(natural_key("t.11 1683"), natural_key("t.100 1733"))

    def test_tome_numbers_sort_inside_a_call_number(self):
        raw = [
            {"index": 1, "barcode": "100"},
            {"index": 2, "barcode": "2"},
            {"index": 3, "barcode": "11"},
        ]
        rows = [
            analytic("100", "0ap!25 j7", "AP25 .J7", "t.10 1682"),
            analytic("2", "0ap!25 j7", "AP25 .J7", "t.2 1667-71"),
            analytic("11", "0ap!25 j7", "AP25 .J7", "t.11 1683"),
        ]
        items, _warnings = build_shelf_items(raw, rows, set(), set(), set(), threshold=20)
        ordered = [item.description for item in sorted(items, key=lambda item: item.analytic_index or 0)]
        self.assertEqual(ordered, ["t.2 1667-71", "t.10 1682", "t.11 1683"])

    def test_newspaper_descriptions_keep_analytic_order(self):
        raw = [
            {"index": 1, "barcode": "111"},
            {"index": 2, "barcode": "222"},
        ]
        rows = [
            analytic("111", "f", "F884 .E9", "1829/1929 cent. ed."),
            analytic("222", "f", "F884 .E9", "82nd no.193 1948"),
        ]
        items, _warnings = build_shelf_items(raw, rows, set(), set(), set(), threshold=20)
        by_barcode = {item.barcode: item for item in items}
        self.assertLess(by_barcode["111"].analytic_index, by_barcode["222"].analytic_index)

    def test_number_and_issue_sort_numerically_when_they_lead(self):
        raw = [
            {"index": 1, "barcode": "10"},
            {"index": 2, "barcode": "2"},
            {"index": 3, "barcode": "i10"},
            {"index": 4, "barcode": "i2"},
        ]
        rows = [
            analytic("10", "a", "AC1 .C62", "no.10"),
            analytic("2", "a", "AC1 .C62", "no.2"),
            analytic("i10", "b", "PS648 .S5", "Issue 10 1991"),
            analytic("i2", "b", "PS648 .S5", "Issue 2 1988"),
        ]
        items, _warnings = build_shelf_items(raw, rows, set(), set(), set(), threshold=20)
        by_barcode = {item.barcode: item for item in items}
        self.assertLess(by_barcode["2"].analytic_index, by_barcode["10"].analytic_index)
        self.assertLess(by_barcode["i2"].analytic_index, by_barcode["i10"].analytic_index)

    def test_unmatched_barcode_shifts_later_indexes(self):
        raw = [
            {"index": 1, "barcode": "111"},
            {"index": 2, "barcode": "222"},
            {"index": 3, "barcode": "333"},
        ]
        rows = [
            analytic("111", "a", "A 1"),
            analytic("333", "b", "B 1"),
        ]
        items, _warnings = build_shelf_items(raw, rows, {"222"}, set(), set(), threshold=20)
        by_barcode = {item.barcode: item for item in items}
        self.assertEqual(by_barcode["222"].shelf_index, 1.5)
        self.assertEqual(by_barcode["222"].notes_text, NOTE_NOT_FOUND)
        self.assertIsNone(by_barcode["222"].analytic_index)
        self.assertEqual(by_barcode["111"].shelf_index, 1)
        self.assertEqual(by_barcode["333"].shelf_index, 2)
        self.assertEqual(by_barcode["333"].analytic_index, 2)
        self.assertEqual(by_barcode["222"].current_location, "Between A 1 and B 1")

    def test_consecutive_unmatched_barcodes_keep_separate_indexes(self):
        raw = [
            {"index": 1, "barcode": "111"},
            {"index": 2, "barcode": "222"},
            {"index": 3, "barcode": "333"},
            {"index": 4, "barcode": "444"},
        ]
        rows = [
            analytic("111", "a", "A 1"),
            analytic("444", "b", "B 1"),
        ]
        items, _warnings = build_shelf_items(raw, rows, {"222", "333"}, set(), set(), threshold=20)
        by_barcode = {item.barcode: item for item in items}
        self.assertEqual(by_barcode["111"].shelf_index, 1)
        self.assertAlmostEqual(by_barcode["222"].shelf_index, 1 + 1 / 3)
        self.assertAlmostEqual(by_barcode["333"].shelf_index, 1 + 2 / 3)
        self.assertEqual(by_barcode["444"].shelf_index, 2)

    def test_volume_order_and_out_of_range_suppression(self):
        raw = [
            {"index": 1, "barcode": "111"},
            {"index": 2, "barcode": "444"},
            {"index": 3, "barcode": "555"},
            {"index": 4, "barcode": "666"},
            {"index": 5, "barcode": "777"},
        ]
        rows = [
            analytic("111", "z", "Z 1"),
            analytic("444", "m", "M", "v.10"),
            analytic("555", "m", "M", "v.2"),
            analytic("666", "m", "M", "v.1"),
            analytic("777", "b", "B 1"),
        ]
        items, _warnings = build_shelf_items(
            raw,
            rows,
            unmatched_barcodes=set(),
            out_of_range_barcodes={"777"},
            not_in_place_barcodes={"111"},
            threshold=0,
        )
        by_barcode = {item.barcode: item for item in items}
        self.assertEqual(
            [by_barcode[barcode].analytic_index for barcode in ("666", "555", "444")],
            [2, 3, 4],
        )
        self.assertEqual(by_barcode["777"].notes_text, NOTE_OUT_OF_RANGE)
        self.assertEqual(by_barcode["777"].order_check, "")
        self.assertEqual(by_barcode["111"].issues_text, ISSUE_NOT_IN_PLACE)
        self.assertIn("Rearrange, should go", by_barcode["666"].order_check)
        self.assertNotIn("B 1", by_barcode["666"].order_check)

    def test_destination_skips_other_misplaced_items(self):
        barcodes = ["A", "B", "C", "D", "E", "F", "X", "Y"]
        raw = [{"index": index, "barcode": barcode} for index, barcode in enumerate(barcodes, start=1)]
        rows = [
            analytic("A", "1", "A 1"),
            analytic("X", "2", "X 1"),
            analytic("Y", "3", "Y 1"),
            analytic("B", "4", "B 1"),
            analytic("C", "5", "C 1"),
            analytic("D", "6", "D 1"),
            analytic("E", "7", "E 1"),
            analytic("F", "8", "F 1"),
        ]
        items, _warnings = build_shelf_items(raw, rows, set(), set(), set(), threshold=3)
        by_barcode = {item.barcode: item for item in items}
        for barcode in ("X", "Y"):
            advice = by_barcode[barcode].order_check
            self.assertIn("A 1", advice)
            self.assertIn("B 1", advice)
            self.assertNotIn("X 1", advice)
            self.assertNotIn("Y 1", advice)
        self.assertEqual(by_barcode["B"].order_check, "")

    def test_same_call_number_is_distinguished_by_copy(self):
        raw = [
            {"index": 1, "barcode": "222"},
            {"index": 2, "barcode": "333"},
            {"index": 3, "barcode": "444"},
            {"index": 4, "barcode": "555"},
            {"index": 5, "barcode": "111"},
        ]
        rows = [
            analytic("222", "a", "ND1042 .C69"),
            analytic("333", "b", "ND1043 .B56"),
            analytic("444", "c", "ND1044 .C1"),
            analytic("555", "d", "ND1045 .D1"),
            analytic("111", "a", "ND1042 .C69"),
        ]
        rows[0]["Item Copy Id"] = "1"
        rows[4]["Item Copy Id"] = "2"
        items, _warnings = build_shelf_items(raw, rows, set(), set(), set(), threshold=1)
        misplaced = next(item for item in items if item.barcode == "111")
        self.assertIn("(copy 1)", misplaced.order_check)
        self.assertIn("ND1043 .B56", misplaced.order_check)

    def test_creator_parse(self):
        text = (
            "Book {Book - Physical} By Keough, Pat, "
            "([Salt String Island, British Columbia, Canada] : Nahanni Productions Inc., [2002])"
        )
        self.assertEqual(parse_creator(text), "Keough, Pat")
        self.assertEqual(parse_creator("Book {Book - Physical}"), "")

    def test_file_roles(self):
        self.assertEqual(classify_input(Path("double_os_raw_report.xlsx")), "raw")
        self.assertEqual(classify_input(Path("missing_items.xlsx")), "missing")
        self.assertEqual(classify_input(Path("Items on shelf but marked not in place.xlsx")), "not_in_place")
        self.assertEqual(classify_input(Path("Items on shelf but out of range.xlsx")), "out_of_range")
        self.assertEqual(classify_input(Path("Barcodes with no matched item.xlsx")), "unmatched")
        self.assertEqual(
            classify_input(Path("Barcodes not in DB list__45940664060001852.xlsx")),
            "unmatched",
        )
        self.assertEqual(classify_input(Path("SCUA_Shelf Report Dump.csv")), "analytic")
        self.assertIsNone(classify_input(Path("~$double_os_raw_report.xlsx")))

    def test_real_alma_export_keeps_its_rows(self):
        path = Path("input/missing_items.xlsx")
        if not path.exists():
            self.skipTest("sample Alma export is not in input")
        headers, rows = read_xlsx_table(path)
        self.assertIn("Barcode", headers)
        self.assertIn("Call Number", headers)
        self.assertGreater(len(rows), 0)
        self.assertTrue(rows[0]["Barcode"])


if __name__ == "__main__":
    unittest.main()
