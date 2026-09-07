import csv
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from consent_parser import COLUMNS, OUTPUT_COLUMNS, IRB, load_irbs, parse_irb, parse_consent_files, write_csv


class ConsentTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def irb(self, name, rows, accepted=None):
        path = self.directory / f"{name}.csv"
        with path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerow(COLUMNS)
            writer.writerows(rows)
        return IRB(name, path, accepted if accepted is not None else ["Both"])

    def row(self, identifier="js1", consent="Both", **changes):
        row = dict(zip(COLUMNS, ["Jane", "Smith", "j@example.org", identifier, "001", "Dr. Lee", consent]))
        row.update(changes)
        return [row[column] for column in COLUMNS]

    def test_multiple_irbs_and_explicit_consent(self):
        a = self.irb("A", [self.row(), self.row("no", "No")])
        b = self.irb("B", [self.row(" JS1 ", " Survey   Only "), self.row("blank", "")], ["Survey Only"])
        c = self.irb("C", [self.row("other", "Yes"), self.row("excluded", "Both")], ["Yes"])
        rows = parse_consent_files([a, b, c])
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["IRBs"], "A; B")
        self.assertEqual(rows[1]["IRBs"], "C")
        self.assertEqual(rows[0]["Subject ID"], "001")

    def test_only_computing_id_joins(self):
        irb = self.irb("A", [self.row(), self.row(), self.row("other")])
        self.assertEqual(len(parse_consent_files([irb])), 2)

    def test_missing_id(self):
        with self.assertRaisesRegex(ValueError, "no Computing ID"):
            parse_irb(self.irb("A", [self.row("")]))
        self.assertEqual(parse_irb(self.irb("B", [self.row("", "No")])), [])

    def test_only_consenting_rows_fill_blanks(self):
        a = self.irb("A", [self.row(**{"Professor": ""}), self.row(consent="No", Professor="Wrong")])
        b = self.irb("B", [self.row()])
        self.assertEqual(parse_consent_files([a, b])[0]["Professor"], "Dr. Lee")

    def test_conflicts(self):
        for column in ("First name", "Last name", "email", "Subject ID", "Professor"):
            with self.subTest(column=column), self.assertRaisesRegex(ValueError, f"conflicting {column}"):
                parse_consent_files([self.irb("A", [self.row(), self.row(**{column: "Other"})])])

    def test_configuration_and_csv_export(self):
        self.irb("A", [self.row()])
        config = self.directory / "irbs.json"
        config.write_text(json.dumps([{"name": "A", "file": "A.csv", "consent_values": ["Both"]}]))
        rows = parse_consent_files(load_irbs(config))
        output = self.directory / "output.csv"
        write_csv(rows, output)
        with output.open(encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            self.assertEqual(reader.fieldnames, list(OUTPUT_COLUMNS))
            self.assertEqual(list(reader), rows)
        with self.assertRaises(FileExistsError):
            write_csv(rows, output)

    def test_invalid_rules_and_names(self):
        for values in ([], [""], ["  "], "Both", [None]):
            with self.subTest(values=values), self.assertRaises(ValueError):
                parse_irb(self.irb("A", [], values))
        irb = self.irb("A", [])
        with self.assertRaisesRegex(ValueError, "unique"):
            parse_consent_files([irb, irb])
        with self.assertRaises(ValueError):
            parse_consent_files([IRB("", irb.file, ["Both"])])

    def test_headers(self):
        irb = self.irb("A", [])
        for content in ("", "First name,email\n", "email,EMAIL\n"):
            Path(irb.file).write_text(content)
            with self.assertRaises(ValueError):
                parse_irb(irb)
        with Path(irb.file).open("w", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow([" " + column.upper() + " " for column in reversed(COLUMNS)])
            writer.writerow(list(reversed(self.row())))
        self.assertEqual(parse_irb(irb)[0]["Computing ID"], "js1")

    def test_excel_sheet_and_zero_padding(self):
        from openpyxl import Workbook
        path = self.directory / "studies.xlsx"
        book = Workbook()
        for name in ("A", "B"):
            sheet = book.create_sheet(name)
            sheet.append(["Title"])
            sheet.append(COLUMNS)
            sheet.append(self.row(**{"Subject ID": 12}))
            sheet["E3"].number_format = "00000"
        book.save(path)
        book.close()
        rows = parse_consent_files([IRB(name, path, ["Both"], name, 2) for name in ("A", "B")])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Subject ID"], "00012")
        self.assertEqual(rows[0]["IRBs"], "A; B")
        with self.assertRaises(ValueError):
            parse_irb(IRB("Missing", path, ["Both"], "Missing"))


if __name__ == "__main__":
    unittest.main()
