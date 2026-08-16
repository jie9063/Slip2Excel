import datetime as dt
import queue
import sys
import tempfile
import threading
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bookkeeping import (  # noqa: E402
    BookkeepingApp,
    Item,
    LocalVisionReader,
    NS_MAIN,
    Receipt,
    XlsxWriter,
    parse_date,
    prepare_entries,
    read_photos_worker,
)


def make_template(path: Path) -> None:
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        "_rels/.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="客戶A" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="1"/></workbook>""",
        "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="4"><c r="K4" s="1"><f>I4*J4</f><v>0</v></c></row></sheetData></worksheet>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


class BookkeepingLogicTests(unittest.TestCase):
    def test_parse_dates(self):
        self.assertEqual(parse_date("115/8/8"), dt.date(2026, 8, 8))
        self.assertEqual(parse_date("8/8/2026"), dt.date(2026, 8, 8))
        self.assertEqual(parse_date("10:30 8 14"), dt.date(dt.date.today().year, 8, 14))
        with self.assertRaises(ValueError):
            parse_date("8/8/26")

    def test_number_and_duplicate_line_validation(self):
        self.assertEqual(LocalVisionReader._number("3斤半"), 3.5)
        self.assertEqual(LocalVisionReader._number("半斤"), 0.5)
        for value in (0, -1, "-5", float("nan"), True):
            with self.assertRaises(ValueError):
                LocalVisionReader._number(value)
        receipt = LocalVisionReader.from_json(
            '{"customer":"客戶A","date":"8/8","items":['
            '{"line":1,"name":"菜A","quantity":1,"price":10},'
            '{"line":1,"name":"菜B","quantity":2,"price":20}]}'
        )
        self.assertEqual(len(receipt.items), 1)
        self.assertEqual(len(receipt.warnings), 1)
        with self.assertRaises(ValueError):
            LocalVisionReader.from_json("[]")
        malformed = LocalVisionReader.from_json(
            '{"customer":"客戶A","date":"8/8","items":["bad",'
            '{"line":1.5,"name":"菜A","quantity":1,"price":10}]}'
        )
        self.assertTrue(malformed.problems)
        self.assertEqual(len(malformed.warnings), 2)

    def test_prepare_entries_deduplicates_and_rejects_conflicts(self):
        first = Receipt("a.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        same = Receipt("b.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        entries, duplicates = prepare_entries([first, same])
        self.assertEqual(len(entries), 1)
        self.assertEqual(duplicates, 1)
        conflict = Receipt("c.jpg", "客戶A", "8/8", [Item("", 3, 70, 1)], [])
        with self.assertRaises(ValueError):
            prepare_entries([first, conflict])

    def test_issue_list_only_shows_failed_or_incomplete_receipts(self):
        class FakeTree:
            def __init__(self):
                self.rows, self.selected = {}, []

            def get_children(self):
                return list(self.rows)

            def delete(self, key):
                self.rows.pop(key, None)

            def insert(self, _parent, _position, iid, values, tags):
                self.rows[iid] = (values, tags)

            def tag_configure(self, *_args, **_kwargs):
                pass

            def selection(self):
                return self.selected

        class FakeLabel:
            def configure(self, **kwargs):
                self.text = kwargs["text"]

        class FakeText:
            def delete(self, *_args):
                self.value = ""

            def insert(self, *_args):
                self.value = _args[-1]

        app = object.__new__(BookkeepingApp)
        app.issues_tree, app.issue_summary, app.issue_detail = FakeTree(), FakeLabel(), FakeText()
        app.receipts = [
            Receipt("failed.jpg", "", "", [], ["missing date"]),
            Receipt("partial.jpg", "客戶A", "8/8", [Item("", 1, 2, 1)], [], ["missing price"]),
            Receipt("ok.jpg", "客戶A", "8/8", [Item("", 1, 2, 1)], []),
        ]
        app.refresh_issues()
        self.assertEqual(set(app.issues_tree.rows), {"0", "1"})
        app.issues_tree.selected = ["1"]
        app.show_issue_detail()
        self.assertIn("partial.jpg", app.issue_detail.value)
        self.assertIn("missing price", app.issue_detail.value)

    def test_background_worker_reports_receipts_without_tkinter(self):
        class FakeReader:
            def __init__(self, model):
                self.model = model

            def check_ready(self):
                pass

            def read(self, photo):
                return Receipt(photo.name, "客戶A", "8/8", [Item("", 1, 2, 1)], [])

        events = queue.Queue()
        with patch("bookkeeping.LocalVisionReader", FakeReader):
            read_photos_worker([Path("one.jpg"), Path("two.jpg")], "model", events, threading.Event())
        received = [events.get_nowait() for _ in range(events.qsize())]
        self.assertEqual([event[0] for event in received], ["receipt", "receipt", "finished"])
        self.assertEqual(received[0][3].file, "one.jpg")

    def test_background_worker_honours_cancel_before_next_photo(self):
        class FakeReader:
            def __init__(self, _model):
                pass

            def check_ready(self):
                pass

            def read(self, _photo):
                raise AssertionError("cancelled work must not read a photo")

        events, cancel = queue.Queue(), threading.Event()
        cancel.set()
        with patch("bookkeeping.LocalVisionReader", FakeReader):
            read_photos_worker([Path("one.jpg")], "model", events, cancel)
        event = events.get_nowait()
        self.assertEqual(event[0], "cancelled")
        self.assertEqual(event[1:], (0, 1))

    def test_polling_updates_ui_from_background_events(self):
        class FakeValue:
            def __init__(self):
                self.value = None

            def set(self, value):
                self.value = value

        class FakeWidget:
            def __init__(self):
                self.calls = []

            def configure(self, **kwargs):
                self.calls.append(kwargs)

        class FakeRoot(FakeWidget):
            def config(self, **kwargs):
                self.calls.append(kwargs)

            def after(self, *_args):
                raise AssertionError("polling should finish without another callback")

        app = object.__new__(BookkeepingApp)
        app.reading = True
        app.read_events = queue.Queue()
        app.cancel_reading_event = threading.Event()
        app.receipts = []
        app.progress_var, app.progress_text = FakeValue(), FakeValue()
        app.read_button = app.cancel_button = app.clear_button = app.write_button = FakeWidget()
        app.root = FakeRoot()
        app.refresh_receipts = lambda: None
        app.validate = lambda receipt: None
        app.read_events.put(("receipt", 1, 1, Receipt("one.jpg", "客戶A", "8/8", [], [])))
        app.read_events.put(("finished", 1, 1))
        with patch("bookkeeping.messagebox.showinfo") as show_info:
            app._poll_read_events()
        self.assertFalse(app.reading)
        self.assertEqual(len(app.receipts), 1)
        self.assertEqual(app.progress_var.value, 1)
        show_info.assert_called_once()


class XlsxWriterTests(unittest.TestCase):
    def test_writes_numbers_preserves_formula_and_sets_recalculation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "template.xlsx", root / "output.xlsx"
            make_template(template)
            XlsxWriter(template, output).write([("客戶A", 4, dt.date(2026, 8, 1), 2, 170)])
            self.assertTrue(zipfile.is_zipfile(output))
            with zipfile.ZipFile(output) as archive:
                workbook = archive.read("xl/workbook.xml").decode("utf-8")
                self.assertIn('fullCalcOnLoad="1"', workbook)
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            cells = {cell.attrib["r"]: cell for cell in sheet.findall(f".//{{{NS_MAIN}}}c")}
            self.assertEqual(cells["I4"].findtext(f"{{{NS_MAIN}}}v"), "2")
            self.assertEqual(cells["J4"].findtext(f"{{{NS_MAIN}}}v"), "170")
            self.assertEqual(cells["K4"].findtext(f"{{{NS_MAIN}}}f"), "I4*J4")

    def test_failed_write_keeps_existing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "template.xlsx", root / "output.xlsx"
            make_template(template)
            output.write_bytes(b"existing-user-output")
            with self.assertRaises(ValueError):
                XlsxWriter(template, output).write([("不存在", 4, dt.date(2026, 8, 1), 1, 1)])
            self.assertEqual(output.read_bytes(), b"existing-user-output")


if __name__ == "__main__":
    unittest.main()
