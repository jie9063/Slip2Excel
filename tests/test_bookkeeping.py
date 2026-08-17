import datetime as dt
from io import BytesIO
import json
import queue
import sys
import tempfile
import threading
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from unittest.mock import patch
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bookkeeping import (  # noqa: E402
    BookkeepingApp,
    assert_write_plan_complete,
    Item,
    LocalVisionReader,
    OpenAIVisionReader,
    classify_duplicate_and_conflicting_receipts,
    NS_MAIN,
    Receipt,
    XlsxWriter,
    parse_date,
    prepare_entries,
    create_vision_reader,
    is_customer_sheet_name,
    merge_customer_mappings,
    remove_template_placeholder_mappings,
    read_template_sheet_names,
    read_photos_worker,
)
from version import VERSION  # noqa: E402


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
    def test_public_version_uses_major_minor_format(self):
        self.assertRegex(VERSION, r"^\d+\.\d{2}$")

    def test_customer_alias_import_merges_without_other_settings(self):
        existing = [{"sheet": "客戶A", "aliases": ["客戶A", "A公司"]}]
        imported = [
            {"sheet": "客戶A", "aliases": ["A公司", "甲公司"]},
            {"sheet": "客戶B", "template_sheet": "表單 (2)", "aliases": ["乙公司"]},
            {"sheet": "", "aliases": ["無效"]},
            "invalid",
        ]
        merged, imported_count = merge_customer_mappings(existing, imported)
        self.assertEqual(imported_count, 3)
        self.assertEqual(merged, [
            {"sheet": "客戶A", "aliases": ["客戶A", "A公司", "甲公司"]},
            {"sheet": "客戶B", "template_sheet": "表單 (2)", "aliases": ["客戶B", "乙公司"]},
        ])

    def test_template_and_summary_tabs_are_not_customer_mappings(self):
        self.assertTrue(is_customer_sheet_name("客戶A"))
        for name in ("空白表單", "空白表單 (4)", "115客戶總表", "工作表1", ""):
            self.assertFalse(is_customer_sheet_name(name))
        cleaned = remove_template_placeholder_mappings([
            {"sheet": "客戶A", "aliases": ["客戶A"]},
            {"sheet": "空白表單 (4)", "aliases": ["空白表單 (4)"]},
            {"sheet": "蛋白大竹", "template_sheet": "空白表單 (4)", "aliases": ["蛋白大竹"]},
        ])
        self.assertEqual(cleaned, [
            {"sheet": "客戶A", "aliases": ["客戶A"]},
            {"sheet": "蛋白大竹", "template_sheet": "空白表單 (4)", "aliases": ["蛋白大竹"]},
        ])

    def test_template_sheet_names_can_seed_customer_mapping(self):
        with tempfile.TemporaryDirectory() as directory:
            template = Path(directory) / "template.xlsx"
            make_template(template)
            names = read_template_sheet_names(template)
        mappings, created = merge_customer_mappings([], [{"sheet": name, "aliases": [name]} for name in names])
        self.assertEqual(created, 1)
        self.assertEqual(mappings[0]["sheet"], names[0])
        self.assertEqual(mappings[0]["aliases"], [names[0]])

    def test_selecting_template_creates_customer_mapping(self):
        class FakeApp:
            def __init__(self):
                self.config = {"customers": []}

        with tempfile.TemporaryDirectory() as directory:
            template = Path(directory) / "template.xlsx"
            make_template(template)
            app = FakeApp()
            with patch("bookkeeping.save_config") as save:
                created = BookkeepingApp._add_template_customers(app, template, show_result=False)
        self.assertEqual(created, 1)
        self.assertEqual(len(app.config["customers"]), 1)
        self.assertEqual(app.config["customers"][0]["aliases"], [app.config["customers"][0]["sheet"]])
        save.assert_called_once_with(app.config)

    def test_first_run_guide_is_shown_once_and_saved(self):
        class FakeApp:
            def __init__(self):
                self.config = {"onboarding_completed": False}

        app = FakeApp()
        with patch("bookkeeping.messagebox.showinfo") as info, patch("bookkeeping.save_config") as save:
            BookkeepingApp._show_first_run_guide(app)
            BookkeepingApp._show_first_run_guide(app)
        self.assertTrue(app.config["onboarding_completed"])
        info.assert_called_once()
        save.assert_called_once_with(app.config)

    def test_parse_dates(self):
        self.assertEqual(parse_date("115/8/8"), dt.date(2026, 8, 8))
        self.assertEqual(parse_date("8/8/2026"), dt.date(2026, 8, 8))
        self.assertEqual(parse_date("87/08/07"), dt.date(dt.date.today().year, 8, 7))
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
        total_row = LocalVisionReader.from_json(
            '{"customer":"客戶A","date":"8/8","items":['
            '{"line":17,"name":"小計","quantity":1,"price":100}]}'
        )
        self.assertFalse(total_row.items)
        self.assertTrue(total_row.warnings)

    def test_landscape_image_is_rotated_for_model_without_changing_source(self):
        with tempfile.TemporaryDirectory() as directory:
            photo = Path(directory) / "landscape.jpg"
            Image.new("RGB", (80, 40), "white").save(photo)
            original = photo.read_bytes()
            image_bytes, mime_type = LocalVisionReader._prepare_image(photo)
            self.assertTrue(photo.exists())
            self.assertEqual(photo.read_bytes(), original)
        self.assertEqual(mime_type, "image/jpeg")
        self.assertNotEqual(image_bytes, original)
        with Image.open(BytesIO(image_bytes)) as prepared:
            self.assertGreater(prepared.height, prepared.width)

    def test_prepare_entries_deduplicates_and_rejects_conflicts(self):
        first = Receipt("a.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        same = Receipt("b.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        entries, duplicates = prepare_entries([first, same])
        self.assertEqual(len(entries), 1)
        self.assertEqual(duplicates, 1)
        conflict = Receipt("c.jpg", "客戶A", "8/8", [Item("", 3, 70, 1)], [])
        with self.assertRaises(ValueError):
            prepare_entries([first, conflict])
        wrong_year = Receipt("wrong-year.jpg", "客戶A", "2023/8/8", [Item("", 3, 70, 1)], [])
        with self.assertRaises(ValueError):
            prepare_entries([first, wrong_year])

    def test_export_plan_rejects_any_silent_deduplication(self):
        first = Receipt("a.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        same = Receipt("b.jpg", "客戶A", "8/8", [Item("", 2, 70, 1)], [])
        entries, duplicates = prepare_entries([first, same])
        with self.assertRaises(ValueError):
            assert_write_plan_complete([first, same], entries, duplicates)
        self.assertEqual(assert_write_plan_complete([first], entries, 0), 1)

    def test_second_same_customer_day_receipt_is_blocked_and_names_first_photo(self):
        first = Receipt("first.jpg", "客戶A", "8/8", [Item("", 1, 20, 1)], [])
        duplicate = Receipt("duplicate.jpg", "客戶A", "8/8", [Item("", 1, 20, 1)], [])
        duplicate_count, repeated_count = classify_duplicate_and_conflicting_receipts([first, duplicate])
        self.assertEqual((duplicate_count, repeated_count), (1, 0))
        self.assertFalse(first.problems)
        self.assertIn("first.jpg", duplicate.problems[0])
        self.assertIn("未寫入 Excel", duplicate.problems[0])

        left = Receipt("left.jpg", "客戶B", "8/9", [Item("", 1, 20, 1)], [])
        right = Receipt("right.jpg", "客戶B", "8/9", [Item("", 2, 20, 1)], [])
        duplicate_count, repeated_count = classify_duplicate_and_conflicting_receipts([left, right])
        self.assertEqual((duplicate_count, repeated_count), (0, 1))
        self.assertFalse(left.problems)
        self.assertIn("left.jpg", right.problems[0])
        self.assertIn("未寫入 Excel", right.problems[0])

    def test_incomplete_item_is_logged_but_recognised_items_remain_writable(self):
        app = object.__new__(BookkeepingApp)
        app.config = {"customers": [{"sheet": "客戶A", "aliases": ["客戶A"]}]}
        receipt = Receipt(
            "partial.jpg", "客戶A", "8/8", [Item("", 1, 20, 1)], [], ["第 2 列缺少單價"],
        )
        app.validate(receipt)
        self.assertFalse(receipt.problems)
        self.assertTrue(receipt.warnings)
        entries, duplicates = prepare_entries([receipt])
        self.assertEqual((len(entries), duplicates), (1, 0))

    def test_write_handler_exports_confirmed_items_from_a_partial_receipt(self):
        app = object.__new__(BookkeepingApp)
        app.reading = False
        app.receipts = [
            Receipt("partial.jpg", "客戶A", "8/8", [Item("", 1, 20, 1)], [], ["第 2 列缺少單價"]),
        ]
        app.config = {"customers": [], "template_file": "template.xlsx", "output_file": "output.xlsx"}
        with patch("bookkeeping.XlsxWriter.write", return_value=False) as write, patch("bookkeeping.messagebox.showinfo") as info:
            BookkeepingApp.write_excel(app)
        entries = write.call_args.args[0]
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0][3:], (1, 20))
        info.assert_called_once()

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

    def test_openai_reader_sends_image_and_reads_structured_response(self):
        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return json.dumps({
                    "output": [{"content": [{"type": "output_text", "text": json.dumps({
                        "customer": "客戶A", "date": "8/8",
                        "items": [{"line": 1, "name": "", "quantity": 2, "price": 70}],
                    })}]}]
                }).encode("utf-8")

        with tempfile.TemporaryDirectory() as directory:
            photo = Path(directory) / "receipt.jpg"
            photo.write_bytes(b"test-image")
            with patch("bookkeeping.urllib.request.urlopen", return_value=FakeResponse()) as open_url:
                receipt = OpenAIVisionReader("gpt-4o", "test-key").read(photo)

        request = open_url.call_args.args[0]
        body = json.loads(request.data.decode("utf-8"))
        self.assertEqual(request.full_url, "https://api.openai.com/v1/responses")
        self.assertEqual(request.get_header("Authorization"), "Bearer test-key")
        self.assertEqual(body["model"], "gpt-4o")
        self.assertTrue(body["input"][0]["content"][1]["image_url"].startswith("data:image/jpeg;base64,"))
        self.assertEqual(receipt.customer, "客戶A")
        self.assertEqual(receipt.items[0].price, 70)

    def test_openai_reader_requires_key_and_factory_uses_selected_provider(self):
        with patch.dict("bookkeeping.os.environ", {}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "API 金鑰"):
                OpenAIVisionReader("gpt-4o").check_ready()
        self.assertIsInstance(create_vision_reader("ollama", "qwen2.5vl:7b"), LocalVisionReader)
        self.assertIsInstance(create_vision_reader("openai", "gpt-4o", "key"), OpenAIVisionReader)

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

    def test_reading_state_does_not_change_global_mouse_cursor(self):
        class FakeWidget:
            def __init__(self):
                self.calls = []

            def configure(self, **kwargs):
                self.calls.append(kwargs)

        class FakeRoot:
            def config(self, **_kwargs):
                raise AssertionError("背景讀取不應改變全域滑鼠游標")

        app = object.__new__(BookkeepingApp)
        app.read_button = FakeWidget()
        app.cancel_button = FakeWidget()
        app.clear_button = FakeWidget()
        app.write_button = FakeWidget()
        app.root = FakeRoot()

        app._set_reading_state(True)

        self.assertTrue(app.reading)
        self.assertEqual(app.read_button.calls[-1], {"state": "disabled"})
        self.assertEqual(app.cancel_button.calls[-1], {"state": "normal"})

    def test_model_choice_only_shows_openai_key_for_openai(self):
        class FakeVar:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class FakeWidget:
            def __init__(self):
                self.calls = []

            def grid(self, **kwargs):
                self.calls.append(("grid", kwargs))

            def grid_remove(self):
                self.calls.append(("grid_remove", {}))

        app = object.__new__(BookkeepingApp)
        app.model_choice_var = FakeVar("OpenAI API｜gpt-4o")
        app.settings_vars = {"vision_provider": FakeVar("ollama")}
        app.openai_key_label = app.openai_key_input = app.openai_key_hint = FakeWidget()

        app._update_model_settings()
        self.assertEqual(app.settings_vars["vision_provider"].get(), "openai")
        self.assertEqual(app.openai_key_input.calls[-1][0], "grid")

        app.model_choice_var.set("Ollama（免費離線）｜qwen2.5vl:7b")
        app._update_model_settings()
        self.assertEqual(app.settings_vars["vision_provider"].get(), "ollama")
        self.assertEqual(app.openai_key_input.calls[-1][0], "grid_remove")


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

    def test_later_batch_appends_to_existing_output_without_losing_first_batch(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "template.xlsx", root / "output.xlsx"
            make_template(template)
            self.assertFalse(XlsxWriter(template, output).write([
                ("客戶A", 4, dt.date(2026, 8, 1), 2, 170),
            ]))
            self.assertTrue(XlsxWriter(template, output).write([
                ("客戶A", 4, dt.date(2026, 8, 2), 3, 180),
            ]))
            with zipfile.ZipFile(output) as archive:
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            cells = {cell.attrib["r"]: cell.findtext(f"{{{NS_MAIN}}}v") for cell in sheet.findall(f".//{{{NS_MAIN}}}c")}
            self.assertEqual(cells["I4"], "2")
            self.assertEqual(cells["J4"], "170")
            self.assertEqual(cells["L4"], "3")
            self.assertEqual(cells["M4"], "180")

    def test_later_batch_blocks_same_customer_date_even_with_a_different_item_row(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "template.xlsx", root / "output.xlsx"
            make_template(template)
            writer = XlsxWriter(template, output)
            writer.write([("客戶A", 4, dt.date(2026, 8, 1), 2, 170)])
            original = output.read_bytes()
            with self.assertRaisesRegex(ValueError, "同一客戶、同一天"):
                writer.write([("客戶A", 5, dt.date(2026, 8, 1), 3, 180)])
            self.assertEqual(output.read_bytes(), original)

    def test_verification_rejects_a_missing_or_wrong_written_value(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "template.xlsx", root / "output.xlsx"
            make_template(template)
            XlsxWriter(template, output).write([("客戶A", 4, dt.date(2026, 8, 1), 2, 170)])
            with self.assertRaisesRegex(ValueError, "Excel 寫入驗證失敗"):
                XlsxWriter._verify_written_entries(
                    output,
                    {"客戶A": "xl/worksheets/sheet1.xml"},
                    [("客戶A", 4, dt.date(2026, 8, 1), 3, 170)],
                )

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
