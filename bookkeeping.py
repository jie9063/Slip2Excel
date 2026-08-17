"""桌面介面、照片辨識、品項比對與 Excel 寫入功能。"""

from __future__ import annotations

import base64
import copy
import datetime as dt
from io import BytesIO
import json
import logging
import math
import mimetypes
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from xml.etree import ElementTree as ET
from PIL import Image, ImageOps
from version import VERSION

APP_NAME = "Slip2Excel"
ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", ROOT))
ICON_PATH = RESOURCE_ROOT / "assets" / "slip2excel-icon.ico"
DATA_ROOT = (Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME) if getattr(sys, "frozen", False) else ROOT
DATA_ROOT.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = DATA_ROOT / "config.json"
LOG_DIR = DATA_ROOT / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    filename=LOG_DIR / "app.log", encoding="utf-8", level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)


DEFAULT_CONFIG = {
    "photo_folder": str(ROOT / "TrainPhoto"),
    "template_file": str(ROOT / "繼民客戶11508應收帳款(空白表單).xlsx"),
    "output_file": str(ROOT / "應收帳款_已填寫.xlsx"),
    "vision_provider": "ollama",
    "local_model": "qwen2.5vl:7b",
    "openai_model": "gpt-4o",
    "onboarding_completed": False,
    "customers": [],
    "products": [],
}


@dataclass
class Item:
    name: str
    quantity: float
    price: float
    line: int


@dataclass
class Receipt:
    file: str
    customer: str
    date: str
    items: list[Item]
    problems: list[str]
    warnings: list[str] = field(default_factory=list)


def is_customer_sheet_name(name: object) -> bool:
    """Exclude template and summary sheets that must never receive receipt data."""
    if not isinstance(name, str) or not name.strip():
        return False
    return not any(token in name for token in ("空白表單", "總表", "工作表"))


def remove_template_placeholder_mappings(customers: object) -> list[dict]:
    """Drop auto-created mappings for template/summary tabs, while preserving renamed tabs."""
    if not isinstance(customers, list):
        return []
    return [
        customer for customer in customers
        if isinstance(customer, dict)
        and (customer.get("template_sheet") or is_customer_sheet_name(customer.get("sheet")))
    ]


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        save_config(copy.deepcopy(DEFAULT_CONFIG))
    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        config = copy.deepcopy(DEFAULT_CONFIG)
    for key, value in DEFAULT_CONFIG.items():
        config.setdefault(key, copy.deepcopy(value))
    # 空白範本已含客戶工作表；首次使用時自動建立客戶名單，免除手動逐一輸入。
    changed = False
    if not config["customers"]:
        template = Path(config["template_file"])
        if template.exists():
            try:
                with zipfile.ZipFile(template) as z:
                    workbook = ET.fromstring(z.read("xl/workbook.xml"))
                names = [sheet.attrib["name"] for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet")]
                config["customers"] = [
                    {"sheet": name, "aliases": [name]}
                    for name in names
                    if is_customer_sheet_name(name)
                ]
                changed = True
            except (OSError, zipfile.BadZipFile, ET.ParseError):
                logging.exception("無法自動載入範本客戶")
    # 已由實際單據驗證的店名別名，以及兩張範本預留頁的新客戶。
    known_aliases = {
        "纖活": ["織活"],
        "大湳": ["八德大湳", "八德大滿", "八德大浦"],
        "強尼中原": ["強尼中原寶號"],
        "內壢": ["蛋白內壢寶號", "蛋白內壇"],
    }
    by_sheet = {customer["sheet"]: customer for customer in config["customers"]}
    for sheet, aliases in known_aliases.items():
        if sheet in by_sheet:
            existing = by_sheet[sheet].setdefault("aliases", [])
            for alias in aliases:
                if alias not in existing:
                    existing.append(alias)
                    changed = True
    for sheet, template_sheet, aliases in [
        ("蛋白大竹", "空白表單 (4)", ["蛋白大竹"]),
        ("桃園南崁", "空白表單 (2)", ["桃園南崁"]),
    ]:
        if sheet not in by_sheet:
            config["customers"].append({"sheet": sheet, "template_sheet": template_sheet, "aliases": aliases})
            changed = True
    cleaned_customers = remove_template_placeholder_mappings(config["customers"])
    if len(cleaned_customers) != len(config["customers"]):
        config["customers"] = cleaned_customers
        changed = True
    if changed:
        save_config(config)
    return config


def save_config(config: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def read_template_sheet_names(path: Path) -> list[str]:
    """Return the worksheet names from an .xlsx template without opening Excel."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [sheet.attrib["name"] for sheet in root.findall(f".//{{{NS_MAIN}}}sheet")]


def merge_customer_mappings(existing: object, imported: object) -> tuple[list[dict], int]:
    """Safely merge customer aliases, returning the merged list and imported count.

    Only ``sheet``, ``template_sheet`` and ``aliases`` are accepted.  This keeps
    old paths and any unrelated settings out of a user's current configuration.
    """
    merged: list[dict] = []
    by_sheet: dict[str, dict] = {}

    def add_entries(entries: object, count_imported: bool) -> int:
        added = 0
        if not isinstance(entries, list):
            return added
        for entry in entries:
            if not isinstance(entry, dict) or not isinstance(entry.get("sheet"), str):
                continue
            sheet = entry["sheet"].strip()
            if not sheet:
                continue
            aliases = entry.get("aliases", [])
            if not isinstance(aliases, list):
                aliases = []
            clean_aliases = [str(alias).strip() for alias in aliases if str(alias).strip()]
            if sheet not in clean_aliases:
                clean_aliases.insert(0, sheet)
            target = by_sheet.get(sheet)
            if target is None:
                target = {"sheet": sheet, "aliases": []}
                template_sheet = entry.get("template_sheet")
                if isinstance(template_sheet, str) and template_sheet.strip():
                    target["template_sheet"] = template_sheet.strip()
                by_sheet[sheet] = target
                merged.append(target)
            elif not target.get("template_sheet") and isinstance(entry.get("template_sheet"), str):
                target["template_sheet"] = entry["template_sheet"].strip()
            before = len(target["aliases"])
            for alias in clean_aliases:
                if alias not in target["aliases"]:
                    target["aliases"].append(alias)
            if count_imported:
                added += len(target["aliases"]) - before
        return added

    add_entries(existing, False)
    return merged, add_entries(imported, True)


def normalise(text: str) -> str:
    return re.sub(r"[\s　()（）\-＿_]", "", str(text)).lower()


def col_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def col_number(column: str) -> int:
    value = 0
    for char in column:
        value = value * 26 + ord(char) - 64
    return value


def parse_date(value: str) -> dt.date:
    value = str(value).strip().replace(".", "/").replace("-", "/")
    parts = [int(x) for x in re.findall(r"\d+", value)]
    today = dt.date.today()
    if len(parts) == 2:
        return dt.date(today.year, parts[0], parts[1])
    # 部分收據在日期前寫時間，例如「10:30 8 14」；最後兩個數字才是月、日。
    if len(parts) > 3 and 1 <= parts[-2] <= 12 and 1 <= parts[-1] <= 31:
        return dt.date(today.year, parts[-2], parts[-1])
    if len(parts) == 3:
        if parts[0] >= 100:
            year, month, day = parts
        elif parts[0] > 31 and 1 <= parts[1] <= 12 and 1 <= parts[2] <= 31:
            # OCR may join a printed form label to a handwritten date, for
            # example reading "8/7" as "87/08/07". A value above 31 cannot
            # be the month, so the final month/day pair is unambiguous.
            year, month, day = today.year, parts[1], parts[2]
        elif parts[2] >= 100:
            month, day, year = parts
        else:
            raise ValueError("三段日期請使用 115/8/8、2026/8/8 或 8/8/2026 格式")
        if year < 1911:  # 民國年
            year += 1911
        return dt.date(year, month, day)
    raise ValueError("日期請使用 8/8、115/8/8 或 2026/8/8 格式")


def prepare_entries(receipts: list[Receipt]) -> tuple[list[tuple[str, int, dt.date, float, float]], int]:
    """Create Excel entries and ignore identical duplicate photos safely."""
    # Excel has one column group for each day of the selected monthly template.
    # The year/month are not part of its cell address, so they must not be part
    # of the duplicate key either; otherwise an OCR year error could overwrite
    # the exact same Excel cell without being detected.
    slots: dict[tuple[str, int, int], tuple[dt.date, float, float]] = {}
    duplicates = 0
    for receipt in receipts:
        date = parse_date(receipt.date)
        for item in receipt.items:
            row = 3 + item.line
            key = (receipt.customer, row, date.day)
            value = (float(item.quantity), float(item.price))
            previous = slots.get(key)
            if previous is None:
                slots[key] = (date, *value)
            elif previous[1:] == value:
                duplicates += 1
            else:
                raise ValueError(
                    f"資料衝突：{receipt.customer}、{date:%m/%d}、第 {item.line} 列有不同數量或單價。"
                    "請保留正確的照片後再匯出。"
                )
    entries = [
        (customer, row, date, quantity, price)
        for (customer, row, _day), (date, quantity, price) in slots.items()
    ]
    return entries, duplicates


def classify_duplicate_and_conflicting_receipts(receipts: list[Receipt]) -> tuple[int, int]:
    """Keep only the first receipt for each customer/day and flag every later one.

    The accounting form has one position for a customer on each day.  The first
    photo encountered owns that position; every following photo is held without
    writing and identifies the retained photo for the user to compare.
    """
    grouped: dict[tuple[str, int], list[Receipt]] = {}
    for receipt in receipts:
        if receipt.problems:
            continue
        try:
            date = parse_date(receipt.date)
        except ValueError:
            continue
        grouped.setdefault((receipt.customer, date.day), []).append(receipt)

    duplicate_photos = repeated_photos = 0
    for (customer, day), group in grouped.items():
        if len(group) < 2:
            continue
        original = group[0]
        original_signature = (
            tuple(sorted((item.line, float(item.quantity), float(item.price)) for item in original.items)),
            tuple(original.warnings),
        )
        for later_receipt in group[1:]:
            later_signature = (
                tuple(sorted((item.line, float(item.quantity), float(item.price)) for item in later_receipt.items)),
                tuple(later_receipt.warnings),
            )
            if later_signature == original_signature:
                later_receipt.problems.append(
                    f"重複單據：{customer}、日期 {day} 日已使用「{original.file}」；此照片未寫入 Excel。"
                )
                duplicate_photos += 1
            else:
                later_receipt.problems.append(
                    f"同日重複單據：{customer}、日期 {day} 日已使用「{original.file}」；"
                    "此照片內容不同，為避免覆蓋資料未寫入 Excel。"
                )
                repeated_photos += 1
    return duplicate_photos, repeated_photos


def assert_write_plan_complete(
    receipts: list[Receipt], entries: list[tuple[str, int, dt.date, float, float]], duplicate_count: int,
) -> int:
    """Reject an export plan unless every approved item has one target cell."""
    expected_entries = sum(len(receipt.items) for receipt in receipts)
    if duplicate_count:
        raise ValueError(
            f"安全檢查發現 {duplicate_count} 筆重複的 Excel 寫入位置，已停止匯出。"
            "請回到「問題紀錄」確認重複照片後再試。"
        )
    if len(entries) != expected_entries:
        raise ValueError(
            f"安全檢查失敗：預計寫入 {expected_entries} 筆資料，實際只建立 {len(entries)} 個 Excel 位置。"
            "為避免漏寫，已停止匯出。"
        )
    return expected_entries


class XlsxWriter:
    """直接修改 XLSX 的 worksheet XML，完整保留範本的圖片、格式和版面。"""

    def __init__(self, template: Path, output: Path):
        self.template, self.output = template, output

    def write(self, entries: list[tuple[str, int, dt.date, float, float]], sheet_renames: dict[str, str] | None = None) -> bool:
        if not self.template.exists():
            raise FileNotFoundError(f"找不到空白表單：{self.template}")
        self.output.parent.mkdir(parents=True, exist_ok=True)
        if self.template.resolve() == self.output.resolve():
            raise ValueError("完成檔不能與空白 Excel 表單設定為同一個檔案。")
        # A later batch must retain the earlier batches.  Use the verified output
        # as the base when it exists; otherwise create the first output from the
        # selected blank form.  All writes still happen in a temporary file.
        appending_to_existing_output = self.output.exists()
        source_file = self.output if appending_to_existing_output else self.template
        if appending_to_existing_output and (not source_file.is_file() or not zipfile.is_zipfile(source_file)):
            raise ValueError("現有完成檔不是可讀取的 Excel 檔案；為保護資料，本次沒有寫入。")
        # 在同一資料夾先建立暫存檔，全部驗證與寫入成功後才取代完成檔。
        # 任何中途錯誤都不會破壞既有完成檔。
        with tempfile.NamedTemporaryFile(
            prefix=f".{self.output.stem}-", suffix=self.output.suffix,
            dir=self.output.parent, delete=False,
        ) as handle:
            temporary = Path(handle.name)
        try:
            shutil.copy2(source_file, temporary)
            with zipfile.ZipFile(temporary, "r") as source:
                files = {info.filename: source.read(info.filename) for info in source.infolist()}

        # Do not serialize workbook.xml with ElementTree: the template contains
        # Office extension namespaces whose original prefixes must be preserved.
            workbook_xml = files["xl/workbook.xml"].decode("utf-8")
            workbook = ET.fromstring(workbook_xml)
            sheet_renames = sheet_renames or {}
            existing_names = {sheet.attrib["name"] for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet")}
            for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
                original = sheet.attrib["name"]
                target = sheet_renames.get(original)
                if target and target not in existing_names:
                    safe_original = re.escape(original)
                    safe_target = (target.replace("&", "&amp;").replace('"', "&quot;")
                                          .replace("<", "&lt;").replace(">", "&gt;"))
                    workbook_xml, count = re.subn(
                        rf'(<sheet\b[^>]*\bname="){safe_original}(")',
                        rf'\g<1>{safe_target}\g<2>', workbook_xml, count=1,
                    )
                    if count != 1:
                        raise ValueError(f"無法建立客戶工作表「{target}」。")
                    sheet.attrib["name"] = target
                    existing_names.remove(original)
                    existing_names.add(target)
            rels = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
            rel_map = {
                relation.attrib["Id"]: relation.attrib["Target"]
                for relation in rels
                if relation.attrib.get("Type", "").endswith("/worksheet")
            }
            sheet_paths = {}
            for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
                relation_id = sheet.attrib[f"{{{NS_REL}}}id"]
                sheet_paths[sheet.attrib["name"]] = "xl/" + rel_map[relation_id].lstrip("/")

            self._assert_target_cells_empty(files, sheet_paths, entries)
            by_sheet: dict[str, list[tuple[int, dt.date, float, float]]] = {}
            for customer, row, date, quantity, price in entries:
                by_sheet.setdefault(customer, []).append((row, date, quantity, price))
            for customer, rows in by_sheet.items():
                if customer not in sheet_paths:
                    raise ValueError(f"範本內沒有客戶工作表「{customer}」")
                path = sheet_paths[customer]
                sheet_xml = files[path].decode("utf-8")
                for row, date, quantity, price in rows:
                    start_col = 9 + (date.day - 1) * 3
                    sheet_xml = self._set_number_xml(sheet_xml, row, start_col, quantity)
                    sheet_xml = self._set_number_xml(sheet_xml, row, start_col + 1, price)
                files[path] = sheet_xml.encode("utf-8")

            files["xl/workbook.xml"] = self._force_recalculation_xml(workbook_xml).encode("utf-8")
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
                for name, content in files.items():
                    target.writestr(name, content)
            self._verify_written_entries(temporary, sheet_paths, entries)
            temporary.replace(self.output)
            return appending_to_existing_output
        except Exception:
            temporary.unlink(missing_ok=True)
            raise

    @staticmethod
    def _assert_target_cells_empty(
        files: dict[str, bytes], sheet_paths: dict[str, str], entries: list[tuple[str, int, dt.date, float, float]],
    ) -> None:
        """Refuse a second receipt for a customer/day already present in the file."""
        by_sheet: dict[str, set[dt.date]] = {}
        for customer, _row, date, _quantity, _price in entries:
            by_sheet.setdefault(customer, set()).add(date)
        occupied: list[str] = []
        for customer, dates in by_sheet.items():
            path = sheet_paths.get(customer)
            if not path:
                # The later worksheet-existence check supplies the clearer error.
                continue
            root = ET.fromstring(files[path])
            cells = {cell.attrib.get("r", ""): cell for cell in root.findall(f".//{{{NS_MAIN}}}c")}
            for date in dates:
                first_column = 9 + (date.day - 1) * 3
                # Form rows 4 through 19 are the 16 writable item rows.  Scan
                # all of them, not just this batch's item lines: a second photo
                # for the same customer/date must never add to the first photo.
                for row in range(4, 20):
                    for column in (first_column, first_column + 1):
                        address = f"{col_name(column)}{row}"
                        cell = cells.get(address)
                        if cell is None:
                            continue
                        value = cell.findtext(f"{{{NS_MAIN}}}v")
                        has_formula = cell.find(f"{{{NS_MAIN}}}f") is not None
                        try:
                            is_blank_value = value is None or math.isclose(float(value), 0, rel_tol=0, abs_tol=1e-9)
                        except (TypeError, ValueError):
                            is_blank_value = False
                        if has_formula or not is_blank_value:
                            shown_value = value if value not in (None, "") else "公式"
                            occupied.append(f"{customer}、{date:%m/%d}（已有資料於 {address}：{shown_value}）")
                            break
                    if occupied:
                        break
                if len(occupied) >= 5:
                    break
            if len(occupied) >= 5:
                break
        if occupied:
            raise ValueError(
                "安全檢查發現同一客戶、同一天已經寫入完成檔：" + "；".join(occupied)
                + "。為避免重複單據，這一批完全沒有寫入 Excel。"
            )

    @staticmethod
    def _verify_written_entries(
        workbook_path: Path,
        sheet_paths: dict[str, str],
        entries: list[tuple[str, int, dt.date, float, float]],
    ) -> None:
        """Reopen the finished XLSX and verify every requested quantity and price.

        This protects users from a misleading success message if an unusual Excel
        template causes a cell write to be lost or redirected.
        """
        by_sheet: dict[str, list[tuple[int, dt.date, float, float]]] = {}
        for customer, row, date, quantity, price in entries:
            by_sheet.setdefault(customer, []).append((row, date, quantity, price))
        problems: list[str] = []
        with zipfile.ZipFile(workbook_path) as archive:
            for customer, rows in by_sheet.items():
                path = sheet_paths.get(customer)
                if not path:
                    problems.append(f"找不到工作表「{customer}」")
                    continue
                root = ET.fromstring(archive.read(path))
                values = {
                    cell.attrib.get("r", ""): cell.findtext(f"{{{NS_MAIN}}}v")
                    for cell in root.findall(f".//{{{NS_MAIN}}}c")
                }
                for row, date, quantity, price in rows:
                    first_column = 9 + (date.day - 1) * 3
                    for column, expected in ((first_column, quantity), (first_column + 1, price)):
                        address = f"{col_name(column)}{row}"
                        actual = values.get(address)
                        try:
                            matches = actual is not None and math.isclose(float(actual), float(expected), rel_tol=0, abs_tol=1e-9)
                        except (TypeError, ValueError):
                            matches = False
                        if not matches:
                            problems.append(f"{customer}!{address} 預期 {expected:g}，實際 {actual or '空白'}")
                            if len(problems) >= 5:
                                break
                    if len(problems) >= 5:
                        break
                if len(problems) >= 5:
                    break
        if problems:
            detail = "；".join(problems)
            raise ValueError(f"Excel 寫入驗證失敗：{detail}。原有完成檔未被覆蓋。")

    @staticmethod
    def _force_recalculation_xml(workbook_xml: str) -> str:
        def rewrite(match: re.Match) -> str:
            attributes = re.sub(
                r'\s(?:calcMode|fullCalcOnLoad|forceFullCalc)="[^"]*"',
                "", match.group(1),
            )
            return (f'<calcPr{attributes} calcMode="auto" '
                    'fullCalcOnLoad="1" forceFullCalc="1"/>')

        updated, count = re.subn(r'<calcPr\b([^>]*)/\s*>', rewrite, workbook_xml, count=1)
        if count:
            return updated
        return workbook_xml.replace(
            "</workbook>",
            '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
            1,
        )

    @staticmethod
    def _set_number_xml(sheet_xml: str, row_number: int, column: int, value: float) -> str:
        """Set a numeric cell without reserializing Office extension namespaces."""
        address = f"{col_name(column)}{row_number}"
        number = str(int(value)) if float(value).is_integer() else str(value)
        cell = f'<c r="{address}" s="1"><v>{number}</v></c>'
        row_pattern = rf'(<row\b[^>]*\br="{row_number}"[^>]*>)(.*?)(</row>)'
        row_match = re.search(row_pattern, sheet_xml, flags=re.DOTALL)
        if row_match is None:
            raise ValueError(f"Excel 範本找不到第 {row_number} 列。")
        body = row_match.group(2)
        existing = re.search(rf'<c\b[^>]*\br="{re.escape(address)}"[^>]*>.*?</c>', body, flags=re.DOTALL)
        if existing:
            body = body[:existing.start()] + cell + body[existing.end():]
        else:
            insert_at = len(body)
            for candidate in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"[^>]*>.*?</c>', body, flags=re.DOTALL):
                if col_number(candidate.group(1)) > column:
                    insert_at = candidate.start()
                    break
            body = body[:insert_at] + cell + body[insert_at:]
        return sheet_xml[:row_match.start(2)] + body + sheet_xml[row_match.end(2):]

    @staticmethod
    def _row_node(sheet_data: ET.Element, row_number: int) -> ET.Element:
        for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            if int(row.attrib.get("r", "0")) == row_number:
                return row
        row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_number)})
        inserted = False
        for index, existing in enumerate(sheet_data.findall(f"{{{NS_MAIN}}}row")):
            if int(existing.attrib.get("r", "0")) > row_number:
                sheet_data.insert(index, row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row)
        return row

    @classmethod
    def _cell(cls, sheet_data: ET.Element, row: int, column: int) -> ET.Element:
        row_node = cls._row_node(sheet_data, row)
        address = f"{col_name(column)}{row}"
        for cell in row_node.findall(f"{{{NS_MAIN}}}c"):
            if cell.attrib.get("r") == address:
                for child in list(cell):
                    cell.remove(child)
                cell.attrib.pop("t", None)
                return cell
        cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": address})
        row_node.append(cell)
        return cell

    @classmethod
    def _ensure_empty(cls, sheet_data: ET.Element, row: int, column: int) -> None:
        """避免把已登錄的數量或單價悄悄覆寫。0 視為尚未填寫。"""
        address = f"{col_name(column)}{row}"
        for row_node in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            if int(row_node.attrib.get("r", "0")) != row:
                continue
            for cell in row_node.findall(f"{{{NS_MAIN}}}c"):
                if cell.attrib.get("r") != address:
                    continue
                value = cell.findtext(f"{{{NS_MAIN}}}v")
                if value not in (None, "", "0", "0.0"):
                    raise ValueError(f"儲存格 {address} 已有資料（{value}）。可能是重複匯入，為避免覆寫已停止。")
            break

    @classmethod
    def _set_number(cls, sheet_data: ET.Element, row: int, column: int, value: float) -> None:
        cell = cls._cell(sheet_data, row, column)
        ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = str(value)

    @classmethod
    def _set_formula(cls, sheet_data: ET.Element, row: int, column: int, formula: str) -> None:
        cell = cls._cell(sheet_data, row, column)
        ET.SubElement(cell, f"{{{NS_MAIN}}}f").text = formula


class LocalVisionReader:
    """透過本機 Ollama 讀取圖片；所有照片只留在使用者電腦。"""

    API_URL = "http://127.0.0.1:11434/api"

    def __init__(self, model: str):
        self.model = model.strip()

    def check_ready(self) -> None:
        request = urllib.request.Request(f"{self.API_URL}/tags", method="GET")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                models = json.loads(response.read().decode("utf-8")).get("models", [])
        except urllib.error.URLError as error:
            raise RuntimeError(
                "找不到本機 Ollama。請先安裝 Ollama 後重新開啟程式。\n"
                "下載位置：https://ollama.com/download"
            ) from error
        available = {item.get("name", "") for item in models}
        if self.model not in available:
            raise RuntimeError(
                f"找不到本機模型「{self.model}」。請開啟 PowerShell 後執行：\n"
                f"ollama pull {self.model}\n\n"
                "首次下載約 3.2 GB；下載完成後回到本程式再按一次讀取照片。"
            )

    def read(self, photo: Path) -> Receipt:
        image_bytes, _mime_type = self._prepare_image(photo)
        encoded = base64.b64encode(image_bytes).decode("ascii")
        prompt = """請辨識這張繁體中文手寫估價單。只輸出 JSON，不要 Markdown。
表格從左到右依序是：品名、數量、單價、金額、備註。請讀取每一列「數量」和它右邊緊鄰的「單價」手寫欄位；金額欄通常是空白，不可拿來當單價。
格式必須是：{"customer":"客戶或店名","date":"YYYY/MM/DD 或 MM/DD","items":[{"line":1,"name":"品名","quantity":"數量，可含斤/包等單位","price":"單價"}]}
line 必須是單據左側印出的列號（1 到 16），不可自行編號。第 17 列以後是表單的小計或總計，不可輸出。僅保留品名欄已有文字、且數量與單價皆實際填寫的列；空白列不可輸出。看不清楚請略過該列，不要猜。"""
        body = {
            "model": self.model,
            "stream": False,
            "format": "json",
            "options": {"temperature": 0},
            "messages": [{"role": "user", "content": prompt, "images": [encoded]}],
        }
        request = urllib.request.Request(
            f"{self.API_URL}/chat", data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST",
        )
        last_text = ""
        # 小型本機模型偶有空白回覆；同張單自動重試一次，無須使用者重按。
        for _attempt in range(2):
            try:
                with urllib.request.urlopen(request, timeout=180) as response:
                    response_json = json.loads(response.read().decode("utf-8"))
            except urllib.error.HTTPError as error:
                detail = error.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"本機模型回覆錯誤 ({error.code})：{detail[:300]}") from error
            except urllib.error.URLError as error:
                raise RuntimeError(f"無法連線至本機模型：{error.reason}") from error
            last_text = response_json.get("message", {}).get("content", "").strip()
            if last_text and not re.fullmatch(r"[@#*\s]+", last_text):
                return self.from_json(last_text, photo.name)
        raise ValueError("本機模型未能讀取此張單據（回覆空白或亂碼）；請重新拍攝清晰照片後再試。")

    @staticmethod
    def _prepare_image(photo: Path) -> tuple[bytes, str]:
        """Return a model-ready image without ever changing the source photo."""
        original = photo.read_bytes()
        mime_type = mimetypes.guess_type(photo.name)[0] or "image/jpeg"
        try:
            with Image.open(photo) as source:
                image = ImageOps.exif_transpose(source)
                if image.width <= image.height:
                    return original, mime_type
                # The supported estimate form is portrait. Gallery exports can
                # contain the same form as landscape pixels without EXIF data.
                image = image.rotate(90, expand=True)
                if image.mode not in {"RGB", "L"}:
                    image = image.convert("RGB")
                buffer = BytesIO()
                image.save(buffer, format="JPEG", quality=95)
                logging.info("已將橫向單據轉正後辨識：%s", photo.name)
                return buffer.getvalue(), "image/jpeg"
        except (OSError, ValueError):
            logging.warning("無法轉正圖片，將使用原始檔辨識：%s", photo.name)
            return original, mime_type

    @staticmethod
    def from_json(text: str, filename: str = "手動輸入") -> Receipt:
        text = text.strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
        try:
            data = json.loads(text)
        except json.JSONDecodeError as error:
            raise ValueError(f"辨識結果不是正確 JSON：{error.msg}") from error
        if not isinstance(data, dict):
            raise ValueError("辨識結果的最外層必須是 JSON 物件")
        problems, warnings, items = [], [], []
        seen_lines: set[int] = set()
        if not data.get("customer"):
            problems.append("找不到客戶名稱")
        if not data.get("date"):
            problems.append("找不到日期")
        raw_items = data.get("items", [])
        if not isinstance(raw_items, list):
            problems.append("品項格式不是清單")
            raw_items = []
        for index, raw in enumerate(raw_items, start=1):
            raw_name = raw.get("name", "未命名品項") if isinstance(raw, dict) else "格式錯誤項目"
            try:
                if not isinstance(raw, dict):
                    raise ValueError("品項必須是 JSON 物件")
                name = str(raw.get("name", "")).strip()
                # 模型偶爾把空白列的序號（13、14…）誤當成品名，直接捨棄。
                if raw.get("quantity") is None or raw.get("price") is None:
                    raise ValueError
                line_value = float(raw.get("line", index))
                if not line_value.is_integer():
                    raise ValueError("列號必須是整數")
                line = int(line_value)
                if not 1 <= line <= 16:
                    raise ValueError("列號必須介於 1 到 16；小計與總計列不可寫入")
                if line in seen_lines:
                    raise ValueError("列號重複")
                quantity = LocalVisionReader._number(raw["quantity"])
                price = LocalVisionReader._number(raw["price"])
                seen_lines.add(line)
                items.append(Item(name, quantity, price, line))
            except (ValueError, TypeError, AttributeError):
                warnings.append(f"略過不完整項目：{raw_name}")
        if not items:
            problems.append("沒有可用的品項")
        return Receipt(filename, str(data.get("customer", "")), str(data.get("date", "")), items, problems, warnings)

    @staticmethod
    def _number(value: object) -> float:
        """接受模型常輸出的「3斤」、「半斤」、「3斤半」等值。"""
        if isinstance(value, bool):
            raise ValueError
        if isinstance(value, (int, float)):
            number = float(value)
            if not math.isfinite(number) or number <= 0:
                raise ValueError
            return number
        text = str(value).strip().replace(" ", "")
        if not text:
            raise ValueError
        if re.search(r"-\d", text):
            raise ValueError
        match = re.search(r"\d+(?:\.\d+)?", text)
        if match:
            number = float(match.group())
            if "半" in text and not text.startswith("半"):
                number += 0.5
            if not math.isfinite(number) or number <= 0:
                raise ValueError
            return number
        if "半" in text:
            return 0.5
        raise ValueError


class OpenAIVisionReader(LocalVisionReader):
    """Read a receipt photo through OpenAI's Responses API.

    The key is deliberately supplied at runtime instead of being saved to config.json.
    """

    API_URL = "https://api.openai.com/v1/responses"
    PROMPT = (
        "Read this handwritten accounting form. Return its customer name, date, and "
        "each filled item line's quantity and unit price. The item name is optional. "
        "Use the line number printed on the form (1 through 16). Rows 17 and later are totals; never return them. Do not invent values."
    )
    RESPONSE_SCHEMA = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "customer": {"type": "string"},
            "date": {"type": "string"},
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "line": {"type": "integer"},
                        "name": {"type": "string"},
                        "quantity": {"type": "number"},
                        "price": {"type": "number"},
                    },
                    "required": ["line", "name", "quantity", "price"],
                },
            },
        },
        "required": ["customer", "date", "items"],
    }

    def __init__(self, model: str, api_key: str = ""):
        super().__init__(model)
        self.api_key = api_key.strip() or os.environ.get("OPENAI_API_KEY", "").strip()

    def _request(self, url: str, *, data: bytes | None = None, method: str = "POST") -> dict:
        if not self.api_key:
            raise RuntimeError("尚未輸入 OpenAI API 金鑰。請在設定頁貼上金鑰，或設定 OPENAI_API_KEY 環境變數。")
        request = urllib.request.Request(
            url, data=data,
            headers={"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
            method=method,
        )
        try:
            with urllib.request.urlopen(request, timeout=180) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            detail = error.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"OpenAI API 請求失敗（{error.code}）：{detail[:300]}") from error
        except urllib.error.URLError as error:
            raise RuntimeError(f"無法連線至 OpenAI API：{error.reason}") from error

    def check_ready(self) -> None:
        if not self.model:
            raise RuntimeError("請填寫 OpenAI 視覺模型名稱。")
        self._request(
            f"https://api.openai.com/v1/models/{urllib.parse.quote(self.model, safe='')}", method="GET"
        )

    def read(self, photo: Path) -> Receipt:
        image_bytes, mime_type = self._prepare_image(photo)
        encoded = base64.b64encode(image_bytes).decode("ascii")
        body = {
            "model": self.model,
            "input": [{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": self.PROMPT},
                    {"type": "input_image", "image_url": f"data:{mime_type};base64,{encoded}", "detail": "high"},
                ],
            }],
            "text": {
                "format": {"type": "json_schema", "name": "receipt", "strict": True, "schema": self.RESPONSE_SCHEMA}
            },
        }
        response = self._request(self.API_URL, data=json.dumps(body).encode("utf-8"))
        for output in response.get("output", []):
            for content in output.get("content", []):
                if content.get("type") == "output_text" and content.get("text", "").strip():
                    return self.from_json(content["text"], photo.name)
        raise ValueError("OpenAI 未回傳可讀取的辨識結果。")


def create_vision_reader(provider: str, model: str, api_key: str = "") -> LocalVisionReader:
    if provider == "openai":
        return OpenAIVisionReader(model, api_key)
    return LocalVisionReader(model)


def read_photos_worker(
    photos: list[Path], model: str, events: queue.Queue, cancel_event: threading.Event,
    provider: str = "ollama", api_key: str = "",
) -> None:
    """Read photos off the Tkinter thread and report only plain data events."""
    total = len(photos)
    try:
        reader = create_vision_reader(provider, model, api_key)
        reader.check_ready()
    except Exception as error:
        events.put(("fatal", str(error)))
        return

    for index, photo in enumerate(photos, start=1):
        if cancel_event.is_set():
            events.put(("cancelled", index - 1, total))
            return
        try:
            receipt = reader.read(photo)
        except Exception as error:
            logging.exception("照片讀取失敗：%s", photo.name)
            receipt = Receipt(photo.name, "", "", [], [str(error)])
        events.put(("receipt", index, total, receipt))
    events.put(("finished", total, total))


class BookkeepingApp:
    def __init__(self):
        self.config = load_config()
        self.receipts: list[Receipt] = []
        self.reading = False
        self.read_events: queue.Queue = queue.Queue()
        self.cancel_reading_event = threading.Event()
        self.root = tk.Tk()
        self.root.title(f"Slip2Excel v{VERSION}")
        if ICON_PATH.exists():
            self.root.iconbitmap(default=str(ICON_PATH))
        self.root.geometry("1050x720")
        self.root.minsize(900, 620)
        self._build()
        self.root.after(250, self._show_first_run_guide)

    def run(self):
        self.root.mainloop()

    def _show_first_run_guide(self):
        """Show a short, non-technical setup guide once for each user profile."""
        if self.config.get("onboarding_completed"):
            return
        messagebox.showinfo(
            "首次使用｜三個步驟完成設定",
            "1. 點上方「設定」。\n"
            "2. 依序選擇「原始照片資料夾」與「空白 Excel 表單」。\n"
            "3. 按「儲存設定」，再回到「讀取與寫入」按「讀取照片資料夾」。\n\n"
            "選擇空白 Excel 表單時，程式會自動建立客戶對照。\n"
            "照片讀完並確認結果後，按「寫入 Excel」即可完成。",
        )
        self.config["onboarding_completed"] = True
        save_config(self.config)

    def _build(self):
        self.root.configure(background="#F1F5F9")
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#F1F5F9")
        style.configure("Surface.TFrame", background="#FFFFFF")
        style.configure("Toolbar.TFrame", background="#FFFFFF")
        style.configure("Title.TLabel", background="#FFFFFF", foreground="#0F172A", font=("Microsoft JhengHei UI", 18, "bold"))
        style.configure("Subtitle.TLabel", background="#FFFFFF", foreground="#64748B", font=("Microsoft JhengHei UI", 10))
        style.configure("Hint.TLabel", background="#F1F5F9", foreground="#64748B", font=("Microsoft JhengHei UI", 9))
        style.configure("TLabel", background="#F1F5F9", foreground="#0F172A", font=("Microsoft JhengHei UI", 10))
        style.configure("Surface.TLabel", background="#FFFFFF", foreground="#0F172A", font=("Microsoft JhengHei UI", 10))
        style.configure("ToolbarHint.TLabel", background="#FFFFFF", foreground="#64748B", font=("Microsoft JhengHei UI", 9))
        style.configure("Primary.TButton", background="#2563EB", foreground="#FFFFFF", bordercolor="#2563EB", padding=(14, 7), font=("Microsoft JhengHei UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#94A3B8")], foreground=[("disabled", "#E2E8F0")])
        style.configure("Secondary.TButton", background="#FFFFFF", foreground="#1E3A8A", bordercolor="#CBD5E1", padding=(12, 7), font=("Microsoft JhengHei UI", 10))
        style.map("Secondary.TButton", background=[("active", "#EFF6FF"), ("disabled", "#F1F5F9")], foreground=[("disabled", "#94A3B8")])
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#0F172A", bordercolor="#CBD5E1", padding=6)
        style.configure("Blue.Horizontal.TProgressbar", troughcolor="#E2E8F0", background="#2563EB", lightcolor="#2563EB", darkcolor="#2563EB", bordercolor="#E2E8F0", thickness=9)
        style.configure("App.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#0F172A", rowheight=30, bordercolor="#CBD5E1", font=("Microsoft JhengHei UI", 10))
        style.map("App.Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", "#0F172A")])
        style.configure("App.Treeview.Heading", background="#EFF6FF", foreground="#1E3A8A", bordercolor="#CBD5E1", relief="flat", font=("Microsoft JhengHei UI", 10, "bold"))
        style.map("App.Treeview.Heading", background=[("active", "#DBEAFE")])
        style.configure("App.TNotebook", background="#F1F5F9", borderwidth=0)
        style.configure("App.TNotebook.Tab", background="#E2E8F0", foreground="#475569", padding=(16, 9), font=("Microsoft JhengHei UI", 10, "bold"))
        style.map("App.TNotebook.Tab", background=[("selected", "#FFFFFF"), ("active", "#DBEAFE")], foreground=[("selected", "#1D4ED8")])
        frame = ttk.Frame(self.root, padding=16, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        header = ttk.Frame(frame, padding=(18, 15), style="Surface.TFrame")
        header.pack(fill="x", pady=(0, 12))
        tk.Frame(header, background="#2563EB", width=4, height=42).pack(side="left", padx=(0, 12))
        header_text = ttk.Frame(header, style="Surface.TFrame")
        header_text.pack(side="left", fill="x", expand=True)
        ttk.Label(header_text, text="Slip2Excel", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header_text, text="讀取照片、確認結果，再寫入 Excel。找不到的資料會保留在問題清單，不會直接寫入。", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 0))
        self.tabs = ttk.Notebook(frame, style="App.TNotebook")
        self.tabs.pack(fill="both", expand=True)
        self.process_tab = ttk.Frame(self.tabs, padding=16, style="App.TFrame")
        self.mapping_tab = ttk.Frame(self.tabs, padding=16, style="App.TFrame")
        self.settings_tab = ttk.Frame(self.tabs, padding=16, style="App.TFrame")
        self.tabs.add(self.process_tab, text="讀取與寫入")
        self.tabs.add(self.mapping_tab, text="問題紀錄")
        self.tabs.add(self.settings_tab, text="設定")
        self._build_process()
        self._build_issues()
        self._build_settings()

    def _build_process(self):
        top = ttk.Frame(self.process_tab, padding=12, style="Toolbar.TFrame")
        top.pack(fill="x")
        self.read_button = ttk.Button(top, text="讀取照片資料夾", command=self.analyse_photos, style="Primary.TButton")
        self.read_button.pack(side="left")
        self.cancel_button = ttk.Button(top, text="取消讀取", command=self.cancel_reading, state="disabled", style="Secondary.TButton")
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.clear_button = ttk.Button(top, text="清除本次結果", command=self.clear_receipts, style="Secondary.TButton")
        self.clear_button.pack(side="left", padx=(8, 0))
        self.write_button = ttk.Button(top, text="寫入 Excel", command=self.write_excel, style="Primary.TButton")
        self.write_button.pack(side="right")
        progress_row = ttk.Frame(self.process_tab, padding=(12, 10), style="Surface.TFrame")
        progress_row.pack(fill="x", pady=(10, 0))
        self.progress_var = tk.IntVar(value=0)
        self.progress_text = tk.StringVar(value="尚未開始讀取")
        self.progress_bar = ttk.Progressbar(progress_row, variable=self.progress_var, maximum=1, style="Blue.Horizontal.TProgressbar")
        self.progress_bar.pack(side="left", fill="x", expand=True)
        ttk.Label(progress_row, textvariable=self.progress_text, width=24, anchor="e", style="Surface.TLabel").pack(side="left", padx=(12, 0))
        columns = ("photo", "customer", "date", "items", "status")
        self.receipt_tree = ttk.Treeview(self.process_tab, columns=columns, show="headings", height=15, style="App.Treeview")
        labels = {"photo": "照片", "customer": "客戶", "date": "日期", "items": "品項數", "status": "結果"}
        widths = {"photo": 270, "customer": 160, "date": 120, "items": 70, "status": 380}
        for column in columns:
            self.receipt_tree.heading(column, text=labels[column])
            self.receipt_tree.column(column, width=widths[column], anchor="w")
        self.receipt_tree.pack(fill="both", expand=True, pady=(12, 8))
        self.detail = tk.Text(self.process_tab, height=9, font=("Microsoft JhengHei UI", 10), background="#FFFFFF", foreground="#0F172A", insertbackground="#0F172A", relief="flat", highlightthickness=1, highlightbackground="#CBD5E1", highlightcolor="#2563EB", padx=10, pady=8)
        self.detail.pack(fill="x")
        self.receipt_tree.bind("<<TreeviewSelect>>", self.show_detail)
        ttk.Label(self.process_tab, text="使用方式：單據第 1、2、3…列會自動寫入 Excel 的固定列，不需要輸入菜名。紅色或有問題的照片會略過，不會寫入。", style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

    def _build_issues(self):
        bar = ttk.Frame(self.mapping_tab, padding=12, style="Toolbar.TFrame")
        bar.pack(fill="x")
        ttk.Button(bar, text="重新整理", command=self.refresh_issues, style="Secondary.TButton").pack(side="left")
        self.issue_summary = ttk.Label(bar, text="尚未讀取照片", style="ToolbarHint.TLabel")
        self.issue_summary.pack(side="left", padx=12)
        columns = ("photo", "customer", "date", "type", "reason")
        self.issues_tree = ttk.Treeview(self.mapping_tab, columns=columns, show="headings", height=15, style="App.Treeview")
        labels = {"photo": "照片", "customer": "客戶", "date": "日期", "type": "狀態", "reason": "原因"}
        widths = {"photo": 280, "customer": 140, "date": 100, "type": 110, "reason": 430}
        for column in columns:
            self.issues_tree.heading(column, text=labels[column])
            self.issues_tree.column(column, width=widths[column], anchor="w")
        self.issues_tree.pack(fill="both", expand=True, pady=(12, 8))
        self.issue_detail = tk.Text(self.mapping_tab, height=9, font=("Microsoft JhengHei UI", 10), background="#FFFFFF", foreground="#0F172A", insertbackground="#0F172A", relief="flat", highlightthickness=1, highlightbackground="#CBD5E1", highlightcolor="#2563EB", padx=10, pady=8)
        self.issue_detail.pack(fill="x")
        self.issues_tree.bind("<<TreeviewSelect>>", self.show_issue_detail)
        ttk.Label(self.mapping_tab, text="這裡只顯示讀取失敗或有不完整項目的照片。點選照片可查看原因；完整可寫入的照片不會顯示。", style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

    def _build_settings_legacy(self):
        self.settings_vars = {
            key: tk.StringVar(value=str(self.config[key]))
            for key in ("photo_folder", "template_file", "output_file", "vision_provider", "local_model", "openai_model")
        }
        self.openai_api_key_var = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        grid = ttk.Frame(self.settings_tab, padding=16, style="Surface.TFrame")
        grid.pack(fill="x")
        ttk.Label(grid, text="辨識服務", style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=7)
        provider_box = ttk.Combobox(
            grid, textvariable=self.settings_vars["vision_provider"], state="readonly",
            values=("ollama", "openai"), width=20,
        )
        provider_box.grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="ollama 為免費離線；openai 會使用 API 額度", style="Surface.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Label(grid, text="OpenAI 視覺模型", style="Surface.TLabel").grid(row=1, column=0, sticky="w", pady=7)
        ttk.Entry(grid, textvariable=self.settings_vars["openai_model"], width=75).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Label(grid, text="OpenAI API 金鑰", style="Surface.TLabel").grid(row=2, column=0, sticky="w", pady=7)
        ttk.Entry(grid, textvariable=self.openai_api_key_var, show="●", width=75).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Label(grid, text="只在本次開啟程式期間使用，不會儲存", style="Surface.TLabel").grid(row=2, column=2, sticky="w")
        labels = [("photo_folder", "原始照片資料夾"), ("template_file", "空白 Excel 表單"), ("output_file", "完成 Excel 儲存位置"), ("local_model", "本機視覺模型")]
        for row, (key, label) in enumerate(labels, start=3):
            ttk.Label(grid, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=7)
            ttk.Entry(grid, textvariable=self.settings_vars[key], width=75).grid(row=row, column=1, sticky="ew", padx=8)
            if key != "local_model":
                ttk.Button(grid, text="選擇", command=lambda item=key: self.choose_path(item), style="Secondary.TButton").grid(row=row, column=2)
        grid.columnconfigure(1, weight=1)
        controls = ttk.Frame(self.settings_tab, style="App.TFrame")
        controls.pack(anchor="w", pady=14)
        ttk.Button(controls, text="檢查本機模型", command=self.check_local_model, style="Secondary.TButton").pack(side="left")
        ttk.Button(controls, text="儲存設定", command=self.save_settings, style="Primary.TButton").pack(side="left", padx=8)
        ttk.Label(self.settings_tab, text="免費離線模式：請先安裝 Ollama，再以 PowerShell 執行：ollama pull qwen2.5vl:7b。首次下載約 6 GB；照片不會上傳。", style="Hint.TLabel", wraplength=780).pack(anchor="w")

    def _build_settings(self):
        self.settings_vars = {
            key: tk.StringVar(value=str(self.config[key]))
            for key in ("photo_folder", "template_file", "output_file", "vision_provider", "local_model", "openai_model")
        }
        self.openai_api_key_var = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        self.model_choice_var = tk.StringVar(value=self._selected_model_label())

        grid = ttk.Frame(self.settings_tab, padding=16, style="Surface.TFrame")
        grid.pack(fill="x")
        ttk.Label(grid, text="辨識模型", style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=7)
        model_box = ttk.Combobox(
            grid, textvariable=self.model_choice_var, values=self._model_choice_labels(),
            state="readonly", width=75,
        )
        model_box.grid(row=0, column=1, sticky="ew", padx=8)
        model_box.bind("<<ComboboxSelected>>", self._on_model_choice)
        ttk.Label(grid, text="選擇後才顯示需要的設定", style="Surface.TLabel").grid(row=0, column=2, sticky="w")

        self.openai_key_label = ttk.Label(grid, text="OpenAI API 金鑰", style="Surface.TLabel")
        self.openai_key_input = ttk.Entry(grid, textvariable=self.openai_api_key_var, show="●", width=75)
        self.openai_key_hint = ttk.Label(grid, text="只在本次開啟程式期間使用，不會儲存", style="Surface.TLabel")

        fields = [
            ("photo_folder", "原始照片資料夾"),
            ("template_file", "空白 Excel 表單"),
            ("output_file", "完成 Excel 儲存位置"),
        ]
        for row, (key, label) in enumerate(fields, start=2):
            ttk.Label(grid, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=7)
            ttk.Entry(grid, textvariable=self.settings_vars[key], width=75).grid(row=row, column=1, sticky="ew", padx=8)
            ttk.Button(grid, text="選擇", command=lambda item=key: self.choose_path(item), style="Secondary.TButton").grid(row=row, column=2)
        grid.columnconfigure(1, weight=1)
        self._update_model_settings()

        controls = ttk.Frame(self.settings_tab, style="App.TFrame")
        controls.pack(anchor="w", pady=14)
        ttk.Button(controls, text="檢查目前模型", command=self.check_local_model, style="Secondary.TButton").pack(side="left")
        ttk.Button(controls, text="儲存設定", command=self.save_settings, style="Primary.TButton").pack(side="left", padx=8)
        ttk.Label(
            self.settings_tab,
            text="本機 Ollama 不上傳照片且免費；選擇 OpenAI 時，照片會送往 OpenAI API 並使用你的 API 額度。",
            style="Hint.TLabel", wraplength=780,
        ).pack(anchor="w")

    def _model_choice_labels(self) -> tuple[str, str]:
        return (
            f"Ollama（免費離線）｜{self.settings_vars['local_model'].get()}",
            f"OpenAI API｜{self.settings_vars['openai_model'].get()}",
        )

    def _selected_model_label(self) -> str:
        labels = self._model_choice_labels()
        return labels[1] if self.settings_vars["vision_provider"].get() == "openai" else labels[0]

    def _on_model_choice(self, _event=None):
        self._update_model_settings()

    def _update_model_settings(self):
        is_openai = self.model_choice_var.get().startswith("OpenAI API")
        self.settings_vars["vision_provider"].set("openai" if is_openai else "ollama")
        if is_openai:
            self.openai_key_label.grid(row=1, column=0, sticky="w", pady=7)
            self.openai_key_input.grid(row=1, column=1, sticky="ew", padx=8)
            self.openai_key_hint.grid(row=1, column=2, sticky="w")
        else:
            self.openai_key_label.grid_remove()
            self.openai_key_input.grid_remove()
            self.openai_key_hint.grid_remove()

    def choose_path(self, key: str):
        if key == "photo_folder":
            selected = filedialog.askdirectory(initialdir=self.settings_vars[key].get() or str(ROOT))
        elif key == "template_file":
            selected = filedialog.askopenfilename(filetypes=[("Excel 表單", "*.xlsx")])
        else:
            selected = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 表單", "*.xlsx")])
        if selected:
            self.settings_vars[key].set(selected)
            if key == "template_file":
                self._add_template_customers(Path(selected), show_result=True)

    def save_settings(self):
        for key, var in self.settings_vars.items():
            self.config[key] = var.get().strip()
        template = Path(self.config["template_file"])
        if template.exists():
            self._add_template_customers(template, show_result=False)
        save_config(self.config)

    def _add_template_customers(self, template: Path, show_result: bool) -> int:
        """Create exact-name mappings for a newly selected Excel template."""
        try:
            sheet_names = read_template_sheet_names(template)
        except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as error:
            logging.warning("Unable to read Excel template %s: %s", template, error)
            if show_result:
                messagebox.showwarning("無法建立客戶對照", "無法讀取這份 Excel 表單，請確認檔案可正常開啟。")
            return 0
        mappings = [{"sheet": name, "aliases": [name]} for name in sheet_names if is_customer_sheet_name(name)]
        existing = remove_template_placeholder_mappings(self.config.get("customers", []))
        removed = len(self.config.get("customers", [])) - len(existing)
        merged, created = merge_customer_mappings(existing, mappings)
        if created or removed:
            self.config["customers"] = merged
            save_config(self.config)
            if show_result:
                messagebox.showinfo("已建立客戶對照", f"已從空白 Excel 表單建立 {created} 個客戶對照。")
        elif show_result:
            messagebox.showinfo("客戶對照已存在", "這份空白 Excel 表單的客戶對照已建立。")
        return created

    def check_local_model(self):
        """Check the provider currently selected in Settings."""
        self.save_settings()
        provider = self.config["vision_provider"]
        model = self.config["openai_model"] if provider == "openai" else self.config["local_model"]
        try:
            create_vision_reader(provider, model, self.openai_api_key_var.get()).check_ready()
        except RuntimeError as error:
            messagebox.showwarning("模型無法使用", str(error))
        else:
            messagebox.showinfo("模型可使用", f"已確認 {provider} 的模型 {model} 可以使用。")
        return
        self.save_settings()
        try:
            LocalVisionReader(self.config["local_model"]).check_ready()
        except RuntimeError as error:
            messagebox.showwarning("尚未準備完成", str(error))
        else:
            messagebox.showinfo("本機模型已就緒", f"「{self.config['local_model']}」可以使用。")

    def import_customers(self):
        template = Path(self.settings_vars.get("template_file", tk.StringVar(value=self.config["template_file"])).get())
        try:
            names = self._sheet_names(template)
        except Exception as error:
            messagebox.showerror("無法讀取範本", str(error)); return
        existing = {entry["sheet"] for entry in self.config["customers"]}
        for name in names:
            if name not in existing and "空白" not in name and "總表" not in name and "工作表" not in name:
                self.config["customers"].append({"sheet": name, "aliases": [name]})
        save_config(self.config); self.refresh_mapping()
        messagebox.showinfo("完成", f"已從範本讀取 {len(names)} 個工作表；請保留需要的客戶與必要別名。")

    @staticmethod
    def _sheet_names(path: Path) -> list[str]:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("xl/workbook.xml"))
        return [sheet.attrib["name"] for sheet in root.findall(f".//{{{NS_MAIN}}}sheet")]

    def add_customer(self):
        sheet = simpledialog.askstring("新增客戶", "Excel 工作表名稱：", parent=self.root)
        if not sheet: return
        aliases = simpledialog.askstring("新增客戶", "照片上可能出現的名稱／別名（逗號分隔）：", initialvalue=sheet, parent=self.root)
        self.config["customers"].append({"sheet": sheet.strip(), "aliases": self._split_aliases(aliases)})
        save_config(self.config); self.refresh_mapping()

    def add_product(self):
        choices = [entry["sheet"] for entry in self.config["customers"]]
        if not choices:
            messagebox.showwarning("請先新增客戶", "請先按「重新載入範本客戶」或新增客戶。"); return
        customer = simpledialog.askstring("新增品項", "客戶工作表名稱：\n" + "、".join(choices), parent=self.root)
        if not customer: return
        if customer not in choices:
            messagebox.showerror("名稱不符", "請輸入清單中的客戶工作表名稱。"); return
        name = simpledialog.askstring("新增品項", "單據上的品名：", parent=self.root)
        row = simpledialog.askinteger("新增品項", "Excel 列號（例如 4）：", minvalue=1, parent=self.root)
        if not name or not row: return
        aliases = simpledialog.askstring("新增品項", "別名（逗號分隔；可留白）：", initialvalue=name, parent=self.root)
        self.config["products"].append({"customer": customer, "name": name.strip(), "aliases": self._split_aliases(aliases), "row": row})
        save_config(self.config); self.refresh_mapping()

    @staticmethod
    def _split_aliases(value: str | None) -> list[str]:
        return [part.strip() for part in (value or "").split(",") if part.strip()]

    def refresh_mapping(self):
        for child in self.mapping_tree.get_children(): self.mapping_tree.delete(child)
        for index, customer in enumerate(self.config["customers"]):
            self.mapping_tree.insert("", "end", iid=f"c:{index}", values=("客戶", customer["sheet"], "", ", ".join(customer.get("aliases", [])), ""))

    def delete_mapping(self):
        chosen = self.mapping_tree.selection()
        if not chosen: return
        for item in sorted(chosen, reverse=True):
            kind, index = item.split(":")
            if kind == "c":
                self.config["customers"].pop(int(index))
        save_config(self.config); self.refresh_mapping()

    def analyse_photos(self):
        if self.reading:
            return
        self.save_settings()
        folder = Path(self.config["photo_folder"])
        photos = sorted([p for p in folder.glob("*") if p.suffix.lower() in {".jpg", ".jpeg", ".png"}]) if folder.exists() else []
        if not photos:
            messagebox.showwarning("沒有照片", "找不到 JPG 或 PNG 照片，請檢查原始照片資料夾。"); return
        self.receipts = []
        self.progress_bar.configure(maximum=len(photos))
        self.progress_var.set(0)
        self.progress_text.set(f"讀取中：0 / {len(photos)}")
        self.refresh_receipts()
        self.read_events = queue.Queue()
        self.cancel_reading_event = threading.Event()
        self._set_reading_state(True)
        provider = self.config["vision_provider"]
        model = self.config["openai_model"] if provider == "openai" else self.config["local_model"]
        worker = threading.Thread(
            target=read_photos_worker,
            args=(photos, model, self.read_events, self.cancel_reading_event, provider, self.openai_api_key_var.get()),
            daemon=True,
        )
        worker.start()
        self.root.after(80, self._poll_read_events)

    def _set_reading_state(self, reading: bool):
        self.reading = reading
        self.read_button.configure(state="disabled" if reading else "normal")
        self.cancel_button.configure(state="normal" if reading else "disabled")
        self.clear_button.configure(state="disabled" if reading else "normal")
        self.write_button.configure(state="disabled" if reading else "normal")

    def cancel_reading(self):
        if not self.reading:
            return
        self.cancel_reading_event.set()
        self.cancel_button.configure(state="disabled")
        self.progress_text.set("正在取消；目前照片完成後會停止")

    def _poll_read_events(self):
        finished = False
        while True:
            try:
                event = self.read_events.get_nowait()
            except queue.Empty:
                break
            kind = event[0]
            if kind == "receipt":
                _, index, total, receipt = event
                self.validate(receipt)
                self.receipts.append(receipt)
                self.progress_var.set(index)
                self.progress_text.set(f"讀取中：{index} / {total}")
                self.refresh_receipts()
            elif kind == "fatal":
                self._finish_reading("無法讀取照片", event[1])
                finished = True
            elif kind == "cancelled":
                _, completed, total = event
                self._finish_reading("已取消讀取", f"已完成 {completed} / {total} 張照片。")
                finished = True
            elif kind == "finished":
                _, completed, total = event
                duplicates, repeated = classify_duplicate_and_conflicting_receipts(self.receipts)
                valid = sum(not item.problems for item in self.receipts)
                self.progress_var.set(completed)
                self.progress_text.set(f"讀取完成：{completed} / {total}")
                notes = [f"共讀取 {completed} 張照片；{valid} 張可寫入。"]
                if duplicates:
                    notes.append(f"已略過 {duplicates} 張完全相同的重複照片。")
                if repeated:
                    notes.append(f"有 {repeated} 張同一客戶同日的照片未寫入；已在問題紀錄列出保留的照片。")
                notes.append("有問題的照片請到「問題紀錄」查看。")
                self._finish_reading("讀取完成", "\n".join(notes))
                finished = True
        if self.reading and not finished:
            self.root.after(80, self._poll_read_events)

    def _finish_reading(self, title: str, message: str):
        self._set_reading_state(False)
        self.refresh_receipts()
        messagebox.showinfo(title, message) if title != "無法讀取照片" else messagebox.showwarning(title, message)

    def validate(self, receipt: Receipt):
        if receipt.problems: return
        customer = self.match_customer(receipt.customer)
        if not customer:
            receipt.problems.append(f"找不到客戶對照：{receipt.customer}")
            return
        receipt.customer = customer
        try:
            parse_date(receipt.date)
        except ValueError as error:
            receipt.problems.append(str(error))
            return

    def match_customer(self, name: str) -> str | None:
        target = normalise(name)
        for customer in self.config["customers"]:
            candidates = [customer["sheet"], *customer.get("aliases", [])]
            if any(normalise(candidate) == target for candidate in candidates): return customer["sheet"]
        # 收據常在店名後加「工房」、「餐飲」等字樣，允許兩字以上的包含比對。
        candidates = []
        for customer in self.config["customers"]:
            for candidate in [customer["sheet"], *customer.get("aliases", [])]:
                candidate_normal = normalise(candidate)
                if len(candidate_normal) >= 2 and (candidate_normal in target or target in candidate_normal):
                    candidates.append(customer["sheet"])
        if len(set(candidates)) == 1:
            return candidates[0]
        return None

    def match_product(self, customer: str, name: str) -> dict | None:
        target = normalise(name)
        matches = [item for item in self.config["products"] if item["customer"] == customer and any(normalise(x) == target for x in [item["name"], *item.get("aliases", [])])]
        return matches[0] if len(matches) == 1 else None

    def refresh_issues(self):
        for child in self.issues_tree.get_children():
            self.issues_tree.delete(child)
        failed = incomplete = 0
        for index, receipt in enumerate(self.receipts):
            if receipt.problems:
                failed += 1
                issue_type = "讀取失敗"
                reason = "；".join(receipt.problems)
            elif receipt.warnings:
                incomplete += 1
                issue_type = "讀取不完整"
                reason = "；".join(receipt.warnings)
            else:
                continue
            self.issues_tree.insert(
                "", "end", iid=str(index),
                values=(receipt.file, receipt.customer or "-", receipt.date or "-", issue_type, reason),
                tags=("incomplete",) if issue_type.startswith("資料不完整") else ("failed",),
            )
        self.issues_tree.tag_configure("failed", foreground="#b00020")
        self.issues_tree.tag_configure("incomplete", foreground="#9a6700")
        if not self.receipts:
            summary = "尚未讀取照片"
        elif not failed and not incomplete:
            summary = "沒有讀取失敗或不完整的照片"
        else:
            summary = f"讀取失敗 {failed} 張｜讀取不完整 {incomplete} 張"
        self.issue_summary.configure(text=summary)

    def show_issue_detail(self, _event=None):
        chosen = self.issues_tree.selection()
        if not chosen:
            return
        receipt = self.receipts[int(chosen[0])]
        lines = [f"照片：{receipt.file}", f"客戶：{receipt.customer or '-'}", f"日期：{receipt.date or '-'}"]
        if receipt.warnings:
            lines += ["", "不完整項目：", *[f"- {warning}" for warning in receipt.warnings]]
        if receipt.problems:
            lines += ["", "讀取失敗原因：", *[f"- {problem}" for problem in receipt.problems]]
        self.issue_detail.delete("1.0", "end")
        self.issue_detail.insert("1.0", "end", "\n".join(lines))

    def refresh_receipts(self):
        for child in self.receipt_tree.get_children(): self.receipt_tree.delete(child)
        for index, receipt in enumerate(self.receipts):
            if receipt.problems:
                status = "；".join(receipt.problems)
            elif receipt.warnings:
                status = "可寫入（略過不完整品項；已記錄）"
            else:
                status = "可寫入"
            self.receipt_tree.insert("", "end", iid=str(index), values=(receipt.file, receipt.customer, receipt.date, len(receipt.items), status), tags=("bad",) if receipt.problems else ())
        self.receipt_tree.tag_configure("bad", foreground="#b00020")
        self.refresh_issues()

    def show_detail(self, _event=None):
        chosen = self.receipt_tree.selection()
        if not chosen: return
        receipt = self.receipts[int(chosen[0])]
        lines = [f"照片：{receipt.file}", f"客戶：{receipt.customer}", f"日期：{receipt.date}", "", "品項："]
        lines += [f"- 第 {item.line} 列｜{item.name}｜數量 {item.quantity:g}｜單價 {item.price:g}" for item in receipt.items]
        if receipt.warnings: lines += ["", "略過項目：", *[f"- {x}" for x in receipt.warnings]]
        if receipt.problems: lines += ["", "問題：", *[f"- {x}" for x in receipt.problems]]
        self.detail.delete("1.0", "end"); self.detail.insert("1.0", "end", "\n".join(lines))

    def clear_receipts(self):
        if self.reading:
            return
        self.receipts = []
        self.progress_var.set(0)
        self.progress_bar.configure(maximum=1)
        self.progress_text.set("尚未開始讀取")
        self.refresh_receipts()
        self.detail.delete("1.0", "end")

    def write_excel(self):
        if self.reading:
            messagebox.showwarning("讀取進行中", "請等待照片讀取完成後再寫入 Excel。")
            return
        valid = [receipt for receipt in self.receipts if not receipt.problems]
        if not valid:
            messagebox.showwarning("沒有可寫入資料", "請先讀取照片，並處理所有缺少品項或日期問題。"); return
        entries = []
        duplicate_count = 0
        try:
            entries, duplicate_count = prepare_entries(valid)
            expected_entry_count = assert_write_plan_complete(valid, entries, duplicate_count)
            sheet_renames = {
                customer["template_sheet"]: customer["sheet"]
                for customer in self.config["customers"]
                if customer.get("template_sheet")
            }
            appended = XlsxWriter(Path(self.config["template_file"]), Path(self.config["output_file"])).write(entries, sheet_renames)
        except Exception as error:
            logging.exception("Excel 寫入失敗")
            messagebox.showerror("寫入失敗", str(error)); return
        skipped = len(self.receipts) - len(valid)
        warning_count = sum(len(receipt.warnings) for receipt in self.receipts)
        logging.info(
            "已安全寫入 %d 張照片、驗證 %d/%d 個 Excel 資料位置；擋下 %d 張、%d 個不完整項目",
            len(valid), len(entries), expected_entry_count, skipped, warning_count,
        )
        notes = [
            "已接續既有完成檔寫入本批資料。" if appended else "已從空白 Excel 表單建立完成檔。",
            f"已安全寫入 {len(valid)} 張照片。",
            f"已回讀驗證 {len(entries)} / {expected_entry_count} 個 Excel 資料位置。",
            f"已擋下 {skipped} 張重複或有問題的照片，沒有寫入 Excel。",
        ]
        if warning_count:
            notes.append(f"已略過並記錄 {warning_count} 個無法確認的品項；其餘已確認品項已寫入。")
        notes.append("請到「問題紀錄」查看略過品項與未寫入照片的原因。")
        messagebox.showinfo("安全寫入完成", "\n".join(notes) + f"\n\n完成檔：\n{self.config['output_file']}")
