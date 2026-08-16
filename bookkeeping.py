"""桌面介面、照片辨識、品項比對與 Excel 寫入功能。"""

from __future__ import annotations

import base64
import copy
import datetime as dt
import json
import logging
import math
import re
import shutil
import tempfile
import tkinter as tk
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass, asdict, field
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"
LOG_DIR = ROOT / "logs"
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
    "local_model": "qwen2.5vl:7b",
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
                    if "空白" not in name and "總表" not in name and "工作表" not in name
                ]
                changed = True
            except (OSError, zipfile.BadZipFile, ET.ParseError):
                logging.exception("無法自動載入範本客戶")
    # 已由實際單據驗證的店名別名，以及兩張範本預留頁的新客戶。
    known_aliases = {
        "纖活": ["織活"],
        "大湳": ["八德大湳", "八德大滿"],
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
    if changed:
        save_config(config)
    return config


def save_config(config: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


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
    slots: dict[tuple[str, int, dt.date], tuple[float, float]] = {}
    duplicates = 0
    for receipt in receipts:
        date = parse_date(receipt.date)
        for item in receipt.items:
            row = 3 + item.line
            key = (receipt.customer, row, date)
            value = (float(item.quantity), float(item.price))
            previous = slots.get(key)
            if previous is None:
                slots[key] = value
            elif previous == value:
                duplicates += 1
            else:
                raise ValueError(
                    f"資料衝突：{receipt.customer}、{date:%Y/%m/%d}、第 {item.line} 列有不同數量或單價。"
                    "請保留正確的照片後再匯出。"
                )
    entries = [(customer, row, date, quantity, price)
               for (customer, row, date), (quantity, price) in slots.items()]
    return entries, duplicates


class XlsxWriter:
    """直接修改 XLSX 的 worksheet XML，完整保留範本的圖片、格式和版面。"""

    def __init__(self, template: Path, output: Path):
        self.template, self.output = template, output

    def write(self, entries: list[tuple[str, int, dt.date, float, float]], sheet_renames: dict[str, str] | None = None) -> None:
        if not self.template.exists():
            raise FileNotFoundError(f"找不到空白表單：{self.template}")
        self.output.parent.mkdir(parents=True, exist_ok=True)
        if self.template.resolve() == self.output.resolve():
            raise ValueError("完成檔不能與空白 Excel 表單設定為同一個檔案。")
        # 在同一資料夾先建立暫存檔，全部驗證與寫入成功後才取代完成檔。
        # 任何中途錯誤都不會破壞既有完成檔。
        with tempfile.NamedTemporaryFile(
            prefix=f".{self.output.stem}-", suffix=self.output.suffix,
            dir=self.output.parent, delete=False,
        ) as handle:
            temporary = Path(handle.name)
        try:
            shutil.copy2(self.template, temporary)
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
            temporary.replace(self.output)
        except Exception:
            temporary.unlink(missing_ok=True)
            raise

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
        encoded = base64.b64encode(photo.read_bytes()).decode("ascii")
        prompt = """請辨識這張繁體中文手寫估價單。只輸出 JSON，不要 Markdown。
表格從左到右依序是：品名、數量、單價、金額、備註。請讀取每一列「數量」和它右邊緊鄰的「單價」手寫欄位；金額欄通常是空白，不可拿來當單價。
格式必須是：{"customer":"客戶或店名","date":"YYYY/MM/DD 或 MM/DD","items":[{"line":1,"name":"品名","quantity":"數量，可含斤/包等單位","price":"單價"}]}
line 必須是單據左側印出的列號（1 到 17），不可自行編號。僅保留品名欄已有文字、且數量與單價皆實際填寫的列；空白列不可輸出。看不清楚請略過該列，不要猜。"""
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
                if not name or name.isdigit() or raw.get("quantity") is None or raw.get("price") is None:
                    raise ValueError
                line_value = float(raw.get("line", index))
                if not line_value.is_integer():
                    raise ValueError("列號必須是整數")
                line = int(line_value)
                if not 1 <= line <= 17:
                    raise ValueError("列號必須介於 1 到 17")
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


class BookkeepingApp:
    def __init__(self):
        self.config = load_config()
        self.receipts: list[Receipt] = []
        self.root = tk.Tk()
        self.root.title("Slip2Excel")
        self.root.geometry("1050x720")
        self.root.minsize(900, 620)
        self._build()

    def run(self):
        self.root.mainloop()

    def _build(self):
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Microsoft JhengHei UI", 16, "bold"))
        style.configure("Hint.TLabel", foreground="#555555")
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="手寫單據自動記帳", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="先讀取照片、確認結果，再寫入 Excel。任何找不到的品項都不會寫入。", style="Hint.TLabel").pack(anchor="w", pady=(2, 12))
        self.tabs = ttk.Notebook(frame)
        self.tabs.pack(fill="both", expand=True)
        self.process_tab = ttk.Frame(self.tabs, padding=12)
        self.mapping_tab = ttk.Frame(self.tabs, padding=12)
        self.settings_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(self.process_tab, text="1. 讀取與寫入")
        self.tabs.add(self.mapping_tab, text="2. 辨識問題")
        self.tabs.add(self.settings_tab, text="設定")
        self._build_process()
        self._build_issues()
        self._build_settings()

    def _build_process(self):
        top = ttk.Frame(self.process_tab)
        top.pack(fill="x")
        ttk.Button(top, text="讀取照片資料夾", command=self.analyse_photos).pack(side="left")
        ttk.Button(top, text="清除本次結果", command=self.clear_receipts).pack(side="left")
        ttk.Button(top, text="寫入 Excel", command=self.write_excel).pack(side="right")
        progress_row = ttk.Frame(self.process_tab)
        progress_row.pack(fill="x", pady=(10, 0))
        self.progress_var = tk.IntVar(value=0)
        self.progress_text = tk.StringVar(value="尚未開始讀取")
        self.progress_bar = ttk.Progressbar(progress_row, variable=self.progress_var, maximum=1)
        self.progress_bar.pack(side="left", fill="x", expand=True)
        ttk.Label(progress_row, textvariable=self.progress_text, width=22, anchor="e").pack(side="left", padx=(10, 0))
        columns = ("photo", "customer", "date", "items", "status")
        self.receipt_tree = ttk.Treeview(self.process_tab, columns=columns, show="headings", height=15)
        labels = {"photo": "照片", "customer": "客戶", "date": "日期", "items": "品項數", "status": "結果"}
        widths = {"photo": 270, "customer": 160, "date": 120, "items": 70, "status": 380}
        for column in columns:
            self.receipt_tree.heading(column, text=labels[column])
            self.receipt_tree.column(column, width=widths[column], anchor="w")
        self.receipt_tree.pack(fill="both", expand=True, pady=(12, 8))
        self.detail = tk.Text(self.process_tab, height=10, font=("Microsoft JhengHei UI", 10))
        self.detail.pack(fill="x")
        self.receipt_tree.bind("<<TreeviewSelect>>", self.show_detail)
        ttk.Label(self.process_tab, text="使用方式：單據第 1、2、3…列會自動寫入 Excel 的固定列，不需要輸入菜名。紅色或有問題的照片會略過，不會寫入。", style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

    def _build_issues(self):
        bar = ttk.Frame(self.mapping_tab)
        bar.pack(fill="x")
        ttk.Button(bar, text="重新整理", command=self.refresh_issues).pack(side="left")
        self.issue_summary = ttk.Label(bar, text="尚未讀取照片", style="Hint.TLabel")
        self.issue_summary.pack(side="left", padx=12)
        columns = ("photo", "customer", "date", "type", "reason")
        self.issues_tree = ttk.Treeview(self.mapping_tab, columns=columns, show="headings", height=15)
        labels = {"photo": "照片", "customer": "客戶", "date": "日期", "type": "狀態", "reason": "原因"}
        widths = {"photo": 280, "customer": 140, "date": 100, "type": 110, "reason": 430}
        for column in columns:
            self.issues_tree.heading(column, text=labels[column])
            self.issues_tree.column(column, width=widths[column], anchor="w")
        self.issues_tree.pack(fill="both", expand=True, pady=(12, 8))
        self.issue_detail = tk.Text(self.mapping_tab, height=9, font=("Microsoft JhengHei UI", 10))
        self.issue_detail.pack(fill="x")
        self.issues_tree.bind("<<TreeviewSelect>>", self.show_issue_detail)
        ttk.Label(self.mapping_tab, text="這裡只顯示讀取失敗或有不完整項目的照片。點選照片可查看原因；完整可寫入的照片不會顯示。", style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

    def _build_settings(self):
        self.settings_vars = {key: tk.StringVar(value=str(self.config[key])) for key in ("photo_folder", "template_file", "output_file", "local_model")}
        grid = ttk.Frame(self.settings_tab)
        grid.pack(fill="x")
        labels = [("photo_folder", "原始照片資料夾"), ("template_file", "空白 Excel 表單"), ("output_file", "完成 Excel 儲存位置"), ("local_model", "本機視覺模型")]
        for row, (key, label) in enumerate(labels):
            ttk.Label(grid, text=label).grid(row=row, column=0, sticky="w", pady=7)
            ttk.Entry(grid, textvariable=self.settings_vars[key], width=75).grid(row=row, column=1, sticky="ew", padx=8)
            if key != "local_model":
                ttk.Button(grid, text="選擇", command=lambda item=key: self.choose_path(item)).grid(row=row, column=2)
        grid.columnconfigure(1, weight=1)
        controls = ttk.Frame(self.settings_tab)
        controls.pack(anchor="w", pady=14)
        ttk.Button(controls, text="檢查本機模型", command=self.check_local_model).pack(side="left")
        ttk.Button(controls, text="儲存設定", command=self.save_settings).pack(side="left", padx=8)
        ttk.Label(self.settings_tab, text="免費離線模式：請先安裝 Ollama，再以 PowerShell 執行：ollama pull qwen2.5vl:7b。首次下載約 6 GB；照片不會上傳。", style="Hint.TLabel", wraplength=780).pack(anchor="w")

    def choose_path(self, key: str):
        if key == "photo_folder":
            selected = filedialog.askdirectory(initialdir=self.settings_vars[key].get() or str(ROOT))
        elif key == "template_file":
            selected = filedialog.askopenfilename(filetypes=[("Excel 表單", "*.xlsx")])
        else:
            selected = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 表單", "*.xlsx")])
        if selected:
            self.settings_vars[key].set(selected)

    def save_settings(self):
        for key, var in self.settings_vars.items():
            self.config[key] = var.get().strip()
        save_config(self.config)

    def check_local_model(self):
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
        self.save_settings()
        folder = Path(self.config["photo_folder"])
        photos = sorted([p for p in folder.glob("*") if p.suffix.lower() in {".jpg", ".jpeg", ".png"}]) if folder.exists() else []
        if not photos:
            messagebox.showwarning("沒有照片", "找不到 JPG 或 PNG 照片，請檢查原始照片資料夾。"); return
        reader = LocalVisionReader(self.config["local_model"])
        try:
            reader.check_ready()
        except RuntimeError as error:
            messagebox.showwarning("無法讀取照片", str(error))
            return
        self.receipts = []
        self.progress_bar.configure(maximum=len(photos))
        self.progress_var.set(0)
        self.progress_text.set(f"讀取中：0 / {len(photos)}")
        self.root.config(cursor="watch"); self.root.update()
        for index, photo in enumerate(photos, start=1):
            try:
                receipt = reader.read(photo)
            except Exception as error:
                logging.exception("照片讀取失敗：%s", photo.name)
                receipt = Receipt(photo.name, "", "", [], [str(error)])
            self.validate(receipt)
            self.receipts.append(receipt)
            self.progress_var.set(index)
            self.progress_text.set(f"讀取中：{index} / {len(photos)}")
            self.refresh_receipts(); self.root.update()
        self.root.config(cursor="")
        self.refresh_receipts()
        valid = sum(not item.problems for item in self.receipts)
        self.progress_text.set(f"讀取完成：{len(photos)} / {len(photos)}")
        messagebox.showinfo("讀取完成", f"共讀取 {len(self.receipts)} 張照片；{valid} 張可寫入，其他請查看結果與 log。")

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
                tags=("failed",) if receipt.problems else ("incomplete",),
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
                status = "可寫入（" + "；".join(receipt.warnings) + "）"
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
        self.receipts = []
        self.progress_var.set(0)
        self.progress_bar.configure(maximum=1)
        self.progress_text.set("尚未開始讀取")
        self.refresh_receipts()
        self.detail.delete("1.0", "end")

    def write_excel(self):
        valid = [receipt for receipt in self.receipts if not receipt.problems]
        if not valid:
            messagebox.showwarning("沒有可寫入資料", "請先讀取照片，並處理所有缺少品項或日期問題。"); return
        entries = []
        duplicate_count = 0
        try:
            entries, duplicate_count = prepare_entries(valid)
            for receipt in []:
                date = parse_date(receipt.date)
                for item in receipt.items:
                    # 單據的第 1 列對應 Excel 第 4 列，因此無須辨識或填入菜名。
                    entries.append((receipt.customer, 3 + item.line, date, item.quantity, item.price))
            sheet_renames = {
                customer["template_sheet"]: customer["sheet"]
                for customer in self.config["customers"]
                if customer.get("template_sheet")
            }
            XlsxWriter(Path(self.config["template_file"]), Path(self.config["output_file"])).write(entries, sheet_renames)
        except Exception as error:
            logging.exception("Excel 寫入失敗")
            messagebox.showerror("寫入失敗", str(error)); return
        skipped = len(self.receipts) - len(valid)
        warning_count = sum(len(receipt.warnings) for receipt in valid)
        logging.info("成功寫入 %d 張照片、%d 筆資料；略過 %d 張、略過 %d 個不完整項目", len(valid), len(entries), skipped, warning_count)
        messagebox.showinfo("完成", f"已寫入 {len(valid)} 張照片、{len(entries)} 筆資料。\n\n完成檔：\n{self.config['output_file']}\n\n略過 {skipped} 張有問題的照片，原因已顯示並記錄於 logs/app.log。")
