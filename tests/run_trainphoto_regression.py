"""End-to-end regression runner for local TrainPhoto images.

This is intentionally separate from the fast unit suite: it calls the local
vision model and writes only a temporary Excel file and JSON report supplied by
the caller.  It never changes source photos, the configured template, or the
user's finished workbook.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bookkeeping import (  # noqa: E402
    BookkeepingApp,
    classify_duplicate_and_conflicting_receipts,
    LocalVisionReader,
    XlsxWriter,
    load_config,
    prepare_entries,
)


def serialise_receipt(receipt) -> dict:
    return {
        "file": receipt.file,
        "customer": receipt.customer,
        "date": receipt.date,
        "items": [
            {"line": item.line, "quantity": item.quantity, "price": item.price}
            for item in receipt.items
        ],
        "problems": receipt.problems,
        "warnings": receipt.warnings,
    }


def run(photo_folder: Path, report_path: Path) -> int:
    config = load_config()
    photos = sorted(
        photo for photo in photo_folder.iterdir()
        if photo.is_file() and photo.suffix.lower() in {".jpg", ".jpeg", ".png"}
    )
    if not photos:
        raise ValueError(f"No supported photos found in {photo_folder}")

    reader = LocalVisionReader(config["local_model"])
    reader.check_ready()
    validator = object.__new__(BookkeepingApp)
    validator.config = config
    receipts = []
    for index, photo in enumerate(photos, start=1):
        try:
            receipt = reader.read(photo)
            BookkeepingApp.validate(validator, receipt)
        except Exception as error:
            from bookkeeping import Receipt
            receipt = Receipt(photo.name, "", "", [], [str(error)])
        receipts.append(receipt)
        print(f"{index}/{len(photos)} {photo.name}: {'OK' if not receipt.problems else 'FAILED'}", flush=True)

    duplicate_photos, conflicting_photos = classify_duplicate_and_conflicting_receipts(receipts)
    valid = [receipt for receipt in receipts if not receipt.problems]
    report = {
        "photo_count": len(photos),
        "valid_receipt_count": len(valid),
        "failed_receipt_count": len(receipts) - len(valid),
        "duplicate_photo_count": duplicate_photos,
        "conflicting_photo_count": conflicting_photos,
        "warning_item_count": sum(len(receipt.warnings) for receipt in valid),
        "receipts": [serialise_receipt(receipt) for receipt in receipts],
    }
    try:
        entries, duplicates = prepare_entries(valid)
        sheet_renames = {
            customer["template_sheet"]: customer["sheet"]
            for customer in config["customers"]
            if customer.get("template_sheet")
        }
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "regression-output.xlsx"
            XlsxWriter(Path(config["template_file"]), output).write(entries, sheet_renames)
            report["xlsx_valid"] = output.is_file()
        report["entry_count"] = len(entries)
        report["duplicate_count"] = duplicates
    except Exception as error:
        report["xlsx_valid"] = False
        report["write_error"] = str(error)

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in report.items() if key != "receipts"}, ensure_ascii=False), flush=True)
    return 0 if report.get("xlsx_valid") and not report["failed_receipt_count"] else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--photos", type=Path, default=Path("TrainPhoto"))
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    raise SystemExit(run(args.photos, args.report))
