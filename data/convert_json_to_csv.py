#!/usr/bin/env python3
"""Convert invoice annotation JSON files into a flat invoice-style CSV."""
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable

ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = ROOT_DIR / "annotations"
DEFAULT_OUTPUT_CSV = ROOT_DIR.parent / "result" / "validate" / "ground_truth" / "annotations.csv"

COLUMNS = [
    "Nomor",
    "Nama File",
    "Due Date",
    "PO Number",
    "Tel",
    "Email",
    "Site",
    "SUB_TOTAL",
    "DISCOUNT",
    "TAX",
    "TOTAL",
    "Bank Name",
    "Branch Name",
    "Bank Account Number",
    "Bank Swift Code",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read structured invoice annotation JSON files and export them as a "
            "clean table (one row per invoice)."
        )
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=DEFAULT_INPUT_DIR,
        type=Path,
        help=f"Directory containing annotation JSON files (default: {DEFAULT_INPUT_DIR}).",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT_CSV,
        type=Path,
        help=f"Destination CSV path (default: {DEFAULT_OUTPUT_CSV}).",
    )
    return parser.parse_args()


def get_text(data: Dict[str, Any], key: str) -> str:
    value = data.get(key)
    if isinstance(value, dict):
        text = value.get("text")
        if isinstance(text, str):
            return text.strip()
    elif isinstance(value, str):
        return value.strip()
    return ""


def normalize_spaces(value: str) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", value).strip()


def value_after_colon(text: str) -> str:
    if not text:
        return ""
    remainder = text.split(":", 1)[1] if ":" in text else text
    return normalize_spaces(remainder)


def extract_labeled_line(block: str, label: str) -> str:
    if not block:
        return ""
    prefix = label.lower()
    for line in block.splitlines():
        stripped = line.strip()
        if not stripped.lower().startswith(prefix):
            continue
        suffix = stripped[len(label) :]
        return normalize_spaces(suffix.lstrip(" :"))
    return ""


def load_annotation(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:  # pragma: no cover - user data validation
        raise RuntimeError(f"Failed to parse {path.name}: {exc}") from exc


def build_record(index: int, path: Path, data: Dict[str, Any]) -> Dict[str, str]:
    buyer_text = get_text(data, "BUYER")
    payment_text = get_text(data, "PAYMENT_DETAILS")
    return {
        "Nomor": str(index),
        "Nama File": f"{path.stem}.jpg",
        "Due Date": value_after_colon(get_text(data, "DUE_DATE")),
        "PO Number": value_after_colon(get_text(data, "PO_NUMBER")),
        "Tel": extract_labeled_line(buyer_text, "Tel"),
        "Email": extract_labeled_line(buyer_text, "Email"),
        "Site": extract_labeled_line(buyer_text, "Site"),
        "SUB_TOTAL": value_after_colon(get_text(data, "SUB_TOTAL")),
        "DISCOUNT": value_after_colon(get_text(data, "DISCOUNT")),
        "TAX": value_after_colon(get_text(data, "TAX")),
        "TOTAL": value_after_colon(get_text(data, "TOTAL")),
        "Bank Name": extract_labeled_line(payment_text, "Bank Name"),
        "Branch Name": extract_labeled_line(payment_text, "Branch Name"),
        "Bank Account Number": extract_labeled_line(
            payment_text, "Bank Account Number"
        ),
        "Bank Swift Code": extract_labeled_line(payment_text, "Bank Swift Code"),
    }


def iter_records(input_dir: Path) -> Iterable[Dict[str, str]]:
    json_files = sorted(input_dir.glob("*.json"))
    for index, json_file in enumerate(json_files, start=1):
        data = load_annotation(json_file)
        if not isinstance(data, dict):
            raise ValueError(f"{json_file.name} does not contain a JSON object")
        yield build_record(index, json_file, data)


def write_csv(records: Iterable[Dict[str, str]], output_path: Path) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=COLUMNS)
        writer.writeheader()
        for record in records:
            writer.writerow(record)
            count += 1
    return count


def main() -> None:
    args = parse_args()
    input_dir = args.input_dir
    if not input_dir.is_dir():
        raise NotADirectoryError(f"{input_dir} is not a directory")
    records = list(iter_records(input_dir))
    row_count = write_csv(records, args.output)
    print(f"Wrote {row_count} rows to {args.output}")


if __name__ == "__main__":
    main()
