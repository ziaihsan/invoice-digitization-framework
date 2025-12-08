#!/usr/bin/env python3
"""Compare OCR CSV output with ground-truth annotations and emit per-column JSON stats."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_OCR = ROOT_DIR.parent / "ocr_results.csv"
DEFAULT_GROUND_TRUTH = ROOT_DIR / "ground_truth" / "annotations.csv"
DEFAULT_JSON_REPORT = ROOT_DIR / "validation_report.json"
DEFAULT_HTML_REPORT = ROOT_DIR / "validation_report.html"
DEFAULT_LOG_FILE = ROOT_DIR / "log.txt"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate OCR results against ground-truth annotations and output per-column JSON metrics."
        )
    )
    parser.add_argument(
        "--ocr",
        type=Path,
        default=DEFAULT_OCR,
        help=f"Path to the OCR CSV file (default: {DEFAULT_OCR}).",
    )
    parser.add_argument(
        "--ground-truth",
        type=Path,
        default=DEFAULT_GROUND_TRUTH,
        help=f"Path to the ground-truth CSV file (default: {DEFAULT_GROUND_TRUTH}).",
    )
    parser.add_argument(
        "--key",
        default="Nama File",
        help="Column used to align OCR rows with ground truth (default: %(default)s)",
    )
    parser.add_argument(
        "--max-samples",
        type=int,
        default=None,
        help="Optional cap on saved missmatch entries per column (default: no limit)",
    )
    parser.add_argument(
        "--json-out",
        type=Path,
        default=DEFAULT_JSON_REPORT,
        help=f"Path to save the JSON report (default: {DEFAULT_JSON_REPORT}).",
    )
    parser.add_argument(
        "--html-out",
        type=Path,
        default=DEFAULT_HTML_REPORT,
        help=f"Path to save the HTML report (default: {DEFAULT_HTML_REPORT}).",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        default=DEFAULT_LOG_FILE,
        help=f"Path to append execution logs (default: {DEFAULT_LOG_FILE}).",
    )
    return parser.parse_args()


def ordered_unique(values: List[str]) -> List[str]:
    seen = set()
    ordered = []
    for value in values:
        if value not in seen:
            seen.add(value)
            ordered.append(value)
    return ordered


def load_csv(path: Path, key: str) -> Tuple[Dict[str, Dict[str, str]], List[str], List[str]]:
    data = {}
    order = []
    path = Path(path)
    with path.open(newline="", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        fieldnames = reader.fieldnames or []
        if key not in fieldnames:
            raise SystemExit(f"Key column '{key}' not found in {path}")
        for line_number, row in enumerate(reader, start=2):
            row_key = (row.get(key) or "").strip()
            if not row_key:
                raise SystemExit(f"Missing key '{key}' at {path}:{line_number}")
            if row_key in data:
                raise SystemExit(f"Duplicate key '{row_key}' detected in {path}")
            data[row_key] = {k: (v or "") for k, v in row.items()}
            order.append(row_key)
    return data, order, fieldnames


def normalize(value: str) -> str:
    return value.strip()


def render_html_report(result: Dict[str, object]) -> str:
    from html import escape

    columns: Dict[str, Dict[str, object]] = result["columns"]  # type: ignore[assignment]
    missing_gt = result.get("missing_in_ground_truth", [])
    missing_ocr = result.get("missing_in_ocr_results", [])

    def fmt_accuracy(value):
        if value is None:
            return "–"
        return f"{value * 100:.2f}%"

    def render_missing(items):
        if not items:
            return "<li>Tidak ada</li>"
        return "".join(f"<li>{escape(str(item))}</li>" for item in items)

    def fmt_ratio(numerator, denominator, unit):
        denom = int(denominator or 0)
        if denom == 0:
            return "Tidak ada data"
        num = 0 if numerator is None else int(numerator)
        label = "baris" if unit == "rows" else "field"
        return f"{num} dari {denom} {label}"

    total_field_comparisons = int(result.get("total_field_comparisons") or 0)
    rows_compared = int(result.get("rows_compared") or 0)
    total_matches = sum(int(col_stats.get("matches", 0)) for col_stats in columns.values())
    total_matches_text = fmt_ratio(total_matches, total_field_comparisons, "fields")
    rows_passing_text = fmt_ratio(result.get("rows_passing"), rows_compared, "rows")
    rows_missing_text = fmt_ratio(result.get("rows_missing"), rows_compared, "rows")
    overall_missing_text = fmt_ratio(result.get("overall_missing"), total_field_comparisons, "fields")

    column_rows = []
    for column_name, stats in columns.items():
        samples = stats.get("missmatch", [])
        if samples:
            sample_items = "".join(
                "<li>"
                f"<strong>{escape(str(sample.get(result['key_column'], '')))}</strong>: "
                f"Expected <code>{escape(str(sample.get('expected', '')))}</code>, "
                f"OCR <code>{escape(str(sample.get('ocr', '')))}</code>"
                "</li>"
                for sample in samples
            )
            sample_html = f"<ul>{sample_items}</ul>"
        else:
            sample_html = "<em>Semua cocok</em>"

        column_rows.append(
            "<tr>"
            f"<td>{escape(column_name)}</td>"
            f"<td>{stats['total_compared']}</td>"
            f"<td>{stats['matches']}</td>"
            f"<td>{stats['mismatches']}</td>"
            f"<td>{fmt_accuracy(stats['accuracy'])}</td>"
            f"<td>{fmt_accuracy(stats.get('missing_rate'))}</td>"
            f"<td>{sample_html}</td>"
            "</tr>"
        )

    column_table_html = "".join(column_rows)

    return f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="utf-8" />
  <title>Laporan Validasi OCR</title>
  <style>
    body {{
      font-family: Arial, Helvetica, sans-serif;
      margin: 2rem;
      background: #fafafa;
      color: #1f1f1f;
    }}
    h1 {{
      margin-top: 0;
    }}
    .summary {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
      margin-bottom: 2rem;
    }}
    .card {{
      background: white;
      border-radius: 8px;
      padding: 1rem;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);
    }}
    .card .subtext {{
      color: #5a5a5a;
      font-size: 0.9rem;
      margin-top: 0.35rem;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      background: white;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);
    }}
    th, td {{
      padding: 0.75rem;
      border-bottom: 1px solid #e5e5e5;
      vertical-align: top;
    }}
    th {{
      background: #f0f4ff;
      text-align: left;
      font-weight: 600;
    }}
    tr:last-child td {{
      border-bottom: none;
    }}
    code {{
      background: #f5f5f5;
      padding: 0.1rem 0.25rem;
      border-radius: 4px;
      font-size: 0.95em;
    }}
    ul {{
      margin: 0.25rem 0 0 1.25rem;
    }}
    em {{
      color: #777;
    }}
  </style>
</head>
<body>
  <h1>Laporan Validasi OCR</h1>
  <p>Kolom acuan: <strong>{escape(str(result['key_column']))}</strong></p>
  <div class="summary">
    <div class="card">
      <strong>Total Ground Truth</strong>
      <div>{result['total_records_ground_truth']}</div>
    </div>
    <div class="card">
      <strong>Total OCR</strong>
      <div>{result['total_records_ocr']}</div>
    </div>
    <div class="card">
      <strong>Record Cocok</strong>
      <div>{result['shared_records']}</div>
    </div>
    <div class="card">
      <strong>Akurasi Keseluruhan</strong>
      <div>{fmt_accuracy(result.get('overall_accuracy'))}</div>
      <div class="subtext">{total_matches_text}</div>
    </div>
    <div class="card">
      <strong>Row Pass-Rate</strong>
      <div>{fmt_accuracy(result.get('row_pass_rate'))}</div>
      <div class="subtext">{rows_passing_text}</div>
    </div>
    <div class="card">
      <strong>Row Missing-Rate</strong>
      <div>{fmt_accuracy(result.get('row_missing_rate'))}</div>
      <div class="subtext">{rows_missing_text}</div>
    </div>
    <div class="card">
      <strong>Overall Missing-Rate</strong>
      <div>{fmt_accuracy(result.get('overall_missing_rate'))}</div>
      <div class="subtext">{overall_missing_text}</div>
    </div>
  </div>
  <div class="card">
    <h2>Record Tidak Ditemukan</h2>
    <h3>Di Ground Truth</h3>
    <ul>{render_missing(missing_gt)}</ul>
    <h3>Di OCR Results</h3>
    <ul>{render_missing(missing_ocr)}</ul>
  </div>
  <h2>Performa per Kolom</h2>
  <table>
    <thead>
      <tr>
        <th>Kolom</th>
        <th>Jumlah Dibandingkan</th>
        <th>Match</th>
        <th>Mismatch</th>
        <th>Akurasi</th>
        <th>Missing-Rate</th>
        <th>Detail Missmatch</th>
      </tr>
    </thead>
    <tbody>
      {column_table_html}
    </tbody>
  </table>
</body>
</html>
"""


def append_log_line(
    log_path: Path,
    accuracy: Optional[float],
    row_pass_rate: Optional[float],
    row_missing_rate: Optional[float],
    overall_missing_rate: Optional[float],
    total_gt: int,
    total_ocr: int,
) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    accuracy_text = f"{accuracy:.4f}" if accuracy is not None else "N/A"
    row_pass_text = f"{row_pass_rate:.4f}" if row_pass_rate is not None else "N/A"
    row_missing_text = f"{row_missing_rate:.4f}" if row_missing_rate is not None else "N/A"
    missing_text = f"{overall_missing_rate:.4f}" if overall_missing_rate is not None else "N/A"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(
            f"{timestamp} | accuracy: {accuracy_text} | row_pass_rate: {row_pass_text} | row_missing_rate: {row_missing_text} | overall_missing_rate: {missing_text} | total_groundtruth: {total_gt} | total_ocr: {total_ocr}\n"
        )


def main() -> None:
    args = parse_args()
    ocr_data, ocr_order, ocr_fields = load_csv(args.ocr, args.key)
    gt_data, gt_order, gt_fields = load_csv(args.ground_truth, args.key)

    columns = ordered_unique(gt_fields + ocr_fields)
    all_keys = ordered_unique(gt_order + ocr_order)

    stats = {
        column: {
            "total_compared": 0,
            "matches": 0,
            "mismatches": 0,
            "accuracy": None,
            "missing": 0,
            "missing_rate": None,
            "missmatch": [],
        }
        for column in columns
    }

    missing_in_ocr = []
    missing_in_ground_truth = []
    total_matches = 0
    total_compared = 0
    rows_compared = 0
    rows_passing = 0
    rows_missing = 0

    for key_value in all_keys:
        gt_row = gt_data.get(key_value)
        ocr_row = ocr_data.get(key_value)

        if gt_row is None:
            missing_in_ground_truth.append(key_value)
        if ocr_row is None:
            missing_in_ocr.append(key_value)

        if gt_row is None or ocr_row is None:
            continue

        row_all_match = True
        row_has_missing = False

        for column in columns:
            gt_value = normalize(gt_row.get(column, ""))
            ocr_value = normalize(ocr_row.get(column, ""))

            column_stat = stats[column]
            column_stat["total_compared"] += 1
            total_compared += 1

            if gt_value == ocr_value:
                column_stat["matches"] += 1
                total_matches += 1
            else:
                column_stat["mismatches"] += 1
                row_all_match = False
                if gt_value and not ocr_value:
                    column_stat["missing"] += 1
                    row_has_missing = True
                missmatch_entries = column_stat["missmatch"]
                if args.max_samples is None or len(missmatch_entries) < args.max_samples:
                    missmatch_entries.append(
                        {
                            args.key: key_value,
                            "expected": gt_value,
                            "ocr": ocr_value,
                        }
                    )

        if columns:
            rows_compared += 1
            if row_all_match:
                rows_passing += 1
            if row_has_missing:
                rows_missing += 1

    overall_accuracy = round(total_matches / total_compared, 4) if total_compared else None
    row_pass_rate = round(rows_passing / rows_compared, 4) if rows_compared else None
    row_missing_rate = round(rows_missing / rows_compared, 4) if rows_compared else None
    overall_missing = sum(column_stat["missing"] for column_stat in stats.values())
    overall_missing_rate = round(overall_missing / total_compared, 4) if total_compared else None

    for column, column_stat in stats.items():
        comparisons = column_stat["total_compared"]
        if comparisons > 0:
            column_stat["accuracy"] = round(column_stat["matches"] / comparisons, 4)
            column_stat["missing_rate"] = round(column_stat["missing"] / comparisons, 4)
        else:
            column_stat["accuracy"] = None
            column_stat["missing_rate"] = None

    result = {
        "key_column": args.key,
        "total_records_ground_truth": len(gt_data),
        "total_records_ocr": len(ocr_data),
        "shared_records": len(all_keys) - len(missing_in_ground_truth) - len(missing_in_ocr),
        "missing_in_ground_truth": missing_in_ground_truth,
        "missing_in_ocr_results": missing_in_ocr,
        "overall_accuracy": overall_accuracy,
        "rows_compared": rows_compared,
        "rows_passing": rows_passing,
        "rows_missing": rows_missing,
        "row_pass_rate": row_pass_rate,
        "row_missing_rate": row_missing_rate,
        "overall_missing": overall_missing,
        "overall_missing_rate": overall_missing_rate,
        "total_field_comparisons": total_compared,
        "columns": stats,
    }

    json_output = json.dumps(result, indent=2)
    generated_reports = []
    if args.json_out:
        args.json_out.write_text(json_output + "\n", encoding="utf-8")
        generated_reports.append(f"JSON: {args.json_out}")
    if args.html_out:
        html_content = render_html_report(result)
        args.html_out.write_text(html_content, encoding="utf-8")
        generated_reports.append(f"HTML: {args.html_out}")
    if args.log_file:
        append_log_line(
            args.log_file,
            result["overall_accuracy"],
            result["row_pass_rate"],
            result["row_missing_rate"],
            result["overall_missing_rate"],
            len(gt_data),
            len(ocr_data),
        )
    if generated_reports:
        print(
            "Validation complete. Reports saved -> "
            + ", ".join(generated_reports)
        )
    else:
        print("Validation complete. No report files were generated.")


if __name__ == "__main__":
    main()
