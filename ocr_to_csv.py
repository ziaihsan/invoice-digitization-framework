#!/usr/bin/env python3
"""
Batch OCR helper that scans images in a directory and writes the extracted text
and any detected key/value pairs to a CSV file.
"""

from __future__ import annotations

import argparse
import csv
import pathlib
import re
import string
import subprocess
import sys
import time
import unicodedata
from dataclasses import dataclass
from typing import Callable, Dict, Iterable, List, Sequence, Set, Tuple

try:
    from PIL import Image, ImageOps, ImageFilter
except ImportError as exc:  # pragma: no cover - import guard
    raise SystemExit(
        "Missing Pillow dependency. Install it with `pip install pillow`."
    ) from exc

try:
    import pytesseract
    from pytesseract import Output, TesseractError
except ImportError as exc:  # pragma: no cover - import guard
    raise SystemExit(
        "Missing pytesseract dependency. Install it with `pip install pytesseract`."
    ) from exc


SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
DEFAULT_OUTPUT = pathlib.Path("result") / "ocr_results.csv"
DEFAULT_KV_PATTERN = r"^\s*([\w\s./%()-]+?)\s*[:\-]\s*(.+)$"
DEFAULT_VALUE_COLUMNS = [
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
FIELD_ALIAS_OVERRIDES = {
    "tol": "tel",
    "tal": "tel",
    "telephone": "tel",
    "phone": "tel",
    "po": "po number",
    "po no": "po number",
    "po num": "po number",
    "po #": "po number",
    "sub total": "sub total",
    "subtotal": "sub total",
    "total amount": "total",
    "total usd": "total",
    "bank": "bank name",
    "branch": "branch name",
    "branch location": "branch name",
    "account number": "bank account number",
    "account no": "bank account number",
    "acct number": "bank account number",
    "swift": "bank swift code",
    "swift code": "bank swift code",
    "swift number": "bank swift code",
    "bankname": "bank name",
    "swiftcode": "bank swift code",
    "bank swiftcode": "bank swift code",
}
FIELD_PREFIX_OVERRIDES = {
    "discount": "discount",
    "sub total": "sub total",
    "tax": "tax",
    "total": "total",
    "bank name": "bank name",
    "branch name": "branch name",
    "bank account number": "bank account number",
    "account number": "bank account number",
    "bank swift code": "bank swift code",
    "swift code": "bank swift code",
    "bank swiftcode": "bank swift code",
}
EXCLUDED_LABEL_SUBSTRINGS = (
    "in words",
    "amount in words",
    "amount words",
)
EMAIL_EXCLUDE_PATTERNS = ("melvin40@example", "melvingo", "melvin")# Exclude static vendor email present in FATURA Template 1 to avoid extraction conflict
KNOWN_TLDS = {
    "com",
    "net",
    "org",
    "info",
    "co",
    "biz",
    "io",
    "id",
    "us",
    "uk",
    "me",
    "gov",
    "edu",
    "int",
    "app",
    "dev",
    "ai",
}
KNOWN_TLD_SUFFIXES = sorted(KNOWN_TLDS, key=len, reverse=True)
KNOWN_CURRENCY_CODES = {
    "USD",
    "EUR",
    "IDR",
    "SGD",
    "AUD",
    "CAD",
    "GBP",
    "JPY",
    "CNY",
    "CHF",
    "INR",
}
TAX_PERCENT_MISMATCH_THRESHOLD = 0.2
DATE_LINE_WHITELIST = "0123456789/-ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
EMAIL_TLD_PATTERN = "|".join(sorted(KNOWN_TLDS, key=len, reverse=True))
INLINE_PHONE_PATTERN = re.compile(
    r"(?:\b(?:tel|tol|tal|telephone|phone)\b)\s*[:\-]\s*([^\n\r|]+)",
    re.IGNORECASE,
)
EMAIL_PATTERN = re.compile(
    rf"[A-Za-z0-9._%+\-]+@(?:[A-Za-z0-9\-]+\.)+(?:{EMAIL_TLD_PATTERN})",
    re.IGNORECASE,
)
AMOUNT_PATTERN = re.compile(r"-?\d[\d,]*(?:\.\d+)?")
DISCOUNT_LEADING_JUNK = re.compile(r"^[\s\"'`~|.,;:_+()*-]+")
TEL_VALUE_SPLIT_PATTERN = re.compile(r"\s*\|\s*")
TEL_DIGIT_SEPARATOR_PATTERN = re.compile(r"(?<=\d)[\s./]+(?=\d)")
TEL_KEYWORD_SPLIT_PATTERN = re.compile(r"(?i)\b(?:fax|email|site)\b")
VAT_PERCENT_PATTERN = re.compile(r"vat\s*\(\s*([0-9]+(?:\.[0-9]+)?)\s*%\s*\)", re.IGNORECASE)
GENERIC_PERCENT_PATTERN = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*%")
INLINE_DISCOUNT_PATTERN = re.compile(
    r"(?i)\bdiscount[^\n\r:]*[:\-]\s*([^\n\r]+)"
)
INLINE_TAX_PATTERN = re.compile(
    r"(?i)\b(?:tax|vat)[a-z]*[^\n\r:]*[:\-]?\s*([^\n\r]+)"
)
FIELD_PATTERNS = {
    "Due Date": re.compile(
        r"due\s+date\s*(?:[:\-]|is)?\s*([^\n\r]+)", re.IGNORECASE
    ),
    "PO Number": re.compile(
        r"po\s*(?:number|no\.?|#)\s*(?:[:\-]|is)?\s*([^\n\r]+)", re.IGNORECASE
    ),
    "Tel": re.compile(
        r"(?:tel|tol|tal|telephone|phone)\s*(?:[:\-]|is)?\s*([^\n\r]+)",
        re.IGNORECASE,
    ),
    "Email": re.compile(
        r"email\s*(?:[:\-]|is)?\s*([^\n\r]+)", re.IGNORECASE
    ),
    "Site": re.compile(
        r"site\s*(?:[:\-]|is)?\s*([^\n\r]+)", re.IGNORECASE
    ),
    "SUB_TOTAL": re.compile(
        r"^\s*sub[_\s-]*total[^\n\r:]*[:\-]?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "DISCOUNT": re.compile(
        r"^\s*discount[^\n\r:]*[:\-]?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "TAX": re.compile(
        r"^\s*(?:tax|vat)[^\n\r:]*[:\-]\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "TOTAL": re.compile(
        r"^\s*total(?!\s+in\s+words)[^\n\r:]*[:\-]?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "Bank Name": re.compile(
        r"^\s*bank\s*name\s*(?:[:\-])?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "Branch Name": re.compile(
        r"^\s*branch\s+name\s*(?:[:\-])?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "Bank Account Number": re.compile(
        r"^\s*bank\s+account\s+number\s*(?:[:\-])?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
    "Bank Swift Code": re.compile(
        r"^\s*bank\s+swift\s*code\s*(?:[:\-])?\s*([^\n\r]+)",
        re.IGNORECASE | re.MULTILINE,
    ),
}
MONTH_LOOKUP = {
    "jan": "Jan",
    "january": "Jan",
    "feb": "Feb",
    "february": "Feb",
    "mar": "Mar",
    "march": "Mar",
    "apr": "Apr",
    "april": "Apr",
    "may": "May",
    "jun": "Jun",
    "june": "Jun",
    "jul": "Jul",
    "july": "Jul",
    "aug": "Aug",
    "august": "Aug",
    "sep": "Sep",
    "sept": "Sep",
    "september": "Sep",
    "oct": "Oct",
    "october": "Oct",
    "nov": "Nov",
    "november": "Nov",
    "dec": "Dec",
    "december": "Dec",
}
DUE_DATE_PATTERN = re.compile(
    r"(\d{1,2})[\s\-/.]+([A-Za-z]{3,})[\s\-/.]+(\d{2,4})"
)
DOMAIN_WITH_PATH_PATTERN = re.compile(
    r"((?:[a-z0-9][a-z0-9-]*\.)+[a-z]{2,})(/[a-z0-9./_\-]*)?",
    re.IGNORECASE,
)
SITE_VALUE_SPLIT_PATTERN = re.compile(r"[|,;\n]+")


def _detect_site_scheme(text: str) -> str:
    lowered = text.lower()
    if "https" in lowered:
        return "https"
    if "http" in lowered:
        return "http"
    return "https"


def _ensure_trailing_slash(url: str) -> str:
    if not url:
        return url
    return url if url.endswith("/") else f"{url}/"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run OCR on a directory of images and export the results to CSV."
    )
    parser.add_argument(
        "--images-dir",
        type=pathlib.Path,
        default=pathlib.Path("data/images"),
        help="Directory that contains images (default: data/images).",
    )
    parser.add_argument(
        "--lang",
        default="eng",
        help="Tesseract language to use (default: eng).",
    )
    parser.add_argument(
        "--kv-pattern",
        default=DEFAULT_KV_PATTERN,
        help=(
            "Regex used to detect key/value lines inside OCR text. "
            "Must contain two capture groups for key and value respectively."
        ),
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recurse into subdirectories when searching for images.",
    )
    parser.add_argument(
        "--value-columns",
        nargs="+",
        help=(
            "Column names to include in the output table (default: standard "
            "invoice-related fields)."
        ),
    )
    return parser.parse_args()


def run_ocr_pipeline(
    images_dir: pathlib.Path,
    *,
    lang: str = "eng",
    kv_pattern: str = DEFAULT_KV_PATTERN,
    recursive: bool = False,
    value_columns: Sequence[str] | None = None,
    output_path: pathlib.Path | None = None,
    progress_callback: Callable[[str], None] | None = None,
    row_callback: Callable[[Dict[str, str], Sequence[str]], None] | None = None,
) -> pathlib.Path:
    """
    Run the OCR workflow programmatically.

    Args:
        images_dir: Directory that contains invoice images.
        lang: Tesseract language code.
        kv_pattern: Regex pattern used for key/value extraction.
        recursive: Whether to recurse into subdirectories.
        value_columns: Optional list of custom columns for the CSV.
        output_path: Where to write the resulting CSV (defaults to DEFAULT_OUTPUT).
        progress_callback: Optional callable that receives progress messages.
        row_callback: Optional callable invoked for each table row before the CSV is written.

    Returns:
        Path to the generated CSV.
    """

    def _notify(message: str) -> None:
        print(message)
        if progress_callback is None:
            return
        try:
            progress_callback(message)
        except Exception as exc:  # pragma: no cover - best effort notifier
            print(f"Progress callback failed: {exc}", file=sys.stderr)

    image_paths = find_images(images_dir, recursive)
    total_images = len(image_paths)
    run_started = time.perf_counter()
    try:
        kv_regex = re.compile(kv_pattern)
    except re.error:
        # Re-raise to let callers decide how to report invalid regex values.
        raise

    column_names = value_columns or DEFAULT_VALUE_COLUMNS
    column_names = [name.strip() for name in column_names if name.strip()]
    if not column_names:
        column_names = DEFAULT_VALUE_COLUMNS

    rows: List[Dict[str, str]] = []
    output = output_path or DEFAULT_OUTPUT

    for index, image_path in enumerate(image_paths, start=1):
        elapsed_text = _format_elapsed(time.perf_counter() - run_started)
        progress_tag = f"[{index}/{total_images} | {elapsed_text}]"
        _notify(f"OCR: {image_path} {progress_tag}")
        image = preprocess_image(image_path)
        text = run_ocr(image, lang)
        ocr_lines = collect_ocr_lines(image, lang)
        fields = extract_key_values(text, kv_regex)
        targeted = apply_targeted_patterns(text)
        fields.update(targeted)
        normalized_fields: Dict[str, str] = {}
        for raw_key, value in fields.items():
            canonical_key = canonicalize_label(raw_key)
            if not canonical_key:
                continue
            normalized_fields[canonical_key] = value
        line_overrides, shipping_contacts = extract_line_overrides(
            image, ocr_lines, lang, normalized_fields
        )
        override_keys: Set[str] = set()
        for label, override_value in line_overrides.items():
            canonical_key = canonicalize_label(label)
            if not canonical_key:
                continue
            override_keys.add(canonical_key)
            if canonical_key == "site" and canonical_key in normalized_fields:
                combined = f"{normalized_fields[canonical_key]} | {override_value}"
                normalized_fields[canonical_key] = combined
            else:
                normalized_fields[canonical_key] = override_value
        for canonical_label, suppressed_values in shipping_contacts.items():
            if canonical_label in override_keys:
                continue
            existing_value = normalized_fields.get(canonical_label, "")
            cleaned_existing = _clean_contact_candidate(
                canonical_label, existing_value
            )
            if cleaned_existing:
                continue
            best_candidate = _select_best_contact_value(
                canonical_label, sorted(suppressed_values)
            )
            if best_candidate:
                normalized_fields[canonical_label] = best_candidate
        redistribute_inline_financial_values(normalized_fields)
        table_row = build_table_row(index, image_path.name, normalized_fields, column_names)
        rows.append(table_row)
        if row_callback is not None:
            try:
                row_callback(table_row, column_names)
            except Exception as exc:  # pragma: no cover - best effort notifier
                print(f"Row callback failed: {exc}", file=sys.stderr)

    write_csv(rows, output, column_names)
    final_elapsed = _format_elapsed(time.perf_counter() - run_started)
    _notify(
        f"Wrote {output} ({len(rows)} rows). [Total time: {final_elapsed}]"
    )
    return output


def find_images(root: pathlib.Path, recursive: bool) -> List[pathlib.Path]:
    if not root.exists():
        raise FileNotFoundError(f"Image directory not found: {root}")

    iterator: Iterable[pathlib.Path]
    iterator = root.rglob("*") if recursive else root.glob("*")

    images = [
        path
        for path in iterator
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    if not images:
        raise FileNotFoundError(
            f"No supported images were found under {root} "
            f"(extensions: {', '.join(sorted(SUPPORTED_EXTENSIONS))})."
        )
    return sorted(images)


def preprocess_image(image_path: pathlib.Path) -> Image.Image:
    image = Image.open(image_path)
    gray = ImageOps.autocontrast(image.convert("L"))
    sharpened = gray.filter(ImageFilter.SHARPEN)
    return sharpened


def run_ocr(image: Image.Image, lang: str) -> str:
    text = pytesseract.image_to_string(image, lang=lang)
    return text.strip()


def extract_key_values(text: str, pattern: re.Pattern[str]) -> Dict[str, str]:
    results: Dict[str, str] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip()
        line = line.lstrip("\"'“”‘’")
        if len(line) < 3:
            continue
        match = pattern.match(line)
        if not match:
            continue
        key = " ".join(match.group(1).split()).strip(" :.-")
        value = match.group(2).strip()
        if not key or not value:
            continue
        existing = results.get(key)
        if existing and value not in existing:
            results[key] = f"{existing} | {value}"
        else:
            results[key] = value
    return results


def normalize_key(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", label.lower()).strip()


def canonicalize_label(label: str) -> str:
    normalized = normalize_key(label)
    if not normalized:
        return normalized
    if any(excluded in normalized for excluded in EXCLUDED_LABEL_SUBSTRINGS):
        return ""
    normalized = FIELD_ALIAS_OVERRIDES.get(normalized, normalized)
    for prefix, target in FIELD_PREFIX_OVERRIDES.items():
        if normalized.startswith(prefix):
            return target
    return normalized


def clean_email_value(value: str) -> str:
    if not value:
        return value
    value = INLINE_PHONE_PATTERN.sub("", value)
    parts = re.split(r"[|,;/\n]+", value)
    filtered: List[str] = []
    for part in parts:
        candidate = part.strip()
        if not candidate:
            continue
        ascii_candidate = (
            unicodedata.normalize("NFKD", candidate)
            .encode("ascii", "ignore")
            .decode("ascii")
        )
        ascii_candidate = ascii_candidate.replace(" ", "")
        if not ascii_candidate:
            continue
        matches = EMAIL_PATTERN.findall(ascii_candidate)
        for match in matches:
            lowered = match.lower()
            if any(excluded in lowered for excluded in EMAIL_EXCLUDE_PATTERNS):
                continue
            if lowered not in filtered:
                filtered.append(lowered)
    return " | ".join(filtered)


def clean_currency_value(value: str, prefer_first: bool = False) -> str:
    if not value:
        return value
    normalized = unicodedata.normalize("NFKC", value)
    text = normalized.strip()
    if not text:
        return text
    has_usd_word = bool(re.search(r"\bUSD\b", text, re.IGNORECASE))
    has_eur_word = bool(
        re.search(r"\bEUR\b", text, re.IGNORECASE)
        or re.search(r"\bEUROS?\b", text, re.IGNORECASE)
    )
    had_dollar_symbol = any(symbol in value for symbol in ("$", "§"))
    had_euro_symbol = "€" in value
    text = text.replace("§", "$").replace("$US", "USD").replace("US$", "USD")
    text = text.replace("€", " EUR ")
    text = re.sub(r"\s+", " ", text)
    matches = AMOUNT_PATTERN.findall(text)
    if not matches:
        return text
    amount = matches[0] if prefer_first else matches[-1]
    amount = amount.replace(" ", "").replace(",,", ",")
    if has_eur_word:
        suffix = "EUR"
    elif has_usd_word:
        suffix = "USD"
    elif had_euro_symbol:
        suffix = "€"
    elif had_dollar_symbol:
        suffix = "$"
    else:
        suffix = "USD"
    return f"{amount} {suffix}"


def _extract_numeric_amount(value: str | None) -> float | None:
    if not value:
        return None
    match = AMOUNT_PATTERN.search(value)
    if not match:
        return None
    numeric = match.group(0).replace(",", "")
    try:
        return float(numeric)
    except ValueError:
        return None


def _infer_currency_suffix(*candidates: str | None) -> str:
    for candidate in candidates:
        if not candidate:
            continue
        if "$" in candidate:
            return "$"
        if "€" in candidate:
            return "€"
        for match in re.finditer(r"\b([A-Za-z]{3})\b", candidate):
            code = match.group(1).upper()
            if code in KNOWN_CURRENCY_CODES:
                return code
    return ""


def _format_currency_amount(amount_value: float, *hints: str | None) -> str:
    suffix = _infer_currency_suffix(*hints) or "USD"
    formatted = f"{amount_value:.2f}"
    return f"{formatted} {suffix}".strip()


def clean_due_date_value(value: str) -> str:
    if not value:
        return value
    text = unicodedata.normalize("NFKC", value).strip()
    if not text:
        return text
    text = text.strip(string.punctuation + " ")
    text = re.sub(r"\s+", " ", text)
    match = DUE_DATE_PATTERN.search(text)
    if not match:
        return text
    try:
        day = int(match.group(1))
    except ValueError:
        return text
    if not 1 <= day <= 31:
        return text
    month_token = match.group(2).lower()
    month = MONTH_LOOKUP.get(month_token)
    if not month and len(month_token) > 3:
        month = MONTH_LOOKUP.get(month_token[:3])
    if not month:
        month = month_token.title()
    year = match.group(3)
    if len(year) == 2:
        year_prefix = "20" if int(year) <= 30 else "19"
        year = f"{year_prefix}{year}"
    return f"{day:02d}-{month}-{year}"


def clean_po_number_value(value: str) -> str:
    if not value:
        return value
    cleaned = value.strip()
    cleaned = re.sub(r"[\s,.;:-]+$", "", cleaned)
    digits = re.findall(r"\d+", cleaned)
    if digits:
        return digits[-1]
    alnum = re.sub(r"[^A-Za-z0-9#/-]+", "", cleaned)
    if alnum:
        return alnum
    return cleaned


def clean_discount_value(value: str) -> str:
    if not value:
        return value
    text = value.strip()
    if not text:
        return ""
    if text.startswith("(-)"):
        return text.rstrip(",.;: ")
    text = (
        text.replace("(.)", " ")
        .replace("(•)", " ")
        .replace("()", " ")
        .replace("(,)", " ")
    )
    text = DISCOUNT_LEADING_JUNK.sub("", text)
    text = text.lstrip("+-").strip().rstrip(",.;: ")
    if not text:
        return "(-)"
    return f"(-) {text}"


def clean_branch_name(value: str) -> str:
    if not value:
        return value
    cleaned = value.strip().strip(" .,:;!?-_")
    return re.sub(r"\s{2,}", " ", cleaned)


def clean_bank_name(value: str) -> str:
    if not value:
        return value
    cleaned = value.strip().strip(" .,:;!?-_")
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    cleaned = re.sub(r"(?i)\bgentral\b", "Central", cleaned)
    cleaned = re.sub(r"(?i)centralbank", "Central Bank", cleaned)
    cleaned = re.sub(r"(?i)bankofeurope", "Bank of Europe", cleaned)
    cleaned = re.sub(r"(?i)bankofcalifornia", "Bank of California", cleaned)
    normalized = normalize_key(cleaned)
    if normalized.startswith("bank swift code") or normalized.startswith("bank swiftcode"):
        return ""
    if normalized.startswith("swift code") or normalized.startswith("swiftcode"):
        return ""
    return cleaned


def clean_bank_swift_code(value: str) -> str:
    if not value:
        return value
    cleaned = value.strip()
    cleaned = re.sub(r"^[^A-Za-z0-9]+", "", cleaned)
    cleaned = re.sub(r"[^A-Za-z0-9]+$", "", cleaned)
    alnum_only = re.sub(r"[^A-Za-z0-9]+", "", cleaned)
    return alnum_only.upper()


def _format_percent(value: str) -> str:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        cleaned = value.strip()
        return cleaned if cleaned.endswith("%") else f"{cleaned}%"
    return f"{numeric:.2f}%"


def clean_tel_value(value: str) -> str:
    if not value:
        return value
    parts = TEL_VALUE_SPLIT_PATTERN.split(value)
    cleaned_parts: List[str] = []
    for part in parts:
        candidate = part.strip()
        if not candidate:
            continue
        candidate = candidate.strip("\"'“”‘’")
        candidate = candidate.rstrip(",.;:!?")
        candidate = candidate.replace("–", "-").replace("—", "-")
        candidate = candidate.replace("{", "(").replace("[", "(")
        candidate = candidate.replace("}", ")").replace("]", ")")
        candidate = TEL_KEYWORD_SPLIT_PATTERN.split(candidate, maxsplit=1)[0]
        if "+" in candidate:
            candidate = candidate[candidate.index("+") :]
        else:
            paren_idx = candidate.find("(")
            if paren_idx != -1:
                candidate = "+" + candidate[paren_idx:]
        candidate = candidate.lstrip(": -")
        candidate = candidate.strip()
        digits_only = re.sub(r"\D", "", candidate)
        if len(digits_only) < 10:
            continue
        core_digits = digits_only[:10]
        area = core_digits[:3]
        middle = core_digits[3:6]
        last = core_digits[6:10]
        formatted = f"+({area}){middle}-{last}"
        cleaned_parts.append(formatted)
    return " | ".join(cleaned_parts)


def clean_tax_value(value: str, subtotal_value: str | None = None) -> str:
    if not value:
        return value
    text = unicodedata.normalize("NFKC", value).strip()
    if not text:
        return text
    value_segment = text
    colon_idx = text.find(":")
    if colon_idx != -1 and colon_idx + 1 < len(text):
        value_segment = text[colon_idx + 1 :]
    else:
        dash_idx = text.find("-")
        if dash_idx != -1 and dash_idx + 1 < len(text):
            value_segment = text[dash_idx + 1 :]
    amount_text = clean_currency_value(value_segment, prefer_first=True)
    tax_amount = _extract_numeric_amount(amount_text)
    if tax_amount is not None and tax_amount <= 0:
        tax_amount = None
    subtotal_clean = clean_currency_value(subtotal_value) if subtotal_value else ""
    subtotal_amount = _extract_numeric_amount(subtotal_clean)

    def _is_plausible(candidate: float | None) -> bool:
        if candidate is None:
            return False
        if candidate <= 0:
            return False
        if candidate > 5000:
            return False
        if (
            subtotal_amount
            and subtotal_amount > 0
            and candidate > subtotal_amount * 0.3
        ):
            return False
        return True

    if not _is_plausible(tax_amount):
        tax_amount = None

    percent_value: float | None = None
    percent_label: str = ""
    match = VAT_PERCENT_PATTERN.search(text)
    if match:
        try:
            percent_value = float(match.group(1))
        except (TypeError, ValueError):
            percent_value = None
        percent_label = _format_percent(match.group(1))
    else:
        alt = GENERIC_PERCENT_PATTERN.search(text)
        if alt:
            try:
                percent_value = float(alt.group(1))
            except (TypeError, ValueError):
                percent_value = None
            percent_label = _format_percent(alt.group(1))

    computed_from_percent: float | None = None
    if (
        percent_value is not None
        and subtotal_amount
        and subtotal_amount > 0
    ):
        computed = round((subtotal_amount * percent_value) / 100.0, 2)
        if _is_plausible(computed):
            computed_from_percent = computed

    amount_choice = tax_amount
    if computed_from_percent is not None:
        use_computed = amount_choice is None
        if not use_computed and subtotal_amount:
            ratio_percent = (amount_choice / subtotal_amount) * 100
            if (
                percent_value is not None
                and abs(ratio_percent - percent_value)
                > TAX_PERCENT_MISMATCH_THRESHOLD
            ):
                use_computed = True
        if use_computed:
            amount_choice = computed_from_percent

    if amount_choice is None:
        return ""

    amount_text = _format_currency_amount(
        amount_choice, amount_text, value_segment, subtotal_value, subtotal_clean
    )
    if not percent_label:
        if subtotal_amount and subtotal_amount > 0:
            computed_percent = (amount_choice / subtotal_amount) * 100
            percent_label = _format_percent(computed_percent)
        else:
            percent_label = "0.00%"
    label = f"VAT ({percent_label})"
    return f"{label}: {amount_text}"


def _extract_inline_field_values(
    value: str, pattern: re.Pattern[str]
) -> Tuple[str, List[str]]:
    if not value:
        return value, []
    remainder = value
    extracted: List[str] = []
    while True:
        match = pattern.search(remainder)
        if not match:
            break
        captured = match.group(1).strip(" .,:;|-")
        if captured:
            extracted.append(captured)
        before = remainder[: match.start()].rstrip(" ,.;|-")
        after = remainder[match.end() :].lstrip(" ,.;|-")
        if before and after:
            remainder = f"{before} {after}"
        else:
            remainder = before or after
        remainder = remainder.strip()
    return remainder.strip(" ,.;|-"), extracted


def redistribute_inline_financial_values(fields: Dict[str, str]) -> None:
    inline_sources = (
        "bank name",
        "branch name",
        "bank account number",
        "bank swift code",
    )
    for source_key in inline_sources:
        value = fields.get(source_key)
        if not value:
            continue
        cleaned_value, discount_values = _extract_inline_field_values(
            value, INLINE_DISCOUNT_PATTERN
        )
        existing_discount = fields.get("discount")
        filtered_discount = [
            candidate
            for candidate in discount_values
            if AMOUNT_PATTERN.search(candidate)
        ]
        if filtered_discount and not (existing_discount and existing_discount.strip()):
            fields["discount"] = " | ".join(filtered_discount)
        cleaned_value, tax_values = _extract_inline_field_values(
            cleaned_value, INLINE_TAX_PATTERN
        )
        existing_tax = fields.get("tax")
        filtered_tax = [
            candidate for candidate in tax_values if AMOUNT_PATTERN.search(candidate)
        ]
        if filtered_tax and not (existing_tax and existing_tax.strip()):
            fields["tax"] = " | ".join(filtered_tax)
        fields[source_key] = cleaned_value.strip(" ,.;|-")


def extract_bank_name_from_lines(text: str) -> str:
    lines = text.splitlines()
    stop_prefixes = (
        "branch name",
        "bank account number",
        "bank swift code",
        "bank swiftcode",
        "swift code",
        "swiftcode",
        "note",
        "sub total",
        "discount",
        "tax",
        "total",
    )
    for idx, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"(?i)bank\s*name\s*(?:[:\-])?\s*(.*)", line)
        if not match:
            continue
        remainder = match.group(1).strip(" .,:;")
        if remainder:
            return remainder
        for candidate in lines[idx + 1 :]:
            cleaned = candidate.strip()
            if not cleaned:
                continue
            normalized = normalize_key(cleaned)
            if not normalized:
                continue
            if any(normalized.startswith(prefix) for prefix in stop_prefixes):
                continue
            return cleaned.strip(" .,:;")
    return ""


def _looks_like_www_token(token: str) -> bool:
    normalized = token.lower()
    if normalized == "www":
        return True
    return bool(normalized) and len(normalized) >= 3 and set(normalized) <= {"w", "v"}


def _normalize_tld(tld: str) -> str:
    letters_only = re.sub(r"[^a-z]", "", tld.lower())
    if not letters_only:
        return tld.lower()
    for candidate in KNOWN_TLD_SUFFIXES:
        if letters_only.startswith(candidate):
            return candidate
    return letters_only


def _normalize_www_prefix(prefix: str) -> str:
    lowered = prefix.lower()
    wv_count = sum(1 for ch in lowered if ch in {"w", "v"})
    other_count = len(lowered) - wv_count
    if wv_count >= 3 and other_count <= 2:
        return "www"
    return lowered


def _normalize_domain_host(domain: str) -> str:
    lowered = domain.lower().strip(".")
    if not lowered:
        return ""
    parts = [part for part in lowered.split(".") if part]
    if len(parts) < 2:
        return lowered
    parts[0] = _normalize_www_prefix(parts[0])
    normalized_tld = _normalize_tld(parts[-1])
    parts[-1] = normalized_tld or parts[-1]
    return ".".join(parts)


def _compact_site_text(text: str) -> str:
    if not text:
        return text
    normalized = text
    normalized = normalized.replace("：", ":").replace("／", "/")
    lowered = normalized.lower()
    http_idx = lowered.find("http")
    if http_idx != -1:
        prefix = normalized[:http_idx]
        fragment = re.sub(r"\s+", "", normalized[http_idx:])
        normalized = prefix + fragment
    else:
        www_idx = lowered.find("www")
        if www_idx != -1:
            prefix = normalized[:www_idx]
            fragment = re.sub(r"\s+", "", normalized[www_idx:])
            normalized = prefix + fragment
    normalized = re.sub(r"(https?:)/{1}(?=[^/])", r"\1//", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"(https?:)//{3,}", r"\1//", normalized, flags=re.IGNORECASE)
    return normalized


def _postprocess_domain_host(host: str) -> str:
    if not host:
        return host
    lowered = host.lower()
    if "www." in lowered and not lowered.startswith("www."):
        idx = lowered.find("www.")
        host = host[idx:]
        lowered = host.lower()
    if lowered.startswith("http"):
        trimmed = re.sub(r"^https?", "", lowered)
        host = trimmed.lstrip("./:")
        lowered = host.lower()
    if lowered.startswith("ww") and not lowered.startswith("www."):
        idx = 0
        while idx < len(host) and lowered[idx] in {"w", "v"}:
            idx += 1
        remainder = host[idx:].lstrip(".")
        if idx >= 2 and remainder:
            host = f"www.{remainder}"
    host = re.sub(r"\.{2,}", ".", host)
    return host


def _site_quality_score(url: str) -> tuple[int, int, int, int]:
    if not url:
        return (0, 0, 0, 0)
    normalized = url.strip()
    lowered = normalized.lower()
    has_scheme = 1 if lowered.startswith(("http://", "https://")) else 0
    host_part = lowered.split("://", 1)[1] if has_scheme else lowered
    host_part = host_part.strip("/")
    if not host_part:
        return (has_scheme, 0, 0, 0)
    host_only = host_part.split("/", 1)[0]
    has_tld = 1 if any(host_only.endswith(f".{tld}") for tld in KNOWN_TLDS) else 0
    has_www = 1 if host_only.startswith("www.") else 0
    contains_space = 1 if any(ch.isspace() for ch in normalized) else 0
    penalty_chars = sum(1 for ch in normalized if ch in {"!", "|", ",", ";", "?"})
    host_length = len(host_only)
    return (
        has_scheme + has_tld * 2 + has_www,
        -contains_space,
        -penalty_chars,
        -host_length,
    )


def _domain_from_token_text(text: str) -> str:
    tokens = [
        re.sub(r"[^a-z0-9-]", "", part.lower())
        for part in re.split(r"[\s/]+", text)
        if part.strip()
    ]
    for idx, token in enumerate(tokens):
        if _looks_like_www_token(token):
            remainder = tokens[idx:]
            break
    else:
        return ""
    if len(remainder) < 2 or remainder[-1] not in KNOWN_TLDS:
        return ""
    remainder[0] = "www"
    if any(not segment for segment in remainder):
        return ""
    return ".".join(remainder)


def _normalize_site_candidate(candidate: str) -> str:
    text = candidate.strip()
    if not text:
        return ""
    text = _compact_site_text(text)
    scheme = _detect_site_scheme(text)
    text = text.replace("\\", "/")
    text = (
        text.replace("—", "-")
        .replace("–", "-")
        .replace("…", ".")
        .replace("“", "")
        .replace("”", "")
        .replace("‘", "")
        .replace("’", "")
    )
    lower_text = text.lower()
    http_index = lower_text.find("http")
    if http_index != -1:
        text = text[http_index:]
    else:
        www_index = lower_text.find("www")
        if www_index != -1:
            text = text[www_index:]
            text = "http://" + text
    match = DOMAIN_WITH_PATH_PATTERN.search(text)
    if match:
        domain = _postprocess_domain_host(_normalize_domain_host(match.group(1)))
        if domain:
            path = (match.group(2) or "").rstrip(".,!?:;")
            url = f"{scheme}://{domain}{path}"
            return _ensure_trailing_slash(url)
    domain_from_tokens = _domain_from_token_text(text)
    if domain_from_tokens:
        normalized_domain = _normalize_domain_host(domain_from_tokens)
        if normalized_domain:
            return _ensure_trailing_slash(f"{scheme}://{normalized_domain}")
    loose_match = re.search(
        r"(www[a-z0-9\-]*\.[a-z]{2,}|[a-z0-9\-]+\.(?:com|net|org|info|biz|co|io|id|us|uk|me|gov|edu|int|app|dev|ai))",
        text,
        re.IGNORECASE,
    )
    if loose_match:
        domain = loose_match.group(1)
        domain = domain.replace("www", "www.", 1) if domain.startswith("www") and "." not in domain[:4] else domain
        domain = _postprocess_domain_host(domain)
        if domain:
            return _ensure_trailing_slash(f"{scheme}://{domain}")
    return ""


def clean_site_value(value: str) -> str:
    if not value:
        return value
    candidates = SITE_VALUE_SPLIT_PATTERN.split(value)
    normalized_values: List[str] = []
    for candidate in candidates:
        normalized = _normalize_site_candidate(candidate)
        if normalized and normalized not in normalized_values:
            normalized_values.append(normalized)
    if len(normalized_values) <= 1:
        return normalized_values[0] if normalized_values else ""
    normalized_values.sort(key=_site_quality_score, reverse=True)
    return normalized_values[0]


@dataclass
class OcrLine:
    text: str
    bbox: Tuple[int, int, int, int]


LINE_FIELD_CONFIGS = {
    "Tel": {
        "keywords": ("tel", "tol", "tal", "telephone", "phone"),
        "config": "--psm 7 -c tessedit_char_whitelist=:+0123456789()- Tel",
        "stop_keywords": ("email", "site"),
        "match_anywhere": True,
    },
    "Email": {
        "keywords": ("email",),
        "config": "--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@._+-:",
        "stop_keywords": ("site",),
    },
    "Site": {
        "keywords": ("site",),
        "config": "--psm 7 -c tessedit_char_whitelist=:/.-ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
    },
    "Due Date": {
        "keywords": ("due date",),
        "strip_keywords": ("due date", "due-date", "duedate"),
        "stop_keywords": ("terms", "po number", "invoice number"),
        "config": f"--psm 7 -c tessedit_char_whitelist={DATE_LINE_WHITELIST}",
    },
}

CONTACT_LINE_LABELS = {"Tel", "Email", "Site"}
CUSTOMER_SECTION_KEYWORDS = (
    "ship to",
    "shipto",
    "bill to",
    "billto",
    "deliver to",
    "deliverto",
)


def _collect_customer_section_centers(lines: Sequence[OcrLine]) -> List[float]:
    centers: List[float] = []
    for line in lines:
        normalized = normalize_key(line.text)
        if any(normalized.startswith(keyword) for keyword in CUSTOMER_SECTION_KEYWORDS):
            _, top, _, bottom = line.bbox
            centers.append((top + bottom) / 2.0)
    return centers


def _is_near_customer_section(
    bbox: Tuple[int, int, int, int],
    section_centers: Sequence[float],
    image_height: int,
) -> bool:
    if not section_centers or image_height <= 0:
        return False
    _, top, _, bottom = bbox
    center = (top + bottom) / 2.0
    max_distance = max(120, int(image_height * 0.22))
    for marker in section_centers:
        if marker < center <= marker + max_distance:
            return True
    return False


def _clean_contact_candidate(canonical_label: str, value: str) -> str:
    if not value:
        return ""
    if canonical_label == "tel":
        return clean_tel_value(value)
    if canonical_label == "email":
        return clean_email_value(value)
    if canonical_label == "site":
        return clean_site_value(value)
    return value.strip()


def _select_best_contact_value(
    canonical_label: str, candidates: Sequence[str]
) -> str:
    if not candidates:
        return ""
    if canonical_label == "tel":
        return max(
            candidates,
            key=lambda candidate: (
                len(re.sub(r"\D", "", candidate)),
                -len(candidate),
            ),
        )
    if canonical_label == "site":
        return max(candidates, key=_site_quality_score)
    if canonical_label == "email":
        return sorted(
            candidates, key=lambda candidate: (-len(candidate), candidate.lower())
        )[0]
    return sorted(candidates)[0]


def _site_scheme_variants(url: str) -> List[str]:
    if url.startswith("http://"):
        return ["https://" + url[len("http://") :]]
    if url.startswith("https://"):
        return ["http://" + url[len("https://") :]]
    return []


def collect_ocr_lines(image: Image.Image, lang: str) -> List[OcrLine]:
    try:
        data = pytesseract.image_to_data(image, lang=lang, output_type=Output.DICT)
    except (TesseractError, OSError):
        return []
    grouped: Dict[Tuple[int, int, int, int], List[Tuple[int, int, int, int, str]]] = {}
    for idx, raw_text in enumerate(data.get("text", [])):
        cleaned = raw_text.strip()
        if not cleaned:
            continue
        key = (
            data["page_num"][idx],
            data["block_num"][idx],
            data["par_num"][idx],
            data["line_num"][idx],
        )
        left = int(data["left"][idx])
        top = int(data["top"][idx])
        width = int(data["width"][idx])
        height = int(data["height"][idx])
        right = left + width
        bottom = top + height
        grouped.setdefault(key, []).append((left, top, right, bottom, raw_text))
    lines: List[OcrLine] = []
    for tokens in grouped.values():
        tokens.sort(key=lambda item: (item[1], item[0]))
        left = min(token[0] for token in tokens)
        top = min(token[1] for token in tokens)
        right = max(token[2] for token in tokens)
        bottom = max(token[3] for token in tokens)
        text = " ".join(token[4] for token in tokens)
        lines.append(OcrLine(text=text, bbox=(left, top, right, bottom)))
    lines.sort(key=lambda line: (line.bbox[1], line.bbox[0]))
    return lines


def _line_matches_keywords(
    line_text: str,
    keywords: Sequence[str],
    exclude_prefixes: Sequence[str],
    match_anywhere: bool = False,
) -> bool:
    normalized_line = normalize_key(line_text)
    normalized_exclusions = [normalize_key(item) for item in exclude_prefixes]
    if any(normalized_line.startswith(prefix) for prefix in normalized_exclusions):
        return False
    normalized_keywords = [
        normalize_key(keyword) for keyword in keywords if normalize_key(keyword)
    ]
    if match_anywhere:
        tokens = normalized_line.split()
        return any(
            token.startswith(keyword)
            for token in tokens
            for keyword in normalized_keywords
        )
    return any(normalized_line.startswith(keyword) for keyword in normalized_keywords)


def _strip_label_from_text(text: str, keywords: Sequence[str]) -> str:
    if not text:
        return text
    for keyword in sorted(keywords, key=len, reverse=True):
        if not keyword:
            continue
        escaped = re.escape(keyword)
        pattern = re.compile(
            rf"(?i)(?:^|[^A-Za-z0-9])({escaped})(?:$|[^A-Za-z0-9])"
        )
        match = pattern.search(text)
        if match:
            remainder = text[match.end(1) :].lstrip(" :.-")
            if remainder:
                return remainder
    colon_idx = text.find(":")
    if colon_idx != -1 and colon_idx + 1 < len(text):
        remainder = text[colon_idx + 1 :].lstrip(" -")
        if remainder:
            return remainder
    lowered = text.lower()
    for keyword in sorted(keywords, key=len, reverse=True):
        key_lower = keyword.lower()
        if not key_lower:
            continue
        idx = lowered.find(key_lower)
        if idx != -1:
            return text[idx + len(key_lower) :].lstrip(" -")
    return text


def _truncate_at_keywords(text: str, stop_keywords: Sequence[str]) -> str:
    if not text or not stop_keywords:
        return text
    lowered = text.lower()
    indices = [
        lowered.find(keyword.lower())
        for keyword in stop_keywords
        if keyword and lowered.find(keyword.lower()) != -1
    ]
    if not indices:
        return text
    cutoff = min(index for index in indices if index >= 0)
    return text[:cutoff].rstrip()


def _is_valid_override(label: str, value: str) -> bool:
    if not value:
        return False
    canonical_label = canonicalize_label(label)
    if label == "Tel":
        digits = re.sub(r"\D", "", clean_tel_value(value))
        return len(digits) >= 7
    if label == "Email":
        return bool(clean_email_value(value))
    if label == "Site":
        return bool(clean_site_value(value))
    if label == "Bank Name":
        cleaned = clean_bank_name(value)
        return len(cleaned) >= 4
    if canonical_label == "due date":
        text = unicodedata.normalize("NFKC", value)
        return bool(DUE_DATE_PATTERN.search(text))
    if canonical_label == "po number":
        return bool(re.search(r"\d+", value))
    if canonical_label in {"sub total", "discount", "tax", "total"}:
        return bool(AMOUNT_PATTERN.search(value))
    return True


def _should_override(
    label: str, value: str, existing_fields: Dict[str, str] | None
) -> bool:
    if not existing_fields:
        return True
    canonical_label = canonicalize_label(label)
    existing_value = existing_fields.get(canonical_label, "")
    if not existing_value:
        return True
    if label == "Tel":
        new_digits = re.sub(r"\D", "", clean_tel_value(value))
        old_digits = re.sub(r"\D", "", clean_tel_value(existing_value))
        if not old_digits:
            return bool(new_digits)
        if len(new_digits) <= len(old_digits):
            return False
        return True
    if label == "Email":
        new_clean = clean_email_value(value)
        old_clean = clean_email_value(existing_value)
        return bool(new_clean) and new_clean != old_clean
    if label == "Site":
        return bool(clean_site_value(value))
    if label == "Bank Name":
        new_clean = clean_bank_name(value)
        old_clean = clean_bank_name(existing_value)
        if len(new_clean) < len(old_clean):
            return False
        if new_clean == old_clean:
            return False
        new_alpha = sum(char.isalpha() for char in new_clean)
        old_alpha = sum(char.isalpha() for char in old_clean)
        return new_alpha >= old_alpha
    return True


def _rerun_line_ocr(
    image: Image.Image, bbox: Tuple[int, int, int, int], lang: str, config: str | None
) -> str:
    left, top, right, bottom = bbox
    pad = 4
    left = max(left - pad, 0)
    top = max(top - pad, 0)
    right = min(image.width, right + pad)
    bottom = min(image.height, bottom + pad)
    if right <= left or bottom <= top:
        return ""
    crop = image.crop((left, top, right, bottom))
    if crop.width <= 0 or crop.height <= 0:
        return ""
    resized = crop.resize((crop.width * 2, crop.height * 2), Image.BICUBIC)
    try:
        text = pytesseract.image_to_string(
            resized, lang=lang, config=config or "--psm 7"
        )
    except (TesseractError, OSError):
        return ""
    return " ".join(text.strip().split())


def extract_line_overrides(
    image: Image.Image,
    lines: Sequence[OcrLine],
    lang: str,
    existing_fields: Dict[str, str] | None = None,
) -> Tuple[Dict[str, str], Dict[str, Set[str]]]:
    overrides: Dict[str, str] = {}
    shipping_contacts: Dict[str, Set[str]] = {}
    section_centers = _collect_customer_section_centers(lines)
    for label, config in LINE_FIELD_CONFIGS.items():
        keywords = config["keywords"]
        canonical_label = canonicalize_label(label)
        exclude_prefixes = config.get("exclude_prefixes", ())
        match_anywhere = config.get("match_anywhere", False)
        for line in lines:
            if not _line_matches_keywords(
                line.text, keywords, exclude_prefixes, match_anywhere=match_anywhere
            ):
                continue
            near_customer_section = label in CONTACT_LINE_LABELS and _is_near_customer_section(
                line.bbox, section_centers, image.height
            )
            if near_customer_section and canonical_label:
                refined_shipping = _rerun_line_ocr(
                    image, line.bbox, lang, config.get("config")
                )
                candidate_sources = [text for text in (refined_shipping, line.text) if text]
                for candidate_shipping in candidate_sources:
                    shipping_value = _strip_label_from_text(
                        candidate_shipping, config.get("strip_keywords", keywords)
                    )
                    shipping_value = _truncate_at_keywords(
                        shipping_value, config.get("stop_keywords", ())
                    ).strip()
                    cleaned_shipping = _clean_contact_candidate(
                        canonical_label, shipping_value
                    )
                    if not cleaned_shipping:
                        continue
                    values = shipping_contacts.setdefault(canonical_label, set())
                    values.add(cleaned_shipping)
                    if canonical_label == "site":
                        for variant in _site_scheme_variants(cleaned_shipping):
                            values.add(variant)
                continue
            refined = _rerun_line_ocr(image, line.bbox, lang, config.get("config"))
            candidate = refined or line.text
            value = _strip_label_from_text(
                candidate, config.get("strip_keywords", keywords)
            )
            value = _truncate_at_keywords(value, config.get("stop_keywords", ()))
            value = value.strip()
            if not _is_valid_override(label, value):
                continue
            if not _should_override(label, value, existing_fields):
                continue
            overrides[label] = value
            break
    return overrides, shipping_contacts


def apply_targeted_patterns(text: str) -> Dict[str, str]:
    sanitized = (
        text.replace("“", "")
        .replace("”", "")
        .replace("‘", "")
        .replace("’", "")
    )
    results: Dict[str, str] = {}
    for label, pattern in FIELD_PATTERNS.items():
        matches = pattern.findall(sanitized)
        if not matches:
            continue
        if label == "Email":
            cleaned_matches = []
            for candidate in matches:
                cleaned = clean_email_value(candidate)
                if cleaned:
                    cleaned_matches.append(cleaned)
            if cleaned_matches:
                deduped = list(dict.fromkeys(cleaned_matches))
                results[label] = " | ".join(deduped)
        else:
            value = matches[0].strip()
            if value:
                if label == "Bank Name":
                    normalized_value = normalize_key(value)
                    if normalized_value.startswith("branch name"):
                        continue
                results[label] = value
    if "Bank Name" not in results:
        fallback_bank = extract_bank_name_from_lines(sanitized)
        if fallback_bank:
            results["Bank Name"] = fallback_bank
    return results


def build_table_row(
    index: int, filename: str, fields: Dict[str, str], column_names: Sequence[str]
) -> Dict[str, str]:
    normalized_fields: Dict[str, str] = {}
    for raw_key, value in fields.items():
        canonical_key = canonicalize_label(raw_key)
        if not canonical_key:
            continue
        normalized_fields[canonical_key] = value
    row: Dict[str, str] = {"Nomor": str(index), "Nama File": filename}
    currency_keys = {"sub total", "total"}
    for column in column_names:
        key = normalize_key(column)
        value = normalized_fields.get(key, "")
        if key == "email":
            value = clean_email_value(value)
        elif key == "site":
            value = clean_site_value(value)
        elif key == "due date":
            value = clean_due_date_value(value)
        elif key == "po number":
            value = clean_po_number_value(value)
        elif key in currency_keys:
            value = clean_currency_value(value)
        elif key == "tax":
            subtotal_raw = normalized_fields.get("sub total", "")
            value = clean_tax_value(value, subtotal_raw)
        if key == "tel":
            value = clean_tel_value(value)
        if key == "discount":
            value = clean_discount_value(value)
        if key == "branch name":
            value = clean_branch_name(value)
        if key == "bank name":
            value = clean_bank_name(value)
        if key == "bank swift code":
            value = clean_bank_swift_code(value)
        row[column] = value
    return row


def write_csv(
    rows: Iterable[Dict[str, str]],
    output_path: pathlib.Path,
    column_names: Sequence[str],
) -> None:
    fieldnames = ["Nomor", "Nama File"] + list(column_names)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def _format_elapsed(seconds: float) -> str:
    total_seconds = int(seconds)
    hours, remainder = divmod(total_seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def run_validation_script(output_path: pathlib.Path) -> int:
    if not output_path.exists():
        print(
            f"Cannot find OCR output at {output_path}; skipping validation.",
            file=sys.stderr,
        )
        return 1

    validator_path = (
        pathlib.Path(__file__).resolve().parent
        / "result"
        / "validate"
        / "validate_ocr.py"
    )
    if not validator_path.exists():
        print(
            f"Validator script not found at {validator_path}; skipping validation.",
            file=sys.stderr,
        )
        return 1

    python_executable = sys.executable or "python3"
    print(f"Running validator: {validator_path}")
    try:
        result = subprocess.run(
            [python_executable, str(validator_path), "--ocr", str(output_path)],
            check=False,
        )
    except OSError as exc:
        print(f"Failed to launch validator: {exc}", file=sys.stderr)
        return 1

    if result.returncode != 0:
        print(
            f"Validator exited with status {result.returncode}. See validation logs for details.",
            file=sys.stderr,
        )
    return result.returncode


def main() -> int:
    args = parse_args()
    try:
        output_path = run_ocr_pipeline(
            args.images_dir,
            lang=args.lang,
            kv_pattern=args.kv_pattern,
            recursive=args.recursive,
            value_columns=args.value_columns,
            output_path=DEFAULT_OUTPUT,
        )
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        return 1
    except re.error as exc:
        print(f"Invalid --kv-pattern: {exc}", file=sys.stderr)
        return 2
    return run_validation_script(output_path)


if __name__ == "__main__":
    raise SystemExit(main())
