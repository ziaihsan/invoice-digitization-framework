#!/usr/bin/env python3
"""
xlwings bridge so the OCR pipeline can be triggered directly from Excel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import pathlib
import sys
import time
from collections import deque
from dataclasses import dataclass
from typing import Deque, Dict, Iterable, List, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - import guard
    raise ImportError(
        "xlwings is not installed. Install it with `pip install xlwings`."
    ) from exc

try:  # pragma: no cover - optional macOS dependency
    from appscript.reference import CommandError as AppscriptCommandError
    from aem.aemsend import EventError as AppscriptEventError
except Exception:  # pragma: no cover - windows/linux fallback
    class AppscriptCommandError(Exception):
        """Fallback CommandError placeholder on non-macOS hosts."""


    class AppscriptEventError(Exception):
        """Fallback EventError placeholder on non-macOS hosts."""


APPLESCRIPT_ERRORS: Tuple[type[BaseException], ...] = (
    AppscriptCommandError,
    AppscriptEventError,
)

from ocr_to_csv import (
    DEFAULT_OUTPUT,
    SUPPORTED_EXTENSIONS,
    run_ocr_pipeline,
    run_validation_script,
)

PROJECT_ROOT = pathlib.Path(__file__).resolve().parent
DEFAULT_IMAGES_DIR = PROJECT_ROOT / "data" / "images"
RESULT_DIR = PROJECT_ROOT / "result"
LOG_PATH = RESULT_DIR / "excel_rpa.log"
LOG_SEPARATOR = "=" * 60
CONTROL_SHEET_NAME = "RPA Control"
RESULTS_SHEET_NAME = "OCR Results"
IMAGES_DIR_CELL = "B2"
LANG_CELL = "B3"
RECURSIVE_CELL = "B4"
VALUE_COLUMNS_CELL = "B5"
STATUS_CELL = "B7"
LAST_RUN_CELL = "B8"


class _BufferingLogHandler(logging.Handler):
    """Collect log records for a single run so they can be prepended later."""

    def __init__(self, target: List[str]) -> None:
        super().__init__()
        self.target = target
        self.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        )

    def emit(self, record: logging.LogRecord) -> None:  # pragma: no cover - log plumbing
        try:
            message = self.format(record)
        except Exception:
            message = record.getMessage()
        self.target.append(message)


def _configure_logger() -> logging.Logger:
    logger = logging.getLogger("excel_rpa")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    logger.propagate = False
    return logger


def _prepend_run_log(lines: Sequence[str]) -> None:
    if not lines:
        return
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    block = "\n".join(lines).rstrip() + "\n"
    if LOG_PATH.exists():
        previous = LOG_PATH.read_text(encoding="utf-8")
        content = f"{block}{LOG_SEPARATOR}\n{previous}"
    else:
        content = block
    LOG_PATH.write_text(content, encoding="utf-8")

def _log_run_summary(
    logger: logging.Logger,
    *,
    total_ms: float,
    row_stats: Tuple[float, float, float] | None,
    rows_written: int | None,
    csv_path: pathlib.Path | None,
) -> None:
    def _fmt(ms: float | None) -> str:
        if ms is None:
            return "N/A"
        seconds = ms / 1000
        return f"{ms:.0f} ms ({seconds:.2f}s)"

    average_ms: float | None = None
    min_ms: float | None = None
    max_ms: float | None = None
    if row_stats:
        average_ms, min_ms, max_ms = row_stats

    summary_lines = [
        "Run summary:",
        f"  Total duration : {_fmt(total_ms)}",
        f"  Avg OCR->Excel : {_fmt(average_ms)}",
        f"  Fastest row    : {_fmt(min_ms)}",
        f"  Slowest row    : {_fmt(max_ms)}",
        f"  Rows processed : {rows_written if rows_written is not None else 'N/A'}",
        f"  Output CSV     : {csv_path or '-'}",
        f"  Log location   : {LOG_PATH}",
    ]
    for line in summary_lines:
        logger.info(line)


def _run_validation(output_path: pathlib.Path, logger: logging.Logger) -> None:
    logger.info("Starting validation script for %s", output_path)
    try:
        exit_code = run_validation_script(output_path)
    except Exception as exc:  # pragma: no cover - subprocess guard
        logger.exception("Validation script failed to execute: %s", exc)
        return
    if exit_code == 0:
        logger.info("Validation completed successfully.")
    else:
        logger.warning(
            "Validation finished with exit code %s. Check result/validate for details.",
            exit_code,
        )


LOGGER = _configure_logger()


@dataclass
class ExcelRPAConfig:
    images_dir: pathlib.Path
    lang: str
    recursive: bool
    value_columns: Sequence[str] | None


@dataclass
class WriteResult:
    rows_written: int
    sheet_name: str
    start_address: str
    target_mode: str


class ExcelStreamWriter:
    """Stream table rows into Excel immediately as OCR results arrive."""

    def __init__(
        self,
        book: xw.Book,
        *,
        target_mode: str,
        sheet_name: str,
        start_cell: str,
    ) -> None:
        normalized_target = (target_mode or "sheet").lower()
        if normalized_target not in {"sheet", "active"}:
            raise RuntimeError(
                "Unknown target mode. Use 'sheet' or 'active'."
            )
        self.target_mode = normalized_target
        if normalized_target == "active":
            sheet, base_range = _get_active_start_range(book)
        else:
            normalized_sheet_name = sheet_name or RESULTS_SHEET_NAME
            sheet, created = _ensure_sheet(book, normalized_sheet_name)
            normalized_cell = _normalize_start_cell(start_cell)
            try:
                base_range = sheet.range(normalized_cell)
            except Exception as exc:
                raise RuntimeError(
                    f"Cell '{normalized_cell}' is not valid on sheet '{normalized_sheet_name}'."
                ) from exc
            if not created and normalized_cell == "A1":
                _clear_existing_sheet(sheet)
        self.sheet = sheet
        self.base_range = base_range
        self.start_address = _range_address(base_range)
        self.header_written = False
        self.header_fields: List[str] = []
        self.data_start_range: xw.Range | None = None
        self.rows_written = 0

    def _ensure_header(self, column_names: Sequence[str]) -> None:
        if self.header_written:
            return
        self.header_fields = ["Nomor", "Nama File"] + list(column_names)
        self.base_range.value = [self.header_fields]
        self.data_start_range = self.base_range.offset(1, 0)
        self.header_written = True

    def append_row(self, row: Dict[str, str], column_names: Sequence[str]) -> None:
        self._ensure_header(column_names)
        if self.data_start_range is None:
            raise RuntimeError("Excel data range is not ready yet.")
        values = [row.get(field, "") for field in self.header_fields]
        target = self.data_start_range.offset(self.rows_written, 0)
        target.value = [values]
        self.rows_written += 1

    def finalize(self) -> WriteResult:
        return WriteResult(
            rows_written=self.rows_written,
            sheet_name=self.sheet.name,
            start_address=self.start_address,
            target_mode=self.target_mode,
        )


def _resolve_path(value: str, book: xw.Book | None) -> pathlib.Path:
    path = pathlib.Path(value).expanduser()
    if path.is_absolute():
        return path
    bases: List[pathlib.Path] = []
    if book and book.fullname:
        bases.append(pathlib.Path(book.fullname).resolve().parent)
    bases.append(PROJECT_ROOT)
    for base in bases:
        candidate = (base / path).expanduser()
        if candidate.exists():
            return candidate
    # Fallback: trust the first base even if it doesn't exist yet (useful for new folders).
    return (bases[0] / path).resolve()


def _coerce_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return bool(value)
    lowered = str(value).strip().lower()
    if lowered in {"1", "true", "yes", "y"}:
        return True
    if lowered in {"0", "false", "no", "n"}:
        return False
    return False


def _split_columns(raw_value) -> Sequence[str] | None:
    if raw_value is None:
        return None
    if isinstance(raw_value, (list, tuple)):
        flattened: List[str] = []
        for item in raw_value:
            if isinstance(item, Iterable) and not isinstance(item, (str, bytes)):
                flattened.extend(str(elem).strip() for elem in item if elem)
            else:
                flattened.append(str(item).strip())
        entries = [entry for entry in flattened if entry]
        return entries or None
    text = str(raw_value)
    normalized = text.replace(";", ",")
    entries = [part.strip() for part in normalized.split(",") if part.strip()]
    return entries or None


def _normalize_start_cell(cell_ref: str | None) -> str:
    text = (cell_ref or "A1").strip()
    if not text:
        return "A1"
    return text.upper()


def _range_address(target_range: xw.Range) -> str:
    try:
        return target_range.get_address(
            row_absolute=False,
            column_absolute=False,
            include_sheetname=False,
        )
    except Exception:
        try:
            return target_range.address.replace("$", "")
        except Exception:
            return "A1"


def _ensure_sheet_visible(sheet: xw.Sheet) -> None:
    try:
        sheet.api.visible.set(True)
        return
    except Exception:
        pass
    try:
        sheet.api.Visible = -1  # xlSheetVisible
        return
    except Exception:
        pass
    try:
        sheet.visible = True
    except Exception:
        pass


def _resolve_images_dir(value, book: xw.Book | None) -> pathlib.Path:
    raw = str(value).strip() if value else ""
    if not raw:
        candidate = DEFAULT_IMAGES_DIR
    else:
        candidate = _resolve_path(raw, book)
    if candidate.exists():
        if candidate.is_dir():
            return candidate
        if candidate.is_file():
            return candidate.parent
    suffix = pathlib.Path(raw).suffix.lower()
    if suffix and suffix in SUPPORTED_EXTENSIONS:
        alt_candidate = DEFAULT_IMAGES_DIR / pathlib.Path(raw).name
        if alt_candidate.exists():
            return alt_candidate.parent
    if raw:
        return candidate
    return DEFAULT_IMAGES_DIR


def _load_config(sheet: xw.Sheet) -> ExcelRPAConfig:
    book = sheet.book
    images_cell_value = sheet.range(IMAGES_DIR_CELL).value or ""
    images_dir = _resolve_images_dir(images_cell_value, book)
    lang = str(sheet.range(LANG_CELL).value or "eng").strip() or "eng"
    recursive = _coerce_bool(sheet.range(RECURSIVE_CELL).value)
    columns = _split_columns(sheet.range(VALUE_COLUMNS_CELL).value)
    return ExcelRPAConfig(
        images_dir=images_dir,
        lang=lang,
        recursive=recursive,
        value_columns=columns,
    )


def _ensure_sheet(book: xw.Book, name: str) -> Tuple[xw.Sheet, bool]:
    def _add_sheet() -> xw.Sheet:
        try:
            last_sheet = book.sheets[-1]
        except Exception:
            return book.sheets.add(name=name)
        return book.sheets.add(name=name, after=last_sheet)

    try:
        sheet = book.sheets[name]
        _ensure_sheet_visible(sheet)
        return sheet, False
    except Exception:
        sheet = _add_sheet()
        sheet.name = name
        _ensure_sheet_visible(sheet)
        return sheet, True


def _clear_existing_sheet(sheet: xw.Sheet) -> None:
    try:
        sheet.clear()
        return
    except Exception:
        pass
    # Fallback: clear by overwriting the current data region (if any)
    try:
        region = sheet.range("A1").expand("table")
        region.clear_contents()
        region.clear_formats()
        return
    except Exception:
        pass
    # Last resort: clear a large bounding box so stale data doesn't linger.
    try:
        sheet.range("A:ZZ").clear_contents()
    except Exception:
        pass



def _get_active_start_range(book: xw.Book) -> Tuple[xw.Sheet, xw.Range]:
    app = book.app
    address: str | None = None
    selection_sheet: xw.Sheet | None = None

    def _capture(candidate) -> bool:
        nonlocal address, selection_sheet
        try:
            if candidate is None:
                return False
            target_address = getattr(candidate, "address", None)
            if not target_address:
                return False
            address = target_address
            try:
                selection_sheet = candidate.sheet
            except Exception:
                selection_sheet = None
            return True
        except Exception:
            return False

    if not _capture(getattr(app, "selection", None)):
        _capture(getattr(book, "selection", None))

    sheet = selection_sheet
    if sheet is None:
        try:
            sheet = book.sheets.active
        except Exception as exc:
            raise RuntimeError(
                "No worksheet is active. Select one in Excel or run again with --target sheet."
            ) from exc
    if sheet is None:
        raise RuntimeError(
            "No worksheet is active. Select one in Excel or run again with --target sheet."
        )

    if not address:
        try:
            address = sheet.api.Application.ActiveCell.Address
        except Exception as exc:
            raise RuntimeError(
                "ActiveCell cannot be determined. Select the destination cell in Excel or run with --target sheet."
            ) from exc

    if not address:
        raise RuntimeError(
            "ActiveCell cannot be determined. Select the destination cell in Excel or run with --target sheet."
        )

    first_area = address.split(",")[0]
    first_cell = first_area.split(":")[0].replace("$", "")
    if not first_cell:
        first_cell = "A1"
    elif first_cell.isdigit():
        first_cell = f"A{first_cell}"
    elif first_cell.isalpha():
        first_cell = f"{first_cell}1"

    try:
        target_range = sheet.range(first_cell)
    except Exception as exc:
        raise RuntimeError(
            "Unable to access the active cell. Ensure the workbook is not showing a modal dialog."
        ) from exc

    return sheet, target_range




def _update_status(sheet: xw.Sheet, message: str) -> None:
    sheet.range(STATUS_CELL).value = message


def _record_last_run(sheet: xw.Sheet) -> None:
    sheet.range(LAST_RUN_CELL).value = dt.datetime.now()


def run_rpa(
    book: xw.Book | None = None,
    *,
    target_mode: str = "sheet",
    sheet_name: str = RESULTS_SHEET_NAME,
    start_cell: str = "A1",
    status_updates: bool = True,
) -> pathlib.Path:
    """
    Execute the OCR pipeline using the configuration stored on the control sheet.
    """

    if book is None:
        book = xw.Book.caller()
    try:
        sheet = book.sheets[CONTROL_SHEET_NAME]
    except (KeyError, IndexError) as exc:
        raise RuntimeError(
            f"Workbook does not contain a sheet named '{CONTROL_SHEET_NAME}'."
        ) from exc
    config = _load_config(sheet)
    logger = LOGGER
    log_buffer: List[str] = []
    log_handler = _BufferingLogHandler(log_buffer)
    logger.addHandler(log_handler)
    run_started = time.perf_counter()
    workbook_label = (
        getattr(book, "fullname", None) or getattr(book, "name", None) or "<unknown>"
    )
    value_columns_summary = (
        ", ".join(config.value_columns) if config.value_columns else "default"
    )
    logger.info(
        "Starting Excel RPA run | workbook=%s | target_mode=%s | sheet_name=%s | start_cell=%s | status_updates=%s | images_dir=%s | lang=%s | recursive=%s | value_columns=%s",
        workbook_label,
        target_mode,
        sheet_name,
        start_cell,
        status_updates,
        config.images_dir,
        config.lang,
        config.recursive,
        value_columns_summary,
    )

    def _maybe_update_status(message: str) -> None:
        if not status_updates:
            return
        try:
            _update_status(sheet, message)
        except Exception:
            pass

    def _maybe_record_last_run() -> None:
        if not status_updates:
            return
        try:
            _record_last_run(sheet)
        except Exception:
            pass

    _maybe_update_status("Running OCR...")
    progress_start_times: Deque[float] = deque()
    row_durations: List[float] = []

    def progress(message: str) -> None:
        _maybe_update_status(message)
        normalized_message = str(message or "").strip()
        if normalized_message.startswith("OCR:"):
            progress_start_times.append(time.perf_counter())
        logger.info("Progress: %s", message)

    writer = ExcelStreamWriter(
        book,
        target_mode=target_mode,
        sheet_name=sheet_name,
        start_cell=start_cell,
    )
    logger.info(
        "ExcelStreamWriter ready | target_mode=%s | sheet=%s | start_cell=%s",
        writer.target_mode,
        writer.sheet.name,
        writer.start_address,
    )


    def row_callback(row: Dict[str, str], columns: Sequence[str]) -> None:
        start_time = progress_start_times.popleft() if progress_start_times else None
        writer.append_row(row, columns)
        end_time = time.perf_counter()
        duration_ms: float | None = None
        if start_time is not None:
            duration = end_time - start_time
            row_durations.append(duration)
            duration_ms = duration * 1000
        logger.info(
            "Excel row written | seq=%s | filename=%s | total_columns=%d%s",
            row.get("Nomor"),
            row.get("Nama File"),
            len(columns) + 2,  # plus Nomor & Nama File
            f" | elapsed_ms={duration_ms:.0f}" if duration_ms is not None else "",
        )

    csv_path: pathlib.Path | None = None
    write_result: WriteResult | None = None
    try:
        logger.info(
            "Running OCR pipeline | images_dir=%s | lang=%s | recursive=%s",
            config.images_dir,
            config.lang,
            config.recursive,
        )
        csv_path = run_ocr_pipeline(
            config.images_dir,
            lang=config.lang,
            recursive=config.recursive,
            value_columns=config.value_columns,
            output_path=DEFAULT_OUTPUT,
            progress_callback=progress,
            row_callback=row_callback,
        )
        logger.info("OCR pipeline finished | output=%s", csv_path)
        write_result = writer.finalize()
        logger.info(
            "Excel write completed | rows=%d | sheet=%s | start=%s | mode=%s",
            write_result.rows_written,
            write_result.sheet_name,
            write_result.start_address,
            write_result.target_mode,
        )
    except Exception as exc:  # pragma: no cover - Excel surface
        _maybe_update_status(f"Error: {exc}")
        elapsed_failed = time.perf_counter() - run_started
        logger.exception(
            "Excel RPA run failed after %.2f seconds: %s",
            elapsed_failed,
            exc,
        )
        raise
    else:
        _maybe_record_last_run()
        if csv_path:
            _maybe_update_status(f"Done: {csv_path.name}")
        elapsed = time.perf_counter() - run_started
        total_ms = elapsed * 1000
        if row_durations:
            durations_ms = [value * 1000 for value in row_durations]
            avg_ms = sum(durations_ms) / len(durations_ms)
            min_ms = min(durations_ms)
            max_ms = max(durations_ms)
            stats_tuple: Tuple[float, float, float] | None = (
                avg_ms,
                min_ms,
                max_ms,
            )
        else:
            stats_tuple = None
        rows_processed = write_result.rows_written if write_result else None
        _log_run_summary(
            logger,
            total_ms=total_ms,
            row_stats=stats_tuple,
            rows_written=rows_processed,
            csv_path=csv_path,
        )
        if csv_path:
            _run_validation(csv_path, logger)
    finally:
        logger.removeHandler(log_handler)
        log_handler.close()
        _prepend_run_log(log_buffer)

    if write_result:
        print(
            "Excel write: "
            f"{write_result.rows_written} rows -> "
            f"{write_result.sheet_name}!{write_result.start_address} "
            f"(target={write_result.target_mode})."
        )
    if csv_path is None:
        raise RuntimeError("OCR pipeline did not return an output path.")
    return csv_path


@xw.sub
def run_ocr_rpa_macro() -> None:
    """
    Macro entry point wired to an Excel button.
    """
    run_rpa()


def _is_object_not_found_error(exc: Exception) -> bool:
    text = str(exc)
    if "-1728" in text:
        return True
    for arg in getattr(exc, "args", []):
        if isinstance(arg, str) and "-1728" in arg:
            return True
    return False


def _handle_appscript_exception(exc: Exception) -> None:
    if _is_object_not_found_error(exc):
        guidance = (
            "Workbook/sheet not found (OSERROR -1728).\n"
            "- Make sure the correct workbook is open.\n"
            "- Use --book with an absolute path if necessary.\n"
            "- Verify the sheet name (--sheet-name).\n"
            "- Or run with --target active to write to the current cursor location."
        )
        print(f"Error: {guidance}", file=sys.stderr)
        sys.exit(2)
    print(f"Error Excel: {exc}", file=sys.stderr)
    sys.exit(2)


def _attach_workbook(book_path: str | None) -> xw.Book:
    if book_path:
        path = pathlib.Path(book_path).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        try:
            return xw.Book(path)
        except Exception as exc:
            raise RuntimeError(
                f"Failed to open workbook '{path}'. Ensure the file is accessible."
            ) from exc
    try:
        app = xw.apps.active
    except Exception as exc:  # pragma: no cover - Excel runtime guard
        raise RuntimeError(
            "Excel is not active. Launch Excel first or provide --book."
        ) from exc
    if app is None:
        raise RuntimeError(
            "Excel is not active. Launch Excel first or provide --book."
        )
    try:
        book = app.books.active
    except Exception as exc:  # pragma: no cover - Excel runtime guard
        raise RuntimeError(
            "No workbook is active. Open a file in Excel or provide --book."
        ) from exc
    if book is None:
        raise RuntimeError(
            "No workbook is active. Open a file in Excel or provide --book."
        )
    return book


def main(argv: Sequence[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Run the OCR RPA workflow directly from the CLI.")
    parser.add_argument(
        "--book",
        metavar="PATH",
        help="Path to the .xlsm workbook containing the RPA Control sheet. If omitted, the script attaches to the active workbook in Excel.",
    )
    parser.add_argument(
        "--target",
        choices=("sheet", "active"),
        default="active",
        help="Writing mode: 'active' (default) starts from the active cell, 'sheet' writes to a specific sheet.",
    )
    parser.add_argument(
        "--sheet-name",
        default=None,
        help="Destination sheet when target=sheet (default 'Sheet1' for CLI usage).",
    )
    parser.add_argument(
        "--start-cell",
        default="A1",
        help="Starting cell when target=sheet, e.g. 'A1' or 'B3' (default 'A1').",
    )
    parser.add_argument(
        "--status-updates",
        choices=("on", "off"),
        default="off",
        help="Control whether status cells on the control sheet are updated (default 'off' via CLI).",
    )
    args = parser.parse_args(argv)
    default_cli_sheet = "Sheet1"
    raw_sheet_name = args.sheet_name or default_cli_sheet
    sheet_name = raw_sheet_name.strip() or default_cli_sheet
    start_cell = (args.start_cell or "A1").strip() or "A1"
    try:
        book = _attach_workbook(args.book)
        csv_path = run_rpa(
            book,
            target_mode=args.target,
            sheet_name=sheet_name,
            start_cell=start_cell,
            status_updates=args.status_updates == "on",
        )
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(2)
    except APPLESCRIPT_ERRORS as exc:  # pragma: no cover - macOS only
        _handle_appscript_exception(exc)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(2)
    except Exception as exc:  # pragma: no cover - unexpected
        print(f"Error tak terduga: {exc}", file=sys.stderr)
        sys.exit(1)
    print(f"Done! Latest results were written to {csv_path}.")


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main()
