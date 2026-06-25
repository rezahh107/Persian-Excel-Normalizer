"""
normalize_excel.py — Production-grade Persian text normaliser for Excel files.

Architecture
------------
* Pure-function core — normalize_text() and normalize_worksheet() are
  side-effect-free and unit-testable.
* CLI is a thin adapter — main() only parses arguments and delegates.
* Structured logging on a dedicated "excel_normalizer.core" logger;
  the root logger is never touched.
* Optional progress / cancellation callbacks are keyword-only with None
  defaults so the public API remains backwards-compatible.
* pathlib is used throughout for OS-agnostic path handling.

Performance notes
-----------------
* normalize_text() uses str.translate() with two pre-built tables
  (_TABLE_REPLACE_ZWNJ / _TABLE_KEEP_ZWNJ) constructed at import time.
  A single translate() call replaces all Arabic→Persian substitutions,
  invisible-char deletions and NBSP normalisation in one O(n) pass.
* openpyxl loads the entire workbook into RAM. For very large files,
  consider splitting large workbooks before processing.

Formula / macro safety
----------------------
* Formula cells are skipped using openpyxl metadata (cell.data_type == "f")
  and the legacy string-prefix guard for compatibility.
* .xlsm files are loaded with keep_vba=True so openpyxl preserves the VBA
  archive while saving. Real macro execution must still be verified in Excel.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import tempfile
from pathlib import Path
from typing import Callable, Optional, TypeVar, overload

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

log: logging.Logger = logging.getLogger("excel_normalizer.core")


class NormalizationCancelled(RuntimeError):
    """Raised when cancellation is requested before saving output."""


_BASE_TABLE: dict[str, Optional[str]] = {
    "\u064a": "\u06cc",   # Arabic yeh        → Persian yeh
    "\u0643": "\u06a9",   # Arabic kaf         → Persian kaf
    "\u0649": "\u06cc",   # Alef maqsura       → Persian yeh
    "\u0629": "\u0647",   # Taa marbuta        → Persian heh
    "\u0623": "\u0627",   # Alef hamza above   → bare alef
    "\u0625": "\u0627",   # Alef hamza below   → bare alef
    "\u0624": "\u0648",   # Waw with hamza     → bare waw
    "\u0626": "\u06cc",   # Yeh with hamza     → Persian yeh
    "\u200d": None,        # ZWJ
    "\u200e": None,        # LRM
    "\u200f": None,        # RLM
    "\u202a": None,        # LRE
    "\u202b": None,        # RLE
    "\u202c": None,        # PDF
    "\u2066": None,        # LRI
    "\u2067": None,        # RLI
    "\u2068": None,        # FSI
    "\u2069": None,        # PDI
    "\xad": None,          # SHY (soft hyphen)
    "\xa0": " ",           # NBSP → regular space
}

_TABLE_REPLACE_ZWNJ = str.maketrans({**_BASE_TABLE, "\u200c": " "})
_TABLE_KEEP_ZWNJ = str.maketrans(_BASE_TABLE)

_MULTI_SPACE_RE: re.Pattern[str] = re.compile(r"[ \t]+")
_SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
_T = TypeVar("_T")


@overload
def normalize_text(text: str, *, replace_zwnj: bool = True) -> str:
    ...


@overload
def normalize_text(text: _T, *, replace_zwnj: bool = True) -> _T:
    ...


def normalize_text(text: object, *, replace_zwnj: bool = True) -> object:
    """Normalise a single Persian string.

    Non-string inputs are returned unchanged to preserve legacy callers that
    pass through mixed cell values before type filtering.
    """
    if not isinstance(text, str):
        return text

    table = _TABLE_REPLACE_ZWNJ if replace_zwnj else _TABLE_KEEP_ZWNJ
    translated = text.translate(table)
    return _MULTI_SPACE_RE.sub(" ", translated).strip()


def _is_formula_cell(cell) -> bool:
    """Return True when an openpyxl cell should be treated as a formula.

    openpyxl marks formula cells with ``data_type == "f"``. The string-prefix
    check is retained as a conservative compatibility guard so existing
    behaviour for text-like values beginning with '=' is not weakened.
    """
    if getattr(cell, "data_type", None) == "f":
        return True
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


def normalize_worksheet(
    sheet: Worksheet,
    *,
    replace_zwnj: bool = True,
    cancel_check: Optional[Callable[[], bool]] = None,
    row_callback: Optional[Callable[[int], None]] = None,
) -> int:
    """Normalise every eligible string cell in *sheet* in-place."""
    changed = 0
    for row in sheet.iter_rows():
        if cancel_check and cancel_check():
            log.info("Cancellation requested — stopping worksheet iteration.")
            break
        cells_in_row = 0
        for cell in row:
            cells_in_row += 1
            if _is_formula_cell(cell):
                continue
            if not isinstance(cell.value, str):
                continue
            normalised = normalize_text(cell.value, replace_zwnj=replace_zwnj)
            if normalised != cell.value:
                cell.value = normalised
                changed += 1
        if row_callback is not None:
            row_callback(cells_in_row)
    return changed


def _estimate_total_cells(workbook: openpyxl.Workbook) -> int:
    """Return a conservative cell-count estimate across all sheets."""
    total = 0
    for ws in workbook.worksheets:
        total += (ws.max_row or 0) * (ws.max_column or 0)
    return max(total, 1)


def _validate_destination(source: Path, destination: Path) -> None:
    """Validate output path for data safety and OOXML extension consistency."""
    source_suffix = source.suffix.lower()
    destination_suffix = destination.suffix.lower()

    if destination_suffix not in _SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported output extension '{destination.suffix}'. "
            "Expected .xlsx or .xlsm."
        )
    if destination_suffix != source_suffix:
        raise ValueError(
            "Output extension must match input extension to avoid Excel format "
            f"mismatch: {source_suffix} -> {destination_suffix}."
        )
    if source.resolve() == destination.resolve():
        raise ValueError(
            "Refusing to overwrite the input workbook in-place. "
            "Choose a different output path."
        )


def _atomic_save(workbook: openpyxl.Workbook, destination: Path) -> None:
    """Save *workbook* to *destination* using same-directory atomic replace.

    The temporary file is created in the destination directory, then promoted
    with ``os.replace()``. Cleanup failures are logged but never allowed to
    hide the original save/replace error.
    """
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)

    fd, raw_tmp = tempfile.mkstemp(
        dir=destination.parent,
        prefix=f".{destination.name}.",
        suffix=destination.suffix or ".xlsx",
    )
    os.close(fd)
    tmp = Path(raw_tmp)

    try:
        workbook.save(tmp)
        os.replace(tmp, destination)
    except Exception:
        try:
            tmp.unlink(missing_ok=True)
        except OSError as cleanup_error:
            log.warning("Could not remove temporary file '%s': %s", tmp, cleanup_error)
        raise


def normalize_workbook(
    source: Path,
    destination: Path,
    *,
    replace_zwnj: bool = True,
    cancel_check: Optional[Callable[[], bool]] = None,
    on_sheet_start: Optional[Callable[[int, int, str], None]] = None,
    on_progress: Optional[Callable[[int, int], None]] = None,
) -> None:
    """Load *source*, normalise all worksheets, and save to *destination*."""
    source = Path(source)
    destination = Path(destination)

    if not source.exists():
        raise FileNotFoundError(f"Input file not found: {source}")
    if source.suffix.lower() not in _SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported extension '{source.suffix}'. Expected .xlsx or .xlsm."
        )
    _validate_destination(source, destination)

    keep_vba = source.suffix.lower() == ".xlsm"
    log.info("Loading workbook: %s", source)
    workbook = openpyxl.load_workbook(source, keep_vba=keep_vba)
    sheet_names: list[str] = workbook.sheetnames
    total_changed = 0

    total_cells = _estimate_total_cells(workbook)
    processed_cells = 0
    was_cancelled = False

    if on_progress:
        on_progress(0, total_cells)

    progress_cb = on_progress

    def _row_cb(n: int) -> None:
        nonlocal processed_cells
        processed_cells += n
        if progress_cb is not None:
            progress_cb(min(processed_cells, total_cells), total_cells)

    for idx, name in enumerate(sheet_names, start=1):
        if cancel_check and cancel_check():
            was_cancelled = True
            log.info("Cancellation before sheet '%s'. Stopping.", name)
            break
        if on_sheet_start:
            on_sheet_start(idx, len(sheet_names), name)
        log.info("[%d/%d] Normalising sheet: '%s'", idx, len(sheet_names), name)
        changed = normalize_worksheet(
            workbook[name],
            replace_zwnj=replace_zwnj,
            cancel_check=cancel_check,
            row_callback=_row_cb if progress_cb else None,
        )
        if cancel_check and cancel_check():
            was_cancelled = True
        log.debug("  -> %d cell(s) modified in '%s'", changed, name)
        total_changed += changed
        if was_cancelled:
            break

    if was_cancelled:
        raise NormalizationCancelled("Normalization cancelled before saving output.")

    if progress_cb is not None:
        progress_cb(total_cells, total_cells)

    log.info("Saving to: %s (atomic write)", destination)
    _atomic_save(workbook, destination)
    log.info("Done. Total cells modified: %d", total_changed)


def build_output_path(source: Path) -> Path:
    """Return *source* with ``_normalized`` appended before the suffix."""
    source = Path(source)
    return source.with_stem(f"{source.stem}_normalized")


def auto_detect_input(directory: Path) -> Path:
    """Return the first .xlsx / .xlsm in *directory* that is not normalised."""
    candidates = sorted(
        p for p in Path(directory).iterdir()
        if p.suffix.lower() in _SUPPORTED_EXTENSIONS
        and not p.stem.endswith("_normalized")
    )
    if not candidates:
        raise FileNotFoundError(
            f"No .xlsx / .xlsm files found in '{directory}'. "
            "Provide an explicit INPUT path or change the working directory."
        )
    return candidates[0]


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="normalize_excel",
        description=(
            "Normalise Persian text inside Excel (.xlsx / .xlsm) files.\n\n"
            "Transformations applied:\n"
            "  * Arabic chars  ->  Persian equivalents\n"
            "  * ZWNJ (U+200C) ->  regular space  (unless --keep-zwnj)\n"
            "  * Invisible Unicode control chars removed\n"
            "  * NBSP (U+00A0) ->  regular space\n"
            "  * Multiple spaces -> single space\n"
            "  * Leading/trailing whitespace stripped\n"
            "  * Formula cells are left untouched\n"
            "  * .xlsm files are saved with VBA archive preservation"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input", nargs="?", type=Path, default=None, metavar="INPUT",
        help="Source .xlsx / .xlsm file. Auto-detected in cwd if omitted.",
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=None, metavar="OUTPUT",
        help="Destination path. Defaults to <INPUT_stem>_normalized<INPUT_suffix>.",
    )
    parser.add_argument(
        "--keep-zwnj", action="store_true", default=False,
        help="Preserve ZWNJ (U+200C) instead of replacing with a space.",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", default=False,
        help="Enable DEBUG-level logging.",
    )
    return parser


def _configure_cli_logging(*, verbose: bool) -> None:
    """Configure the dedicated package logger for CLI use only."""
    pkg_logger = logging.getLogger("excel_normalizer")
    pkg_logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    if not pkg_logger.handlers:
        handler = logging.StreamHandler(stream=sys.stderr)
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
                datefmt="%H:%M:%S",
            )
        )
        pkg_logger.addHandler(handler)
    pkg_logger.propagate = False


def main(argv: Optional[list[str]] = None) -> int:
    """CLI entry point. Returns 0 on success, 1 on error."""
    parser = _build_parser()
    args = parser.parse_args(argv)
    _configure_cli_logging(verbose=args.verbose)

    try:
        if args.input is None:
            source = auto_detect_input(Path.cwd())
            log.info("Auto-detected input: %s", source)
        else:
            source = args.input
        destination: Path = args.output or build_output_path(source)
        normalize_workbook(source, destination, replace_zwnj=not args.keep_zwnj)
    except NormalizationCancelled as exc:
        log.warning("%s", exc)
        return 1
    except (FileNotFoundError, ValueError) as exc:
        log.error("%s", exc)
        return 1
    except OSError as exc:
        log.error("I/O error: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
