# -*- coding: utf-8 -*-
"""
PDF to Excel Converter  -  Coordinate-based extraction (no Java required)
=========================================================================
Handles structured PDFs where table rows are free-floating text positioned
by x-coordinate (not enclosed in drawn borders).

Supports:
  - Multi-page PDFs (same header on every page)
  - Wrapping Particulars field (multi-line per transaction)
  - Auto-detection of column x-boundaries from the header row
  - Professional Excel formatting

Usage (CLI):
    python pdf_to_excel_converter.py --pdf input.pdf --excel output.xlsx

Usage (import):
    from pdf_to_excel_converter import pdf_to_excel
    pdf_to_excel("input.pdf", "output.xlsx")
"""

from backend_common.constants import UNUSED_BANK_FOOTER_DATA, SAMPLE_BANK_COLUMN_MAPPING
import argparse
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import fitz
import time


# ---------------------------------------------------------------------------
# Constants & Helper
# ---------------------------------------------------------------------------


# Tolerance in points for grouping words onto the same visual line
LINE_Y_TOLERANCE = 4


# Shared helpers from bank_processor
from .bank_processor import (
    _normalize_text,
    _collapse_whitespace,
    _resolve_pdf_path,
    _is_date,
)


# ---------------------------------------------------------------------------
# Helper: group words by y-position into lines
# ---------------------------------------------------------------------------

def _group_words_into_lines(words: list[dict], y_tolerance: float = LINE_Y_TOLERANCE) -> list[list[dict]]:
    """
    Given a list of pdfplumber word dicts, group them into visual lines
    sorted by their vertical (top) position.

    Returns a list of lines, each line being a list of word dicts
    sorted left-to-right by x0.
    """
    if not words:
        return []

    # Sort words by top position
    sorted_words = sorted(words, key=lambda w: w["top"])

    lines = []
    current_line = [sorted_words[0]]
    current_top = sorted_words[0]["top"]

    for word in sorted_words[1:]:
        if abs(word["top"] - current_top) <= y_tolerance:
            current_line.append(word)
        else:
            lines.append(sorted(current_line, key=lambda w: w["x0"]))
            current_line = [word]
            current_top = word["top"]

    if current_line:
        lines.append(sorted(current_line, key=lambda w: w["x0"]))

    return lines


# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------

def _find_header_matches(line: list[dict], column_mapping: dict | None = None) -> list[tuple[str, list[dict]]]:
    """
    Given a line of word dicts, identify which column_mapping patterns match and where.
    Returns a list of (canonical_key, matched_words_list) sorted by the x0 of the match.
    """
    active_mapping = column_mapping if column_mapping else SAMPLE_BANK_COLUMN_MAPPING
    normalized_words = [{"word": w, "text": _normalize_text(w["text"])} for w in line]
    matches = []
    i = 0
    n = len(normalized_words)

    while i < n:
        matched = False
        # Match longest phrase first (up to 3 words)
        for phrase_len in [3, 2, 1]:
            if i + phrase_len <= n:
                phrase_tokens = [normalized_words[j]["text"] for j in range(i, i + phrase_len)]
                phrase_str = " ".join(phrase_tokens)

                for key, patterns in active_mapping.items():
                    normalized_patterns = [_normalize_text(p) for p in patterns]
                    if phrase_str in normalized_patterns:
                        matched_words = [normalized_words[j]["word"] for j in range(i, i + phrase_len)]
                        matches.append((key, matched_words))
                        i += phrase_len
                        matched = True
                        break
            if matched:
                break
        if not matched:
            i += 1

    return matches


def _is_header_line(matches: list[tuple[str, list[dict]]]) -> bool:
    """Evaluate whether a line's matches qualify it as the table header."""
    matched_keys = {key for key, _ in matches}
    has_date = "DATE" in matched_keys
    has_desc = "DESCRIPTION" in matched_keys
    
    # Must have date + description + at least one other column, OR match 4 or more columns
    if has_date and has_desc and len(matched_keys) >= 3:
        return True
    if len(matched_keys) >= 4:
        return True
    return False


def _find_header_line(lines: list[list[dict]], column_mapping: dict | None = None) -> tuple[int, list[dict], list[tuple[str, list[dict]]]] | None:
    """
    Find the line index containing the transaction table header.
    Returns (line_index, header_words, matches) or None.
    """
    for idx, line in enumerate(lines):
        matches = _find_header_matches(line, column_mapping=column_mapping)
        if _is_header_line(matches):
            return idx, line, matches
    return None


def _refine_column_boundaries(
    matches: list[tuple[str, list[dict]]],
    lines: list[list[dict]],
    header_idx: int
) -> list[tuple[str, float, float]]:
    """
    Refine column boundaries by merging word X-intervals from data lines starting with a date
    and mapping them back to the detected header columns.
    """
    col_positions = []
    seen_keys = set()
    for key, words in matches:
        if key not in seen_keys:
            seen_keys.add(key)
            hx0 = min(w["x0"] for w in words)
            hx1 = max(w["x1"] for w in words)
            col_positions.append({"name": key, "hx0": hx0, "hx1": hx1})

    col_positions.sort(key=lambda x: x["hx0"])

    data_words = []
    for line in lines[header_idx + 1:]:
        if not line:
            continue
        line_text = " ".join(w["text"] for w in line).strip()
        if any(kw in line_text.lower() for kw in ["hdfc bank", "generated on", "closing balance", "statement summary", "page"]):
            continue
        leftmost = line[0]
        if _is_date(leftmost["text"]):
            for w in line:
                data_words.append(w)

    intervals = []
    for w in sorted(data_words, key=lambda x: x["x0"]):
        x0, x1 = w["x0"], w["x1"]
        if not intervals:
            intervals.append([x0, x1])
        else:
            last = intervals[-1]
            if x0 <= last[1] + 5:
                last[1] = max(last[1], x1)
            else:
                intervals.append([x0, x1])

    for col in col_positions:
        best_interval = None
        best_overlap = 0.0
        for ix0, ix1 in intervals:
            overlap = max(0.0, min(col["hx1"], ix1) - max(col["hx0"], ix0))
            if overlap > best_overlap:
                best_overlap = overlap
                best_interval = (ix0, ix1)

        best_dist = float("inf")
        if best_overlap == 0.0:
            for ix0, ix1 in intervals:
                if ix0 > col["hx1"]:
                    dist = ix0 - col["hx1"]
                elif col["hx0"] > ix1:
                    dist = col["hx0"] - ix1
                else:
                    dist = 0.0
                if dist < best_dist:
                    best_dist = dist
                    best_interval = (ix0, ix1)

        if best_overlap > 0 or (best_interval and best_dist < 40):
            col["rx0"] = best_interval[0]
            col["rx1"] = best_interval[1]
        else:
            col["rx0"] = col["hx0"]
            col["rx1"] = col["hx1"]

    result = []
    n = len(col_positions)
    for i, col in enumerate(col_positions):
        name = col["name"]
        if i == 0:
            x_start = 0.0
        else:
            x_start = (col_positions[i - 1]["rx1"] + col["rx0"]) / 2.0

        if i == n - 1:
            x_end = 9999.0
        else:
            x_end = (col["rx1"] + col_positions[i + 1]["rx0"]) / 2.0

        result.append((name, x_start, x_end))

    return result


def _assign_word_to_column(word: dict, col_boundaries: list[tuple[str, float, float]]) -> str | None:
    """Return the column name for a word based on its x0 position."""
    x = word["x0"]
    for name, x_start, x_end in col_boundaries:
        if x_start <= x < x_end:
            return name
    if col_boundaries:
        return min(col_boundaries, key=lambda c: abs((c[1] + c[2])/2 - x))[0]
    return None


# ---------------------------------------------------------------------------
# Row reconstruction
# ---------------------------------------------------------------------------

def _reconstruct_rows(
    data_lines: list[list[dict]],
    col_boundaries: list[tuple[str, float, float]],
) -> list[dict]:
    """
    Convert a flat list of visual lines (below the header) into structured
    transaction rows.
    """
    if not col_boundaries:
        return []

    tran_date_col = None
    for col in col_boundaries:
        if col[0] == "DATE":
            tran_date_col = col
            break
    if not tran_date_col:
        tran_date_col = col_boundaries[0]

    rows = []
    current_row: dict | None = None

    for line in data_lines:
        if not line:
            continue

        leftmost_word = line[0]
        leftmost_x = leftmost_word["x0"]
        date_end = tran_date_col[2]

        in_date_col = (leftmost_x <= date_end + 5)
        is_date = _is_date(leftmost_word["text"])
        is_new_row = in_date_col and is_date

        if is_new_row:
            if current_row is not None:
                rows.append(current_row)
            current_row = {col[0]: "" for col in col_boundaries}

            for word in line:
                col_name = _assign_word_to_column(word, col_boundaries)
                if col_name:
                    sep = " " if current_row[col_name] else ""
                    current_row[col_name] += sep + word["text"]
        else:
            if current_row is not None:
                particulars_col = _find_particulars_col(col_boundaries)
                continuation_words = [w["text"] for w in line
                                       if _assign_word_to_column(w, col_boundaries) == particulars_col]
                if continuation_words:
                    sep = " " if current_row.get(particulars_col, "") else ""
                    current_row[particulars_col] = current_row.get(particulars_col, "") + sep + " ".join(continuation_words)
                else:
                    extra = " ".join(w["text"] for w in line)
                    particulars_col_name = particulars_col or list(current_row.keys())[3]
                    sep = " " if current_row.get(particulars_col_name, "") else ""
                    current_row[particulars_col_name] = current_row.get(particulars_col_name, "") + sep + extra

    if current_row is not None:
        rows.append(current_row)

    return rows


def _find_particulars_col(col_boundaries: list[tuple[str, float, float]]) -> str:
    """Return the name of the description column."""
    for name, _, _ in col_boundaries:
        if name == "DESCRIPTION":
            return name
    for name, _, _ in col_boundaries:
        if "description" in name.lower() or "particulars" in name.lower():
            return name
    if len(col_boundaries) >= 4:
        return col_boundaries[3][0]
    return col_boundaries[-1][0]


# ---------------------------------------------------------------------------
# Per-page extraction
# ---------------------------------------------------------------------------

def _extract_page_rows(page, page_num: int, prev_boundaries=None, prev_start_y=None, prev_end_y=None, column_mapping: dict | None = None) -> tuple:
    """
    Extract transaction rows from a single page.

    Returns:
        (col_boundaries, rows, table_start_y, table_end_y, original_header_texts)
        where original_header_texts is a dict mapping canonical key → original PDF text
        (only populated when a header row was actually found on this page; else {}).
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    if not words:
        print("  Page %d: no words found." % page_num)
        return [], [], None, None, {}

    lines = _group_words_into_lines(words)

    result = _find_header_line(lines, column_mapping=column_mapping)
    col_boundaries = []
    header_idx = -1
    table_start_y = None
    table_end_y = None
    original_header_texts = {}  # canonical_key -> original text from PDF

    if result is not None:
        header_idx, header_words, matches = result
        col_boundaries = _refine_column_boundaries(matches, lines, header_idx)
        # Build original header map: join all matched word texts for each canonical key
        original_header_texts = {
            key: " ".join(w["text"] for w in words)
            for key, words in matches
        }
        header_y = min(w["top"] for w in header_words)
        table_start_y = header_y
        print("  Page %d: header at line %d (y=%.1f), %d columns: %s" % (
            page_num, header_idx + 1, header_y, len(col_boundaries),
            [c[0] for c in col_boundaries]
        ))
    else:
        if prev_boundaries is not None:
            col_boundaries = prev_boundaries
            table_start_y = prev_start_y
            table_end_y = prev_end_y
            print("  Page %d: header not found. Reusing previous boundaries." % page_num)
        else:
            print("  Page %d: header not found and no previous boundaries available. Skipping." % page_num)
            return [], [], None, None, {}

    footer_y = 9999.0
    for line in lines:
        line_text = " ".join(w["text"] for w in line).strip()
        norm_line = line_text.lower().replace(" ", "").replace("-", "").replace(":", "").replace(".", "")
        is_footer = any(kw in norm_line for kw in UNUSED_BANK_FOOTER_DATA)
        if is_footer:
            line_y = min(w["top"] for w in line)
            if line_y > 400 and line_y < footer_y:
                footer_y = line_y

    if footer_y == 9999.0 and prev_end_y is not None:
        footer_y = prev_end_y
    else:
        table_end_y = footer_y

    filtered_data_lines = []
    for idx, line in enumerate(lines):
        if not line:
            continue
        line_y = min(w["top"] for w in line)

        if header_idx != -1:
            if idx <= header_idx:
                continue
        else:
            if table_start_y is not None and line_y < table_start_y - 10:
                continue

        if line_y >= footer_y - 2:
            continue

        line_text = " ".join(w["text"] for w in line).strip()
        if line_text.isdigit():
            continue
        if len(line_text) > 200 and line_text.count("BANK") > 3:
            continue
        if line_text.startswith("Generated on:"):
            continue

        filtered_data_lines.append(line)

    rows = _reconstruct_rows(filtered_data_lines, col_boundaries)
    print("  Page %d: extracted %d rows." % (page_num, len(rows)))
    return col_boundaries, rows, table_start_y, table_end_y, original_header_texts


# ---------------------------------------------------------------------------
# Full PDF extraction
# ---------------------------------------------------------------------------

def extract_bank_statement(pdf_path: str, column_mapping: dict | None = None) -> tuple:
    """
    Extract all transaction rows from all pages of the PDF.

    Args:
        pdf_path:       Path to the source PDF.
        column_mapping: Optional bank-specific column mapping dict
                        (same structure as the module-level COLUMN_MAPPING).
                        When None, the static COLUMN_MAPPING is used.

    Returns:
        (df, original_headers) where:
          - df is a pandas DataFrame with canonical column names
          - original_headers is a dict mapping canonical_key → original PDF header text
    """
    # Use provided mapping, or fall back to the module-level static mapping
    active_column_mapping = column_mapping if column_mapping else SAMPLE_BANK_COLUMN_MAPPING
    all_rows = []
    col_names = None
    active_boundaries = None
    table_start_y = None
    table_end_y = None
    original_headers: dict = {}   # populated from the first page that has a header

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print("  PDF has %d page(s)." % total_pages)

        for page_num, page in enumerate(pdf.pages, start=1):
            col_boundaries, rows, start_y, end_y, page_original_headers = _extract_page_rows(
                page, page_num,
                prev_boundaries=active_boundaries,
                prev_start_y=table_start_y,
                prev_end_y=table_end_y,
                column_mapping=active_column_mapping,
            )

            if col_boundaries:
                if active_boundaries is None or (start_y is not None and start_y != table_start_y):
                    active_boundaries = col_boundaries
                    table_start_y = start_y
                    if end_y is not None:
                        table_end_y = end_y

            # Capture original header names only from the first page that has them
            if col_boundaries and col_names is None:
                col_names = [c[0] for c in col_boundaries]
                if page_original_headers:
                    original_headers = page_original_headers

            all_rows.extend(rows)

    if not all_rows:
        return pd.DataFrame(), original_headers

    df = pd.DataFrame(all_rows)

    # Reorder to match original column order
    if col_names:
        existing_cols = [c for c in col_names if c in df.columns]
        df = df[existing_cols]

    # --- Post-processing ---

    # Drop rows where DATE is empty AND all amount columns are empty
    amount_cols = [c for c in df.columns if c in ("WITHDRAWAL_AMOUNT", "DEPOSIT_AMOUNT", "CLOSING_BALANCE")]
    date_col = "DATE" if "DATE" in df.columns else (df.columns[0] if len(df.columns) > 0 else None)
    if date_col and amount_cols:
        noise_mask = (
            df[date_col].str.strip().eq("") &
            df[amount_cols].apply(lambda r: r.str.strip().eq("").all(), axis=1)
        )
        df = df[~noise_mask]

    df = df.reset_index(drop=True)
    return df, original_headers


# ---------------------------------------------------------------------------
# Excel output with formatting
# ---------------------------------------------------------------------------

def _style_worksheet(ws) -> None:
    """Apply professional formatting to an openpyxl worksheet."""
    header_fill = PatternFill("solid", fgColor="1F3864")
    alt_fill    = PatternFill("solid", fgColor="EBF3FB")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    data_font   = Font(name="Calibri", size=10)

    thin   = Side(style="thin", color="B8CCE4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = border
            if row_idx == 1:
                cell.fill      = header_fill
                cell.font      = header_font
                cell.alignment = center
            else:
                cell.font      = data_font
                cell.alignment = left
                if row_idx % 2 == 0:
                    cell.fill = alt_fill

    # Auto-fit column widths (capped at 60)
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(cell.value or "")) for cell in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    ws.freeze_panes = "A2"


def write_to_excel(df: pd.DataFrame, output_path: str) -> None:
    """Write DataFrame to a formatted Excel file."""
    if df.empty:
        print("  No data to write.")
        return

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)

    wb = load_workbook(output_path)
    for ws in wb.worksheets:
        _style_worksheet(ws)
    wb.save(output_path)

    print("  Excel saved -> %s" % output_path)
    print("  Total rows: %d" % len(df))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def pdf_to_excel(pdf_path: str, excel_path: str) -> None:
    """
    Convert a structured PDF bank statement to Excel.

    Args:
        pdf_path:   Path to the source PDF.
        excel_path: Destination path for the output Excel file.
    """
    pdf_path   = str(Path(pdf_path).resolve())
    excel_path = str(Path(excel_path).resolve())

    if not os.path.exists(pdf_path):
        raise FileNotFoundError("PDF not found: %s" % pdf_path)

    print("\n" + "=" * 60)
    print("  PDF  -> %s" % pdf_path)
    print("  XLSX -> %s" % excel_path)
    print("=" * 60)

    print("\n[1/2] Extracting transactions from PDF ...")
    df = extract_bank_statement(pdf_path)

    if df.empty:
        print("\n  [!] No data extracted. PDF may be scanned/image-based.")
        sys.exit(1)

    print("\n[2/2] Writing to Excel ...")
    write_to_excel(df, excel_path)

    print("\n[OK]  Done! %d transaction rows written.\n" % len(df))




def is_structured_pdf(pdf_path: str,column_mapping: dict | None = None,max_pages: int = 1,minimum_matches: int = 4) -> bool:
    active_mapping = (column_mapping if column_mapping else SAMPLE_BANK_COLUMN_MAPPING)
    try:
        doc = fitz.open(pdf_path)
        pages_to_check = min(max_pages, len(doc))
        # Flatten all header aliases
        expected_headers = set()
        for aliases in active_mapping.values():
            for alias in aliases:
                expected_headers.add(
                    alias.lower().strip()
                )

        for page_no in range(pages_to_check):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue
            text = text.lower()
            matches = []
            for header in expected_headers:
                if header in text:
                    matches.append(header)
            if len(matches) >= minimum_matches:
                doc.close()
                return True
        doc.close()
    except Exception as exc:
        pass
    return False


def pdf_to_document(pdf_path: str, column_mapping: dict | None = None) -> tuple:
    """
    Extract transaction rows from a structured PDF bank statement and
    return them directly as a pandas DataFrame — no Excel conversion.

    Returns an empty DataFrame if the PDF is unstructured / image-based.

    Args:
        pdf_path:       Absolute or relative path to the source PDF.
        column_mapping: Optional bank-specific column mapping dict fetched
                        from the edocsmart_bank_mapping collection. When None,
                        the module-level static COLUMN_MAPPING is used.

    Returns:
        (df, original_headers) where:
          - df is a pandas DataFrame with canonical column names
            (e.g. DATE, DESCRIPTION, WITHDRAWAL_AMOUNT …)
          - original_headers is a dict mapping canonical_key → original PDF
            header text (e.g. {"DATE": "Tran Date", "DESCRIPTION": "Narration"})
        Both are empty when no transactions are found.
    """
    pdf_path = _resolve_pdf_path(pdf_path)

    print("\n[pdf_to_document] Extracting transactions from: %s" % pdf_path)
    df, original_headers = extract_bank_statement(pdf_path, column_mapping=column_mapping)

    if df.empty:
        print("  [!] No structured data found. PDF may be image-based.")
    else:
        print("  [OK] %d transaction rows extracted." % len(df))

    return df, original_headers


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert a structured PDF bank statement to Excel"
    )
    parser.add_argument(
        "--pdf",
        default=r"C:\Tejas\Projects\RnD\pdf_to_excel\kalupur_bank.pdf",
        help="Path to the input PDF file.",
    )
    parser.add_argument(
        "--excel",
        default=r"C:\Tejas\Projects\RnD\pdf_to_excel\output.xlsx",
        help="Path for the output Excel file.",
    )
    args = parser.parse_args()

    pdf_to_excel(args.pdf, args.excel)
