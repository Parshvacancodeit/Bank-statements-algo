# -*- coding: utf-8 -*-
"""
Axis Bank PDF Statement Extractor
==================================
Dedicated extractor for Axis Bank structured PDF statements.

Axis Bank PDFs use native drawn-border tables (not free-floating text),
so pdfplumber's ``extract_tables()`` API is used instead of the
coordinate-based word approach used for HDFC / Kalupur statements.

Axis Bank table headers (as they appear in the PDF):
    Tran Date | Chq No | Particulars | Debit | Credit | Balance | Init.Br

These are resolved to canonical keys using the column_mapping fetched from
the database (edocsmart_bank_mapping collection), which has the structure:
    { "DATE": ["tran date"], "DESCRIPTION": ["particulars", ...], ... }

Skipped sentinel rows:
    - OPENING BALANCE
    - TRANSACTION TOTAL
    - CLOSING BALANCE
    - Fully empty rows (no date AND no amounts)

Public API (mirrors structured_pdf_processor.py):
    is_axis_bank_pdf(pdf_path, column_mapping)              -> bool
    extract_axis_bank_statement(pdf_path, column_mapping)   -> (DataFrame, original_headers)
    pdf_to_document_axis(pdf_path, column_mapping)          -> (DataFrame, original_headers)

Shared utilities imported from structured_pdf_processor:
    _collapse_whitespace  -- whitespace-only normaliser (case/punctuation preserved)
    _is_date              -- generic date-string detector
    _resolve_pdf_path     -- path resolution + FileNotFoundError guard
"""

import pandas as pd
import pdfplumber
import fitz
from backend_common.constants import BANK_STATEMENT_CANONICAL_ORDER

# Shared helpers from bank_processor — no duplication needed here.
from .bank_processor import (
    _collapse_whitespace,
    _is_date,           # noqa: F401  (re-exported so callers can import from here)
    _resolve_pdf_path,
    _normalize_text,
)


# ---------------------------------------------------------------------------
# Axis Bank–specific constants
# ---------------------------------------------------------------------------

# Sentinel row values in the Description column that must be skipped
_SKIP_PARTICULARS = frozenset([
    "OPENING BALANCE",
    "CLOSING BALANCE",
    "TRANSACTION TOTAL",
])

# Minimum number of recognised columns required to accept a row as the header
_MIN_MATCHED_COLS = 4


# Fallback mapping used when no column_mapping is provided from the database.
# Keys are canonical names, values are lists of alias strings as they appear
# in Axis Bank PDFs (matching is case-insensitive).
_AXIS_DEFAULT_COLUMN_MAPPING: dict[str, list[str]] = {
    "DATE":              ["tran date", "date"],
    "VALUE_DATE":        ["value date", "val date"],
    "TRANSACTION_NO":    ["chq no", "instr no", "ref no"],
    "DESCRIPTION":       ["particulars", "narration", "description"],
    "WITHDRAWAL_AMOUNT": ["debit"],
    "DEPOSIT_AMOUNT":    ["credit"],
    "CLOSING_BALANCE":   ["balance"],
}


# ---------------------------------------------------------------------------
# Axis Bank–specific internal helpers
# ---------------------------------------------------------------------------

def _build_alias_lookup(column_mapping: dict[str, list[str]]) -> dict[str, str]:
    """
    Convert a DB-style column mapping  ``{canonical_key: [alias, ...]}``
    into a flat lookup ``{normalised_alias: canonical_key}`` for O(1) header
    cell matching inside ``_map_header_row``.

    Normalisation:  lowercase + whitespace collapse (via ``_collapse_whitespace``).
    When two aliases normalise to the same string the first canonical key wins.

    Args:
        column_mapping: Dict fetched from edocsmart_bank_mapping.mapping,
                        e.g. {"DATE": ["tran date"], "DESCRIPTION": ["particulars"]}.

    Returns:
        Flat lookup dict, e.g. {"tran date": "DATE", "particulars": "DESCRIPTION"}.
    """
    lookup: dict[str, str] = {}
    for canonical, aliases in column_mapping.items():
        for alias in aliases:
            normalised = _collapse_whitespace(str(alias)).lower()
            if normalised and normalised not in lookup:
                lookup[normalised] = canonical
    return lookup


def _map_header_row(raw_row: list, alias_lookup: dict[str, str]) -> dict[str, int] | None:
    """
    Inspect *raw_row* (the first row of a pdfplumber table) and return a
    mapping of ``canonical_key → column_index`` when it matches the expected
    header pattern.

    Each cell is normalised (whitespace-collapsed, lowercased) and looked up
    in *alias_lookup* (built from the DB column mapping via ``_build_alias_lookup``).

    Returns None when fewer than ``_MIN_MATCHED_COLS`` columns are recognised,
    so the caller knows this row is not a bank-statement header.
    """
    col_map: dict[str, int] = {}
    for idx, cell in enumerate(raw_row):
        normalised = _collapse_whitespace(str(cell or "")).lower()
        canonical = alias_lookup.get(normalised)
        if canonical and canonical not in col_map:   # keep first occurrence
            col_map[canonical] = idx

    if len(col_map) < _MIN_MATCHED_COLS:
        return None
    return col_map


def _extract_table_rows(
    raw_table: list[list],
    col_map: dict[str, int],
    original_pdf_headers: dict[str, str],
) -> list[dict]:
    """
    Convert a raw pdfplumber table (list of rows) into a list of canonical
    dicts using the column-index mapping produced by ``_map_header_row``.

    Skips:
      - Row 0  only when it is a header row (detected via ``_map_header_row``).
        On page 1 the first row is the column header; on subsequent pages it
        may already be a transaction row — so we never assume blindly.
      - Sentinel summary rows  (OPENING BALANCE, TRANSACTION TOTAL, CLOSING BALANCE)
      - Fully empty rows       (no date AND no debit AND no credit)

    Cell values are whitespace-collapsed via ``_collapse_whitespace`` so that
    multi-line descriptions are flattened into a single readable string.
    """
    rows: list[dict] = []

    # Skip row 0 only when it is actually a header row.
    # _map_header_row returns a col_map when enough header aliases are matched,
    # or None when the row looks like ordinary data (e.g. page 2 tables that
    # start directly with a transaction).
    _alias_lookup = _build_alias_lookup(_AXIS_DEFAULT_COLUMN_MAPPING)
    first_row = raw_table[0] if raw_table else []
    is_header_row = _map_header_row(first_row, _alias_lookup) is not None
    start_index = 1 if is_header_row else 0

    for raw_row in raw_table[start_index:]:
        record: dict[str, str] = {}
        for canonical, idx in col_map.items():
            cell_val = ""
            if idx < len(raw_row) and raw_row[idx] is not None:
                cell_val = _collapse_whitespace(str(raw_row[idx]))
            record[canonical] = cell_val

        # Drop sentinel / summary rows
        desc = record.get("DESCRIPTION", "")
        if desc.upper() in _SKIP_PARTICULARS:
            continue

        # Drop fully empty rows (no date and no financial amounts)
        if (
            not record.get("DATE")
            and not record.get("WITHDRAWAL_AMOUNT")
            and not record.get("DEPOSIT_AMOUNT")
        ):
            continue

        rows.append(record)

    return rows


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def is_axis_bank_pdf(pdf_path: str,column_mapping: dict | None = None) -> bool:
    """
    Quickly decide whether *pdf_path* is an Axis Bank structured statement.
    Uses PyMuPDF text extraction instead of pdfplumber.extract_tables().
    Scans the first 3 pages and searches for known column aliases.

    Args:
        pdf_path:       Path to the PDF file.
        column_mapping: DB column mapping
                        {canonical: [aliases]}.

    Returns:
        bool
    """
    active_mapping = (column_mapping if column_mapping else _AXIS_DEFAULT_COLUMN_MAPPING)
    _MIN_HEADER_MATCHES=3
    alias_lookup = _build_alias_lookup(active_mapping)
    try:
        doc = fitz.open(pdf_path)
        pages_to_check = min(3, len(doc))
        for page_no in range(pages_to_check):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue
            normalized_text = _normalize_text(text)
            matched_columns = []
            for alias in alias_lookup.keys():
                normalized_alias = _normalize_text(alias)

                if normalized_alias in normalized_text:
                    matched_columns.append(alias)

            # Same behaviour as _map_header_row()
            if len(matched_columns) >= _MIN_HEADER_MATCHES:
                doc.close()
                return True

        doc.close()

    except Exception as exc:

        print(
            "[is_axis_bank_pdf] Could not inspect PDF '%s': %s"
            % (pdf_path, exc)
        )

    return False



def extract_axis_bank_statement(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Extract all transaction rows from every page of an Axis Bank PDF.

    Uses pdfplumber's native ``extract_tables()`` API (drawn-border tables),
    which gives clean structured data without coordinate-based word grouping.

    Args:
        pdf_path:       Path to the PDF file.
        column_mapping: DB column mapping  ``{canonical: [aliases]}``.
                        Falls back to ``_AXIS_DEFAULT_COLUMN_MAPPING`` when None.

    Returns:
        (df, original_headers) where:
          - df               : DataFrame with canonical column names
          - original_headers : dict  canonical_key -> original PDF header text
                               e.g. {"DATE": "Tran Date", "DESCRIPTION": "Particulars"}

    Both are empty when no transactions are found.
    """
    active_mapping = column_mapping if column_mapping else _AXIS_DEFAULT_COLUMN_MAPPING
    alias_lookup   = _build_alias_lookup(active_mapping)

    all_rows:         list[dict] = []
    original_headers: dict       = {}
    col_map:          dict | None = None

    with pdfplumber.open(pdf_path) as pdf:
        print("  [Axis Bank] PDF has %d page(s)." % len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            if not tables:
                print("  [Axis Bank] Page %d: no tables found." % page_num)
                continue

            for table in tables:
                if not table:
                    continue

                # Try to detect the header on this table's first row
                detected = _map_header_row(table[0], alias_lookup)
                if detected is not None:
                    col_map = detected
                    # Capture original (PDF) header texts once
                    if not original_headers:
                        original_headers = {
                            canonical: _collapse_whitespace(str(table[0][idx] or ""))
                            for canonical, idx in col_map.items()
                        }
                    print(
                        "  [Axis Bank] Page %d: header detected -> columns=%s"
                        % (page_num, list(col_map.keys()))
                    )

                if col_map is not None:
                    page_rows = _extract_table_rows(table, col_map, original_headers)
                    print("  [Axis Bank] Page %d: extracted %d row(s)." % (page_num, len(page_rows)))
                    all_rows.extend(page_rows)
                else:
                    print(
                        "  [Axis Bank] Page %d: table found but header not recognised. Skipping."
                        % page_num
                    )

    if not all_rows:
        return pd.DataFrame(), original_headers

    df = pd.DataFrame(all_rows)
    df = df[[c for c in BANK_STATEMENT_CANONICAL_ORDER if c in df.columns]]
    df = df.reset_index(drop=True)

    return df, original_headers


def pdf_to_document_axis(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Public entry point that mirrors ``pdf_to_document()`` in structured_pdf_processor.py.

    Extracts transaction rows from an Axis Bank structured PDF using the
    column mapping fetched from the database.

    Args:
        pdf_path:       Absolute or relative path to the source PDF.
        column_mapping: DB column mapping  ``{canonical: [aliases]}``
                        fetched from edocsmart_bank_mapping.mapping.
                        Falls back to ``_AXIS_DEFAULT_COLUMN_MAPPING`` when None.

    Returns:
        (df, original_headers) where:
          - df               : DataFrame with canonical column names
          - original_headers : dict  canonical_key -> original PDF header text

    Raises:
        FileNotFoundError: if *pdf_path* does not exist (via ``_resolve_pdf_path``).
    """
    pdf_path = _resolve_pdf_path(pdf_path)

    print("\n[pdf_to_document_axis] Extracting transactions from: %s" % pdf_path)
    df, original_headers = extract_axis_bank_statement(pdf_path, column_mapping=column_mapping)

    if df.empty:
        print("  [!] No structured Axis Bank data found. PDF may be image-based.")
    else:
        print("  [OK] %d transaction row(s) extracted." % len(df))

    return df, original_headers
