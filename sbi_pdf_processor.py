# -*- coding: utf-8 -*-
"""
SBI (State Bank of India) PDF Statement Extractor
===================================================
Dedicated extractor for SBI structured PDF statements.

SBI PDFs use native drawn-border tables (not free-floating text),
so pdfplumber's ``extract_tables()`` API is used instead of the
coordinate-based word approach used for HDFC / Kalupur statements.

SBI table column layout (7 columns, 0-indexed):
    Col 0: Txn Date       | Col 1: Value Date | Col 2: Description
    Col 3: Ref/Chq No     | Col 4: Debit      | Col 5: Credit
    Col 6: Balance

Note on header detection:
    SBI PDFs historically do not print a full column header row inside the drawn table.
    Every page starts with ``['', '', '', '', '', '', 'Balance']`` as row 0.
    In such classic cases, we fall back to positional index-based mapping.
    For newer/custom formats, the header columns are dynamically resolved using column_mapping.

Skipped sentinel rows:
    - Fully empty rows (no date AND no debit AND no credit)
    - Summary table rows (Statement Summary, Brought Forward, etc.)
    - Rows where the date cell is empty but no financial data is present
    - Dash-only rows (SBI uses "-" as a placeholder for empty amount cells)

Public API:
    is_sbi_pdf(pdf_path, column_mapping)              -> bool
    extract_sbi_statement(pdf_path, column_mapping)   -> (DataFrame, original_headers)
    pdf_to_document_sbi(pdf_path, column_mapping)     -> (DataFrame, original_headers)

Shared utilities imported from bank_processor:
    _collapse_whitespace  -- whitespace-only normaliser (case/punctuation preserved)
    _is_date              -- generic date-string detector
    _resolve_pdf_path     -- path resolution + FileNotFoundError guard
    _normalize_text       -- lowercase, strip, and remove special characters
"""

import logging
import difflib
import pandas as pd
import pdfplumber
import fitz
from backend_common.constants import BANK_STATEMENT_CANONICAL_ORDER

# Shared helpers from bank_processor
from .bank_processor import (
    _collapse_whitespace,
    _is_date,
    _resolve_pdf_path,
    _normalize_text,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# SBI-specific constants
# ---------------------------------------------------------------------------

# The original human-readable PDF header labels for each canonical key.
_SBI_ORIGINAL_HEADERS: dict[str, str] = {
    "DATE":               "Txn Date",
    "VALUE_DATE":         "Value Date",
    "DESCRIPTION":        "Description",
    "TRANSACTION_NO":     "Ref No./Cheque No.",
    "WITHDRAWAL_AMOUNT":  "Debit",
    "DEPOSIT_AMOUNT":     "Credit",
    "CLOSING_BALANCE":    "Balance",
}

# Expected number of columns in a valid SBI transaction table.
_SBI_EXPECTED_COLS = 7

# The "Balance" header text that SBI prints in column 6 of row 0 on each page.
_SBI_BALANCE_HEADER = "balance"

# Dash placeholder that SBI uses instead of empty amount cells.
_DASH_VALUE = "-"

# Canonical keys for the summary / totals table at the end of the statement.
_SKIP_SUMMARY_KEYWORDS = frozenset([
    "statement summary",
    "brought forward",
    "dr count",
    "cr count",
    "total debits",
    "total credits",
    "closing balance",
])

# Fallback mapping used when column_mapping is provided from the database.
_SBI_DEFAULT_COLUMN_MAPPING: dict[str, list[str]] = {
    "DATE":               ["txn date", "tran date", "date"],
    "VALUE_DATE":         ["value date", "val date"],
    "TRANSACTION_NO":     ["ref no", "chq no", "ref no./cheque no."],
    "DESCRIPTION":        ["description", "particulars", "narration", "transaction reference", "transaction details"],
    "WITHDRAWAL_AMOUNT":  ["debit", "withdrawal"],
    "DEPOSIT_AMOUNT":     ["credit", "deposit"],
    "CLOSING_BALANCE":    ["balance", "closing balance"],
}


# ---------------------------------------------------------------------------
# SBI-specific internal helpers
# ---------------------------------------------------------------------------

def _match_cell_to_canonical(
    cell_text: str,
    norm_mapping: dict[str, list[str]],
    threshold: float = 0.7,
) -> tuple[str, float, str] | None:
    """
    Match normalized cell text against normalized mapping aliases.
    Matches exact aliases, partial substring matches, and falls back to fuzzy matching.

    Returns:
        tuple (canonical_field, score, match_type) or None if no match found.
    """
    cell_norm = _normalize_text(cell_text)
    if not cell_norm:
        return None

    # 1. Exact match check
    for canonical, norm_aliases in norm_mapping.items():
        for norm_alias in norm_aliases:
            if norm_alias == cell_norm:
                return canonical, 1.0, "exact"

    # 2. Substring match check
    for canonical, norm_aliases in norm_mapping.items():
        for norm_alias in norm_aliases:
            if norm_alias in cell_norm or cell_norm in norm_alias:
                return canonical, 0.9, "partial"

    # 3. Fuzzy match fallback
    best_canonical = None
    best_score = 0.0
    for canonical, norm_aliases in norm_mapping.items():
        for norm_alias in norm_aliases:
            score = difflib.SequenceMatcher(None, cell_norm, norm_alias).ratio()
            if score >= threshold and score > best_score:
                best_canonical = canonical
                best_score = score

    if best_canonical:
        return best_canonical, best_score, "fuzzy"

    return None


def _match_sbi_headers(
    header_row: list,
    column_mapping: dict[str, list[str]],
    threshold: float = 0.7,
) -> dict[str, int]:
    """
    Dynamically map header row cells to canonical fields using column_mapping.
    Supports case-insensitive normalization, multiple aliases, and fuzzy fallback.
    """
    resolved_indices: dict[str, int] = {}
    detected_headers = []
    unmatched_columns = []

    # Normalize the aliases in column_mapping
    norm_mapping = {
        canonical: [_normalize_text(alias) for alias in aliases]
        for canonical, aliases in column_mapping.items()
    }

    # Match each cell to a canonical field
    for col_idx, cell in enumerate(header_row):
        if cell is None:
            continue
        cell_text = str(cell).strip()
        match_info = _match_cell_to_canonical(cell_text, norm_mapping, threshold)

        if match_info:
            canonical, score, match_type = match_info
            if canonical in resolved_indices:
                prev_idx = resolved_indices[canonical]
                logger.warning(
                    "[SBI] Duplicate match for %s. Already matched to col %d, now also col %d. Keeping col %d.",
                    canonical, prev_idx, col_idx, prev_idx
                )
            else:
                resolved_indices[canonical] = col_idx
                detected_headers.append(
                    f"Col {col_idx} ('{cell_text}') -> {canonical} ({match_type} match, score {score:.2f})"
                )
        else:
            unmatched_columns.append(f"Col {col_idx} ('{cell_text}')")

    # Logging output: only log reports if we have a valid header row match (>= 3 matches)
    if len(resolved_indices) >= 3:
        logger.info("[SBI] Dynamic Header Detection Report:")
        for det in detected_headers:
            logger.info("  [+] Resolved: %s", det)
        if unmatched_columns:
            logger.info("  [-] Unmatched: %s", ", ".join(unmatched_columns))

        # Check for missing mandatory fields
        mandatory = ["DATE", "DESCRIPTION", "CLOSING_BALANCE"]
        missing = [field for field in mandatory if field not in resolved_indices]
        if "WITHDRAWAL_AMOUNT" not in resolved_indices and "DEPOSIT_AMOUNT" not in resolved_indices:
            missing.append("WITHDRAWAL_AMOUNT/DEPOSIT_AMOUNT")

        if missing:
            logger.warning("[SBI] Missing mandatory fields for dynamic mapping: %s", ", ".join(missing))

    return resolved_indices


def _is_sbi_table(
    raw_table: list[list],
    column_mapping: dict,
    threshold: float = 0.7,
) -> tuple[bool, int | None, dict[str, int] | None, bool]:
    """
    Determine whether *raw_table* is a valid SBI transaction table.

    Returns:
        tuple (is_valid, header_row_idx, resolved_indices, is_classic_fallback)
    """
    if not raw_table or not raw_table[0]:
        return False, None, None, False

    # 1. Try to find a header row dynamically
    # Check the first 3 rows of the table for headers
    for idx in range(min(3, len(raw_table))):
        row = raw_table[idx]
        resolved_indices = _match_sbi_headers(row, column_mapping, threshold)

        # If we matched at least 3 canonical fields, consider this as a valid header row
        if len(resolved_indices) >= 3:
            return True, idx, resolved_indices, False

    # 2. Backward compatibility fallback: check for classic 7-column layout
    first_row = raw_table[0]
    if len(first_row) == _SBI_EXPECTED_COLS:
        last_cell = _collapse_whitespace(str(first_row[-1] or "")).lower()
        if last_cell == _SBI_BALANCE_HEADER:
            default_indices = {
                "DATE": 0,
                "VALUE_DATE": 1,
                "DESCRIPTION": 2,
                "TRANSACTION_NO": 3,
                "WITHDRAWAL_AMOUNT": 4,
                "DEPOSIT_AMOUNT": 5,
                "CLOSING_BALANCE": 6,
            }
            return True, 0, default_indices, True

    return False, None, None, False


def _is_summary_row(row: list) -> bool:
    """
    Return True if *row* belongs to the SBI statement summary table
    (the totals block printed at the very end of the PDF).

    Matches any row where at least one cell contains a known summary keyword.
    """
    for cell in row:
        if cell is None:
            continue
        cell_text = _collapse_whitespace(str(cell)).lower()
        for keyword in _SKIP_SUMMARY_KEYWORDS:
            if keyword in cell_text:
                return True
    return False


def _clean_amount(value: str) -> str:
    """
    Normalise an SBI amount cell.
    """
    cleaned = _collapse_whitespace(value)
    return "" if cleaned == _DASH_VALUE else cleaned


def _extract_sbi_table_rows(
    raw_table: list[list],
    col_indices: dict[str, int],
    start_row_idx: int,
) -> list[dict]:
    """
    Convert a raw pdfplumber SBI table (list of rows) into a list of canonical
    dicts using the dynamically resolved column index mapping.

    Skips:
      - Rows before start_row_idx.
      - Fully empty rows  — no date AND no debit AND no credit amount.
      - Summary / totals rows  — detected via ``_is_summary_row``.
    """
    rows: list[dict] = []

    for raw_row in raw_table[start_row_idx:]:
        # Skip summary / totals table rows
        if _is_summary_row(raw_row):
            continue

        record: dict[str, str] = {}
        for canonical, idx in col_indices.items():
            cell_val = ""
            if idx < len(raw_row) and raw_row[idx] is not None:
                raw_text = _collapse_whitespace(str(raw_row[idx]))
                # Normalise dash-placeholder amounts to empty string
                if canonical in ("WITHDRAWAL_AMOUNT", "DEPOSIT_AMOUNT", "CLOSING_BALANCE"):
                    cell_val = _clean_amount(raw_text)
                elif canonical == "TRANSACTION_NO":
                    cell_val = _clean_amount(raw_text)
                else:
                    cell_val = raw_text
            record[canonical] = cell_val

        # Drop fully empty rows (no date AND no financial amounts)
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

def is_sbi_pdf(
    pdf_path: str,
    column_mapping: dict | None = None
) -> bool:
    """
    Quickly determine whether the PDF appears to contain a structured
    bank statement table.
    """
    mapping = column_mapping or _SBI_DEFAULT_COLUMN_MAPPING
    structured_keywords = set()
    for aliases in mapping.values():
        for alias in aliases:
            val = alias.strip().lower()
            if val:
                structured_keywords.add(val)

    try:
        doc = fitz.open(pdf_path)
        pages_to_check = min(3, len(doc))
        for page_no in range(pages_to_check):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue
            text = text.lower()
            print(f"The text is {text}")
            matches = sum(1 for keyword in structured_keywords if keyword in text)
            print(f"The matches are {matches}")
            if matches >= 2:
                doc.close()
                return True
        doc.close()
    except Exception as exc:
        logger.warning("[is_sbi_pdf] Could not inspect PDF '%s': %s", pdf_path, exc)
    return False


def extract_sbi_statement(
    pdf_path: str,
    column_mapping: dict | None = None,
    fuzzy_threshold: float = 0.7,
) -> tuple[pd.DataFrame, dict]:
    """
    Extract all transaction rows from every page of an SBI PDF.

    Uses pdfplumber's native ``extract_tables()`` API (drawn-border tables).
    Column resolution is fully dynamic using `column_mapping`.
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping = column_mapping or _SBI_DEFAULT_COLUMN_MAPPING
    logger.info("[SBI] Extracting statement from '%s'", pdf_path)

    all_rows:         list[dict] = []
    original_headers: dict       = {}
    active_col_indices: dict     = {}
    active_num_cols: int | None  = None

    with pdfplumber.open(pdf_path) as pdf:
        logger.info("[SBI] PDF has %d page(s).", len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            if not tables:
                logger.info("[SBI] Page %d: no tables found.", page_num)
                continue

            for table in tables:
                if not table:
                    continue

                # Detect if table is SBI transaction table
                is_valid, header_idx, resolved_indices, is_classic = _is_sbi_table(
                    table, mapping, fuzzy_threshold
                )

                if not is_valid:
                    # Fallback to reusing active indices on subsequent pages
                    if active_col_indices and len(table[0]) == active_num_cols:
                        resolved_indices = active_col_indices
                        start_row_idx = 0
                        logger.info("[SBI] Page %d: Reusing active column mapping on header-less page.", page_num)
                    else:
                        logger.info(
                            "[SBI] Page %d: table found but not an SBI transaction table. Skipping.",
                            page_num
                        )
                        continue
                else:
                    if is_classic:
                        if active_col_indices:
                            resolved_indices = active_col_indices
                            start_row_idx = 1
                            logger.info("[SBI] Page %d: Reusing active mapping on classic fallback page.", page_num)
                        else:
                            active_col_indices = resolved_indices
                            active_num_cols = len(table[0])
                            start_row_idx = 1
                            logger.info("[SBI] Page %d: Set active mapping to default positional indices.", page_num)
                    else:
                        active_col_indices = resolved_indices
                        active_num_cols = len(table[0])
                        start_row_idx = header_idx + 1
                        logger.info("[SBI] Page %d: Discovered header row at index %d dynamically.", page_num, header_idx)

                # Capture original headers if they are not yet captured
                if resolved_indices and header_idx is not None and not original_headers:
                    header_row = table[header_idx]
                    for canonical, col_idx in resolved_indices.items():
                        if col_idx < len(header_row) and header_row[col_idx] is not None:
                            original_headers[canonical] = str(header_row[col_idx]).strip()

                page_rows = _extract_sbi_table_rows(table, resolved_indices, start_row_idx)
                logger.info("[SBI] Page %d: extracted %d transaction row(s).", page_num, len(page_rows))
                all_rows.extend(page_rows)

    if not all_rows:
        ret_headers = original_headers or _SBI_ORIGINAL_HEADERS.copy()
        return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), ret_headers

    df = pd.DataFrame(all_rows)
    df = df[[c for c in BANK_STATEMENT_CANONICAL_ORDER if c in df.columns]]
    df = df.reset_index(drop=True)

    return df, original_headers


def pdf_to_document_sbi(
    pdf_path: str,
    column_mapping: dict | None = None,
    fuzzy_threshold: float = 0.7,
) -> tuple[pd.DataFrame, dict]:
    """
    Public entry point that extracts SBI transactions.
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    logger.info("[pdf_to_document_sbi] Extracting transactions from: %s", pdf_path)
    df, original_headers = extract_sbi_statement(pdf_path, column_mapping=column_mapping, fuzzy_threshold=fuzzy_threshold)

    if df.empty:
        logger.warning("[pdf_to_document_sbi] No structured SBI data found. PDF may be empty/unstructured.")
    else:
        logger.info("[pdf_to_document_sbi] %d transaction row(s) extracted.", len(df))

    return df, original_headers
