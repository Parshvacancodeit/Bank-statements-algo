# -*- coding: utf-8 -*-
"""
Kotak Mahindra Bank PDF Statement Extractor
=============================================
Dedicated extractor for Kotak Mahindra Bank structured PDF statements.

Kotak Mahindra Bank PDF layout observations:
  - pdfplumber's extract_tables() reliably returns the full transaction table
    including all data rows with proper cell boundaries.
  - Each page carries a repeated header:
        Row 0: "Savings Account Transactions" (merged title)
        Row 1: "#", "Date", "Description", "Chq/Ref. No.",
                "Withdrawal (Dr.)", "Deposit (Cr.)", "Balance"
  - The transaction table area starts at y ≈ 329.5 on page 1 (after account
    info / summary block) and at y ≈ 80 on subsequent pages.
  - Date format: DD MMM YYYY  (e.g. "01 Jan 2026", "03 Jan 2026").
  - Descriptions may wrap across lines; pdfplumber joins them with '\n'.
  - An "Opening Balance" sentinel row is present on page 1 with '-' values
    for #, Date, Chq/Ref, Withdrawal, and Deposit columns.
  - x0 positions (approximate):
        #                x ≈  39
        Date             x ≈  72
        Description      x ≈ 119
        Chq/Ref. No.     x ≈ 274
        Withdrawal (Dr.) x ≈ 396
        Deposit (Cr.)    x ≈ 460
        Balance          x ≈ 529

Header column names (exact PDF text):
    # | Date | Description | Chq/Ref. No. |
    Withdrawal (Dr.) | Deposit (Cr.) | Balance

Canonical output keys (project-wide standard):
    DATE | TRANSACTION_NO | DESCRIPTION | WITHDRAWAL_AMOUNT |
    DEPOSIT_AMOUNT | CLOSING_BALANCE

Public API:
    is_kotak_pdf(pdf_path, column_mapping)              -> bool
    extract_kotak_statement(pdf_path, column_mapping)    -> (DataFrame, headers)
    pdf_to_document_kotak(pdf_path, column_mapping)      -> (DataFrame, headers)

Shared utilities imported from bank_processor:
    _collapse_whitespace  -- whitespace normaliser
    _is_date              -- generic date-string detector
    _resolve_pdf_path     -- path resolution + FileNotFoundError guard
"""

import re
import logging
import pandas as pd
import pdfplumber
import fitz

logger = logging.getLogger(__name__)

from backend_common.constants import BANK_STATEMENT_CANONICAL_ORDER
from .bank_processor import (
    _collapse_whitespace,
    _is_date,
    _resolve_pdf_path,
    _normalize_text,
    _group_words_by_line,
    _classify_word,
)

# Default column mapping: canonical key → PDF header aliases (case-insensitive)
_KOTAK_DEFAULT_COLUMN_MAPPING: dict[str, list[str]] = {
    "DATE":               ["date"],
    "TRANSACTION_NO":     ["chq/ref. no.", "chq/ref no", "chq ref no", "cheque no", "ref no"],
    "DESCRIPTION":        ["description", "particulars", "narration"],
    "WITHDRAWAL_AMOUNT":  ["withdrawal (dr.)", "withdrawal dr", "withdrawal", "debit"],
    "DEPOSIT_AMOUNT":     ["deposit (cr.)", "deposit cr", "deposit", "credit"],
    "CLOSING_BALANCE":    ["balance"],
}

# Minimum aliases that must match to recognise a line as the header row
_MIN_HEADER_MATCHES = 3

# Vertical tolerance (pt) for grouping words on the same visual line
_LINE_Y_TOLERANCE: float = 3.0

# Maximum vertical gap (pt) for attaching amount values to a transaction row
_AMOUNT_ATTACH_THRESHOLD: float = 15.0

# Maximum vertical distance (pt) between a narration word and the
# nearest transaction's date line for the narration to be attached
_NARRATION_ATTACH_THRESHOLD: float = 30.0

# Sentinel description values that indicate summary rows to skip
_SKIP_DESCRIPTIONS = frozenset([
    "opening balance",
    "closing balance",
    "transaction total",
    "total",
])

# Kotak-specific keywords used to detect the bank's identity.
# The KKBK IFSC prefix is the most reliable marker and appears on page 1.
_KOTAK_IDENTITY_KEYWORDS = [
    "kotak mahindra bank",
    "kotak mahindra",
    "www.kotak.bank.in",
    "kotak.com",
    "kkbk",        # Kotak IFSC prefix (e.g. KKBK0000837)
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _is_kotak_date(text: str) -> bool:
    """
    Return True for Kotak date strings (DD MMM YYYY) as well as
    the generic _is_date patterns.

    Kotak dates: "01 Jan 2026", "03 Jan 2026", "18 May 2026"
    """
    text = text.strip()
    # DD MMM YYYY (Kotak-specific format)
    if re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{4}$", text):
        return True
    return _is_date(text)


def _detect_header_row(
    words: list[dict],
    column_mapping: dict,
) -> dict | None:
    """
    Scan page words for the first line that looks like the transaction-table
    header (at least _MIN_HEADER_MATCHES known column aliases present).

    Returns:
        dict with keys "top", "bottom", "words" — or None if not found.
    """
    # Build flat set of normalised aliases
    all_aliases: set[str] = set()
    for aliases in column_mapping.values():
        for alias in aliases:
            all_aliases.add(_normalize_text(alias))

    lines = _group_words_by_line(words)

    for i, line in enumerate(lines):
        # Kotak header is a single row, but gather nearby lines (within 12 pt)
        # in case of multi-line header wrapping
        candidate_words: list[dict] = list(line["words"])
        for j in range(1, 3):
            if i + j < len(lines) and (lines[i + j]["top"] - line["top"]) < 12:
                candidate_words.extend(lines[i + j]["words"])

        line_text = " ".join(_normalize_text(w["text"]) for w in candidate_words)
        matched = [alias for alias in all_aliases if alias in line_text]

        if len(matched) >= _MIN_HEADER_MATCHES:
            return {
                "top":    line["top"],
                "bottom": max(w["bottom"] for w in candidate_words),
                "words":  candidate_words,
            }

    return None


def _build_column_ranges(
    page,
    header_info: dict,
    column_mapping: dict,
) -> dict[str, tuple[float, float]]:
    """
    Build {normalised_alias: (x_start, x_end)} from the header word positions.

    First tries to detect ranges dynamically using table cells from page.find_tables().
    Falls back to a robust static layout mapping when no tables are found.
    """
    tables = page.find_tables()
    if tables:
        for table in tables:
            # Match the table that overlaps with the header region vertically
            if abs(table.bbox[1] - header_info["top"]) < 50 or abs(table.bbox[3] - header_info["bottom"]) < 50:
                ranges = _build_column_ranges_from_table(table, column_mapping)
                if len(ranges) >= 3:
                    logger.debug("[KOTAK] Built column ranges from table cells: %s", ranges)
                    return ranges

    logger.debug("[KOTAK] Table-based range detection failed; using fallback ranges")

    alias_to_canonical = {}
    for canonical, aliases in column_mapping.items():
        for alias in aliases:
            alias_to_canonical[_normalize_text(alias)] = canonical

    # Standard x-coordinate boundaries derived from Kotak PDF inspection
    canonical_coords = {
        "DATE":               ( 60.0, 119.0),
        "TRANSACTION_NO":     (260.0, 390.0),
        "DESCRIPTION":        (119.0, 260.0),
        "WITHDRAWAL_AMOUNT":  (390.0, 450.0),
        "DEPOSIT_AMOUNT":     (450.0, 520.0),
        "CLOSING_BALANCE":    (520.0, 570.0),
    }

    ranges = {}
    for alias, canonical in alias_to_canonical.items():
        if canonical in canonical_coords:
            ranges[alias] = canonical_coords[canonical]

    return ranges


def _build_column_ranges_from_table(
    table,
    column_mapping: dict,
) -> dict[str, tuple[float, float]]:
    """Build column x-ranges using table cell boundaries from pdfplumber."""
    alias_to_canonical = {}
    for canonical, aliases in column_mapping.items():
        for alias in aliases:
            alias_to_canonical[_normalize_text(alias)] = canonical

    extracted_rows = table.extract()
    if not extracted_rows:
        return {}

    # The header row is typically the second row (row index 1) because
    # row 0 is the "Savings Account Transactions" title
    header_text_row = None
    header_row_idx = None
    for idx, row in enumerate(extracted_rows):
        # Check if this row contains recognisable header aliases
        row_text = " ".join(_normalize_text(cell or "") for cell in row)
        matched = [alias for alias in alias_to_canonical if alias in row_text]
        if len(matched) >= _MIN_HEADER_MATCHES:
            header_text_row = row
            header_row_idx = idx
            break

    if header_text_row is None:
        return {}

    # Get the cells for the header row
    # Cells are (x0, top, x1, bottom) tuples sorted by position
    header_cells = sorted(table.cells, key=lambda c: (c[1], c[0]))

    # Group cells by row (same top coordinate within tolerance)
    rows_by_top: dict[float, list] = {}
    for cell in header_cells:
        placed = False
        for existing_top in rows_by_top:
            if abs(existing_top - cell[1]) < 5.0:
                rows_by_top[existing_top].append(cell)
                placed = True
                break
        if not placed:
            rows_by_top[cell[1]] = [cell]

    # Get the cells from the header row
    sorted_tops = sorted(rows_by_top.keys())
    if header_row_idx is not None and header_row_idx < len(sorted_tops):
        header_row_cells = sorted(rows_by_top[sorted_tops[header_row_idx]], key=lambda c: c[0])
    else:
        return {}

    ranges = {}
    for idx, cell in enumerate(header_row_cells):
        if idx >= len(header_text_row):
            break
        cell_text = _normalize_text(header_text_row[idx] or "")
        x0, y0, x1, y1 = cell

        # Check if cell_text matches any alias
        matched_canonical = None
        for alias, canonical in alias_to_canonical.items():
            if alias == cell_text or alias in cell_text or cell_text in alias:
                matched_canonical = canonical
                break

        if matched_canonical:
            ranges[cell_text] = (x0, x1)

    return ranges


def _match_cell_to_canonical(cell_text: str, norm_mapping: dict[str, list[str]]) -> str | None:
    """
    Match normalized cell text against normalized mapping aliases.
    Matches exact aliases as well as partial matches.
    """
    cell_norm = _normalize_text(cell_text)
    if not cell_norm:
        return None
    for canonical, norm_aliases in norm_mapping.items():
        for norm_alias in norm_aliases:
            if norm_alias == cell_norm or norm_alias in cell_norm:
                return canonical
    return None


# ---------------------------------------------------------------------------
# Table-based extraction (primary strategy for Kotak)
# ---------------------------------------------------------------------------

def _extract_transactions_from_table(table, column_mapping, col_indices=None) -> tuple[list[dict], dict | None]:
    """
    Extract transaction rows from a pdfplumber table object.

    Kotak tables are well-structured with clear cell boundaries, so we can
    rely on the table extraction directly rather than word-level parsing.

    Returns:
        tuple (transactions, col_indices) where:
          - transactions : list of transaction dicts with canonical keys.
          - col_indices  : dict mapping canonical keys to cell indices.
    """
    extracted_rows = table.extract()
    if not extracted_rows:
        return [], col_indices

    # Pre-normalize the aliases in column_mapping for fast lookup
    norm_mapping = {
        canonical: [_normalize_text(alias) for alias in aliases]
        for canonical, aliases in column_mapping.items()
    }

    # Find the header row to determine column indices
    header_idx = None
    new_col_indices = {}
    for idx, row in enumerate(extracted_rows):
        row_col_indices = {}
        for col_idx, cell in enumerate(row):
            if not cell:
                continue
            matched_canonical = _match_cell_to_canonical(cell, norm_mapping)
            if matched_canonical:
                row_col_indices[matched_canonical] = col_idx

        if len(row_col_indices) >= _MIN_HEADER_MATCHES:
            header_idx = idx
            new_col_indices = row_col_indices
            break

    if header_idx is not None:
        col_indices = new_col_indices
        start_row_idx = header_idx + 1
    else:
        if not col_indices:
            # Fallback to standard 7-column Kotak mapping if table has 7 columns
            if len(extracted_rows[0]) == 7:
                col_indices = {
                    "DATE": 1,
                    "DESCRIPTION": 2,
                    "TRANSACTION_NO": 3,
                    "WITHDRAWAL_AMOUNT": 4,
                    "DEPOSIT_AMOUNT": 5,
                    "CLOSING_BALANCE": 6,
                }
            else:
                return [], col_indices
        start_row_idx = 0

    transactions = []
    for row in extracted_rows[start_row_idx:]:
        # Skip rows where the # column is '-' or empty (sentinel/header rows)
        serial = (row[0] or "").strip() if row else ""
        if serial in ("-", ""):
            # Check if this might be a continuation row for description only
            desc_idx = col_indices.get("DESCRIPTION")
            if desc_idx is not None and transactions:
                desc_text = (row[desc_idx] or "").strip()
                if desc_text:
                    # Append continuation text to the previous transaction
                    existing = transactions[-1]["DESCRIPTION"]
                    if existing:
                        transactions[-1]["DESCRIPTION"] = existing + " " + desc_text
                    else:
                        transactions[-1]["DESCRIPTION"] = desc_text
            continue

        # Extract date
        date_idx = col_indices.get("DATE")
        date_val = (row[date_idx] or "").strip() if date_idx is not None else ""

        if not _is_kotak_date(date_val):
            continue

        # Extract description — clean up newlines from pdfplumber wrapping
        desc_idx = col_indices.get("DESCRIPTION")
        desc_val = (row[desc_idx] or "").strip() if desc_idx is not None else ""
        desc_val = re.sub(r"\s+", " ", desc_val)

        # Extract cheque/reference number
        txn_idx = col_indices.get("TRANSACTION_NO")
        txn_val = (row[txn_idx] or "").strip() if txn_idx is not None else ""

        # Extract amounts
        wd_idx = col_indices.get("WITHDRAWAL_AMOUNT")
        wd_val = (row[wd_idx] or "").strip() if wd_idx is not None else ""

        dep_idx = col_indices.get("DEPOSIT_AMOUNT")
        dep_val = (row[dep_idx] or "").strip() if dep_idx is not None else ""

        bal_idx = col_indices.get("CLOSING_BALANCE")
        bal_val = (row[bal_idx] or "").strip() if bal_idx is not None else ""

        # Skip sentinel rows
        if desc_val.lower() in _SKIP_DESCRIPTIONS:
            continue

        # Skip rows with no financial data
        if not wd_val and not dep_val and not bal_val:
            continue

        txn = {
            "DATE":              date_val,
            "DESCRIPTION":       desc_val,
            "TRANSACTION_NO":    txn_val,
            "WITHDRAWAL_AMOUNT": wd_val,
            "DEPOSIT_AMOUNT":    dep_val,
            "CLOSING_BALANCE":   bal_val,
        }
        transactions.append(txn)

    return transactions, col_indices


# ---------------------------------------------------------------------------
# Word-level fallback extraction
# ---------------------------------------------------------------------------

def _extract_page_transactions(
    page,
    column_ranges: dict[str, tuple[float, float]],
    table_start_y: float,
    table_end_y: float,
) -> list[dict]:
    """
    Fallback: Extract transaction rows from a single Kotak Bank PDF page
    using word-level x0 classification.

    This is used when pdfplumber's table extraction fails to return
    structured data.

    Returns:
        List of transaction dicts with canonical keys.
    """
    words = page.extract_words()

    transactions: list[dict]     = []
    narration_buffer: list[dict] = []
    current_txn: dict | None     = None

    # Map column range aliases to canonical keys
    _alias_to_canonical: dict[str, str] = {}
    for canonical, aliases in _KOTAK_DEFAULT_COLUMN_MAPPING.items():
        for alias in aliases:
            _alias_to_canonical[_normalize_text(alias)] = canonical

    for word in words:
        text   = word["text"].strip()
        x0     = word["x0"]
        top    = word["top"]
        bottom = word["bottom"]

        if top < table_start_y or top > table_end_y:
            continue

        col_alias = _classify_word(x0, column_ranges)
        if col_alias is None:
            continue

        canonical = _alias_to_canonical.get(col_alias, col_alias.upper())
        word_center = (top + bottom) / 2.0

        # A date word starts a new transaction
        if canonical == "DATE" and _is_kotak_date(text):
            current_txn = {
                "center":            word_center,
                "DATE":              text,
                "DESCRIPTION":       "",
                "TRANSACTION_NO":    "",
                "WITHDRAWAL_AMOUNT": "",
                "DEPOSIT_AMOUNT":    "",
                "CLOSING_BALANCE":   "",
            }
            transactions.append(current_txn)
            continue

        # Description / narration words — buffer for proximity assignment
        if canonical == "DESCRIPTION":
            narration_buffer.append({"text": text, "center": word_center})
            continue

        if current_txn is None:
            continue

        distance = abs(current_txn["center"] - word_center)

        # Cheque / reference number
        if canonical == "TRANSACTION_NO" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            sep = " " if current_txn["TRANSACTION_NO"] else ""
            current_txn["TRANSACTION_NO"] += sep + text

        # Amount columns
        elif canonical == "WITHDRAWAL_AMOUNT" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["WITHDRAWAL_AMOUNT"] = text

        elif canonical == "DEPOSIT_AMOUNT" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["DEPOSIT_AMOUNT"] = text

        elif canonical == "CLOSING_BALANCE" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            if re.fullmatch(r"[\d,]+\.?\d*", text):
                current_txn["CLOSING_BALANCE"] = text

    # Sort transactions by vertical position
    transactions = sorted(transactions, key=lambda t: t["center"])

    # Assign buffered narration words to their correct transaction
    for nword in narration_buffer:
        y = nword["center"]
        target_txn = None
        for i, txn in enumerate(transactions):
            y_curr = txn["center"]
            y_next = transactions[i+1]["center"] if i + 1 < len(transactions) else float("inf")
            if (y_curr - 6.0) <= y < (y_next - 6.0):
                target_txn = txn
                break

        # Fallback: close to the first transaction
        if target_txn is None and transactions:
            if y < transactions[0]["center"] - 6.0 and (transactions[0]["center"] - y) <= _NARRATION_ATTACH_THRESHOLD:
                target_txn = transactions[0]

        if target_txn is not None:
            existing = target_txn["DESCRIPTION"]
            if not existing:
                target_txn["DESCRIPTION"] = nword["text"]
            else:
                target_txn["DESCRIPTION"] = existing + " " + nword["text"]

    # Post-filter: remove sentinel / empty rows
    clean: list[dict] = []
    for txn in transactions:
        txn["DESCRIPTION"] = txn["DESCRIPTION"].strip()
        txn["DESCRIPTION"] = re.sub(r"\s+", " ", txn["DESCRIPTION"])
        desc_lower = txn["DESCRIPTION"].lower()

        if desc_lower in _SKIP_DESCRIPTIONS:
            continue

        if (
            not txn["WITHDRAWAL_AMOUNT"]
            and not txn["DEPOSIT_AMOUNT"]
            and not txn["CLOSING_BALANCE"]
        ):
            continue

        clean.append(txn)

    return clean


# ---------------------------------------------------------------------------
# Bank-format detection
# ---------------------------------------------------------------------------

def _detect_header_structure(pdf_path: str, column_mapping: dict) -> bool:
    """
    Return True if the PDF has a recognisable Kotak header AND contains a
    Kotak-specific identity marker.

    Scans up to 5 pages for the identity markers (the KKBK IFSC code appears
    on page 1, but "Kotak Mahindra Bank" may only appear in footer pages).
    Header-alias matching is restricted to the first 3 pages.
    """
    all_aliases: set[str] = set()
    for aliases in column_mapping.values():
        for alias in aliases:
            all_aliases.add(_normalize_text(alias))

    try:
        doc = fitz.open(pdf_path)
        pages_to_check_header = min(3, len(doc))
        pages_to_check_identity = min(5, len(doc))
        has_kotak_identity = False
        has_header_match = False

        # First pass: check identity across up to 5 pages
        for page_no in range(pages_to_check_identity):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue

            normalized_text = _normalize_text(text)

            # Check for Kotak-specific identity markers
            for keyword in _KOTAK_IDENTITY_KEYWORDS:
                if _normalize_text(keyword) in normalized_text:
                    has_kotak_identity = True
                    break

            # Check for header aliases (only in first 3 pages)
            if page_no < pages_to_check_header:
                matched = [
                    alias
                    for alias in all_aliases
                    if alias in normalized_text
                ]
                logger.debug("[KOTAK] Page %s matched aliases: %s", page_no + 1, matched)

                if len(matched) >= _MIN_HEADER_MATCHES:
                    has_header_match = True

            # Early exit if both conditions are met
            if has_kotak_identity and has_header_match:
                doc.close()
                return True

        doc.close()
    except Exception as exc:
        logger.warning("[KOTAK] Structure check error: %s", exc)

    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def is_kotak_pdf(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> bool:
    """
    Return True when pdf_path looks like a Kotak Mahindra Bank structured statement.

    Detection strategy:
    Look for at least _MIN_HEADER_MATCHES known column aliases AND
    a Kotak-specific identity keyword within the first 3 pages.
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping = (column_mapping or _KOTAK_DEFAULT_COLUMN_MAPPING)
    result = _detect_header_structure(pdf_path, mapping)
    logger.info("[KOTAK] is_kotak_pdf('%s') → %s", pdf_path, result)
    return result


def extract_kotak_statement(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Extract all transactions from a Kotak Mahindra Bank PDF statement.

    Primary strategy: Use pdfplumber's table extraction (Kotak PDFs have
    well-defined table borders).
    Fallback: Word-level x0 classification when table extraction fails.

    Args:
        pdf_path:       Path to the (unlocked) PDF file.
        column_mapping: Optional DB column mapping {canonical: [alias, ...]}.
                        Falls back to _KOTAK_DEFAULT_COLUMN_MAPPING when None.

    Returns:
        (df, original_headers) where:
          - df               : DataFrame with canonical column names
          - original_headers : dict  canonical_key → original PDF header text
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping  = column_mapping or _KOTAK_DEFAULT_COLUMN_MAPPING
    logger.info("[KOTAK] Extracting statement from '%s'", pdf_path)

    all_records:      list[dict] = []
    original_headers: dict       = {}
    column_ranges:    dict       = {}
    active_col_indices: dict     = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # ----------------------------------------------------------
            # Primary strategy: table-based extraction
            # ----------------------------------------------------------
            tables = page.find_tables()
            table_extracted = False

            for table in tables:
                table_records, active_col_indices = _extract_transactions_from_table(table, mapping, active_col_indices)
                if table_records:
                    logger.debug(
                        "[KOTAK] Page %d: %d transactions from table",
                        page_num, len(table_records),
                    )
                    all_records.extend(table_records)
                    table_extracted = True

                    # Capture original headers from the table
                    if not original_headers:
                        extracted_rows = table.extract()
                        for row in extracted_rows:
                            row_text = " ".join(_normalize_text(cell or "") for cell in row)
                            if "date" in row_text and "description" in row_text:
                                for cell in row:
                                    if cell and cell.strip():
                                        original_headers[cell.strip()] = cell.strip()
                                break

            if table_extracted:
                continue

            # ----------------------------------------------------------
            # Fallback: word-level extraction
            # ----------------------------------------------------------
            page_words = page.extract_words()
            header_info = _detect_header_row(page_words, mapping)

            if header_info is None:
                if not column_ranges:
                    logger.warning("[KOTAK] No header detected on page %d", page_num)
                    continue
                table_start_y = 50.0
            else:
                # Build column x-ranges once from the first detected header
                if not column_ranges:
                    column_ranges = _build_column_ranges(page, header_info, mapping)
                    logger.debug("[KOTAK] Column ranges: %s", column_ranges)
                    for word in header_info["words"]:
                        original_headers[word["text"]] = word["text"]
                table_start_y = header_info["bottom"] + 2.0

            table_end_y   = 810.0

            page_records = _extract_page_transactions(
                page, column_ranges, table_start_y, table_end_y
            )
            logger.debug("[KOTAK] Page %d: %d transactions (fallback)", page_num, len(page_records))
            all_records.extend(page_records)

    if not all_records:
        logger.warning("[KOTAK] No transactions found in '%s'", pdf_path)
        return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), {}

    df = pd.DataFrame(all_records)
    df = df[[c for c in BANK_STATEMENT_CANONICAL_ORDER if c in df.columns]]
    df = df.reset_index(drop=True)
    logger.info("[KOTAK] Total transactions extracted: %d", len(df))
    return df, original_headers


def pdf_to_document_kotak(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Public API entry point — matches the contract used by other bank processors.

    Returns:
        (df, original_headers) — see extract_kotak_statement for details.
    """
    return extract_kotak_statement(pdf_path, column_mapping=column_mapping)
