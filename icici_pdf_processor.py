# -*- coding: utf-8 -*-
"""
ICICI Bank PDF Statement Extractor
====================================
Dedicated extractor for ICICI Bank structured PDF statements.

ICICI Bank PDF layout observations (from bankofbaroda.pdf analysis):
  - pdfplumber's extract_tables() returns only the 7-column header row; the
    actual transaction rows are free-floating text (not inside drawn borders).
  - Each transaction occupies 1 primary line + 1-to-N narration continuation
    lines printed below it.
  - Primary transaction line columns (approximate x0 positions):
        S.No.       x ≈  30
        Date        x ≈  61   (format: DD.MM.YYYY)
        (Remarks    x ≈ 192   — but the primary line has amounts, not remarks)
        Withdrawal  x ≈ 417
        Deposit     x ≈ 483
        Balance     x ≈ 535
  - Narration / remarks lines sit between transaction rows at x0 ≈ 192+.
  - Header row on page 1 is at y ≈ 270; pages 2+ repeat it at y ≈ 84–94.
  - Date hint: a "Credit trxn" / "Debit trxn" label may appear as a sub-line
    (used internally for debit/credit detection).

Header column names (exact PDF text):
    S No. | Transaction Date | Cheque Number | Transaction Remarks |
    Withdrawal Amount (INR) | Deposit Amount (INR) | Balance (INR)

Canonical output keys (project-wide standard):
    DATE | TRANSACTION_NO | DESCRIPTION | WITHDRAWAL_AMOUNT |
    DEPOSIT_AMOUNT | CLOSING_BALANCE

Public API:
    is_icici_pdf(pdf_path, column_mapping)              -> bool
    extract_icici_statement(pdf_path, column_mapping)   -> (DataFrame, headers)
    pdf_to_document_icici(pdf_path, column_mapping)     -> (DataFrame, headers)

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
_ICICI_DEFAULT_COLUMN_MAPPING: dict[str, list[str]] = {
    "DATE":               ["transaction date", "date", "txn date"],
    "TRANSACTION_NO":     ["cheque number", "chq no", "cheque no", "ref no"],
    "DESCRIPTION":        ["transaction remarks", "particulars", "narration", "remarks"],
    "WITHDRAWAL_AMOUNT":  ["withdrawal amount (inr)", "withdrawal", "debit", "dr"],
    "DEPOSIT_AMOUNT":     ["deposit amount (inr)", "deposit", "credit", "cr"],
    "CLOSING_BALANCE":    ["balance (inr)", "balance"],
}

# Minimum aliases that must match to recognise a line as the header row
_MIN_HEADER_MATCHES = 3

# Vertical tolerance (pt) for grouping words on the same visual line
_LINE_Y_TOLERANCE: float = 3.0

# Maximum vertical gap (pt) between a narration fragment and its transaction
_NARRATION_ATTACH_THRESHOLD: float = 50.0

# Maximum vertical gap (pt) for attaching amount values to a transaction row
_AMOUNT_ATTACH_THRESHOLD: float = 15.0

# X-coordinate threshold that separates left columns (S.No., Date) from
# narration/remark words.  Words with x0 >= this value on a non-date,
# non-amount line are treated as narration continuation.
_NARRATION_X_MIN: float = 150.0

# Approximate x0 ranges for each column (mid-page, from PDF inspection).
# These are used to classify words when no dynamic header is available.
# Format: (x_start, x_end)  — half-open intervals [x_start, x_end)
_ICICI_FALLBACK_COLUMN_RANGES: dict[str, tuple[float, float]] = {
    "date":                ( 55.0, 120.0),   # DD.MM.YYYY
    "cheque number":       (120.0, 185.0),   # cheque / ref number (often blank)
    "transaction remarks": (185.0, 395.0),   # narration / remarks
    "withdrawal amount (inr)": (395.0, 460.0),
    "deposit amount (inr)":    (460.0, 510.0),
    "balance (inr)":           (510.0, 600.0),
}

# Sentinel description values that indicate summary rows to skip
_SKIP_DESCRIPTIONS = frozenset([
    "opening balance",
    "closing balance",
    "transaction total",
    "total"
])


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

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

    # ICICI header spans 3 visual sub-lines — look for the aggregate match
    # by combining text from lines that are close together (within 15 pt)
    for i, line in enumerate(lines):
        # Gather text from this line and the 2 immediately following lines
        # (ICICI header uses 3 stacked rows for a single logical header)
        candidate_words: list[dict] = list(line["words"])
        for j in range(1, 4):
            if i + j < len(lines) and (lines[i + j]["top"] - line["top"]) < 15:
                candidate_words.extend(lines[i + j]["words"])

        line_text = " ".join(_normalize_text(w["text"]) for w in candidate_words)
        matched = [alias for alias in all_aliases if alias in line_text]

        if len(matched) >= _MIN_HEADER_MATCHES:
            # Return the span of the entire multi-row header
            return {
                "top":    line["top"],
                "bottom": max(w["bottom"] for w in candidate_words),
                "words":  candidate_words,
            }

    return None


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
    header_text_row = extracted_rows[0]

    header_cells = sorted(table.cells, key=lambda c: c[0])

    ranges = {}
    for idx, cell in enumerate(header_cells):
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


def _build_column_ranges(
    page,
    header_info: dict,
    column_mapping: dict,
) -> dict[str, tuple[float, float]]:
    """
    Build {normalised_alias: (x_start, x_end)} from the header word positions.

    First tries to detect ranges dynamically using table cells from page.find_tables().
    Falls back to a robust static layout mapping of aliases when no tables are found.
    """
    tables = page.find_tables()
    if tables:
        for table in tables:
            # Match the table that overlaps with the header region vertically
            if abs(table.bbox[1] - header_info["top"]) < 50 or abs(table.bbox[3] - header_info["bottom"]) < 50:
                ranges = _build_column_ranges_from_table(table, column_mapping)
                if len(ranges) >= 3:
                    logger.debug("[ICICI] Built column ranges from table cells: %s", ranges)
                    return ranges

    logger.debug("[ICICI] Table-based range detection failed; using fallback ranges")

    alias_to_canonical = {}
    for canonical, aliases in column_mapping.items():
        for alias in aliases:
            alias_to_canonical[_normalize_text(alias)] = canonical

    # Standard vertical alignment / coordinate boundaries of ICICI statements
    canonical_coords = {
        "DATE":               ( 49.0, 119.0),
        "TRANSACTION_NO":     (119.0, 189.0),
        "DESCRIPTION":        (189.0, 390.0),
        "WITHDRAWAL_AMOUNT":  (390.0, 456.0),
        "DEPOSIT_AMOUNT":     (456.0, 522.0),
        "CLOSING_BALANCE":    (522.0, 574.0),
    }

    ranges = {}
    for alias, canonical in alias_to_canonical.items():
        if canonical in canonical_coords:
            ranges[alias] = canonical_coords[canonical]

    return ranges




def _is_icici_date(text: str) -> bool:
    """
    Return True for ICICI date strings (DD.MM.YYYY or DD-MM-YYYY) as well as
    the generic _is_date patterns.
    """
    text = text.strip()
    if re.match(r"^\d{2}\.\d{2}\.\d{4}$", text):
        return True
    return _is_date(text)


# ---------------------------------------------------------------------------
# Page-level transaction extraction
# ---------------------------------------------------------------------------

def _extract_page_transactions(
    page,
    column_ranges: dict[str, tuple[float, float]],
    table_start_y: float,
    table_end_y: float,
) -> list[dict]:
    """
    Extract transaction rows from a single ICICI Bank PDF page.

    Strategy:
    1. Collect words inside the table area.
    2. A date word in the date column starts a new transaction row.
    3. Narration words are buffered and assigned to the correct transaction
       using the visual grid preceding rule (Y_i - 6.0 <= y < Y_next - 6.0).
    4. Amount / balance words are attached immediately.
    5. ICICI PDFs wrap narration across multiple lines — join with space
       to keep them readable and separated.

    Returns:
        List of transaction dicts with canonical keys.
    """
    words = page.extract_words()

    transactions: list[dict]     = []
    narration_buffer: list[dict] = []
    current_txn: dict | None     = None

    # Map column range aliases to canonical keys for clean assignment
    _alias_to_canonical: dict[str, str] = {}
    for canonical, aliases in _ICICI_DEFAULT_COLUMN_MAPPING.items():
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

        # -----------------------------------------------------------------
        # A date word starts a new transaction
        # -----------------------------------------------------------------
        if canonical == "DATE" and _is_icici_date(text):
            current_txn = {
                "center":           word_center,
                "DATE":             text,
                "DESCRIPTION":      "",
                "TRANSACTION_NO":   "",
                "WITHDRAWAL_AMOUNT": "",
                "DEPOSIT_AMOUNT":   "",
                "CLOSING_BALANCE":  "",
            }
            transactions.append(current_txn)
            continue

        # -----------------------------------------------------------------
        # Narration / remarks words — buffer ALL (even before first date)
        # because ICICI sometimes prints a narration line above the date row.
        # -----------------------------------------------------------------
        if canonical == "DESCRIPTION":
            narration_buffer.append({"text": text, "center": word_center})
            continue

        if current_txn is None:
            continue

        distance = abs(current_txn["center"] - word_center)

        # -----------------------------------------------------------------
        # Cheque / reference number
        # -----------------------------------------------------------------
        if canonical == "TRANSACTION_NO" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            sep = " " if current_txn["TRANSACTION_NO"] else ""
            current_txn["TRANSACTION_NO"] += sep + text

        # -----------------------------------------------------------------
        # Amount columns
        # -----------------------------------------------------------------
        elif canonical == "WITHDRAWAL_AMOUNT" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["WITHDRAWAL_AMOUNT"] = text

        elif canonical == "DEPOSIT_AMOUNT" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["DEPOSIT_AMOUNT"] = text

        elif canonical == "CLOSING_BALANCE" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            if re.fullmatch(r"[\d,]+\.?\d*", text):
                current_txn["CLOSING_BALANCE"] = text

    # Sort transactions by center y-coordinate for range mapping
    transactions = sorted(transactions, key=lambda t: t["center"])

    # ------------------------------------------------------------------
    # Assign buffered narration words to their correct transaction using
    # the visual grid preceding rule.
    # ------------------------------------------------------------------
    for nword in narration_buffer:
        y = nword["center"]
        target_txn = None
        for i, txn in enumerate(transactions):
            y_curr = txn["center"]
            y_next = transactions[i+1]["center"] if i + 1 < len(transactions) else float("inf")
            if (y_curr - 6.0) <= y < (y_next - 6.0):
                target_txn = txn
                break

        # Fallback to closest if slightly above the first transaction but close to it
        if target_txn is None and transactions:
            if y < transactions[0]["center"] - 6.0 and (transactions[0]["center"] - y) <= 15.0:
                target_txn = transactions[0]

        if target_txn is not None:
            existing = target_txn["DESCRIPTION"]
            if not existing:
                target_txn["DESCRIPTION"] = nword["text"]
            else:
                target_txn["DESCRIPTION"] = existing + " " + nword["text"]

    # ------------------------------------------------------------------
    # Post-filter: remove sentinel / empty rows
    # ------------------------------------------------------------------
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
    Return True if any of the first 3 pages has a recognisable ICICI header.

    Uses PyMuPDF instead of pdfplumber for much faster text extraction.
    """
    all_aliases: set[str] = set()
    for aliases in column_mapping.values():
        for alias in aliases:
            all_aliases.add(
                _normalize_text(alias)
            )
    try:
        doc = fitz.open(pdf_path)
        pages_to_check = min(3, len(doc))
        for page_no in range(pages_to_check):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue
            normalized_text = _normalize_text(text)
            matched = [
                alias
                for alias in all_aliases
                if alias in normalized_text
            ]
            logger.debug("[ICICI] Page %s matched aliases: %s",page_no + 1,matched)
            if len(matched) >= _MIN_HEADER_MATCHES:
                doc.close()
                return True
        doc.close()
    except Exception as exc:
        logger.warning(
            "[ICICI] Structure check error: %s",
            exc
        )
    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def is_icici_pdf(
    pdf_path: str,
    column_mapping: dict | None = None
) -> bool:
    """
    Return True when pdf_path looks like an ICICI Bank structured statement.
    Detection strategy:
    Look for at least _MIN_HEADER_MATCHES known column aliases
    within the first 3 pages.
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping = (column_mapping or _ICICI_DEFAULT_COLUMN_MAPPING)
    result = _detect_header_structure(pdf_path,mapping)
    logger.info("[ICICI] is_icici_pdf('%s') → %s",pdf_path,result)
    return result


def extract_icici_statement(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Extract all transactions from an ICICI Bank PDF statement.

    Args:
        pdf_path:       Path to the (unlocked) PDF file.
        column_mapping: Optional DB column mapping {canonical: [alias, ...]}.
                        Falls back to _ICICI_DEFAULT_COLUMN_MAPPING when None.

    Returns:
        (df, original_headers) where:
          - df               : DataFrame with canonical column names
          - original_headers : dict  canonical_key → original PDF header text
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping  = column_mapping or _ICICI_DEFAULT_COLUMN_MAPPING
    logger.info("[ICICI] Extracting statement from '%s'", pdf_path)

    all_records:      list[dict] = []
    original_headers: dict       = {}
    column_ranges:    dict       = {}  # built from the first header found

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_words = page.extract_words()

            # Re-detect the header on every page (ICICI repeats it)
            header_info = _detect_header_row(page_words, mapping)

            if header_info is None:
                if not column_ranges:
                    logger.warning("[ICICI] No header detected in '%s'", pdf_path)
                    return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), {}
                logger.debug("[ICICI] Page %d: no header, skipping", page_num)
                continue

            # Build column x-ranges once from the first detected header
            if not column_ranges:
                column_ranges = _build_column_ranges(page, header_info, mapping)
                logger.debug("[ICICI] Column ranges: %s", column_ranges)
                for word in header_info["words"]:
                    original_headers[word["text"]] = word["text"]

            # Table area on this page starts just below its header
            table_start_y = header_info["bottom"] + 2.0
            table_end_y   = 750.0

            page_records = _extract_page_transactions(
                page, column_ranges, table_start_y, table_end_y
            )
            logger.debug("[ICICI] Page %d: %d transactions", page_num, len(page_records))
            all_records.extend(page_records)

    if not all_records:
        logger.warning("[ICICI] No transactions found in '%s'", pdf_path)
        return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), {}

    df = pd.DataFrame(all_records)
    df = df[[c for c in BANK_STATEMENT_CANONICAL_ORDER if c in df.columns]]
    df = df.reset_index(drop=True)
    logger.info("[ICICI] Total transactions extracted: %d", len(df))
    return df, original_headers


def pdf_to_document_icici(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Public API entry point — matches the contract used by other bank processors.

    Returns:
        (df, original_headers) — see extract_icici_statement for details.
    """
    return extract_icici_statement(pdf_path, column_mapping=column_mapping)
