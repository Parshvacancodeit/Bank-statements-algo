import re
import logging
import pandas as pd
import pdfplumber
from backend_common.constants import BANK_STATEMENT_CANONICAL_ORDER
import fitz
from .bank_processor import (
    _collapse_whitespace,
    _is_date,
    _resolve_pdf_path,
    _normalize_text,
    _group_words_by_line,
    _classify_word,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Bank of Baroda – constants
# ---------------------------------------------------------------------------

# Default column mapping: canonical key → list of header aliases as seen in the PDF
_BOB_DEFAULT_COLUMN_MAPPING: dict[str, list[str]] = {
    "DATE":               ["date"],
    "DESCRIPTION":        ["narration"],
    "TRANSACTION_NO":     ["chq.no.", "chq no", "cheque no"],
    "WITHDRAWAL_AMOUNT":  ["withdrawal", "withdrawal (dr)", "debit"],
    "DEPOSIT_AMOUNT":     ["deposit", "deposit (cr)", "credit"],
    "CLOSING_BALANCE":    ["balance"],
}

# Tolerance (points) for grouping words that share the same visual line
_LINE_Y_TOLERANCE: float = 3.0

# Maximum vertical distance (points) between a narration word and the
# nearest transaction's date line for the narration to be attached
_NARRATION_ATTACH_THRESHOLD: float = 30.0

# Maximum vertical distance for assigning amount values to a transaction row
_AMOUNT_ATTACH_THRESHOLD: float = 15.0


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------




def _detect_header_row(words: list[dict], column_mapping: dict) -> dict | None:
    """
    Scan all visual lines on a page and return metadata for the first line
    that looks like the transaction table header.

    A line qualifies as the header when at least 3 column aliases from
    *column_mapping* are found in its normalised text.

    Returns a dict with keys:
        - "top"    : minimum word top-y in the header line
        - "bottom" : maximum word bottom-y in the header line
        - "words"  : list of word dicts that make up the header line
    Returns None if no header is found.
    """
    lines = _group_words_by_line(words)

    # Build a flat set of normalised aliases to match against
    all_aliases: set[str] = set()
    for aliases in column_mapping.values():
        for alias in aliases:
            all_aliases.add(_normalize_text(alias))

    for line in lines:
        # Build a single normalised string for the whole line
        normalized_words = [_normalize_text(w["text"]) for w in line["words"] if w["text"].strip()]
        line_text = " ".join(normalized_words)

        # Count how many distinct aliases appear in the line text
        matched = [alias for alias in all_aliases if alias in line_text]
        if len(matched) >= 3:
            return {
                "top":    min(w["top"]    for w in line["words"]),
                "bottom": max(w["bottom"] for w in line["words"]),
                "words":  line["words"],
            }

    return None


def _build_column_ranges(header_words: list[dict]) -> dict[str, tuple[float, float]]:
    """
    Build a mapping of  column_name → (x_start, x_end)  from the header words.

    Each column's x-range spans from the midpoint between the previous and
    current header x0 (or 0 for the first column) to the midpoint between the
    current and next header x0 (or page-right for the last column).

    This midpoint strategy guarantees non-overlapping ranges so that narration
    words printed slightly right of the DATE header are still correctly
    classified as NARRATION rather than DATE.

    Special handling:
    - Sub-labels like "(DR)" / "(CR)" are merged with their parent header.
    - The last column extends 50 pt past the rightmost header word.

    Returns:
        dict mapping normalised column name → (x_start, x_end)
    """
    # Step 1: Sort header words left-to-right and normalise their text
    sorted_words = sorted(header_words, key=lambda w: w["x0"])
    columns = [
        {
            "text":   _normalize_text(w["text"]),
            "x0":     w["x0"],
            "x1":     w["x1"],
        }
        for w in sorted_words
    ]

    # Step 2: Merge "(DR)" / "(CR)" sub-labels with their preceding header
    merged: list[dict] = []
    skip_next = False
    for i, col in enumerate(columns):
        if skip_next:
            skip_next = False
            continue
        text = col["text"]
        x1   = col["x1"]
        # If the very next token is "dr" or "cr", absorb it
        if i + 1 < len(columns) and columns[i + 1]["text"] in ("dr", "cr"):
            text     = text + " " + columns[i + 1]["text"]
            x1       = columns[i + 1]["x1"]
            skip_next = True
        merged.append({"text": text, "x0": col["x0"], "x1": x1})

    # Step 3: Derive non-overlapping x ranges using midpoints between adjacent
    # header x0 values.  This ensures a narration word printed at x0=80 (which
    # is between the DATE header at x0=43 and the NARRATION header at x0=141)
    # is correctly classified as NARRATION rather than DATE.
    ranges: dict[str, tuple[float, float]] = {}
    for i, col in enumerate(merged):
        # First column always starts at x=0 to catch values left of the header
        if i == 0:
            x_start = 0.0
        else:
            # Midpoint between this column's x0 and the previous column's x0
            x_start = (merged[i - 1]["x0"] + col["x0"]) / 2.0

        # Last column extends well past the right edge of the page
        if i == len(merged) - 1:
            x_end = col["x1"] + 50.0
        else:
            # Midpoint between this column's x0 and the next column's x0
            x_end = (col["x0"] + merged[i + 1]["x0"]) / 2.0

        ranges[col["text"]] = (x_start, x_end)

    return ranges

# ---------------------------------------------------------------------------
# Bank-format detection
# ---------------------------------------------------------------------------

def _detect_header_structure(pdf_path: str, column_mapping: dict) -> bool:
    """
    Return True if any page of *pdf_path* contains at least 3 recognised
    Bank of Baroda column header aliases.
    """
    all_aliases: set[str] = set()
    for aliases in column_mapping.values():
        for alias in aliases:
            all_aliases.add(_collapse_whitespace(alias).lower())

    try:
        doc = fitz.open(pdf_path)
        for page_no in range(len(doc)):
            page = doc[page_no]
            text = page.get_text("text")
            if not text:
                continue
            normalized_text = _collapse_whitespace(text).lower()
            matched_aliases = [alias for alias in all_aliases if alias in normalized_text]
            logger.debug("[BOB] Page %s matched aliases: %s", page_no + 1, matched_aliases)
            if len(matched_aliases) >= 3:
                doc.close()
                return True

        doc.close()

    except Exception as exc:
        logger.warning("[BOB] Structure check error: %s", exc)

    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def is_bob_pdf(pdf_path: str, column_mapping: dict | None = None) -> bool:
    """
    Return True when *pdf_path* looks like a Bank of Baroda statement.
    Detection works by finding at least 3 recognised column header aliases
    (DATE, NARRATION, WITHDRAWAL, DEPOSIT, BALANCE) on any page.
    """

    pdf_path = _resolve_pdf_path(pdf_path)
    mapping = column_mapping or _BOB_DEFAULT_COLUMN_MAPPING
    result = _detect_header_structure(pdf_path, mapping)
    logger.info("[BOB] is_bob_pdf('%s') → %s", pdf_path, result)
    return result




def _extract_page_transactions(
    page,
    column_ranges: dict[str, tuple[float, float]],
    table_start_y: float,
    table_end_y: float,
) -> list[dict]:
    """
    Extract all transaction rows from a single PDF page.

    Strategy:
    1. Collect all words inside the table area (between table_start_y and table_end_y).
    2. Classify each word to a column using its x0 position.
    3. A word in the DATE column that matches a date pattern starts a new transaction.
    4. Narration words are buffered and later matched to the nearest transaction by
       vertical distance (centre-to-centre).
    5. Amount words (WITHDRAWAL, DEPOSIT, BALANCE) are attached to the nearest
       transaction within _AMOUNT_ATTACH_THRESHOLD points.

    Returns a list of transaction dicts (keys: DATE, NARRATION, CHQ_NO,
    WITHDRAWAL, DEPOSIT, BALANCE).
    """
    words = page.extract_words()

    transactions: list[dict]  = []
    narration_buffer: list[dict] = []
    current_txn: dict | None  = None

    for word in words:
        text   = word["text"].strip()
        x0     = word["x0"]
        top    = word["top"]
        bottom = word["bottom"]

        # Skip words outside the transaction table area
        if top < table_start_y or top > table_end_y:
            continue

        column = _classify_word(x0, column_ranges)
        if column is None:
            continue

        word_center = (top + bottom) / 2.0

        # ------------------------------------------------------------------
        # A valid date in the DATE column → start a new transaction row
        # ------------------------------------------------------------------
        if column == "date" and _is_date(text):
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

        # ------------------------------------------------------------------
        # Narration words – buffer ALL of them (even before the first date
        # row) for proximity-based assignment at the end.
        #
        # NOTE: BOB PDFs print narration text starting at x0≈80, which
        # falls inside the DATE column zone (0–92 pt).  We therefore treat
        # non-date words from BOTH the date and narration columns as
        # narration text.
        #
        # IMPORTANT: Do NOT skip narration words when current_txn is None.
        # The BOB layout places the narration line ONE ROW ABOVE the date
        # row, so narration words arrive before we have seen a date.
        # Proximity matching at the end correctly assigns them.
        # ------------------------------------------------------------------
        if column in ("narration", "date") and not _is_date(text):
            narration_buffer.append({"text": text, "center": word_center})
            continue

        # For non-narration columns there is nothing to do until a
        # transaction row has been started
        if current_txn is None:
            continue

        txn_center = current_txn["center"]
        distance   = abs(txn_center - word_center)

        # ------------------------------------------------------------------
        # CHQ/reference number
        # ------------------------------------------------------------------
        if column == "chq no" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            sep = " " if current_txn["TRANSACTION_NO"] else ""
            current_txn["TRANSACTION_NO"] += sep + text

        # ------------------------------------------------------------------
        # Amount columns – attach to the nearest transaction row
        # ------------------------------------------------------------------
        elif column == "withdrawal dr" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["WITHDRAWAL_AMOUNT"] = text

        elif column == "deposit cr" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            current_txn["DEPOSIT_AMOUNT"] = text

        elif column == "balance" and distance <= _AMOUNT_ATTACH_THRESHOLD:
            # "Cr" / "Dr" suffixes may appear right after the amount – skip them
            if re.fullmatch(r"[\d,]+\.?\d*", text):
                current_txn["CLOSING_BALANCE"] = text

    # ------------------------------------------------------------------
    # Assign each buffered narration word to the nearest transaction.
    #
    # Joining strategy: BOB PDFs hard-wrap long narration strings (e.g.
    # UPI payment references) mid-token across visual lines.  The first
    # line ends mid-word (e.g. "bhuvaharshpa") and the next line starts
    # with the remaining characters ("tel7915").  We therefore:
    #   - Use NO separator when the previous narration ends with a
    #     partial token (no trailing space / punctuation).
    #   - Use a space only when the previous text already ends cleanly
    #     (i.e. ends with a space, digit followed by letter, or similar).
    # ------------------------------------------------------------------
    for nword in narration_buffer:
        best_txn      = None
        best_distance = float("inf")
        for txn in transactions:
            dist = abs(txn["center"] - nword["center"])
            if dist < best_distance:
                best_distance = dist
                best_txn      = txn

        if best_txn is not None and best_distance <= _NARRATION_ATTACH_THRESHOLD:
            existing = best_txn["DESCRIPTION"]
            if not existing:
                # First narration fragment – no separator needed
                best_txn["DESCRIPTION"] = nword["text"]
            else:
                # Join without space: BOB wraps UPI strings mid-character.
                # A space is only inserted when the existing text ends with
                # a whitespace or when it clearly ends a complete word
                # (ends with digit/letter followed by a whitespace boundary).
                # For safety we always join without space — the raw UPI
                # reference is one continuous token.
                best_txn["DESCRIPTION"] = existing + nword["text"]

    # ------------------------------------------------------------------
    # Post-filter: remove opening/closing balance rows and empty rows
    # ------------------------------------------------------------------
    clean: list[dict] = []
    for txn in transactions:
        txn["DESCRIPTION"] = txn["DESCRIPTION"].strip()

        desc_lower = txn["DESCRIPTION"].lower()
        if desc_lower.startswith("opening") or desc_lower.startswith("closing"):
            continue

        # Skip rows that have no financial data at all
        if not txn["WITHDRAWAL_AMOUNT"] and not txn["DEPOSIT_AMOUNT"] and not txn["CLOSING_BALANCE"]:
            continue

        clean.append(txn)

    return clean


def extract_bob_statement(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Extract all transactions from a Bank of Baroda PDF statement.

    Args:
        pdf_path:       Path to the (unlocked) PDF file.
        column_mapping: Optional DB column mapping ``{canonical: [alias, ...]}``.
                        Falls back to ``_BOB_DEFAULT_COLUMN_MAPPING`` when None.

    Returns:
        (df, original_headers) where:
          - df               : pandas DataFrame with canonical column names
          - original_headers : dict  canonical_key → original PDF header text
    """
    pdf_path = _resolve_pdf_path(pdf_path)
    mapping  = column_mapping or _BOB_DEFAULT_COLUMN_MAPPING
    logger.info("[BOB] Extracting statement from '%s'", pdf_path)

    all_records: list[dict] = []
    original_headers: dict  = {}
    column_ranges:    dict  = {}  # built from the first page that has a header

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_words  = page.extract_words()

            # Detect the header row on this page.  Every page of a BOB
            # statement carries a repeated header so we re-detect it each
            # time rather than reusing page-1 coordinates.  This ensures
            # table_start_y is correct for pages 2, 3, … which begin
            # their data near y=80 rather than the page-1 value of y=373.
            header_info = _detect_header_row(page_words, mapping)

            if header_info is None:
                if not column_ranges:
                    # No header found on the very first page – cannot extract
                    logger.warning("[BOB] No header row detected in '%s'", pdf_path)
                    return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), {}
                # Subsequent page without a visible header – skip it
                logger.debug("[BOB] Page %d: no header, skipping", page_num)
                continue

            # Build column x-ranges once (all pages share the same layout)
            if not column_ranges:
                column_ranges = _build_column_ranges(header_info["words"])
                logger.debug("[BOB] Column ranges: %s", column_ranges)
                # Capture original header text for downstream use
                for word in header_info["words"]:
                    original_headers[word["text"]] = word["text"]

            # Table area on this page starts just below its own header
            table_start_y = header_info["bottom"] + 2.0
            table_end_y   = 9999.0

            page_records = _extract_page_transactions(
                page, column_ranges, table_start_y, table_end_y
            )
            logger.debug("[BOB] Page %d: %d transactions found", page_num, len(page_records))
            all_records.extend(page_records)

    if not all_records:
        logger.warning("[BOB] No transactions found in '%s'", pdf_path)
        return pd.DataFrame(columns=BANK_STATEMENT_CANONICAL_ORDER), {}

    df = pd.DataFrame(all_records)

    # Keep only the canonical columns that are present, in the correct order.
    # This mirrors the pattern used in axis_bank_pdf_processor.py and ensures
    # that the DataFrame schema is consistent across all bank processors.
    df = df[[c for c in BANK_STATEMENT_CANONICAL_ORDER if c in df.columns]]
    df = df.reset_index(drop=True)
    logger.info("[BOB] Total transactions extracted: %d", len(df))
    return df, original_headers


def pdf_to_document_bob(
    pdf_path: str,
    column_mapping: dict | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Public API entry point — matches the contract used by other bank processors.

    Returns:
        (df, original_headers) — see ``extract_bob_statement`` for details.
    """
    return extract_bob_statement(pdf_path, column_mapping=column_mapping)
