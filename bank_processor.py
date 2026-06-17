# -*- coding: utf-8 -*-
"""
Bank Processor General Utilities
=================================
Common functions shared among Axis Bank, BOB, ICICI, Kotak, and Structured PDF processors.
"""

import os
import re
from pathlib import Path
from backend_common.constants import BANK_STATEMENT_CANONICAL_ORDER

# ---------------------------------------------------------------------------
# Helper: String and Date Normalisation
# ---------------------------------------------------------------------------

def _collapse_whitespace(text: str) -> str:
    """
    Collapse all whitespace sequences (including newlines) into a single space
    and strip leading/trailing whitespace.
    """
    if not text:
        return ""
    return " ".join(text.split())


def _normalize_text(text: str) -> str:
    """
    Lowercase the text, remove non-alphanumeric characters,
    collapse multiple spaces, and strip.
    """
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _resolve_pdf_path(pdf_path: str) -> str:
    """
    Resolve *pdf_path* to an absolute path and verify the file exists.
    """
    resolved = str(Path(pdf_path).resolve())
    if not os.path.exists(resolved):
        raise FileNotFoundError("PDF not found: %s" % resolved)
    return resolved


def _is_date(text: str) -> bool:
    """Check if the text represents a valid transaction date format."""
    text = text.strip()
    if not text:
        return False
    # 1. Matches DD/MM/YY or DD/MM/YYYY or DD-MM-YYYY (e.g. 01/05/26, 01/05/2026, 01-05-2026)
    if re.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$', text):
        return True
    # 2. Matches DD-MMM-YYYY or DD-MMM-YY (e.g. 02-Mar-2026, 02-Mar-26)
    if re.match(r'^\d{1,2}-[a-zA-Z]{3}-\d{2,4}$', text, re.IGNORECASE):
        return True
    # 3. Matches DD mmm YYYY or DD mmm YY (e.g. 02 Mar 2026, 2 Mar 26)
    if re.match(r'^\d{1,2}\s+[a-zA-Z]{3}\s+\d{2,4}$', text, re.IGNORECASE):
        return True
    return False


# ---------------------------------------------------------------------------
# Helper: Word and Column Classification
# ---------------------------------------------------------------------------

def _group_words_by_line(words: list[dict], y_tolerance: float = 3.0) -> list[dict]:
    """
    Group pdfplumber word dicts into visual lines by their 'top' position.
    Returns a list of line dicts sorted top-to-bottom.
    """
    lines: list[dict] = []
    for word in words:
        placed = False
        word_top = word["top"]
        for line in lines:
            if abs(line["top"] - word_top) <= y_tolerance:
                line["words"].append(word)
                placed = True
                break
        if not placed:
            lines.append({"top": word_top, "words": [word]})

    # Sort words within each line left-to-right
    for line in lines:
        line["words"] = sorted(line["words"], key=lambda w: w["x0"])

    return sorted(lines, key=lambda l: l["top"])


def _classify_word(x0: float, column_ranges: dict[str, tuple[float, float]]) -> str | None:
    """Return the first column name whose x-range contains x0, or None."""
    for col_name, (x_start, x_end) in column_ranges.items():
        if x_start <= x0 < x_end:
            return col_name
    return None
