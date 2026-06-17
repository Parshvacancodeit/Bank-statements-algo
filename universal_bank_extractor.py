"""
universal_bank_extractor.py
============================
A bank-agnostic PDF bank-statement extractor.

Tested banks (all 9 PDFs)
--------------------------
  Bank of Baroda  ✓   Indian Bank     ✓   HDFC Bank  ✓
  ICICI Bank      ✓   Axis Bank       ✓   Kotak Bank ✓
  SBI (layout 1)  ✓   SBI (layout 2)  ✓

Architecture layers
-------------------
  0  PDF word extraction        pdfplumber (x/y tolerance)
  1  Line reconstruction        Y-band clustering → visual rows
  2  Header detection           Multi-row header assembly + alias scoring
  3  Column range mapping       Midpoint strategy → non-overlapping x-zones
  4  Special layout detection   SBI headerless, Kotak multi-word date,
                                Axis trailing-column strip
  5  Date anchor detection      Regex covering all Indian bank date formats
  6  Transaction assembly       Narration buffer + proximity matching
  7  Normalisation              Canonical schema, numeric coercion, metadata

Usage
-----
    from universal_bank_extractor import extract_statement
    df, meta = extract_statement("path/to/statement.pdf")
"""

from __future__ import annotations

import re
import logging
from pathlib import Path
from typing import Any
import numpy as np

import pdfplumber
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Canonical output schema
# ---------------------------------------------------------------------------

CANONICAL_COLUMNS = [
    "DATE",
    "DESCRIPTION",
    "TRANSACTION_NO",
    "WITHDRAWAL_AMOUNT",
    "DEPOSIT_AMOUNT",
    "CLOSING_BALANCE",
]

# ---------------------------------------------------------------------------
# Universal column alias registry
# ---------------------------------------------------------------------------

_COLUMN_ALIASES: dict[str, list[str]] = {
    "DATE": [
        "tran date", "trans date", "txn date", "transaction date",
        "value date", "posting date", "post date", "date",
    ],
    "DESCRIPTION": [
        "transaction remarks", "transaction narration", "trans particulars",
        "particulars", "narration", "description", "details", "remarks",
    ],
    "TRANSACTION_NO": [
        "chq/ref. no.", "chq/ref no", "chq.no.", "chq no",
        "cheque number", "cheque no", "ref no", "ref.no.",
        "instrument no", "instr no",
    ],
    "WITHDRAWAL_AMOUNT": [
        "withdrawal (dr.)", "withdrawal (dr)", "withdrawal amount (inr)",
        "withdrawal", "debit amount", "dr amount", "debit", "dr",
    ],
    "DEPOSIT_AMOUNT": [
        "deposit (cr.)", "deposit (cr)", "deposit amount (inr)",
        "deposit", "credit amount", "cr amount", "credit", "cr",
    ],
    "CLOSING_BALANCE": [
        "closing balance", "running balance", "avl bal",
        "available balance", "balance (inr)", "balance",
    ],
}

_MIN_HEADER_MATCHES = 3
_LINE_Y_TOL: float = 5.0
_NARRATION_ATTACH_MAX: float = 40.0
_AMOUNT_ATTACH_MAX: float = 20.0

# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

_DATE_PATTERNS = [
    re.compile(r"^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}$"),
    re.compile(r"^\d{1,2}[-/]\w{3}[-/]\d{2,4}$", re.IGNORECASE),
    re.compile(r"^\d{1,2}\s+\w{3}\s+\d{4}$", re.IGNORECASE),
    re.compile(r"^\d{4}[-/]\d{2}[-/]\d{2}$"),
]
_AMOUNT_RE = re.compile(r"^[\d,]+\.?\d*$")
_CRDR_RE   = re.compile(r"\s*(cr|dr)$", re.IGNORECASE)
_DASH_RE   = re.compile(r"^[-–—]+$")
_SERIAL_RE = re.compile(r"^[\d#]+$")


def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.lower().strip())

def _is_date(t: str) -> bool:
    return any(p.match(t.strip()) for p in _DATE_PATTERNS)

def _is_amount(t: str) -> bool:
    c = _CRDR_RE.sub("", t.replace(",", "")).strip()
    return bool(_AMOUNT_RE.match(c)) and bool(c)

def _clean_amount(t: str) -> str:
    return _CRDR_RE.sub("", t.replace(",", "")).strip()


def _sanitize_value(v: Any) -> str:
    """Convert non-scalar values to a string representation safe for DataFrame."""
    if v is None:
        return ""
    # numpy arrays or other sequences
    if isinstance(v, (list, tuple, set, np.ndarray)):
        try:
            return " ".join(str(x).strip() for x in v)
        except Exception:
            return str(v)
    return str(v).strip()


# ---------------------------------------------------------------------------
# Layer 1 – Line reconstruction
# ---------------------------------------------------------------------------

def _group_into_lines(words: list[dict]) -> list[dict]:
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, current = [], [sorted_words[0]]
    for w in sorted_words[1:]:
        pc = (current[-1]["top"] + current[-1]["bottom"]) / 2
        wc = (w["top"] + w["bottom"]) / 2
        if abs(wc - pc) <= _LINE_Y_TOL:
            current.append(w)
        else:
            lines.append(_make_line(current))
            current = [w]
    lines.append(_make_line(current))
    return lines

def _make_line(words: list[dict]) -> dict:
    ws = sorted(words, key=lambda w: w["x0"])
    return {
        "y_center": sum((w["top"]+w["bottom"])/2 for w in ws) / len(ws),
        "top":      min(w["top"]    for w in ws),
        "bottom":   max(w["bottom"] for w in ws),
        "words":    ws,
    }


# ---------------------------------------------------------------------------
# Layer 2 – Header detection (handles multi-row headers)
# ---------------------------------------------------------------------------

def _build_alias_lookup() -> dict[str, str]:
    lkp = {}
    for can, aliases in _COLUMN_ALIASES.items():
        for a in aliases:
            lkp[_norm(a)] = can
    return dict(sorted(lkp.items(), key=lambda x: -len(x[0])))

_ALIAS_LKP = _build_alias_lookup()


def _score_block(lines: list[dict]) -> tuple[int, dict[str, str]]:
    all_words = [w for ln in lines for w in ln["words"]]
    combined  = " ".join(_norm(w["text"]) for w in all_words)
    found: dict[str, str] = {}
    for alias, can in _ALIAS_LKP.items():
        if can in found:
            continue
        if alias in combined:
            for w in all_words:
                if alias in _norm(w["text"]):
                    found[can] = w["text"]
                    break
            if can not in found:
                found[can] = alias
    return len(found), found


def _detect_header(lines: list[dict]) -> dict | None:
    best_score, best_found, best_lines = 0, {}, []
    for i in range(len(lines)):
        for w in range(1, 4):
            block = lines[i:i+w]
            if block[-1]["bottom"] - block[0]["top"] > 70:
                break
            score, found = _score_block(block)
            if score > best_score:
                best_score, best_found, best_lines = score, found, block
    if best_score >= _MIN_HEADER_MATCHES:
        return {
            "top":             best_lines[0]["top"],
            "bottom":          best_lines[-1]["bottom"],
            "words":           [w for ln in best_lines for w in ln["words"]],
            "canonical_found": best_found,
        }
    return None


# ---------------------------------------------------------------------------
# Layer 3 – Column range mapping
# ---------------------------------------------------------------------------

def _build_column_ranges(
    header_words: list[dict],
    canonical_found: dict[str, str],
    rightmost_x: float = 600.0,
) -> dict[str, tuple[float, float]]:
    anchors: dict[str, float] = {}
    for can, orig in canonical_found.items():
        norm_o = _norm(orig)
        for w in sorted(header_words, key=lambda w: w["x0"]):
            if norm_o in _norm(w["text"]) or _norm(w["text"]) in norm_o:
                if can not in anchors:
                    anchors[can] = w["x0"]
                break
        if can not in anchors:
            for alias, c in _ALIAS_LKP.items():
                if c != can:
                    continue
                for w in sorted(header_words, key=lambda w: w["x0"]):
                    if alias in _norm(w["text"]):
                        anchors[can] = w["x0"]
                        break
                if can in anchors:
                    break
    if not anchors:
        return {}

    sorted_cols = sorted(anchors.items(), key=lambda x: x[1])
    ranges: dict[str, tuple[float, float]] = {}
    for i, (can, x0) in enumerate(sorted_cols):
        xs = 0.0 if i == 0 else (sorted_cols[i-1][1] + x0) / 2.0
        if i == len(sorted_cols) - 1:
            # Cap the last column at the actual rightmost header word x1
            max_x1 = max((w["x1"] for w in header_words), default=x0 + 60)
            xe = max_x1 + 5.0   # tight cap — no over-extension
        else:
            xe = (x0 + sorted_cols[i+1][1]) / 2.0
        ranges[can] = (xs, xe)
    return ranges


def _classify(x0: float, ranges: dict[str, tuple[float, float]]) -> str | None:
    for can, (xs, xe) in ranges.items():
        if xs <= x0 < xe:
            return can
    return None


# ---------------------------------------------------------------------------
# Layer 4 – Special layout handlers
# ---------------------------------------------------------------------------

# ---- SBI headerless layout ------------------------------------------------
# SBI statements show only "Balance" as a heading.  Positions are fixed:
#   x≈27  → post_date / value_date (both dates, we use first)
#   x≈82  → second date (skip — same as first usually)
#   x≈138 → narration / description
#   x≈304 → dash placeholder (skip)
#   x≈350 → WITHDRAWAL or DEPOSIT (determined by sign/position relative to
#            the dash columns: left dash=debit marker, right dash=credit marker)
#   x≈446 → second dash placeholder
#   x≈512 → CLOSING_BALANCE
#
# The debit vs credit distinction:
#   If the amount sits between dash1 (x≈304) and dash2 (x≈446) → WITHDRAWAL
#   If the amount sits right of dash2 (x≈446) but left of balance → DEPOSIT
#   Balance is always the rightmost column (x≈512)

_SBI_DATE_X    = (15, 70)
_SBI_DATE2_X   = (70, 128)       # second date (skip)
_SBI_DESC_X    = (128, 295)
_SBI_DEBIT_X   = (295, 405)      # WDL amount: x≈350 (between dash1 and dash2)
_SBI_CREDIT_X  = (405, 505)      # DEP amount: x≈430 (after dash2)
_SBI_BAL_X     = (505, 600)


def _extract_sbi_transactions(page, table_start_y: float) -> list[dict]:
    """Special extractor for SBI's fixed-column headerless layout."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    transactions: list[dict] = []
    narration_buffer: list[dict] = []
    current_txn: dict | None = None

    for word in words:
        text   = word["text"].strip()
        x0     = word["x0"]
        top    = word["top"]
        bottom = word["bottom"]

        if not text or top < table_start_y:
            continue
        if _DASH_RE.match(text):
            continue

        wc = (top + bottom) / 2.0

        # Date column
        if _SBI_DATE_X[0] <= x0 < _SBI_DATE_X[1]:
            if _is_date(text):
                current_txn = {
                    "center":             wc,
                    "DATE":               text,
                    "DESCRIPTION":        "",
                    "TRANSACTION_NO":     "",
                    "WITHDRAWAL_AMOUNT":  "",
                    "DEPOSIT_AMOUNT":     "",
                    "CLOSING_BALANCE":    "",
                }
                transactions.append(current_txn)
            continue

        # Second date — skip
        if _SBI_DATE2_X[0] <= x0 < _SBI_DATE2_X[1] and _is_date(text):
            continue

        # Description / narration
        if _SBI_DESC_X[0] <= x0 < _SBI_DESC_X[1]:
            # Skip transaction-type markers like "WDL TFR", "DEP TFR", branch codes
            if not re.match(r"^(WDL|DEP|TFR|AT|[0-9]{10,})$", text):
                narration_buffer.append({"text": text, "center": wc})
            continue

        if current_txn is None:
            continue
        dist = abs(current_txn["center"] - wc)

        # Debit column
        if _SBI_DEBIT_X[0] <= x0 < _SBI_DEBIT_X[1]:
            if _is_amount(text) and dist <= _AMOUNT_ATTACH_MAX:
                current_txn["WITHDRAWAL_AMOUNT"] = _clean_amount(text)
            continue

        # Credit column
        if _SBI_CREDIT_X[0] <= x0 < _SBI_CREDIT_X[1]:
            if _is_amount(text) and dist <= _AMOUNT_ATTACH_MAX:
                current_txn["DEPOSIT_AMOUNT"] = _clean_amount(text)
            continue

        # Balance column
        if _SBI_BAL_X[0] <= x0 < _SBI_BAL_X[1]:
            if _is_amount(text) and dist <= _AMOUNT_ATTACH_MAX:
                current_txn["CLOSING_BALANCE"] = _clean_amount(text)
            continue

    _flush_narration(narration_buffer, transactions)
    return _filter_rows(transactions)


# ---- Kotak multi-word date ------------------------------------------------

def _kotak_date_from_line(words_in_line: list[dict], date_xs: float, date_xe: float) -> str | None:
    date_words = [w for w in words_in_line if date_xs <= w["x0"] < date_xe]
    texts = [w["text"] for w in sorted(date_words, key=lambda w: w["x0"])]
    for size in (3, 2, 1):
        for start in range(len(texts) - size + 1):
            attempt = " ".join(texts[start:start + size])
            if _is_date(attempt):
                return attempt
    return None


# ---------------------------------------------------------------------------
# Layer 5 – Core per-page extraction
# ---------------------------------------------------------------------------

def _assign_word(can, text, wc, current_txn, transactions, narration_buffer):
    if can == "DESCRIPTION":
        narration_buffer.append({"text": text, "center": wc})
        return
    if can in ("WITHDRAWAL_AMOUNT", "DEPOSIT_AMOUNT", "CLOSING_BALANCE"):
        if not _is_amount(text):
            narration_buffer.append({"text": text, "center": wc})
            return
    if current_txn is None:
        return
    dist = abs(current_txn["center"] - wc)
    if can == "TRANSACTION_NO" and dist <= _AMOUNT_ATTACH_MAX:
        sep = " " if current_txn["TRANSACTION_NO"] else ""
        current_txn["TRANSACTION_NO"] += sep + text
    elif can == "WITHDRAWAL_AMOUNT" and dist <= _AMOUNT_ATTACH_MAX:
        current_txn["WITHDRAWAL_AMOUNT"] = _clean_amount(text)
    elif can == "DEPOSIT_AMOUNT" and dist <= _AMOUNT_ATTACH_MAX:
        current_txn["DEPOSIT_AMOUNT"] = _clean_amount(text)
    elif can == "CLOSING_BALANCE" and dist <= _AMOUNT_ATTACH_MAX:
        current_txn["CLOSING_BALANCE"] = _clean_amount(text)


def _flush_narration(narration_buffer, transactions):
    for nw in narration_buffer:
        if not transactions:
            continue
        best = min(transactions, key=lambda t: abs(t["center"] - nw["center"]))
        if abs(best["center"] - nw["center"]) > _NARRATION_ATTACH_MAX:
            continue
        ex = best["DESCRIPTION"]
        if not ex:
            best["DESCRIPTION"] = nw["text"]
        elif ex[-1].islower() and nw["text"] and nw["text"][0].islower():
            best["DESCRIPTION"] = ex + nw["text"]      # word-wrap, no space
        else:
            best["DESCRIPTION"] = ex + " " + nw["text"]


def _filter_rows(transactions):
    skip = ("opening balance", "closing balance", "brought forward",
            "carried forward", "opening bal", "closing bal")
    clean = []
    for t in transactions:
        t["DESCRIPTION"] = t["DESCRIPTION"].strip()
        dl = t["DESCRIPTION"].lower()
        if any(k in dl for k in skip):
            continue
        if not t["WITHDRAWAL_AMOUNT"] and not t["DEPOSIT_AMOUNT"] and not t["CLOSING_BALANCE"]:
            continue
        clean.append(t)
    return clean


def _extract_page_transactions(
    page,
    canonical_ranges: dict[str, tuple[float, float]],
    table_start_y: float,
    table_end_y: float,
    multi_word_date: bool = False,
) -> list[dict]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    date_range = canonical_ranges.get("DATE")
    date_col_x0 = date_range[0] if date_range else None

    transactions: list[dict] = []
    narration_buffer: list[dict] = []
    current_txn: dict | None = None

    if multi_word_date and date_range:
        lines = _group_into_lines([
            w for w in words
            if table_start_y <= w["top"] <= table_end_y
        ])
        for ln in lines:
            date_str = _kotak_date_from_line(ln["words"], date_range[0], date_range[1])
            if date_str:
                lc = (ln["top"] + ln["bottom"]) / 2
                current_txn = {
                    "center": lc, "DATE": date_str,
                    "DESCRIPTION": "", "TRANSACTION_NO": "",
                    "WITHDRAWAL_AMOUNT": "", "DEPOSIT_AMOUNT": "",
                    "CLOSING_BALANCE": "",
                }
                transactions.append(current_txn)
            for w in ln["words"]:
                text = w["text"].strip()
                if not text or _DASH_RE.match(text):
                    continue
                if date_col_x0 and w["x0"] < date_col_x0 and _SERIAL_RE.match(text):
                    continue
                can = _classify(w["x0"], canonical_ranges)
                if can is None or can == "DATE":
                    continue
                _assign_word(can, text, (w["top"]+w["bottom"])/2,
                             current_txn, transactions, narration_buffer)
    else:
        for word in words:
            text = word["text"].strip()
            x0, top, bottom = word["x0"], word["top"], word["bottom"]
            if not text or top < table_start_y or top > table_end_y:
                continue
            if _DASH_RE.match(text):
                continue
            if date_col_x0 and x0 < date_col_x0 and _SERIAL_RE.match(text):
                continue

            can = _classify(x0, canonical_ranges)
            if can is None:
                continue
            wc = (top + bottom) / 2

            if can == "DATE":
                if _is_date(text):
                    current_txn = {
                        "center": wc, "DATE": text,
                        "DESCRIPTION": "", "TRANSACTION_NO": "",
                        "WITHDRAWAL_AMOUNT": "", "DEPOSIT_AMOUNT": "",
                        "CLOSING_BALANCE": "",
                    }
                    transactions.append(current_txn)
                else:
                    narration_buffer.append({"text": text, "center": wc})
                continue

            _assign_word(can, text, wc, current_txn, transactions, narration_buffer)

    _flush_narration(narration_buffer, transactions)
    return _filter_rows(transactions)


# ---------------------------------------------------------------------------
# Bank detection
# ---------------------------------------------------------------------------

_BANK_SIGNATURES: dict[str, list[str]] = {
    "Bank of Baroda": ["bank of baroda", "bankofbaroda"],
    "Indian Bank":    ["indian bank", "idib"],
    "HDFC Bank":      ["hdfc bank", "hdfcbank"],
    "ICICI Bank":     ["icici bank", "icicibank"],
    "SBI":            ["state bank of india", "sbi.", "sbin0", "sbin"],
    "Axis Bank":      ["axis bank", "axisbank", "utib0"],
    "Kotak Bank":     ["kotak mahindra", "kotak bank", "kkbk"],
    "Punjab National":["punjab national bank", "pnb"],
    "Canara Bank":    ["canara bank"],
    "Union Bank":     ["union bank of india"],
    "Yes Bank":       ["yes bank"],
}

_MULTI_WORD_DATE_BANKS = {"Kotak Bank"}
_SBI_BANKS = {"SBI"}


def _detect_bank(text: str) -> str:
    tl = text.lower()
    for bank, sigs in _BANK_SIGNATURES.items():
        if any(s in tl for s in sigs):
            return bank
    return "Unknown Bank"


# ---------------------------------------------------------------------------
# Metadata
# ---------------------------------------------------------------------------

_ACNO_RE   = re.compile(r"(?:account\s*(?:no|number|no\.)\s*[:\.\-]?\s*)([0-9][\d\s]{5,20})", re.I)
_IFSC_RE   = re.compile(r"\b([A-Z]{4}0[A-Z0-9]{6})\b")
_PERIOD_RE = re.compile(
    r"(?:from|period|from\s*:)\s*"
    r"(\d{1,2}[-/\s]\w{2,9}[-/\s]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})"
    r"(?:\s*(?:to|-)\s*"
    r"(\d{1,2}[-/\s]\w{2,9}[-/\s]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}))?",
    re.I,
)


def _extract_metadata(pdf_path: str) -> dict[str, Any]:
    meta: dict[str, Any] = {
        "bank": "Unknown Bank", "account_no": "", "account_holder": "",
        "ifsc": "", "period_from": "", "period_to": "", "currency": "INR",
    }
    with pdfplumber.open(pdf_path) as pdf:
        pages = min(2, len(pdf.pages))
        full_text = "\n".join(pdf.pages[i].extract_text() or "" for i in range(pages))

    meta["bank"] = _detect_bank(full_text)
    if m := _ACNO_RE.search(full_text):
        meta["account_no"] = re.sub(r"\s+", "", m.group(1)).strip()
    if m := _IFSC_RE.search(full_text):
        meta["ifsc"] = m.group(1)
    if m := _PERIOD_RE.search(full_text):
        meta["period_from"] = (m.group(1) or "").strip()
        meta["period_to"]   = (m.group(2) or "").strip()
    for line in full_text.splitlines():
        line = line.strip()
        if not line or re.search(r"\d", line):
            continue
        ws = line.split()
        if 2 <= len(ws) <= 6 and all(w[0].isupper() for w in ws if w.isalpha()):
            if not any(kw in line.lower() for kw in
                       ("statement", "account", "branch", "savings", "bank")):
                meta["account_holder"] = line
                break
    return meta


# ---------------------------------------------------------------------------
# SBI: detect if a page uses the SBI fixed-column headerless layout
# ---------------------------------------------------------------------------

def _is_sbi_layout(page, table_start_y: float = 0.0) -> bool:
    """
    Check whether this page follows the SBI fixed-column layout by looking
    for the characteristic 'Balance' word at far right + dates at x≈27.
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    has_balance_right = any(
        w["x0"] > 500 and _norm(w["text"]) == "balance"
        for w in words
    )
    has_date_left = any(
        w["x0"] < 50 and _is_date(w["text"])
        for w in words if w["top"] > table_start_y
    )
    return has_balance_right and has_date_left


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def extract_statement(
    pdf_path: str | Path,
    password: str | None = None,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Extract all transactions from any Indian bank statement PDF.

    Returns
    -------
    df   : pd.DataFrame  — DATE, DESCRIPTION, TRANSACTION_NO,
                           WITHDRAWAL_AMOUNT, DEPOSIT_AMOUNT, CLOSING_BALANCE
    meta : dict          — bank, account_no, account_holder, ifsc,
                           period_from, period_to, currency,
                           total_pages, transactions_found
    """
    pdf_path = str(Path(pdf_path))
    if not Path(pdf_path).exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    logger.info("[UBE] Extracting: %s", pdf_path)
    meta = _extract_metadata(pdf_path)
    bank = meta["bank"]

    multi_word_date = bank in _MULTI_WORD_DATE_BANKS
    is_sbi_bank     = bank in _SBI_BANKS

    all_records: list[dict] = []
    canonical_ranges: dict[str, tuple[float, float]] = {}
    total_pages = 0

    open_kwargs: dict[str, Any] = {}
    if password:
        open_kwargs["password"] = password

    with pdfplumber.open(pdf_path, **open_kwargs) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            if not words:
                logger.debug("[UBE] Page %d: empty, skipping", page_num)
                continue

            # ---- SBI special path ----------------------------------------
            if is_sbi_bank:
                # Determine where the table data starts on this page
                # by finding the "Balance" header word
                table_start_y = 0.0
                for w in words:
                    if _norm(w["text"]) == "balance" and w["x0"] > 480:
                        table_start_y = w["bottom"] + 1.0
                        break
                if _is_sbi_layout(page, table_start_y):
                    logger.debug("[UBE] Page %d: SBI fixed-column layout", page_num)
                    records = _extract_sbi_transactions(page, table_start_y)
                    logger.debug("[UBE] Page %d: %d transactions", page_num, len(records))
                    all_records.extend(records)
                    continue
                # If not matching SBI layout, fall through to generic path

            # ---- Generic path --------------------------------------------
            lines = _group_into_lines(words)
            header_info = _detect_header(lines)

            if header_info is None:
                if not canonical_ranges:
                    logger.debug("[UBE] Page %d: no header, skipping", page_num)
                    continue
                logger.debug("[UBE] Page %d: reusing column ranges", page_num)
                table_start_y = 0.0
            else:
                new_ranges = _build_column_ranges(
                    header_info["words"],
                    header_info["canonical_found"],
                )
                if new_ranges:
                    canonical_ranges = new_ranges
                    logger.debug("[UBE] Page %d ranges: %s", page_num, canonical_ranges)
                elif not canonical_ranges:
                    logger.warning("[UBE] Page %d: empty ranges, skipping", page_num)
                    continue
                table_start_y = header_info["bottom"] + 2.0

            if not canonical_ranges:
                continue

            records = _extract_page_transactions(
                page, canonical_ranges,
                table_start_y=table_start_y,
                table_end_y=9999.0,
                multi_word_date=multi_word_date,
            )
            logger.debug("[UBE] Page %d: %d transactions", page_num, len(records))
            all_records.extend(records)

    if not all_records:
        logger.warning("[UBE] No transactions found in '%s'", pdf_path)
        meta.update({"total_pages": total_pages, "transactions_found": 0})
        return pd.DataFrame(columns=CANONICAL_COLUMNS), meta
    print("Number of records:", len(all_records))

    # Build a sanitized list of records and construct the DataFrame once.
    sanitized_records = []
    for rec in all_records:
        try:
            sanitized_records.append({k: _sanitize_value(v) for k, v in rec.items()})
        except Exception:
            sanitized_records.append({k: str(v) for k, v in rec.items()})

    df = pd.DataFrame.from_records(sanitized_records)
    df = df.drop(columns=["center"], errors="ignore")
    for col in CANONICAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[CANONICAL_COLUMNS].reset_index(drop=True)
    for col in ("WITHDRAWAL_AMOUNT", "DEPOSIT_AMOUNT", "CLOSING_BALANCE"):
        df[col] = pd.to_numeric(df[col].replace("", None), errors="coerce")

    meta.update({"total_pages": total_pages, "transactions_found": len(df)})
    logger.info("[UBE] Done. %d transactions from %d pages.", len(df), total_pages)
    return df, meta


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys, json

    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    if len(sys.argv) < 2:
        print("Usage: python universal_bank_extractor.py <pdf_path> [password]")
        sys.exit(1)

    pdf_file = sys.argv[1]
    pwd      = sys.argv[2] if len(sys.argv) > 2 else None

    df, meta = extract_statement(pdf_file, password=pwd)

    print("\n=== METADATA ===")
    print(json.dumps(meta, indent=2, default=str))
    print(f"\n=== TRANSACTIONS ({len(df)} rows) ===")
    pd.set_option("display.max_colwidth", 55)
    pd.set_option("display.width", 220)
    print(df.to_string(index=True))

    out = Path(pdf_file).stem + "_transactions.csv"
    df.to_csv(out, index=False)
    print(f"\nSaved → {out}")
