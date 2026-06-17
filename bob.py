"""
Bank Statement PDF → Excel Converter
Supports: Axis Bank, Bank of Baroda (BOB), HDFC Bank
"""

import re
import sys
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
try:
    from universal_bank_extractor import extract_statement as ube_extract_statement
    _HAS_UNIVERSAL = True
except Exception:
    _HAS_UNIVERSAL = False


# ─────────────────────────────────────────────────────────────────────────────
# Shared helpers
# ─────────────────────────────────────────────────────────────────────────────

DATE_RE_DMY  = re.compile(r'^\d{2}[-/]\d{2}[-/]\d{4}$')   # 02-01-2024 / 02/01/2024
DATE_RE_DMYY = re.compile(r'^\d{2}/\d{2}/\d{2}$')          # 01/05/26  (HDFC)

def is_date(text: str) -> bool:
    t = (text or '').strip()
    return bool(DATE_RE_DMY.match(t) or DATE_RE_DMYY.match(t))

def clean_amount(text: str) -> str:
    """Remove commas, trailing Cr/Dr tags, keep number."""
    t = re.sub(r'[,\s]', '', (text or '').strip())
    t = re.sub(r'(Cr|Dr)$', '', t, flags=re.IGNORECASE)
    return t if re.match(r'^-?\d+\.?\d*$', t) else ''

def normalize_text(text: str) -> str:
    return re.sub(r'\s+', ' ', (text or '').replace('\n', ' ')).strip()


# ─────────────────────────────────────────────────────────────────────────────
# Bank detection
# ─────────────────────────────────────────────────────────────────────────────

def detect_bank(pdf_path: str) -> str:
    """Return 'axis', 'bob', or 'hdfc' based on PDF content."""
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages[:2]:
            text += (page.extract_text() or '').lower()

    # Check most-specific identifiers first to avoid false matches
    # HDFC: their own IFSC code or brand name (no spaces due to PDF extraction)
    if 'hdfcbank' in text or 'hdfc0001' in text or 'hdfc bank' in text:
        return 'hdfc'
    # BOB: bank of baroda or BARB IFSC prefix
    if 'bank of baroda' in text or 'bankofbaroda' in text or 'barb0' in text:
        return 'bob'
    # Axis: own IFSC code UTIB0001 in header area, or 'axis bank'
    if 'axis bank' in text or 'axis' in text or 'utib0001' in text:
        return 'axis'
    # Kotak, ICICI, SBI, Indian Bank
    if 'kotak' in text or 'kotak mahindra' in text:
        return 'kotak'
    if 'icici' in text or 'icici bank' in text:
        return 'icici'
    if 'state bank of india' in text or 'sbi' in text:
        return 'sbi'
    if 'indian bank' in text or 'idib' in text:
        return 'indian'
    return 'unknown'


# ─────────────────────────────────────────────────────────────────────────────
# AXIS BANK extractor
# ─────────────────────────────────────────────────────────────────────────────
# Layout: clean ruled table, 7 columns
# Header: Tran Date | Chq No | Particulars | Debit | Credit | Balance | Init.Br
# pdfplumber.extract_tables() works reliably page-by-page.
# Page 2+ has no repeated header row — data starts directly.

AXIS_CANONICAL = ['Date', 'Chq No', 'Description', 'Debit', 'Credit', 'Balance', 'Branch']
AXIS_SKIP_DESC = {'opening balance', 'closing balance', 'transaction total'}


def _axis_parse_table_rows(rows, has_header: bool) -> list[dict]:
    records = []
    start = 1 if has_header else 0
    for row in rows[start:]:
        if not row or len(row) < 6:
            continue
        date = normalize_text(row[0] or '')
        desc = normalize_text(row[2] or '')
        if not is_date(date):
            continue
        if desc.lower() in AXIS_SKIP_DESC:
            continue
        records.append({
            'Date':        date,
            'Chq No':      normalize_text(row[1] or ''),
            'Description': desc,
            'Debit':       clean_amount(row[3] or ''),
            'Credit':      clean_amount(row[4] or ''),
            'Balance':     clean_amount(row[5] or ''),
            'Branch':      normalize_text(row[6] or '') if len(row) > 6 else '',
        })
    return records


def extract_axis(pdf_path: str) -> list[dict]:
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                continue
            for table in tables:
                if not table:
                    continue
                # Detect if first row is header
                first = [normalize_text(c or '').lower() for c in table[0]]
                has_header = any(h in first for h in ['tran date', 'date', 'particulars'])
                records.extend(_axis_parse_table_rows(table, has_header))
    return records


# ─────────────────────────────────────────────────────────────────────────────
# BANK OF BARODA extractor
# ─────────────────────────────────────────────────────────────────────────────
# Layout: ruled table BUT each row is a separate mini-table (1 row each)
# because horizontal rules exist per row.  We collect all tables per page,
# identify the header table to get column mapping, then iterate single-row tables.
# Header: DATE | NARRATION | CHQ.NO. | WITHDRAWAL (DR) | DEPOSIT (CR) | BALANCE

BOB_CANONICAL = ['Date', 'Description', 'Chq No', 'Withdrawal', 'Deposit', 'Balance']
BOB_SKIP_NARR = {'opening balance', 'closing balance'}
BOB_HEADER_KEYS = {'date', 'narration', 'withdrawal', 'deposit', 'balance', 'chq.no.', 'chq no'}


def _bob_is_header_table(table) -> bool:
    if not table or len(table) < 2:
        return False
    for row in table[:3]:
        row_text = ' '.join(normalize_text(c or '').lower() for c in row)
        matches = sum(1 for k in BOB_HEADER_KEYS if k in row_text)
        if matches >= 3:
            return True
    return False


def _bob_parse_row(row) -> dict | None:
    """Parse a single BOB table row [date, narration, chqno, withdrawal, deposit, balance]."""
    if not row or len(row) < 5:
        return None
    date = normalize_text(row[0] or '')
    if not is_date(date):
        return None
    narr = normalize_text(row[1] or '')
    if narr.lower() in BOB_SKIP_NARR:
        return None
    bal_raw = normalize_text(row[5] if len(row) > 5 else (row[4] or ''))
    return {
        'Date':        date,
        'Description': narr,
        'Chq No':      normalize_text(row[2] or ''),
        'Withdrawal':  clean_amount(row[3] or ''),
        'Deposit':     clean_amount(row[4] or ''),
        'Balance':     clean_amount(bal_raw),
    }


def extract_bob(pdf_path: str) -> list[dict]:
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                # Skip header/summary tables
                if _bob_is_header_table(table):
                    # Still try to extract data rows inside (row index 2+)
                    for row in table[2:]:
                        parsed = _bob_parse_row(row)
                        if parsed:
                            records.append(parsed)
                    continue
                # Single-row transaction tables
                for row in table:
                    parsed = _bob_parse_row(row)
                    if parsed:
                        records.append(parsed)
    return records


# ─────────────────────────────────────────────────────────────────────────────
# HDFC BANK extractor
# ─────────────────────────────────────────────────────────────────────────────
# Layout: outer border only — NO horizontal rules between rows.
# pdfplumber merges entire page into 1-2 table rows with all values newline-joined.
# Strategy: word-level extraction, group by date words in the Date column (x0 < 65).
# Columns (from header words):
#   Date:           x0~28–68
#   Narration:      x0~68–255
#   Chq/Ref.No.:    x0~255–357
#   Value Dt:       x0~357–420
#   Withdrawal Amt: x0~420–490
#   Deposit Amt:    x0~490–562
#   Closing Balance:x0~562–631

HDFC_CANONICAL = ['Date', 'Description', 'Ref No', 'Value Date', 'Withdrawal', 'Deposit', 'Balance']

# Column x-boundaries derived from header positions (constant across pages)
HDFC_COLS = [
    ('date',        28,   68),
    ('narration',   68,  258),
    ('ref',        258,  360),
    ('value_dt',   360,  422),
    ('withdrawal', 422,  492),
    ('deposit',    492,  564),
    ('balance',    564,  640),
]

HDFC_TABLE_START_Y = 225   # below the "From / To" header line


def _hdfc_classify(x0: float) -> str | None:
    for col_name, x_start, x_end in HDFC_COLS:
        if x_start <= x0 < x_end:
            return col_name
    return None


def _hdfc_extract_page(page) -> list[dict]:
    words = page.extract_words()
    # Only words inside the table area
    table_words = [w for w in words if w['top'] > HDFC_TABLE_START_Y]

    # Anchor rows: words in the 'date' column that look like dates
    date_anchors = [
        w for w in table_words
        if _hdfc_classify(w['x0']) == 'date' and is_date(w['text'])
    ]
    if not date_anchors:
        return []

    # Build transaction buckets keyed by their top-y
    transactions: list[dict] = []
    for anchor in date_anchors:
        transactions.append({
            '_top':      anchor['top'],
            '_bottom':   anchor['top'] + 20,   # estimated row height; refined below
            'date':      anchor['text'],
            'narration': [],
            'ref':       [],
            'value_dt':  [],
            'withdrawal':[],
            'deposit':   [],
            'balance':   [],
        })

    # Sort anchors by y
    transactions.sort(key=lambda t: t['_top'])

    # Assign each non-date word to the nearest transaction by y-distance
    for w in table_words:
        col = _hdfc_classify(w['x0'])
        if col is None or col == 'date':
            continue
        word_cy = (w['top'] + w['bottom']) / 2

        best, best_dist = None, float('inf')
        for txn in transactions:
            dist = abs(txn['_top'] - word_cy)
            if dist < best_dist:
                best_dist = dist
                best = txn

        # Only attach if reasonably close (within ~60 pt = ~2 narration wrap lines)
        if best is not None and best_dist < 60:
            best[col].append(w['text'])

    records = []
    for txn in transactions:
        narr = ' '.join(txn['narration']).strip()
        # Skip summary / header bleed
        if narr.lower() in {'opening balance', 'closing balance', 'statement summary :-',
                             'opening balance dr count cr count debits credits closing bal'}:
            continue

        records.append({
            'Date':        txn['date'],
            'Description': narr,
            'Ref No':      ' '.join(txn['ref']).strip(),
            'Value Date':  ' '.join(txn['value_dt']).strip(),
            'Withdrawal':  clean_amount(' '.join(txn['withdrawal'])),
            'Deposit':     clean_amount(' '.join(txn['deposit'])),
            'Balance':     clean_amount(' '.join(txn['balance'])),
        })
    return records


def extract_hdfc(pdf_path: str) -> list[dict]:
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            records.extend(_hdfc_extract_page(page))
    return records


# ─────────────────────────────────────────────────────────────────────────────
# Excel writer
# ─────────────────────────────────────────────────────────────────────────────

HEADER_FILL   = PatternFill('solid', start_color='1F4E79')
HEADER_FONT   = Font(bold=True, color='FFFFFF', name='Arial', size=10)
DATA_FONT     = Font(name='Arial', size=9)
ALT_FILL      = PatternFill('solid', start_color='D9E2F3')
DEBIT_FONT    = Font(name='Arial', size=9, color='C00000')
CREDIT_FONT   = Font(name='Arial', size=9, color='375623')
THIN          = Side(style='thin', color='BFBFBF')
THIN_BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

META_LABEL_FONT = Font(bold=True, name='Arial', size=9)
META_VALUE_FONT = Font(name='Arial', size=9)

BANK_COLORS = {
    'axis':  '97002E',   # Axis maroon
    'bob':   'FF6600',   # BOB orange
    'hdfc':  '004C8F',   # HDFC blue
    'unknown': '404040',
}


def _add_meta_sheet(wb, bank: str, pdf_path: str, record_count: int):
    ws = wb.create_sheet('Info', 0)
    ws.sheet_view.showGridLines = False

    color = BANK_COLORS.get(bank, '404040')
    title_fill = PatternFill('solid', start_color=color)
    title_font = Font(bold=True, name='Arial', size=14, color='FFFFFF')

    bank_names = {'axis': 'Axis Bank', 'bob': 'Bank of Baroda', 'hdfc': 'HDFC Bank', 'unknown': 'Unknown Bank'}

    ws.merge_cells('A1:D1')
    ws['A1'] = f"{bank_names.get(bank, bank.upper())} — Statement Summary"
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    meta = [
        ('Source File',       Path(pdf_path).name),
        ('Detected Bank',     bank_names.get(bank, bank)),
        ('Total Transactions', record_count),
    ]
    for i, (label, value) in enumerate(meta, start=3):
        ws.cell(i, 1, label).font = META_LABEL_FONT
        ws.cell(i, 2, str(value)).font = META_VALUE_FONT
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 36


def write_excel(records: list[dict], headers: list[str], bank: str, pdf_path: str, out_path: str):
    wb = openpyxl.Workbook()

    # Create Transactions sheet first (as active), then Info sheet
    ws = wb.active
    ws.title = 'Transactions'

    _add_meta_sheet(wb, bank, pdf_path, len(records))
    ws.sheet_view.showGridLines = False

    color = BANK_COLORS.get(bank, '404040')
    header_fill = PatternFill('solid', start_color=color)

    # Write header
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(1, col_idx, h)
        cell.font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 20

    # Amount columns (0-indexed positions that contain monetary values)
    amount_headers = {'Debit', 'Credit', 'Balance', 'Withdrawal', 'Deposit'}
    amount_cols = {i + 1 for i, h in enumerate(headers) if h in amount_headers}

    # Write data
    for row_idx, rec in enumerate(records, 2):
        alt = (row_idx % 2 == 0)
        for col_idx, h in enumerate(headers, 1):
            val = rec.get(h, '')
            cell = ws.cell(row_idx, col_idx, val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical='top', wrap_text=(h == 'Description'))

            if alt:
                cell.fill = ALT_FILL

            # Amount formatting
            if col_idx in amount_cols and val:
                cell.alignment = Alignment(horizontal='right', vertical='top')
                is_debit = h in ('Debit', 'Withdrawal')
                is_credit = h in ('Credit', 'Deposit')
                if is_debit:
                    cell.font = DEBIT_FONT
                elif is_credit:
                    cell.font = CREDIT_FONT

    # Column widths
    width_map = {
        'Date': 14, 'Value Date': 13, 'Chq No': 10, 'Ref No': 22,
        'Description': 48, 'Narration': 48,
        'Debit': 14, 'Credit': 14, 'Balance': 16,
        'Withdrawal': 15, 'Deposit': 14,
        'Branch': 9,
    }
    for col_idx, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width_map.get(h, 14)

    # Freeze header row
    ws.freeze_panes = 'A2'

    # Summary row
    last_row = len(records) + 2
    ws.cell(last_row, 1, 'TOTAL').font = Font(bold=True, name='Arial', size=9)
    ws.cell(last_row, 1).fill = PatternFill('solid', start_color='F2F2F2')

    debit_h  = next((h for h in ['Debit', 'Withdrawal'] if h in headers), None)
    credit_h = next((h for h in ['Credit', 'Deposit'] if h in headers), None)
    for col_idx, h in enumerate(headers, 1):
        if h in (debit_h, credit_h):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(last_row, col_idx)
            cell.value = f'=SUM({col_letter}2:{col_letter}{last_row - 1})'
            cell.font = Font(bold=True, name='Arial', size=9,
                             color='C00000' if h == debit_h else '375623')
            cell.alignment = Alignment(horizontal='right')
            cell.fill = PatternFill('solid', start_color='F2F2F2')
            cell.border = THIN_BORDER

    wb.save(out_path)
    print(f"  ✓ Saved: {out_path}  ({len(records)} transactions)")


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

BANK_HEADERS = {
    'axis':    AXIS_CANONICAL,
    'bob':     BOB_CANONICAL,
    'hdfc':    HDFC_CANONICAL,
    'unknown': ['Date', 'Description', 'Chq No', 'Debit', 'Credit', 'Balance'],
    'kotak':   ['Date', 'Description', 'Chq No', 'Debit', 'Credit', 'Balance'],
    'icici':   ['Date', 'Description', 'Chq No', 'Debit', 'Credit', 'Balance'],
    'sbi':     ['Date', 'Description', 'Chq No', 'Debit', 'Credit', 'Balance'],
    'indian':  ['Date', 'Description', 'Chq No', 'Debit', 'Credit', 'Balance'],
}

EXTRACTORS = {
    'axis':    extract_axis,
    'bob':     extract_bob,
    'hdfc':    extract_hdfc,
}


def process_pdf(pdf_path: str, out_dir: str = '.') -> str:
    pdf_path = str(pdf_path)
    print(f"\nProcessing: {pdf_path}")

    bank = detect_bank(pdf_path)
    print(f"  Detected bank: {bank.upper()}")

    extractor = EXTRACTORS.get(bank, extract_axis)
    records   = extractor(pdf_path)
    print(f"  Extracted {len(records)} transactions")

    # If no records found with bank-specific extractor, try the universal extractor
    if not records and _HAS_UNIVERSAL:
        try:
            print("  → Trying universal extractor fallback...")
            df, meta = ube_extract_statement(pdf_path)
            if df is not None and len(df) > 0:
                # Map universal columns to our expected header names
                mapped = []
                for _, row in df.iterrows():
                    mapped.append({
                        'Date': str(row.get('DATE', '')),
                        'Description': str(row.get('DESCRIPTION', '')),
                        'Chq No': str(row.get('TRANSACTION_NO', '')),
                        'Debit': str(row.get('WITHDRAWAL_AMOUNT', '')),
                        'Credit': str(row.get('DEPOSIT_AMOUNT', '')),
                        'Balance': str(row.get('CLOSING_BALANCE', '')),
                    })
                records = mapped
                bank = (meta.get('bank') or 'unknown').lower()
                print(f"  Universal extractor found {len(records)} transactions (bank: {bank})")
        except Exception as e:
            print(f"  ✗ Universal extractor failed: {e}")

    if not records:
        print("  ⚠ No transactions found — skipping Excel output.")
        return ''

    headers  = BANK_HEADERS.get(bank, BANK_HEADERS['unknown'])
    stem     = Path(pdf_path).stem
    out_path = str(Path(out_dir) / f"{stem}_transactions.xlsx")

    write_excel(records, headers, bank, pdf_path, out_path)
    return out_path


def main():
    # Default behaviour: process all PDF files in the same folder as this script
    script_dir = Path(__file__).parent
    if len(sys.argv) > 1:
        pdf_files = sys.argv[1:]
    else:
        pdf_files = [str(p) for p in script_dir.glob('*.pdf')]

    out_dir = str(script_dir)
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    outputs = []
    for pdf in pdf_files:
        if not Path(pdf).exists():
            print(f"  ✗ File not found: {pdf}")
            continue
        out = process_pdf(pdf, out_dir)
        if out:
            outputs.append(out)

    print(f"\nDone. {len(outputs)} file(s) written.")
    return outputs


if __name__ == '__main__':
    main()