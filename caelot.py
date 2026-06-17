try:
    import camelot
    _HAS_CAMELOT = True
except ModuleNotFoundError:
    _HAS_CAMELOT = False

import pandas as pd
import re

PDF_FILE = "/Users/parshvapatel/Downloads/bank_statement_pdfs/axis_bank.pdf"
OUTPUT_FILE = "bank_statement.xlsx"

# Keywords commonly found in transaction tables
HEADER_KEYWORDS = [
    "date",
    "value date",

    "description",
    "narration",
    "particulars",
    "remarks",
    "withdrawal",
    "deposit",
    "debit",
    "credit",
    "balance",
    "amount"
]


def is_transaction_table(df):
    """
    Check if table looks like a bank transaction table.
    """
    try:
        header_text = " ".join(
            str(x).lower().strip()
            for x in df.iloc[0].tolist()
        )

        matches = sum(
            1 for keyword in HEADER_KEYWORDS
            if keyword in header_text
        )

        return matches >= 2
    except:
        return False


def clean_dataframe(df):
    """
    Clean extracted table.
    """
    df = df.replace("\n", " ", regex=True)

    # First row becomes header
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

    # Remove empty rows
    df = df.dropna(how="all")

    return df


print("Extracting tables...")
if not _HAS_CAMELOT:
    try:
        from universal_bank_extractor import extract_statement
    except Exception:
        raise ModuleNotFoundError(
            "camelot is not installed. Install it with: `pip3 install camelot-py[cv]`\n"
            "On macOS also install Ghostscript (e.g. `brew install ghostscript`).\n"
            "Alternatively ensure `universal_bank_extractor.py` is present in the same folder."
        )

    print("`camelot` not available — falling back to universal_bank_extractor (pdfplumber)")
    df, meta = extract_statement(PDF_FILE)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} rows to {OUTPUT_FILE} (using universal_bank_extractor)")
    raise SystemExit(0)

# Try lattice first (works for bordered tables). If lattice fails, fall back
# to stream mode. If both fail or find no tables, fall back to the pdfplumber
# extractor in `universal_bank_extractor`.
tables = []
try:
    tables = camelot.read_pdf(
        PDF_FILE,
        pages="all",
        flavor="lattice"
    )
except Exception as e:
    print("camelot lattice mode failed:", e)

if len(tables) == 0:
    try:
        tables = camelot.read_pdf(
            PDF_FILE,
            pages="all",
            flavor="stream"
        )
    except Exception as e:
        print("camelot stream mode failed:", e)
        tables = []

if len(tables) == 0:
    print("camelot did not return any tables — falling back to universal_bank_extractor (pdfplumber)")
    try:
        from universal_bank_extractor import extract_statement
        df, meta = extract_statement(PDF_FILE)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"Saved {len(df)} rows to {OUTPUT_FILE} (using universal_bank_extractor)")
        raise SystemExit(0)
    except Exception as e:
        raise RuntimeError("Both camelot parsing and universal_bank_extractor fallback failed: " + str(e))

all_transactions = []

for i, table in enumerate(tables):

    df = table.df

    if is_transaction_table(df):
        print(f"Transaction table found: {i + 1}")

        cleaned_df = clean_dataframe(df)

        all_transactions.append(cleaned_df)

if not all_transactions:
    raise Exception(
        "No transaction tables detected. Try flavor='stream'."
    )

final_df = pd.concat(
    all_transactions,
    ignore_index=True
)

# Remove duplicate headers that appear on new pages
date_pattern = re.compile(
    r"date|value date",
    re.IGNORECASE
)

first_col = str(final_df.columns[0])

final_df = final_df[
    ~final_df.iloc[:, 0]
    .astype(str)
    .str.contains(date_pattern, na=False)
]

final_df.to_excel(
    OUTPUT_FILE,
    index=False
)

print(f"Saved {len(final_df)} rows to {OUTPUT_FILE}")