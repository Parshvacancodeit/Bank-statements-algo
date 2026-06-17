# Bank Statement Extractor - User Documentation

This guide explains how `extract_all_banks.py` works in simple words and how to use it.

---

## 1. How to run it

At the very top of `extract_all_banks.py`, there is a configuration section:
```python
# ==============================================================================
# CONFIGURATION: Set this path to a directory containing PDF files,
# or directly to a single PDF bank statement file.
INPUT_PATH = "/Users/parshvapatel/Downloads/bank_statement_pdfs"
# ==============================================================================
```

- **To extract all PDFs in a folder**: Set `INPUT_PATH` to the folder directory.
- **To extract a single PDF**: Set `INPUT_PATH` directly to the path of that specific PDF file (e.g. `"/Users/parshvapatel/Downloads/bank_statement_pdfs/axis_bank.pdf"`).

Run the script using Python:
```bash
python extract_all_banks.py
```

It will process the files and generate a beautifully formatted, colored Excel file named `bank_statements_all.xlsx` in the same directory!

---

## 2. How the code works (in simple words)

The code consists of **4 main steps**:

### Step A: Detection (Which Bank is this?)
When you give a PDF to the script, it reads the first page's text and searches for specific "signatures" (keywords like "Axis Bank", "State Bank of India", "HDFC", etc.). Based on which signatures appear, it selects the correct bank processor.

### Step B: Extraction Processors
Once the bank is identified, the script runs the specific processor for that bank:

1. **Table-Based Extraction (Axis Bank, Kotak, SBI Layout 1 & 2)**:
   - These banks have structured grid tables in the PDF.
   - The script uses `pdfplumber` to extract tables, maps the header columns (e.g., matching "debit" or "withdrawal" to the withdrawal column), and processes each row sequentially.
   
2. **Coordinate-Based Extraction (Bank of Baroda, HDFC, ICICI, Indian Bank)**:
   - These statements either do not have borders or have complex layouts that confuse standard table extractors.
   - For these, the script extracts every single word along with its exact spatial **(x, y) coordinates** on the page.
   - It groups words vertically into lines.
   - It partitions columns based on predefined horizontal coordinate ranges (e.g., if a word starts between x=30 and x=76, it belongs to the `DATE` column).
   - In banks like HDFC, ICICI, Bank of Baroda, and Indian Bank, transaction descriptions often span multiple lines. The script uses a **proximity-alignment algorithm** page-by-page. If text is printed on a new line but aligns vertically with a preceding date, it gets automatically merged into that transaction's description.

### Step C: Sanitization & Clean-up
The script cleans all the columns before exporting:
- Strips commas, currency markers (like `Cr` and `Dr`), and hyphens to convert amount text into clean numbers.
- Ignores summary/header rows like "Opening Balance", "Brought Forward", or "Carried Forward".
- Standardizes all sheets into a clean 7-column layout:
  `DATE`, `VALUE_DATE`, `DESCRIPTION`, `TRANSACTION_NO`, `WITHDRAWAL_AMOUNT`, `DEPOSIT_AMOUNT`, `CLOSING_BALANCE`.

### Step D: Styled Excel Output
Finally, it writes the clean tables into an Excel file using the `openpyxl` library:
- Assigns a professional deep navy header theme with bold white text.
- Formats withdrawal, deposit, and balance numbers as currency format (`#,##0.00`) and aligns them to the right.
- Left-aligns date and description text.
- Automatically adjusts column widths so that text fits neatly and is not cut off.
- Freezes the top header row so it stays visible as you scroll down.
