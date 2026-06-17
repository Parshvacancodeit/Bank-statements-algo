from __future__ import annotations
import logging

logger = logging.getLogger(__name__)



def _pdf_to_document_axis(pdf_path, column_mapping=None):
    from .axis_bank_pdf_processor import pdf_to_document_axis
    return pdf_to_document_axis(pdf_path, column_mapping=column_mapping)


def _pdf_to_document(pdf_path, column_mapping=None):
    from .general_structured_pdf_processor import pdf_to_document
    return pdf_to_document(pdf_path, column_mapping=column_mapping)


def _is_axis_bank_pdf(pdf_path, column_mapping=None):
    from .axis_bank_pdf_processor import is_axis_bank_pdf
    return is_axis_bank_pdf(pdf_path, column_mapping=column_mapping)


def _is_structured_pdf(pdf_path, column_mapping=None):
    from .general_structured_pdf_processor import is_structured_pdf
    return is_structured_pdf(pdf_path, column_mapping=column_mapping)


def _pdf_to_document_bob(pdf_path, column_mapping=None):
    from .bob_pdf_processor import pdf_to_document_bob
    return pdf_to_document_bob(pdf_path, column_mapping=column_mapping)


def _is_bob_pdf(pdf_path, column_mapping=None):
    from .bob_pdf_processor import is_bob_pdf
    return is_bob_pdf(pdf_path, column_mapping=column_mapping)


def _pdf_to_document_icici(pdf_path, column_mapping=None):
    from .icici_pdf_processor import pdf_to_document_icici
    return pdf_to_document_icici(pdf_path, column_mapping=column_mapping)


def _is_icici_pdf(pdf_path, column_mapping=None):
    from .icici_pdf_processor import is_icici_pdf
    return is_icici_pdf(pdf_path, column_mapping=column_mapping)


def _pdf_to_document_sbi(pdf_path, column_mapping=None):
    from .sbi_pdf_processor import pdf_to_document_sbi
    return pdf_to_document_sbi(pdf_path, column_mapping=column_mapping)


def _is_sbi_pdf(pdf_path, column_mapping=None):
    from .sbi_pdf_processor import is_sbi_pdf
    return is_sbi_pdf(pdf_path, column_mapping=column_mapping)


def _pdf_to_document_kotak(pdf_path, column_mapping=None):
    from .kotak_pdf_processor import pdf_to_document_kotak
    return pdf_to_document_kotak(pdf_path, column_mapping=column_mapping)


def _is_kotak_pdf(pdf_path, column_mapping=None):
    from .kotak_pdf_processor import is_kotak_pdf
    return is_kotak_pdf(pdf_path, column_mapping=column_mapping)


# ---------------------------------------------------------------------------
# Master index – add / change a bank's processor in one place
# ---------------------------------------------------------------------------

# Maps a canonical bank key → extraction function
EXTRACTION_INDEX: dict[str, callable] = {
    "AXIS":    _pdf_to_document_axis,
    "HDFC":    _pdf_to_document,
    "KALUPUR": _pdf_to_document,
    "BOB":     _pdf_to_document_bob,
    "ICICI":   _pdf_to_document_icici,
    "SBI":     _pdf_to_document_sbi,
    "KOTAK":   _pdf_to_document_kotak,
}

# Maps a canonical bank key → structure-check function
STRUCTURE_INDEX: dict[str, callable] = {
    "AXIS":    _is_axis_bank_pdf,
    "HDFC":    _is_structured_pdf,
    "KALUPUR": _is_structured_pdf,
    "BOB":     _is_bob_pdf,
    "ICICI":   _is_icici_pdf,
    "SBI":     _is_sbi_pdf,
    "KOTAK":   _is_kotak_pdf,
}

# Default fallback used when the bank is not found in either index
_DEFAULT_EXTRACTION_FN = _pdf_to_document
_DEFAULT_STRUCTURE_FN  = _is_structured_pdf


# ---------------------------------------------------------------------------
# Public routing functions
# ---------------------------------------------------------------------------

def is_structured_pdf_for_bank(
    pdf_path: str,
    bank_name: str | None,
    column_mapping: dict | None = None,
) -> bool:
    """
    Determine whether *pdf_path* contains a structured bank statement that can
    be extracted without OCR, using the processor appropriate for *bank_name*.

    Args:
        pdf_path:       Path to the PDF file on disk.
        bank_name:      Bank name string as stored in the DB (case-insensitive).
        column_mapping: ``{canonical_key: [alias, ...]}`` from edocsmart_bank_mapping.

    Returns:
        True  – PDF is structured and can be extracted digitally.
        False – PDF is image-based; fall back to the OCR/AI flow.
    """
    bank_key = bank_name
    logger.info(
        "[BANK_INDEX] is_structured_pdf_for_bank: bank='%s' key='%s' path='%s'",
        bank_name, bank_key, pdf_path,
    )

    bank_data_structure_function = STRUCTURE_INDEX.get(bank_key, _DEFAULT_STRUCTURE_FN)
    result = bank_data_structure_function(pdf_path, column_mapping=column_mapping)
    
    logger.info("[BANK_INDEX] structure-check result: %s", result)
    return result


def pdf_to_document_for_bank(
    pdf_path: str,
    bank_name: str | None,
    column_mapping: dict | None = None,
) -> tuple:
    """
    Returns:
        (df, original_headers) where:
          - df               : pandas DataFrame with canonical column names
          - original_headers : dict  canonical_key → original PDF header text
        Both are empty when no structured data is found.

    """
    bank_key = bank_name
    logger.info("[BANK_INDEX] pdf_to_document_for_bank: bank='%s' key='%s' path='%s'",bank_name, bank_key, pdf_path,)
    bank_data_extraction_function = EXTRACTION_INDEX.get(bank_key, _DEFAULT_EXTRACTION_FN)
    return bank_data_extraction_function(pdf_path, column_mapping=column_mapping)
