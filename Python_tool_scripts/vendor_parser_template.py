# ============================================================
# Vendor Parser Template
# ------------------------------------------------------------
# Copy this file, rename it:
#   <vendor_name>_parser.py
#
# Then fill in the extraction rules inside parse_invoice().
#
# REQUIRED RETURN FORMAT:
#   {
#       "vendor": "Vendor Name",
#       "invoice_number": "12345",
#       "jobname": "Some Job",
#       "date": "2025-01-01",
#       "total": "1234.56",
#       "page": 1
#   }
#
# If multi-invoice PDF, return a LIST of dictionaries.
# ============================================================

import fitz  # PyMuPDF
import re

# ------------------------------------------------------------
# Extract ALL text from a PDF — page by page
# ------------------------------------------------------------
def extract_pdf_text(pdf_path):
    """Return a list: [ (page_number, text), ... ]"""
    doc = fitz.open(pdf_path)
    pages = []

    for i, page in enumerate(doc):
        text = page.get_text("text")
        pages.append((i + 1, text))

    doc.close()
    return pages


# ------------------------------------------------------------
# MAIN PARSER FUNCTION
# ------------------------------------------------------------
def parse_invoice(pdf_path):
    """
    Extract fields from a vendor PDF.
    This is the function InvoiceSorter uses.

    Return:
        - dict (single invoice)
        - list of dicts (multi-invoice PDFs)
    """

    pages = extract_pdf_text(pdf_path)

    # --------------------------------------------------------
    # DEFAULT RESULT (empty shell — you will fill rules)
    # --------------------------------------------------------
    results = []

    for page_num, text in pages:

        # Build a dictionary for this page.
        # You will modify the extraction logic per vendor.
        inv = {
            "vendor": "",           # MUST fill
            "invoice_number": "",   # MUST fill
            "jobname": "",          # if unknown leave ""
            "date": "",             # MUST fill
            "total": "",            # MUST fill
            "page": page_num,
        }

        # ----------------------------------------------------
        # 🧩 STEP 1 — FILL IN Simple Defaults
        # ----------------------------------------------------
        inv["vendor"] = detect_vendor_name(text)
        inv["invoice_number"] = detect_invoice_number(text)
        inv["date"] = detect_invoice_date(text)
        inv["total"] = detect_invoice_total(text)
        inv["jobname"] = detect_jobname(text)

        results.append(inv)

    # If only one invoice → return dict
    if len(results) == 1:
        return results[0]

    return results


# ============================================================
# HELPER EXTRACTION FUNCTIONS — You customize these per vendor
# ============================================================

def detect_vendor_name(text):
    """Override this for each vendor."""
    # Example detects uppercase line at top
    match = re.search(r"^([A-Z][A-Z0-9 ]{3,})$", text, re.MULTILINE)
    return match.group(1).strip() if match else ""


def detect_invoice_number(text):
    """Override with vendor-specific rules."""
    # Example: look for 'Invoice #' or 'Inv #'
    patterns = [
        r"Invoice\s*#\s*([A-Za-z0-9\-]+)",
        r"Inv\s*#\s*([A-Za-z0-9\-]+)",
        r"Invoice Number[:\s]*([A-Za-z0-9\-]+)"
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def detect_invoice_date(text):
    """Override with vendor-specific date patterns."""
    # Example: 01/01/2025 or 2025-01-01
    patterns = [
        r"(\d{2}/\d{2}/\d{4})",
        r"(\d{4}-\d{2}-\d{2})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return m.group(1).strip()
    return ""


def detect_invoice_total(text):
    """Override for vendor-specific totals."""
    # Example:
    #   Total: $1,234.56
    patterns = [
        r"Total[:\s]*\$?([0-9\.,]+)",
        r"Amount Due[:\s]*\$?([0-9\.,]+)",
        r"Balance Due[:\s]*\$?([0-9\.,]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).replace(",", "").strip()
    return ""


def detect_jobname(text):
    """Leave blank or vendor-specific logic."""
    # Basic idea: look for a job code or name manually later
    return ""
