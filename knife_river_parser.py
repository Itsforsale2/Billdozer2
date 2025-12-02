import fitz  # PyMuPDF
import re


# ============================================================
# LOW-LEVEL: EXTRACT TEXT FROM A SINGLE PAGE
# ============================================================
def _extract_page_text(doc, page_index: int) -> str:
    """
    Return raw text for a single page, as a string.
    """
    page = doc.load_page(page_index)
    txt = page.get_text("text")  # pure text mode
    return txt.replace("\x00", "").strip()


# ============================================================
# PARSE ONE KNIFE RIVER INVOICE PAGE
# ============================================================
def _parse_knife_river_page(raw_text: str, page_number: int) -> dict:
    """
    Given the raw text for ONE page of a Knife River invoice,
    extract: vendor, invoice_number, jobname, date, total, raw_text, page.
    """
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

    # -----------------------------
    # Vendor is always Knife River
    # -----------------------------
    vendor = "Knife River"

    # -----------------------------
    # Invoice Number:
    # first line that is 6+ digits only (e.g. 968457, 940775)
    # -----------------------------
    invoice_number = ""
    for ln in lines:
        if re.fullmatch(r"\d{6,}", ln):
            invoice_number = ln
            break

    # -----------------------------
    # Date:
    # first MM/DD/YY pattern, e.g. 09/08/25
    # -----------------------------
    date = ""
    for ln in lines:
        m = re.search(r"\b\d{2}/\d{2}/\d{2}\b", ln)
        if m:
            date = m.group(0)
            break

    # ============================================================
    # TOTAL (PATCHED — CORRECTLY CAPTURE NUMBERS WITH COMMAS)
    # ============================================================
    total = ""

    # Matches:
    #   440.70
    #   1,025.28
    #   12,345.67
    total_matches = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b", raw_text)

    if total_matches:
        # Take the LAST money value on the page (Knife River puts total last)
        total = total_matches[-1].replace(",", "")

    # ============================================================
    # JOBNAME:
    # Line directly before "ORIGINAL" most of the time.
    # ============================================================
    jobname = ""
    for idx, ln in enumerate(lines):
        if ln.upper() == "ORIGINAL":
            candidate = lines[idx - 1] if idx - 1 >= 0 else ""
            candidate2 = lines[idx - 2] if idx - 2 >= 0 else ""

            bad_words = {"INVOICE", "TICKET", "PAYABLE COPY", "SUBTOTAL", "TOTAL"}

            cand_up = candidate.upper()
            cand2_up = candidate2.upper()

            if candidate and cand_up not in bad_words:
                jobname = candidate
            elif candidate2 and cand2_up not in bad_words:
                jobname = candidate2

            break

    if not jobname:
        jobname = ""

    # -----------------------------
    # Build result dict for this page
    # -----------------------------
    return {
        "vendor": vendor,
        "invoice_number": invoice_number,
        "jobname": jobname,
        "date": date,
        "total": total,
        "raw_text": raw_text,
        "page": page_number,
    }


# ============================================================
# PUBLIC API – THIS IS WHAT YOUR UI SHOULD CALL
# ============================================================
def parse_invoice(pdf_path: str):
    """
    Parse a Knife River PDF that may contain multiple invoices
    (one invoice per page).
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise RuntimeError(f"Could not open PDF '{pdf_path}': {e}")

    results = []

    for page_index in range(len(doc)):
        raw_text = _extract_page_text(doc, page_index)
        parsed = _parse_knife_river_page(raw_text, page_index + 1)
        results.append(parsed)

    doc.close()
    return results


# ============================================================
# OPTIONAL: CONSOLE OUTPUT WHEN RUN DIRECTLY
# ============================================================
def _build_display_text(parsed_pages):
    blocks = []
    for p in parsed_pages:
        header = [
            f"===== INVOICE PAGE {p['page']} =====",
            f"Vendor:         {p['vendor']}",
            f"Invoice Number: {p['invoice_number']}",
            f"Jobname:        {p['jobname']}",
            f"Date:           {p['date']}",
            f"Total:          {p['total']}",
            "",
            "--- RAW TEXT ---",
            p["raw_text"],
            "",
        ]
        blocks.append("\n".join(header))
    return "\n".join(blocks).rstrip()


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python knife_river_parser.py file.pdf")
        sys.exit(0)

    pdf = sys.argv[1]
    pages = parse_invoice(pdf)
    print(_build_display_text(pages))
