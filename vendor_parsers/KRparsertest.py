import re
import fitz  # PyMuPDF

PDF_PATH = r"C:\Python\Sample_PDFs\Knife River\KR 2 pages.pdf"


def extract_pages(pdf_path):
    """Return a list of text blocks, one per page."""
    doc = fitz.open(pdf_path)
    return [page.get_text("text") for page in doc]


def get_job_name(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        if line.upper() == "ORIGINAL" and i > 0:
            prev = lines[i - 1]
            if not prev.isdigit():
                return prev
    return None


def get_invoice_date(text):
    # First MM/DD/YY on page
    m = re.search(r"\b\d{2}/\d{2}/\d{2}\b", text)
    return m.group() if m else None


def get_vendor_name(_text):
    return "Knife River"


def get_total(text):
    """Locate TOTAL on the current page only."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    total_index = None
    for i, line in enumerate(lines):
        if line.strip().upper() == "TOTAL":
            total_index = i
            break

    if total_index is None:
        return None

    # Numbers after TOTAL on THIS page
    amounts = []
    for line in lines[total_index + 1:]:
        matches = re.findall(r"\d[\d,]*\.\d{2}", line)
        for m in matches:
            amounts.append(float(m.replace(",", "")))

    if not amounts:
        return None

    # Totals often repeat 2–3 times
    for amt in amounts:
        if amounts.count(amt) >= 2:
            return amt

    return max(amounts)


# ------------------------------------------------------------
# MAIN — PROCESS EACH PAGE AS AN INVOICE
# ------------------------------------------------------------
if __name__ == "__main__":
    print(f"Reading PDF: {PDF_PATH}")
    print("--------------------------------------------------\n")

    pages = extract_pages(PDF_PATH)

    invoice_number = 1
    for text in pages:
        job = get_job_name(text)
        date = get_invoice_date(text)
        vendor = get_vendor_name(text)
        total = get_total(text)

        print(f"INVOICE #{invoice_number}")
        print("--------------------------------------------------")
        print("Job Name :", job)
        print("Vendor   :", vendor)
        print("Date     :", date)
        print("Total    :", total)
        print("--------------------------------------------------\n")

        invoice_number += 1
