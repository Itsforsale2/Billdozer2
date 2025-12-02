import fitz  # PyMuPDF
import re
import os


# ============================================================
# LOW-LEVEL: EXTRACT TEXT FROM A SINGLE PAGE
# ============================================================
def _extract_page_text(doc, page_index: int) -> str:
    """
    Return raw text for a single page.
    """
    page = doc.load_page(page_index)
    txt = page.get_text("text")
    return txt.replace("\x00", "").strip()


# ============================================================
# MISS0ULA LANDFILL — INVOICE NUMBER EXTRACTION (FIXED)
# ============================================================
def _extract_missoula_invoice_number(raw_text: str) -> str:
    """
    Improved Missoula Landfill invoice number extractor.
    Always returns the correct 6–7 digit invoice number.
    """
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

    # ---------------------------------------------------------
    # 1) Most reliable spot: lines immediately after SIGNATURE
    # ---------------------------------------------------------
    for i, ln in enumerate(lines):
        if "SIGNATURE" in ln.upper():
            for j in range(i + 1, min(i + 5, len(lines))):
                # A pure 6–7 digit number
                if re.fullmatch(r"\d{6,7}", lines[j]):
                    return lines[j]

    # ---------------------------------------------------------
    # 2) Standalone 6–7 digit line (skip "01", weight numbers)
    # ---------------------------------------------------------
    for ln in lines:
        if re.fullmatch(r"\d{6,7}", ln):
            return ln

    # ---------------------------------------------------------
    # 3) Fallback: any 6–7 digit number anywhere
    # ---------------------------------------------------------
    m = re.search(r"\b(\d{6,7})\b", raw_text)
    return m.group(1) if m else ""


# ============================================================
# EXTRACT DATE
# ============================================================
def _extract_missoula_date(lines) -> str:
    """
    First MM/DD/YY style date.
    """
    date_pattern = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2}\b")
    for ln in lines:
        m = date_pattern.search(ln)
        if m:
            return m.group(0)
    return ""


# ============================================================
# EXTRACT TOTAL
# ============================================================
def _extract_missoula_total(raw_text: str) -> str:
    """
    Choose largest non-zero $XXX.XX.
    """
    money_matches = re.findall(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})", raw_text)
    if not money_matches:
        return ""

    parsed = [(float(x.replace(",", "")), x) for x in money_matches]
    nonzero = [p for p in parsed if p[0] > 0]

    if nonzero:
        _, best_str = max(nonzero, key=lambda t: t[0])
    else:
        _, best_str = max(parsed, key=lambda t: t[0])

    return best_str.replace(",", "")


# ============================================================
# JOBNAME LOGIC
# ============================================================
def _extract_missoula_jobname(lines) -> str:
    """
    Job name appears after the 2nd date.
    """
    date_pattern = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2}\b")
    date_indices = []

    for idx, ln in enumerate(lines):
        if date_pattern.search(ln):
            date_indices.append(idx)

    if len(date_indices) < 2:
        return _fallback_jobname(lines)

    start_idx = date_indices[1] + 1
    weight_keywords = {"GROSS", "TARE", "NET", "WEIGHT", "SCALE", "INBOUND"}

    for i in range(start_idx, min(start_idx + 6, len(lines))):
        ln = lines[i].strip()
        up = ln.upper()

        if not ln:
            continue
        if ln.isdigit():  # skip scale numbers like "01"
            continue
        if any(w in up for w in weight_keywords):
            continue

        return ln

    return _fallback_jobname(lines)


def _fallback_jobname(lines):
    """
    Simple uppercase jobname fallback.
    """
    for ln in lines:
        if re.fullmatch(r"[A-Z0-9 ]{3,30}", ln):
            up = ln.upper()
            bad = [
                "PAYMENT", "GRANT", "CREEK", "EXCAVATING", "MISSOULA",
                "LANDFILL", "GROSS", "TARE", "NET", "WEIGHT", "INVOICE",
                "INBOUND", "SCALE", "SIGNATURE"
            ]
            if not any(b in up for b in bad):
                return ln
    return ""


# ============================================================
# CORE PARSER FOR 1 PAGE
# ============================================================
def _parse_missoula_landfill_page(raw_text: str, page_number: int) -> dict:
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

    vendor = "Missoula Landfill"
    invoice_number = _extract_missoula_invoice_number(raw_text)
    date = _extract_missoula_date(lines)
    total = _extract_missoula_total(raw_text)
    jobname = _extract_missoula_jobname(lines)

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
# PUBLIC API – PARSE ENTIRE PDF
# ============================================================
def parse_invoice(pdf_path: str):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise RuntimeError(f"Could not open PDF '{pdf_path}': {e}")

    results = []
    for page_index in range(len(doc)):
        raw_text = _extract_page_text(doc, page_index)
        parsed = _parse_missoula_landfill_page(raw_text, page_index + 1)
        results.append(parsed)

    doc.close()
    return results


# ============================================================
# CONSOLE PREVIEW BUILDER
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


# ============================================================
# FILENAME BUILDER (FORMAT A)
# ============================================================
def _clean_for_filename(value: str) -> str:
    if not value:
        return ""
    v = value.strip()
    v = v.replace("/", "-")
    v = v.replace(" ", "")
    v = re.sub(r'[\\/:*?"<>|]', "", v)
    return v


def build_output_filename(inv: dict) -> str:
    """
    Build filename in Format A:
    Vendor_Job_Date_Invoice_Total.pdf
    """
    vendor = _clean_for_filename(inv.get("vendor", "Vendor"))
    job = _clean_for_filename(inv.get("jobname", "Job"))
    date = _clean_for_filename(inv.get("date", ""))
    invoice = _clean_for_filename(inv.get("invoice_number", ""))
    total = _clean_for_filename(inv.get("total", ""))

    if not invoice:
        invoice = "NOINV"

    parts = [vendor, job, date, invoice, total]
    parts = [p for p in parts if p]  # keep invoice even if empty originally

    base = "_".join(parts)
    return f"{base}.pdf"


# ============================================================
# SAVE SPLIT PAGES INTO PROCESSED/
# ============================================================
def save_split_invoices(pdf_path: str) -> int:
    parsed_pages = parse_invoice(pdf_path)
    if not parsed_pages:
        return 0

    doc = fitz.open(pdf_path)
    folder = os.path.dirname(pdf_path)
    processed_folder = os.path.join(folder, "processed")
    os.makedirs(processed_folder, exist_ok=True)

    count = 0
    for inv in parsed_pages:
        page_index = inv.get("page", 1) - 1

        try:
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_index, to_page=page_index)

            file_name = build_output_filename(inv)
            out_path = os.path.join(processed_folder, file_name)

            new_doc.save(out_path)
            new_doc.close()
            count += 1

        except Exception:
            continue

    doc.close()
    return count


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage:")
        print("  python missoula_landfill_parser.py file.pdf")
        sys.exit(0)

    target = sys.argv[1]

    if os.path.isdir(target):
        total_files = 0
        total_invoices = 0
        for name in os.listdir(target):
            if not name.lower().endswith(".pdf"):
                continue
            pdf_full = os.path.join(target, name)
            created = save_split_invoices(pdf_full)
            total_files += 1
            total_invoices += created
            print(f"{name}: saved {created} invoice PDF(s).")

        print(f"\nDone. {total_invoices} invoice PDF(s) saved from {total_files} file(s).")

    else:
        created = save_split_invoices(target)
        print(f"Saved {created} invoice PDF(s) into 'processed' folder.\n")

        pages = parse_invoice(target)
        print(_build_display_text(pages))
