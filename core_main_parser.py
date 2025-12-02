import fitz
import re

def parse_invoice(pdf_path):
    doc = fitz.open(pdf_path)
    results = []

    SKIP_PREFIXES = [
        "job #",
        "bill of lading",
        "shipped via",
        "invoice#",
        "invoice #",
        "invoice",
        "date ordered",
        "date shipped",
    ]

    DATE_RE = re.compile(r"^[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}$")

    for page_index, page in enumerate(doc):
        text = page.get_text("text")
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        # -------------------------------
        # INVOICE NUMBER
        # -------------------------------
        m_inv = re.search(r"Invoice\s*#\s*\n\s*([A-Z0-9]+)", text, re.IGNORECASE)
        invoice_number = m_inv.group(1).strip() if m_inv else ""

        # -------------------------------
        # DATE
        # -------------------------------
        m_date = re.search(r"Invoice Date\s*\n\s*([0-9/]+)", text)
        date_str = m_date.group(1).strip() if m_date else ""

        # -------------------------------
        # TOTAL
        # -------------------------------
        m_total = re.search(r"Total Amount Due\s*\n\s*\$?([0-9,]+\.[0-9]+)", text)
        total = m_total.group(1).replace(",", "") if m_total else ""

        # -------------------------------
        # JOB NAME (FINAL WORKING VERSION)
        # -------------------------------
        jobname = "UNKNOWN"

        for idx, line in enumerate(lines):
            if "customer po #" in line.lower() and "job name" in line.lower():

                for nxt in lines[idx+1:]:

                    low = nxt.lower()

                    # skip any junk lines
                    if any(low.startswith(prefix) for prefix in SKIP_PREFIXES):
                        continue

                    # skip dates (this was the problem)
                    if DATE_RE.match(nxt):
                        continue

                    # skip invoice-like codes
                    if re.match(r"^[A-Z0-9]{6,}$", nxt):
                        continue

                    # FOUND REAL JOB NAME
                    jobname = nxt.strip()
                    break

                break

        # ------------------------------------------------------------
        # PATCH 3 — SECONDARY JOBNAME FALLBACK (correct insertion)
        # ------------------------------------------------------------
        if jobname == "UNKNOWN":
            # Look for a standalone "Job Name" anchor
            for idx, line in enumerate(lines):
                if line.lower() == "job name":

                    for nxt in lines[idx+1:]:
                        low = nxt.lower()

                        # skip junk terms
                        if any(low.startswith(prefix) for prefix in SKIP_PREFIXES):
                            continue

                        # skip dates
                        if DATE_RE.match(nxt):
                            continue

                        # skip invoice-like codes
                        if re.match(r"^[A-Z0-9]{6,}$", nxt):
                            continue

                        # skip literal job markers
                        if low in ("job #", "job#", "job no", "job number"):
                            continue

                        jobname = nxt.strip()
                        break

                if jobname != "UNKNOWN":
                    break
        # ------------------------------------------------------------
        # END PATCH 3
        # ------------------------------------------------------------

        results.append({
            "vendor": "Core & Main",
            "invoice_number": invoice_number,
            "jobname": jobname,
            "date": date_str,
            "total": total,
            "page": page_index + 1,
        })


    return results
