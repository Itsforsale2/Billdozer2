# ============================================================
# KNSplit.py
# Splits Knife River invoices (1 invoice per page)
# ============================================================

import os
import fitz  # PyMuPDF


def split(pdf_path: str, parsed_pages: list, output_folder: str):
    """
    Split a Knife River PDF where each page is one invoice.

    Arguments:
        pdf_path: full path to original PDF
        parsed_pages: list returned from knife_river_parser.parse_invoice()
        output_folder: folder to write split PDFs into

    Returns:
        number_of_files_created
    """

    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(pdf_path)
    count = 0

    for inv in parsed_pages:

        page_number = inv["page"] - 1

        # Extract invoice fields
        vendor = inv.get("vendor", "Unknown").replace(" ", "")
        job = inv.get("jobname", "").replace("/", "-").replace(" ", "")
        date = inv.get("date", "").replace("/", "-")
        total = inv.get("total", "")

        # Build filename
        filename = f"{vendor}_{job}_{date}_{total}.pdf"
        out_path = os.path.join(output_folder, filename)

        # Extract page
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=page_number, to_page=page_number)
        new_doc.save(out_path)
        new_doc.close()

        count += 1

    doc.close()
    return count
