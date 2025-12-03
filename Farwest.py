import fitz
import re
import os
from openpyxl import Workbook
from tkinter import Tk, filedialog


# ============================================================
#  TEXT EXTRACTION
# ============================================================
def extract_pdf_text(pdf_path):
    """Return all readable text from a text-based PDF."""
    pdf_path = os.path.normpath(pdf_path)

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise RuntimeError(f"Cannot open PDF (corrupted or unreadable): {e}")

    all_text = []
    for page in doc:
        try:
            text = page.get_text("text")
            all_text.append(text)
        except:
            pass

    return "\n".join(all_text)


# ============================================================
#  FARWEST FIELD EXTRACTION
# ============================================================
def extract_invoice_number(text):
    """
    Invoice # is always on a line immediately following 'Invoice #'
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for i, ln in enumerate(lines):
        if ln.lower() == "invoice #":
            if i + 1 < len(lines):
                return lines[i + 1].replace("#", "").strip()
    return None


def extract_job_name(text):
    """
    Job name appears immediately after the line 'JOB'
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for i, ln in enumerate(lines):
        if ln.lower() == "job":
            if i + 1 < len(lines):
                return lines[i + 1].strip()
    return None


# ============================================================
#  ITEM EXTRACTION (FARWEST)
# ============================================================
def extract_line_items(text, job_name, invoice_number):
    """
    Farwest format looks like:

        Tons    Amount
        Date        Description          $/Ton
        65.31       641.34               10/1/2025    3/4 Base   9.82
        14.67       144.06               10/22/2025   3/4 Base   9.82

    We detect blocks using this pattern:
        quantity
        extended price
        date
        description
        unit price
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    items = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Quantity tons (float)
        if re.fullmatch(r"\d+(\.\d+)?", line):
            qty = float(line)

            # Extended amount
            if i + 1 < len(lines) and re.fullmatch(r"\d+(\.\d+)?", lines[i + 1]):
                extended = float(lines[i + 1])
            else:
                i += 1
                continue

            # Date (MM/DD/YYYY style)
            if i + 2 < len(lines) and re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", lines[i + 2]):
                date = lines[i + 2]
            else:
                i += 1
                continue

            # Description (string)
            description = lines[i + 3] if (i + 3 < len(lines)) else ""

            # Unit Price
            if i + 4 < len(lines) and re.fullmatch(r"\d+(\.\d+)?", lines[i + 4]):
                unit_price = float(lines[i + 4])
            else:
                unit_price = 0.0

            # Build record
            items.append({
                "job_name": job_name,
                "invoice_number": invoice_number,
                "date": date,
                "description": description,
                "quantity_tons": qty,
                "unit_price": unit_price,
                "extended_price": extended
            })

            # Move pointer
            i += 5
            continue

        i += 1

    return items


# ============================================================
#  BATCH PROCESSING
# ============================================================
def batch_process_folder():
    """
    Find all subfolders named after this script and process all PDFs.
    Example: script name = Farwest.py → looks for folders named 'Farwest'
    """

    SCRIPT_NAME = os.path.splitext(os.path.basename(__file__))[0].strip()

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    top_folder = filedialog.askdirectory(title=f"Select top folder (contains subfolders with '{SCRIPT_NAME}')")
    if not top_folder:
        print("Batch cancelled.")
        return None

    matched_folders = []

    for root_dir, dirs, files in os.walk(top_folder):
        for d in dirs:
            if d.lower() == SCRIPT_NAME.lower():
                matched_folders.append(os.path.join(root_dir, d))

    if not matched_folders:
        print(f"\nNo folders named '{SCRIPT_NAME}' found under:\n{top_folder}")
        return None

    print("\nFound matching folders:")
    for f in matched_folders:
        print(" -", f)

    all_items = []

    for folder in matched_folders:
        pdfs = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ]

        for pdf in pdfs:
            print(f"\nProcessing: {os.path.basename(pdf)}")
            text = extract_pdf_text(pdf)

            invoice_number = extract_invoice_number(text)
            job_name = extract_job_name(text)
            items = extract_line_items(text, job_name, invoice_number)

            if items:
                all_items.extend(items)
            else:
                print("   ⚠ No line items found in this invoice.")

    print(f"\nExtracted line items: {len(all_items)}")
    return all_items


# ============================================================
#  EXPORT TO EXCEL
# ============================================================
def export_batch_to_excel(all_items):
    if not all_items:
        print("Nothing to export.")
        return

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Batch Results As"
    )

    if not save_path:
        print("Export cancelled.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Farwest Loads"

    ws.append([
        "Job Name", "Invoice Number", "Date",
        "Description", "Quantity (Tons)",
        "Unit Price", "Extended Price"
    ])

    for it in all_items:
        ws.append([
            it["job_name"],
            it["invoice_number"],
            it["date"],
            it["description"],
            it["quantity_tons"],
            it["unit_price"],
            it["extended_price"]
        ])

    wb.save(save_path)
    print(f"\nSaved:\n{save_path}")


# ============================================================
#  MAIN
# ============================================================
if __name__ == "__main__":
    print("Select the TOP folder (e.g., 2025)...")
    items = batch_process_folder()
    if items:
        export_batch_to_excel(items)
