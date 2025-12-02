import fitz
import re
import os
from openpyxl import Workbook
from tkinter import Tk, filedialog

# ============================================================
#  TEXT EXTRACTION (with path fix)
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
        except Exception as e:
            print(f"   ERROR reading page: {e}")

    return "\n".join(all_text)


# ============================================================
#  JOB NAME EXTRACTION
# ============================================================
def extract_job_name_from_text(text):
    """
    Extract the job name from a Knife River invoice.
    Job name = line immediately ABOVE the word 'ORIGINAL'.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    for i, line in enumerate(lines):
        if line.lower() == "original":
            if i > 0:
                candidate = lines[i - 1]
                # Exclude numeric-only or empty values
                if not candidate.isdigit() and len(candidate) > 1:
                    return candidate

    return None


# ============================================================
#  ITEM BLOCK HELPERS
# ============================================================
def is_truck_code(line):
    return bool(re.fullmatch(r"[A-Z]{3}\d", line))

def is_quantity_line(line):
    return bool(re.fullmatch(r"\d+(\.\d+)?\s*TN", line.upper()))

def is_unit_price(line):
    return bool(re.fullmatch(r"\d+(\.\d{1,4})?", line))

def is_extended_price(line):
    return bool(re.fullmatch(r"\d+(\.\d{2})", line))


# ============================================================
#  ITEM BLOCK EXTRACTION
# ============================================================
def extract_item_blocks(text):
    """Extract 6-line load blocks from invoice text."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    blocks = []
    current = []

    # Words we never want in a load block
    bad_words = {
        "item", "description", "special", "instructions",
        "subtotal", "total", "sales", "discount",
        "taxable", "nontaxable", "kr-mtn", "quantity"
    }

    for line in lines:

        if any(word in line.lower() for word in bad_words):
            continue

        # Ticket number starts block
        if line.isdigit() and 5 <= len(line) <= 7:
            current = [line]
            continue

        # Continue building block
        if current:
            current.append(line)

            if len(current) == 6:
                ticket, desc, truck, qty, unit, ext = current

                if (
                    is_truck_code(truck)
                    and is_quantity_line(qty)
                    and is_unit_price(unit)
                    and is_extended_price(ext)
                ):
                    blocks.append(current)

                current = []

    return blocks


# ============================================================
#  PARSE BLOCK INTO STRUCTURED DATA
# ============================================================
def parse_block(block, jobname):
    ticket, desc, truck, qty, unit, ext = block
    return {
        "job_name": jobname,
        "ticket": ticket,
        "description": desc,
        "truck": truck,
        "quantity_tons": float(qty.replace("TN", "").strip()),
        "unit_price": float(unit),
        "extended_price": float(ext),
    }


# ============================================================
#  BATCH PROCESSING (folder → many PDFs)
# ============================================================
def batch_process_folder():
    """Let user pick a folder and process all PDFs inside."""

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
    if not folder:
        print("Batch cancelled.")
        return None

    folder = os.path.normpath(folder)

    pdf_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    ]

    if not pdf_files:
        print("No PDF files found.")
        return None

    print(f"\nFound {len(pdf_files)} PDF files:")
    for f in pdf_files:
        print(" -", os.path.basename(f))

    all_items = []
    bad_files = 0

    for pdf in pdf_files:
        pdf = os.path.normpath(pdf)
        print(f"\nProcessing: {os.path.basename(pdf)}")

        # Try opening PDF
        try:
            text = extract_pdf_text(pdf)
        except Exception as e:
            print(f"   Skipping file (unreadable): {e}")
            bad_files += 1
            continue

        # FILTER 1 — must contain ORIGINAL
        if "ORIGINAL" not in text.upper():
            print("   Skipping (not a haul invoice – no ORIGINAL found)")
            continue

        # FILTER 2 — must contain at least one truck code
        words = text.split()
        has_truck = any(re.fullmatch(r"[A-Z]{3}\d", w) for w in words)
        if not has_truck:
            print("   Skipping (no truck codes found)")
            continue

        jobname = extract_job_name_from_text(text)
        blocks = extract_item_blocks(text)

        if not blocks:
            print("   No load blocks found on this file.")
            continue

        # Parse load blocks
        for block in blocks:
            item = parse_block(block, jobname)
            all_items.append(item)

    print(f"\nFinished batch.")
    print(f"Extracted loads: {len(all_items)}")
    print(f"Unreadable PDFs skipped: {bad_files}")

    return all_items


# ============================================================
#  EXPORT TO EXCEL
# ============================================================
def export_batch_to_excel(all_items):
    """Save ALL parsed items to one Excel workbook."""

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
    ws.title = "Loads"

    ws.append([
        "Job Name", "Ticket", "Description", "Truck",
        "Quantity (Tons)", "Unit Price", "Extended Price"
    ])

    for item in all_items:
        ws.append([
            item["job_name"],
            item["ticket"],
            item["description"],
            item["truck"],
            item["quantity_tons"],
            item["unit_price"],
            item["extended_price"]
        ])

    wb.save(save_path)
    print(f"\nBatch Excel exported to:\n{save_path}")


# ============================================================
#  MAIN
# ============================================================
if __name__ == "__main__":
    print("Select a folder to batch process all PDFs...")

    all_items = batch_process_folder()

    if all_items:
        print("\nExporting to Excel...")
        export_batch_to_excel(all_items)
    else:
        print("\nNo items extracted.")
