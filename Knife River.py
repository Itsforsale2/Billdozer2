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
#  BATCH PROCESSING (auto-find matching subfolders)
# ============================================================
def batch_process_folder():
    """
    User selects a YEAR folder (e.g., '2025').
    Script finds ALL subfolders named exactly like the script filename.
    Then processes all PDFs inside them.
    """

    # ----------------------------
    # Get this script's name
    # e.g. "Knife River.py" → "Knife River"
    # ----------------------------
    SCRIPT_NAME = os.path.splitext(os.path.basename(__file__))[0].strip()

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    top_folder = filedialog.askdirectory(title=f"Select top folder (contains subfolders with '{SCRIPT_NAME}')")
    if not top_folder:
        print("Batch cancelled.")
        return None

    top_folder = os.path.normpath(top_folder)

    # ============================================================
    # === NEW FOLDER SEARCH LOGIC: find all subfolders matching script name
    # ============================================================
    matched_folders = []

    for root_dir, dirs, files in os.walk(top_folder):
        for d in dirs:
            if d.lower() == SCRIPT_NAME.lower():
                full_path = os.path.join(root_dir, d)
                matched_folders.append(full_path)

    if not matched_folders:
        print(f"\nNo folders named '{SCRIPT_NAME}' found under:\n{top_folder}")
        return None

    print("\nFound matching folders:")
    for f in matched_folders:
        print(" -", f)

    # ============================================================
    # Process PDFs from all matched folders
    # ============================================================
    all_items = []
    bad_files = 0

    for folder in matched_folders:
        print(f"\n📂 Processing folder: {folder}")

        pdf_files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ]

        if not pdf_files:
            print("   No PDF files found in this folder.")
            continue

        for pdf in pdf_files:
            pdf = os.path.normpath(pdf)
            print(f"\nProcessing: {os.path.basename(pdf)}")

            try:
                text = extract_pdf_text(pdf)
            except Exception as e:
                print(f"   Skipping unreadable file: {e}")
                bad_files += 1
                continue

            if "ORIGINAL" not in text.upper():
                print("   Skipping (not a haul invoice – no ORIGINAL found)")
                continue

            words = text.split()
            has_truck = any(re.fullmatch(r"[A-Z]{3}\d", w) for w in words)
            if not has_truck:
                print("   Skipping (no truck codes found)")
                continue

            jobname = extract_job_name_from_text(text)
            blocks = extract_item_blocks(text)

            if not blocks:
                print("   No load blocks found.")
                continue

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
    print("Select the TOP folder (e.g., 2025)...")

    all_items = batch_process_folder()

    if all_items:
        print("\nExporting to Excel...")
        export_batch_to_excel(all_items)
    else:
        print("\nNo items extracted.")
