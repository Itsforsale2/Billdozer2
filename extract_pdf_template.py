import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# ============================================================
# extract_pdf_template.py
# ------------------------------------------------------------
# This script helps you create vendor parser templates.
# Pick a sample PDF → extract text → view & save it cleanly.
# ============================================================


def extract_pdf_text(pdf_path):
    """Return list: [ (page_number, text), ... ]"""
    doc = fitz.open(pdf_path)
    pages = []

    for i, page in enumerate(doc):
        text = page.get_text("text")
        pages.append((i + 1, text))

    doc.close()
    return pages


def format_extracted_text(pages):
    """
    Build a clean and numbered text dump to help writing templates.
    Adds page separators and line numbers.
    """
    out_lines = []

    for page_num, text in pages:
        out_lines.append("=" * 60)
        out_lines.append(f"===================== PAGE {page_num} =====================")
        out_lines.append("=" * 60)

        for i, line in enumerate(text.splitlines()):
            line_num = str(i + 1).rjust(4)
            out_lines.append(f"{line_num}: {line}")

        out_lines.append("")  # blank line

    return "\n".join(out_lines)


def pick_pdf():
    """Open file picker to select a PDF."""
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Select sample vendor PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return pdf_path


def main():
    pdf_path = pick_pdf()
    if not pdf_path:
        print("No PDF selected. Exiting.")
        return

    print(f"\nExtracting: {pdf_path}\n")

    try:
        pages = extract_pdf_text(pdf_path)
    except Exception as e:
        print(f"Could not read PDF:\n{e}")
        return

    formatted = format_extracted_text(pages)

    # Print to console
    print(formatted)

    # Save .txt next to the PDF
    out_txt = os.path.splitext(pdf_path)[0] + "_EXTRACTED.txt"
    try:
        with open(out_txt, "w", encoding="utf-8") as f:
            f.write(formatted)
        print(f"\nSaved extract to:\n{out_txt}")
    except Exception as e:
        print(f"Could not save text output:\n{e}")

    # Optional success dialog
    try:
        messagebox.showinfo("Done", f"Text extracted.\nSaved:\n{out_txt}")
    except:
        pass


if __name__ == "__main__":
    main()
