import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog

def main():
    # Bare minimum Tk root (hidden)
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Select a PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_path:
        print("No file selected.")
        return

    print(f"\n--- TEXT FROM: {pdf_path} ---\n")

    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            print(f"\n=== PAGE {i+1} ===\n")
            print(page.get_text("text"))
        doc.close()
    except Exception as e:
        print(f"Error reading PDF: {e}")

if __name__ == "__main__":
    main()
