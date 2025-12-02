import fitz

def extract(pdf):
    doc = fitz.open(pdf)
    for i, page in enumerate(doc):
        print("\n" + "="*60)
        print(f"PAGE {i+1}")
        print("="*60)
        print(page.get_text("text"))
        print("="*60)

pdf_path = r"C:/Python/Sample_PDFs/4 April/Core & Main/INVOICES_20250430_012458.PDF"
extract(pdf_path)
