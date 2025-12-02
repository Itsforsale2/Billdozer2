import tkinter as tk
from tkinter import filedialog

import core_main_parser

root = tk.Tk()
root.withdraw()

print("\n=== Core & Main Parser Test ===\n")

pdf = filedialog.askopenfilename(
    title="Select Core & Main PDF",
    filetypes=[("PDF Files", "*.pdf")]
)

print(f"\nSelected PDF:\n{pdf}\n")

result = core_main_parser.parse_invoice(pdf)
print("\n--- PARSED RESULT ---\n")
print(result)
print("\nDone.\n")
