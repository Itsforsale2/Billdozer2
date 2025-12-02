import os
import re
import tkinter as tk
from tkinter import filedialog


# ================================================================
# Convert vendor folder name → Python-safe parser module name
# ================================================================
def make_parser_name(folder_name: str) -> str:
    """
    Folder name stays EXACT on disk.
    Parser module name rules:
        - lowercase
        - remove '&'
        - convert '-' to space
        - remove all non-alphanumeric except space
        - collapse multiple spaces
        - convert spaces to underscores
        - append '_parser'
    """

    name = folder_name.lower()

    # Remove '&' entirely (Core & Main → core main)
    name = name.replace("&", "")

    # Hyphens become spaces (Tire-rama → tire rama)
    name = name.replace("-", " ")

    # Remove non-alphanumeric except spaces
    name = re.sub(r"[^a-z0-9 ]+", "", name)

    # Collapse multiple spaces
    name = re.sub(r"\s+", " ", name).strip()

    # Spaces → underscores
    name = name.replace(" ", "_")

    # Append required suffix
    return name + "_parser"


# ================================================================
# MAIN TOOL SCRIPT
# ================================================================
def main():
    root = tk.Tk()
    root.withdraw()

    print("\n=== Vendor Folder Name Extractor ===")

    base_folder = filedialog.askdirectory(title="Select Vendor Root Folder")
    if not base_folder:
        print("No folder selected. Exiting.")
        return

    print(f"\nSelected folder:\n{base_folder}\n")

    vendor_folders = [
        name for name in os.listdir(base_folder)
        if os.path.isdir(os.path.join(base_folder, name))
    ]

    if not vendor_folders:
        print("No vendor folders found.")
        return

    print("\n--- EXACT Vendor Folder Names (unchanged) ---\n")
    for vendor in vendor_folders:
        print(f" - {vendor}")

    print("\n\n--- Copy-Paste Ready ALIASES Dictionary ---\n")

    for vendor in vendor_folders:
        alias_key = vendor.lower().strip()
        parser_name = make_parser_name(vendor)
        print(f'    "{alias_key}": "{parser_name}",')

    print("\nDone.\n")


if __name__ == "__main__":
    main()
