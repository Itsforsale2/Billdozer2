import os
import fitz  # PyMuPDF
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ------------------------------------------------------------
# Utility: extract text from PDF (page by page)
# ------------------------------------------------------------
def extract_pdf_text(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text("text")
        pages.append((i + 1, text))
    doc.close()
    return pages


class TemplateWizard(tk.Tk):
    """
    Wizard to build vendor-specific parser rules.

    FLOW:
      1. Load sample PDF.
      2. Highlight values in the text viewer.
      3. Click "Set as Vendor / Invoice / Date / Total / Jobname / Work #".
         - Wizard records the VALUE and an ANCHOR near it.
      4. Click "Preview Parser" to see the generated vendor_parser.py code.
      5. Click "Test Rules" to run the regex patterns on the sample PDF and
         see what will be extracted.
      6. Click "Save Parser..." to write vendor_parser.py to disk.
    """

    def __init__(self):
        super().__init__()
        self.title("Invoice Parser Template Wizard")
        self.geometry("1200x800")

        # Selected PDF info
        self.pdf_path = None
        self.pages = []         # list of (page_num, text)
        self.current_page_idx = 0

        # Field configs: each has {"value", "anchor", "pattern"}
        self.fields = {
            "vendor":      {"value": "", "anchor": "", "pattern": ""},
            "invoice_num": {"value": "", "anchor": "", "pattern": ""},
            "date":        {"value": "", "anchor": "", "pattern": ""},
            "total":       {"value": "", "anchor": "", "pattern": ""},
            "jobname":     {"value": "", "anchor": "", "pattern": ""},
            "work_number": {"value": "", "anchor": "", "pattern": ""},
        }

        self.build_ui()

    # --------------------------------------------------------
    # UI LAYOUT
    # --------------------------------------------------------
    def build_ui(self):
        # Top controls (load PDF, page navigation, vendor name for file)
        top_frame = tk.Frame(self)
        top_frame.pack(side="top", fill="x", padx=8, pady=4)

        btn_load = tk.Button(top_frame, text="Load Sample PDF", command=self.load_pdf)
        btn_load.pack(side="left")

        self.page_label = tk.Label(top_frame, text="Page: - / -")
        self.page_label.pack(side="left", padx=10)

        btn_prev = tk.Button(top_frame, text="Prev Page", command=self.prev_page)
        btn_prev.pack(side="left")

        btn_next = tk.Button(top_frame, text="Next Page", command=self.next_page)
        btn_next.pack(side="left", padx=(4, 20))

        tk.Label(top_frame, text="Vendor name (used for parser filename):").pack(side="left")
        self.vendor_name_entry = tk.Entry(top_frame, width=30)
        self.vendor_name_entry.pack(side="left", padx=4)

        # Split main area: left = PDF text view, right = field info + preview
        main_pane = tk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True)

        left_frame = tk.Frame(main_pane)
        right_frame = tk.Frame(main_pane)
        main_pane.add(left_frame, stretch="always")
        main_pane.add(right_frame, stretch="always")

        # LEFT: text widget (PDF text)
        left_top = tk.Frame(left_frame)
        left_top.pack(fill="x", padx=4, pady=4)

        anchor_help = (
            "Instructions:\n"
            "1. Use the mouse to highlight a VALUE (e.g. invoice number).\n"
            "2. Click one of the 'Set as ...' buttons.\n"
            "   The wizard will also grab nearby text as an ANCHOR.\n\n"
            "Anchor = stable label near the value, used as a hook in regex.\n"
            "Example: 'Invoice # 1635034'\n"
            "  Anchor: 'Invoice #'    Value: '1635034'\n"
            "Later, regex searches for 'Invoice #' and captures what follows."
        )
        tk.Label(left_top, text=anchor_help, justify="left", anchor="w").pack(fill="x")

        # Text widget with scrollbars
        text_frame = tk.Frame(left_frame)
        text_frame.pack(fill="both", expand=True, padx=4, pady=4)

        self.text = tk.Text(text_frame, wrap="none")
        self.text.pack(side="left", fill="both", expand=True)

        yscroll = tk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        yscroll.pack(side="right", fill="y")
        self.text.configure(yscrollcommand=yscroll.set)

        xscroll = tk.Scrollbar(left_frame, orient="horizontal", command=self.text.xview)
        xscroll.pack(side="bottom", fill="x")
        self.text.configure(xscrollcommand=xscroll.set)

        # Buttons to set fields from selection
        field_btn_frame = tk.Frame(left_frame)
        field_btn_frame.pack(fill="x", padx=4, pady=(0, 4))

        tk.Button(field_btn_frame, text="Set as Vendor",
                  command=lambda: self.set_field_from_selection("vendor")).pack(side="left", padx=2)
        tk.Button(field_btn_frame, text="Set as Invoice #",
                  command=lambda: self.set_field_from_selection("invoice_num")).pack(side="left", padx=2)
        tk.Button(field_btn_frame, text="Set as Date",
                  command=lambda: self.set_field_from_selection("date")).pack(side="left", padx=2)
        tk.Button(field_btn_frame, text="Set as Total",
                  command=lambda: self.set_field_from_selection("total")).pack(side="left", padx=2)
        tk.Button(field_btn_frame, text="Set as Jobname",
                  command=lambda: self.set_field_from_selection("jobname")).pack(side="left", padx=2)
        tk.Button(field_btn_frame, text="Set as Work #",
                  command=lambda: self.set_field_from_selection("work_number")).pack(side="left", padx=2)

        # RIGHT: field summary + parser preview + test output
        right_top = tk.Frame(right_frame)
        right_top.pack(fill="both", expand=True, padx=4, pady=4)

        # Field summary
        field_frame = tk.LabelFrame(right_top, text="Field Rules (Value + Anchor + Regex)")
        field_frame.pack(fill="x", padx=4, pady=4)

        self.field_labels = {}
        for key, label in [
            ("vendor", "Vendor"),
            ("invoice_num", "Invoice #"),
            ("date", "Date"),
            ("total", "Total"),
            ("jobname", "Jobname"),
            ("work_number", "Work #"),
        ]:
            row = tk.Frame(field_frame)
            row.pack(fill="x", padx=2, pady=1)
            tk.Label(row, text=f"{label}:", width=10, anchor="w").pack(side="left")
            var = tk.StringVar()
            lbl = tk.Label(row, textvariable=var, anchor="w", justify="left")
            lbl.pack(side="left", fill="x", expand=True)
            self.field_labels[key] = var

        # Buttons for preview / test / save
        btn_frame = tk.Frame(right_top)
        btn_frame.pack(fill="x", padx=4, pady=4)

        tk.Button(btn_frame, text="Preview Parser Code",
                  command=self.preview_parser).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Test Rules on Sample PDF",
                  command=self.test_rules).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Save vendor_parser.py",
                  command=self.save_parser).pack(side="left", padx=4)

        # Parser preview text
        preview_frame = tk.LabelFrame(right_top, text="Generated parser code (preview)")
        preview_frame.pack(fill="both", expand=True, padx=4, pady=4)

        self.preview_text = tk.Text(preview_frame, height=16, wrap="none")
        self.preview_text.pack(side="left", fill="both", expand=True)

        prev_scroll = tk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        prev_scroll.pack(side="right", fill="y")
        self.preview_text.configure(yscrollcommand=prev_scroll.set)

        # Test result text
        test_frame = tk.LabelFrame(right_frame, text="Test Extraction Result")
        test_frame.pack(fill="both", expand=True, padx=4, pady=4)

        self.test_output = tk.Text(test_frame, height=10, wrap="word")
        self.test_output.pack(fill="both", expand=True)

    # --------------------------------------------------------
    # PDF navigation
    # --------------------------------------------------------
    def load_pdf(self):
        path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.pages = extract_pdf_text(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read PDF:\n{e}")
            return

        self.pdf_path = path
        self.current_page_idx = 0
        self.show_current_page()

    def show_current_page(self):
        self.text.delete("1.0", tk.END)
        if not self.pages:
            self.page_label.config(text="Page: - / -")
            return
        page_num, text = self.pages[self.current_page_idx]
        self.text.insert("1.0", text)
        self.page_label.config(
            text=f"Page: {self.current_page_idx + 1} / {len(self.pages)}"
        )

    def prev_page(self):
        if not self.pages:
            return
        if self.current_page_idx > 0:
            self.current_page_idx -= 1
            self.show_current_page()

    def next_page(self):
        if not self.pages:
            return
        if self.current_page_idx < len(self.pages) - 1:
            self.current_page_idx += 1
            self.show_current_page()

    # --------------------------------------------------------
    # Set field from highlighted selection
    # --------------------------------------------------------
    def set_field_from_selection(self, field_key):
        try:
            selected = self.text.get("sel.first", "sel.last")
        except tk.TclError:
            messagebox.showwarning("No selection", "Please highlight a value first.")
            return

        selected = selected.strip()
        if not selected:
            messagebox.showwarning("Empty selection", "Please select some text (not just spaces).")
            return

        # Grab full page text
        full_text = self.text.get("1.0", "end-1c")

        # Find position of selected text in full_text
        # (If it appears multiple times, we'll use the first occurrence.)
        pos = full_text.find(selected)
        if pos == -1:
            # Fallback: just use the selection with no anchor
            anchor = ""
            pattern = re.escape(selected)
        else:
            anchor, pattern = self.build_anchor_and_pattern(full_text, pos, selected, field_key)

        self.fields[field_key]["value"] = selected
        self.fields[field_key]["anchor"] = anchor
        self.fields[field_key]["pattern"] = pattern

        # Update label
        self.update_field_label(field_key)

    def build_anchor_and_pattern(self, full_text, value_pos, value, field_key):
        """
        Build a short anchor from text just BEFORE the value,
        and a regex pattern that uses that anchor as a hook.

        Strategy:
          - Look ~40 characters to the left of the value.
          - Use the last line before the value as anchor.
          - Trim to last ~25 characters so it's short.
          - Build pattern:  ANCHOR + r"\\s*([^\\n]+)"
        """
        window = 40
        start = max(0, value_pos - window)
        left_chunk = full_text[start:value_pos]

        # Use last line fragment as anchor
        line = left_chunk.splitlines()[-1] if left_chunk else ""
        line = line.strip()

        if len(line) > 25:
            line = line[-25:]

        anchor = line

        # If anchor is empty (value at start of line), just match the value itself
        if not anchor:
            pattern = re.escape(value)
        else:
            # Escape special regex chars in anchor, then capture up to newline
            pattern = re.escape(anchor) + r"\s*([^\n]+)"

        return anchor, pattern

    def update_field_label(self, field_key):
        info = self.fields[field_key]
        value = info["value"]
        anchor = info["anchor"]
        pattern = info["pattern"]

        text = f"Value: '{value}'"
        if anchor:
            text += f" | Anchor: '{anchor}'\nRegex: {pattern}"
        else:
            text += f"\nRegex (no anchor): {pattern}"

        self.field_labels[field_key].set(text)

    # --------------------------------------------------------
    # Generate parser code string
    # --------------------------------------------------------
    def generate_parser_code(self):
        vendor_name_for_file = self.vendor_name_entry.get().strip() or "vendor"
        # Safe module name
        safe_vendor = "".join(ch if ch.isalnum() else "_" for ch in vendor_name_for_file.lower())
        if not safe_vendor:
            safe_vendor = "vendor"

        vendor_title = vendor_name_for_file or "Vendor Name"

        # Build helper detect_* functions with patterns
        def pattern_for(field_key, default_comment):
            info = self.fields[field_key]
            if info["pattern"]:
                return f'    m = re.search(r"{info["pattern"]}", text, re.MULTILINE)\n' \
                       f'    if m:\n        return m.group(1).strip()\n' \
                       f'    return ""\n'
            else:
                return f'    # TODO: pattern not set via wizard for {field_key}\n' \
                       f'    # {default_comment}\n' \
                       f'    return ""\n'

        code = f'''# Auto-generated parser for {vendor_title}
# Generated by TemplateWizard

import fitz  # PyMuPDF
import re

def extract_pdf_text(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text("text")
        pages.append((i + 1, text))
    doc.close()
    return pages


def parse_invoice(pdf_path):
    pages = extract_pdf_text(pdf_path)
    results = []

    for page_num, text in pages:
        inv = {{
            "vendor": detect_vendor_name(text),
            "invoice_number": detect_invoice_number(text),
            "jobname": detect_jobname(text),
            "date": detect_invoice_date(text),
            "total": detect_invoice_total(text),
            "work_number": detect_work_number(text),
            "page": page_num,
        }}
        results.append(inv)

    return results[0] if len(results) == 1 else results


def detect_vendor_name(text):
'''
        # Vendor: for now, if user selected a specific literal vendor line,
        # we can just return that literal as default if we don't find a pattern.
        vend_info = self.fields["vendor"]
        if vend_info["pattern"]:
            code += f'    m = re.search(r"{vend_info["pattern"]}", text, re.MULTILINE)\n' \
                    f'    if m:\n        return m.group(1).strip()\n' \
                    f'    return "{vend_info["value"]}"\n\n'
        else:
            code += f'    # TODO: pattern not set via wizard for vendor\n' \
                    f'    # Returning literal from wizard as fallback.\n' \
                    f'    return "{vend_info["value"]}"\n\n'

        code += '''def detect_invoice_number(text):
'''
        code += pattern_for("invoice_num", "Example: look for 'Invoice # 12345'")

        code += '''def detect_invoice_date(text):
'''
        code += pattern_for("date", "Example: capture date near label like 'Date:'")

        code += '''def detect_invoice_total(text):
'''
        code += pattern_for("total", "Example: capture total near 'Total' or 'Amount Due'")

        code += '''def detect_jobname(text):
'''
        code += pattern_for("jobname", "Example: capture job description near 'Job' or 'Project'")

        code += '''def detect_work_number(text):
'''
        code += pattern_for("work_number", "Optional: work order number if present")

        return safe_vendor + "_parser.py", code

    # --------------------------------------------------------
    # Preview parser code in the UI
    # --------------------------------------------------------
    def preview_parser(self):
        filename, code = self.generate_parser_code()
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", f"# File: {filename}\n\n{code}")

    # --------------------------------------------------------
    # Test rules on current PDF
    # --------------------------------------------------------
    def test_rules(self):
        self.test_output.delete("1.0", tk.END)
        if not self.pdf_path:
            self.test_output.insert("1.0", "No PDF loaded.")
            return

        # Reuse the same logic as generated parser, but inline here
        try:
            pages = extract_pdf_text(self.pdf_path)
        except Exception as e:
            self.test_output.insert("1.0", f"Error reading PDF:\n{e}")
            return

        def apply_pattern(field_key, text):
            info = self.fields[field_key]
            pat = info["pattern"]
            if not pat:
                return ""
            try:
                m = re.search(pat, text, re.MULTILINE)
                if m:
                    return m.group(1).strip()
            except re.error as e:
                return f"[regex error: {e}]"
            return ""

        results = []
        for page_num, text in pages:
            inv = {
                "vendor":      apply_pattern("vendor", text) or self.fields["vendor"]["value"],
                "invoice_num": apply_pattern("invoice_num", text),
                "date":        apply_pattern("date", text),
                "total":       apply_pattern("total", text),
                "jobname":     apply_pattern("jobname", text),
                "work_number": apply_pattern("work_number", text),
                "page":        page_num,
            }
            results.append(inv)

        for inv in results:
            self.test_output.insert(
                tk.END,
                f"Page {inv['page']}:\n"
                f"  Vendor:      {inv['vendor']}\n"
                f"  Invoice #:   {inv['invoice_num']}\n"
                f"  Date:        {inv['date']}\n"
                f"  Total:       {inv['total']}\n"
                f"  Jobname:     {inv['jobname']}\n"
                f"  Work #:      {inv['work_number']}\n"
                f"{'-'*40}\n"
            )

    # --------------------------------------------------------
    # Save parser code to a .py file
    # --------------------------------------------------------
    def save_parser(self):
        filename, code = self.generate_parser_code()
        path = filedialog.asksaveasfilename(
            defaultextension=".py",
            initialfile=filename,
            filetypes=[("Python files", "*.py"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(code)
            messagebox.showinfo("Saved", f"Parser saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")


if __name__ == "__main__":
    app = TemplateWizard()
    app.mainloop()
