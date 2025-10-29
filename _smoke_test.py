import os, sys, time
import fitz
import tkinter as tk
from pathlib import Path

sys.path.insert(0, r"C:\Users\jeffm\OneDrive\Documents\GitHub\GS-Commons-Configuration-Updater")
import pdf_processor as appmod

base = Path(r"C:\Users\jeffm\OneDrive\Documents\GitHub\GS-Commons-Configuration-Updater")
work = base / "_smoke"
work.mkdir(exist_ok=True)

# Create a 2-page PDF for testing
src_pdf = work / "sample.pdf"
if not src_pdf.exists():
    doc = fitz.open()
    for i in range(2):
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_text((72, 100), f"Page {i+1}")
    doc.save(src_pdf.as_posix())
    doc.close()

root = tk.Tk()
root.withdraw()
app = appmod.PDFProcessorApp(root)

# Configure options: remove first page only, copy mode
app.pdf_files = [src_pdf.as_posix()]
app.scan_ocr_var.set(True)  # scan text
app.ignore_first_page_scan_var.set(False)
app.ocr_pdfs_var.set(False)
app.remove_first_page_var.set(True)
app.save_mode.set("copy")
app.prefix_var.set("test_")
app.suffix_var.set("_out")
app.clean_char_var.set("None")

# Run processing synchronously
app.process_pdfs()

# Locate output
out_dir = src_pdf.parent / "processed_pdfs"
expected_name = app.generate_clean_filename(src_pdf.name)
out_path = out_dir / expected_name

# Verify output file exists and has 1 page
result = {
    "exists": out_path.exists(),
    "path": out_path.as_posix(),
    "pages": None,
}
if out_path.exists():
    with fitz.open(out_path.as_posix()) as d:
        result["pages"] = len(d)
print(result)
