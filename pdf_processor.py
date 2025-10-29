import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import io
import threading
import time
import datetime
import re
import uuid
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import openpyxl

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced PDF Processor")
        self.root.geometry("800x850")

        # --- State Management ---
        self.processing_thread = None
        self.is_paused = threading.Event()  # set => paused
        self.is_stopped = threading.Event()  # set => stop
        self.start_time = 0
        self.elapsed_seconds = 0
        self.timer_id = None
        self.completed_count = 0
        self.total_files = 0

        # OCR settings
        self.ocr_dpi_var = tk.IntVar(value=200)
        self.ocr_lang_var = tk.StringVar(value="eng")

        self.setup_gui()
        self.reset_job() # Initialize UI state

    def setup_gui(self):
        # --- Style ---
        style = ttk.Style(self.root)
        style.theme_use('clam')

        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1) # Log area row

        # --- Top Action Buttons ---
        top_button_frame = ttk.Frame(main_frame)
        top_button_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(top_button_frame, text="Setup & Instructions", command=self.show_instructions).pack(side=tk.LEFT)

        # --- 1. File Selection ---
        file_frame = ttk.LabelFrame(main_frame, text="1. Select Files", padding="10")
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(file_frame, selectmode=tk.EXTENDED, height=6)
        self.file_listbox.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
        ttk.Button(file_frame, text="Select PDF Files", command=self.select_files).grid(row=1, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(file_frame, text="Select Folder", command=self.select_folder).grid(row=1, column=1, sticky="ew", padx=(5, 0))

        # --- 2. Options ---
        options_frame = ttk.LabelFrame(main_frame, text="2. Choose Operations", padding="10")
        options_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)

        # Scan & Process Options
        scan_frame = ttk.LabelFrame(options_frame, text="Scan PDFs", padding="10")
        scan_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.scan_ocr_var = tk.BooleanVar()
        self.ignore_first_page_scan_var = tk.BooleanVar()
        ttk.Checkbutton(scan_frame, text="Scan for OCR", variable=self.scan_ocr_var).pack(anchor="w")
        ttk.Checkbutton(scan_frame, text="Ignore first page during scan", variable=self.ignore_first_page_scan_var).pack(anchor="w")
        
        process_frame = ttk.LabelFrame(options_frame, text="Process PDFs", padding="10")
        process_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.ocr_pdfs_var = tk.BooleanVar()
        self.remove_first_page_var = tk.BooleanVar()
        ttk.Checkbutton(process_frame, text="OCR PDFs", variable=self.ocr_pdfs_var).pack(anchor="w")
        ttk.Checkbutton(process_frame, text="Remove first page", variable=self.remove_first_page_var).pack(anchor="w")

        # OCR Settings
        ocr_settings = ttk.Frame(process_frame)
        ocr_settings.pack(anchor="w", fill="x", pady=(8, 0))
        ttk.Label(ocr_settings, text="OCR DPI:").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(ocr_settings, from_=100, to=600, increment=25, textvariable=self.ocr_dpi_var, width=6).grid(row=0, column=1, sticky="w", padx=(5, 15))
        ttk.Label(ocr_settings, text="Tesseract Lang:").grid(row=0, column=2, sticky="w")
        ttk.Entry(ocr_settings, textvariable=self.ocr_lang_var, width=10).grid(row=0, column=3, sticky="w", padx=(5, 0))

        # --- 3. Output Options ---
        output_frame = ttk.LabelFrame(main_frame, text="3. Output & Naming", padding="10")
        output_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        
        self.save_mode = tk.StringVar(value="copy")
        ttk.Radiobutton(output_frame, text="Create a new copy", variable=self.save_mode, value="copy").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(output_frame, text="Overwrite existing file", variable=self.save_mode, value="overwrite").grid(row=0, column=1, sticky="w")

        ttk.Label(output_frame, text="Prefix:").grid(row=1, column=0, sticky="w", pady=5)
        self.prefix_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.prefix_var).grid(row=1, column=1, sticky="ew")

        ttk.Label(output_frame, text="Suffix:").grid(row=2, column=0, sticky="w")
        self.suffix_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.suffix_var).grid(row=2, column=1, sticky="ew")
        
        ttk.Label(output_frame, text="Clean Filename:").grid(row=3, column=0, sticky="w", pady=5)
        self.clean_char_var = tk.StringVar()
        ttk.Combobox(output_frame, textvariable=self.clean_char_var, values=["None", "Replace spaces/special chars with _", "Replace spaces/special chars with -"], state="readonly").grid(row=3, column=1, sticky="ew")
        self.clean_char_var.set("None")

        # --- 4. Live Log ---
        log_frame = ttk.LabelFrame(main_frame, text="4. Live Log", padding="10")
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 10))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', height=10)
        self.log_area.grid(row=0, column=0, sticky="nsew")

        # --- 5. Progress and Controls ---
        progress_frame = ttk.LabelFrame(main_frame, text="5. Progress & Controls", padding="10")
        progress_frame.grid(row=5, column=0, sticky="ew")
        progress_frame.columnconfigure(0, weight=1)
        
        # Counters and Timers
        stats_frame = ttk.Frame(progress_frame)
        stats_frame.grid(row=0, column=0, columnspan=5, sticky="ew", pady=5)
        self.completed_label = ttk.Label(stats_frame, text="Completed: 0")
        self.completed_label.pack(side=tk.LEFT, padx=5)
        self.remaining_label = ttk.Label(stats_frame, text="Remaining: 0")
        self.remaining_label.pack(side=tk.LEFT, padx=5)
        self.elapsed_label = ttk.Label(stats_frame, text="Elapsed: 00:00:00")
        self.elapsed_label.pack(side=tk.RIGHT, padx=5)
        self.etr_label = ttk.Label(stats_frame, text="ETR: --:--:--")
        self.etr_label.pack(side=tk.RIGHT, padx=5)
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=1, column=0, columnspan=5, sticky="ew", pady=5)

        # Control buttons
        control_frame = ttk.Frame(progress_frame)
        control_frame.grid(row=2, column=0, columnspan=5, pady=(10, 0))
        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_processing_thread)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.pause_resume_button = ttk.Button(control_frame, text="Pause", command=self.toggle_pause_resume)
        self.pause_resume_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_processing)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.reset_button = ttk.Button(control_frame, text="Reset / New Job", command=self.reset_job)
        self.reset_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
    
    # --- GUI Actions ---
    def select_files(self):
        files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])
        if files: self.pdf_files.extend(files); self.update_listbox()

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self.pdf_files.extend([os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")])
            self.update_listbox()

    def update_listbox(self):
        self.pdf_files = sorted(list(set(self.pdf_files)))
        self.file_listbox.delete(0, tk.END)
        for f in self.pdf_files: self.file_listbox.insert(tk.END, os.path.basename(f))
        self.remaining_label.config(text=f"Remaining: {len(self.pdf_files)}")

    def log_message(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.config(state='disabled')
        self.log_area.see(tk.END)

    def show_instructions(self):
        win = tk.Toplevel(self.root)
        win.title("Setup and Instructions")
        win.geometry("700x600")
        
        text_area = scrolledtext.ScrolledText(win, wrap=tk.WORD, padx=10, pady=10)
        text_area.pack(expand=True, fill='both')
        text_area.insert(tk.INSERT, self.get_instructions_text())
        text_area.config(state='disabled')

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=10)
        win.transient(self.root)
        win.grab_set()

    # --- Processing Control ---
    def reset_job(self):
        if self.processing_thread and self.processing_thread.is_alive(): self.stop_processing()
        
        self.pdf_files = []
        self.file_listbox.delete(0, tk.END)
        self.log_message("System reset. Ready for a new job.")
        
        self.progress_bar['value'] = 0
        self.completed_label.config(text="Completed: 0")
        self.remaining_label.config(text="Remaining: 0")
        self.elapsed_label.config(text="Elapsed: 00:00:00")
        self.etr_label.config(text="ETR: --:--:--")

        self.start_button.config(state="normal")
        self.pause_resume_button.config(text="Pause", state="disabled")
        self.stop_button.config(state="disabled")
        self.reset_button.config(state="normal")
        
        self.is_paused.clear()
        self.is_stopped.clear()
        if self.timer_id: self.root.after_cancel(self.timer_id); self.timer_id = None

    def start_processing_thread(self):
        if not self.pdf_files:
            messagebox.showwarning("No Files", "Please select PDF files or a folder first.")
            return
        if not (self.scan_ocr_var.get() or self.ocr_pdfs_var.get() or self.remove_first_page_var.get()):
            messagebox.showwarning("No Operations", "Please select at least one operation to perform.")
            return
        
        self.start_button.config(state="disabled")
        self.pause_resume_button.config(state="normal")
        self.stop_button.config(state="normal")
        
        self.is_stopped.clear()
        self.is_paused.clear()
        
        self.processing_thread = threading.Thread(target=self.process_pdfs)
        self.processing_thread.daemon = True
        self.processing_thread.start()
        
        self.start_time = time.time()
        self.update_timers()

    def toggle_pause_resume(self):
        if self.is_paused.is_set():
            self.is_paused.clear()
            self.pause_resume_button.config(text="Pause")
            self.log_message(">>> Resuming processing...")
            self.start_time = time.time() - self.elapsed_seconds # Adjust start time
            self.update_timers() # Restart timer loop
        else:
            self.is_paused.set()
            self.pause_resume_button.config(text="Resume")
            self.log_message(">>> Pausing processing...")
            if self.timer_id: self.root.after_cancel(self.timer_id); self.timer_id = None

    def stop_processing(self):
        if self.processing_thread and self.processing_thread.is_alive():
            if messagebox.askyesno("Confirm Stop", "Are you sure you want to stop the current job?"):
                self.is_stopped.set()
                self.is_paused.clear() # Allow thread to see the stop signal
                self.log_message(">>> STOP signal sent. Finishing current file...")

    # --- Core Logic ---
    def process_pdfs(self):
        total_files = len(self.pdf_files)
        self.log_data = []
        self.completed_count = 0
        self.total_files = total_files
        
        # --- Tesseract Path Check ---
        try:
            pytesseract.get_tesseract_version()
        except pytesseract.TesseractNotFoundError:
            self.root.after(0, lambda: self.log_message("[ERROR] Tesseract not found. Please check installation and PATH."))
            self.root.after(0, lambda: messagebox.showerror("Tesseract Error", "Tesseract is not installed or it's not in your system's PATH. Please fix this via the instructions."))
            self.root.after(0, self.job_finished, "Error")
            return

        for i, file_path in enumerate(self.pdf_files):
            if self.is_stopped.is_set():
                self.root.after(0, lambda: self.log_message("--- JOB STOPPED BY USER ---"))
                break
            
            # Pause gate: block while paused; abort if stopped
            if not self._wait_if_paused_or_stopped():
                break
            
            filename = os.path.basename(file_path)
            self.root.after(0, lambda f=filename: self.log_message(f"--- Starting: {f} ---"))
            
            log_entry = {
                "Original Filename": filename,
                "Scanned for OCR": "No", "First Page Ignored During Scan": "N/A", "Pages Lacking OCR": 0,
                "File OCR'd": "No", "First Page Removed": "No", "New Filename": filename
            }

            try:
                # --- Scan Phase ---
                if self.scan_ocr_var.get():
                    log_entry["Scanned for OCR"] = "Yes"
                    log_entry["First Page Ignored During Scan"] = "Yes" if self.ignore_first_page_scan_var.get() else "No"
                    try:
                        with fitz.open(file_path) as doc:
                            start_page = 1 if self.ignore_first_page_scan_var.get() and len(doc) > 1 else 0
                            for page_num in range(start_page, len(doc)):
                                if not doc[page_num].get_text("text"):
                                    log_entry["Pages Lacking OCR"] += 1
                        self.root.after(0, lambda p=log_entry['Pages Lacking OCR']: self.log_message(f"Scan result: {p} page(s) lack text."))
                    except Exception as scan_err:
                        self.root.after(0, lambda e=scan_err: self.log_message(f"[WARN] Scan failed: {e}"))

                # --- Process Phase ---
                if self.ocr_pdfs_var.get() or self.remove_first_page_var.get():
                    temp_path = self._make_temp_path(file_path)
                    wrote_temp = False

                    try:
                        with fitz.open(file_path) as doc:
                            # Remove first page if requested
                            if self.remove_first_page_var.get():
                                if len(doc) > 1:
                                    self.root.after(0, lambda: self.log_message("Removing first page..."))
                                    doc.delete_page(0)
                                    log_entry["First Page Removed"] = "Yes"
                                else:
                                    self.root.after(0, lambda: self.log_message("Skipping first page removal (only 1 page)."))

                            if self.ocr_pdfs_var.get():
                                self.root.after(0, lambda: self.log_message("Performing OCR... (this may take a while)"))
                                out_doc = fitz.open()
                                dpi = max(72, int(self.ocr_dpi_var.get()))
                                scale = dpi / 72.0
                                mat = fitz.Matrix(scale, scale)
                                for pg_index in range(len(doc)):
                                    if not self._wait_if_paused_or_stopped():
                                        raise RuntimeError("Processing stopped by user")
                                    page = doc[pg_index]
                                    pix = page.get_pixmap(matrix=mat, alpha=False)
                                    png_bytes = pix.tobytes("png")
                                    pil_img = Image.open(io.BytesIO(png_bytes))
                                    pdf_bytes = pytesseract.image_to_pdf_or_hocr(pil_img, extension='pdf', lang=self.ocr_lang_var.get())
                                    one_page_pdf = fitz.open("pdf", pdf_bytes)
                                    out_doc.insert_pdf(one_page_pdf)
                                out_doc.save(temp_path)
                                out_doc.close()
                                log_entry["File OCR'd"] = "Yes"
                                wrote_temp = True
                            else:
                                # Only removal requested; save modified document
                                doc.save(temp_path)
                                wrote_temp = True

                        if wrote_temp:
                            final_filename = self.generate_clean_filename(filename)
                            log_entry["New Filename"] = final_filename

                            if self.save_mode.get() == "copy":
                                output_folder = os.path.join(os.path.dirname(file_path), "processed_pdfs")
                                os.makedirs(output_folder, exist_ok=True)
                                output_path = os.path.join(output_folder, final_filename)
                                output_path = self._resolve_collision(output_path)
                            else:
                                output_path = os.path.join(os.path.dirname(file_path), final_filename)

                            os.replace(temp_path, output_path)
                            self.root.after(0, lambda p=output_path: self.log_message(f"Saved to: {p}"))
                    finally:
                        try:
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                        except Exception:
                            pass

            except Exception as e:
                self.root.after(0, lambda f=filename, err=e: self.log_message(f"[ERROR] processing {f}: {err}"))
            
            self.completed_count += 1
            self.root.after(0, self.update_progress, self.completed_count, total_files)
            self.log_data.append(log_entry)

        # --- Finalize ---
        final_status = "Stopped" if self.is_stopped.is_set() else "Completed"
        self.root.after(0, self.job_finished, final_status)
    
    def job_finished(self, status):
        if status == "Completed":
            self.log_message(f"--- JOB COMPLETED ---")
            messagebox.showinfo("Success", "All tasks are complete. You can now save the log file.")
            self.create_and_save_report(self.log_data)
        elif status == "Error":
            self.log_message(f"--- JOB HALTED DUE TO ERROR ---")
        else: # Stopped
            self.log_message(f"--- JOB FINISHED ---")

        self.start_button.config(state="disabled")
        self.pause_resume_button.config(text="Pause", state="disabled")
        self.stop_button.config(state="disabled")
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
            
    # --- Helper & UI Update Functions ---
    def update_progress(self, completed, total):
        self.progress_bar['value'] = (completed / total) * 100 if total else 0
        self.completed_label.config(text=f"Completed: {completed}")
        self.remaining_label.config(text=f"Remaining: {max(0, total - completed)}")

    def update_timers(self):
        if self.is_paused.is_set() or self.is_stopped.is_set():
            return
        
        self.elapsed_seconds = time.time() - self.start_time
        self.elapsed_label.config(text=f"Elapsed: {str(datetime.timedelta(seconds=int(self.elapsed_seconds)))}")

        completed = self.completed_count
        total = self.total_files if self.total_files else len(self.pdf_files)
        if completed > 0 and total:
            time_per_file = self.elapsed_seconds / completed
            remaining_files = max(0, total - completed)
            etr_seconds = time_per_file * remaining_files
            self.etr_label.config(text=f"ETR: {str(datetime.timedelta(seconds=int(etr_seconds)))}")
        
        self.timer_id = self.root.after(1000, self.update_timers)

    def _wait_if_paused_or_stopped(self):
        while self.is_paused.is_set():
            if self.is_stopped.is_set():
                return False
            time.sleep(0.1)
        return not self.is_stopped.is_set()

    def _make_temp_path(self, file_path):
        base_dir = os.path.dirname(file_path)
        return os.path.join(base_dir, f".tmp_{uuid.uuid4().hex}.pdf")

    def _resolve_collision(self, output_path):
        if not os.path.exists(output_path):
            return output_path
        base, ext = os.path.splitext(output_path)
        idx = 1
        while True:
            candidate = f"{base} ({idx}){ext}"
            if not os.path.exists(candidate):
                return candidate
            idx += 1

    def generate_clean_filename(self, filename):
        name, ext = os.path.splitext(filename)
        new_name = self.prefix_var.get() + name + self.suffix_var.get()
        
        clean_mode = self.clean_char_var.get()
        if clean_mode == "Replace spaces/special chars with _":
            new_name = re.sub(r'[\s\W]+', '_', new_name).strip('_')
        elif clean_mode == "Replace spaces/special chars with -":
            new_name = re.sub(r'[\s\W]+', '-', new_name).strip('-')
            
        return new_name + ext

    def create_and_save_report(self, data):
        if not data: return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            title="Save Processing Log",
            initialfile=f"pdf_processing_log_{datetime.datetime.now():%Y-%m-%d}.xlsx"
        )
        if not filepath: return

        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Processing Log"
            
            headers = list(data[0].keys())
            sheet.append(headers)
            
            for row_data in data:
                sheet.append(list(row_data.values()))
            
            workbook.save(filepath)
            self.log_message(f"Log file saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save log file:\n{e}")

    def get_instructions_text(self):
        return """
WHAT THIS SCRIPT DOES:
This application provides a graphical interface to perform batch operations on PDF files. You can select multiple files or an entire folder.

Key Features:
- Scan PDFs to check if they contain searchable text (i.e., if they are image-only).
- Process PDFs by performing Optical Character Recognition (OCR) to make them text-searchable.
- Remove the first page from PDFs.
- Flexible output options: overwrite files or save modified versions as new copies.
- Advanced file naming: add prefixes/suffixes and clean filenames by replacing special characters.
- Live progress tracking with pause, resume, and stop controls.
- Generates a detailed Excel log file (.xlsx) of all operations performed.

----------------------------------------------------
SETUP FOR WINDOWS:
----------------------------------------------------
1. Install Python:
   - Go to python.org and download the latest version of Python.
   - During installation, MAKE SURE to check the box that says "Add Python to PATH".

2. Install Tesseract-OCR:
   - Download the installer from the official Tesseract project: https://github.com/UB-Mannheim/tesseract/wiki
   - Run the installer. Note the installation path (e.g., C:\\Program Files\\Tesseract-OCR).
   - The script will try to find Tesseract automatically. If it fails, you may need to add this path to your system's PATH environment variable.

3. Install Required Python Libraries:
   - Open Command Prompt (cmd) or PowerShell.
   - Run the following command:
     pip install PyMuPDF pytesseract pillow openpyxl

----------------------------------------------------
SETUP FOR MACOS:
----------------------------------------------------
1. Install Homebrew (if you don't have it):
   - Open the Terminal app.
   - Paste and run the command from the official Homebrew website: https://brew.sh

2. Install Python and Tesseract:
   - With Homebrew installed, run the following commands in your Terminal:
     brew install python
     brew install tesseract

3. Install Required Python Libraries:
   - In the Terminal, run this command (use 'pip3' if 'pip' is linked to an older Python version):
     pip3 install PyMuPDF pytesseract pillow openpyxl

----------------------------------------------------
HOW TO RUN THE SCRIPT:
----------------------------------------------------
1. Save the code as a Python file (e.g., pdf_processor_advanced.py).
2. Open your terminal or command prompt.
3. Navigate to the directory where you saved the file.
   - Example: cd Downloads
4. Run the script using Python:
   - On Windows: python pdf_processor_advanced.py
   - On macOS: python3 pdf_processor_advanced.py
5. The application window will appear. Follow the steps in the GUI.
"""

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()
