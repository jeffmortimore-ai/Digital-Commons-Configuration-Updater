import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# -------------------- Shared control flags --------------------
pause_flag = threading.Event()
pause_flag.set()

stop_flag = threading.Event()
stop_flag.clear()

start_time = None
processed_count = 0
total_records = 0


# -------------------- Logging --------------------
def log(output_box, text):
    output_box.insert(tk.END, text + "\n")
    output_box.see(tk.END)
    output_box.update()


# -------------------- Time + Progress Tracking --------------------
def update_status():
    if start_time is None or total_records == 0:
        return
    while not stop_flag.is_set():
        if pause_flag.is_set():
            elapsed = time.time() - start_time
            avg_per_record = elapsed / processed_count if processed_count > 0 else 0
            remaining = (total_records - processed_count) * avg_per_record if avg_per_record > 0 else 0
            elapsed_str = time.strftime("%H:%M:%S", time.gmtime(elapsed))
            remaining_str = time.strftime("%H:%M:%S", time.gmtime(remaining)) if remaining > 0 else "--:--:--"
            percent_complete = (processed_count / total_records) * 100 if total_records > 0 else 0
            timer_label.config(
                text=f"⏱️ Elapsed: {elapsed_str}   ⏳ Est. Remaining: {remaining_str}   ({processed_count}/{total_records})"
            )
            progress_bar["value"] = percent_complete
            percent_label.config(text=f"{percent_complete:.1f}%")
        time.sleep(1)


# -------------------- Core Form Updater Logic --------------------
def update_forms(file_path, output_box):
    """Update Digital Commons configuration pages based on spreadsheet."""
    global processed_count, total_records, start_time
    results = []
    start_time = time.time()

    try:
        # Connect to Chrome in debug mode
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        df = pd.read_excel(file_path)
        if "URL" not in df.columns:
            log(output_box, "⚠️ Spreadsheet must contain a 'URL' column.")
            return

        field_names = [col for col in df.columns if col not in ["URL", "Page Title"]]
        total_records = len(df)
        log(output_box, f"✅ Connected to Chrome. Processing {total_records} records...\n")

        # Start timer thread
        threading.Thread(target=update_status, daemon=True).start()

        for idx, row in df.iterrows():
            if stop_flag.is_set():
                log(output_box, "⏹️ Stop signal received. Halting process...")
                break

            while not pause_flag.is_set():
                log(output_box, "⏸️ Paused. Waiting to resume...")
                time.sleep(2)
                if stop_flag.is_set():
                    break
            if stop_flag.is_set():
                break

            url = row["URL"]
            log(output_box, f"➡️ ({idx+1}/{total_records}) Opening: {url}")
            context = "(unknown)"
            status = "Failed"
            message = ""

            try:
                driver.get(url)
                WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "form")))

                # Synchronize only listed fields
                for field in field_names:
                    value = str(row[field]) if pd.notna(row[field]) else ""
                    try:
                        element = driver.find_element(By.NAME, field)
                        tag = element.tag_name.lower()

                        if tag == "input":
                            input_type = element.get_attribute("type")
                            if input_type in ["text", "email", "number", "url", "search", "password"]:
                                element.clear()
                                if value:
                                    element.send_keys(value)
                            elif input_type in ["checkbox", "radio"]:
                                checked = element.is_selected()
                                should_check = bool(value)
                                if should_check != checked:
                                    element.click()
                            else:
                                driver.execute_script("arguments[0].value = arguments[1];", element, value)

                        elif tag == "textarea":
                            element.clear()
                            if value:
                                element.send_keys(value)

                        elif tag == "select":
                            # Select matching option or clear selection if blank
                            if value:
                                driver.execute_script(
                                    "for (let opt of arguments[0].options) { "
                                    "if (opt.text.trim() === arguments[1].trim()) { opt.selected = true; return; } }",
                                    element, value
                                )
                            else:
                                driver.execute_script("arguments[0].selectedIndex = -1;", element)

                    except Exception as e:
                        log(output_box, f"   ⚠️ Could not update field '{field}': {e}")

                # Submit form
                try:
                    driver.find_element(By.XPATH, "//input[@type='submit']").click()
                    log(output_box, "   💾 Submitted form.")
                except Exception:
                    log(output_box, "   ⚠️ Submit button not found.")

                WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(2)

                # Regenerate configuration
                try:
                    context = url.split("context=")[1].split("&")[0]
                    regen_url = f"https://digitalcommons.georgiasouthern.edu/cgi/user_config.cgi?context={context}&x_regenerate=1"
                    driver.get(regen_url)
                    log(output_box, f"   🔁 Regenerating context: {context}")
                    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(2)
                except Exception as e:
                    log(output_box, f"   ⚠️ Regeneration failed: {e}")
                    message += f"Regeneration issue: {e}. "

                status = "Success"
                message += "Updated and regenerated successfully."
                log(output_box, f"✅ {context} updated successfully.")

            except Exception as e:
                message = f"Error: {e}"
                log(output_box, f"❌ Error processing {url}: {e}")

            results.append({
                "URL": url,
                "Context": context,
                "Page Title": driver.title if driver.title else "",
                "Status": status,
                "Message": message
            })

            processed_count = idx + 1

        # Save results
        result_df = pd.DataFrame(results)
        out_path = os.path.splitext(file_path)[0] + "_results.xlsx"
        result_df.to_excel(out_path, index=False)
        log(output_box, f"\n📊 Results saved to: {out_path}")

        driver.quit()
        log(output_box, "\n🎉 Process finished or stopped.")

    except Exception as e:
        log(output_box, f"\n❌ Fatal error: {e}")


# -------------------- Control Buttons --------------------
def toggle_pause():
    if pause_flag.is_set():
        pause_flag.clear()
        pause_btn.config(text="▶️ Resume", bg="lightgreen")
        log(output_box, "⏸️ Script paused. Will pause after the current URL completes.")
    else:
        pause_flag.set()
        pause_btn.config(text="⏸️ Pause", bg="khaki")
        log(output_box, "▶️ Script resumed.")


def stop_script():
    stop_flag.set()
    pause_flag.set()
    stop_btn.config(state="disabled", bg="lightgray")
    log(output_box, "⏹️ Stop requested. Will finish current record then exit.")


def run_updater(os_type):
    global start_time, processed_count, total_records
    file_path = filedialog.askopenfilename(
        title="Select Updated Spreadsheet",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    stop_flag.clear()
    pause_flag.set()
    processed_count = 0
    total_records = 0
    progress_bar["value"] = 0
    percent_label.config(text="0.0%")
    timer_label.config(text="⏱️ Elapsed: 00:00:00   ⏳ Est. Remaining: --:--:--")
    stop_btn.config(state="normal", bg="lightcoral")
    pause_btn.config(text="⏸️ Pause", bg="khaki")

    log(output_box, f"\nStarting Form Updater for {os_type}...")
    threading.Thread(target=update_forms, args=(file_path, output_box), daemon=True).start()


# -------------------- Setup & Instructions (Scrollable, Copyable, Non-blocking) --------------------
def show_instructions():
    """Display setup and usage instructions in a scrollable, copyable, independent window."""
    instructions = """
🧩 Digital Commons Configuration Form Updater — Setup & Instructions
================================================================
This tool updates form field values from Digital Commons 
configuration pages. It connects to a Chrome browser that 
you open manually in debug mode.

------------------------------------
📦 1. Install Requirements
------------------------------------
Install Python (if not already installed):

🖥️ On macOS:
------------------------------------
1. Open Terminal and install Homebrew (if missing):
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

2. Install Python 3:
   brew install python

🪟 On Windows:
------------------------------------
1. Download Python from:
   https://www.python.org/downloads/windows/

2. During setup, check:
   ✅ "Add Python to PATH"

3. Click "Install Now"

------------------------------------
📚 2. Install Dependencies
------------------------------------
1. Once Python is installed, open Terminal (Mac) or Command Prompt (Windows)
   and run the following commands exactly as shown:

   python3 -m pip install --upgrade pip
   python3 -m pip install selenium webdriver-manager pandas openpyxl tk

   Windows users may need to use "python" instead of "python3".

------------------------------------
🌐 3. Install Google Chrome
------------------------------------
1. Download and install Chrome if you don’t already have it:
   https://www.google.com/chrome/

------------------------------------
⚙️ 4. Start Chrome in Debug Mode
------------------------------------
1. Close all Chrome windows completely.

2. Run ONE of the following commands exactly as shown:

🖥️ macOS:
------------------------------------
Using Terminal:
/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222 --user-data-dir="~/chrome-debug"

🪟 Windows:
------------------------------------
Using Command Prompt:
"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\chrome-debug"

OR 

Using PowerShell:
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome-debug"

Chrome will open in a new window.

------------------------------------
⚙️ 5. Log into Digital Commons
------------------------------------

1. Once Chrome is open in debug mode, navigate to Georgia Southern Commons:
   https://digitalcommons.georgiasouthern.edu/

2. Log into your Digital Commons administrator account.

------------------------------------
📘 6. Prepare the Spreadsheet
------------------------------------
• Each spreadsheet should include in cell A1:  URL
• Cell A2 and following should contain a complete Digital Commons URL, e.g.:
  https://digitalcommons.georgiasouthern.edu/cgi/user_config.cgi?context=allenepaulson
• Include only the columns for fields you want to update.  
• If a cell is empty, the corresponding form field will be cleared.  

------------------------------------
▶️ 7. Run the Script
------------------------------------
Once Chrome is open in debug mode and logged in:
   - Double-click this Python file, or run:
     python3 ConfigurationScraper_UI_Standalone.py
        (or "python" on Windows)
2. In the window:
   - Click “Show Instructions” if you need to review these steps.
   - Click “Run on Mac” or “Run on Windows.”
   - Select one or more Excel files.
3. The script will visit each URL and update all form fields included in the Excel file.

------------------------------------
📊 8. Results
------------------------------------
After processing, a new file will be created with “_results.xlsx” 
added to its name, containing a record of each update.

------------------------------------
🎉 That’s It!
------------------------------------
You can now use this tool anytime to batch update and 
regenerate Digital Commons configuration pages.
"""

    win = tk.Toplevel(root)
    win.title("Setup & Instructions")
    win.geometry("820x720")
    win.resizable(True, True)

    text_area = scrolledtext.ScrolledText(
        win,
        wrap=tk.WORD,
        width=95,
        height=40,
        font=("Helvetica", 11),
        bg="white",
        fg="black"
    )
    text_area.insert(tk.END, instructions)
    text_area.config(state="normal")  # Allow copying
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    close_btn = tk.Button(win, text="Close Instructions", command=win.destroy, width=20)
    close_btn.pack(pady=10)

    win.focus_force()  # Focus window but allow main window to remain interactive


# -------------------- Main Window --------------------
root = tk.Tk()
root.title("Digital Commons Configuration Form Updater")
root.geometry("830x740")

title_label = tk.Label(root, text="Digital Commons Configuration Form Updater", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

timer_label = tk.Label(root, text="⏱️ Elapsed: 00:00:00   ⏳ Est. Remaining: --:--:--", font=("Helvetica", 11))
timer_label.pack(pady=5)

progress_frame = tk.Frame(root)
progress_frame.pack(pady=5)

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=700, mode="determinate")
progress_bar.pack(side="left", padx=10)
percent_label = tk.Label(progress_frame, text="0.0%")
percent_label.pack(side="left")

desc_label = tk.Label(
    root,
    text="Updates Digital Commons configuration forms automatically using data from an Excel spreadsheet.",
    wraplength=780,
    justify="center"
)
desc_label.pack(pady=5)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

mac_btn = tk.Button(btn_frame, text="Run on Mac", width=18, command=lambda: run_updater("Mac"))
mac_btn.grid(row=0, column=0, padx=8)

win_btn = tk.Button(btn_frame, text="Run on Windows", width=18, command=lambda: run_updater("Windows"))
win_btn.grid(row=0, column=1, padx=8)

pause_btn = tk.Button(btn_frame, text="⏸️ Pause", width=15, bg="khaki", command=toggle_pause)
pause_btn.grid(row=0, column=2, padx=8)

stop_btn = tk.Button(btn_frame, text="⏹️ Stop", width=15, bg="lightcoral", command=stop_script)
stop_btn.grid(row=0, column=3, padx=8)

help_btn = tk.Button(root, text="Show Setup & Instructions", command=show_instructions, width=30)
help_btn.pack(pady=5)

output_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=25)
output_box.pack(padx=10, pady=10)

root.mainloop()
