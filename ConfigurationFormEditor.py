import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
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
def log(output_box, text, tag="info"):
    output_box.insert(tk.END, text + "\n", tag)
    output_box.see(tk.END)
    output_box.update()


# -------------------- Time + Progress Tracking --------------------
def update_status():
    """Updates elapsed time, estimated time, and progress bar."""
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
        timer_label.update()
        progress_bar.update()
        percent_label.update()


# -------------------- Core Form Editor Logic --------------------
def update_forms(file_path, output_box):
    global processed_count, total_records, start_time

    results = []
    start_time = time.time()

    try:
        # Connect to Chrome
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        df = pd.read_excel(file_path)
        if "URL" not in df.columns:
            log(output_box, "❌ Spreadsheet must contain a 'URL' column.", "error")
            return

        field_names = [col for col in df.columns if col not in ["URL", "Page Title"]]
        total_records = len(df)
        log(output_box, f"✅ Connected to Chrome. Processing {total_records} records...\n", "info")

        # Start status tracker thread
        threading.Thread(target=update_status, daemon=True).start()

        for idx, row in df.iterrows():
            if stop_flag.is_set():
                log(output_box, "⏹️ Stop signal received. Halting process...", "warn")
                break

            while not pause_flag.is_set():
                log(output_box, "⏸️ Paused. Waiting to resume...", "warn")
                time.sleep(2)
                if stop_flag.is_set():
                    break

            if stop_flag.is_set():
                break

            url = row["URL"]
            log(output_box, f"➡️ ({idx+1}/{total_records}) Opening: {url}", "info")
            context = "(unknown)"
            status = "Failed"
            message = ""

            try:
                driver.get(url)
                WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "form")))

                # Fill out fields
                for field in field_names:
                    value = str(row[field]) if pd.notna(row[field]) else ""
                    if not value:
                        continue
                    try:
                        element = driver.find_element(By.NAME, field)
                        tag = element.tag_name.lower()
                        if tag == "input":
                            input_type = element.get_attribute("type")
                            if input_type in ["text", "email", "number", "url", "search"]:
                                element.clear()
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
                            element.send_keys(value)
                        elif tag == "select":
                            driver.execute_script(
                                "for (let opt of arguments[0].options) { "
                                "if (opt.text.trim() === arguments[1].trim()) { opt.selected = true; break; } }",
                                element, value
                            )
                    except Exception as e:
                        log(output_box, f"   ⚠️ Could not update field '{field}': {e}", "warn")

                # Submit form
                try:
                    driver.find_element(By.XPATH, "//input[@type='submit']").click()
                    log(output_box, "   💾 Submitted form.", "info")
                except Exception:
                    log(output_box, "   ⚠️ Submit button not found.", "warn")

                WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(2)

                # Regenerate
                try:
                    context = url.split("context=")[1].split("&")[0]
                    regen_url = f"https://digitalcommons.georgiasouthern.edu/cgi/user_config.cgi?context={context}&x_regenerate=1"
                    driver.get(regen_url)
                    log(output_box, f"   🔁 Regenerating context: {context}", "info")
                    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(2)
                except Exception as e:
                    log(output_box, f"   ⚠️ Regeneration failed: {e}", "warn")
                    message += f"Regeneration issue: {e}. "

                status = "Success"
                message += "Updated and regenerated successfully."
                log(output_box, f"✅ {context} updated successfully.", "success")

            except Exception as e:
                message = f"Error: {e}"
                log(output_box, f"❌ Error processing {url}: {e}", "error")

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
        log(output_box, f"\n📊 Results saved to: {out_path}", "info")

        driver.quit()
        log(output_box, "\n🎉 Process finished or stopped.", "success")

    except Exception as e:
        log(output_box, f"\n❌ Fatal error: {e}", "error")


# -------------------- Control Buttons --------------------
def toggle_pause():
    if pause_flag.is_set():
        pause_flag.clear()
        pause_btn.config(text="▶️ Resume", bg="green")
        log(output_box, "⏸️ Script paused. Will pause after the current URL completes.", "warn")
    else:
        pause_flag.set()
        pause_btn.config(text="⏸️ Pause", bg="goldenrod")
        log(output_box, "▶️ Script resumed.", "info")


def stop_script():
    stop_flag.set()
    pause_flag.set()
    stop_btn.config(state="disabled", bg="gray")
    log(output_box, "⏹️ Stop requested. Will finish current record then exit.", "warn")


# -------------------- Run Handler --------------------
def run_editor(os_type):
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
    stop_btn.config(state="normal", bg="red")
    pause_btn.config(text="⏸️ Pause", bg="goldenrod")

    log(output_box, f"\nStarting Form Editor for {os_type}...", "info")
    threading.Thread(target=update_forms, args=(file_path, output_box), daemon=True).start()


# -------------------- Instructions --------------------
def show_instructions():
    instructions = """
🧩 Digital Commons Configuration Form Editor — with Time Tracker + Progress Bar
===============================================================================
Now includes:
⏱️ Elapsed and remaining time tracker
📊 Visual progress bar with % complete
⏸️ Pause / ▶️ Resume / ⏹️ Stop controls

Setup:
1. Install Python dependencies:
   python3 -m pip install selenium webdriver-manager pandas openpyxl

2. Start Chrome in debug mode:
   macOS:
   /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222 --user-data-dir="~/chrome-debug"

   Windows:
   "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\chrome-debug"

3. Log into Digital Commons in that Chrome window before starting.
"""
    win = tk.Toplevel(root)
    win.title("Setup & Instructions")
    win.geometry("900x700")
    text_area = scrolledtext.ScrolledText(win, wrap=tk.WORD, width=100, height=38)
    text_area.pack(expand=True, fill="both", padx=10, pady=10)
    text_area.insert(tk.END, instructions)
    text_area.config(state=tk.DISABLED)


# -------------------- UI Setup --------------------
root = tk.Tk()
root.title("Digital Commons Configuration Form Editor — Progress Bar Edition")
root.geometry("850x740")

title_label = tk.Label(
    root,
    text="Digital Commons Configuration Form Editor",
    font=("Helvetica", 16, "bold")
)
title_label.pack(pady=5)

timer_label = tk.Label(root, text="⏱️ Elapsed: 00:00:00   ⏳ Est. Remaining: --:--:--", fg="lightblue", bg="#1e1e1e", font=("Helvetica", 12))
timer_label.pack(pady=5)

# Progress bar
progress_frame = tk.Frame(root, bg="#1e1e1e")
progress_frame.pack(pady=5)

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=700, mode="determinate")
progress_bar.pack(side="left", padx=10)
percent_label = tk.Label(progress_frame, text="0.0%", fg="white", bg="#1e1e1e")
percent_label.pack(side="left")

desc_label = tk.Label(
    root,
    text="Automates configuration updates with time tracking, progress visualization, and full control.",
    wraplength=800,
    justify="center"
)
desc_label.pack(pady=5)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

mac_btn = tk.Button(btn_frame, text="Run on Mac", width=18, command=lambda: run_editor("Mac"))
mac_btn.grid(row=0, column=0, padx=8)

win_btn = tk.Button(btn_frame, text="Run on Windows", width=18, command=lambda: run_editor("Windows"))
win_btn.grid(row=0, column=1, padx=8)

pause_btn = tk.Button(btn_frame, text="⏸️ Pause", width=15, bg="goldenrod", command=toggle_pause)
pause_btn.grid(row=0, column=2, padx=8)

stop_btn = tk.Button(btn_frame, text="⏹️ Stop", width=15, bg="red", fg="white", command=stop_script)
stop_btn.grid(row=0, column=3, padx=8)

help_btn = tk.Button(root, text="Show Setup & Instructions", command=show_instructions, width=30)
help_btn.pack(pady=5)

output_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=25)
output_box.pack(padx=10, pady=10)

# Colors and theme
output_box.tag_config("info", foreground="white")
output_box.tag_config("success", foreground="limegreen")
output_box.tag_config("warn", foreground="gold")
output_box.tag_config("error", foreground="red")

root.configure(bg="#1e1e1e")
output_box.configure(bg="#2b2b2b", fg="white", insertbackground="white")

root.mainloop()
