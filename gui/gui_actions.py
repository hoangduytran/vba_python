# gui_actions.py
import os
import glob
import time
import threading
import shutil
import tkinter as tk
from tkinter import messagebox, filedialog
import logging
from multiprocessing import Pool

from gv import Gvar as gv
from mpp_logger import LOG_LEVELS, DynamicLevelFilter
from worker import worker_logging_setup
import worker

logger = None

def update_gui_filter():
    """
    Update the GUI handler filter based on the current log level and is_exact flag.
    If is_exact is True, only log records with exactly the chosen level are shown;
    if False, all records with levels greater than or equal to the chosen level are displayed.
    """
    current_level = LOG_LEVELS.get(gv.log_level_var.get(), logging.INFO)
    is_exact = gv.is_exact_var.get()
    print(f'is_exact:{is_exact}')

    # Clear existing DynamicLevelFilter filters.
    gv.root.gui_handler.filters = [f for f in gv.root.gui_handler.filters
                                   if not isinstance(f, DynamicLevelFilter)]
    gv.root.gui_handler.addFilter(DynamicLevelFilter(current_level, is_exact))
    reload_log_text()  # Make sure we reload

def select_log_level(selected):
    
    level = LOG_LEVELS.get(selected, logging.INFO)

    level_int = logging._nameToLevel[selected]
    print(f'select_log_level: {selected}, level:{level}')

    gv.mp_logging.select_log_level(level)
    update_gui_filter()
    reload_log_text()   # <<--- Add this line
    logger.info(f"Log level changed to {selected}")

def save_log():
    """
    Open a save file dialog to store the log.
    If the chosen extension is JSON, copy the temporary JSON file.
    Otherwise, write the log text from the log_text_widget.
    """
    path = filedialog.asksaveasfilename(
        title="Lưu Log vào tập tin",
        defaultextension=".log",
        filetypes=[("Tệp văn bản (*.log)", "*.log"), ("Tệp JSON (*.json)", "*.json")]
    )
    if path:
        try:
            if path.lower().endswith(".json"):
                shutil.copyfile(gv.root.mp_logging.log_temp_file_path, path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(gv.root.log_container.log_text.get("1.0", tk.END))
            messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
            logger.info("Log đã được lưu thành công")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")

def load_vba_file():
    """
    Open a file dialog to allow the user to select a VBA file.
    Save the chosen file path into the global state and log the selection.
    """
    init_dir = gv.root.excel_directory if gv.root.excel_directory else os.getcwd()
    path = filedialog.askopenfilename(
        title="Chọn tệp VBA",
        defaultextension=".bas",
        initialdir=init_dir,
        filetypes=[("Tệp VBA (*.bas)", "*.bas"), ("Tất cả các tệp", "*.*")]
    )
    if path:
        gv.root.vba_file = path
        globals()["global_vba_file_path"] = path
        logger.info(f"Đã tải tệp VBA: {path}")
        messagebox.showinfo("Thông báo", f"Đã tải tệp VBA: {path}")

def load_excel_directory():
    """
    Open a directory dialog to allow the user to select the directory containing Excel files.
    Log the number of Excel files found (if any).
    """
    init_dir = os.path.dirname(gv.root.vba_file) if gv.root.vba_file else os.getcwd()
    directory = filedialog.askdirectory(
        title="Chọn thư mục chứa các tệp Excel",
        initialdir=init_dir
    )
    if directory:
        gv.root.excel_directory = directory
        excel_files = glob.glob(os.path.join(directory, "*.xlsx"))
        if excel_files:
            logger.info(f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
            messagebox.showinfo("Thông báo", f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
            logger.info("Không tìm thấy tệp Excel nào trong thư mục đã chọn.")

def update_progress():
    """
    Update the progress bar and percentage label based on processed files.
    Poll the progress_queue and update the counter accordingly.
    Reschedules itself every 1000ms.
    """
    if not gv.root.running:
        return
    if gv.root.progress_queue:
        while not gv.root.progress_queue.empty():
            gv.root.progress_queue.get()
            gv.root.progress_count += 1
            gv.progress_bar["value"] = gv.root.progress_count
            percent = int((gv.root.progress_count / gv.root.total_files) * 100) if gv.root.total_files > 0 else 0
            gv.progress_label.config(text=f"{percent}%")
    gv.root.after_id_progress = gv.root.after(1000, update_progress)

def run_vba_on_all():
    """
    Process Excel files by:
      - Reading file paths from the UI.
      - Searching for Excel files in the chosen directory.
      - Batching the files and launching a multiprocessing Pool.
      - Updating progress and logging each step.
    """
    logger.info("Bắt đầu chạy VBA trên các tệp Excel.")
    dev_dir = os.environ.get('DEV') or os.getcwd()
    test_dir = os.path.join(dev_dir, 'test_files')
    if not gv.root.excel_directory:
        gv.root.excel_directory = os.path.join(test_dir, 'excel')
    if not gv.root.vba_file:
        gv.root.vba_file = os.path.join(gv.root.excel_directory, 'test_macro.bas')
    globals()["global_vba_file_path"] = gv.root.vba_file

    excel_files = glob.glob(os.path.join(gv.root.excel_directory, "*.xlsx"))
    if not excel_files:
        messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
        return

    gv.root.total_files = len(excel_files)
    gv.root.progress_count = 0
    gv.progress_bar["maximum"] = gv.root.total_files
    gv.progress_bar["value"] = 0

    num_processes = os.cpu_count() - 2 or 1
    batch_size = gv.root.total_files // num_processes
    remainder = gv.root.total_files % num_processes
    batches = []
    start = 0
    for i in range(num_processes):
        extra = 1 if i < remainder else 0
        end = start + batch_size + extra
        batches.append(excel_files[start:end])
        start = end

    logger.info(f"Bắt đầu chạy VBA trên {gv.root.total_files} tệp, chia thành {num_processes} batch")

    if gv.root.mp_logging.queue is None:
        raise ValueError("Hàng đợi logging chia sẻ chưa được thiết lập!")
    from mpp_logger import get_mp_logger
    gv.root.progress_queue = get_mp_logger().manager.Queue()
    shared_queue = gv.root.mp_logging.queue

    pool = Pool(
        processes=num_processes,
        initializer=worker_logging_setup,
        initargs=(shared_queue, gv.root.mp_logging.log_level.value)
    )
    for batch in batches:
        pool.apply_async(worker.process_batch, args=(batch, gv.root.progress_queue, gv.root.mp_logging.queue))
    pool.close()
    pool.join()

    time.sleep(1)
    logger.info(f"Đã chạy VBA trên {gv.root.total_files} tệp Excel.")
    reload_log_text()

def run_vba_on_all_thread():
    """
    Start the VBA processing (run_vba_on_all) in a separate thread so the UI remains responsive.
    """
    gv.root.stop_event.clear()
    gv.root.vba_thread = threading.Thread(target=run_vba_on_all)
    gv.root.vba_thread.start()

def copy_text():
    """
    Copy the selected text from the log_text_widget to the clipboard.
    """
    try:
        selected_text = gv.root.log_container.log_text.get("sel.first", "sel.last")
        gv.root.clipboard_clear()
        gv.root.clipboard_append(selected_text)
    except tk.TclError:
        pass

def paste_text():
    """
    Paste text from the clipboard into the log_text_widget at the current cursor position.
    """
    try:
        clipboard_text = gv.root.clipboard_get()
        gv.root.log_container.log_text.insert(tk.INSERT, clipboard_text)
    except tk.TclError:
        pass

def select_fonts():
    """
    Allow the user to select a font for the log_text_widget.
    Opens a top-level window with font options.
    """
    top = tk.Toplevel(gv.root)
    top.title("Chọn phông chữ")
    tk.Label(top, text="Phông chữ:").pack(side="left", padx=5, pady=5)
    font_options = ["Arial", "Courier New", "Times New Roman", "Verdana", "Tahoma"]
    var = tk.StringVar(value=gv.root.font.actual("family"))
    option_menu = tk.OptionMenu(top, var, *font_options)
    option_menu.pack(side="left", padx=5, pady=5)
    def update_font():
        gv.root.font.configure(family=var.get())
        top.destroy()
    tk.Button(top, text="OK", command=update_font).pack(side="left", padx=5, pady=5)

def font_size_up():
    """
    Increase the font size of the log_text_widget.
    """
    current_size = gv.root.font.actual("size")
    gv.root.font.configure(size=current_size + 2)

def font_size_down():
    """
    Decrease the font size of the log_text_widget (not going below 1).
    """
    current_size = gv.root.font.actual("size")
    new_size = current_size - 2 if current_size > 2 else 1
    gv.root.font.configure(size=new_size)

def toggle_wrap():
    """
    Toggle the wrap setting of the log_text_widget between 'word' and 'none'.
    """
    current_wrap = gv.root.log_container.log_text.cget("wrap")
    new_wrap = "none" if current_wrap == "word" else "word"
    gv.root.log_container.log_text.configure(wrap=new_wrap)


def exit_app():
    """
    Cleanly shut down the application:
      - Stop background threads.
      - Shutdown logging.
      - Cancel scheduled UI callbacks.
      - Destroy the main window.
    """
    gv.root.running = False
    if hasattr(gv.root, 'vba_thread') and gv.root.vba_thread is not None and gv.root.vba_thread.is_alive():
        gv.root.vba_thread.join(timeout=5)
    gv.root.mp_logging.shutdown()
    if gv.root.after_id_progress is not None:
        gv.root.after_cancel(gv.root.after_id_progress)
    gv.root.destroy()


def reload_log_text():
    """
    Clear the log display and reload all log entries from mp_logging.log_store
    that satisfy the filter, then convert each record to text using the PrettyFormatter.
    """
    # Get the text widget (assumed to be a Tkinter Text widget)
    widget = gv.root.log_container.log_text
    widget.configure(state="normal")
    widget.delete("1.0", tk.END)
    widget.configure(state="disabled")
    
    # Retrieve the current log level and exact matching flag from the GUI's shared variables.
    current_level = LOG_LEVELS.get(gv.log_level_var.get(), logging.INFO)
    is_exact = gv.is_exact_var.get()
    
    # Define a filter function for the raw log record.
    def passes_filter(rec):
        # Compare using the raw record's 'levelno' attribute.
        return rec.levelno == current_level if is_exact else rec.levelno >= current_level

    # Filter the raw log records stored in the MemoryLogHandler.
    filtered_records = [rec for rec in gv.root.mp_logging.log_store if passes_filter(rec)]
    
    # Format each record using the PrettyFormatter.
    # The PrettyFormatter uses create_log_record with diacritics=True.
    formatted_logs = [
        gv.root.mp_logging.pretty_formatter.format(rec)
        for rec in filtered_records
    ]
    
    # Join the formatted log entries with newline characters.
    full_text = "\n".join(formatted_logs)
    
    # Insert the combined log text into the log container.
    gv.root.log_container.insert_log(full_text)



# Register all callback functions in an action_list dictionary.
action_list = {
    "update_gui_filter": update_gui_filter,
    "select_log_level": select_log_level,
    "save_log": save_log,
    "load_vba_file": load_vba_file,
    "load_excel_directory": load_excel_directory,
    "update_progress": update_progress,
    "run_macro_thread": run_vba_on_all_thread,
    "run_macro": run_vba_on_all,
    "copy_text": copy_text,
    "paste_text": paste_text,
    "select_fonts": select_fonts,
    "font_size_up": font_size_up,
    "font_size_down": font_size_down,
    "toggle_wrap": toggle_wrap,
    "reload_log_text": reload_log_text,
    "exit_app": exit_app,
}
