# gui.py
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox, filedialog
import logging
import glob, os, threading
from multiprocessing import Pool
import worker
from worker import worker_logging_setup
from mpp_logger import get_mp_logger
from logtext import LogText

# Declare global logger; it will be assigned later.
logger = None

# Common widget style for all widgets.
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}

# Log-level options.
LOG_LEVELS = {
    "NO_LOGGING": 100,   # Higher than CRITICAL, so nothing shows in GUI.
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}

# A simple TextHandler for the Tkinter Text widget.
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.is_gui_handler = True  # mark this handler as our GUI handler

    def emit(self, record):
        try:
            msg = self.format(record) + "\n"
            self.text_widget.after(0, self.append, msg)
        except Exception:
            self.handleError(record)

    def append(self, msg):
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg)
        self.text_widget.configure(state="disabled")
        self.text_widget.yview(tk.END)

class MainWindow(tk.Tk):
    def __init__(self, mp_logging):
        super().__init__()
        self.mp_logging = mp_logging

        # Assign the global logger from mp_logging.
        global logger
        logger = self.mp_logging.logger

        # Example log entry.
        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")

        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Variables for tracking progress.
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None

        self.stop_event = threading.Event()

        # ----------------------
        # Left area: Taskbar
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Replace debug checkbox with a dropdown for log level selection.
        self.log_level_var = tk.StringVar(value="INFO")
        self.log_level_menu = ttk.OptionMenu(self.taskbar, self.log_level_var, "INFO", *LOG_LEVELS.keys(), command=self.select_log_level)
        self.log_level_menu.config(width=20)
        self.log_level_menu.pack(pady=5, anchor="w")

        self.create_taskbar_buttons()

        # ----------------------
        # Right area: Main display (Log text & progress)
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        self.log_container = LogText(right_area)
        self.log_container.pack(fill="both", expand=True)

        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))

        self.after_id_progress = self.after(500, self.update_progress)

        # Attach a GUI TextHandler.
        if self.mp_logging.listener is not None:
            from mpp_logger import get_mp_logger  # if needed
            gui_handler = TextHandler(self.log_container.log_text)
            gui_handler.setFormatter(self.mp_logging.json_formatter)
            current_level = LOG_LEVELS.get(self.log_level_var.get(), logging.INFO)
            gui_handler.addFilter(lambda record: record.levelno >= current_level)
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (gui_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def create_taskbar_buttons(self):
        buttons_config = [
            {"text": "Lưu Log vào tập tin", "command": self.save_log},
            {"text": "Tải tệp VBA", "command": self.load_vba_file},
            {"text": "Tải thư mục Excel", "command": self.load_excel_directory},
            {"text": "Chạy VBA trên tất cả các tệp Excel", "command": self.run_vba_on_all_thread},
            {"text": "Thoát Ứng dụng", "command": self.exit_app}
        ]
        for btn_conf in buttons_config:
            btn = tk.Button(self.taskbar, text=btn_conf["text"], command=btn_conf["command"], **COMMON_WIDGET_STYLE)
            btn.pack(pady=3, fill="x", anchor="w")

    def select_log_level(self, selected):
        level = LOG_LEVELS.get(selected, logging.INFO)
        self.mp_logging.select_log_level(level)
        logger.info(f"Log level changed to {selected}")

    def save_log(self):
        path = filedialog.asksaveasfilename(title="Lưu Log vào tập tin", defaultextension=".txt",
                                              filetypes=[("Tệp văn bản", "*.txt"), ("Tất cả các tệp", "*.*")])
        if path:
            try:
                with open(self.mp_logging.log_temp_file_path, "r", encoding="utf-8") as src, \
                     open(path, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
                messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
                logger.info("Log đã được lưu thành công")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")

    def load_vba_file(self):
        init_dir = self.excel_directory if self.excel_directory else os.getcwd()
        path = filedialog.askopenfilename(title="Chọn tệp VBA", defaultextension=".bas", initialdir=init_dir,
                                          filetypes=[("Tệp VBA", "*.bas"), ("Tất cả các tệp", "*.*")])
        if path:
            self.vba_file = path
            globals()["global_vba_file_path"] = path
            logger.info(f"Đã tải tệp VBA: {path}")
            messagebox.showinfo("Thông báo", f"Đã tải tệp VBA: {path}")

    def load_excel_directory(self):
        init_dir = os.path.dirname(self.vba_file) if self.vba_file else os.getcwd()
        directory = filedialog.askdirectory(title="Chọn thư mục chứa các tệp Excel", initialdir=init_dir)
        if directory:
            self.excel_directory = directory
            excel_files = glob.glob(os.path.join(directory, "*.xlsx"))
            if excel_files:
                logger.info(f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
                messagebox.showinfo("Thông báo", f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
                logger.info("Không tìm thấy tệp Excel nào trong thư mục đã chọn.")

    def update_progress(self):
        if not self.running:
            return
        if self.progress_queue:
            while not self.progress_queue.empty():
                self.progress_queue.get()
                self.progress_count += 1
                self.progress_bar["value"] = self.progress_count
                percent = int((self.progress_count / self.total_files) * 100) if self.total_files > 0 else 0
                self.progress_label.config(text=f"{percent}%")
        self.after_id_progress = self.after(1000, self.update_progress)

    def run_vba_on_all_thread(self):
        self.stop_event.clear()
        self.vba_thread = threading.Thread(target=self.run_vba_on_all)
        self.vba_thread.start()

    def run_vba_on_all(self):
        logger.info("Bắt đầu chạy VBA trên các tệp Excel.")
        dev_dir = os.environ.get('DEV') or os.getcwd()
        test_dir = os.path.join(dev_dir, 'test_files')
        if not self.excel_directory:
            self.excel_directory = os.path.join(test_dir, 'excel')
        if not self.vba_file:
            self.vba_file = os.path.join(self.excel_directory, 'test_macro.bas')
        globals()["global_vba_file_path"] = self.vba_file

        excel_files = glob.glob(os.path.join(self.excel_directory, "*.xlsx"))
        if not excel_files:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
            return

        self.total_files = len(excel_files)
        self.progress_count = 0
        self.progress_bar["maximum"] = self.total_files
        self.progress_bar["value"] = 0

        num_processes = os.cpu_count() - 2 or 1
        batch_size = self.total_files // num_processes
        remainder = self.total_files % num_processes
        batches = []
        start = 0
        for i in range(num_processes):
            extra = 1 if i < remainder else 0
            end = start + batch_size + extra
            batches.append(excel_files[start:end])
            start = end

        logger.info(f"Bắt đầu chạy VBA trên {self.total_files} tệp, chia thành {num_processes} batch")

        if self.mp_logging.queue is None:
            raise ValueError("Hàng đợi logging chia sẻ chưa được thiết lập!")
        from mpp_logger import get_mp_logger
        self.progress_queue = get_mp_logger().manager.Queue()
        shared_queue = self.mp_logging.queue

        from multiprocessing import Pool
        pool = Pool(processes=num_processes,
                    initializer=worker_logging_setup,
                    initargs=(shared_queue, self.mp_logging.log_level.value))
        for batch in batches:
            pool.apply_async(worker.process_batch, args=(batch, self.progress_queue, self.mp_logging.queue))
        pool.close()
        pool.join()

        import time
        time.sleep(1)
        logger.info(f"Đã chạy VBA trên {self.total_files} tệp Excel.")

    def exit_app(self):
        self.running = False
        if hasattr(self, 'vba_thread') and self.vba_thread.is_alive():
            self.vba_thread.join(timeout=5)
        self.mp_logging.shutdown()
        if self.after_id_progress is not None:
            self.after_cancel(self.after_id_progress)
        self.destroy()
