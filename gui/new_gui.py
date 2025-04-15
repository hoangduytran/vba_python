import logging
import tkinter as tk
from tkinter import ttk
from gv import Gvar as gv, COMMON_WIDGET_STYLE
from gui_actions import action_list  # Import callback actions
from mpp_logger import get_mp_logger, LOG_LEVELS, TextHandler, PrettyFormatter, DynamicLevelFilter
from logtext import LogText, ToolTip
import threading

logger = None

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        global logger

        logger = gv.logger
        gv.root = self
        self.mp_logging = gv.mp_logging

        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")
        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # State variables
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None
        self.stop_event = threading.Event()
        self.vba_thread = None
        self.after_id_progress = None

        # Tạo style TTK (nếu cần)
        self.style = ttk.Style(self)
        self.style.configure(
            "App.TMenubutton",
            font=("Arial", 18, "normal"),
            padding=8
        )

        # -----------
        # LEFT SIDEBAR
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Tạo các nút khác (save_log, load_vba_file, v.v.)
        # (The log_level_menu and exact_check have been removed here)
        buttons_config = [
            {"text": "Nạp tập tin VBA", "action": "load_vba_file"},
            {"text": "Chọn thư mục Excel", "action": "load_excel_directory"},
            {"text": "Chạy VBA trên các tập tin Excel", "action": "run_macro_thread"},
            {"text": "Thoát ứng dụng", "action": "exit_app"}
        ]        

        for conf in buttons_config:
            btn = tk.Button(
                self.taskbar,
                text=conf["text"],
                command=action_list[conf["action"]],
                **COMMON_WIDGET_STYLE
            )
            btn.pack(pady=3, fill="x", anchor="w")

        # -----------
        # MAIN RIGHT AREA
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Create LogText (this now includes the log_level_menu & checkbox in its own toolbar)
        self.log_container = LogText(right_area, self.mp_logging)
        self.log_container.pack(fill="both", expand=True)

        # Create progress bar
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        gv.progress_bar = self.progress_bar

        # Label for percentage
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))        
        self.progress_label.pack(pady=(0, 5))        
        gv.progress_label = self.progress_label

        # Logging GUI handler
        if self.mp_logging.listener is not None:
            self.gui_handler = TextHandler(self.log_container.log_text)
            gui_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            self.gui_handler.setFormatter(gui_formatter)
            # For initial filter
            # (If you want to keep the default set to DEBUG, do so in your actions or here)
            self.gui_handler.addFilter(DynamicLevelFilter(logging.DEBUG, True))
            self.mp_logging.listener.handlers += (self.gui_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        # Periodic progress update
        self.after_id_progress = self.after(500, action_list["update_progress"])

        # On window close
        self.protocol("WM_DELETE_WINDOW", action_list["exit_app"])
