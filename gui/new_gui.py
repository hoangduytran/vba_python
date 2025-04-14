# new_gui.py
import logging
import tkinter as tk
from tkinter import ttk
from gv import Gvar as gv
from gui_actions import action_list  # Import the callbacks from gui_actions.py
from mpp_logger import get_mp_logger, LOG_LEVELS, TextHandler, PrettyFormatter, DynamicLevelFilter  # and other logging utilities
from logtext import LogText  # Your custom LogText widget class
import threading

COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}
logger = None

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        global logger

        logger = gv.logger
        # Register the main window and mp_logging instance in gv.
        gv.root = self
        self.mp_logging = gv.mp_logging
        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")
        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Application state variables.
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None
        self.stop_event = threading.Event()
        self.vba_thread = None
        self.after_id_progress = None

        # ----------------------
        # LEFT TASKBAR (Controls)
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Log level dropdown control.
        self.log_level_var = tk.StringVar(value="INFO")
        self.log_level_menu = ttk.OptionMenu(
            self.taskbar,
            self.log_level_var,
            "INFO",
            *LOG_LEVELS.keys(),
            command=action_list["select_log_level"]
        )
        self.log_level_menu.config(width=20)
        self.log_level_menu.pack(pady=5, anchor="w")

        # Checkbox for exact filter control.
        self.is_exact_var = tk.BooleanVar(value=True)
        self.exact_check = tk.Checkbutton(
            self.taskbar,
            text="Chính Xác",
            variable=self.is_exact_var,
            command=action_list["update_gui_filter"]
        )
        self.exact_check.pack(pady=5, anchor="w")

        # List of buttons and their associated actions (registered in action_list).
        btns_config = [
            {"text": "Lưu Log vào tập tin", "action": "save_log"},
            {"text": "Tải tệp VBA", "action": "load_vba_file"},
            {"text": "Tải thư mục Excel", "action": "load_excel_directory"},
            {"text": "Chạy VBA trên tất cả các tệp Excel", "action": "run_macro_thread"},
            {"text": "Thoát Ứng dụng", "action": "exit_app"}
        ]

        for config in btns_config:
            btn = tk.Button(
                self.taskbar,
                text=config["text"],
                command=action_list[config["action"]],
                **COMMON_WIDGET_STYLE
            )
            btn.pack(pady=3, fill="x", anchor="w")
            if config["action"] == "save_log":
                gv.button_save_log = btn
            elif config["action"] == "load_vba_file":
                gv.button_load_vba_file = btn
            elif config["action"] == "load_excel_directory":
                gv.button_load_excel_directory = btn
            elif config["action"] == "run_macro_thread":
                gv.button_run_macro = btn
            elif config["action"] == "exit_app":
                gv.button_exit_app = btn

        # ----------------------
        # RIGHT DISPLAY AREA
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # LogText widget (for displaying log messages)
        self.log_container = LogText(right_area, self.mp_logging)
        self.log_container.pack(fill="both", expand=True)
        gv.log_text_widget = self.log_container.log_text

        # Progress bar widget.
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        gv.progress_bar = self.progress_bar

        # Progress label showing percentage.
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))
        gv.progress_label = self.progress_label

        # Setup logging GUI handler if needed (omitted detailed configuration for brevity)
        if self.mp_logging.listener is not None:
            self.gui_handler = TextHandler(self.log_container.log_text)
            gui_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            self.gui_handler.setFormatter(gui_formatter)
            current_level = LOG_LEVELS.get(self.log_level_var.get(), logging.INFO)
            is_exact = self.is_exact_var.get()
            self.gui_handler.addFilter(DynamicLevelFilter(current_level, is_exact))
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (self.gui_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        # Schedule the periodic progress update (callback in gui_actions.py).
        self.after_id_progress = self.after(500, action_list["update_progress"])

        # Set the WM_DELETE_WINDOW protocol to trigger the exit callback.
        self.protocol("WM_DELETE_WINDOW", action_list["exit_app"])

if __name__ == "__main__":
    from mpp_logger import get_mp_logger
    mp_logging = get_mp_logger()
    app = MainWindow(mp_logging)
    app.mainloop()
