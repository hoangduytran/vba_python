# new_gui.py

import logging
import tkinter as tk
from tkinter import ttk
from gv import Gvar as gv, COMMON_WIDGET_STYLE
from gui_actions import action_list  # Import các hàm xử lý sự kiện
from mpp_logger import get_mp_logger, LOG_LEVELS, TextHandler, PrettyFormatter, DynamicLevelFilter
from logtext import LogText, ToolTip
import threading

logger = None

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        global logger

        # Khởi tạo logger và các biến toàn cục
        logger = gv.logger
        gv.root = self
        self.mp_logging = gv.mp_logging

        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) đã khởi động.")
        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Các biến trạng thái ban đầu
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None
        self.stop_event = threading.Event()
        self.vba_thread = None
        self.after_id_progress = None

        # Tạo style TTK cho giao diện
        self.style = ttk.Style(self)
        self.style.configure(
            "App.TMenubutton",
            font=("Arial", 18, "normal"),
            padding=8
        )

        # Thanh công cụ bên trái
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Danh sách cấu hình các nút bấm và chức năng tương ứng
        buttons_config = [
            {"text": "Nạp tập tin VBA", "action": "load_vba_file"},
            {"text": "Chọn thư mục Excel", "action": "load_excel_directory"},
            {"text": "Chạy VBA trên các tập tin Excel", "action": "run_macro_thread"},
            {"text": "Thoát ứng dụng", "action": "exit_app"}
        ]

        # Tạo các nút bấm trong thanh công cụ
        for conf in buttons_config:
            btn = tk.Button(
                self.taskbar,
                text=conf["text"],
                command=action_list[conf["action"]],
                **COMMON_WIDGET_STYLE
            )
            btn.pack(pady=3, fill="x", anchor="w")

        # Vùng hiển thị chính bên phải
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Hiển thị cửa sổ Log (bao gồm menu chọn log level và các tùy chọn liên quan)
        self.log_container = LogText(right_area, self.mp_logging)
        self.log_container.pack(fill="both", expand=True)

        # Thanh tiến độ (progress bar)
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        gv.progress_bar = self.progress_bar

        # Nhãn hiển thị phần trăm hoàn thành
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))
        gv.progress_label = self.progress_label

        # Cấu hình handler của logger để hiển thị log lên giao diện
        if self.mp_logging.listener is not None:
            self.gui_handler = TextHandler(self.log_container.log_text)
            gui_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            self.gui_handler.setFormatter(gui_formatter)
            # Thêm filter động để quản lý mức độ log
            self.gui_handler.addFilter(DynamicLevelFilter(logging.DEBUG, True))
            self.mp_logging.listener.handlers += (self.gui_handler,)
        else:
            print("Cảnh báo: Không có listener logging nào đang hoạt động.")

        # Cập nhật thanh tiến độ định kỳ mỗi 500ms
        self.after_id_progress = self.after(500, action_list["update_progress"])

        # Xử lý khi cửa sổ đóng (thoát ứng dụng)
        self.protocol("WM_DELETE_WINDOW", action_list["exit_app"])
