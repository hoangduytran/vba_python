import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox
import logging
import glob, os, threading
from multiprocessing import Manager, Pool
import worker
from mpp_logger import get_mp_logger, DEBUG_LOG  # Import hàm và instance global của logging


# Kiểu giao diện chung cho các widget
COMMON_WIDGET_STYLE = {
    "font": ("Arial", 18, "bold"),
    "width": 25,
    "height": 3
}

# Một TextHandler đơn giản để cập nhật widget Text của GUI.
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.after(0, self.append, msg)

    def append(self, msg):
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg)
        self.text_widget.configure(state="disabled")
        self.text_widget.yview(tk.END)

class MainWindow(tk.Tk):
    def __init__(self, mp_logging):
        super().__init__()
        # Lấy instance LoggingMultiProcess thông qua get_mp_logger()
        self.mp_logging = get_mp_logger()

        # Ghi log về việc khởi chạy ứng dụng (theo điều kiện is_debug)
        DEBUG_LOG("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")

        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Các biến theo dõi tiến trình.
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None

        # Tạo sự kiện dừng khi chạy tác vụ dài.
        self.stop_event = threading.Event()

        # ----------------------
        # KHU VỰC CÔNG CỤ (Bên trái)
        # ----------------------
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Checkbutton để bật/tắt chế độ gỡ lỗi.
        self.debug_var = tk.BooleanVar(value=self.mp_logging.is_debug)
        self.debug_check = tk.Checkbutton(
            self.taskbar,
            text="Gỡ Lỗi (ON/OFF)",
            variable=self.debug_var,
            command=self.toggle_debug,
            **COMMON_WIDGET_STYLE
        )
        self.debug_check.pack(pady=5, anchor="w")

        # Gọi hàm tạo nút trên taskbar
        self.create_taskbar_buttons()

        # ----------------------
        # KHU VỰC CHÍNH (Hiển thị log và thanh tiến trình)
        # ----------------------
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Tạo widget Text để hiển thị log.
        # Tạo widget Text để hiển thị log.
        # Now use this font object in your Text widget
        # Create a font from the common settings
        common_font = tkFont.Font(font=COMMON_WIDGET_STYLE["font"])
        # Change the weight to normal (non-bold)
        common_font.configure(weight="normal")

        self.log_text = tk.Text(right_area, wrap="none", font=common_font)
        self.log_text.config(state="disabled")
        self.log_text.pack(fill="both", expand=True)

        # Thanh cuộn cho widget Text.
        self.v_scroll = tk.Scrollbar(self.log_text, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.v_scroll.set)
        self.v_scroll.pack(side="right", fill="y")

        # Thanh tiến trình và nhãn hiển thị phần trăm.
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))

        # Lên lịch cập nhật tiến trình.
        self.after_id_progress = self.after(500, self.update_progress)

        # Gắn TextHandler vào QueueListener để cập nhật log trong GUI.
        if self.mp_logging.listener is not None:
            from logging import Formatter
            text_handler = TextHandler(self.log_text)
            text_handler.setFormatter(self.mp_logging.default_formatter)
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (text_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        # Gán sự kiện đóng cửa sổ.
        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def create_taskbar_buttons(self):
        """
        Tạo các nút trên taskbar dựa trên danh sách cấu hình.
        
        Mỗi nút có 2 thuộc tính chính: 'text' và 'command'.
        Danh sách này chứa các cấu hình cho các nút (ngoại trừ checkbox).
        Các nút sẽ được tạo và hiển thị theo thứ tự của danh sách.
        """
        # Danh sách các nút với văn bản và callback tương ứng.
        # Lưu ý: Bạn có thể thêm hoặc thay đổi các nút ở đây.
        buttons_config = [
            {"text": "Lưu Log vào tập tin", "command": self.save_log},
            {"text": "Tải tệp VBA", "command": self.load_vba_file},
            {"text": "Tải thư mục Excel", "command": self.load_excel_directory},
            {"text": "Chạy VBA trên tất cả các tệp Excel", "command": self.run_vba_on_all_thread},
            {"text": "Thoát Ứng dụng", "command": self.exit_app}
        ]
        # Tạo các nút và đóng gói (pack) chúng theo thứ tự.
        for btn_conf in buttons_config:
            btn = tk.Button(self.taskbar, text=btn_conf["text"], command=btn_conf["command"], **COMMON_WIDGET_STYLE)
            btn.pack(pady=3, fill="x", anchor="w")

    def toggle_debug(self):
        """
        Thay đổi trạng thái của is_debug trong LoggingMultiProcess.
        """
        self.mp_logging.is_debug = self.debug_var.get()
        status = "bật" if self.mp_logging.is_debug else "tắt"
        DEBUG_LOG(f"Chế độ gỡ lỗi được {status}")

    def save_log(self):
        """
        Lưu nội dung log từ tệp tạm vào tập tin do người dùng chỉ định.
        """
        path = tk.filedialog.asksaveasfilename(
            title="Lưu Log vào tập tin", defaultextension=".txt",
            filetypes=[("Tệp văn bản", "*.txt"), ("Tất cả các tệp", "*.*")]
        )
        if path:
            try:
                with open(self.mp_logging.log_temp_file_path, "r", encoding="utf-8") as src, \
                     open(path, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
                messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
                DEBUG_LOG("Log đã được lưu thành công")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")

    def load_vba_file(self):
        """
        Cho phép người dùng chọn tệp VBA và cập nhật tham chiếu.
        """
        init_dir = self.excel_directory if self.excel_directory else os.getcwd()
        path = tk.filedialog.askopenfilename(
            title="Chọn tệp VBA", defaultextension=".bas",
            initialdir=init_dir,
            filetypes=[("Tệp VBA", "*.bas"), ("Tất cả các tệp", "*.*")]
        )
        if path:
            self.vba_file = path
            globals()["global_vba_file_path"] = path
            DEBUG_LOG(f"Đã tải tệp VBA: {path}")
            messagebox.showinfo("Thông báo", f"Đã tải tệp VBA: {path}")

    def load_excel_directory(self):
        """
        Cho phép người dùng chọn thư mục chứa các tệp Excel và hiển thị số lượng file được tìm thấy.
        """
        init_dir = os.path.dirname(self.vba_file) if self.vba_file else os.getcwd()
        directory = tk.filedialog.askdirectory(
            title="Chọn thư mục chứa các tệp Excel", initialdir=init_dir
        )
        if directory:
            self.excel_directory = directory
            excel_files = glob.glob(os.path.join(directory, "*.xlsx"))
            if excel_files:
                DEBUG_LOG(f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
                messagebox.showinfo("Thông báo", f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
                DEBUG_LOG("Không tìm thấy tệp Excel nào trong thư mục đã chọn.")

    def update_progress(self):
        """
        Cập nhật thanh tiến trình và nhãn % hoàn thành.
        """
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
        """
        Khởi chạy tác vụ chạy VBA trên các tệp Excel trong luồng riêng để giữ GUI luôn phản hồi.
        """
        self.stop_event.clear()
        self.vba_thread = threading.Thread(target=self.run_vba_on_all)
        self.vba_thread.start()

    def run_vba_on_all(self):
        """
        Xử lý tất cả các tệp Excel trong thư mục được chọn bằng cách sử dụng nhiều tiến trình.
        Nếu không có tệp hay thư mục được chọn, sử dụng giá trị mặc định.
        """
        DEBUG_LOG("Bắt đầu chạy VBA trên các tệp Excel.")
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

        DEBUG_LOG(f"Bắt đầu chạy VBA trên {self.total_files} tệp, chia thành {num_processes} batch")

        if self.mp_logging.queue is None:
            raise ValueError("shared_log_queue must be set!")
        mgr = Manager()
        self.progress_queue = mgr.Queue()

        pool = Pool(processes=num_processes)
        for batch in batches:
            pool.apply_async(worker.process_batch, args=(batch, self.progress_queue, self.mp_logging.queue))
        pool.close()
        pool.join()

        import time
        time.sleep(1)  # Thêm thời gian chờ để QueueListener xử lý các log chờ.

        DEBUG_LOG(f"Đã chạy VBA trên {self.total_files} tệp Excel.")

    def exit_app(self):
        """
        Xử lý đóng ứng dụng:
          - Dừng các tác vụ chạy nền (nếu có).
          - Tắt hệ thống logging thông qua instance LoggingMultiProcess.
          - Hủy bỏ các tác vụ đã lên lịch và đóng cửa sổ GUI.
        """
        self.running = False
        if hasattr(self, 'vba_thread') and self.vba_thread.is_alive():
            self.vba_thread.join(timeout=5)
        self.mp_logging.shutdown()
        if self.after_id_progress is not None:
            self.after_cancel(self.after_id_progress)
        self.destroy()
