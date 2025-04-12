import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox, filedialog
import logging
import glob, os, threading
from multiprocessing import Manager, Pool
import worker
from worker import worker_logging_setup
from mpp_logger import get_mp_logger, DEBUG_LOG, IsDebugFilter  # Nhập hàm và instance global của logging
from logtext import LogText  # Nhập lớp LogText đã được định nghĩa riêng

# Kiểu giao diện chung cho các widget
COMMON_WIDGET_STYLE = {
    "font": ("Arial", 18, "bold"),
    "width": 25,
    "height": 3
}

class MainWindow(tk.Tk):
    def __init__(self, mp_logging):
        super().__init__()
        # Lấy instance của LoggingMultiProcess thông qua get_mp_logger()
        self.mp_logging = get_mp_logger()

        # Ghi log khởi chạy ứng dụng (theo trạng thái is_debug)
        DEBUG_LOG("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")

        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Các biến theo dõi tiến trình
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None

        # Tạo sự kiện dừng cho các tác vụ chạy nền dài
        self.stop_event = threading.Event()

        # ----------------------
        # KHU VỰC CÔNG CỤ (Bên trái)
        # ----------------------
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Checkbutton để bật/tắt chế độ gỡ lỗi
        self.debug_var = tk.BooleanVar(value=self.mp_logging.is_debug.value)
        self.debug_check = tk.Checkbutton(
            self.taskbar,
            text="Gỡ Lỗi (ON/OFF)",
            variable=self.debug_var,
            command=self.toggle_debug,
            **COMMON_WIDGET_STYLE
        )
        self.debug_check.pack(pady=5, anchor="w")

        # Gọi hàm tạo các nút trên taskbar
        self.create_taskbar_buttons()

        # ----------------------
        # KHU VỰC CHÍNH (Hiển thị log và thanh tiến trình)
        # ----------------------
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Thay vì tạo widget Text trực tiếp, tạo instance của LogText để bao gồm cả toolbar và vùng log
        self.log_container = LogText(right_area)
        self.log_container.pack(fill="both", expand=True)

        # Thanh tiến trình và nhãn hiển thị % hoàn thành
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))

        # Lên lịch cập nhật tiến trình
        self.after_id_progress = self.after(500, self.update_progress)

        # Gắn TextHandler vào QueueListener để cập nhật log trong GUI,
        # sử dụng vùng Text của LogText (được đặt trong self.log_container.log_text)
        if self.mp_logging.listener is not None:
            from logging import Formatter
            text_handler = TextHandler(self.log_container.log_text)
            text_handler.setFormatter(self.mp_logging.default_formatter)
            # Kết hợp với các handler đã có của QueueListener
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (text_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        # Gán sự kiện đóng cửa sổ
        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def create_taskbar_buttons(self):
        """
        Tạo các nút trên thanh công cụ dựa trên danh sách cấu hình.
        Mỗi nút có thuộc tính 'text' và 'command'.
        Danh sách này chứa cấu hình cho các nút (ngoại trừ checkbox).
        Các nút được tạo và đóng gói theo thứ tự.
        """
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

    def toggle_debug(self):
        """
        Thay đổi trạng thái của chế độ gỡ lỗi dựa trên giá trị của checkbox.
        Khi tắt (False), các thông báo log DEBUG sẽ không được xuất ra.
        """
        # Lấy giá trị boolean từ GUI.
        new_value = self.debug_var.get()
        # Cập nhật cờ is_debug (sử dụng biến bool đơn giản, không cần .value).
        self.mp_logging.is_debug = new_value  
        # Điều chỉnh mức logging của logger chính.
        if new_value:
            self.mp_logging.logger.setLevel(logging.DEBUG)
        else:
            self.mp_logging.logger.setLevel(logging.INFO)
        status = "bật" if new_value else "tắt"
        DEBUG_LOG(f"Chế độ gỡ lỗi được {status}")

    def save_log(self):
        """
        Lưu nội dung log từ tệp tạm vào tập tin do người dùng chỉ định.
        """
        path = filedialog.asksaveasfilename(
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
        path = filedialog.askopenfilename(
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
        directory = filedialog.askdirectory(
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
        Cập nhật thanh tiến trình và nhãn hiển thị phần trăm hoàn thành.
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
        Khởi chạy tác vụ chạy VBA trên các tệp Excel trong luồng riêng để giữ giao diện luôn phản hồi.
        """
        self.stop_event.clear()
        self.vba_thread = threading.Thread(target=self.run_vba_on_all)
        self.vba_thread.start()

    def run_vba_on_all(self):
        """
        Xử lý tất cả các tệp Excel trong thư mục đã chọn bằng cách sử dụng nhiều tiến trình.
        Nếu không có tệp hoặc thư mục, sử dụng giá trị mặc định.
        """
        # khi chạy thì comment mấy dòng dưới đây lại, 
        # từ đây
        DEBUG_LOG("Bắt đầu chạy VBA trên các tệp Excel.")
        dev_dir = os.environ.get('DEV') or os.getcwd()
        test_dir = os.path.join(dev_dir, 'test_files')
        if not self.excel_directory:
            self.excel_directory = os.path.join(test_dir, 'excel')
        if not self.vba_file:
            self.vba_file = os.path.join(self.excel_directory, 'test_macro.bas')
        globals()["global_vba_file_path"] = self.vba_file
        # cho đến đây, để lấy thông tin thực

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
            raise ValueError("Hàng đợi logging chia sẻ chưa được thiết lập!")
        # Tạo hàng đợi tiến trình riêng, sử dụng Manager của mp_logging
        mp_logger = get_mp_logger()  # Lấy instance của LoggingMultiProcess
        self.progress_queue = mp_logger.manager.Queue()
        shared_queue = mp_logger.queue  # Hàng đợi logging chia sẻ
        shared_is_debug = mp_logger.is_debug

        pool = Pool(processes=num_processes, 
                    initializer=worker_logging_setup, 
                    initargs=(shared_queue, shared_is_debug))

        for batch in batches:
            pool.apply_async(worker.process_batch, args=(batch, self.progress_queue, self.mp_logging.queue))
        pool.close()
        pool.join()

        import time
        time.sleep(1)  # Thêm thời gian chờ để QueueListener xử lý các log chờ

        DEBUG_LOG(f"Đã chạy VBA trên {self.total_files} tệp Excel.")

    def exit_app(self):
        """
        Xử lý đóng ứng dụng:
          - Dừng các tác vụ chạy nền (nếu có)
          - Tắt hệ thống logging thông qua instance LoggingMultiProcess
          - Hủy bỏ các tác vụ đã lên lịch và đóng cửa sổ GUI
        """
        self.running = False
        if hasattr(self, 'vba_thread') and self.vba_thread.is_alive():
            self.vba_thread.join(timeout=5)
        self.mp_logging.shutdown()
        if self.after_id_progress is not None:
            self.after_cancel(self.after_id_progress)
        self.destroy()

# Định nghĩa một TextHandler đơn giản (nếu chưa có trong dự án của bạn)
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
        self.text_widget.configure(state="normal")
        self.text_widget.yview(tk.END)
