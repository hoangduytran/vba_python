# gui.py

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
import glob, os, threading
from multiprocessing import Pool
import worker
from worker import worker_logging_setup
from mpp_logger import get_mp_logger, ExactLevelFilter
from logtext import LogText  # Lớp LogText do bạn định nghĩa (phần giao diện hiển thị log)

# Khai báo biến toàn cục logger, sẽ được gán trong MainWindow
logger = None

# Kiểu giao diện chung cho các widget (đã định nghĩa)
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}

# Định nghĩa các mức log với tên tiếng Việt; 
# "NO_LOGGING" được đặt thành 100 để không hiển thị log nào khi được chọn.
LOG_LEVELS = {
    "NO_LOGGING": 100,   # Không hiển thị log nào trong GUI.
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}

# Lớp TextHandler để xử lý ghi log vào widget Text của Tkinter.
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.is_gui_handler = True  # Đánh dấu đây là handler dùng cho giao diện

    def emit(self, record):
        try:
            # Gọi hàm format để định dạng bản ghi log theo định dạng của PrettyFormatter hay tương tự,
            # sau đó thêm dấu xuống dòng.
            msg = self.format(record) + "\n"
            # Dùng phương thức after của widget để chèn log vào Text một cách bất đồng bộ.
            self.text_widget.after(0, self.append, msg)
        except Exception:
            self.handleError(record)

    def append(self, msg):
        # Cho phép chỉnh sửa widget, chèn thông điệp log, sau đó khóa lại widget và cuộn xuống cuối.
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg)
        self.text_widget.configure(state="disabled")
        self.text_widget.yview(tk.END)

# Lớp MainWindow: cửa sổ chính giao diện Tkinter của ứng dụng.
class MainWindow(tk.Tk):
    def __init__(self, mp_logging):
        super().__init__()
        self.mp_logging = mp_logging
        
        # Gán logger toàn cục từ mp_logging.logger cho biến toàn cục logger.
        global logger
        logger = self.mp_logging.logger

        # Ghi log cho khởi động ứng dụng (đã có thông điệp bằng tiếng Việt)
        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")

        # Cài đặt các thuộc tính cho cửa sổ.
        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Khởi tạo các biến theo dõi quá trình xử lý.
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None
        self.stop_event = threading.Event()

        # ----------------------
        # Khu vực thanh công cụ bên trái.
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Tạo một dropdown (OptionMenu) để người dùng chọn mức log cần hiển thị.
        # Giá trị mặc định là "INFO". Khi thay đổi, callback select_log_level sẽ được gọi.
        self.log_level_var = tk.StringVar(value="INFO")
        self.log_level_menu = ttk.OptionMenu(
            self.taskbar, 
            self.log_level_var, 
            "INFO", 
            *LOG_LEVELS.keys(), 
            command=self.select_log_level
        )
        self.log_level_menu.config(width=20)
        self.log_level_menu.pack(pady=5, anchor="w")

        # Tạo các nút trên thanh công cụ.
        self.create_taskbar_buttons()

        # ----------------------
        # Khu vực hiển thị chính bên phải.
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Khởi tạo LogText widget để hiển thị log; LogText nhận đối tượng mp_logging để có thể thao tác với hệ thống log.
        self.log_container = LogText(right_area, self.mp_logging)
        self.log_container.pack(fill="both", expand=True)

        # Tạo thanh tiến trình và nhãn hiển thị % hoàn thành.
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))

        # Lên lịch cập nhật tiến trình định kỳ.
        self.after_id_progress = self.after(500, self.update_progress)

        # Gắn một TextHandler vào QueueListener để hiển thị log trong vùng Text của giao diện.
        if self.mp_logging.listener is not None:
            self.gui_handler = TextHandler(self.log_container.log_text)
            # Sử dụng PrettyFormatter để hiển thị log theo định dạng dễ đọc.
            from mpp_logger import PrettyFormatter  # Nếu có định nghĩa PrettyFormatter
            gui_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            self.gui_handler.setFormatter(gui_formatter)
            # Thêm filter để hiển thị chỉ các bản ghi có mức log >= mức đã chọn.
            current_level = LOG_LEVELS.get(self.log_level_var.get(), logging.INFO)
            self.gui_handler.addFilter(ExactLevelFilter(current_level))
            # Kết hợp handler mới vào danh sách các handler đã có của QueueListener.
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (self.gui_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        # Đăng ký sự kiện đóng cửa sổ.
        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def create_taskbar_buttons(self):
        """
        Tạo các nút trên thanh công cụ dựa trên cấu hình:
          - "Lưu Log vào tập tin"
          - "Tải tệp VBA"
          - "Tải thư mục Excel"
          - "Chạy VBA trên tất cả các tệp Excel"
          - "Thoát Ứng dụng"
        Các nút được đóng gói theo kiểu chung.
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

    def select_log_level(self, selected):
        level = LOG_LEVELS.get(selected, logging.INFO)
        self.mp_logging.select_log_level(level)
        # If you saved a reference to the GUI handler, update its filter explicitly:
        self.gui_handler.filters = []
        self.gui_handler.addFilter(ExactLevelFilter(level))
        logger.info(f"Log level changed to {selected}")

    def save_log(self):
        """
        Mở hộp thoại lưu file với hai định dạng: JSON và TXT.
        Nếu người dùng chọn JSON thì copy file log tạm (chứa log dưới định dạng JSON) sang file được chọn.
        Nếu người dùng chọn TXT, lưu nội dung của vùng Text hiển thị log.
        """
        import shutil
        path = filedialog.asksaveasfilename(
            title="Lưu Log vào tập tin",
            defaultextension=".txt",
            filetypes=[("Tệp văn bản (*.txt)", "*.txt"), ("Tệp JSON (*.json)", "*.json")]
        )
        if path:
            try:
                if path.lower().endswith(".json"):
                    # Nếu chọn JSON: copy file log tạm đã được ghi dưới định dạng JSON.
                    shutil.copyfile(self.mp_logging.log_temp_file_path, path)
                else:
                    # Nếu chọn TXT: lấy nội dung của vùng Text và lưu ra file.
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(self.log_container.log_text.get("1.0", tk.END))
                messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
                logger.info("Log đã được lưu thành công")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")

    def load_vba_file(self):
        """
        Cho phép người dùng chọn tệp VBA và cập nhật đường dẫn của tệp đó.
        Sau đó, ghi log thông báo đã tải tệp VBA.
        """
        init_dir = self.excel_directory if self.excel_directory else os.getcwd()
        path = filedialog.askopenfilename(
            title="Chọn tệp VBA",
            defaultextension=".bas",
            initialdir=init_dir,
            filetypes=[("Tệp VBA (*.bas)", "*.bas"), ("Tất cả các tệp", "*.*")]
        )
        if path:
            self.vba_file = path
            globals()["global_vba_file_path"] = path
            logger.info(f"Đã tải tệp VBA: {path}")
            messagebox.showinfo("Thông báo", f"Đã tải tệp VBA: {path}")

    def load_excel_directory(self):
        """
        Cho phép người dùng chọn thư mục chứa tệp Excel.
        Sau đó kiểm tra số lượng tệp Excel có trong thư mục và ghi log phù hợp.
        """
        init_dir = os.path.dirname(self.vba_file) if self.vba_file else os.getcwd()
        directory = filedialog.askdirectory(
            title="Chọn thư mục chứa các tệp Excel",
            initialdir=init_dir
        )
        if directory:
            self.excel_directory = directory
            import glob
            excel_files = glob.glob(os.path.join(directory, "*.xlsx"))
            if excel_files:
                logger.info(f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
                messagebox.showinfo("Thông báo", f"Đã tải thư mục Excel: {directory}, loaded {len(excel_files)} files.")
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy tệp Excel nào trong thư mục đã chọn.")
                logger.info("Không tìm thấy tệp Excel nào trong thư mục đã chọn.")

    def update_progress(self):
        """
        Cập nhật thanh tiến trình và nhãn phần trăm hoàn thành dựa trên số lượng tệp đã xử lý.
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
        Khởi chạy tác vụ chạy macro VBA trên tất cả các tệp Excel trong một luồng riêng để giao diện không bị treo.
        """
        self.stop_event.clear()
        self.vba_thread = threading.Thread(target=self.run_vba_on_all)
        self.vba_thread.start()

    def run_vba_on_all(self):
        """
        Thực hiện các bước xử lý:
          - Đọc thông tin từ giao diện (đường dẫn tệp VBA và thư mục Excel)
          - Tìm tệp Excel trong thư mục đã chọn
          - Phân lô các tệp Excel và khởi chạy các tiến trình xử lý
          - Cập nhật tiến trình và ghi log cho các bước trên.
        """
        logger.info("Bắt đầu chạy VBA trên các tệp Excel.")
        dev_dir = os.environ.get('DEV') or os.getcwd()
        test_dir = os.path.join(dev_dir, 'test_files')
        if not self.excel_directory:
            self.excel_directory = os.path.join(test_dir, 'excel')
        if not self.vba_file:
            self.vba_file = os.path.join(self.excel_directory, 'test_macro.bas')
        globals()["global_vba_file_path"] = self.vba_file

        import glob
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
        pool = Pool(
            processes=num_processes,
            initializer=worker_logging_setup,
            initargs=(shared_queue, self.mp_logging.log_level.value)
        )
        for batch in batches:
            pool.apply_async(worker.process_batch, args=(batch, self.progress_queue, self.mp_logging.queue))
        pool.close()
        pool.join()

        import time
        time.sleep(1)
        logger.info(f"Đã chạy VBA trên {self.total_files} tệp Excel.")

    def exit_app(self):
        """
        Xử lý đóng ứng dụng:
          - Dừng các tác vụ chạy nền (nếu có)
          - Tắt hệ thống logging thông qua instance mp_logging
          - Hủy bỏ các tác vụ đã lên lịch và đóng cửa sổ GUI
        """
        self.running = False
        if hasattr(self, 'vba_thread') and self.vba_thread.is_alive():
            self.vba_thread.join(timeout=5)
        self.mp_logging.shutdown()
        if self.after_id_progress is not None:
            self.after_cancel(self.after_id_progress)
        self.destroy()
