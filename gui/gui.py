import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
import glob, os, threading
from multiprocessing import Pool
import worker
from worker import worker_logging_setup
from mpp_logger import get_mp_logger, LOG_LEVELS, TextHandler, DynamicLevelFilter, PrettyFormatter
from logtext import LogText  # Lớp LogText do bạn định nghĩa, dùng để hiển thị log trong giao diện

# Khai báo biến toàn cục logger, sẽ được gán trong MainWindow
logger = None

# Kiểu giao diện chung cho các widget
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}

# -------------------------------------------------------------------------------
# Lớp MainWindow: Giao diện chính của ứng dụng.
# -------------------------------------------------------------------------------
class MainWindow(tk.Tk):
    def __init__(self, mp_logging):
        super().__init__()
        self.mp_logging = mp_logging
        global logger
        # Gán biến logger toàn cục bằng logger từ hệ thống logging đa tiến trình
        logger = self.mp_logging.logger

        logger.info("Ứng dụng Chạy VBA trên Excel (Tkinter) started.")

        # Cài đặt thông tin của cửa sổ giao diện
        self.title("Ứng dụng Chạy VBA trên Excel (Tkinter)")
        self.geometry("900x700")
        self.running = True

        # Các biến theo dõi tiến trình
        self.total_files = 0
        self.progress_count = 0
        self.progress_queue = None
        self.vba_file = None
        self.excel_directory = None
        self.stop_event = threading.Event()

        # ----------------------
        # Khu vực thanh công cụ (Taskbar) bên trái
        self.taskbar = tk.Frame(self, bd=2, relief=tk.RIDGE, padx=5, pady=5)
        self.taskbar.pack(side="left", fill="y")

        # Tạo dropdown cho mức log
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
        
        # Tạo checkbox cho cờ is_exact với nhãn "Chính Xác" và tooltip (nếu cần)
        self.is_exact_var = tk.BooleanVar(value=True)
        self.exact_check = tk.Checkbutton(
            self.taskbar,
            text="Chính Xác",       # Hiển thị tiếng Việt đầy đủ
            variable=self.is_exact_var,
            command=self.update_gui_filter  # Khi người dùng thay đổi, cập nhật bộ lọc GUI
        )
        self.exact_check.pack(pady=5, anchor="w")

        # Tạo các nút trên thanh công cụ
        self.create_taskbar_buttons()

        # ----------------------
        # Khu vực hiển thị chính (Right area)
        right_area = tk.Frame(self, bd=2, relief=tk.SUNKEN, padx=10, pady=10)
        right_area.pack(side="left", fill="both", expand=True)

        # Khởi tạo LogText widget để hiển thị log (giao diện hiển thị log)
        self.log_container = LogText(right_area, self.mp_logging)
        self.log_container.pack(fill="both", expand=True)

        # Tạo thanh tiến trình và nhãn % hoàn thành
        self.progress_bar = ttk.Progressbar(right_area, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_label = tk.Label(right_area, text="0%", font=("Arial", 12))
        self.progress_label.pack(pady=(0, 5))

        self.after_id_progress = self.after(500, self.update_progress)

        # Gắn TextHandler cho giao diện (GUI) vào QueueListener của hệ thống logging
        if self.mp_logging.listener is not None:
            # Tạo một instance của TextHandler để ghi log vào vùng Text của giao diện
            gui_handler = TextHandler(self.log_container.log_text)
            from mpp_logger import PrettyFormatter  # Sử dụng định nghĩa PrettyFormatter từ mpp_logger
            gui_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            gui_handler.setFormatter(gui_formatter)
            # Tạo bộ lọc động sử dụng mức log được chọn và cờ is_exact
            current_level = LOG_LEVELS.get(self.log_level_var.get(), logging.INFO)
            gui_handler.addFilter(DynamicLevelFilter(current_level, self.is_exact_var.get()))
            # Lưu lại tham chiếu của gui_handler để dễ dàng cập nhật bộ lọc sau này
            self.gui_handler = gui_handler
            # Gắn gui_handler vào danh sách các handler của QueueListener (chỉ một lần)
            self.mp_logging.listener.handlers = self.mp_logging.listener.handlers + (gui_handler,)
        else:
            print("Cảnh báo: Không có listener hoạt động.")

        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def create_taskbar_buttons(self):
        """
        Tạo các nút trên thanh công cụ với các chức năng:
          - Lưu Log vào tập tin
          - Tải tệp VBA
          - Tải thư mục Excel
          - Chạy VBA trên tất cả các tệp Excel
          - Thoát Ứng dụng
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

    def update_gui_filter(self):
        """
        Cập nhật bộ lọc của GUI handler dựa trên mức log hiện tại và cờ is_exact.
        Nếu is_exact True, chỉ hiển thị các bản ghi có mức bằng đúng mức đã chọn;
        nếu is_exact False, hiển thị tất cả các bản ghi có mức lớn hơn hoặc bằng mức đã chọn.
        """
        current_level = LOG_LEVELS.get(self.log_level_var.get(), logging.INFO)
        is_exact = self.is_exact_var.get()
        # Loại bỏ các filter cũ đã được gắn vào GUI handler.
        self.gui_handler.filters = [f for f in self.gui_handler.filters if not isinstance(f, DynamicLevelFilter)]
        # Thêm filter mới với các giá trị cập nhật.
        self.gui_handler.addFilter(DynamicLevelFilter(current_level, is_exact))

    def select_log_level(self, selected):
        """
        Hàm gọi khi người dùng chọn một mức log mới từ dropdown.
        Cập nhật mức log của hệ thống logging và cập nhật bộ lọc trên GUI.
        """
        level = LOG_LEVELS.get(selected, logging.INFO)
        self.mp_logging.select_log_level(level)
        self.update_gui_filter()  # Cập nhật bộ lọc của GUI theo mức log và is_exact hiện tại.
        logger.info(f"Log level changed to {selected}")

    def save_log(self):
        """
        Mở hộp thoại lưu file với hai định dạng: JSON và TXT.
        Nếu người dùng chọn JSON, file log tạm (đã ghi log dưới định dạng JSON)
        sẽ được copy sang file được chọn.
        Nếu người dùng chọn TXT, nội dung hiển thị trong log_text sẽ được lưu.
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
                    shutil.copyfile(self.mp_logging.log_temp_file_path, path)
                else:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(self.log_container.log_text.get("1.0", tk.END))
                messagebox.showinfo("Thông báo", "Log đã được lưu thành công.")
                logger.info("Log đã được lưu thành công")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu log: {e}")

    def load_vba_file(self):
        """
        Cho phép người dùng chọn tệp VBA và cập nhật đường dẫn của tệp.
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
        Kiểm tra và ghi log số lượng tệp Excel được tìm thấy trong thư mục.
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
        Cập nhật thanh tiến trình và nhãn % hoàn thành dựa trên số lượng tệp đã xử lý.
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
        Khởi chạy tác vụ chạy macro VBA trên tất cả các tệp Excel
        trong một luồng riêng để giao diện không bị treo.
        """
        self.stop_event.clear()
        self.vba_thread = threading.Thread(target=self.run_vba_on_all)
        self.vba_thread.start()

    def run_vba_on_all(self):
        """
        Thực hiện xử lý:
          - Đọc thông tin từ giao diện.
          - Tìm các tệp Excel trong thư mục đã chọn.
          - Phân lô các tệp Excel và khởi chạy các tiến trình xử lý.
          - Cập nhật tiến trình và ghi log các bước thực hiện.
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
          - Dừng các tác vụ chạy nền.
          - Tắt hệ thống logging thông qua instance mp_logging.
          - Hủy bỏ các tác vụ đã lên lịch và đóng cửa sổ giao diện.
        """
        self.running = False
        if hasattr(self, 'vba_thread') and self.vba_thread.is_alive():
            self.vba_thread.join(timeout=5)
        self.mp_logging.shutdown()
        if self.after_id_progress is not None:
            self.after_cancel(self.after_id_progress)
        self.destroy()
