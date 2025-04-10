import os
import logging
import mpp_logger
from mpp_logger import DEBUG_LOG, LoggingMultiProcess, SafeQueueHandler, get_mp_logger
import win32com.client as win32

class DummyLogging:
    def __init__(self, logger, debug_flag):
        self.logger = logger
        self.debug_flag = debug_flag  # Đây là một giá trị Manager.Value

    def DEBUG_LOG(self, msg):
        # Chỉ ghi log nếu cờ debug chia sẻ có giá trị True
        if self.debug_flag.value:
            self.logger.debug(msg)

def worker_logging_setup(shared_queue, shared_is_debug):
    """
    Khởi tạo logging cho tiến trình con mà không tạo một Manager mới.
    Hàm này cấu hình logger của tiến trình con (được đặt tên là LoggingMultiProcess.MAIN_LOGGER)
    bằng một SafeQueueHandler sử dụng shared_queue, sau đó ghi đè biến _mp_logger của module
    và hàm DEBUG_LOG để các lần gọi DEBUG_LOG() trong tiến trình con sử dụng logger của tiến trình con.
    """
    # Truy xuất (hoặc tạo) logger cho tiến trình con này.
    worker_logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
    worker_logger.handlers.clear()
    
    # Tạo một SafeQueueHandler mới sử dụng shared_queue.
    new_handler = LoggingMultiProcess.get_worker_handler(shared_queue)
    worker_logger.addHandler(new_handler)
    worker_logger.setLevel(logging.DEBUG)
        
    # Ghi đè biến _mp_logger cấp module bằng thực thể DummyLogging của chúng ta.
    mpp_logger._mp_logger = DummyLogging(worker_logger, shared_is_debug)
    
    print(f"Worker logging setup running, using queue: {shared_queue}")

def process_excel_file(file_path):
    """
    Xử lý một tệp Excel đơn lẻ.
    Logger của tiến trình con (đã được cấu hình trong worker_logging_setup()) được kỳ vọng sử dụng một SafeQueueHandler
    để các thông điệp được gửi thông qua DEBUG_LOG() được chuyển qua hàng đợi chia sẻ.

    Lưu ý: Đảm bảo rằng hàm khởi tạo của tiến trình con đã cấu hình lại _mp_logger toàn cục sao cho DEBUG_LOG() hoạt động đầy đủ.
    """
    try:
        # In ra trực tiếp để kiểm tra
        print(f"PRINTING Worker ({os.getpid()}): Bắt đầu xử lý tệp: {file_path}")        
        # Ghi log từ tiến trình con sử dụng DEBUG_LOG
        DEBUG_LOG(f"Worker ({os.getpid()}): Bắt đầu xử lý tệp: {file_path}")
        
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False

        DEBUG_LOG(f"Mở file {file_path}")
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        
        # Xác định đường dẫn đến file VBA macro (macro_module.bas)
        macro_file = os.path.abspath("macro_module.bas")
        DEBUG_LOG(f"Nhập module VBA từ {macro_file} vào {file_path}")
        wb.VBProject.VBComponents.Import(macro_file)
        
        DEBUG_LOG(f"Chạy macro 'ProcessWorkbook' trên {file_path}")
        excel.Application.Run("ProcessWorkbook")
        
        wb.Save()
        wb.Close()
        excel.Application.Quit()

        # Ghi log thành công
        result_message = f"Worker ({os.getpid()}): Đã xử lý thành công {file_path}"
        print(f"PRINTING result_message {result_message}")        
        DEBUG_LOG(result_message)        
        return result_message

    except Exception as e:
        error_message = f"Worker ({os.getpid()}): Lỗi khi xử lý {file_path}: {str(e)}"
        DEBUG_LOG(error_message)
        # Ném exception để đảm bảo lỗi được truyền lên.
        raise Exception(error_message)

def process_batch(batch, progress_queue, log_q):
    """
    Xử lý một nhóm các tệp Excel (batch).

    Các bước thực hiện:
      - Gọi worker_logging_setup để cấu hình logging cho tiến trình con, sử dụng hàng đợi log chia sẻ.
      - Duyệt qua các tệp trong batch, xử lý từng tệp và cập nhật tiến trình.
      - Trả về tổng số tệp đã xử lý.
    """
    count = 0
    for file_path in batch:
        process_excel_file(file_path)
        count += 1
        progress_queue.put(1)
        print(f"Worker ({os.getpid()}): Đã xử lý {file_path} – gửi thông báo cập nhật tiến trình.")
    return count
