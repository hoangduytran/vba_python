import logging
import sys
import tempfile
from multiprocessing import Manager
from logging.handlers import QueueHandler, QueueListener

class SafeQueueHandler(QueueHandler):
    """
    Một QueueHandler đảm bảo rằng hàng đợi được sử dụng (queue) còn hợp lệ trước khi phát hành các bản ghi log.
    Điều này giúp tránh lỗi (ví dụ trong quá trình tắt hệ thống) nếu queue là None.
    """
    def __init__(self, queue, formatter=None):
        super().__init__(queue)
        if formatter:
            self.setFormatter(formatter)
    
    def emit(self, record):
        # Kiểm tra nếu queue không còn hợp lệ thì không làm gì.
        if not self.queue:
            return
        try:
            self.queue.put_nowait(record)
        except Exception:
            self.handleError(record)

class LoggingMultiProcess:
    """
    Bao bọc cấu hình logging cho đa tiến trình.
    
    Lớp này tạo ra:
      - Một Manager và một hàng đợi logging chia sẻ.
      - Một QueueListener xử lý các bản ghi log và chuyển chúng đến FileHandler và StreamHandler.
      - Cấu hình logger toàn cục ("main_logger") với một SafeQueueHandler.
    
    Lớp cũng định nghĩa một cờ is_debug (mặc định True) và một phương thức DEBUG_LOG để ghi log ở mức debug
    chỉ khi is_debug được bật.
    """
    # Đặt tên logger cố định.
    MAIN_LOGGER = "main_logger"
    DEFAULT_FORMAT = "%(asctime)s - %(processName)s - %(levelname)s - %(message)s"

    def __init__(self):
        # Tạo Manager và hàng đợi chia sẻ logging.
        self.manager = Manager()
        self.queue = self.manager.Queue()
        # Use a Manager.Value for is_debug so that it is shared among processes.
        self.is_debug = self.manager.Value('b', True)

        # Tạo một tệp tạm thời để ghi log.
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".log")
        self.log_temp_file_path = temp_file.name
        temp_file.close()

        # Định nghĩa formatter mặc định.
        self.default_formatter = logging.Formatter(self.DEFAULT_FORMAT)

        # Cài đặt FileHandler sử dụng formatter mặc định.
        file_handler = logging.FileHandler(self.log_temp_file_path, mode="w", encoding="utf-8")
        file_handler.setFormatter(self.default_formatter)

        # Cài đặt StreamHandler để xuất log ra stdout.
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(self.default_formatter)

        # Tạo và bắt đầu QueueListener với cả FileHandler và StreamHandler.
        self.listener = QueueListener(self.queue, file_handler, stream_handler)
        self.listener.start()
        print("Tệp log tạm thời:", self.log_temp_file_path)

        # Khởi tạo logger toàn cục "main_logger" sử dụng tên cố định.
        self.logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
        self.logger.setLevel(logging.DEBUG)
        # Gắn SafeQueueHandler để gửi các bản ghi log vào hàng đợi chia sẻ.
        safe_handler = SafeQueueHandler(self.queue, formatter=self.default_formatter)
        self.logger.addHandler(safe_handler)

    def DEBUG_LOG(self, msg):
        """
        Ghi một thông báo debug bằng cách sử dụng logger toàn cục nếu is_debug được bật.
        """
        if self.is_debug:
            self.logger.debug(msg)

    def reinit(self):
        """
        Trả về một instance mới của SafeQueueHandler bọc hàng đợi chia sẻ với formatter mặc định.
        Phương thức này hữu ích cho việc khởi tạo logger trong các tiến trình con.
        """
        return SafeQueueHandler(self.queue, formatter=self.default_formatter)

    @classmethod
    def get_worker_handler(cls, queue):
        """
        Trả về một instance mới của SafeQueueHandler với formatter mặc định sử dụng hàng đợi được cung cấp.
        Hữu ích cho việc cấu hình logging trong các tiến trình con.
        """
        formatter = logging.Formatter(cls.DEFAULT_FORMAT)
        return SafeQueueHandler(queue, formatter=formatter)

    def shutdown(self):
        """
        Dừng QueueListener và tắt Manager.
        Phương thức này được gọi khi ứng dụng kết thúc để giải phóng các tài nguyên.
        """
        if self.listener:
            try:
                self.listener.stop()
            except BrokenPipeError:
                pass
            except Exception as e:
                print("Lỗi khi dừng QueueListener:", e)
        if self.manager:
            self.manager.shutdown()


# Bộ truy cập toàn cục cho một instance singleton của LoggingMultiProcess.
_mp_logger = None

def get_mp_logger():
    """
    Trả về instance của LoggingMultiProcess.
    Đảm bảo rằng chỉ tạo một instance duy nhất.
    """
    global _mp_logger
    if _mp_logger is None:
        _mp_logger = LoggingMultiProcess()
    return _mp_logger

def DEBUG_LOG(msg):
    """
    Hàm toàn cục để ghi một thông báo debug sử dụng cấu hình logging chia sẻ.
    Chỉ ghi khi is_debug được bật.
    """
    get_mp_logger().DEBUG_LOG(msg)
