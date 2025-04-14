import logging
import sys
import tempfile
from multiprocessing import Manager
from logging.handlers import QueueHandler, QueueListener
import json

# -------------------------------------------------------------------------------
# ExactLevelFilter: Bộ lọc chỉ cho phép các bản ghi log có mức (level)
# chính xác bằng giá trị đã chỉ định (ví dụ, chỉ cho phép bản ghi với mức DEBUG nếu được đặt là DEBUG).
# -------------------------------------------------------------------------------
class ExactLevelFilter(logging.Filter):
    def __init__(self, level):
        super().__init__()
        self.level = level  # Gán mức log mong muốn

    def filter(self, record):
        # Trả về True chỉ khi mức log của bản ghi (record.levelno) bằng chính xác self.level
        return record.levelno == self.level

# -------------------------------------------------------------------------------
# JsonFormatter: Định dạng bản ghi log thành chuỗi JSON (dành cho việc ghi vào file).
# Các khóa trong JSON sẽ được xuất dưới dạng ASCII (không dấu), nhằm tránh các vấn đề mã hóa.
# -------------------------------------------------------------------------------
class JsonFormatter(logging.Formatter):
    def format(self, record):
        # Sử dụng record.msg thay vì getMessage() để tránh trùng lặp thông tin định dạng
        record.message = record.msg  
        # Tính toán luôn giá trị record.asctime theo định dạng datefmt đã cấu hình
        record.asctime = self.formatTime(record, self.datefmt)
        # Tạo dictionary các trường log với các khóa ASCII
        log_record = {
            "time stamp": record.asctime,
            "process name": record.processName,
            "filename": record.pathname,
            "function": f"ham:{record.funcName}()",
            "line number": f"dong so:{record.lineno}",
            "level": record.levelname,
            "message": record.message
        }
        # Trả về chuỗi JSON của dictionary trên
        return json.dumps(log_record)

# -------------------------------------------------------------------------------
# PrettyFormatter: Định dạng bản ghi log theo kiểu “dễ đọc” (human-friendly)
# dùng cho output lên terminal và giao diện (log_text).
# Các thông tin sẽ được hiển thị dưới dạng văn bản nhiều dòng, có dấu tiếng Việt.
# -------------------------------------------------------------------------------
class PrettyFormatter(logging.Formatter):
    def format(self, record):
        # Sử dụng record.msg để lấy thông điệp gốc (không bị định dạng lại)
        record.message = record.msg  
        # Tính toán thời gian (asctime) theo định dạng đã cấu hình
        record.asctime = self.formatTime(record, self.datefmt)
        # Tạo chuỗi nhiều dòng với các trường thông tin và nhãn tiếng Việt có dấu
        s = (
            f"mã thời gian: {record.asctime}\n"
            f"tên tiến trình: {record.processName}\n"
            f"tên tệp tin: {record.pathname}\n"
            f"hàm:{record.funcName}()\n"
            f"dòng số:{record.lineno}\n"
            f"cấp độ: {record.levelname}\n"
            f"thông điệp: {record.message}\n"
        )
        return s

# -------------------------------------------------------------------------------
# MemoryLogHandler: Handler tùy chỉnh lưu trữ mỗi bản ghi log (theo định dạng JSON)
# vào danh sách nội bộ (log_store) để sau này có thể xuất ra file nếu cần.
# -------------------------------------------------------------------------------
class MemoryLogHandler(logging.Handler):
    def __init__(self, log_store):
        super().__init__()
        self.log_store = log_store  # log_store là một danh sách lưu trữ các bản ghi log dạng dict

    def emit(self, record):
        try:
            # Sử dụng JsonFormatter để định dạng bản ghi log thành chuỗi JSON
            formatter = JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            json_record = formatter.format(record)
            # Chuyển chuỗi JSON thành đối tượng dict và thêm vào danh sách log_store
            self.log_store.append(json.loads(json_record))
        except Exception:
            self.handleError(record)

# -------------------------------------------------------------------------------
# LoggingMultiProcess: Lớp bao bọc cấu hình logging đa tiến trình
#
# Chức năng:
#   - Tạo Manager và hàng đợi (queue) chia sẻ giữa các tiến trình.
#   - Thiết lập mức log chia sẻ (log_level) dùng cho giao diện.
#   - Tạo file log tạm (temporary file) để ghi log bằng định dạng JSON.
#   - Cài đặt các formatter:
#         + PrettyFormatter để hiển thị cho terminal và giao diện.
#         + JsonFormatter để ghi log ra file (dạng JSON).
#   - Tạo các handler:
#         + Terminal handler (StreamHandler) sử dụng PrettyFormatter.
#         + File handler (FileHandler) sử dụng JsonFormatter.
#         + QueueHandler để gửi log vào hàng đợi.
#         + MemoryLogHandler để lưu các bản ghi log vào log_store.
#   - Thiết lập QueueListener để lắng nghe hàng đợi và chuyển log đến terminal và file.
#   - Cung cấp phương thức select_log_level để cập nhật mức log và bộ lọc của các handler.
# -------------------------------------------------------------------------------
class LoggingMultiProcess:
    MAIN_LOGGER = "main_logger"

    def __init__(self):
        # Tạo Manager và hàng đợi chia sẻ cho log
        self.manager = Manager()
        self.queue = self.manager.Queue()
        # Mức log chia sẻ dùng cho giao diện (mặc định là INFO)
        self.log_level = self.manager.Value('i', logging.INFO)

        # Tạo file tạm để ghi log ra file (dạng JSON)
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".log")
        self.log_temp_file_path = temp_file.name
        temp_file.close()

        # Tạo các formatter:
        self.pretty_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
        self.json_formatter = JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")

        # Tạo handler xuất log ra terminal sử dụng PrettyFormatter
        terminal_handler = logging.StreamHandler(sys.stdout)
        terminal_handler.setFormatter(self.pretty_formatter)

        # Tạo file handler ghi log ra file sử dụng JsonFormatter (mỗi dòng là JSON dictionary)
        file_handler = logging.FileHandler(self.log_temp_file_path, mode="w", encoding="utf-8")
        file_handler.setFormatter(self.json_formatter)

        # Thiết lập QueueListener để lắng nghe hàng đợi và gửi log đến terminal và file
        self.listener = QueueListener(self.queue, terminal_handler, file_handler)
        self.listener.start()
        print("Temporary log file:", self.log_temp_file_path)

        # Danh sách nội bộ để lưu trữ các bản ghi log dưới dạng dict (sau khi định dạng JSON)
        self.log_store = []

        # Tạo logger toàn cục với tên cố định MAIN_LOGGER
        self.logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
        # Đặt mức logger là DEBUG để ghi nhận tất cả các bản ghi; bộ lọc sẽ kiểm soát hiển thị cho giao diện.
        self.logger.setLevel(logging.DEBUG)
        # Gắn QueueHandler để gửi các bản ghi log vào hàng đợi chia sẻ
        safe_handler = self.get_worker_handler(self.queue)
        self.logger.addHandler(safe_handler)
        # Gắn MemoryLogHandler để lưu các bản ghi log (theo định dạng JSON) vào log_store
        memory_handler = MemoryLogHandler(self.log_store)
        self.logger.addHandler(memory_handler)
        # (Lưu ý: Việc xuất log ra terminal và file được xử lý thông qua QueueListener.
        #  Handler dành cho GUI sẽ được thêm sau trong module GUI.)

    def select_log_level(self, new_level):
        """
        Cập nhật mức log chia sẻ (dùng cho giao diện) và cập nhật các bộ lọc trên handler.
        Chỉ cho phép hiển thị các bản ghi log có mức chính xác bằng new_level.
        """
        self.log_level.value = new_level
        # Cập nhật bộ lọc cho các handler của logger
        for handler in self.logger.handlers:
            # Loại bỏ các ExactLevelFilter cũ nếu có
            handler.filters = [f for f in handler.filters if not isinstance(f, ExactLevelFilter)]
            # Thêm bộ lọc mới chỉ cho phép bản ghi log có mức bằng new_level
            handler.addFilter(ExactLevelFilter(new_level))
        # Cập nhật bộ lọc cho các handler của QueueListener (terminal và file handler)
        for handler in self.listener.handlers:
            handler.filters = [f for f in handler.filters if not isinstance(f, ExactLevelFilter)]
            handler.addFilter(ExactLevelFilter(new_level))

    @classmethod
    def get_worker_handler(cls, queue):
        """
        Trả về một instance mới của QueueHandler sử dụng hàng đợi được cung cấp.
        Ở đây sử dụng PrettyFormatter để định dạng bản ghi cho các output (terminal, GUI).
        """
        handler = QueueHandler(queue)
        return handler

    def shutdown(self):
        """
        Tắt QueueListener và shutdown Manager để giải phóng các tài nguyên.
        """
        if self.listener:
            try:
                self.listener.stop()
            except BrokenPipeError:
                pass
            except Exception as e:
                print("Error stopping QueueListener:", e)
        if self.manager:
            self.manager.shutdown()

# -------------------------------------------------------------------------------
# Hàm truy cập singleton toàn cục cho đối tượng LoggingMultiProcess.
# -------------------------------------------------------------------------------
_mp_logger = None
def get_mp_logger():
    """
    Trả về instance của LoggingMultiProcess.
    Đảm bảo rằng chỉ có duy nhất 1 instance được tạo.
    """
    global _mp_logger
    if _mp_logger is None:
        _mp_logger = LoggingMultiProcess()
    return _mp_logger

# (Hàm DEBUG_LOG đã được loại bỏ; sử dụng trực tiếp get_mp_logger().logger.info(...) để ghi log.)
