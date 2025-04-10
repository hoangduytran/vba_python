import logging
from mpp_logger import DEBUG_LOG, LoggingMultiProcess, SafeQueueHandler, get_mp_logger

def worker_logging_setup(queue):
    """
    Thiết lập cấu hình logging trong tiến trình con.

    Thao tác:
      - In ra thông báo debug về queue đang được sử dụng.
      - Lấy logger toàn cục với tên được định nghĩa tĩnh trong LoggingMultiProcess.
      - Đặt mức log là DEBUG, loại bỏ mọi handler cũ và tắt propagate.
      - Gắn một SafeQueueHandler mới (được lấy từ phương thức get_worker_handler)
        để gửi log vào hàng đợi chia sẻ.
    """
    print("Worker logging setup running, using queue:", queue)
    worker_logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
    worker_logger.setLevel(logging.DEBUG)
    for h in worker_logger.handlers[:]:
        worker_logger.removeHandler(h)
    worker_logger.propagate = False
    handler = LoggingMultiProcess.get_worker_handler(queue)
    worker_logger.addHandler(handler)

def process_excel_file(file_path):
    """
    Xử lý một file Excel.

    Thao tác:
      - Lấy logger đã được cấu hình bởi worker_logging_setup.
      - Ghi log debug cho các bước xử lý (ví dụ: khởi tạo Excel, mở file).
      - Trả về kết quả xử lý file.
      - Nếu gặp lỗi, ghi log lỗi và ném ngoại lệ.
    """
    # Lấy logger được cấu hình bởi worker_logging_setup
    worker_logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
    try:
        print("Processing file:", file_path)  # In thông báo ra màn hình (để debug).
        worker_logger.debug(f"Khởi tạo Excel cho {file_path}")
        worker_logger.debug(f"Mở file {file_path}")
        result = f"Đã xử lý: {file_path}"
        worker_logger.debug(result)
        return result
    except Exception as e:
        worker_logger.error(f"Lỗi khi xử lý {file_path}: {e}")
        raise

def process_batch(batch, progress_queue, log_q):
    """
    Xử lý một nhóm các file Excel (batch).

    Thao tác:
      - Gọi worker_logging_setup để cấu hình logging cho tiến trình con, sử dụng hàng đợi log chia sẻ.
      - Duyệt qua các file trong batch, xử lý từng file và cập nhật tiến trình.
      - Trả về tổng số file đã xử lý.
    """
    worker_logging_setup(log_q)
    count = 0
    for file_path in batch:
        process_excel_file(file_path)
        count += 1
        progress_queue.put(1)
    return count
