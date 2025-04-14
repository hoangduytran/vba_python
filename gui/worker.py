# worker.py
import os
import logging
import mpp_logger
from mpp_logger import get_mp_logger, LoggingMultiProcess
import win32com.client as win32

# Declare global logger variable.
logger = None

def worker_logging_setup(shared_queue, shared_log_level):
    """
    Configures the worker process logger.
    Clears inherited handlers and installs a QueueHandler that sends log records to the shared queue.
    """
    worker_logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
    worker_logger.handlers.clear()
    new_handler = LoggingMultiProcess.get_worker_handler(shared_queue)
    worker_logger.addHandler(new_handler)
    worker_logger.setLevel(logging.DEBUG)
    # Optionally add filters if needed; in a worker we typically log everything.
    worker_logger.info(f"Worker logging setup complete. GUI filter level = {shared_log_level}")
    
    # Set the global logger for this worker.
    global logger
    logger = worker_logger

def process_excel_file(file_path):
    try:
        print(f"Worker ({os.getpid()}): Bắt đầu xử lý tệp: {file_path}")
        logger.info(f"Worker ({os.getpid()}): Bắt đầu xử lý tệp: {file_path}")
        
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        logger.info(f"Mở file {file_path}")
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        
        macro_file = os.path.abspath("macro_module.bas")
        logger.info(f"Nhập module VBA từ {macro_file} vào {file_path}")
        wb.VBProject.VBComponents.Import(macro_file)
        
        logger.info(f"Chạy macro 'ProcessWorkbook' trên {file_path}")
        excel.Application.Run("ProcessWorkbook")
        
        wb.Save()
        wb.Close()
        excel.Application.Quit()

        result_message = f"Worker ({os.getpid()}): Đã xử lý thành công {file_path}"
        print(result_message)
        logger.info(result_message)
        return result_message

    except Exception as e:
        error_message = f"Worker ({os.getpid()}): Lỗi khi xử lý {file_path}: {str(e)}"
        logger.error(error_message)
        raise Exception(error_message)

def process_batch(batch, progress_queue, log_q):
    count = 0
    for file_path in batch:
        process_excel_file(file_path)
        count += 1
        progress_queue.put(1)
        print(f"Worker ({os.getpid()}): Đã xử lý {file_path} – gửi thông báo cập nhật tiến trình.")
    return count
