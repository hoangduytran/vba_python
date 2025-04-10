import os
import logging
import mpp_logger
from mpp_logger import DEBUG_LOG, LoggingMultiProcess, SafeQueueHandler, get_mp_logger

class DummyLogging:
    def __init__(self, logger, debug_flag):
        self.logger = logger
        self.debug_flag = debug_flag  # This is a Manager.Value

    def DEBUG_LOG(self, msg):
        # Only log if the shared debug flag is True
        if self.debug_flag.value:
            self.logger.debug(msg)

def worker_logging_setup(shared_queue, shared_is_debug):
    """
    Initialize logging for the worker process without creating a new Manager.
    This function configures the worker's logger (named LoggingMultiProcess.MAIN_LOGGER)
    with a SafeQueueHandler using the shared_queue, then overrides the module's _mp_logger
    and DEBUG_LOG so that calls to DEBUG_LOG() in the worker use the worker's logger.
    """
    # Retrieve (or create) the logger for this worker.
    worker_logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
    worker_logger.handlers.clear()
    
    # Create a new SafeQueueHandler using the shared queue.
    new_handler = LoggingMultiProcess.get_worker_handler(shared_queue)
    worker_logger.addHandler(new_handler)
    worker_logger.setLevel(logging.DEBUG)
        
    # Override the module-level _mp_logger with our dummy.
    mpp_logger._mp_logger = DummyLogging(worker_logger, shared_is_debug)
    
    print(f"Worker logging setup running, using queue: {shared_queue}")

def process_excel_file(file_path):
    """
    Processes a single Excel file.
    The worker’s logger (previously set up in worker_logging_setup()) is expected to use a SafeQueueHandler
    so that messages sent via DEBUG_LOG() are forwarded through the shared queue.

    Note: Ensure that the worker initializer has reconfigured the global _mp_logger so that DEBUG_LOG() is fully effective.
    """
    try:
        # Also print directly for testing purposes
        print(f"PRINTING Worker ({os.getpid()}): Starting processing of file: {file_path}")        
        # Log from the worker using DEBUG_LOG
        DEBUG_LOG(f"Worker ({os.getpid()}): Starting processing of file: {file_path}")
        
        # Simulate processing steps. For an actual implementation,
        # uncomment and replace the code below with real Excel automation.
        # For example:
        # excel = win32.gencache.EnsureDispatch("Excel.Application")
        # excel.Visible = False
        # wb = excel.Workbooks.Open(os.path.abspath(file_path))
        # macro_file = globals().get("global_vba_file_path", os.path.abspath("macro_module.bas"))
        # DEBUG_LOG(f"Worker ({os.getpid()}): Importing VBA module from: {macro_file}")
        # wb.VBProject.VBComponents.Import(macro_file)
        # DEBUG_LOG(f"Worker ({os.getpid()}): Running macro 'ProcessWorkbook' on: {file_path}")
        # excel.Application.Run("ProcessWorkbook")
        # wb.Save()
        # wb.Close()
        # excel.Application.Quit()

        # Log success
        result_message = f"Worker ({os.getpid()}): Successfully processed {file_path}"
        print(f"PRINTING result_message {result_message}")
        DEBUG_LOG(result_message)        
        return result_message

    except Exception as e:
        error_message = f"Worker ({os.getpid()}): Error processing {file_path}: {str(e)}"
        DEBUG_LOG(error_message)
        # Raising exception ensures that the error is propagated up.
        raise Exception(error_message)

def process_batch(batch, progress_queue, log_q):
    """
    Xử lý một nhóm các file Excel (batch).

    Thao tác:
      - Gọi worker_logging_setup để cấu hình logging cho tiến trình con, sử dụng hàng đợi log chia sẻ.
      - Duyệt qua các file trong batch, xử lý từng file và cập nhật tiến trình.
      - Trả về tổng số file đã xử lý.
    """
    count = 0
    for file_path in batch:
        process_excel_file(file_path)
        count += 1
        progress_queue.put(1)
    return count
