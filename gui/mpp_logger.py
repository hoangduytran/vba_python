# mpp_logger.py
import logging
import sys
import tempfile
from multiprocessing import Manager
from logging.handlers import QueueHandler, QueueListener
import json

# --- Custom JSON Formatter ---
class JsonFormatter(logging.Formatter):
    def format(self, record):
        # Compute the message.
        record.message = record.getMessage()
        # Always compute asctime using the configured datefmt.
        record.asctime = self.formatTime(record, self.datefmt)
        log_record = {
            "time stamp": record.asctime,
            "process name": record.processName,
            "filename": record.pathname,
            "function": f"hàm:{record.funcName}()",
            "line number": f"dòng số:{record.lineno}",
            "level": record.levelname,
            "message": record.message
        }
        return json.dumps(log_record)

# --- Filter for the GUI handler ---
class GuiLogFilter(logging.Filter):
    def __init__(self, allowed_level):
        super().__init__()
        self.allowed_level = allowed_level

    def filter(self, record):
        return record.levelno >= self.allowed_level

# --- Main multiprocess logging class ---
class LoggingMultiProcess:
    MAIN_LOGGER = "main_logger"

    def __init__(self):
        # Create a Manager and shared queue.
        self.manager = Manager()
        self.queue = self.manager.Queue()
        # Instead of a Boolean, use a shared integer for the log level.
        # Default GUI log level is INFO.
        self.log_level = self.manager.Value('i', logging.INFO)

        # Create a temporary file for logging.
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".log")
        self.log_temp_file_path = temp_file.name
        temp_file.close()

        # Create a JSON formatter with a fixed date format.
        self.json_formatter = JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")

        # Set up two handlers: one for the terminal and one for the file.
        terminal_handler = logging.StreamHandler(sys.stdout)
        terminal_handler.setFormatter(self.json_formatter)

        file_handler = logging.FileHandler(self.log_temp_file_path, mode="w", encoding="utf-8")
        file_handler.setFormatter(self.json_formatter)

        # Create and start a QueueListener that listens on the shared queue.
        self.listener = QueueListener(self.queue, terminal_handler, file_handler)
        self.listener.start()
        print("Temporary log file:", self.log_temp_file_path)

        # Create the global logger.
        self.logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
        self.logger.setLevel(logging.DEBUG)  # Capture all; handlers and GUI filter decide what to show.
        # Attach a QueueHandler so that log records are sent to the shared queue.
        safe_handler = self.get_worker_handler(self.queue)
        self.logger.addHandler(safe_handler)
        # (Do not filter here; the file and terminal handlers get all messages.)

        # Note: The GUI text handler will be added later in the GUI module.

    def select_log_level(self, new_level):
        """
        Update the shared GUI log level. new_level should be one of
        logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL.
        """
        self.log_level.value = new_level
        # Update any GUI handler(s) attached to self.logger.
        for handler in self.logger.handlers:
            if getattr(handler, "is_gui_handler", False):
                handler.filters = []
                handler.addFilter(GuiLogFilter(new_level))

    @classmethod
    def get_worker_handler(cls, queue):
        """
        Returns a new QueueHandler that sends log records to the given queue,
        with the JSON formatter applied.
        """
        handler = QueueHandler(queue)
        handler.setFormatter(JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z"))
        return handler

    def shutdown(self):
        """
        Stops the QueueListener and shuts down the Manager.
        """
        if self.listener:
            try:
                self.listener.stop()
            except BrokenPipeError:
                # Ignore BrokenPipeError on shutdown as it is benign in this context.
                pass
            except Exception as e:
                print("Error stopping QueueListener:", e)
        if self.manager:
            self.manager.shutdown()

# Global singleton accessor.
_mp_logger = None

def get_mp_logger():
    global _mp_logger
    if _mp_logger is None:
        _mp_logger = LoggingMultiProcess()
    return _mp_logger

# (DEBUG_LOG is removed; use get_mp_logger().logger.info(…) etc directly.)
