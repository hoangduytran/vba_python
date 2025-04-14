# mpp_logger.py
import logging
import sys
import tempfile
from multiprocessing import Manager
from logging.handlers import QueueHandler, QueueListener
import json

class ExactLevelFilter(logging.Filter):
    def __init__(self, level):
        super().__init__()
        self.level = level

    def filter(self, record):
        return record.levelno == self.level

# --- Formatter that produces JSON output (for storage) ---
class JsonFormatter(logging.Formatter):
    def format(self, record):
        record.message = record.getMessage()
        # Always compute timestamp.
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

# --- Formatter for human-friendly output ---
class PrettyFormatter(logging.Formatter):
    def format(self, record):
        record.message = record.getMessage()
        record.asctime = self.formatTime(record, self.datefmt)
        s = (
            f"time stamp: {record.asctime}\n"
            f"process name: {record.processName}\n"
            f"filename: {record.pathname}\n"
            f"function: hàm:{record.funcName}()\n"
            f"line number: dòng số:{record.lineno}\n"
            f"level: {record.levelname}\n"
            f"message: {record.message}\n"
        )
        return s

# --- A custom handler that saves each log record to an internal list ---
class MemoryLogHandler(logging.Handler):
    def __init__(self, log_store):
        super().__init__()
        self.log_store = log_store

    def emit(self, record):
        try:
            # Format the record as JSON then parse it into a dict.
            formatter = JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
            json_record = formatter.format(record)
            self.log_store.append(json.loads(json_record))
        except Exception:
            self.handleError(record)

# --- Main multiprocess logging wrapper ---
class LoggingMultiProcess:
    MAIN_LOGGER = "main_logger"

    def __init__(self):
        # Create a Manager and shared queue.
        self.manager = Manager()
        self.queue = self.manager.Queue()
        # Shared log level used for GUI filtering; default is INFO.
        self.log_level = self.manager.Value('i', logging.INFO)

        # Create a temporary file for file-logging.
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".log")
        self.log_temp_file_path = temp_file.name
        temp_file.close()

        # Create formatters.
        self.pretty_formatter = PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")
        self.json_formatter = JsonFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z")

        # Create handlers that display logs in a human-friendly way.
        terminal_handler = logging.StreamHandler(sys.stdout)
        terminal_handler.setFormatter(self.pretty_formatter)

        # Change file_handler to use JSON formatter so that each log line is a JSON dictionary.
        file_handler = logging.FileHandler(self.log_temp_file_path, mode="w", encoding="utf-8")
        file_handler.setFormatter(self.json_formatter)

        # Set up a QueueListener that listens on the shared queue.
        self.listener = QueueListener(self.queue, terminal_handler, file_handler)
        self.listener.start()
        print("Temporary log file:", self.log_temp_file_path)

        # Internal list to store JSON-formatted logs.
        self.log_store = []

        # Create the global logger.
        self.logger = logging.getLogger(LoggingMultiProcess.MAIN_LOGGER)
        self.logger.setLevel(logging.DEBUG)
        # Attach a QueueHandler so logs are sent to the shared queue.
        safe_handler = self.get_worker_handler(self.queue)
        self.logger.addHandler(safe_handler)
        # Attach the MemoryLogHandler to record JSON logs internally.
        memory_handler = MemoryLogHandler(self.log_store)
        self.logger.addHandler(memory_handler)
        # (The terminal and file output is handled by the QueueListener.)
        # The GUI handler will be attached later by the GUI module.

    def select_log_level(self, new_level):
        self.log_level.value = new_level
        # Update the filters on logger handlers (like the QueueHandler and MemoryLogHandler)
        for handler in self.logger.handlers:
            # Remove any previous ExactLevelFilter instances.
            handler.filters = [f for f in handler.filters if not isinstance(f, ExactLevelFilter)]
            # Add the new filter.
            handler.addFilter(ExactLevelFilter(new_level))
        # Also update the filters on handlers used by the QueueListener (terminal and file output)
        for handler in self.listener.handlers:
            handler.filters = [f for f in handler.filters if not isinstance(f, ExactLevelFilter)]
            handler.addFilter(ExactLevelFilter(new_level))

    @classmethod
    def get_worker_handler(cls, queue):
        # Use PrettyFormatter for display in all output handlers.
        handler = QueueHandler(queue)
        handler.setFormatter(PrettyFormatter(datefmt="%Y-%m-%dT%H:%M:%S%z"))
        return handler

    def shutdown(self):
        if self.listener:
            try:
                self.listener.stop()
            except BrokenPipeError:
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

# (Remove DEBUG_LOG helper; use get_mp_logger().logger.info(…) directly.)
