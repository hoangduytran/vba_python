# gv.py
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}
FONT_BASIC = ("Arial", 15, "normal")
font_options = ["Arial", "Courier New", "Times New Roman", "Verdana", "Tahoma"]

def create_log_record(record, with_diacritics=False):
    """
    Creates a log record dictionary from a logging record.

    Args:
        record: A logging record object.
        with_diacritics (bool): If True, return keys with Vietnamese diacritics;
                                if False, return keys without diacritics.

    Returns:
        dict: A dictionary with the formatted log information.
    """
    if with_diacritics:
        return {
            "thời điểm": record.asctime,
            "tên tiến trình": record.processName,
            "tên tệp tin": record.pathname,
            "hàm": f"{record.funcName}()",
            "số dòng": record.lineno,
            "cấp độ": record.levelname,
            "thông điệp": record.message
        }
    else:
        return {
            "thoi diem": record.asctime,
            "ten tien trinh": record.processName,
            "ten tep tin": record.pathname,
            "ham": f"{record.funcName}()",
            "so dong": record.lineno,
            "cap do": record.levelname,
            "thong diep": record.message
        }


class Gvar:
    # Main window and logging
    root = None               # Main Tk instance (set in gui.py)
    mp_logging = None         # Reference to the logging object from mp_logging
    logger = None

    # Widget references (buttons)
    button_save_log = None
    button_load_vba_file = None
    button_load_excel_directory = None
    button_run_macro = None
    button_exit_app = None

    # Display widgets
    log_text_widget = None    # Reference to the log display (LogText widget)
    progress_bar = None       # The progress bar widget
    progress_label = None     # The label showing the percentage progress

    # Application state variables
    excel_directory = None    # Directory containing Excel files
    vba_file = None           # Path to the VBA file
    total_files = 0           # Total Excel files to process
    progress_count = 0        # Count of processed files
    progress_queue = None     # Queue for processing progress
    log_level_var = None
    is_exact_var = None