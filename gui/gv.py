# gv.py
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}
FONT_BASIC = ("Arial", 15, "normal")
font_options = ["Arial", "Courier New", "Times New Roman", "Verdana", "Tahoma"]

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