# main.py
from mpp_logger import get_mp_logger
from gui import MainWindow

# Declare global logger variable.
logger = None

def main():
    global logger
    # Get the multiprocess logging instance.
    mp_logger = get_mp_logger()
    # Assign its logger to the global variable.
    logger = mp_logger.logger
    # Pass the multiprocess logger to the GUI.
    window = MainWindow(mp_logger)
    window.protocol("WM_DELETE_WINDOW", window.exit_app)
    window.mainloop()
    mp_logger.shutdown()

if __name__ == '__main__':
    main()
