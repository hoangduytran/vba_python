# main.py
from mpp_logger import get_mp_logger
from gui import MainWindow

# Global logger variable; will be assigned by main.
logger = None

def main():
    global logger
    mp_logger = get_mp_logger()
    logger = mp_logger.logger
    # Pass mp_logger to MainWindow (in gui.py, the global logger will be set).
    window = MainWindow(mp_logger)
    window.protocol("WM_DELETE_WINDOW", window.exit_app)
    window.mainloop()
    mp_logger.shutdown()

if __name__ == '__main__':
    main()
