from mpp_logger import get_mp_logger, DEBUG_LOG
from gui import MainWindow  # module GUI của bạn

def main():
    # Tạo một instance của LoggingMultiProcess. Điều này thiết lập hàng đợi chia sẻ, listener và logger.
    # Giờ đây, logger chính ("main_logger") có sẵn dưới dạng mp_logging.logger.
    mp_logger = get_mp_logger()
    # Truyền toàn bộ instance mp_logging vào MainWindow.
    window = MainWindow(mp_logger)
    window.protocol("WM_DELETE_WINDOW", window.exit_app)
    window.mainloop()
    #
    # Tắt hệ thống logging sau khi GUI kết thúc.
    mp_logger.shutdown()

if __name__ == '__main__':
    main()
