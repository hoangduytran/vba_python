# main.py

# Nhập hàm get_mp_logger() từ module mpp_logger, hàm này trả về đối tượng quản lý logging đa tiến trình.
from mpp_logger import get_mp_logger

# Nhập lớp MainWindow từ module gui, đây là cửa sổ giao diện chính của ứng dụng.
from gui import MainWindow

# Khai báo biến toàn cục 'logger'; biến này sẽ được gán bằng logger chính sau khi khởi tạo trong hàm main.
logger = None

def main():
    global logger  # Sử dụng biến toàn cục 'logger' để có thể gán giá trị bên trong hàm main.
    
    # Lấy instance của LoggingMultiProcess qua hàm get_mp_logger(), đối tượng này chứa cấu hình logging toàn cục,
    # hàng đợi logging chia sẻ và các handler liên quan đến logging cho các tiến trình.
    mp_logger = get_mp_logger()
    
    # Gán logger toàn cục bằng logger của instance mp_logger.
    # Sau dòng này, biến logger sẽ được sử dụng ở các nơi khác trong ứng dụng để gọi các hàm logger.info, logger.error, ...
    logger = mp_logger.logger
    
    # Tạo cửa sổ giao diện chính của ứng dụng, truyền đối tượng mp_logger vào MainWindow
    # để giao diện có thể sử dụng hệ thống logging đa tiến trình đã được cấu hình.
    window = MainWindow(mp_logger)
    
    # Thiết lập sự kiện đóng cửa sổ: khi người dùng tắt cửa sổ (sự kiện "WM_DELETE_WINDOW"),
    # hàm exit_app của MainWindow sẽ được gọi để thực hiện các tác vụ dọn dẹp (shutdown hệ thống logging, kết thúc vòng lặp, …)
    window.protocol("WM_DELETE_WINDOW", window.exit_app)
    
    # Khởi chạy vòng lặp chính của Tkinter để giao diện hiển thị và xử lý sự kiện.
    window.mainloop()
    
    # Sau khi vòng lặp kết thúc (ứng dụng đã đóng), gọi phương thức shutdown() của mp_logger để tắt và giải phóng
    # các tài nguyên liên quan đến logging (ví dụ: dừng QueueListener, tắt Manager, …).
    mp_logger.shutdown()

# Kiểm tra nếu file được chạy trực tiếp (không phải được import vào module khác), sau đó gọi hàm main().
if __name__ == '__main__':
    main()

