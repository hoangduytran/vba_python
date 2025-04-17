# Định nghĩa kiểu dáng chung cho các widget giao diện
COMMON_WIDGET_STYLE = {"font": ("Arial", 18, "bold"), "width": 25, "height": 3}
FONT_BASIC = ("Arial", 15, "normal")
font_options = ["Arial", "Courier New", "Times New Roman", "Verdana", "Tahoma"]

def create_log_record(record, with_diacritics=False):
    """
    Tạo một bản ghi log dưới dạng từ điển từ một đối tượng ghi log.

    Tham số:
        record: Đối tượng bản ghi log.
        with_diacritics (bool): Nếu là True, trả về các khóa có dấu tiếng Việt;
                                nếu là False, trả về các khóa không dấu.

    Trả về:
        dict: Một từ điển chứa thông tin log đã định dạng.
    """
    if with_diacritics:
        return {
            "thời điểm": record.asctime,
            "tên tiến trình": record.processName,
            "tên tệp tin": record.pathname,
            "tên hàm": f"{record.funcName}()",
            "dòng số": record.lineno,
            "cấp độ": record.levelname,
            "thông điệp": record.message
        }
    else:
        return {
            "thoi diem": record.asctime,
            "ten tien trinh": record.processName,
            "ten tep tin": record.pathname,
            "ten ham": f"{record.funcName}()",
            "dong so": record.lineno,
            "cap do": record.levelname,
            "thong diep": record.message
        }

# Lớp dùng để lưu các biến toàn cục trong ứng dụng
class Gvar:
    # Cửa sổ chính và logging
    root = None               # Tham chiếu đến Tk chính (thiết lập trong gui.py)
    mp_logging = None         # Tham chiếu đến đối tượng logging từ mp_logging
    logger = None             # Logger chính

    # Tham chiếu đến các nút giao diện
    button_save_log = None
    button_load_vba_file = None
    button_load_excel_directory = None
    button_run_macro = None
    button_exit_app = None

    # Các widget hiển thị
    log_text_widget = None    # Tham chiếu đến widget hiển thị log (LogText)
    progress_bar = None       # Thanh tiến trình
    progress_label = None     # Nhãn hiển thị phần trăm tiến trình

    # Các biến trạng thái của ứng dụng
    excel_directory = None    # Thư mục chứa các tệp Excel
    vba_file = None           # Đường dẫn đến tệp VBA
    total_files = 0           # Tổng số tệp Excel cần xử lý
    progress_count = 0        # Số lượng tệp đã xử lý
    progress_queue = None     # Hàng đợi xử lý tiến trình
    log_level_var = None      # Biến lưu cấp độ log
    is_exact_var = None       # Biến lưu trạng thái tìm kiếm chính xác hay không
