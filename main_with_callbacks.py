import os
import glob
import win32com.client as win32
from multiprocessing import Pool
import inspect

# Cờ toàn cục để bật/tắt log gỡ lỗi (đặt True để bật log gỡ lỗi)
is_debug = True

# Danh sách toàn cục để lưu kết quả từ tất cả các batch
global_results = []

def DEBUG_LOG(message):
    """
    Ghi log thông báo gỡ lỗi kèm theo tên của hàm gọi.

    Tham số:
        message (str): Thông báo cần ghi log.
    """
    if is_debug:
        # Lấy tên hàm của caller (một cấp trên)
        caller = inspect.stack()[1].function
        print(f"[DEBUG] [{caller}] {message}")

def process_excel_file(file_path):
    """
    Xử lý một file Excel bằng cách mở file, nhập module VBA từ file macro_module.bas,
    chạy macro 'ProcessWorkbook', lưu và đóng file.

    Tham số:
        file_path (str): Đường dẫn đến file Excel cần xử lý.

    Trả về:
        str: Thông báo cho biết file đã được xử lý thành công.

    Ném ra:
        Exception: Nếu có bất kỳ lỗi nào xảy ra trong quá trình xử lý.
    """
    try:
        DEBUG_LOG(f"Khởi tạo Excel cho {file_path}")
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False

        DEBUG_LOG(f"Mở file {file_path}")
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        
        # Xác định đường dẫn đến file VBA macro (macro_module.bas)
        macro_file = os.path.abspath("macro_module.bas")
        DEBUG_LOG(f"Nhập module VBA từ {macro_file} vào {file_path}")
        wb.VBProject.VBComponents.Import(macro_file)
        
        DEBUG_LOG(f"Chạy macro 'ProcessWorkbook' trên {file_path}")
        excel.Application.Run("ProcessWorkbook")
        
        wb.Save()
        wb.Close()
        excel.Application.Quit()
        
        result = f"Đã xử lý: {file_path}"
        DEBUG_LOG(result)
        return result
    except Exception as e:
        error_msg = f"Lỗi khi xử lý {file_path}: {e}"
        DEBUG_LOG(error_msg)
        raise Exception(error_msg)

def process_batch_callback(batch, success_callback, error_callback):
    """
    Xử lý một nhóm (batch) các file Excel sử dụng các hàm callback để xử lý thành công và lỗi.

    Mỗi file được xử lý riêng lẻ. Nếu xử lý thành công, success_callback được gọi.
    Nếu xảy ra lỗi, error_callback được gọi. Kết quả được lưu vào danh sách toàn cục,
    và hàm sẽ in ra số lượng thay đổi thành công trong nhóm.

    Tham số:
        batch (list): Danh sách các đường dẫn file cần xử lý trong nhóm.
        success_callback (function): Hàm callback được gọi khi xử lý thành công.
                                     (Chữ ký: success_callback(file_path, result))
        error_callback (function): Hàm callback được gọi khi xảy ra lỗi.
                                   (Chữ ký: error_callback(file_path, error))

    Trả về:
        int: Số file được xử lý thành công (số lượng thay đổi).
    """
    batch_change_count = 0
    DEBUG_LOG(f"Bắt đầu xử lý batch với {len(batch)} file")
    
    for file_path in batch:
        DEBUG_LOG(f"Đang xử lý file: {file_path}")
        try:
            result = process_excel_file(file_path)
            batch_change_count += 1
            global_results.append(result)
            # Gọi hàm callback khi xử lý thành công file
            success_callback(file_path, result)
        except Exception as e:
            # Gọi hàm callback khi có lỗi xảy ra trong file
            error_callback(file_path, e)
            global_results.append(f"Lỗi: {file_path}: {e}")
    
    DEBUG_LOG(f"Hoàn thành xử lý batch. Tổng số thay đổi thành công: {batch_change_count}")
    print(f"Batch hoàn thành: {batch_change_count} thay đổi đã được xử lý.")
    return batch_change_count

# Các hàm callback mẫu
def success_callback(file_path, result):
    """
    Hàm callback xử lý thành công file.

    Tham số:
        file_path (str): Đường dẫn của file đã xử lý.
        result (str): Thông báo kết quả từ quá trình xử lý.
    """
    DEBUG_LOG(f"Callback thành công cho {file_path} với kết quả: {result}")
    print(f"THÀNH CÔNG: {file_path} đã được xử lý.")

def error_callback(file_path, error):
    """
    Hàm callback xử lý lỗi khi xử lý file.

    Tham số:
        file_path (str): Đường dẫn của file xảy ra lỗi.
        error (Exception): Lỗi (exception) phát sinh trong quá trình xử lý.
    """
    DEBUG_LOG(f"Callback lỗi cho {file_path} với lỗi: {error}")
    print(f"LỖI: {file_path} gặp lỗi: {error}")

# Ví dụ sử dụng (cho mục đích minh họa)
if __name__ == "__main__":
    # Chỉ định thư mục chứa các file Excel cần xử lý
    directory = r"C:\path\to\excel\files"  # Cập nhật đường dẫn theo nhu cầu của bạn
    # Sử dụng glob để lấy tất cả các file Excel có đuôi .xlsx trong thư mục chỉ định
    excel_files = glob.glob(os.path.join(directory, "*.xlsx"))
    
    if not excel_files:
        print("Không tìm thấy file Excel nào trong thư mục chỉ định.")
    else:
        # Ví dụ: xử lý tất cả các file trong một batch.
        # Bạn có thể chia thành nhiều batch tùy nhu cầu.
        changes = process_batch_callback(excel_files, success_callback, error_callback)
        print(f"Tổng số thay đổi trong batch: {changes}")
