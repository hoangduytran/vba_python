"""
Module: Excel Batch Processor
Mô tả:
    - Lấy tất cả các file Excel (.xlsx) từ một thư mục cụ thể sử dụng thư viện glob.
    - Chia danh sách file thành các nhóm (batch) dựa trên số lõi CPU trừ đi 2 (ít nhất 1 lõi).
    - Mỗi tiến trình sẽ xử lý một batch file Excel: mở file, nhập module VBA từ file macro_module.bas,
      chạy macro 'ProcessWorkbook', lưu và đóng file.
    - In ra kết quả xử lý của từng file.
Yêu cầu:
    - Chạy trên Windows có cài đặt Microsoft Excel.
    - Cài đặt thư viện pywin32 (pip install pywin32).
    - Bật tùy chọn "Trust access to the VBA project object model" trong Excel.


Giải thích chi tiết:
Module Docstring:
Mô tả mục đích của module, các bước xử lý và yêu cầu để chạy mã.

process_excel_file(file_path):

Mở file Excel bằng win32com.client.

Nhập module VBA từ file macro_module.bas và chạy macro ProcessWorkbook.

Lưu và đóng file Excel, trả về thông báo kết quả xử lý.

process_batch(batch):

Hàm nhận vào một batch (danh sách file) và gọi hàm xử lý cho từng file, trả về danh sách kết quả.

Khối if name == "main":

Sử dụng glob để lấy danh sách file Excel từ thư mục chỉ định.

Tính số lõi CPU cần dùng (số lõi trừ đi 2) và chia danh sách file thành các batch, đảm bảo phân bổ phần dư.

Sử dụng multiprocessing.Pool để xử lý các batch song song và in kết quả xử lý từng file.

Mã nguồn trên giúp tự động hoá việc xử lý nhiều file Excel đồng thời, tối ưu tài nguyên CPU và dễ dàng bảo trì, mở rộng nếu cần.
"""

import os
import glob
import win32com.client as win32
from multiprocessing import Pool


def process_excel_file(file_path):
    """
    Mở file Excel, nhập module VBA từ file macro_module.bas, chạy macro 'ProcessWorkbook',
    lưu và đóng file.

    Parameters:
        file_path (str): Đường dẫn đến file Excel cần xử lý.

    Returns:
        str: Thông báo kết quả xử lý cho file.
    """
    try:
        # Khởi tạo instance Excel và ẩn giao diện người dùng
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False

        # Mở file Excel theo đường dẫn tuyệt đối
        wb = excel.Workbooks.Open(os.path.abspath(file_path))

        # Xác định đường dẫn đến file macro VBA (cần đảm bảo macro_module.bas tồn tại)
        macro_file = os.path.abspath("macro_module.bas")

        # Nhập module VBA vào dự án của workbook
        wb.VBProject.VBComponents.Import(macro_file)

        # Chạy macro 'ProcessWorkbook' được định nghĩa trong file macro_module.bas
        excel.Application.Run("ProcessWorkbook")

        # Lưu thay đổi, đóng workbook và thoát Excel
        wb.Save()
        wb.Close()
        excel.Application.Quit()

        return f"Processed: {file_path}"
    except Exception as e:
        return f"Error processing {file_path}: {e}"


def process_batch(batch):
    """
    Xử lý một batch các file Excel.

    Parameters:
        batch (list): Danh sách đường dẫn file Excel cần xử lý trong batch.

    Returns:
        list: Danh sách thông báo kết quả xử lý cho từng file trong batch.
    """
    results = []
    for file_path in batch:
        result = process_excel_file(file_path)
        results.append(result)
    return results


if __name__ == "__main__":
    # Chỉ định thư mục chứa các file Excel cần xử lý
    directory = r"C:\path\to\excel\files"  # Thay đổi đường dẫn theo nhu cầu của bạn, đường dẫn đến thư mục chứa các văn bản Excel

    # Lấy danh sách tất cả các file .xlsx trong thư mục được chỉ định
    excel_files = glob.glob(os.path.join(directory, "*.xlsx"))

    if not excel_files:
        print("No Excel files found in the specified directory.")
    else:
        # Xác định số lõi CPU cần dùng: tổng số lõi trừ đi 2 (ít nhất 1 lõi)
        num_cores = os.cpu_count() - 2
        if num_cores < 1:
            num_cores = 1

        total_files = len(excel_files)

        # Chia danh sách file thành các batch, mỗi batch xử lý bởi một tiến trình riêng
        batches = []
        batch_size = total_files // num_cores  # Số file tối thiểu trên mỗi batch
        remainder = total_files % num_cores  # Phần dư file sau khi chia đều

        start = 0
        for i in range(num_cores):
            # Nếu có dư, thêm 1 file vào các batch đầu tiên
            end = start + batch_size + (1 if i < remainder else 0)
            batches.append(excel_files[start:end])
            start = end

        # Xử lý các batch đồng thời bằng multiprocessing Pool
        with Pool(processes=num_cores) as pool:
            batch_results = pool.map(process_batch, batches)

        # In kết quả ra màn hình (flatten danh sách kết quả)
        for batch in batch_results:
            for result in batch:
                print(result)
