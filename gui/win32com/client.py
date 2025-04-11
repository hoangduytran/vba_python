# File: win32com/client.py
# Mô phỏng một phiên bản giả của win32com.client dùng để test ứng dụng.
import time

class FakeVBComponents:
    def Import(self, macro_file):
        print(f"Fake VBComponents: Nhập macro từ '{macro_file}'")
        time.sleep(1)  # Thêm độ trễ 1 giây

class FakeVBProject:
    def __init__(self):
        self.VBComponents = FakeVBComponents()

class FakeWorkbook:
    def __init__(self, path):
        self.path = path
        self.VBProject = FakeVBProject()

    def Save(self):
        print(f"Fake Workbook: Lưu tệp '{self.path}'")
        time.sleep(1)  # Thêm độ trễ 1 giây

    def Close(self):
        print(f"Fake Workbook: Đóng tệp '{self.path}'")
        time.sleep(1)  # Thêm độ trễ 1 giây

class FakeWorkbooks:
    def Open(self, path):
        print(f"Fake Workbooks: Mở tệp '{path}'")
        return FakeWorkbook(path)

class FakeExcel:
    def __init__(self):
        self.Visible = False
        self.Application = self  # Giả lập thuộc tính Application
        self.Workbooks = FakeWorkbooks()

    def Run(self, macro_name):
        print(f"Fake Excel: Chạy macro '{macro_name}'")
        time.sleep(1)  # Thêm độ trễ 1 giây

    def Quit(self):
        print("Fake Excel: Thoát Excel")
        time.sleep(1)  # Thêm độ trễ 1 giây

class FakeCache:
    def EnsureDispatch(self, prog_id):
        print(f"Fake win32com: EnsureDispatch('{prog_id}') được gọi")
        time.sleep(1)  # Thêm độ trễ 1 giây
        return FakeExcel()

# Tạo đối tượng gencache để mô phỏng việc gọi win32com.client.gencache.EnsureDispatch()
gencache = FakeCache()
