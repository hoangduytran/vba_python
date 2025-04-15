# vba_python

# Thi hành VBA 

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)

## Tổng Quan

Thi hành VBA macro cho Excel worksheets

## Nội Dung

- [Cài Đặt](#installation)
- [Sử Dụng](#usage)
- [Đóng Góp](#contributing)
- [Giấy Phép](#license)
- [Liên Lạc](#contact)

## installation (Cài Đặt)

Cách lấy về máy và cài đặt phần phụ thuộc.

```bash
git clone https://github.com/hoangduytran/vba_python.git
cd vba_python
pip install -r requirements.txt
```

## usage (Sử Dụng)

### Giao Diện Đồ Họa

Ứng dụng có giao diện đồ họa và được chia nhỏ như sau:

**Thanh Công Cụ (Taskbar)** nằm bên trái cửa sổ chính (trong `new_gui.py`), chứa các nút:
- **Nạp tập tin VBA**
- **Chọn thư mục Excel**
- **Chạy VBA trên các tập tin Excel**
- **Thoát Ứng dụng**

**Khu Vực Log và Thanh Tiến Trình (bên phải)**  
- **LogText** (định nghĩa trong `logtext.py`):  
  - Chứa **log_level_menu** (drop-down) và **exact_check** (checkbox) ở **phần toolbar trên**. Chúng cho phép người dùng thay đổi mức log cần hiển thị (DEBUG, INFO, WARNING, ERROR, CRITICAL, v.v.) và chọn “duy nhất cấp độ” hay “từ cấp độ này trở lên.” Xem bảng dưới đây: 
  
 | Cấp Độ | Giá trị số biểu thị | Tác Dụng                                                                                  |
|-----------|---------------|------------------------------------------------------------------------------------------|
| NOTSET (Không Đặt)   | 0             | Không có mức cụ thể được thiết lập. Thường dùng nội bộ để bỏ qua các logs, thanh lọc chúng đi.          |
| DEBUG (Điều Tra Lỗi)    | 10            | Chi tiết các thông tin gỡ lỗi; mức thấp nhất, thường dùng cho phát triển và chẩn đoán.     |
| INFO  (Thông Tin)    | 20            | Thông tin chung về hoạt động của ứng dụng; dùng cho việc ghi nhận các sự kiện thông thường. |
| WARNING (Cảnh Báo)  | 30            | Cảnh báo cho biết có thể đã xảy ra vấn đề nhưng ứng dụng vẫn tiếp tục chạy.                |
| ERROR (Lỗi)    | 40            | Ghi nhận lỗi nghiêm trọng khiến một chức năng không thể thực hiện đúng được.              |
| CRITICAL  (Nghiêm Trọng) | 50            | Các lỗi cực kỳ nghiêm trọng, chỉ ra rằng chương trình có thể không còn khả năng chạy.       |

   - Dựa vào nội dung trên, khi phát triển thêm chức năng thì lưu ý dùng các lệnh `logger.debug, logger.info, logger.warning, logger.error, logger.critical` cho hợp lý với mức độ nghiêm trọng mà  dòng nhật ký biểu tả.
 
  - Vùng log (Text) thể hiện nội dung log theo **bộ lọc động** (DynamicLevelFilter).  
  - Thanh công cụ phụ (toolbar) trong LogText (có các nút emoji) giúp thực hiện lưu log, xóa log, sao chép, dán, chọn font, tăng/giảm cỡ chữ, bật/tắt xuống dòng, v.v.
  - Khi lưu log thì có hai lựa chọn, nếu lưu bằng '.log' thì chỉ những gì hiện có trong hộp văn bản log_text sẽ được lưu ra. Tức là những gì mình thanh lọc bằng các lựa chọn trong mức độ và cờ 'duy nhất cấp độ'. Cái này cho phép mình thu gọn các log cần thiết để xem kỹ mà không bị những cái không quan tâm làm rối trí, hòng nhằm điều tra và xử lý tiếp.  
  - Nếu lưu bằng '.json' thì TOÀN BỘ các dòng logs sẽ được lưu ra, nghĩa là nội dung của bản log temp, sẽ được viết ra, dưới định dạng JSON. Cái này hòng cho phép mình viết mã Python, sử dụng JSON phân tích để xem cái nào cần, cái nào không v.v..

- **Thanh Tiến Trình (progress bar)**:  
  Hiển thị tiến độ xử lý tệp Excel (phần trăm) khi chạy VBA trên nhiều file song song.

---

### Phân Tách `gui_actions.py`
Mọi logic như **“save_log,” “load_vba_file,” “select_log_level,” “run_vba_on_all,”**… được cài đặt trong **`gui_actions.py`**.  
Các file **`logtext.py`** và **`new_gui.py`** chỉ gọi `action_list["tên_action"]` để thực thi, đảm bảo tách biệt giữa giao diện và logic.

---

#### Lưu ý:
- Ban đầu, để phục vụ thử nghiệm, có thể có một số dòng mã tạm (debug) trong `main.py` hoặc `new_gui.py` (ví dụ, đặt đường dẫn macro mặc định). Khi vào chạy thực tế, cần **comment** những dòng đó, cho phép người dùng chọn thực sự qua giao diện.

| **Thành phần** | **Mô tả**                                                                                                       |
|----------------|-------------------------------------------------------------------------------------------------------------------|
| **Taskbar (Thanh Công Cụ)** | Bên trái cửa sổ, chứa nút nạp VBA, chọn thư mục, lưu log,...                                                                               |
| **LogText (trong `logtext.py`)** | Chứa vùng hiển thị log, thanh công cụ (toolbar) với nút emoji, **log_level_menu** và **exact_check**                           |
| **Progress Bar (Thanh Tiến Trình)** | Hiển thị phần trăm hoàn thành của quy trình (song song)                                                                          |
 
- Nếu bạn đang giả lập `win32com.client` trong quá trình dev/test, cần gỡ module giả lập trước khi triển khai thật để dùng bản chính thức.  
- Xem trong `main.py` hoặc `new_gui.py`, có thể có các dòng “DEV” gán đường dẫn tạm. Hãy **comment** chúng khi chạy sản xuất để lấy thông tin thực từ giao diện.


Nằm trong thư mục con `gui` cho nên phải `cd gui` rồi chạy mã

`
python3 main.py
`

> **Cần Làm:**  
> Trước khi chạy ứng dụng phiên bản sản xuất, người dùng cần xóa thư mục (module) `win32com/client.py` dùng để mô phỏng (testing mock) win32com đi, trước khi thi hành. Điều này đảm bảo rằng ứng dụng sẽ sử dụng phiên bản chính thức của win32com.client để tương tác với Excel.

> Đồng thời xem trong bản gui.py, thấy các dòng có đề như sau:
```
        # khi chạy thì comment mấy dòng dưới đây lại, 
        # từ đây
        DEBUG_LOG("Bắt đầu chạy VBA trên các tệp Excel.")
        dev_dir = os.environ.get('DEV') or os.getcwd()
        test_dir = os.path.join(dev_dir, 'test_files')
        if not self.excel_directory:
            self.excel_directory = os.path.join(test_dir, 'excel')
        if not self.vba_file:
            self.vba_file = os.path.join(self.excel_directory, 'test_macro.bas')
        globals()["global_vba_file_path"] = self.vba_file
        # cho đến đây, để lấy thông tin thực
```
> và comment chúng lại để lấy thông tin thực từ các điều khiển. Mấy dòng này chỉ là để chạy khi thử nghiệm trong khi xây dựng. Để comment chúng lại, trong vscode, thì chọn mấy dòng này và bấm <kbd>Command</kbd>+<kbd>/</kbd> (macOS), trên Windows chắc là dùng <kbd>Ctrl</kbd>+<kbd>/</kbd>.


##### Lưu ý bản `main.py` và `main_with_callbacks.py`:

Đây chỉ là một sườn bài, bạn phải áp dụng nó cụ thể vào trong trường hợp của bạn. Phải biết lấy Python về máy, cài đặt nó theo nhu cầu. Rồi lại phải biết chạy lệnh

`python3 -m pip install -r requirements.txt`

để lấy các bản thư viện phụ thuộc yêu cầu.

Bên cạnh đó, phải học cách viết các bản `macro` trong ngôn ngữ `VBA` (Visual Basic for Applications: Visual Basic cho Ứng Dụng). Đồng thời phải thông thuộc cấu trúc và các hàm của Excel để có thể sử dụng chúng thành thục.

##### Giải thích chi tiết:

Mô tả mục đích của `module`, các bước xử lý và yêu cầu để chạy mã.
```
process_excel_file(file_path):
```

Mở file Excel bằng win32com.client.

Nhập module VBA từ file `macro_module.bas` và chạy macro `ProcessWorkbook`.

Lưu và đóng file Excel, trả về thông báo kết quả xử lý.

```
process_batch(batch):
```

Hàm nhận vào một nhóm (danh sách các tập tin) và gọi hàm xử lý cho từng tập tin một, trả về một danh sách kết quả.

Khối 
```
if __name__ == "__main__":
```

Sử dụng `glob` (hàm liệt kê danh sách thư mục) để lấy danh sách các tập tin Excel từ một thư mục chỉ định nào đó.

Tính số lõi CPU cần dùng (số lõi trừ đi 2 - để 2 cái còn lại cho hệ điều hành, đừng tham lấy hết để làm trì trệ màn hình) và chia danh sách các tập tin thành các nhóm (batch), đảm bảo phân bổ phần dư (khi chia số tập tin 1000 ra thành 24 nhóm chẳng hạn, sẽ cho 24 nhóm, mỗi nhóm là 41 tập tin, và số dư là 16, phải đảm bảo 16 tập tin còn lại sẽ được phân bổ).

Sử dụng `multiprocessing.Pool` (bể đa trình xử lý) để xử lý các nhóm song song và in kết quả xử lý từng tập tin một.

Mã nguồn trên giúp tự động hoá việc xử lý nhiều văn bản Excel đồng thời, tối ưu tài nguyên CPU và dễ dàng bảo trì, mở rộng nếu cần.



## contributing (Đóng Góp)

Nếu biết về lập trình và `forking repository` (tạo một bản sao cho mình tại kho riêng của mình) và lấy về máy bằng lệnh `git clone`, rồi chỉnh sửa và tạo một `pull-request` thì mình có thể giúp bạn hội nhập các thay đổi để những người khác cũng dùng chung được các thay đổi của bạn, một điều rất đáng hoan nghênh.

## contact (Liên Lạc)
e-mail: hoangduytran1960@gmail.com
Facebook: https://www.facebook.com/hoangduy.tran

