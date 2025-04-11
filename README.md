# vba_python

# Thi hành VBA 

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)

## Tổng Quan

Thi hành VBA macro cho Excel worksheets

## Table of Contents

- [Cài Đặt](#installation)
- [Sử Dụng](#usage)
- [Đóng Góp](#contributing)
- [Giấy Phép](#license)
- [Liên Lạc](#contact)

## installation

Cách lấy về máy và cài đặt phần phụ thuộc.

```bash
git clone https://github.com/hoangduytran/vba_python.git
cd vba_python
pip install -r requirements.txt
```

## usage (Sử Dụng)

### Giao Diện Đồ Họa

Nằm trong thư mục con `gui` cho nên phải `cd gui` rồi chạy mã

`
python3 main.py
`

Ứng dụng bao gồm một giao diện đồ họa thân thiện và hiện đại, được thiết kế để hỗ trợ người dùng thao tác một cách dễ dàng. Các thành phần chính bao gồm:

1. **Thanh Công Cụ (Taskbar):**
   - **Các nút điều khiển:**
     - **Bật/Tắt hiển thị lỗi:** Chức năng này cho phép bật/tắt các dòng `DEBUG_LOG` gắn trong mã hiển thị trên 3 phương tiện:
        1. Văn bản tạm thời của hệ thống (temp)
        2. Cổng thiết bị cuối, nơi thi hành `python3 main.py`
        3. Cửa sổ log ở bên phải, nơi có thể sử dụng để quan sát, lưu văn bản log vào một tập tin và đọc, quan sát lỗi, kết quả chạy. Cái này cho phép bạn sử dụng mã `DEBUG_LOG` để liệt kê tiến trình cùng các dòng điều tra lỗi trong khi phát triển phần mềm, viết thêm những chức năng mới cho bản thân. 
     - **Lưu Log vào tập tin:** Lưu lại nội dung log hiện tại vào một tập tin để lưu trữ hoặc chia sẻ.
     - **Tải tệp VBA:** Cho phép người dùng chọn tệp chứa VBA macro để nhập vào Excel.
     - **Tải thư mục Excel:** Lựa chọn thư mục chứa các tệp Excel cần xử lý.
     - **Chạy VBA trên tất cả các tệp Excel:** Khởi chạy chế độ xử lý VBA cho tất cả các tập tin Excel trong thư mục đã chọn, sử dụng đa tiến trình nhằm tối ưu hoá tài nguyên CPU.
     - **Thoát Ứng Dụng:** Đóng ứng dụng và giải phóng các tài nguyên liên quan.

2. **Khu Vực Log và Tiến Trình:**
   - **Khung Log (LogText):**
     Một khung chứa vùng hiển thị log kết hợp với thanh công cụ phụ. Vùng log cho phép:
     - Hiển thị thông báo log chi tiết và các thông báo hệ thống.
     - Thực hiện các thao tác lưu, sao chép, dán văn bản.
     - Điều chỉnh phông chữ, kích thước phông chữ, màu nền và màu điều khiển.
     - Các nút trên thanh công cụ của khung log được thiết kế với kích thước lớn (sử dụng emoji và tooltip tiếng Việt) để giúp người dùng hiểu rõ chức năng của từng nút.

   - **Thanh Tiến Trình:**
     Hiển thị phần trăm hoàn thành của quy trình xử lý các tệp Excel. Thanh tiến trình được cập nhật tự động khi tiến trình con gửi thông báo qua hàng đợi tiến trình.

3. **Các Cài Đặt Giao Diện:**
   Các thuộc tính như phông chữ, màu sắc và kích thước của các control đã được cài đặt mặc định khi ứng dụng khởi chạy theo sở thích của người dùng.

| Thành phần              | Mô tả                                                                                  |
|-------------------------|----------------------------------------------------------------------------------------|
| Taskbar (Thanh Công Cụ) | Nơi chứa các nút chức năng chính như tải tệp VBA, chạy VBA, lưu log,...                |
| LogText                 | Khung chứa vùng hiển thị log kết hợp với thanh công cụ phụ hỗ trợ các thao tác lưu, copy, paste, thay đổi phông, tăng, giảm cỡ phông chữ, bật tắt xuống dòng để có thể quan sát toàn bộ dòng trong cửa sổ.|
| Progress Bar            | Thanh tiến trình hiển thị phần trăm hoàn thành của quy trình xử lý các tệp Excel          |

> **Lưu ý:**  
> Trước khi chạy ứng dụng phiên bản sản xuất, người dùng cần xóa thư mục (module) `win32com/client.py` dùng để mô phỏng (testing mock) win32com đi, trước khi thi hành. Điều này đảm bảo rằng ứng dụng sẽ sử dụng phiên bản chính thức của win32com.client để tương tác với Excel.


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

