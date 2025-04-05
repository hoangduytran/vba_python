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

Describe how to install and set up your project. Include any prerequisites or dependencies. For example:

```bash
git clone https://github.com/hoangduytran/vba_python.git
cd vba_python
pip install -r requirements.txt
```

## usage (Sử Dụng)

Giải thích chi tiết:
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

Nên nhớ, đây chỉ là sườn bài, bạn phải áp dụng nó cụ thể vào trong trường hợp của bạn. Phải biết lấy Python về máy, cài đặt nó theo nhu cầu. Rồi lại phải biết chạy lệnh

`python3 -m pip install -r requirements.txt`

để lấy các bản thư viện phụ thuộc yêu cầu.

Bên cạnh đó, phải học cách viết các bản `macro` trong ngôn ngữ `VBA` (Visual Basic for Applications: Visual Basic cho Ứng Dụng). Đồng thời phải thông thuộc cấu trúc và các hàm của Excel để có thể sử dụng chúng thành thục.

## contributing (Đóng Góp)

Nếu biết về lập trình và `forking repository` (tạo một bản sao cho mình tại kho riêng của mình) và lấy về máy bằng lệnh `git clone`, rồi chỉnh sửa và tạo một `pull-request` thì mình có thể giúp bạn hội nhập các thay đổi để những người khác cũng dùng chung được các thay đổi của bạn, một điều rất đáng hoan nghênh.

## contact (Liên Lạc)
e-mail: hoangduytran1960@gmail.com
Facebook: https://www.facebook.com/hoangduy.tran

