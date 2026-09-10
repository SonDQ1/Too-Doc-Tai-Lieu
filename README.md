# Công cụ tạo giấy chứng nhận kiểm định

Công cụ desktop (Windows) giúp lập hàng loạt **giấy chứng nhận kiểm định** từ
biên bản kiểm định, **giữ nguyên 100% định dạng** của mẫu Word: font, cỡ chữ,
bảng biểu, căn lề, header/footer — chỉ thay đúng phần nội dung cần thay.

## Hai chức năng

| Tab | Đầu vào | Kết quả |
|---|---|---|
| **Word → Word** | Một file Word biên bản (nguồn) + một mẫu Word | File Word nhiều trang, mỗi thiết bị một trang |
| **Excel → Word** | Cả thư mục biên bản Excel + một mẫu Word | File Word nhiều trang, mỗi thiết bị một trang |

Ngoài ra: đánh **số giấy chứng nhận tự tăng**, **xem trước** danh sách trang sẽ
tạo, **xem mẫu** dưới dạng ảnh trang giấy, và **gợi ý mẫu** phù hợp với dữ liệu.

## Các loại phương tiện đo đã hỗ trợ

Máy biến dòng trung thế · Máy biến dòng hạ thế · Máy biến áp đo lường ·
Công tơ điện xoay chiều 3 pha · Cân đồng hồ lò xo · Cân kỹ thuật · Cân bàn ·
Cân đĩa · Taximet

Đang cập nhật: Giấy chứng nhận kiểm xạ · Thiết bị X-quang · Cột đo xăng dầu

## Cách chạy

```bash
pip install lxml openpyxl xlrd pywin32 pymupdf
python ThayTheDuLieu.py
```

Chạy nhanh không cần giao diện:

```bash
python ThayTheDuLieu.py <fileNguon.docx> <mau.docx> <fileXuat.docx> --so="4111 - 26 (YB)"
python ThayTheDuLieu.py --excel <thưMụcExcel> <mau.docx> <fileXuat.docx> --so=400
python ThayTheDuLieu.py --mau          # liệt kê các mẫu đang nhận
```

## Đóng gói thành file .exe cho máy không có Python

```bash
pyinstaller --onefile --windowed --name "ThayTheDuLieu" ^
  --hidden-import win32com.client --hidden-import pythoncom ^
  --hidden-import openpyxl --hidden-import xlrd --hidden-import fitz ^
  --add-data "MauWord;MauWord" ThayTheDuLieu.py
```

Bản gửi khách gồm: `ThayTheDuLieu.exe` + thư mục `MauWord/` + `CauHinhThoiHan.txt`
đặt cạnh nhau.

## Cấu hình thời hạn hiệu lực

`CauHinhThoiHan.txt` quy định thời hạn của giấy theo phương pháp thực hiện:

```
ĐLVN 30: 2019|2      # 2 năm
ĐLVN 01: 2019|18t    # 18 tháng
```

Thời hạn = ngày kiểm định + số năm/tháng tương ứng. Nếu biên bản Excel đã ghi
sẵn ô "Hiệu lực đến" thì lấy thẳng ngày đó, không dùng cấu hình.

Sửa được ngay trong chương trình bằng nút **"Cấu hình thời hạn..."**, hoặc mở
file bằng Notepad.

## Ghi chú về thư mục mẫu

Thư mục `MauWord/` chứa các mẫu Word thật của đơn vị (có tên đơn vị sử dụng,
tên kiểm định viên, số tem) nên **không được đưa lên kho mã công khai**.
Chương trình ưu tiên đọc thư mục `MauWord/` đặt cạnh file chạy, nên có thể bổ
sung hoặc sửa mẫu mà không cần build lại.

## Yêu cầu môi trường

- Windows, Microsoft Word (chỉ cần khi dùng mẫu `.doc` đời cũ hoặc xem mẫu dạng ảnh)
- File Excel `.xls` đời cũ đọc trực tiếp bằng thư viện `xlrd`, **không cần Excel**
