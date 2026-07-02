# Trợ Lý Quyết Toán RPA

Phần mềm desktop Windows giúp điều phối luồng làm việc giữa **GPT Custom** và
**file Excel output**. Người dùng bấm nút để mở trợ lý GPT Custom bằng file
`.bat` đã cấu hình sẵn; sau khi tải file output về thư mục Downloads, phần mềm
tự phát hiện và **tự mở file trong Excel** để người dùng kiểm tra/chỉnh sửa.
File tải về **giữ nguyên trong Downloads**; chỉ khi người dùng bấm **Đã kiểm
tra xong**, phần mềm mới **sao chép thêm một bản vào thư mục Output**.

> Giai đoạn hiện tại **chưa ghi thật** vào file theo dõi hàng ngày — việc kiểm
> tra và chỉnh sửa dữ liệu thực hiện trực tiếp trong Excel.

---

## 1. Yêu cầu hệ thống

- Windows 10/11
- Python 3.11 trở lên (khuyến nghị 3.11 hoặc 3.12)
- Không cần internet, không gọi API ChatGPT, không can thiệp trình duyệt

## 2. Cài đặt Python

1. Tải Python tại: https://www.python.org/downloads/windows/
2. Khi cài, **tích chọn** "Add Python to PATH".
3. Kiểm tra sau khi cài, mở PowerShell/CMD và gõ:

   ```bat
   python --version
   ```

## 3. Cài đặt thư viện

Mở PowerShell/CMD tại thư mục dự án và chạy:

```bat
pip install -r requirements.txt
```

Bao gồm: `PySide6`, `watchdog`, `openpyxl`.

## 4. Chạy chương trình

```bat
python main.py
```

Lần chạy đầu tiên, chương trình sẽ tự:

- Tạo cấu trúc thư mục tại `D:\RPA_QuyetToan` (nếu chưa có)
- Tạo file cấu hình `D:\RPA_QuyetToan\Config\settings.json`
- Khởi tạo database `D:\RPA_QuyetToan\Database\app_state.db`
- Bắt đầu theo dõi thư mục Downloads

## 5. Cấu hình `settings.json`

File cấu hình mặc định nằm ở `D:\RPA_QuyetToan\Config\settings.json`:

```json
{
  "app_root": "D:\\RPA_QuyetToan",
  "bat_path": "D:\\RPA_QuyetToan\\Launcher\\Mo_Tro_Ly_Quyet_Toan.bat",
  "download_folder": "D:\\RPA_QuyetToan\\Downloads",
  "output_folder": "D:\\RPA_QuyetToan\\Outputs",
  "daily_tracking_file": "D:\\RPA_QuyetToan\\Daily\\file_theo_doi_hang_ngay.xlsx",
  "allowed_extensions": [".xlsx", ".xlsm", ".csv"],
  "output_file_patterns": [
    "input_quyet_toan*.xlsx",
    "output*.xlsx",
    "input_trip*.xlsx",
    "*.xlsx"
  ],
  "download_stable_seconds": 3
}
```

Ý nghĩa các trường:

| Trường | Mô tả |
| --- | --- |
| `bat_path` | Đường dẫn file `.bat` mở trợ lý GPT Custom |
| `download_folder` | Thư mục Chrome/GPT tải file về (được theo dõi); file gốc giữ nguyên tại đây |
| `output_folder` | Thư mục lưu bản đã kiểm tra (sao chép khi bấm "Đã kiểm tra xong"), chia theo ngày |
| `daily_tracking_file` | File theo dõi hàng ngày (giai đoạn này chưa ghi) |
| `allowed_extensions` | Các đuôi file được chấp nhận |
| `output_file_patterns` | Mẫu tên file output cần bắt |
| `download_stable_seconds` | Số giây tối thiểu để coi file đã tải xong |

Bạn cũng có thể chỉnh các đường dẫn ngay trên giao diện, tại tab **Cài đặt**,
rồi bấm **Lưu cấu hình**.

## 6. Đặt thư mục tải về của Chrome

Để phần mềm bắt được file output, hãy trỏ thư mục tải về của Chrome về đúng
`download_folder`:

1. Mở Chrome → **Cài đặt** (Settings)
2. Vào **Tải xuống** (Downloads)
3. Ở mục **Vị trí** (Location), bấm **Thay đổi** và chọn:
   `D:\RPA_QuyetToan\Downloads`
4. (Khuyến nghị) **Tắt** tùy chọn "Hỏi nơi lưu từng tệp trước khi tải".

## 7. Các bước sử dụng

Mọi thao tác hằng ngày nằm ở tab **Chức năng**, theo 2 bước:

Bạn có thể bỏ qua Bước 1 và bấm **Chọn file có sẵn** ở Bước 2 để dùng lại một
file `.xlsx`, `.xlsm` hoặc `.csv` cũ. File được chọn sẽ được ghi nhận để kiểm tra
và xác nhận như file vừa tải từ trợ lý.

1. **Bước 1** — Bấm **Mở trợ lý quyết toán** (chạy file `.bat`), gửi dữ liệu
   lên **GPT Custom** và chờ xử lý.
2. **Tải file output** về (Chrome lưu vào `Downloads`).
3. Phần mềm tự phát hiện và **tự mở file trong Excel**. File gốc **giữ nguyên
   trong Downloads**.
4. **Bước 2** — Kiểm tra/chỉnh sửa toàn bộ dữ liệu và trạng thái nhập cho từng
   dòng ngay trong Excel, sau đó **lưu và đóng file**.
5. Bấm **Đã kiểm tra xong** ở thẻ Bước 2 → phần mềm **sao chép một bản vào
   `Outputs\YYYY-MM-DD\`** và chuyển trạng thái **Đã kiểm tra & hoàn tất**.
   File gốc vẫn được giữ nguyên trong thư mục ban đầu.
   - Nếu lỡ đóng file, bấm **Mở lại file** để mở lại.
   - Nếu file còn đang mở trong Excel, app sẽ nhắc lưu và đóng trước.

Tab **Cài đặt** chứa các đường dẫn/thư mục, lịch sử xử lý và nhật ký.

## 8. Cấu trúc thư mục dữ liệu

```
D:\RPA_QuyetToan
├── App\           QuyetToanAssistant.exe (khi đóng gói)
├── Config\        settings.json
├── Launcher\      Mo_Tro_Ly_Quyet_Toan.bat
├── Downloads\     Nơi Chrome/GPT tải file về; file gốc giữ nguyên tại đây
├── Outputs\       Bản đã kiểm tra (sao chép khi bấm "Đã kiểm tra xong"), chia theo ngày
├── Daily\         file_theo_doi_hang_ngay.xlsx
├── Database\      app_state.db
└── Logs\          app_YYYYMMDD.log
```

## 9. Cấu trúc mã nguồn

```
project/
├── main.py                 Điểm khởi chạy
├── requirements.txt
├── README.md
└── app/
    ├── __init__.py
    ├── config.py           Nạp/lưu settings.json, tạo thư mục
    ├── database.py         SQLite: processed_files
    ├── file_utils.py       Kiểm tra khóa, chờ tải xong, hash, sao chép sang Output
    ├── watcher.py          Theo dõi Downloads bằng watchdog
    ├── logger_setup.py     Cấu hình logging
    ├── theme.py            Bộ style QSS (tông sáng, hiện đại)
    └── main_window.py      Giao diện PySide6 (sidebar + 2 tab)
```

## 10. Đóng gói bằng PyInstaller (tùy chọn)

```bat
pip install pyinstaller
pyinstaller --noconfirm --windowed --name QuyetToanAssistant main.py
```

File `.exe` sẽ nằm trong thư mục `dist\QuyetToanAssistant\`.

## 11. Ghi chú vận hành

- File tải về **giữ nguyên trong Downloads**; phần mềm **không tự xóa/di chuyển**
  file này.
- Mỗi file chỉ có **một bản trong Output**: nếu bạn kiểm tra & xác nhận lại cùng
  một file, phần mềm **không tạo thêm file mới** mà **thay bản cũ** bằng bản mới
  (tên đặt lại theo mốc thời gian mới nhất). Các file tải về khác nhau vẫn có bản
  Output riêng.
- Dùng giờ local của Windows (không dùng UTC).
- Có thể tắt/mở lại app bất kỳ lúc nào; file gần nhất được khôi phục trên giao diện.
- Khi có file mới, phần mềm tự mở file trong Excel; bạn kiểm tra/chỉnh sửa xong,
  lưu và đóng file rồi bấm **Đã kiểm tra xong** để lưu bản vào Output.
- Mọi lỗi được ghi vào log tại `D:\RPA_QuyetToan\Logs\app_YYYYMMDD.log`.
