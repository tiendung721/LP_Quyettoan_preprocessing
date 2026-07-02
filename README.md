# Trợ Lý Quyết Toán RPA

Phần mềm desktop Windows giúp điều phối luồng làm việc giữa **GPT Custom** và
**file Excel output**. Người dùng bấm nút để mở trợ lý GPT Custom bằng file
`.bat` đã cấu hình sẵn; sau khi tải file output về thư mục Downloads, phần mềm
tự phát hiện, sao lưu bản gốc, chuyển vào thư mục output chuẩn, cho phép mở file
để kiểm tra/chỉnh sửa, rồi duyệt (xem trước) dữ liệu từ đúng file đã chỉnh sửa.

> Giai đoạn hiện tại **chưa ghi thật** vào file theo dõi hàng ngày — chỉ đọc và
> hiển thị preview dữ liệu.

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
  "backup_folder": "D:\\RPA_QuyetToan\\Backup",
  "daily_tracking_file": "D:\\RPA_QuyetToan\\Daily\\file_theo_doi_hang_ngay.xlsx",
  "allowed_extensions": [".xlsx", ".xlsm", ".csv"],
  "output_file_patterns": [
    "input_quyet_toan*.xlsx",
    "output*.xlsx",
    "input_trip*.xlsx",
    "*.xlsx"
  ],
  "download_stable_seconds": 3,
  "auto_open_after_download": false
}
```

Ý nghĩa các trường:

| Trường | Mô tả |
| --- | --- |
| `bat_path` | Đường dẫn file `.bat` mở trợ lý GPT Custom |
| `download_folder` | Thư mục Chrome/GPT tải file output về (được theo dõi) |
| `output_folder` | Thư mục lưu file làm việc (working), chia theo ngày |
| `backup_folder` | Thư mục lưu bản gốc (backup), chia theo ngày |
| `daily_tracking_file` | File theo dõi hàng ngày (giai đoạn này chưa ghi) |
| `allowed_extensions` | Các đuôi file được chấp nhận |
| `output_file_patterns` | Mẫu tên file output cần bắt |
| `download_stable_seconds` | Số giây tối thiểu để coi file đã tải xong |
| `auto_open_after_download` | Tự mở file ngay khi phát hiện xong |

Bạn cũng có thể chỉnh các đường dẫn ngay trên giao diện (Vùng 1) rồi bấm
**Lưu cấu hình**.

## 6. Đặt thư mục tải về của Chrome

Để phần mềm bắt được file output, hãy trỏ thư mục tải về của Chrome về đúng
`download_folder`:

1. Mở Chrome → **Cài đặt** (Settings)
2. Vào **Tải xuống** (Downloads)
3. Ở mục **Vị trí** (Location), bấm **Thay đổi** và chọn:
   `D:\RPA_QuyetToan\Downloads`
4. (Khuyến nghị) **Tắt** tùy chọn "Hỏi nơi lưu từng tệp trước khi tải".

## 7. Các bước sử dụng

1. Bấm **Mở trợ lý quyết toán** (chạy file `.bat`).
2. Gửi file lên **GPT Custom** và chờ xử lý.
3. **Tải file output** về (Chrome lưu vào `Downloads`).
4. Phần mềm tự phát hiện, sao lưu bản gốc, chuyển vào `Outputs\YYYY-MM-DD\`
   và hiển thị ở **Vùng 3**.
5. Bấm **Mở file kết quả** để kiểm tra/chỉnh sửa trong Excel.
6. **Lưu và đóng** file Excel.
7. Bấm **Đã kiểm tra và dùng file này** (nếu file còn mở, app sẽ nhắc bạn
   lưu và đóng trước).
8. Bấm **Duyệt dữ liệu từ file output** để xem preview 20 dòng đầu, số sheet,
   số dòng, số cột.

## 8. Cấu trúc thư mục dữ liệu

```
D:\RPA_QuyetToan
├── App\           QuyetToanAssistant.exe (khi đóng gói)
├── Config\        settings.json
├── Launcher\      Mo_Tro_Ly_Quyet_Toan.bat
├── Downloads\     Nơi Chrome/GPT tải file output về
├── Outputs\       File làm việc (working), chia theo ngày
├── Backup\        Bản gốc (backup_original_...), chia theo ngày
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
    ├── file_utils.py       Kiểm tra khóa, chờ tải xong, hash, di chuyển file
    ├── watcher.py          Theo dõi Downloads bằng watchdog
    ├── excel_preview.py    Đọc Excel/CSV, lấy preview
    ├── logger_setup.py     Cấu hình logging
    └── main_window.py      Giao diện PySide6
```

## 10. Đóng gói bằng PyInstaller (tùy chọn)

```bat
pip install pyinstaller
pyinstaller --noconfirm --windowed --name QuyetToanAssistant main.py
```

File `.exe` sẽ nằm trong thư mục `dist\QuyetToanAssistant\`.

## 11. Ghi chú vận hành

- Phần mềm **không tự xóa** và **không ghi đè** file cũ; luôn tạo backup bản gốc.
- Dùng giờ local của Windows (không dùng UTC).
- Có thể tắt/mở lại app bất kỳ lúc nào; file gần nhất được khôi phục trên giao diện.
- Khi có file output mới, nút **Duyệt dữ liệu** bị khóa cho tới khi bạn bấm
  **Đã kiểm tra và dùng file này**.
- Mọi lỗi được ghi vào log tại `D:\RPA_QuyetToan\Logs\app_YYYYMMDD.log`.
