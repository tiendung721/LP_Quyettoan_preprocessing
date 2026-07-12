# Trợ Lý Quyết Toán RPA

Phần mềm desktop Windows điều phối luồng làm việc giữa **GPT Custom**, **file
JSON bóc tách**, **file theo dõi hàng ngày** và **phần mềm quyết toán**.

Luồng gồm 4 bước, mỗi bước đúng **một nút bấm**:

1. **Mở trợ lý quyết toán** — chạy file `.bat` mở GPT Custom để bóc tách chứng từ.
2. **Xem file bóc tách dữ liệu** — file JSON tải về được mở bằng màn hình kiểm
   tra/sửa dữ liệu của phần mềm; nút này mở lại dữ liệu bất cứ lúc nào, kèm dòng
   **“Lưu lần cuối”** hiển thị đúng thời điểm bạn bấm **Lưu**.
3. **Nhập lên file hàng ngày** — lấy **bản lưu mới nhất** của file bóc tách và cập
   nhật file theo dõi hàng ngày.
4. **Nhập dữ liệu mới lên phần mềm quyết toán** — chạy file `.bat` khởi động luồng
   **PAD RPA**.

File bóc tách **giữ nguyên trong Downloads** và là bản gốc duy nhất bạn chỉnh sửa.
Mỗi lần nhập ở Bước 3, phần mềm tự lưu một bản sao của đúng dữ liệu đã dùng vào
thư mục Output để đối chiếu về sau.

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
  "pad_bat_path": "D:\\RPA_QuyetToan\\Launcher\\Chay_PAD_Quyet_Toan.bat",
  "download_folder": "D:\\RPA_QuyetToan\\Downloads",
  "output_folder": "D:\\RPA_QuyetToan\\Outputs",
  "daily_tracking_file": "D:\\RPA_QuyetToan\\Daily\\file_theo_doi_hang_ngay.xlsx",
  "allowed_extensions": [".json"],
  "output_file_patterns": [
    "boc_tach*.json",
    "rpa_input*.json",
    "*.json"
  ],
  "download_stable_seconds": 3
}
```

Ý nghĩa các trường:

| Trường | Mô tả |
| --- | --- |
| `bat_path` | Đường dẫn file `.bat` mở trợ lý GPT Custom (Bước 1) |
| `pad_bat_path` | Đường dẫn file `.bat` chạy luồng PAD RPA (Bước 4) |
| `download_folder` | Thư mục Chrome/GPT tải file về (được theo dõi); file bóc tách giữ nguyên tại đây |
| `output_folder` | Nơi lưu bản sao của dữ liệu đã dùng cho mỗi lần nhập, chia theo ngày |
| `daily_tracking_file` | File theo dõi hàng ngày được cập nhật ở Bước 3 |
| `allowed_extensions` | Các đuôi file được chấp nhận |
| `output_file_patterns` | Mẫu tên file bóc tách cần bắt |
| `download_stable_seconds` | Số giây tối thiểu để coi file đã tải xong |

### File `.bat` chạy PAD RPA (Bước 4)

Lần chạy đầu, phần mềm **tự tạo sẵn** file mẫu tại
`D:\RPA_QuyetToan\Launcher\Chay_PAD_Quyet_Toan.bat`. Hãy mở file đó và **điền lệnh
gọi flow PAD của bạn** vào (xóa phần cảnh báo mặc định). Nếu file `.bat` chưa được
điền, khi chạy nó chỉ hiện thông báo nhắc cấu hình chứ không làm gì.

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

Mọi thao tác hằng ngày nằm ở tab **Chức năng**, theo 4 bước, mỗi bước một nút:

1. **Bước 1 — Mở trợ lý quyết toán.** Bấm nút (chạy file `.bat`), gửi chứng từ
   lên **GPT Custom** và chờ bóc tách xong, rồi **tải file JSON về**
   (Chrome lưu vào `Downloads`).
2. **Bước 2 — Xem file bóc tách dữ liệu.** Phần mềm tự phát hiện file mới và
   mở màn hình kiểm tra dữ liệu JSON. Kiểm tra/chỉnh sửa dữ liệu rồi bấm **Lưu**.
   Bấm nút này để mở lại dữ liệu bất cứ lúc nào; dòng **“Lưu lần cuối”** ngay
   dưới nút cho biết bản hiện tại được lưu lúc nào.
3. **Bước 3 — Nhập lên file hàng ngày.** Phần mềm lấy **bản lưu mới nhất** của
   file bóc tách, tự ghép Phiếu cân với Bill theo container, tạo/cập nhật SQT PM
   và ghép khoản chi theo Ngày tháng + Container.
   - Trước khi nhập, một bản sao của dữ liệu đang dùng được lưu vào
     `Outputs\YYYY-MM-DD\` để đối chiếu về sau.
   - Nếu một container có nhiều Bill, phần mềm yêu cầu chọn Bill đúng.
   - Chứng từ đã xử lý được nhận diện bằng MD5 và không bị nhập trùng.
   - Dữ liệu chưa đủ điều kiện được **hiện ra ngay sau khi nhập** để bạn bổ sung;
     phần còn lại sẽ tự được thử ghép lại ở những lần nhập sau.
4. **Bước 4 — Nhập dữ liệu mới lên phần mềm quyết toán.** Bấm nút, xác nhận, phần
   mềm chạy file `.bat` khởi động luồng **PAD RPA**. Trong lúc RPA chạy, không
   dùng chuột/bàn phím và không mở phần mềm quyết toán bằng tay.

Không cần chọn file thủ công: khi mở lại app, phần mềm tự nhận lại file bóc tách
đang dùng dở, hoặc file bóc tách mới nhất trong thư mục tải về.

Tab **Cài đặt** chứa các đường dẫn/thư mục, lịch sử xử lý và nhật ký.

## 8. Cấu trúc thư mục dữ liệu

```
D:\RPA_QuyetToan
├── App\           QuyetToanAssistant.exe (khi đóng gói)
├── Config\        settings.json
├── Launcher\      Mo_Tro_Ly_Quyet_Toan.bat, Chay_PAD_Quyet_Toan.bat
├── Downloads\     Nơi Chrome/GPT tải file về; file bóc tách giữ nguyên tại đây
├── Outputs\       Bản sao dữ liệu đã dùng cho mỗi lần nhập, chia theo ngày
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
    ├── database.py         SQLite: file đã xử lý, MD5 và dữ liệu chờ
    ├── daily_import.py     Đọc, match và ghi file quyết toán hàng ngày
    ├── daily_import_ui.py  Hộp chọn Bill và quản lý dữ liệu chờ
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

- File JSON bóc tách **giữ nguyên trong Downloads**; phần mềm **không tự xóa/di chuyển**
  file này. Bạn sửa và lưu trực tiếp bằng màn hình kiểm tra dữ liệu của phần mềm.
- Mỗi lần bấm **Nhập lên file hàng ngày**, một bản sao của dữ liệu đang dùng được
  lưu vào `Outputs\YYYY-MM-DD\` (đặt tên theo mốc thời gian) để đối chiếu về sau.
- Dùng giờ local của Windows (không dùng UTC).
- Có thể tắt/mở lại app bất kỳ lúc nào; phần mềm tự nhận lại file bóc tách đang
  dùng, hoặc file mới nhất trong thư mục tải về.
- Chỉ các dòng có **Trạng thái kiểm tra = OK** và **Trạng thái nhập = Chưa nhập**
  mới đủ điều kiện cho luồng RPA ở Bước 4.
- Mọi lỗi được ghi vào log tại `D:\RPA_QuyetToan\Logs\app_YYYYMMDD.log`.

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

$ChromePaths = @(
  "C:\Program Files\Google\Chrome\Application\chrome.exe",
  "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)

$Chrome = $ChromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Chrome) {
  Write-Host "Không tìm thấy Chrome. Cần cài Google Chrome trước."
  exit
}

$ProfileDir = "D:\RPA_ChatGPT_Profile"

if (-not (Test-Path $ProfileDir)) {
  New-Item -ItemType Directory -Path $ProfileDir | Out-Null
}

Start-Process $Chrome -ArgumentList @(
  "--user-data-dir=$ProfileDir",
  "--profile-directory=Default",
  "https://chatgpt.com"
)
