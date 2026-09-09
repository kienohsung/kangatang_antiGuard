# KangatangGuard - Hệ Thống Phòng Vệ & Diệt Virus Macro Excel Chuyên Dụng

[![Version](https://img.shields.io/badge/version-v3.5.2-blue.svg)](https://github.com/kienohsung/kangatang_antiGuard)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Office%202010--365-green.svg)](https://github.com/kienohsung/kangatang_antiGuard)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE)

**KangatangGuard** là giải pháp bảo mật chuyên sâu cấp độ doanh nghiệp (Production-grade) nhằm ngăn chặn, phát hiện thời gian thực và tiêu diệt triệt để dòng virus macro nguy hiểm **Kangatang / Laroux / mypersonnel** trên Microsoft Excel.

---

## 🚀 Tính Năng Nổi Bật

- 🛡️ **Tự Động Phòng Vệ Thời Gian Thực (Real-Time Excel Add-in):**
  - Tự động kiểm tra và làm sạch tệp Excel ngay khi vừa mở.
  - Tích hợp bộ nhớ đệm thông minh ScanCache chống quét lặp lại khi người dùng lưu tệp hoặc AutoSave.
  - Hiển thị menu điều khiển chuyên dụng **KangatangGuard** trên thanh công cụ Excel.

- ⚡ **Quét Luồng Trực Tiếp Không Độ Trễ (Real-Time Directory Streaming):**
  - Quét ngay lập tức lần lượt từng thư mục và tệp Excel theo mô hình luồng (*Streaming Pipeline*).
  - Loại bỏ hoàn toàn độ trễ chờ đợi thu thập trước danh sách tệp.

- 🔄 **Kiến Trúc Tách Ngoài Tiến Trình (Out-of-Process Worker):**
  - Tác vụ quét thư mục nặng được bàn giao cho tiến trình ngầm độc lập chạy trên cửa sổ riêng với thanh tiến độ % thời gian thực.
  - Ứng dụng Excel của người dùng **hoàn toàn mượt mà 100%**, không bao giờ bị treo Book1 - Excel (Not Responding).

- 🩺 **Cơ Chế Tự Phục Hồi Tức Thì (COM Auto-Recovery & PID Tracking):**
  - Quản lý vòng đời tiến trình Excel COM riêng biệt bằng PID đích danh.
  - Tự động phát hiện và hồi sinh phiên làm việc trong **0.5 giây** nếu gặp tệp hỏng hoặc có mật khẩu riêng, bảo đảm **một tệp lỗi không bao giờ làm gián đoạn các tệp còn lại**.
  - Định kỳ làm mới bộ nhớ sau mỗi 30 tệp, kiểm soát RAM luôn dưới 40MB.

- 🏢 **Hỗ Trợ Toàn Diện Office 365 (Dual-Registration Defense):**
  - Tự động nạp song song vào thư mục XLSTART và đăng ký khóa Registry HKCU\Software\Microsoft\Office\<Version>\Excel\Options\OPEN*.
  - Bảo vệ vững chắc ngay cả với các phiên bản Office 365 không khởi tạo thư mục XLSTART.

- 🔒 **Bảo Toàn Dữ Liệu Tuyệt Đối (Zero Data Loss):**
  - Mở tệp ở chế độ an toàn ReadOnly = True khi kiểm tra.
  - Tự động tạo bản sao lưu có gắn tem thời gian vào thư mục _Backup_Kangatang trước khi tiến hành khử mã độc.

- 🌐 **Hỗ Trợ Đầy Đủ Tiếng Việt Có Dấu & Đường Dẫn Mạng UNC:**
  - Tương thích 100% với tên tệp tiếng Việt, tiếng Hàn, tiếng Trung và đường dẫn mạng nội bộ \\server\share.

---

## 📂 Cấu Trúc Mã Nguồn

`	ext
kangatang_antiGuard/
├── addin/                                # Thư mục mã nguồn và cài đặt Add-in
│   ├── Install_Addin.bat                 # Kịch bản 1-click cài đặt Add-in (yêu cầu Admin)
│   ├── Install_KangatangGuard.ps1       # Trình biên dịch mã nguồn sang KangatangGuard.xlam
│   ├── KangatangGuard_Code.vba           # Toàn bộ mã nguồn VBA chuẩn hóa của Add-in
│   ├── Kangatang_FolderScanner.ps1       # Trình quét luồng ngầm độc lập v3.5.2
│   ├── Uninstall_Addin.bat               # Kịch bản gỡ bỏ Add-in an toàn và sạch sẽ
│   └── addin_kangatang_test_3.9.2026.zip # Gói nén phân phối Add-in
├── doc/                                  # Tài liệu kiến trúc và lịch sử nâng cấp
│   ├── 20260817_kangatang_cleaner_fix_plan.md
│   ├── 20260903_kangatang_addin_plan.md
│   ├── 20260903_office365_non_xlstart_defense_plan.md
│   ├── 20260908_fix_folder_scan_unicode_v3.3.0.md
│   ├── 20260909_decouple_scanner_architecture_v3.4.0.md
│   └── 20260909_stream_scan_architecture_v3.5.0.md
├── Diet_Virus_Kangatang.bat              # Công cụ quét và diệt virus độc lập dạng CLI
├── Chay_Diet_Virus.bat                   # Kịch bản khởi chạy diệt virus
├── clean_excel_virus.ps1                 # Script PowerShell diệt virus cơ bản
├── .gitignore                            # Cấu hình bỏ qua tệp tạm và thư mục sao lưu
└── README.md                             # Tài liệu hướng dẫn sử dụng dự án
`

---

## 🛠️ Hướng Dẫn Cài Đặt & Sử Dụng

### Cách 1: Cài đặt Excel Add-in Tự Động Phòng Vệ (Khuyên dùng)
1. Mở thư mục ddin/.
2. Nhấp chuột phải vào tệp Install_Addin.bat chọn **Run as administrator**.
3. Hệ thống sẽ tự động cấu hình Registry, biên dịch KangatangGuard.xlam và triển khai vào máy tính.
4. Mở Excel bình thường:
   - Mọi tệp bạn mở sẽ được bảo vệ tự động.
   - Menu **KangatangGuard** sẽ xuất hiện trên thanh điều khiển của Excel để bạn quét thủ công hoặc quét thư mục khi cần.

### Cách 2: Quét & Diệt Tổng Lực Bằng Script Độc Lập
- Nhấp đúp (hoặc Run as administrator) tệp Diet_Virus_Kangatang.bat ở thư mục gốc.
- Chương trình sẽ quét toàn diện các tệp khởi động hệ thống, bộ nhớ đệm và các ổ đĩa để triệt phá ổ dịch virus.

### Cách Gỡ Bỏ Add-in (Khi không còn nhu cầu)
- Chạy tệp ddin/Uninstall_Addin.bat với quyền Administrator để dọn dẹp sạch sẽ Add-in và Registry khỏi Office.

---

## 📝 Nhật Ký Phiên Bản (Changelog)

- **v3.5.3:**
  - Tích hợp **Native C# Watchdog Engine (25s timeout)** trên luồng ngầm CLR, tự động ngắt và hồi sinh Excel khi gặp tệp treo/hỏng hoặc modal dialog ngầm trên mạng SMB.
  - Thêm **Pre-flight Write Lock Check** (.NET FileStream) phát hiện tệp đang được người dùng khác mở ghi để bỏ qua an toàn, ngăn chặn xung đột khóa mạng.
  - Chặn triệt để hộp thoại Modal: truyền dummy password và tắt cảnh báo tương thích/bảo mật (`CheckCompatibility = $false`).
  - Thêm cơ chế **Auto-Recovery ngay trong khối catch khử khuẩn**, khôi phục kênh COM tự động khi gặp sự cố giữa chừng.
- **v3.5.2:**
  - Sửa triệt để lỗi biên dịch Duplicate Option statement trong VBA.
  - Tích hợp chuẩn hóa chuỗi tiêm mã phòng vệ Inject-CleanVbaModule.
- **v3.5.1:**
  - Khắc phục lỗi Unable to get the Open property of the Workbooks class khi quét thư mục có hàng nghìn tệp.
  - Thêm cơ chế **COM Auto-Recovery** và quản lý PID tiến trình nguyên tử.
- **v3.5.0:**
  - Nâng cấp cơ chế quét luồng trực tiếp Scan-FolderStream (Zero pre-indexing delay).
- **v3.4.0:**
  - Tách trình quét thư mục thành tiến trình ngầm độc lập (Out-of-Process Background Worker) chống treo giao diện Excel.
- **v3.3.0 - v3.0.0:**
  - Hỗ trợ phòng vệ đa tầng Office 365, giao diện Unicode tiếng Việt có dấu qua Windows API MessageBoxW.

---

## 📄 Bản Quyền & Giấy Phép

Dự án được phát hành dưới giấy phép [MIT License](LICENSE). Tự do sử dụng và triển khai cho môi trường cá nhân cũng như doanh nghiệp.
