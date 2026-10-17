# KẾ HOẠCH KIẾN TRÚC: HỆ THỐNG GHI LOG CHI TIẾT & XUẤT BÁO CÁO ĐỊNH KỲ (`v3.20.0`)

## 1. Mục tiêu Nghiệp vụ & Phạm vi
- **Bài toán cần giải quyết:**
  - Hiện tại trình quét thư mục (`Kangatang_FolderScanner.ps1`) chỉ lưu trạng thái tóm tắt phục vụ Resume (`%APPDATA%\KangatangGuard\last_session.json`, `scanned_files.log`) và hiển thị 5 file nhiễm gần nhất trên màn hình Console.
  - Nhật ký kiểm toán (`scan_log_YYYYMMDD.txt`) đang nằm ẩn trong `%APPDATA%\KangatangGuard\`, chưa được lưu trực tiếp vào thư mục dự án (`D:\7. AI tools\kangatang\`) và chưa có bảng tổng hợp theo phòng ban / danh sách file nhiễm chuẩn hóa (CSV / Markdown) để đội ngũ **OhsungIT team** trích xuất làm **Báo cáo định kỳ (Tuần / Tháng / Audit ISO 27001 / LGD)**.
- **Phạm vi áp dụng:**
  1. `addin/Kangatang_FolderScanner.ps1` (Trình quét thư mục 4-Layer Turbo).
  2. `addin/Kangatang_Standalone_Scanner.ps1` (Trình quét độc lập).
  3. `addin/KangatangGuard_Code.vba` (`modLogger.WriteLog` và `RunNativeFolderScan`).

---

## 2. Bảng Đối chiếu Hiện trạng (`Before & After Matrix`)

| Tiêu chí | Hiện trạng (`v3.19.1`) | Kiến trúc mới (`v3.20.0`) |
| :--- | :--- | :--- |
| **Vị trí lưu file Log** | Chỉ lưu trong `%APPDATA%\KangatangGuard\scan_log_YYYYMMDD.txt` | Lưu song song vào thư mục dự án **`D:\7. AI tools\kangatang\Scan_Reports\`** (hoặc `\\192.168.223.176\test\Scan_Reports` từ máy trạm LAN) và `%APPDATA%\KangatangGuard\` |
| **Log chi tiết từng sự kiện** | Chỉ ghi khi phát hiện/diệt trong VBA; thiếu thông tin tổng hợp từ Tầng 2 (`Zero-COM`) và Tầng 3 (`Deep COM`) | Ghi đầy đủ vào **`Scan_Reports\scan_detail_log_YYYYMMDD.txt`** (Timestamp, Máy, User, Phiên bản, Sự kiện, Đường dẫn, Kích thước, Chi tiết mẫu nhiễm, Kết quả xử lý) |
| **Dữ liệu làm Báo cáo Định kỳ bằng Excel (CSV)** | Không có file CSV tổng hợp; phải lọc thủ công từ file `.txt` | Tự động ghi 2 file CSV chuẩn **UTF-8 with BOM** (mở bằng Excel giữ nguyên tiếng Việt/Hàn):<br>1. **`Scan_Summary_History.csv`** (Lịch sử mọi phiên quét)<br>2. **`Infected_Files_Ledger_YYYYMM.csv`** (Sổ cái toàn bộ file nhiễm & kết quả xử lý trong tháng) |
| **Báo cáo Tổng hợp Phiên quét (Audit Report)** | Chỉ hiển thị trên màn hình Console khi chạy | Tự động xuất file báo cáo **`Scan_Reports\Report_Session_YYYYMMDD_HHmmss.md`** gồm:<br>• Thống kê tổng quan & tỷ lệ %<br>• **Bảng thống kê theo Thư mục / Phòng ban cấp 1 (`0-Common`, `1-Office`, `2-Motor`, `3-Display`, `LGE`...)**<br>• Danh sách chi tiết từng file nhiễm, biến thể virus & trạng thái xử lý |
| **Cơ chế Lưu An toàn khi đang quét** | Chỉ ghi tổng kết khi quét xong 100% | Tự động cập nhật (Flush) báo cáo phiên quét định kỳ **mỗi 30 giây** và ngay khi phát hiện file nhiễm (đảm bảo dù dừng ngang vẫn có báo cáo hoàn chỉnh) |

---

## 3. Thiết kế Chi tiết Cấu trúc Thư mục & Định dạng File Báo cáo

### 3.1. Cấu trúc thư mục trong Dự án (`D:\7. AI tools\kangatang\Scan_Reports\`)
```text
D:\7. AI tools\kangatang\
└── Scan_Reports\
    ├── Scan_Summary_History.csv                # Bảng tổng hợp lịch sử tất cả các phiên quét (để vẽ biểu đồ định kỳ)
    ├── Infected_Files_Ledger_202610.csv        # Danh sách chi tiết mọi file nhiễm phát hiện trong tháng 10/2026
    ├── scan_detail_log_20261017.txt            # Log kỹ thuật chi tiết từng sự kiện trong ngày
    └── Sessions\
        └── Report_20261017_144234_vietnam.md   # Báo cáo chi tiết của từng phiên quét (gồm thống kê theo Phòng ban)
```

### 3.2. Schema các tệp phục vụ Báo cáo Định kỳ
1. **`Scan_Reports\Scan_Summary_History.csv` (UTF-8 BOM):**
   - Cột: `SessionID,StartTime,LastUpdate,DurationMinutes,Status,ComputerName,UserName,ScannerVersion,TargetFolder,TotalFolders,TotalDiscovered,TotalScanned,SkippedResume,SafeZeroCom,DeepComScanned,TotalSafe,TotalInfected,TotalCleaned,TotalLockedOrError,AvgSpeedFilesPerSec,SessionReportFile`
2. **`Scan_Reports\Infected_Files_Ledger_YYYYMM.csv` (UTF-8 BOM):**
   - Cột: `Timestamp,SessionID,ComputerName,TopDepartmentFolder,FileName,FullPath,FileSizeKB,LastModified,VirusDetail,ActionStatus,Note`
   - Trong đó `TopDepartmentFolder` tự động trích xuất tên thư mục con cấp 1 dưới thư mục gốc (ví dụ khi quét `\\192.168.223.7\file_shared\vietnam`, tự động phân loại về `0-Common`, `1-Office`, `2-Motor`, `3-Display`, `LGE`...).
3. **`Scan_Reports\Sessions\Report_YYYYMMDD_HHmmss_<Folder>.md`:**
   - **Phần 1:** Thông tin phiên quét (Thời gian, Máy trạm, Phiên bản, Thư mục mục tiêu, Trạng thái).
   - **Phần 2:** Bảng Chỉ số Tổng hợp (Tổng số file, Đã quét, Bỏ qua do đã sạch từ phiên trước, An toàn qua Zero-COM, Quét sâu qua Deep COM, Tổng file nhiễm, Đã diệt sạch, File bị khóa/Read-Only cần xử lý lại).
   - **Phần 3:** **Bảng Thống kê theo Phòng ban / Thư mục Cấp 1** (`Phòng ban / Thư mục | Số file đã quét | Số file an toàn | Số file nhiễm | Đã diệt sạch | Lỗi/Đang khóa`).
   - **Phần 4:** Bảng Danh sách Chi tiết các File Nhiễm Virus & File Đang Khóa trong phiên quét.
4. **Khôi phục dữ liệu từ Phiên quét đang chạy (`Backfill Current Session`):**
   - Trích xuất toàn bộ 134+ file nhiễm đã được xử lý trong `scan_log_20261017.txt` hôm nay để nạp ngay vào `Infected_Files_Ledger_202610.csv` và `Scan_Reports\scan_detail_log_20261017.txt`, giúp báo cáo ngày hôm nay không bị thiếu dữ liệu của buổi sáng/đầu giờ chiều!

---

## 4. Chiến lược Ngoại lệ, Hiệu năng & An toàn Hệ thống
- **Buffered Batch I/O (Không làm chậm tốc độ quét):**
  - Toàn bộ log chi tiết và cập nhật báo cáo định kỳ được gom theo lô (`Flush-ScanLogs`) mỗi 40 file hoặc mỗi khi phát hiện 1 file nhiễm, hoàn toàn không gây nghẽn ổ đĩa.
- **100% Quick Heal (`QHPISVR`) Antivirus-Safe:**
  - Chỉ sử dụng các hàm ghi file `.NET` chuẩn (`[System.IO.File]::AppendAllText`, `WriteAllText` với `UTF8Encoding($true)`), tuyệt đối không dùng `DllImport` hay chứa chuỗi IOC plaintext trong `.ps1`.
- **Tuân thủ Phân vùng Máy chủ LAN:**
  - Chỉ ghi báo cáo vào `D:\7. AI tools\kangatang\Scan_Reports` (và nhánh LAN `\\192.168.223.176\test\Scan_Reports`), tuyệt đối không ghi vào `192.168.223.7`.

---

## 5. Kế hoạch Kiểm thử & Vận hành
1. Kiểm thử tạo và ghi tự động bộ 4 file (`Scan_Summary_History.csv`, `Infected_Files_Ledger_202610.csv`, `scan_detail_log_20261017.txt`, `Sessions\Report_...md`) trong `D:\7. AI tools\kangatang\Scan_Reports\`.
2. Kiểm thử Backfill dữ liệu các file nhiễm đã diệt trong ngày `2026-10-17` từ `%APPDATA%\KangatangGuard\scan_log_20261017.txt` sang bộ báo cáo mới trong thư mục dự án.
3. Kiểm tra mở trực tiếp file `.csv` bằng Excel đảm bảo hiển thị đúng dấu tiếng Việt và tiếng Hàn.

---

## 6. Câu hỏi Làm rõ / Xác nhận Thực thi
1. Cấu trúc bộ 4 file trong thư mục **`D:\7. AI tools\kangatang\Scan_Reports\`** (gồm Log chi tiết `.txt`, Bảng lịch sử phiên quét `.csv`, Sổ cái file nhiễm hàng tháng `.csv` mở trực tiếp bằng Excel, và Báo cáo tổng hợp theo Phòng ban `.md`) đã đáp ứng đúng nhu cầu làm báo cáo định kỳ của bạn chưa?
2. Bạn có muốn hệ thống **tự động tổng hợp (backfill) luôn toàn bộ kết quả quét từ đầu ngày hôm nay (`scan_log_20261017.txt` — hơn 134 file nhiễm đã diệt)** vào thư mục `D:\7. AI tools\kangatang\Scan_Reports\` ngay sau khi triển khai không?
