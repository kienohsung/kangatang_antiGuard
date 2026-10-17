# Báo Cáo Kiến Trúc v3.17.0: Khắc Phục Triệt Để Lỗi Antivirus (Seqrite Endpoint Security / Quick Heal) Nhận Nhầm & Cách Ly (`Zero-PS1 Native VBA Folder Scanner`)

- **Ngày thực hiện:** 2026-10-17
- **Phiên bản phát hành:** `v3.17.0`
- **Tác giả:** Tech Lead & Senior Full-stack Architect

---

## 1. Hiện Trạng & Nguyên Nhân Gốc Rễ (Root Cause Analysis)

### 1.1. Hiện tượng trên máy trạm Client (`DPL-Prod-mrHuy`)
Trên máy trạm có cài đặt phần mềm diệt virus doanh nghiệp **Seqrite Endpoint Security (Quick Heal)**, tệp:
- `C:\Users\<User>\AppData\Roaming\KangatangGuard\Kangatang_FolderScanner.ps1`
bị cảnh báo nhận nhầm (False-Positive) và đưa vào danh sách cách ly (Quarantine).

### 1.2. Nguyên nhân gốc rễ (4 Yếu tố kích hoạt Heuristic/Behavioral EDR)
1. **Vị trí lưu trữ nhạy cảm (`%APPDATA%\KangatangGuard\*.ps1`):** Các trình EDR/Antivirus (đặc biệt là Seqrite / Quick Heal) áp dụng luật giám sát rất gắt gao với mọi tệp `.ps1` được sinh ra hoặc thực thi từ thư mục người dùng `%APPDATA%\Roaming\...`.
2. **Excel gọi `WScript.Shell.Run "powershell.exe -ExecutionPolicy Bypass -File %APPDATA%\..."`:** Hành vi tiến trình Office (`EXCEL.EXE`) khởi chạy `powershell.exe -ExecutionPolicy Bypass` trỏ tới `%APPDATA%` trùng khớp với mẫu hành vi (Behavioral Signature) của dòng mã độc *Office Macro Dropper*.
3. **Kỹ thuật C# Win32 P/Invoke (`Add-Type`) trong `.ps1`:** Tệp `Kangatang_FolderScanner.ps1` cũ (`v3.8.3`, dài 1.118 dòng) nhúng mã C# biên dịch động tại runtime gọi các hàm API hệ thống cấp thấp (`[DllImport("kernel32.dll")] GetConsoleMode / SetConsoleMode`, `[DllImport("ole32.dll")] CoRegisterMessageFilter`, `[DllImport("user32.dll")] GetWindowThreadProcessId`, `Process.Kill()`).
4. **Chuỗi dấu hiệu (IOC Strings) của virus nằm dạng văn bản thuần (Plaintext) trong `.ps1`:** Các chuỗi `"kangatang"`, `"mypersonnel"`, `"laroux"`, `"auto_open"`, `"VBComponents.Remove"` xuất hiện dày đặc kết hợp với thao tác can thiệp Registry `AccessVBOM` khiến bộ quét tĩnh (Static Heuristic Scanner) của Seqrite đánh dấu tệp `.ps1` là công cụ can thiệp hệ thống đáng ngờ.

---

## 2. Giải Pháp Kiến Trúc `v3.17.0` (`100% Native VBA Folder Scanner & Zero-PS1 in %APPDATA%`)

1. 🛡️ **Tích hợp Trình Quét Thư Mục 100% Native VBA (`RunNativeFolderScan`) ngay trong `KangatangGuard.xlam`:**
   - Toàn bộ tính năng **Quét Thư Mục (`OnRibbonScanFolder`)** và **Tiếp Tục Quét Nhanh (`OnRibbonResumeScan` - Fast Resume)** chạy trực tiếp bằng mã VBA thuần bên trong `KangatangGuard.xlam`.
   - `KangatangGuard.xlam` **không bao giờ lưu, không sao chép và không gọi `powershell.exe`** đối với bất kỳ tệp `.ps1` nào trong `%APPDATA%\KangatangGuard`.
   - Trong quá trình quét thư mục, `RunNativeFolderScan` tắt `Application.EnableEvents = False` (và `Workbooks.Open` từ VBA mặc định không kích hoạt `Auto_Open`), giúp mở tệp cần kiểm tra an toàn tuyệt đối mà vẫn giữ nguyên quyền truy cập `wbTarget.VBProject` để bóc tách mã độc và sao lưu về máy chủ `192.168.223.176`.
2. 🛡️ **Cơ chế Tự Động Dọn Sạch Tệp `.ps1` Cũ Trong `%APPDATA%\KangatangGuard` (`Active Self-Cleanup`):**
   - Ngay khi khởi động Excel (`CleanAddInCollisions`), khi cập nhật tự động qua LAN (`PerformLanUpdate`), hoặc khi chạy `Install_KangatangGuard.ps1` / `Install_Client.ps1`, hệ thống tự động xóa bỏ hoàn toàn tệp `Kangatang_FolderScanner.ps1` còn sót lại trong `%APPDATA%\KangatangGuard` và xóa giá trị Registry `HKCU\Software\KangatangGuard\ScannerScript`.
3. 🛡️ **Tái cấu trúc `Kangatang_FolderScanner.ps1` & `Kangatang_Standalone_Scanner.ps1` thành Trình Gọi COM Thuần Khiết (`Zero P/Invoke, Zero IOC`):**
   - Loại bỏ **100%** khối mã C# `Add-Type` P/Invoke (`kernel32.dll`, `ole32.dll`, `user32.dll`, `Process.Kill()`) và loại bỏ **100%** chuỗi tên virus nhạy cảm khỏi các tệp `.ps1` công cụ trên máy chủ LAN.
   - Các tệp `.ps1` chỉ còn là trình bao (wrapper ~100 dòng) khởi tạo COM `Excel.Application` và ủy quyền trực tiếp cho hàm `RunNativeFolderScan` / `ShieldStartupSweep` bên trong `KangatangGuard.xlam`.

---

## 3. Bảng Đối Chiếu Hiện Trạng (`Before & After Matrix`)

| Hạng mục Kiến trúc | Trước (`v3.16.0` / Scanner `v3.8.3`) | Sau (`v3.17.0` Production-grade) |
| :--- | :--- | :--- |
| **Cơ chế Quét Thư Mục từ Ribbon Excel** | Gọi `WScript.Shell.Run "powershell.exe -ExecutionPolicy Bypass -File %APPDATA%\KangatangGuard\Kangatang_FolderScanner.ps1"` | Chạy **100% Native VBA (`RunNativeFolderScan`)** ngay trong `KangatangGuard.xlam`, không gọi `powershell.exe` |
| **Tệp `.ps1` trong `%APPDATA%\KangatangGuard`** | Luôn tồn tại `Kangatang_FolderScanner.ps1` (dễ bị Seqrite / Quick Heal quét và cách ly) | **0 tệp `.ps1`** trong `%APPDATA%\KangatangGuard` (tự động xóa sạch bản cũ khi khởi động/cập nhật) |
| **Mã C# Win32 P/Invoke & Chuỗi IOC trong `.ps1`** | 1.118 dòng có `DllImport("kernel32.dll")`, `ole32.dll`, `user32.dll`, `Process.Kill()` và chuỗi IOC virus | **Loại bỏ 100%** P/Invoke, `Process.Kill()` và chuỗi IOC; chỉ còn ~100 dòng COM Wrapper sạch |
| **Hỗ trợ Fast Resume (`last_session.json` & `scanned_files.log`)** | Quản lý qua PowerShell bên ngoài | Tích hợp trực tiếp trong `RunNativeFolderScan` (VBA UTF-8 `ADODB.Stream` + `Scripting.Dictionary`) |

---

## 4. Kết Quả Kiểm Chứng Tự Động (End-to-End COM Verification)

- `LegacyPs1ExistsAfterCleanAddInCollisions = False`: Tệp `.ps1` cũ trong `%APPDATA%\KangatangGuard` được tự động dọn sạch ngay khi `CleanAddInCollisions` thực thi.
- `Pass1_SessionStatus = COMPLETED | TotalFiles = 2 | TotalScanned = 2 | TotalCleaned = 1`: `RunNativeFolderScan` phát hiện chính xác `InfectedBook.xlsm`, sao lưu về `Quarantine_Backup` (máy `223.176`), gỡ bỏ module độc hại `Kangatang` (`InfectedBookStillHasKangatang = False`) và giữ nguyên `CleanBook.xlsx`.
- `Pass2_SessionStatus = COMPLETED | TotalFiles = 2 | TotalScanned = 2 | Skipped = 2`: Chế độ `Fast Resume` (`bResume = True`) bỏ qua tức thì các tệp không thay đổi chữ ký (`Size|DateLastModified`).
- Đã đồng bộ gói phát hành `v3.17.0` lên LAN Hub `\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang` và bản Clone cục bộ `distribution/`.
