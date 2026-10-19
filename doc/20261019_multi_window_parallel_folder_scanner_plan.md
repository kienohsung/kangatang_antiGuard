# KẾ HOẠCH KIẾN TRÚC: HỖ TRỢ CHẠY ĐA CỬA SỔ POWERSHELL SONG SONG CHO NHIỀU THƯ MỤC (`v3.22.0`)

## 1. Mục tiêu Nghiệp vụ & Phạm vi
- **Bài toán thực tế:** Hiện tại khi đang chạy quét một thư mục (trên máy chủ `223.7` hoặc máy quản trị `223.176`), nếu người dùng bấm chạy lại tệp [`Chay_Quet_Truc_Tiep_Tren_223.7.bat`](file:///d:/7.%20AI%20tools/kangatang/addin/Chay_Quet_Truc_Tiep_Tren_223.7.bat) để quét song song một thư mục khác, cửa sổ PowerShell đang quét bị đóng cưỡng bức và các tiến trình `EXCEL.EXE` Worker đang chạy bị tiêu diệt.
- **Mục tiêu `v3.22.0` (Multi-Instance Concurrent Folder Scanner):**
  1. Cho phép mở **nhiều cửa sổ PowerShell cùng lúc** (mỗi cửa sổ quét một thư mục độc lập).
  2. Chặn trùng lặp thông minh theo **từng thư mục (`Per-Folder Mutex`)**: chỉ ngăn khi 2 cửa sổ chọn **cùng một thư mục đang quét**, hoàn toàn tự do khi chọn các thư mục khác nhau.
  3. Cô lập vòng đời tiến trình `EXCEL.EXE` Worker theo từng cửa sổ (`PID-Scoped Worker Tracking`) để cửa sổ mới khởi chạy hoặc kết thúc tuyệt đối không tắt nhầm `EXCEL.EXE` của cửa sổ đang chạy.
  4. Đồng bộ an toàn đa tiến trình (`Cross-Process Mutex Locking`) khi ghi vào các tệp nhật ký/báo cáo dùng chung (`scanned_files.log`, `Scan_Summary_History.csv`, `Infected_Files_Ledger_YYYYMM.csv`, `scan_detail_log_YYYYMMDD.txt`).

---

## 2. Bảng Đối chiếu Hiện trạng (`Before & After Matrix` — `v3.21.5` ➔ `v3.22.0`)

| Thành phần | Hiện trạng (`v3.21.5`) | Kiến trúc Mới (`v3.22.0`) |
| :--- | :--- | :--- |
| **1. Khởi chạy từ [`Chay_Quet_Truc_Tiep_Tren_223.7.bat`](file:///d:/7.%20AI%20tools/kangatang/addin/Chay_Quet_Truc_Tiep_Tren_223.7.bat#L12-L13)** | 🛑 Chạy `wmic ... terminate` giết mọi `powershell.exe` đang chạy `Kangatang_FolderScanner` và `taskkill /F /IM excel.exe` giết toàn bộ Excel Worker của phiên đang chạy. | 🛡️ **(VACCINE)** Tuyệt đối **KHÔNG** tắt `powershell.exe` hay `excel.exe` nếu phát hiện đang có phiên `Kangatang_FolderScanner` hoạt động. Chỉ dọn `excel.exe` mồ côi khi hiện tại có `0` cửa sổ quét nào đang mở, và tách cửa sổ mới độc lập (`start`). |
| **2. Chốt chặn Single-Instance trong [`Kangatang_FolderScanner.ps1`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1#L73-L114)** | 🛑 Dùng khóa toàn cục `Local\KangatangGuard_FolderScanner_Mutex` (chỉ cho phép 1 cửa sổ duy nhất trên toàn máy) và `Stop-Process -Name EXCEL` không phân biệt tại lúc khởi động. | 🛡️ **(VACCINE)** Chuyển sang **Khóa theo Thư mục Mục tiêu (`Local\KangatangGuard_Folder_<Hash>`)** đặt *sau* bước chọn thư mục: cho phép mở $N$ cửa sổ cho $N$ thư mục khác nhau; nếu chọn trùng đúng thư mục đang quét thì hiện thông báo cảnh báo rõ ràng. |
| **3. Quản lý & Dọn dẹp `EXCEL.EXE` Worker** | Dùng `KangatangGuard_Worker_$MasterPid_$workerId.xlam` nhưng khi khởi động lại quét sạch mọi `EXCEL.EXE` có `MainWindowHandle -eq 0`. | 🛡️ **(VACCINE)** Mỗi cửa sổ chỉ quản lý và dọn dẹp các `EXCEL.EXE` Worker do chính `MasterPid` của nó sinh ra (hoặc mồ côi từ `MasterPid` đã chết), bảo vệ 100% Worker của các cửa sổ đang chạy song song. Tự động cân bằng số Worker (`Adaptive Worker Scaling`) khi mở nhiều cửa sổ. |
| **4. Chế độ `[NO] Quét mới từ đầu` & Ghi Log dùng chung** | 🛑 Chọn `[NO]` sẽ xóa trắng toàn bộ `scanned_files.log` (`WriteAllText(..., "")`), làm mất lịch sử của các thư mục khác. Ghi file CSV/Log chưa có khóa liên tiến trình. | 🛡️ **(VACCINE)** Chọn `[NO]` chỉ bỏ qua/lọc lại dữ liệu Resume của **riêng thư mục đang chọn**, giữ nguyên 100% lịch sử các thư mục khác trong `scanned_files.log`. Bọc khóa `Local\KangatangGuard_LogWrite_Mutex` (`< 10ms`) quanh mọi thao tác ghi `scanned_files.log`, `Scan_Summary_History.csv` và `Sessions\Report_<SessionID>.md`. |

---

## 3. Thiết kế Chi tiết Backend / PowerShell Engine & Batch Launcher

### 3.1. Tinh chỉnh [`Chay_Quet_Truc_Tiep_Tren_223.7.bat`](file:///d:/7.%20AI%20tools/kangatang/addin/Chay_Quet_Truc_Tiep_Tren_223.7.bat) và [`Chay_Diet_Virus_Ngoai.bat`](file:///d:/7.%20AI%20tools/kangatang/addin/Chay_Diet_Virus_Ngoai.bat)
- Kiểm tra xem hiện có tiến trình `powershell.exe` nào đang chạy `Kangatang_FolderScanner.ps1` hay không:
  - **Nếu KHÔNG có cửa sổ nào đang chạy (`0` instance):** Cho phép dọn sạch các tiến trình `excel.exe` chạy ngầm mồ côi và xóa các tệp `KangatangGuard_Worker_*.xlam` cũ trước khi đồng bộ tệp `.xlam` mới.
  - **Nếu ĐANG CÓ cửa sổ quét khác hoạt động ($\ge 1$ instance):** Tuyệt đối **không** gọi `wmic terminate` hay `taskkill /F /IM excel.exe`, không ghi đè tệp `.xlam` đang mở, và không ghi đè ngược `scanned_files.log` đang ghi dở.
- Khởi chạy cửa sổ PowerShell mới bằng lệnh `start "KangatangGuard Scanner" powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File ... -InteractivePrompt` để mỗi lần chạy `.bat` sẽ bật lên một cửa sổ độc lập.

### 3.2. Thiết kế `Per-Folder Mutex` & `PID-Scoped Worker Cleanup` trong [`Kangatang_FolderScanner.ps1`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1)
1. **Dọn dẹp Worker mồ côi thông minh theo `MasterPid`:**
   - Khi khởi động, quét danh sách tệp `%APPDATA%\KangatangGuard\KangatangGuard_Worker_<OldPid>_*.xlam`.
   - Kiểm tra xem `<OldPid>` còn sống hay không (`Get-Process -Id $OldPid`). Chỉ khi `<OldPid>` đã kết thúc mới xóa tệp `KangatangGuard_Worker_<OldPid>_*.xlam`.
   - Chỉ dọn `EXCEL.EXE` chạy ngầm (`MainWindowHandle -eq 0`) khi trên máy **không có bất kỳ tiến trình `powershell.exe` nào khác** đang chạy `Kangatang_FolderScanner.ps1`.
2. **Khóa `Per-Folder Mutex` sau khi xác định `$TargetFolder`:**
   - Sau khi người dùng chọn `$TargetFolder` từ hộp thoại (hoặc nhận từ tham số dòng lệnh), chuẩn hóa đường dẫn `$TargetFolder` về chữ thường và tính mã băm SHA256 ngắn (`$folderHash`).
   - Tạo Mutex định danh riêng cho thư mục đó: `Local\KangatangGuard_Folder_$folderHash`.
   - Nếu `WaitOne(0, $false)` trả về `$false` (tức là chính thư mục này đang được quét ở một cửa sổ khác), hiển thị hộp thoại thông báo: *"Thư mục này hiện đang được quét ở một cửa sổ PowerShell khác. Vui lòng chọn thư mục khác!"* và đóng cửa sổ mới an toàn.
3. **Định danh phiên độc lập (`SessionID`) & Cân bằng tải Worker (`Adaptive Worker Budget`):**
   - Định dạng `SessionID`: `SCAN_yyyyMMdd_HHmmss_<COMPUTERNAME>_P<PID>` (đảm bảo mỗi cửa sổ có 1 dòng riêng trong `Scan_Summary_History.csv` và 1 tệp báo cáo riêng `Sessions\Report_<SessionID>_<Folder>.md`).
   - Đếm số cửa sổ `Kangatang_FolderScanner.ps1` đang mở (`$activeScanners`):
     - Nếu chỉ có `1` cửa sổ: cấp tối đa `6 Workers` (hoặc `4`/`2` tùy cấu hình phần cứng).
     - Nếu có `2` cửa sổ chạy song song: mỗi cửa sổ mới dùng tối đa `3 Workers` (trên máy 16 luồng) để tổng số `EXCEL.EXE` không làm nghẽn RAM/CPU.
     - Nếu có $\ge 3$ cửa sổ chạy song song: mỗi cửa sổ mới dùng `2 Workers`.

---

## 4. Thiết kế Giao diện Console & Trải nghiệm Người dùng (Live Dashboard)
- Tiêu đề cửa sổ (`WindowTitle`) và khung đầu của **Live Dashboard** hiển thị rõ **`[PID: $PID | Thư mục: $safeLeafName]`** để người dùng dễ dàng phân biệt khi xếp 2–4 cửa sổ PowerShell cạnh nhau trên màn hình.
- Hộp thoại hỏi `Tiếp tục phiên quét (Resume)?` chỉ đếm và tác động lên các tệp thuộc nhánh `$TargetFolder` của cửa sổ đó. Nếu người dùng bấm `[NO]` (Quét mới từ đầu), chỉ loại bỏ các đường dẫn thuộc `$TargetFolder` khỏi `$resumeMap` của cửa sổ hiện tại và chỉ xóa các dòng thuộc `$TargetFolder` trong `scanned_files.log`, **bảo toàn 100%** tiến độ đã quét của các thư mục khác.

---

## 5. Chiến lược Ngoại lệ, Bảo mật & Toàn vẹn Dữ liệu (Cross-Process File Safety)
- 🛡️ **Named Log Mutex (`Local\KangatangGuard_LogWrite_Mutex`):**
  - Toàn bộ thao tác đọc-sửa-ghi (`Read-Modify-Write`) hoặc nối chuỗi (`AppendAllText`) vào các tệp dùng chung giữa các cửa sổ (`scanned_files.log`, `last_session.json`, `Scan_Summary_History.csv`, `Infected_Files_Ledger_YYYYMM.csv`, `scan_detail_log_YYYYMMDD.txt`, `LATEST_SCAN_REPORT.md`) được bọc trong hàm `Invoke-WithLogMutex`:
    - Xin quyền `WaitOne(3000)` trên `Local\KangatangGuard_LogWrite_Mutex`.
    - Thực thi ghi đĩa nguyên tử trong `< 5ms`.
    - Giải phóng `ReleaseMutex()` ngay trong khối `finally`.
- 🛡️ **Bảo vệ tiến trình đang chạy ở `99.9%` trên `223.7`:**
  - Chỉ cập nhật mã nguồn và triển khai khi phiên quét hiện tại đã hoàn tất (hoặc triển khai sao cho cửa sổ mới mở ra không chạm vào phiên đang chạy).

---

## 6. Kế hoạch Kiểm thử & Vận hành An toàn
1. Kiểm tra cú pháp tĩnh (`Parser::ParseFile`) đảm bảo `0` lỗi cú pháp trên PowerShell 4.0 (`Windows Server 2012 R2`) và PowerShell 5.1.
2. Kiểm thử thực tế kịch bản mở **2 cửa sổ `Kangatang_FolderScanner.ps1` song song** trên 2 thư mục khác nhau:
   - Xác minh cửa sổ 1 không bị đóng khi mở cửa sổ 2.
   - Xác minh cả 2 cửa sổ cùng ghi vào `Scan_Summary_History.csv` và `scanned_files.log` với 2 `SessionID` riêng biệt mà không xung đột khóa file.
   - Xác minh khi đóng cửa sổ 2, các `EXCEL.EXE` Worker của cửa sổ 1 vẫn hoạt động bình thường.
3. Đồng bộ bản phát hành `v3.22.0` lên `\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang\`, `distribution\` và GitHub.

---

## 7. Câu hỏi Cần Làm Rõ (`Open Questions`)
- Thiết kế đã đầy đủ và tương thích ngược 100%. Khi mở nhiều cửa sổ cùng lúc, mặc định hệ thống sẽ tự động chia sẻ số lượng Excel COM Worker (`6` Workers cho 1 cửa sổ, `3` Workers/cửa sổ khi chạy 2 cửa sổ, `2` Workers/cửa sổ khi chạy $\ge 3$ cửa sổ) để máy chủ `223.7` luôn mượt mà. Bạn có muốn giữ mức phân bổ tự động này không?
