# TÀI LIỆU KIẾN TRÚC BẢN VÁ v3.5.4: KHẮC PHỤC TRIỆT ĐỂ LỖI CỬA SỔ POWERSHELL CHỚP TẮT KHI BẤM QUÉT THƯ MỤC

**Ngày thực hiện:** 09/09/2026  
**Phiên bản:** `v3.5.4`  
**Mục tiêu:** Đảm bảo tính năng "Quét toàn bộ thư mục chỉ định" từ menu Add-in trong Excel khởi chạy trình quét ngầm PowerShell mượt mà, cửa sổ console hiển thị ổn định, không bị chớp nhoáng rồi tắt.

---

## 1. PHÂN TÍCH NGUYÊN NHÂN GỐC RỄ (ROOT CAUSE ANALYSIS)

Khi người dùng nhấn chọn **"Quét toàn bộ thư mục chỉ định"** trên thanh Ribbon / Menu của KangatangGuard, cửa sổ PowerShell bật lên một phần tích tắc rồi lập tức tắt. Qua điều tra kỹ thuật chuyên sâu, nguyên nhân bao gồm 3 điểm mấu chốt:

### 1.1. Tệp `Kangatang_FolderScanner.ps1` chưa được triển khai vào `%APPDATA%\KangatangGuard\`
- Trong mã nguồn VBA của Add-in, hàm `LaunchBackgroundScanner` trỏ đến đường dẫn mặc định:
  `scannerScript = Environ("APPDATA") & "\" & LOG_SUBFOLDER & "\Kangatang_FolderScanner.ps1"`
- Tuy nhiên, trong môi trường thực tế, thư mục `%APPDATA%\KangatangGuard\` chỉ mới chứa các tệp nhật ký `scan_log_*.txt` mà chưa có tệp script `Kangatang_FolderScanner.ps1`. Thư mục `XLSTART` cũng chỉ có `KangatangGuard.xlam`.
- Khi PowerShell khởi động với tham số `-File "<đường_dẫn_không_tồn_tại>"`, PowerShell báo lỗi:
  `The argument '...Kangatang_FolderScanner.ps1' to the -File parameter does not exist.`
  Do không có cờ giữ cửa sổ, PowerShell lập tức hủy tiến trình và đóng cửa sổ console.

### 1.2. Lỗi Command Line Escaping `\"` trong Windows API
- Trong hàm `ScanFolderDialog` và `LaunchBackgroundScanner`, mã nguồn VBA có đoạn:
  ```vba
  If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
  ```
- Sau đó ghép lệnh:
  ```vba
  cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & folderPath & """"
  ```
- Kết quả tạo ra dòng lệnh kết thúc bằng: `-TargetFolder "D:\Folder\"`
- Theo đặc tả của trình phân tích dòng lệnh chuẩn Windows (`CommandLineToArgvW`), ký tự `\"` được xem là ký tự thoát (escaped literal quote), chứ không phải dấu đóng chuỗi!
- Vì vậy, PowerShell nhận diện tham số bị sai định dạng, gây ra lỗi:
  `PositionalParameterNotFound: A positional parameter cannot be found that accepts argument...`
  và crash ngay khi vừa khởi chạy.

### 1.3. Thiếu cờ `-NoExit` trong lệnh chạy PowerShell
- Lệnh gọi chạy PowerShell trong VBA trước đây chỉ có `-NoProfile -ExecutionPolicy Bypass -File ...`.
- Khi có bất kỳ lỗi cú pháp, lỗi thiếu tệp, hoặc khi tiến trình kết thúc, Windows sẽ tự động đóng ngay cửa sổ console, khiến người dùng không thể đọc được thông báo lỗi hay theo dõi kết quả.

---

## 2. THIẾT KẾ BẢN VÁ PHÒNG THỦ (DEFENSIVE DESIGN v3.5.4)

### 2.1. Cơ chế Tự Phục Hồi Tệp Scanner (Self-Healing File Fallback)
Trong VBA `LaunchBackgroundScanner`:
1. Kiểm tra sự tồn tại của `Kangatang_FolderScanner.ps1` tại `%APPDATA%\KangatangGuard\`.
2. Nếu chưa có, lần lượt tìm kiếm tại các vị trí:
   - `ThisWorkbook.Path & "\Kangatang_FolderScanner.ps1"` (thư mục chứa Add-in đang chạy)
   - `%APPDATA%\Microsoft\Excel\XLSTART\Kangatang_FolderScanner.ps1`
   - `%APPDATA%\Microsoft\AddIns\Kangatang_FolderScanner.ps1`
3. Nếu tìm thấy ở bất kỳ vị trí nào, tự động sao chép sang `%APPDATA%\KangatangGuard\Kangatang_FolderScanner.ps1`.
4. Nếu vẫn không tìm thấy: Hiển thị thông báo tiếng Việt có dấu cảnh báo người dùng và dừng lại, tuyệt đối không gọi lệnh PowerShell rỗng gây lỗi.

### 2.2. Chuẩn hóa Tham số Đường dẫn (Sanitize Path Trailing Slashes)
Trong VBA:
- Cắt bỏ triệt để toàn bộ dấu gạch chéo ngược `\` ở cuối chuỗi `folderPath` trước khi bọc dấu nháy kép truyền vào CLI.
- Trong `Kangatang_FolderScanner.ps1`, logic đã có sẵn kiểm tra ổ đĩa gốc `^[a-zA-Z]:$` để tự bổ sung `\` an toàn.

### 2.3. Bổ sung cờ `-NoExit` cho Trình Quét Ngầm
- Chuyển đổi lệnh gọi sang:
  `cmd = "powershell.exe -NoExit -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & cleanFolder & """"`
- Giúp cửa sổ PowerShell luôn hiển thị trực quan, hỗ trợ người dùng theo dõi tiến độ quét và báo cáo kết quả.

### 2.4. Tự Động Phân Phối Tệp Scanner Trong Bộ Cài Đặt
Trong `Install_KangatangGuard.ps1`:
- Tự động sao chép `Kangatang_FolderScanner.ps1` vào cả 3 thư mục:
  1. `%APPDATA%\KangatangGuard\`
  2. `%APPDATA%\Microsoft\Excel\XLSTART\`
  3. `%APPDATA%\Microsoft\AddIns\`
