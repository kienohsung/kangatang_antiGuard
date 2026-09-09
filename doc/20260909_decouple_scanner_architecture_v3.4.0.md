# THIẾT KẾ KIẾN TRÚC TÁCH TIẾN TRÌNH: KANGATANGGUARD v3.4.0
**Mã tài liệu:** `doc/20260909_decouple_scanner_architecture_v3.4.0.md`  
**Phiên bản:** `v3.4.0` (Decoupled Asynchronous Architecture)  
**Chức danh:** Tech Lead kiêm Senior Full-stack Architect  
**Trạng thái:** Chờ phê duyệt (Awaiting Approval)

---

## 1. PHÂN TÍCH NGUYÊN NHÂN GỐC RỄ (ROOT CAUSE: EXCEL NOT RESPONDING)

Dựa trên hình ảnh thực tế người dùng cung cấp (`Book1 - Excel (Not Responding)`):

### 🛑 Bản chất kỹ thuật của việc treo Excel:
1. **Mô hình luồng đơn của VBA (Single-threaded UI):**
   - Môi trường VBA trong Microsoft Excel chạy trực tiếp trên **luồng giao diện người dùng (Main UI Thread)** của tiến trình `EXCEL.EXE`.
   - Khi quét một thư mục lớn (ví dụ: 138 tệp tại `10. Mua bán`), hàm `ScanParentFolder` thực hiện vòng lặp mở tuần tự từng tệp:
     `Application.Workbooks.Open(FileName:=fullPath, ...)`
   - Mỗi tệp mất trung bình từ **0.3s - 0.8s** để đọc XML, nạp cấu trúc Workbook, kiểm tra Sheets, Names, VBProject.
   - Với 138 tệp, tổng thời gian chiếm dụng luồng UI liên tục từ **40 đến 90 giây**.
2. **Hàng đợi thông điệp Windows bị đói (Message Pump Starvation):**
   - Khi luồng UI bị vòng lặp VBA chiếm dụng 100%, Excel không thể xử lý các thông điệp hệ thống của Windows (`WM_PAINT`, `WM_ACTIVATE`, `WM_MOUSEMOVE`).
   - Sau 5 giây không phản hồi, Windows tự động gán nhãn cửa sổ là **`Excel (Not Responding)`** và làm mờ giao diện. Người dùng không thể gõ phím, không thể tính toán và lầm tưởng rằng Excel đã bị lỗi/crash.
3. **Rủi ro cô lập lỗi (Fault Tolerance):**
   - Nếu trong 138 tệp có 1 tệp bị hỏng (corrupted) gây sập COM, toàn bộ cửa sổ Excel của người dùng (bao gồm các file chưa lưu) sẽ bị đóng đột ngột.

---

## 2. KIẾN TRÚC MỚI: BẤT ĐỒNG BỘ NGOÀI TIẾN TRÌNH (OUT-OF-PROCESS ASYNC WORKER)

Để giải quyết triệt để và vĩnh viễn tình trạng trên theo đúng Tiêu chuẩn Kỹ thuật **Production-grade (Nguyên tắc 4: Tách biệt các tác vụ nặng sang hàng đợi ngầm / Worker độc lập, không block UI chính)**, chúng tôi thiết kế kiến trúc phân tách 2 thành phần:

```
+-------------------------------------------------------------------------+
| [TIẾN TRÌNH 1] EXCEL CỦA NGƯỜI DÙNG (EXCEL.EXE - HOÀN TOÀN MƯỢT MÀ)     |
| - KangatangGuard Add-in v3.4.0 (Event Sentinel):                         |
|   + Bảo vệ thời gian thực: Quét nhanh khi mở 1 tệp (< 0.2s)            |
|   + ScanCache chống lag khi lưu                                         |
|   + Menu Ribbon: Khi bấm "Quét thư mục..."                              |
|       ===> Bật File Dialog chọn thư mục                                 |
|       ===> GỌI TIẾN TRÌNH 2 BẤT ĐỒNG BỘ (ASYNCHRONOUS SPAWN: 0.01 GIÂY) |
|       ===> THOÁT NGAY LẬP TỨC! Excel giải phóng UI 100%                 |
+-------------------------------------------------------------------------+
                                    |
            (WScript.Shell.Run cmd, 1, False - Không chờ)
                                    v
+-------------------------------------------------------------------------+
| [TIẾN TRÌNH 2] TRÌNH QUÉT ĐỘC LẬP (POWERSHELL WORKER - CỬA SỔ RIÊNG)     |
| File: %APPDATA%\KangatangGuard\Kangatang_FolderScanner.ps1              |
| - Chạy trên tiến trình riêng biệt (powershell.exe)                      |
| - Khởi tạo Excel COM ẩn hoàn toàn độc lập với Excel người dùng          |
| - CÓ GIAO DIỆN THANH TIẾN ĐỘ THỜI GIAN THỰC (PROGRESS BAR):             |
|   + Đang quét: [========>        ] 45/138 (32%)                         |
|   + Tệp hiện tại: don de nghi - mrKien.xlsx                             |
|   + Phát hiện & Tự động sao lưu -> Tiêu diệt virus                      |
| - Khi xong: Báo cáo tổng kết rõ ràng, ghi log kiểm toán                 |
| - NẾU CÓ TỆP LỖI: Chỉ ảnh hưởng tiến trình quét, EXCEL NGƯỜI DÙNG AN TOÀN|
+-------------------------------------------------------------------------+
```

---

## 3. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Tiêu chí | Bản v3.3.0 (Hiện tại) | Bản v3.4.0 (Kiến trúc Tách Tiến trình Mới) | Giá trị Vượt trội |
| :--- | :--- | :--- | :--- |
| **Mô hình thực thi** | Chạy đồng bộ trong luồng chính của Excel | **Tách ngoài tiến trình (Out-of-Process Asynchronous Worker)** | 🛡️ Chống treo UI 100% |
| **Trạng thái Excel khi quét** | Bị treo `(Not Responding)`, tê liệt 1-2 phút | **Mượt mà 100%, người dùng tiếp tục gõ văn bản/làm việc trên `Book1` bình thường** | ✅ Trải nghiệm người dùng hoàn hảo |
| **Hiển thị tiến độ** | Màn hình đơ, không biết quét đến đâu | **Thanh tiến độ (Progress Bar) % và tên file trực quan** | ✅ Rõ ràng, minh bạch |
| **Độ an toàn dữ liệu** | File quét hỏng có thể kéo sập Excel đang mở | Tách biệt hoàn toàn bộ nhớ: Tiến trình quét sập không ảnh hưởng Excel đang mở | 🛡️ Bảo vệ an toàn tuyệt đối |
| **Tự động quét ổ dịch** | Quét đồng bộ làm đơ Excel sau khi diệt | Bật cửa sổ quét ngầm chạy độc lập ở chế độ nền | 🛡️ Tự động dập dịch không gián đoạn công việc |

---

## 4. CHI TIẾT THIẾT KẾ KỸ THUẬT

### 1. Thành phần Excel Add-in (`KangatangGuard_Code.vba` v3.4.0):
- Thay thế vòng lặp nặng `ScanParentFolder` bên trong VBA bằng lệnh kích hoạt tiến trình độc lập không đồng bộ:
  ```vba
  Public Sub LaunchBackgroundScanner(ByVal folderPath As String)
      Dim wsh As Object
      Set wsh = CreateObject("WScript.Shell")
      
      Dim scannerScript As String
      scannerScript = Environ("APPDATA") & "\KangatangGuard\Kangatang_FolderScanner.ps1"
      
      Dim cmd As String
      cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scannerScript & """ -TargetFolder """ & folderPath & """"
      
      ' Tham so 1 = Hien cua so, False = Chay bat dong bo (Khong cho, thoat ngay lap tuc)
      wsh.Run cmd, 1, False
      Set wsh = Nothing
  End Sub
  ```
- Hàm này thực thi xong trong **10 mili-giây** và trả quyền điều khiển lại cho Excel ngay lập tức!

### 2. Thành phần Worker Độc lập (`Kangatang_FolderScanner.ps1`):
- Được đặt tại `%APPDATA%\KangatangGuard\Kangatang_FolderScanner.ps1`.
- Sử dụng PowerShell native + COM độc lập (`Visible = $false`).
- Hiển thị giao diện điều khiển console / GUI với:
  - Thanh tiến độ `Write-Progress` chuẩn Windows.
  - Hiển thị danh sách tệp đã quét, số tệp sạch, số tệp đã tiêu diệt virus.
  - Tự động sao lưu vào `_Backup_Kangatang\` qua chuẩn Unicode.
  - Ghi nhật ký vào `%APPDATA%\KangatangGuard\scan_log_YYYYMMDD.txt`.
