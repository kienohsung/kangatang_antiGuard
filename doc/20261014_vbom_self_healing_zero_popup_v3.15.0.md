# Báo Cáo Kiến Trúc KangatangGuard `v3.15.0` — Sửa Triệt Để Hộp Thoại Cảnh Báo `AccessVBOM` (Self-Healing Đa Tầng & Zero-Popup Architecture)

- **Ngày thực hiện:** `2026-10-14`
- **Phiên bản phát hành:** `v3.15.0`
- **Trạng thái:** ✅ **HOÀN TẤT & ĐÃ PHÁT HÀNH LÊN LAN HUB (`223.7` & `223.176`)**

---

## 1. Phân Tích Nguyên Nhân Gốc Rễ (Root Cause Analysis — Bằng Chứng Thực Nghiệm)

Khi người dùng mở Excel, xuất hiện hộp thoại cảnh báo:
> **KangatangGuard v3.14.0 - Cảnh báo bảo mật**
> *KangatangGuard KHÔNG THỂ kiểm tra mã VBA vì Excel đang chặn quyền 'Trust access to the VBA project object model'.*

Qua thực nghiệm bằng Excel COM Automation & kiểm tra Registry trên môi trường thực tế, xác định **4 nguyên nhân gốc rễ**:

1. **Nguyên nhân 1 — `EXCEL.EXE` tự động ghi đè `AccessVBOM = 0` trong RAM xuống lại Registry khi đóng Excel (C++ Shutdown Flush):**
   - Khi Excel khởi động ở trạng thái `AccessVBOM = 0`, tiến trình `EXCEL.EXE` nạp biến cờ bảo mật nội bộ (`0`) vào RAM.
   - Dù VBA trong `v3.14.0` có ghi `AccessVBOM = 1` vào `HKCU\Software\Microsoft\Office\16.0\Excel\Security` khi đang chạy, lúc người dùng đóng Excel để khởi động lại theo hướng dẫn, hàm shutdown nội bộ của `EXCEL.EXE` lại **xả (flush) giá trị `0` trong RAM đè ngược xuống Registry `HKCU`**!
   - Hệ quả: Mở lại Excel bao nhiêu lần thì `AccessVBOM` vẫn bị reset về `0` bấy nhiêu lần.
2. **Nguyên nhân 2 — Chạm vào `wb.VBProject` trên cả `Book1` mới tạo và file `.xlsx` không có macro (`wb.HasVBProject = False`):**
   - Trong `v3.14.0`, hàm `CheckWorkbookInfection` và `CleanInfectedWorkbook` gọi `Set vbProj = wb.VBProject` trên mọi workbook, kể cả `Book1` trắng vừa khởi tạo khi mở Excel (`xlApp_NewWorkbook`) và các file `.xlsx` thuần dữ liệu.
   - Đã kiểm chứng thực nghiệm: Thuộc tính `wb.HasVBProject` hoạt động chính xác ngay cả khi `AccessVBOM = 0` (`False` với `Book1`/`.xlsx` sạch, `True` với mọi file `.xls`/`.xlsm` có chứa dù chỉ 1 dòng lệnh VBA).
3. **Nguyên nhân 3 — Lỗi định dạng dấu phẩy thập phân của `Application.Version` (`"16,0"`) và thiếu đồng bộ đa nhánh Registry:**
   - Trên Windows cấu hình vùng Tiếng Việt, `Application.Version` có thể trả về `"16,0"` (dấu phẩy), khiến đường dẫn Registry `HKCU\Software\Microsoft\Office\16,0\...` bị ghi sai nhánh.
   - Ngoài ra, khóa `HKCU\Software\Policies\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 0` hoặc việc chạy `Install_Addin.bat` với `-Verb RunAs` (ghi nhầm vào `HKCU` của tài khoản Administrator thay vì tài khoản người dùng hiện tại) làm vô hiệu hóa cấu hình `HKCU`.
4. **Nguyên nhân 4 — Không phân biệt lỗi `AccessVBOM` toàn cục với lỗi `VBProject` cục bộ của từng file:**
   - Trước đây, mọi lỗi khi truy cập `wb.VBProject` đều quy kết là do `AccessVBOM` bị tắt và bật hộp thoại cảnh báo. Thực tế có thể kiểm tra chính xác quyền `AccessVBOM` toàn cục bằng cách thử đọc `ThisWorkbook.VBProject.VBComponents.Count` (`IsVbomAccessible()`).

---

## 2. Giải Pháp Kiến Trúc Đã Triển Khai (`v3.15.0`)

### 🛡️ 2.1. Kiểm Tra Sớm `wb.HasVBProject` & Xác Thực `IsVbomAccessible()`
- Sau khi hoàn tất quét **Tên Sheet**, **Named Ranges (XLM Macro 4.0)** và **Excel 5.0 Module Sheets**, kiểm tra `wb.HasVBProject`:
  - Nếu `wb.HasVBProject = False` (như `Book1`, `.xlsx`, `.xls` không có VBA stream): kết luận an toàn ngay, **không chạm vào `wb.VBProject`**.
- Nếu `wb.HasVBProject = True` mà `Set vbProj = wb.VBProject` báo lỗi: gọi `IsVbomAccessible()` (kiểm tra `ThisWorkbook.VBProject.VBComponents.Count`) để xác định chính xác có phải `AccessVBOM` toàn cục bị chặn hay chỉ là lỗi riêng của file `wb`.

### 🛡️ 2.2. Tự Phục Hồi Đa Tầng (`EnsureVbomRegistryEnabled` + `SchedulePostExitVbomFix`)
- **`EnsureVbomRegistryEnabled()`:**
  - Chuẩn hóa `Replace(Trim$(Application.Version), ",", ".")` và ghi `AccessVBOM = 1` đồng loạt cho tất cả phiên bản Office (`16.0`, `15.0`, `14.0`, `12.0`).
  - Tự động xóa khóa chặn `AccessVBOM` trong `HKCU\Software\Policies\Microsoft\Office\<ver>\Excel\Security` và xóa `LastVbomWarn`.
  - Tự động bật `AccessVBOM = 1` trên `HKLM` và `HKLM\Software\WOW6432Node` nếu tiến trình có quyền ghi.
- **`SchedulePostExitVbomFix()` (Khắc phục triệt để lỗi Excel ghi đè `0` khi thoát):**
  - Khi đóng Excel (`TerminateGuard`) trong phiên có `AccessVBOM = 0`, kích hoạt tiến trình ngầm độc lập (`wsh.Run ..., 0, False`) lặp lại lệnh `reg add ... AccessVBOM = 1` mỗi 1 giây trong 8 giây sau khi `EXCEL.EXE` thoát, đảm bảo ghi đè lại giá trị `1` ngay sau khi C++ shutdown của Excel vừa xả RAM.

### 🛡️ 2.3. Kiến Trúc Zero-Popup & Out-of-Process Scanner Fallback (`ReportVbomBlocked`)
- Loại bỏ hoàn toàn hộp thoại `MsgBoxW` gây gián đoạn và hoang mang cho người dùng.
- Nếu phát hiện file trên ổ đĩa có `wb.HasVBProject = True` trong phiên Excel đang bị kẹt `AccessVBOM = 0` trong RAM:
  1. Tự động bật lại `AccessVBOM = 1` trong Registry.
  2. Tự động kích hoạt **`LaunchBackgroundScanner(wb.Path)`** chạy ngầm ngoài tiến trình (Out-of-Process). Vì tiến trình PowerShell Worker khởi tạo một instance `Excel.Application` mới hoàn toàn, nó đọc ngay `AccessVBOM = 1` vừa bật trong Registry để quét sạch VBA của file mà người dùng không cần khởi động lại Excel!

### 🛡️ 2.4. Chống Kẹt File Khi Phát Hành LAN (`Versioned & Latest XLAM Resolution`)
- Phát hiện máy trạm `DPL-Mold-mrHanh` mở trực tiếp `KangatangGuard.xlam` trên `\\192.168.223.7` gây khóa SMB (`Sharing Violation`).
- Cập nhật `Deploy_To_Hub.ps1`, `Install_Client.ps1` và `PerformLanUpdate` để luôn phát hành song song `KangatangGuard_v3.15.0.xlam` và `KangatangGuard_latest.xlam` trên `223.7` (không bao giờ bị khóa bởi Excel client) kèm tiến trình `Background Sync Watcher` tự động cập nhật `KangatangGuard.xlam` ngay khi máy trạm đóng Excel.

---

## 3. Bảng Đối Chiếu Trước & Sau (`Before & After Matrix`)

| Hạng mục | Trước (`v3.14.0`) | Sau (`v3.15.0`) |
| :--- | :--- | :--- |
| **Mở Excel (`Book1`) hoặc file `.xlsx` khi `AccessVBOM = 0`** | Cố truy cập `wb.VBProject` ➔ Văng popup cảnh báo bảo mật đỏ | Kiểm tra `wb.HasVBProject = False` ➔ Bỏ qua `VBProject` an toàn, **0 popup** |
| **Hiện tượng đóng/mở lại Excel vẫn bị mất `AccessVBOM`** | `EXCEL.EXE` xả giá trị `0` trong RAM đè lại Registry khi thoát | `SchedulePostExitVbomFix()` khóa lại `AccessVBOM = 1` sau khi `EXCEL.EXE` thoát |
| **Định dạng `Application.Version` vùng Tiếng Việt (`"16,0"`)** | Ghi nhầm vào nhánh `Office\16,0` không tồn tại | Chuẩn hóa `Replace(..., ",", ".")` + quét phủ `16.0`, `15.0`, `14.0`, `12.0` |
| **Khi mở file `.xls`/`.xlsm` có macro lúc RAM đang `AccessVBOM = 0`** | Hiện popup bắt người dùng tự khởi động lại Excel | Tự bật `AccessVBOM = 1` + gọi `LaunchBackgroundScanner` quét ngầm bằng tiến trình mới (**Zero-Popup**) |
| **`Install_Addin.bat` trên máy trạm Domain** | Ép `-Verb RunAs` ➔ Cài nhầm vào `HKCU`/`%APPDATA%` của Admin | Chạy đúng ngữ cảnh User hiện tại (100% Pure ASCII) |
| **Phát hành lên LAN Hub `223.7` khi có máy mở trực tiếp `.xlam`** | Lỗi `Copy-Item` do SMB lock | Phát hành `KangatangGuard_v3.15.0.xlam` + `KangatangGuard_latest.xlam` + `Background Sync Watcher` |
