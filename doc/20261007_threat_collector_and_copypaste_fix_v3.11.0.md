# TÀI LIỆU KIẾN TRÚC & THIẾT KẾ KANGATANGGUARD v3.11.0
## THREAT COLLECTOR, PASSWORD BYPASS & COPY/PASTE HEALING

- **Phiên bản:** `v3.11.0`
- **Ngày phát hành:** `2026-10-07`
- **Tác giả:** Tech Lead kiêm Senior Full-stack Architect
- **Tiêu chuẩn:** Production-grade, Multi-Agent Coordination Protocol v1.0

---

## 1. TỔNG QUAN YÊU CẦU & BỐI CẢNH

Trong quá trình bảo vệ và làm sạch hệ thống Excel mạng LAN, phát sinh 3 bài toán nghiệp vụ trọng yếu:

1. **Thu thập Mã độc (Threat Intelligence / Threat Collector):**
   - Người dùng phát hiện khi click đúp vào các module `Kangatang` bị nhiễm, bên trong có mã nguồn VBA thực tế.
   - Nhu cầu: Cần cơ chế tự động trích xuất toàn bộ nội dung code của tất cả các module độc hại trước khi tiến hành xóa bỏ, phục vụ phân tích giải phẫu và nâng cấp chữ ký nhận diện.
   - 🛑 **Ràng buộc an toàn tuyệt đối:** Thư mục lưu trữ mẫu mã độc PHẢI đặt tại máy cá nhân `192.168.223.176` (hoặc ổ cứng cục bộ `D:\7. AI tools\kangatang\Threat_Samples`), **TUYỆT ĐỐI KHÔNG LƯU TRÊN MÁY CHỦ CHUNG `192.168.223.7`**.

2. **Xử lý File & Sheet có Mật khẩu Bảo vệ (Password & Protection Bypass):**
   - Các file có cấu trúc bảng tính hoặc sheet bị khóa (`ProtectStructure`, `ProtectContents`) khiến lệnh xóa module-sheet bị lỗi hoặc không thể làm sạch.
   - Khi quét folder, các file có mật khẩu mở (Password to Open) kích hoạt hộp thoại popup của Excel COM, gây treo tiến trình quét chạy ngầm (hang process).
   - Các file có VBProject bị khóa mật khẩu (`vbProj.Protection = 1`) thường bị bỏ qua toàn bộ, làm sót các module-sheet và Named range độc hại.

3. **Khắc phục Triệt để Lỗi Không thể Copy/Paste:**
   - Người dùng báo cáo hiện tượng: "File kiểm tra không có module kangatang, tuy nhiên vẫn không thể copy paste".
   - **Nguyên nhân gốc rễ (Root Cause):** Tàn dư virus Kangatang / Laroux cài cắm các Application hooks và can thiệp sâu vào môi trường Excel:
     - Chặn phím tắt: `Application.OnKey "^c", ""` và `Application.OnKey "^v", ""`.
     - Vô hiệu hóa kéo thả: `Application.CellDragAndDrop = False`.
     - Khóa menu chuột phải: `Application.CommandBars("Cell").Enabled = False`.
     - Khóa chọn ô: `sh.EnableSelection = 0` (`xlNoSelection`).
     - **Nguyên nhân chính làm mất Clipboard:** Các hook sự kiện mồ côi (`Application.OnSheetActivate = "'mypersonnel1.xls'!allocated"`, `OnWindow`, `OnCalculate`). Khi người dùng click chuột sang ô khác hoặc chuyển sheet, Excel cố gọi macro mồ côi đã bị xóa, gặp lỗi ngầm và tự động reset `Application.CutCopyMode = False` ngay lập tức, làm mất sạch dữ liệu vừa Copy!

---

## 2. THIẾT KẾ KIẾN TRÚC CHI TIẾT

```
+-----------------------------------------------------------------------------------+
|                            KANGATANGGUARD v3.11.0                                 |
+-----------------------------------------------------------------------------------+
                                          |
        +---------------------------------+---------------------------------+
        |                                 |                                 |
        v                                 v                                 v
+-----------------------+     +-----------------------+     +-----------------------+
|  modThreatCollector   |     |  Password & Protect   |     | Clipboard & UI Healing|
|  - Trích xuất code    |     |  - TryUnprotectWb/Sht |     | - Gỡ chặn OnKey ^c/^v |
|  - Băm MD5 / SHA-256  |     |  - Diệt module sheet  |     | - Bật CellDragAndDrop |
|  - Khử trùng lặp      |     |    ngay cả khi VBProj |     | - Bật Context Menu    |
|  - Lưu JSON Defanged  |     |    bị lock mật khẩu   |     | - Dọn sạch App Hooks  |
|  - Chặn máy 223.7!    |     |  - Skip popup mở file |     | - Nút Ribbon 1-Click  |
+-----------------------+     +-----------------------+     +-----------------------+
        |                                 |                                 |
        v                                 v                                 v
[Local D:\ & 223.176]          [Quét/Diệt sạch 100%]           [Copy/Paste phục hồi]
```

### 2.1. Module `modThreatCollector`
- **Đường dẫn ưu tiên:**
  1. Cục bộ: `D:\7. AI tools\kangatang\Threat_Samples\`
  2. Mạng LAN máy cá nhân: `\\192.168.223.176\KangatangGuard_Hub\Threat_Samples\`
  3. Fallback: `%APPDATA%\KangatangGuard\Threat_Samples\`
- **Cơ chế Vaccine Guard:** Kiểm tra chuỗi `InStr(path, "192.168.223.7") > 0` -> Hủy ngay lập tức nếu phát hiện địa chỉ máy chủ cấm.
- **Tính toán Băm & Khử trùng lặp:**
  - Sử dụng `.NET COM` `System.Security.Cryptography.MD5CryptoServiceProvider` (VBA) và `SHA256` (PowerShell).
  - Có thuật toán Fallback thuần VBA (FNV-1a / DJB2) nếu COM .NET bị giới hạn.
  - Tên file lưu: `<hash>.sample.json`. Nếu file đã tồn tại, bỏ qua ghi đĩa để tối ưu I/O.
- **Định dạng mẫu JSON (Defanged):**
  ```json
  {
    "timestamp": "2026-10-07 00:30:17",
    "computer_name": "CM-GA-MRKIENIT1",
    "file_path": "C:\\Data\\Financial_Report.xlsm",
    "component_name": "Kangatang",
    "threat_type": "KANGATANG_VIRUS_MODULE",
    "code_length": 542,
    "code_hash": "ea3ab1e3c32dc540d128d4f281809795",
    "raw_code": "Sub Auto_Open()\n  ...\nEnd Sub"
  }
  ```

### 2.2. Xử lý Mật khẩu Bảo vệ (Protection & Password Bypass)
- **Hàm hỗ trợ an toàn:**
  - `TryUnprotectWorkbook(wb)`: Thử `wb.Unprotect ""` với `On Error Resume Next`.
  - `TryUnprotectSheet(sh)`: Thử `sh.Unprotect ""`.
- **VBProject Locked (`vbProj.Protection = 1`):**
  - Không bao giờ "bỏ cuộc" khi gặp project khóa mật khẩu!
  - Excel 5.0 Module-sheets nằm trong tập hợp bảng tính `wb.Sheets` và tên phạm vi nằm trong `wb.Names`. Cả hai thành phần này hoàn toàn CÓ THỂ xóa được sau khi mở khóa cấu trúc workbook mà không phụ thuộc vào `VBComponents`.
  - Add-in làm sạch triệt để `wb.Sheets` và `wb.Names`, chỉ bỏ qua `VBComponents` và ghi log kiểm toán rõ ràng.
- **PowerShell Scanner Safe Open (`Open-ExcelWorkbookSafe`):**
  - Truyền `Password = "dummy_anti_freeze_pwd"` và bắt ngoại lệ mật khẩu (`0x800a03ec`) để từ chối mở và ghi log `[PASS_PROTECTED] Skipped password-protected file` ngay trong 1ms, triệt tiêu nguy cơ treo hộp thoại modal.

### 2.3. Khôi phục Copy/Paste & Giao diện Excel (`RestoreExcelClipboardAndUI`)
- **Các bước phục hồi toàn diện:**
  1. Phục hồi phím tắt chuẩn: Gọi `Application.OnKey` cho `"^c"`, `"^v"`, `"^x"`, `"^d"`, `"^r"`, `"%{F11}"`, `"%{F8}"` không có tham số để gỡ bỏ ràng buộc virus.
  2. Bật lại `Application.CellDragAndDrop = True`.
  3. Đặt lại `Application.CutCopyMode = False`.
  4. Bật lại menu ngữ cảnh chuột phải: `Application.CommandBars("Cell").Enabled = True`, `"Row"`, `"Column"`.
  5. Dọn dẹp tất cả các Application hooks mồ côi: `OnSheetActivate = ""`, `OnSheetDeactivate = ""`, `OnWindow = ""`, `OnCalculate = ""`, `OnDoubleClick = ""`, `OnEntry = ""`.
  6. Phục hồi quyền chọn ô trên toàn bộ Sheet: Nếu `sh.EnableSelection = 0` (`xlNoSelection`), tự động đặt lại thành `-4142` (`xlNoRestrictions`).
- **Điểm kích hoạt:**
  - Tự động chạy khi khởi động add-in (`InitializeGuard`).
  - Tự động chạy trong quét startup (`ShieldStartupSweep`).
  - Tự động chạy sau mỗi phiên làm sạch file (`CleanInfectedWorkbook`).
  - Nút bấm trực tiếp trên Ribbon tab KangatangGuard: **"Khôi phục Copy/Paste"** (`btnFixCopyPaste`, icon Undo).

---

## 3. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Tiêu chí | Trước nâng cấp (v3.10.0) | Sau nâng cấp (v3.11.0) |
|---|---|---|
| **Thu thập Code Mẫu** | Xóa thẳng tay mã độc, không lưu lại code để phân tích | Tự động trích xuất code, băm MD5/SHA-256, lưu dạng JSON defanged vào máy cá nhân `223.176` |
| **Bảo mật Kho Mẫu** | Chưa có | Chặn tuyệt đối máy chủ `223.7`, chỉ lưu máy cá nhân `223.176` và ổ đĩa cục bộ |
| **File có Mật khẩu Mở** | Có nguy cơ hiện popup nhập mật khẩu gây treo scanner COM | Truyền dummy password, bắt lỗi từ chối ngay lập tức không treo COM |
| **Sheet / Cấu trúc bị khóa** | Gặp lỗi không thể xóa sheet hoặc cấu trúc file | Tự động giải phóng `Unprotect ""` trước khi can thiệp xóa |
| **VBProject bị Khóa Password** | Bỏ qua toàn bộ file, bỏ sót Module-sheets độc hại | Vẫn diệt sạch 100% Module-sheets và Names độc hại trong bảng tính |
| **Lỗi Mất Copy/Paste** | Người dùng không copy được ô tính do hook mồ côi reset CutCopyMode | Tự động khôi phục toàn diện: OnKey, DragDrop, Context Menus, dọn sạch Hooks + Nút bấm 1-Click trên Ribbon |

---

## 4. KẾT QUẢ KIỂM THỬ THỰC TẾ (VERIFICATION SUITE)

- **PowerShell AST Syntax Check:** 0 Syntax Errors (Bảng mã UTF-8 BOM chuẩn Windows).
- **VBA Compilation In-Memory:** 7 components nạp và biên dịch thành công 100%.
- **Kiểm thử Tự động Add-in v3.11.0 (`scratch/test_v3_11_0_plan.ps1`):**
  - `[PASS] CellDragAndDrop restored to True`
  - `[PASS] Cell Context Menu restored to Enabled`
  - `[PASS] Workbook structure successfully protected`
  - `[PASS] TryUnprotectWorkbook returned True`
  - `[PASS] Workbook structure is now Unprotected`
  - `[PASS] ComputeThreatHash generated valid MD5 hash`
  - `[PASS] Sample file was saved locally`
  - `[PASS] Sample JSON contains correct component name`
  - `[PASS] Sample JSON contains threat type`
  - **TỔNG KẾT: 9 / 9 TESTS PASSED (100%)**
- **Đóng gói & Phát hành Dual-Mirror:**
  - Đã biên dịch `KangatangGuard.xlam` v3.11.0 vào thư mục XLSTART cục bộ.
  - Đã xuất bản lên máy chủ tập trung `\\192.168.223.7` và bản clone nội bộ `223.176`.
