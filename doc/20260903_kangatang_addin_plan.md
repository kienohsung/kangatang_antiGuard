# BÁO CÁO RÀ SOÁT KIẾN TRÚC & KẾ HOẠCH TRIỂN KHAI EXCEL ADD-IN KANGATANGGUARD
**Mã tài liệu:** `doc/20260903_kangatang_addin_plan.md`  
**Phiên bản kiến trúc:** `v3.1.0`  
**Chức danh phê duyệt:** Senior Full-stack Architect / Tech Lead  
**Trạng thái:** Sẵn sàng nghiệm thu kỹ thuật (Technical Audit Ready)

---

## 1. TỔNG QUAN HỆ THỐNG & MỤC TIÊU V3.1.0

Hệ thống phòng chống và tiêu diệt virus macro **Kangatang / Laroux / mypersonnel** tiến hóa từ mô hình quét theo yêu cầu (On-demand) sang mô hình **Bảo vệ Đa tầng Kết hợp (Hybrid Multi-Layer Defense)**:

```
+-----------------------------------------------------------------------------------+
|                           KIẾN TRÚC BẢO VỆ ĐA TẦNG v3.1.0                         |
+-----------------------------------------------------------------------------------+
|  TẦNG 1: THỜI GIAN THỰC (Real-time Guard)                                         |
|  -> Excel Add-in (KangatangGuard.xlam) đặt trong XLSTART                          |
|  -> Tự động hook Application.WorkbookOpen để kiểm tra ngay khi mở file            |
+-----------------------------------------------------------------------------------+
|  TẦNG 2: QUÉT TOÀN DIỆN HỆ THỐNG THEO YÊU CẦU (On-Demand Deep Clean)             |
|  -> Script Polyglot Diet_Virus_Kangatang.bat v3.1.0                               |
|  -> Quét sâu toàn bộ AppData của mọi User Profile, XLSTART hệ thống, và thư mục   |
+-----------------------------------------------------------------------------------+
|  TẦNG 3: BẢO VỆ DỮ LIỆU & KIỂM TOÁN (Data Integrity & Audit Log)                  |
|  -> Tự động Snapshot Backup trước khi xóa module/sheet nhiễm                      |
|  -> Ghi log kiểm toán truy vết tại %APPDATA%\KangatangGuard                       |
+-----------------------------------------------------------------------------------+
```

---

## 2. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Tiêu chí Đánh giá | Phiên bản v3.0.0 (Hiện tại) | Phiên bản v3.1.0 (Kiến trúc Mới) | Đánh giá Giá trị |
| :--- | :--- | :--- | :--- |
| **Cơ chế Kích hoạt** | Quét thủ công 1-Click khi nghi ngờ có file nhiễm | **Bảo vệ thời gian thực:** Tự động quét và diệt ngay khi mở file Excel | 🛡️ Phòng ngừa lây lan chéo |
| **Giao diện Người dùng** | Terminal CLI PowerShell | **Menu Add-in Ribbon** trực tiếp trong Excel + Hộp thoại GUI chọn thư mục | ✅ Trải nghiệm người dùng cao |
| **Bảo vệ Toàn vẹn Dữ liệu** | Lưu đè trực tiếp sau khi xóa mã độc (`$wb.Save()`) | **Tự động sao lưu Snapshot** vào `_Backup_Kangatang\` trước khi can thiệp | 🛡️ Zero-loss / Khả năng Rollback |
| **Phạm vi Quét Mã độc** | Module VBA + Sheet ẩn | **Module VBA + Sheet ẩn + Hidden Named Ranges (Name rác do virus tạo)** | 🛡️ Diệt tận gốc không sót rác |
| **Quyền AccessVBOM** | Bắt buộc người dùng tự cấu hình Trust Center thủ công | **Tự động cấu hình Registry** (có cơ chế hoàn nguyên an toàn khi kết thúc) | ✅ Tự động hóa 100% |
| **Khả năng Truy vết (Audit)** | Chỉ in thông tin ra màn hình Console tạm thời | **Xuất file Log kiểm toán có dấu thời gian** lưu trữ dài hạn tại `%APPDATA%` | 🛡️ Observability-ready |

---

## 3. RÀ SOÁT RỦI RO KIẾN TRÚC & GIẢI PHÁP PHÒNG THỦ (VACCINE PATTERNS)

Qua quá trình rà soát mã nguồn và kiểm tra luồng thực thi, chúng tôi xác định **5 rủi ro trọng yếu** và thiết kế các mẫu phòng thủ tương ứng:

### 🛑 Rủi ro 1: Race Condition khi mở file đồng thời với khởi động Excel
- **Nguyên nhân:** Khi người dùng double-click một file `.xls` ngoài Desktop khi Excel chưa chạy, Excel sẽ nạp Add-in trong `XLSTART` đồng thời mở file đó. Nếu việc đăng ký sự kiện `Application_WorkbookOpen` diễn ra sau khi workbook đó đã mở xong, file sẽ bị lọt lưới quét.
- 🛡️ **VACCINE 1 (Defensive Initialization):** Trong thủ tục `InitializeGuard()`, ngay sau khi thiết lập hook sự kiện, bổ sung vòng lặp quét toàn bộ danh sách `Application.Workbooks` hiện hữu để bắt sạch các file đã mở trước đó.

### 🛑 Rủi ro 2: Nguy cơ lây nhiễm ngược vào Add-in (Add-in Tampering)
- **Nguyên nhân:** Virus Kangatang có cơ chế tự nhân bản vào các file workbook nằm trong thư mục `XLSTART`. Nếu file `KangatangGuard.xlam` không được khóa, virus có thể ghi đè hoặc chèn mã độc vào chính Add-in.
- 🛡️ **VACCINE 2 (Immutability & Whitelist Protection):**
  1. Trình cài đặt tự động gắn cờ **Read-Only** ở cấp độ tệp hệ điều hành (`attrib +R`) cho file `KangatangGuard.xlam`.
  2. Thuộc tính `ThisWorkbook.IsAddin = True` được bật để ngăn sửa đổi giao diện.
  3. Cơ chế Whitelist tường minh bỏ qua chính Add-in và file `PERSONAL.XLSB` hợp lệ.

### 🛑 Rủi ro 3: Xung đột Compile Error do trùng lặp `Option Explicit` khi Inject code
- **Nguyên nhân:** Khi khởi tạo Workbook rỗng bằng COM Automation, một số phiên bản Excel tự động chèn sẵn chỉ thị `Option Explicit`. Nếu script PowerShell inject thêm mã nguồn chứa `Option Explicit`, VBA Compiler sẽ báo lỗi cú pháp.
- 🛡️ **VACCINE 3 (Clean Component Preparation):** Xóa sạch toàn bộ dòng mã hiện có trong `CodeModule` bằng lệnh `CodeModule.DeleteLines(1, CodeModule.CountOfLines)` trước khi gọi `AddFromString()`.

### 🛑 Rủi ro 4: File bị khóa quyền ghi (Read-Only) hoặc Chế độ Protected View
- **Nguyên nhân:** Các file đính kèm từ Email hoặc mạng nội bộ thường mở dưới dạng Read-Only hoặc Protected View. Lệnh `wb.Save()` sẽ gây ra Exception gián đoạn luồng làm việc.
- 🛡️ **VACCINE 4 (Defensive Save Strategy):** Thêm bẫy lỗi cô lập quanh khối lưu file. Nếu file ở trạng thái Read-Only, Add-in sẽ thông báo yêu cầu người dùng dùng tính năng "Save As" sang vị trí an toàn thay vì nuốt lỗi (fail silently).

### 🛑 Rủi ro 5: Thư mục XLSTART bị phân mảnh giữa nhiều phiên bản Office
- **Nguyên nhân:** Trên máy có thể tồn tại song song nhiều đường dẫn `XLSTART` (Office 32-bit trong `Program Files (x86)`, Office 64-bit trong `Program Files`, và User Profile cá nhân).
- 🛡️ **VACCINE 5 (Universal Discovery Engine):** Tận dụng engine tìm kiếm động `Get-AllXLStartPaths` đã chứng minh hiệu quả từ v3.0.0 để rà soát tất cả các thư mục khởi động, đảm bảo tiêu diệt mầm bệnh `mypersonnel1.xls` ở mọi ngóc ngách.

---

## 4. CHI TIẾT CÁC THÀNH PHẦN THỰC THI (DELIVERABLES)

```
d:\7. AI tools\kangatang\
├── Diet_Virus_Kangatang.bat          # Script Polyglot v3.1.0 (Quét sâu toàn máy, GUI, Backup, Log)
├── addin\                            # Bộ giải pháp Excel Add-in thời gian thực
│   ├── KangatangGuard_Code.vba       # Mã nguồn chuẩn VBA (ThisWorkbook, clsAppEvents, Scanner, Logger)
│   ├── Install_KangatangGuard.ps1    # Engine tạo .xlam tự động bằng COM & phân phối vào XLSTART
│   ├── Install_Addin.bat             # Trình cài đặt 1-Click tự nâng quyền Administrator
│   └── Uninstall_Addin.bat           # Trình gỡ bỏ 1-Click an toàn, bảo lưu dữ liệu log
└── doc\
    └── 20260903_kangatang_addin_plan.md # Tài liệu kiến trúc & rà soát kỹ thuật
```

---

## 5. QUY TRÌNH KIỂM CHỨNG & NGHIỆM THU (VERIFICATION CHECKLIST)

- [x] **Kiểm tra cú pháp PowerShell:** Xác thực bằng `[System.Management.Automation.Language.Parser]` đạt kết quả `SYNTAX_OK_NO_ERRORS`.
- [x] **Chuẩn hóa mã hóa ký tự:** 100% file `.bat`, `.ps1`, `.vba` được lưu dưới chuẩn **UTF-8 with BOM**, bảo đảm hiển thị tiếng Việt hoàn hảo trên Windows Terminal và VBA Editor.
- [x] **Bảo toàn dữ liệu & Đồng bộ:** Toàn bộ sản phẩm được đồng bộ hai chiều giữa thư mục phát triển `d:\7. AI tools\kangatang` và thư mục làm việc của người dùng `C:\Users\mrKienIT\Desktop\python\coding\AI tools\kangatang`.
- [x] **Khả năng Gỡ cài đặt (Reversible):** Cung cấp sẵn `Uninstall_Addin.bat` để hoàn nguyên hệ thống về trạng thái ban đầu mà không để lại tác dụng phụ.
