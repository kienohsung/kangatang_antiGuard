# 🏛️ BÀI HỌC KIẾN TRÚC & TIẾN HÓA HỆ THỐNG QUA CÁC PHIÊN BẢN (`03_version_lessons_and_architecture.md`)
> Quay lại chỉ mục chính: [`../LEARN.md`](file:///d:/7.%20AI%20tools/kangatang/LEARN.md) | Tọa độ code: [`../AGENT_CONTEXT.md`](file:///d:/7.%20AI%20tools/kangatang/AGENT_CONTEXT.md)

---

## 1. 6 BÀI HỌC KIẾN TRÚC CỐT LÕI (`ARC-01` ➔ `ARC-06`)

### `ARC-01`: Tương thích với EDR / Antivirus Doanh nghiệp (`Zero-IOC & Zero-P/Invoke PS1 in Trusted AddIns`)
- **Bài học:** Trong môi trường doanh nghiệp có cài đặt EDR/Antivirus (như **Seqrite Endpoint Security / Quick Heal**), tuyệt đối không lưu `.ps1` trong `%APPDATA%\KangatangGuard` và không nhúng C# Win32 P/Invoke (`DllImport`) hay chuỗi IOC mã độc dạng văn bản thuần vào file `.ps1`.
- **Chuẩn thiết kế (`v3.17.3+`):** Toàn bộ logic nhận diện và bóc tách mã độc nằm gọn **100% bên trong mã VBA đã biên dịch của `.xlam` (`ScanSingleFileForWorker`)**. Cửa sổ PowerShell ngoài tiến trình (`%APPDATA%\Microsoft\AddIns\Kangatang_FolderScanner.ps1`) chỉ đóng vai trò giao diện hiển thị luồng tiến độ trực quan (Streaming UI) và điều phối COM Worker, đảm bảo **0% P/Invoke**, **0% chuỗi IOC** và **Zero-Hang** cho Excel chính.

### `ARC-02`: Phân tách Hạ tầng LAN: Máy chủ Phân phối (`223.7`) vs Kho Cách ly & Mẫu Mã độc (`223.176`)
- **Bài học:** Không bao giờ dùng chung một thư mục chia sẻ mạng cho cả việc phân phối phần mềm sạch lẫn việc lưu trữ file nhiễm virus (Quarantine Backup / Threat Samples).
- **Chuẩn thiết kế (`v3.13.0` – `v3.14.0`):**
  - **`\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang`:** Chỉ phục vụ phân phối gói cài đặt và tự động cập nhật (`KangatangGuard.xlam`, `version.json`, `Install_Client.ps1`).
  - **`192.168.223.176` (`D:\7. AI tools\kangatang\` / `\\192.168.223.176\test\`):** Chuyên trách thu thập tập trung `Quarantine_Backup\` và `Threat_Samples\`. Có chốt chặn cứng trong code (`FORBIDDEN_SERVER = "192.168.223.7"`) để không bao giờ ghi nhầm mẫu virus lên `223.7`.

### `ARC-03`: Phòng vệ Đa tầng và An toàn Call Stack trong Excel VBA (`Defensive VBE Engineering`)
- **Bài học:** Engine VBE của Excel rất nhạy cảm khi thao tác tự động trên chính các thành phần VBA/Sheet của workbook đang mở:
  1. Không bao giờ chạm vào `wb.VBProject` khi workbook còn chứa `Legacy Module Sheet` (`xlModule = -4141`).
  2. Không bao giờ xóa `Module Sheet` trực tiếp bên trong sự kiện `WorkbookOpen` hoặc `BeforeSave` (phải dùng `Application.OnTime` Two-Phase Purge hoặc vòng lặp 2 bước `WorkerDeleteOneModuleSheet` + `WorkerFinishCleanAndClose`).
  3. Không bao giờ đặt `Application.AutomationSecurity = 3` từ bên trong một hàm VBA đang chạy rồi đọc `VBProject` của workbook `.xlsm` vừa mở (dùng `Application.EnableEvents = False` thay thế).

### `ARC-04`: Tôn trọng Môi trường Của Người Dùng (`Non-Intrusive Coexistence`)
- **Bài học:** Một Add-in bảo mật tốt phải hoạt động "tàng hình", không làm gián đoạn luồng làm việc quen thuộc của người dùng:
  1. **Bảo toàn Clipboard (`CutCopyMode <> 0`):** Không chạy các hàm khôi phục phím tắt/UI khi người dùng đang trong trạng thái Copy/Cut.
  2. **Chung sống với Add-in khác (`MyExcel.xlam`):** Không `.Reset` thanh menu chính (`Worksheet Menu Bar`, `Standard`) và không ghi đè các khóa `OPEN*` của Add-in khác trong Registry.

### `ARC-05`: Chuẩn hóa Vòng đời Add-in theo Chuẩn Microsoft (`Canonical AddIns Path`)
- **Bài học:** Việc cài đặt `.xlam` đồng thời ở cả `XLSTART` và `AddIns` hoặc dùng mã VBA can thiệp vào trạng thái `oAi.Installed` sẽ gây xung đột với hộp thoại `Excel Options -> Add-ins -> Go`.
- **Chuẩn thiết kế (`v3.16.0+`):** Chỉ lưu một bản `.xlam` duy nhất tại `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam` và nạp qua khóa Registry `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN*`.

### `ARC-06`: Một Nguồn Sự Thật Duy Nhất Cho Số Phiên Bản (`Single Source of Truth Version Gate`)
- **Bài học:** Khi dự án có nhiều tệp (`.vba`, `.ps1`, `.bat`, `.json`, `.txt`), việc sửa số phiên bản thủ công chắc chắn sẽ dẫn đến lệch phiên bản giữa các thành phần.
- **Chuẩn thiết kế (`v3.17.1+`):** Mọi đợt phát hành đều đi qua `addin/Deploy_To_Hub.ps1` (Bước `[0/6]`), tự động đồng bộ biến `$ReleaseVersion` sang toàn bộ 15 tệp trước khi biên dịch `.xlam`, kết hợp đọc phiên bản động từ `version.json` / Registry.

---

## 2. BẢNG TIẾN HÓA KIẾN TRÚC QUA CÁC PHIÊN BẢN (`v3.1` ➔ `v3.20.2`)

| Phiên bản | Ngày | Cột mốc Kiến trúc & Cải tiến Trọng tâm | Tài liệu chi tiết trong `doc/` |
| :--- | :--- | :--- | :--- |
| **`v3.1` – `v3.3`** | `09/2026` | Khởi tạo Add-in thời gian thực, hỗ trợ Office 365 không có `XLSTART`, xử lý đường dẫn tiếng Việt Unicode (`Uni` + `MessageBoxW`). | `20260903_kangatang_addin_plan.md`<br>`20260908_fix_folder_scan_unicode_v3.3.0.md` |
| **`v3.4` – `v3.6`** | `09/2026` | Tách luồng quét thư mục, bổ sung cơ chế `Fast Resume` (`last_session.json` & `scanned_files.log`) và `ScanCache` chống lag khi Save. | `20260909_stream_scan_architecture_v3.5.0.md`<br>`20260909_resume_scan_session_architecture_v3.6.0.md` |
| **`v3.7` – `v3.8.8`** | `09/2026` | Thiết lập Trung tâm Phân phối LAN (`223.7`), tự động cập nhật `PerformLanUpdate`, tích hợp Tab Ribbon riêng biệt với Icon HD 32x32 (`customUI14.xml`). | `20260909_dedicated_lan_server_dual_mirror_architecture_v3.8.0.md`<br>`20260917_deduplicate_ui_dedicated_ribbon_tab_v3.8.8.md` |
| **`v3.10.0`** | `10/2026` | **Runtime Shield:** Quét dọn ổ dịch `XLSTART`/`AddIns`, Ping-Guard `IsUncReachable`, và thuật toán **Two-Phase Purge** xử lý `Module Sheet` Excel 4.0/5.0 không gây crash. | `20261006_runtime_shield_v3.10.0.md` |
| **`v3.11` – `v3.12.2`** | `10/2026` | **Surgical Clean & Clipboard Routing:** Chỉ xóa hàm nhiễm IOC giữ lại hàm sạch, mở khóa sheet mật khẩu rỗng, định tuyến `Ctrl+C`/`Ctrl+V` qua `Two-Step PasteSpecial`. | `20261007_threat_collector_and_copypaste_fix_v3.11.0.md` |
| **`v3.13` – `v3.14.0`** | `10/2026` | **Cross-Workbook Copy/Paste & Centralized Vault `223.176`:** Bảo toàn `CutCopyMode` khi chuyển giữa 2 file Excel; chuyển toàn bộ `Quarantine_Backup` và `Threat_Samples` về máy `223.176` (cấm lưu trên `223.7`). | `20261008_cross_workbook_copypaste_and_centralized_threat_collection_v3.13.0.md`<br>`20261008_redirect_backup_to_223_176_and_purge_223_7_v3.14.0.md` |
| **`v3.15.0`** | `10/2026` | **VBOM Self-Healing & Third-Party Coexistence:** Tự động sửa Registry `AccessVBOM = 1` không hiện popup; bảo vệ menu và khóa `OPEN*` của `MyExcel.xlam`. | `20261014_vbom_self_healing_zero_popup_v3.15.0.md` |
| **`v3.16.0`** | `10/2026` | **Instant Add-In Manager Integration (`< 5ms` Zero-Spin):** Khắc phục triệt để lỗi quay tròn và không tích chọn được Add-in trong `Excel Options -> Add-ins -> Go`. | `20261017_fix_addin_manager_install_v3.16.0.md` |
| **`v3.17.0` – `v3.17.3`** | `10/2026` | **Out-of-Process Streaming Console + Native VBA Worker (`Two-Step Smart Open`):** Khởi chạy cửa sổ PowerShell độc lập hiển thị tiến độ trực tiếp từng tệp (`ScanSingleFileForWorker`), 0% IOC/P-Invoke chống Seqrite/Quick Heal cách ly, đồng bộ tự động số phiên bản trên 15 file và cập nhật thương hiệu `OhsungIT team`. | `20261017_fix_antivirus_false_positive_v3.17.0.md` |
| **`v3.18.0`** | `10/2026` | **4-Layer Turbo Folder Scanner (Tăng tốc 50x–100x):** Duyệt cuốn chiếu (`Streaming Queue`), Bộ lọc tĩnh `Zero-COM OpenXML Pre-Filter` (`1.5 ms/tệp`, không cần mở Excel với 85% tệp `.xlsx` sạch), kết hợp `Lazy On-Demand Multi-Worker COM Pool` (`2–4` tiến trình Excel song song) xử lý `.xls`/`.xlsm`, rút ngắn thời gian quét 20.000 tệp từ `> 24 giờ` xuống `~10–20 phút`. | `20261017_turbo_folder_scanner_architecture_v3.18.0.md` |
| **`v3.19.0` – `v3.19.3`** | `10/2026` | **Non-Scrolling Live Dashboard UI + Immortal Worker Pool + Single-Instance Mutex:** Giao diện bảng điều khiển cố định tự co giãn (`Adaptive Console Width Guard`), hàm băm `Pure Native VBA ComputeThreatHash` (0% `.NET CLR mscorwks.dll` crash), `Immortal Worker Pool` + `LOH 4MB Cap` chịu tải `66.000+` tệp, và `Single-Instance Mutex` chặn bật trùng cửa sổ PowerShell. | `20261017_live_dashboard_progress_ui_v3.19.0.md` |
| **`v3.20.0`** | `10/2026` | **Enterprise Periodic Scan Reporting & Project Audit Logging (`Scan_Reports\`):** Tự động ghi nhật ký kiểm toán chi tiết và xuất bộ 4 báo cáo định kỳ (`LATEST_SCAN_REPORT.md`, `Scan_Summary_History.csv`, `Infected_Files_Ledger_YYYYMM.csv`, `scan_detail_log_YYYYMMDD.txt`) kèm thống kê theo từng Phòng ban / Thư mục cấp 1 trực tiếp vào `D:\7. AI tools\kangatang\Scan_Reports\`. | `20261017_detailed_scan_report_logging_plan.md` |
| **`v3.20.1` – `v3.20.2`** | `10/2026` | **RAM-Only DDE Isolation, Zero-Freeze Worker Calculation (`< 1.3s` User Excel Open) & UTF-8 Resume Session Reader:** Cách ly hoàn toàn Worker ngầm với Windows Explorer DDE (`IgnoreRemoteRequests = True` trong RAM + xóa ngay bit `0x40` trong Registry `HKCU\...\Excel\Options\Options`), chuyển `Application.Calculation = xlCalculationManual` (`-4135`) + `DoEvents` khi xóa hàng loạt sheet ẩn `Kangatang_N`, và bổ sung `ReadUtf8FileText` (`ADODB.Stream` UTF-8 + Fallback) cho `IsScannerCurrentlyRunning` & `ResumeScanDialog`. | `AGENT_CONTEXT.md` |

