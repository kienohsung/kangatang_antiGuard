# ⚡ AGENT QUICK-START CONTEXT & ARCHITECTURE MAP (`AGENT_CONTEXT.md`)
> **MỤC ĐÍCH:** Đọc tệp này ĐẦU TIÊN trong mỗi phiên làm việc để nắm toàn bộ kiến trúc, quy tắc bất biến và tọa độ dòng lệnh (`StartLine`–`EndLine`) mà **không cần đọc toàn bộ file `KangatangGuard_Code.vba` (4.059 dòng)** — tiết kiệm **85–90% token** và tăng tốc xử lý.
> 🧠 **TRA CỨU BUGS, THUẬT TOÁN & BÀI HỌC CÁC PHIÊN BẢN:** Xem ngay **[`LEARN.md`](file:///d:/7.%20AI%20tools/kangatang/LEARN.md)** (`learn/01_bugs_and_root_causes.md`, `learn/02_algorithms_and_optimizations.md`, `learn/03_version_lessons_and_architecture.md`).

---

## 1. THÔNG TIN PHIÊN BẢN & HẠ TẦNG MẠNG LAN (CRITICAL TOPOLOGY)

- **Phiên bản hiện tại:** `v3.21.0` (`Executive Multi-Sheet Excel Report [1-Page A4 Dashboard] + Endpoint MAC Telemetry + Last Saved By Forensics`)
- **Phân tách vai trò 2 máy chủ LAN (BẤT BIẾN):**
  | Máy chủ | Đường dẫn UNC / Local | Vai trò duy nhất | Quy tắc CẤM (Hard Constraint) |
  | :--- | :--- | :--- | :--- |
  | **`192.168.223.7`** *(Distribution Hub)* | `\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang` | Phân phối gói cài đặt & tự động cập nhật (`KangatangGuard.xlam`, `version.json`, `Install_Client.ps1`, `Kangatang_FolderScanner.ps1`) | 🛑 **TUYỆT ĐỐI KHÔNG** lưu `Quarantine_Backup`, `Threat_Samples` hoặc `Scan_Reports` trên `223.7`! |
  | **`192.168.223.176`** *(Master & Threat Vault)* | Local: `D:\7. AI tools\kangatang\`<br>LAN Share: `\\192.168.223.176\test\Quarantine_Backup`, `\Threat_Samples` & `\Scan_Reports` | Lưu trữ tập trung mẫu mã độc (`Threat_Samples/`), bản sao lưu file nhiễm (`Quarantine_Backup/`) và Báo cáo Quản trị Đa Sheet / Nhật ký (`Scan_Reports/`) | Không dùng làm Hub cập nhật chính cho Client (chỉ lưu bản Clone tại `distribution/`) |

- **Vị trí cài đặt chuẩn trên máy Client/Master:**
  - Add-in & Trình quét ngoài tiến trình: `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam` và `%APPDATA%\Microsoft\AddIns\Kangatang_FolderScanner.ps1` (**KHÔNG** chép `.xlam` vào `XLSTART`).
  - Tự động nạp qua Registry: `HKCU\Software\Microsoft\Office\<16.0|15.0|14.0>\Excel\Options\OPEN* = "KangatangGuard.xlam"` (bảo toàn các Add-in khác như `MyExcel.xlam`).
  - Thư mục Log & Session cục bộ: `%APPDATA%\KangatangGuard\` (`scan_log_YYYYMMDD.txt`, `last_session.json`, `scanned_files.log`, `scan_request.json`). **TUYỆT ĐỐI KHÔNG lưu file `.ps1` trong `%APPDATA%\KangatangGuard\`.**
  - Thư mục Báo cáo Quản trị & Log chi tiết trong dự án (`v3.20.0` – `v3.21.0`): `D:\7. AI tools\kangatang\Scan_Reports\` (`Bao_Cao_Quan_Tri_KangatangGuard.xlsx` [6 Sheets, 1-Page A4 Dashboard], `Endpoint_Registry.csv` [Khóa chính `MAC Address`], `LATEST_SCAN_REPORT.md`, `Scan_Summary_History.csv`, `Infected_Files_Ledger_YYYYMM.csv`, `scan_detail_log_YYYYMMDD.txt`).

---

## 2. BẢN ĐỒ TỌA ĐỘ DÒNG CODE (`addin/KangatangGuard_Code.vba` — 4.355 DÒNG)

> 💡 **Mẹo tiết kiệm Token:** Khi dùng `view_file` trên `addin/KangatangGuard_Code.vba`, hãy chỉ định đúng `StartLine` và `EndLine` theo bảng dưới đây:

### A. Cấu trúc 6 Section (Được `Install_KangatangGuard.ps1` tách tự động khi biên dịch `.xlam`)
| Section Name | Loại Module trong `.xlam` | Dòng (`Start`–`End`) | Chức năng chính |
| :--- | :--- | :--- | :--- |
| `ThisWorkbook` | Document Module | `L13` – `L41` | Hook `Workbook_Open` (tự động phân luồng `_Worker_` vs User Excel), `Workbook_AddinInstall`, `Workbook_BeforeClose` |
| `clsAppEvents` | Class Module | `L44` – `L114` | Bắt sự kiện toàn cục `xlApp_WorkbookOpen`, `BeforeSave`, `NewWorkbook`, `SheetActivate`, `WorkbookActivate` |
| `modKangatangScanner` | Standard Module | `L117` – `L2784` | Engine quét/diệt virus, Out-of-Process Launcher (`IsScannerCurrentlyRunning`, `ReadUtf8FileText`), Worker Macros (`InitWorkerMode`, `ClearExcelIgnoreDdeRegistryBit`, `ScanSingleFileForWorker`), Truy vết `GetWorkbookLastAuthor`, Đồng bộ `RegisterEndpointHeartbeat` (`MAC/IP`), Copy/Paste Shortcuts, Ribbon Callbacks, LAN Auto-Update |
| `modKangatangShield` | Standard Module | `L2787` – `L4083` | Nhận diện chữ ký/IOC virus, Startup Sweep, Quarantine, CleanAddInCollisions, Two-Phase Module-Sheet Purge |
| `modLogger` | Standard Module | `L4086` – `L4137` | Ghi nhật ký kiểm toán `WriteLog` ra `%APPDATA%\KangatangGuard\scan_log_YYYYMMDD.txt` và đồng bộ về `D:\7. AI tools\kangatang\Scan_Reports\scan_detail_log_YYYYMMDD.txt` |
| `modThreatCollector` | Standard Module | `L4140` – `L4355` | Băm Pure Native VBA Dual-Seed 32-bit (`ComputeThreatHash`) & thu thập mẫu mã độc JSON (`CollectThreatSample`) về máy `223.176` |

### B. Tọa độ các hàm trọng yếu trong `modKangatangScanner` (`L117` – `L2784`)
- **Hằng số phiên bản & cấu hình (`CURRENT_VERSION = "3.21.0"`, `CENTRAL_BACKUP_HUB`):** `L120` – `L160`
- **Unicode & Headless-safe Dialog (`Uni`, `MsgBoxW`):** `L161` – `L208`
- **ScanCache chống lag khi Save (`IsScanCacheExpired`, `UpdateScanCache`):** `L209` – `L245`
- **Khởi tạo & Ribbon Callbacks (`InitializeGuard`, `EnsureGuardAlive`, `Ribbon_*`):** `L246` – `L366`
- **Kiểm tra nhiễm virus (`IsWhitelisted`, `CheckWorkbookInfection`):** `L367` – `L513`
- **Truy vết Tác giả Lưu cuối & Quét Workbook (`GetWorkbookLastAuthor`, `ScanWorkbook`):** `L518` – `L645`
- **Khởi chạy Trình Quét Ngoài Tiến Trình (`LaunchBackgroundScanner`, `WriteScanRequestJson`, `EnsureFolderScannerScript`):** `L650` – `L935`
- **Worker Macros cho PowerShell Turbo (`InitWorkerMode`, `ClearExcelIgnoreDdeRegistryBit`, `GetPreFilterRegex`, `ScanSingleFileForWorker`, `WorkerDeleteOneModuleSheet`, `WorkerFinishCleanAndClose`):** `L940` – `L1130`
- **Trình Quét Thư Mục Native VBA Dự Phòng (`RunNativeFolderScan`, `CollectExcelFilesRecursive`, `WriteScanSessionJson`):** `L1135` – `L1490`
- **Sao lưu cách ly tập trung về `223.176` (`BackupBeforeClean`):** `L1495` – `L1580`
- **Định tuyến Copy/Paste/Cut & Khôi phục UI (`Kangatang_ShortcutCopy/Paste/Cut`, `EnsureClipboardShortcutsAndMenus`, `RestoreExcelClipboardAndUI`):** `L1582` – `L1775`
- **Làm sạch Workbook nhiễm (`TryUnprotectWorkbook/Sheet`, `CleanInfectedWorkbook`):** `L1780` – `L2035`
- **Hộp thoại Quét Thư Mục & Fast Resume (`ScanFolderDialog`, `ResumeScanDialog`):** `L2000` – `L2120`
- **Cập nhật tự động qua LAN & Heartbeat MAC/IP (`CheckForLanUpdates`, `RegisterEndpointHeartbeat`, `PerformLanUpdate`):** `L2200` – `L2615`

---

## 3. 7 QUY TẮC KỸ THUẬT BẤT BIẾN (GOLDEN INVARIANTS — DO NOT BREAK)

1. 🛡️ **Bảo vệ bộ nhớ đệm Copy/Paste (Cùng file & Khác file):**
   - Trong `clsAppEvents` (`SheetActivate`, `WorkbookActivate`, `WorkbookOpen`), `EnsureClipboardShortcutsAndMenus` và `RestoreExcelClipboardAndUI`: **BẮT BUỘC** kiểm tra `If Application.CutCopyMode <> 0 Then Exit Sub` ở đầu hàm.
   - Trong `Kangatang_ShortcutPaste`: Khi `Application.CutCopyMode = 1` (`xlCopy`), dùng quy trình dán 2 bước `Selection.PasteSpecial -4123` (`xlPasteFormulas`) + `Selection.PasteSpecial -4122` (`xlPasteFormats`) (fallback `-4163` `xlPasteValues`).
2. 🛡️ **Chung sống hòa bình với Add-in bên thứ 3 (`MyExcel.xlam`):**
   - **TUYỆT ĐỐI KHÔNG** gọi `.Reset` trên `CommandBars("Worksheet Menu Bar")` hoặc `CommandBars("Standard")`. Chỉ được phép `.Reset` trên 4 menu ngữ cảnh chuột phải: `"Cell"`, `"Row"`, `"Column"`, `"XLM Cell"`.
   - Khi đăng ký `OPEN*` trong Registry (`Install_KangatangGuard.ps1`, `Install_Client.ps1`, `CleanAddInCollisions`), chỉ thêm `"KangatangGuard.xlam"` vào slot `OPEN*` trống, giữ nguyên các khóa `OPEN*` của Add-in khác.
3. 🛡️ **Chống treo/quay tròn hộp thoại Add-Ins (`File -> Options -> Add-ins -> Go`):**
   - `InitializeGuard()` phải hoàn tất trong `< 5ms` và hoãn quét khởi động qua `Application.OnTime` (`ShieldStartupSweep`).
   - Trong `CleanAddInCollisions()`, **TUYỆT ĐỐI KHÔNG** gán `oAi.Installed = False` đối với `KangatangGuard.xlam`.
4. 🛡️ **Chống Antivirus (Seqrite / Quick Heal) nhận nhầm (`Zero-PS1 in %APPDATA%\KangatangGuard`):**
   - Không bao giờ lưu, sao chép hoặc chạy tệp `.ps1` trong `%APPDATA%\KangatangGuard` (chỉ đặt bản wrapper sạch tại `%APPDATA%\Microsoft\AddIns\Kangatang_FolderScanner.ps1`).
   - Không bao giờ đưa mã C# `Add-Type` P/Invoke (`DllImport("kernel32.dll")`, `ole32.dll`, `user32.dll`, `Process.Kill()`) hoặc chuỗi tên virus vào các tệp `.ps1` trong `addin/` và `distribution/`.
5. 🛡️ **Quy tắc truy cập `VBProject` & Two-Step Smart Open:**
   - Khi mở tệp từ bên trong VBA (`ScanSingleFileForWorker` / `RunNativeFolderScan`), **KHÔNG** đặt `Application.AutomationSecurity = 3` vì VBE sẽ ném lỗi `0x800A9C68` (`40040`) khi truy cập `wbTarget.VBProject` của tệp `.xlsm`. Chỉ cần đặt `Application.EnableEvents = False`.
   - Luôn mở tệp ở chế độ `ReadOnly:=True` trước để kiểm tra nhanh (`CheckWorkbookInfection`), chỉ mở lại với `ReadOnly:=False` khi phát hiện tệp thực sự nhiễm virus (`Two-Step Smart Open`).
6. 🛡️ **Chống crash Excel khi gặp Legacy Excel 4.0/5.0 Module Sheet:**
   - Nếu `LegacyModuleSheetCount(wb) > 0`, **TUYỆT ĐỐI KHÔNG** truy cập `wb.VBProject` ngay trong luồng sự kiện. Bắt buộc chuyển giao cho quy trình Two-Phase Purge (`DeferModuleSheetPurge` hoặc `WorkerDeleteOneModuleSheet` + `WorkerFinishCleanAndClose`).
7. 🛡️ **Chuẩn hóa tệp `.bat` / `.cmd` & Console Output:**
   - 100% Pure ASCII (tiếng Anh hoặc tiếng Việt không dấu hoàn toàn trong cả `REM`, `echo`, `title` và `Write-Host`).

---

## 4. QUY TRÌNH BIÊN DỊCH, TRIỂN KHAI & KIỂM THỬ 1 LỆNH (CHEAT SHEET)

- **Biên dịch `.vba` ➔ `.xlam` + Đồng bộ `distribution/` + Phát hành lên LAN Hub `223.7`:**
  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File "d:\7. AI tools\kangatang\addin\Deploy_To_Hub.ps1"
  ```
- **Khóa file khi phối hợp đa Agent:** Kiểm tra và cập nhật [`doc/_agent_file_locks.md`](file:///d:/7.%20AI%20tools/kangatang/doc/_agent_file_locks.md).
