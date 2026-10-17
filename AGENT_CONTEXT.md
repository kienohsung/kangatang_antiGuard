# ⚡ AGENT QUICK-START CONTEXT & ARCHITECTURE MAP (`AGENT_CONTEXT.md`)
> **MỤC ĐÍCH:** Đọc tệp này ĐẦU TIÊN trong mỗi phiên làm việc để nắm toàn bộ kiến trúc, quy tắc bất biến và tọa độ dòng lệnh (`StartLine`–`EndLine`) mà **không cần đọc toàn bộ file `KangatangGuard_Code.vba` (3.651 dòng)** — tiết kiệm **85–90% token** và tăng tốc xử lý.

---

## 1. THÔNG TIN PHIÊN BẢN & HẠ TẦNG MẠNG LAN (CRITICAL TOPOLOGY)

- **Phiên bản hiện tại:** `v3.17.0` (`100% Native VBA Folder Scanner & Zero-PS1 Antivirus-Safe Architecture`)
- **Phân tách vai trò 2 máy chủ LAN (BẤT BIẾN):**
  | Máy chủ | Đường dẫn UNC / Local | Vai trò duy nhất | Quy tắc CẤM (Hard Constraint) |
  | :--- | :--- | :--- | :--- |
  | **`192.168.223.7`** *(Distribution Hub)* | `\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang` | Phân phối gói cài đặt & tự động cập nhật (`KangatangGuard.xlam`, `version.json`, `Install_Client.ps1`) | 🛑 **TUYỆT ĐỐI KHÔNG** lưu `Quarantine_Backup` hoặc `Threat_Samples` trên `223.7`! |
  | **`192.168.223.176`** *(Master & Threat Vault)* | Local: `D:\7. AI tools\kangatang\`<br>LAN Share: `\\192.168.223.176\test\Quarantine_Backup` & `\Threat_Samples` | Lưu trữ tập trung mẫu mã độc (`Threat_Samples/*.sample.json`) và bản sao lưu file nhiễm (`Quarantine_Backup/`) từ mọi máy LAN | Không dùng làm Hub cập nhật chính cho Client (chỉ lưu bản Clone tại `distribution/`) |

- **Vị trí cài đặt chuẩn trên máy Client/Master:**
  - Add-in duy nhất: `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam` (**KHÔNG** chép `.xlam` vào `XLSTART`).
  - Tự động nạp qua Registry: `HKCU\Software\Microsoft\Office\<16.0|15.0|14.0>\Excel\Options\OPEN* = "KangatangGuard.xlam"` (bảo toàn các Add-in khác như `MyExcel.xlam`).
  - Thư mục Log & Session: `%APPDATA%\KangatangGuard\` (`scan_log_YYYYMMDD.txt`, `last_session.json`, `scanned_files.log`). **TUYỆT ĐỐI KHÔNG lưu file `.ps1` trong `%APPDATA%\KangatangGuard\`.**

---

## 2. BẢN ĐỒ TỌA ĐỘ DÒNG CODE (`addin/KangatangGuard_Code.vba` — 3.651 DÒNG)

> 💡 **Mẹo tiết kiệm Token:** Khi dùng `view_file` trên `addin/KangatangGuard_Code.vba`, hãy chỉ định đúng `StartLine` và `EndLine` theo bảng dưới đây:

### A. Cấu trúc 6 Section (Được `Install_KangatangGuard.ps1` tách tự động khi biên dịch `.xlam`)
| Section Name | Loại Module trong `.xlam` | Dòng (`Start`–`End`) | Chức năng chính |
| :--- | :--- | :--- | :--- |
| `ThisWorkbook` | Document Module | `L13` – `L36` | Hook `Workbook_Open`, `Workbook_AddinInstall`, `Workbook_BeforeClose` |
| `clsAppEvents` | Class Module | `L39` – `L109` | Bắt sự kiện toàn cục `xlApp_WorkbookOpen`, `BeforeSave`, `NewWorkbook`, `SheetActivate`, `WorkbookActivate` |
| `modKangatangScanner` | Standard Module | `L112` – `L2093` | Engine quét/diệt virus, Native Folder Scanner, Copy/Paste Shortcuts, Ribbon Callbacks, LAN Auto-Update |
| `modKangatangShield` | Standard Module | `L2096` – `L3378` | Nhận diện chữ ký/IOC virus, Startup Sweep, Quarantine, CleanAddInCollisions, Two-Phase Module-Sheet Purge |
| `modLogger` | Standard Module | `L3381` – `L3413` | Ghi nhật ký kiểm toán `WriteLog` ra `%APPDATA%\KangatangGuard\scan_log_YYYYMMDD.txt` |
| `modThreatCollector` | Standard Module | `L3416` – `L3651` | Băm MD5 (`ComputeThreatHash`) & thu thập mẫu mã độc JSON (`CollectThreatSample`) về máy `223.176` |

### B. Tọa độ các hàm trọng yếu trong `modKangatangScanner` (`L112` – `L2093`)
- **Hằng số phiên bản & cấu hình (`CURRENT_VERSION`, `CENTRAL_BACKUP_HUB`):** `L115` – `L155`
- **Unicode & Headless-safe Dialog (`Uni`, `MsgBoxW`):** `L156` – `L203`
- **ScanCache chống lag khi Save (`IsScanCacheExpired`, `UpdateScanCache`):** `L204` – `L240`
- **Khởi tạo & Ribbon Callbacks (`InitializeGuard`, `EnsureGuardAlive`, `Ribbon_*`):** `L241` – `L356`
- **Kiểm tra nhiễm virus (`IsWhitelisted`, `CheckWorkbookInfection`):** `L357` – `L503`
- **Quét Workbook đơn lẻ (`ScanWorkbook`):** `L509` – `L625`
- **Trình Quét Thư Mục 100% Native VBA (`RunNativeFolderScan`, `CollectExcelFilesRecursive`, `WriteScanSessionJson`):** `L632` – `L976`
- **Sao lưu cách ly tập trung về `223.176` (`BackupBeforeClean`):** `L997` – `L1082`
- **Định tuyến Copy/Paste/Cut & Khôi phục UI (`Kangatang_ShortcutCopy/Paste/Cut`, `EnsureClipboardShortcutsAndMenus`, `RestoreExcelClipboardAndUI`):** `L1084` – `L1276`
- **Làm sạch Workbook nhiễm (`TryUnprotectWorkbook/Sheet`, `CleanInfectedWorkbook`):** `L1282` – `L1532`
- **Hộp thoại Quét Thư Mục & Fast Resume (`ScanFolderDialog`, `ResumeScanDialog`):** `L1554` – `L1664`
- **Cập nhật tự động qua LAN (`CheckForLanUpdates`, `IsNewerVersion`, `PerformLanUpdate`):** `L1771` – `L2093`

### C. Tọa độ các hàm trọng yếu trong `modKangatangShield` (`L2096` – `L3378`)
- **Nhận diện chữ ký & IOC (`IsVirusName`, `IsVirusReference`, `CodeHasIOC`, `LegacyModuleSheetCount`):** `L2130` – `L2224`
- **Gỡ hook virus & Quét khởi động (`NeutralizeVirusHooks`, `ShieldStartupSweep`, `SweepStartupFolders`):** `L2290` – `L2622`
- **Cách ly tệp khởi động độc hại về `223.176` (`GetQuarantineDir`, `QuarantineFile`):** `L2512` – `L2587`
- **Đồng bộ Add-in & Dọn xung đột + Xóa `.ps1` cũ (`CleanAddInCollisions`):** `L2668` – `L2857`
- **Kiểm tra UNC & Tự phục hồi `AccessVBOM` (`IsUncReachable`, `IsVbomAccessible`, `EnsureVbomRegistryEnabled`):** `L2858` – `L3046`
- **Bóc tách hàm nhiễm & Xóa Sheet an toàn (`RemoveInfectedProcedures`, `DeleteSheetSafe`):** `L3047` – `L3175`
- **Two-Phase Purge cho Module-Sheet Excel 5.0 (`DeferModuleSheetPurge`, `ShieldPurgeModuleSheets`, `ShieldAfterPurge`):** `L3190` – `L3378`

---

## 3. 7 QUY TẮC KỸ THUẬT BẤT BIẾN (GOLDEN INVARIANTS — DO NOT BREAK)

1. 🛡️ **Bảo vệ bộ nhớ đệm Copy/Paste (Cùng file & Khác file):**
   - Trong `clsAppEvents` (`SheetActivate`, `WorkbookActivate`, `WorkbookOpen`), `EnsureClipboardShortcutsAndMenus` và `RestoreExcelClipboardAndUI`: **BẮT BUỘC** kiểm tra `If Application.CutCopyMode <> 0 Then Exit Sub` ở đầu hàm. Mọi thao tác chạm vào `Application.OnKey`, `Application.CellDragAndDrop = True`, `Sheet.EnableSelection` hoặc `CommandBars` khi `CutCopyMode <> 0` sẽ làm Excel hủy tức thì vùng nét đứt (marching ants) và xóa clipboard.
   - Trong `Kangatang_ShortcutPaste`: Khi `Application.CutCopyMode = 1` (`xlCopy`), dùng quy trình dán 2 bước `Selection.PasteSpecial -4123` (`xlPasteFormulas`) + `Selection.PasteSpecial -4122` (`xlPasteFormats`) (fallback `-4163` `xlPasteValues`). Không dùng `ActiveSheet.Paste` hoặc `xlPasteAll (-4104)` vì gây lỗi `1004` trên Office 2016/2019/365.
2. 🛡️ **Chung sống hòa bình với Add-in bên thứ 3 (`MyExcel.xlam`):**
   - **TUYỆT ĐỐI KHÔNG** gọi `.Reset` trên `CommandBars("Worksheet Menu Bar")` hoặc `CommandBars("Standard")`. Chỉ được phép `.Reset` trên 4 menu ngữ cảnh chuột phải: `"Cell"`, `"Row"`, `"Column"`, `"XLM Cell"`.
   - Khi đăng ký `OPEN*` trong Registry (`Install_KangatangGuard.ps1`, `Install_Client.ps1`, `CleanAddInCollisions`), chỉ thêm `"KangatangGuard.xlam"` vào slot `OPEN*` trống, giữ nguyên các khóa `OPEN*` của Add-in khác.
3. 🛡️ **Chống treo/quay tròn hộp thoại Add-Ins (`File -> Options -> Add-ins -> Go`):**
   - `InitializeGuard()` phải hoàn tất trong `< 5ms` và hoãn quét khởi động qua `Application.OnTime` (`ShieldStartupSweep`).
   - Trong `CleanAddInCollisions()`, **TUYỆT ĐỐI KHÔNG** gán `oAi.Installed = False` đối với `KangatangGuard.xlam`.
4. 🛡️ **Chống Antivirus (Seqrite / Quick Heal) nhận nhầm (`Zero-PS1 in %APPDATA%`):**
   - Không bao giờ lưu, sao chép hoặc chạy tệp `.ps1` trong `%APPDATA%\KangatangGuard`.
   - Không bao giờ đưa mã C# `Add-Type` P/Invoke (`DllImport("kernel32.dll")`, `ole32.dll`, `user32.dll`, `Process.Kill()`) hoặc chuỗi tên virus vào các tệp `.ps1` trong `addin/` và `distribution/`.
5. 🛡️ **Quy tắc truy cập `VBProject` bên trong `RunNativeFolderScan`:**
   - Khi mở tệp từ bên trong VBA (`RunNativeFolderScan`), **KHÔNG** đặt `Application.AutomationSecurity = 3` vì VBE sẽ ném lỗi `0x800A9C68` (`40040`) khi truy cập `wbTarget.VBProject` của tệp `.xlsm`. Chỉ cần đặt `Application.EnableEvents = False` (vì `Workbooks.Open` gọi từ VBA mặc định không bao giờ tự chạy `Auto_Open`).
6. 🛡️ **Chống crash Excel khi gặp Legacy Excel 4.0/5.0 Module Sheet:**
   - Nếu `LegacyModuleSheetCount(wb) > 0`, **TUYỆT ĐỐI KHÔNG** truy cập `wb.VBProject` ngay trong luồng sự kiện. Bắt buộc chuyển giao cho quy trình Two-Phase Purge (`DeferModuleSheetPurge` qua `Application.OnTime`).
7. 🛡️ **Chuẩn hóa tệp `.bat` / `.cmd`:**
   - 100% Pure ASCII (tiếng Anh hoặc tiếng Việt không dấu hoàn toàn trong cả `REM`, `echo`, `title`).

---

## 4. QUY TRÌNH BIÊN DỊCH, TRIỂN KHAI & KIỂM THỬ 1 LỆNH (CHEAT SHEET)

- **Biên dịch `.vba` ➔ `.xlam` + Đồng bộ `distribution/` + Phát hành lên LAN Hub `223.7`:**
  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File "d:\7. AI tools\kangatang\addin\Deploy_To_Hub.ps1"
  ```
  *(Lưu ý: `Deploy_To_Hub.ps1` đã tự động gọi `addin\Install_KangatangGuard.ps1` ở Bước 1, đóng gói `.xlam` có nhúng Ribbon CustomUI XML, tạo `version.json`, đồng bộ `distribution/` và đẩy lên `\\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang`).*
- **Khóa file khi phối hợp đa Agent:** Kiểm tra và cập nhật [`doc/_agent_file_locks.md`](file:///d:/7.%20AI%20tools/kangatang/doc/_agent_file_locks.md).
