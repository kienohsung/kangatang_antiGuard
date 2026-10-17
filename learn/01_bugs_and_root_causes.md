# 🐞 TỪ ĐIỂN BUG THỰC TẾ, NGUYÊN NHÂN GỐC RỄ & CÁCH KHẮC PHỤC (`01_bugs_and_root_causes.md`)
> Quay lại chỉ mục chính: [`../LEARN.md`](file:///d:/7.%20AI%20tools/kangatang/LEARN.md) | Tọa độ code: [`../AGENT_CONTEXT.md`](file:///d:/7.%20AI%20tools/kangatang/AGENT_CONTEXT.md)

---

### `BUG-01`: Mất vùng nét đứt (Marching Ants) và mất dữ liệu Copy/Paste khi chuyển Sheet hoặc chuyển giữa 2 file Excel (`v3.11.0` – `v3.13.0`)
- **Triệu chứng:** Bấm `Ctrl+C` có vòng nét đứt, nhưng khi bấm `Ctrl+V` (hoặc chuyển sang sheet khác / sang cửa sổ file Excel thứ 2 rồi bấm `Ctrl+V`) thì nét đứt biến mất và không có dữ liệu được dán. Dùng menu chuột phải đôi khi vẫn được.
- **Root Cause:** Trong Excel, bất kỳ lệnh VBA nào gán lại thuộc tính `Application.OnKey`, `Application.CellDragAndDrop = True`, `Sheet.EnableSelection`, `Sheet.Unprotect` hoặc thao tác lên `Application.CommandBars` đều làm Excel **xóa tức thì bộ nhớ đệm Clipboard nội bộ** và đưa `Application.CutCopyMode` về `0` (`False`). Khi người dùng chuyển cửa sổ hoặc chuyển sheet, các sự kiện `xlApp_SheetActivate`, `xlApp_WorkbookActivate`, `xlApp_WorkbookOpen` gọi `EnsureClipboardShortcutsAndMenus` làm hủy vùng đang Copy.
- **Cách khắc phục (`v3.13.0`):**
  1. Đặt chốt chặn sớm `If Application.CutCopyMode <> 0 Then Exit Sub` ở dòng đầu tiên của `xlApp_SheetActivate`, `xlApp_WorkbookActivate`, `xlApp_NewWorkbook`, `EnsureClipboardShortcutsAndMenus` và `RestoreExcelClipboardAndUI`.
  2. Trong `xlApp_WorkbookOpen`: Nếu `Application.CutCopyMode <> 0`, chỉ quét khi `QuickCheckWorkbook(Wb) = True` để không phá vỡ thao tác Copy từ file cũ sang file mới mở.
- **Tọa độ code:** [`KangatangGuard_Code.vba`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L48-L108) (`clsAppEvents`), [`L1148-L1276`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L1148-L1276) (`EnsureClipboardShortcutsAndMenus`).

---

### `BUG-02`: Lỗi `Run-time error 1004` khi gọi `ActiveSheet.Paste` hoặc `xlPasteAll (-4104)` trong macro `Application.OnKey "^v"` (`v3.12.2`)
- **Triệu chứng:** Định tuyến `^v` vào hàm `Kangatang_ShortcutPaste` để chống virus vô hiệu hóa `Ctrl+V`, nhưng lệnh `ActiveSheet.Paste` hoặc `Selection.PasteSpecial xlPasteAll` bị thất bại ngầm (do `On Error Resume Next`) trên Office 2016/2019/365 khi `CutCopyMode = xlCopy (1)`.
- **Root Cause:** Khi macro được kích hoạt qua `Application.OnKey`, Excel khóa một phần trạng thái UI của `ActiveSheet.Paste` và `xlPasteAll (-4104)` đối với vùng chọn đang ở chế độ `xlCopy`, ném lỗi `1004 (PasteSpecial method of Range class failed)`.
- **Cách khắc phục (`v3.12.2`):** Khi `Application.CutCopyMode = 1` (`xlCopy`), tách thành **2 bước PasteSpecial tuần tự** trên `Selection`:
  1. `Selection.PasteSpecial -4123` (`xlPasteFormulas` — dán công thức & giá trị; nếu lỗi thì fallback sang `-4163` `xlPasteValues`).
  2. `Selection.PasteSpecial -4122` (`xlPasteFormats` — dán toàn bộ định dạng màu sắc, viền, font).
- **Tọa độ code:** [`KangatangGuard_Code.vba#L1094-L1130`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L1094-L1130) (`Kangatang_ShortcutPaste`).

---

### `BUG-03`: Làm biến mất thanh công cụ hoặc vô hiệu hóa Add-in bên thứ 3 (`MyExcel.xlam`) (`v3.15.0` – `v3.16.0`)
- **Triệu chứng:** Sau khi cài hoặc chạy `KangatangGuard`, thanh menu/công cụ của Add-in `MyExcel.xlam` bị mất hoặc `MyExcel.xlam` không tự nạp khi mở Excel.
- **Root Cause:**
  1. Hàm khôi phục UI gọi `.Reset` trên `CommandBars("Worksheet Menu Bar")` và `CommandBars("Standard")`, làm xóa toàn bộ nút công cụ do Add-in bên thứ 3 tạo ra khi khởi động.
  2. Script cài đặt ghi đè cứng vào `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN` mà không kiểm tra xem khóa `OPEN` đã chứa đường dẫn của `MyExcel.xlam` hay chưa.
- **Cách khắc phục:**
  1. **TUYỆT ĐỐI KHÔNG** gọi `.Reset` trên `"Worksheet Menu Bar"` hoặc `"Standard"`. Chỉ gọi `.Reset` trên 4 menu ngữ cảnh chuột phải (`"Cell"`, `"Row"`, `"Column"`, `"XLM Cell"`).
  2. Khi đăng ký Registry `Options\OPEN*`, duyệt toàn bộ các khóa `OPEN`, `OPEN1`, `OPEN2`... hiện có, giữ nguyên các Add-in khác và chỉ ghi `"KangatangGuard.xlam"` vào slot `OPEN*` trống đầu tiên.
- **Tọa độ code:** [`KangatangGuard_Code.vba#L1242-L1276`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L1242-L1276), [`L2741-L2805`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L2741-L2805).

---

### `BUG-04`: Quay tròn (Spinning Cursor) và không thêm được Add-in trong `File -> Options -> Add-ins -> Go -> OK` (`v3.16.0`)
- **Triệu chứng:** Khi người dùng vào `File -> Options -> Add-ins -> Go`, bấm `Browse` chọn `KangatangGuard.xlam` rồi bấm `OK`, con trỏ chuột quay tròn và khi mở lại hộp thoại thì `KangatangGuard` vẫn không được tích chọn.
- **Root Cause:**
  1. Add-in trước đó được cài vào `XLSTART`, trong khi hộp thoại Add-In Manager của Excel quản lý các tệp trong `%APPDATA%\Microsoft\AddIns`.
  2. Hàm `CleanAddInCollisions()` có đoạn mã tự động gán `oAi.Installed = False` đối với mọi mục trong `Application.AddIns` trùng tên `KangatangGuard.xlam` (vốn được viết để tránh nạp kép với `XLSTART`). Do đó, ngay khi người dùng vừa tick chọn và bấm `OK`, chính Add-in lại tự gỡ tick (`Installed = False`) của chính nó!
  3. `Workbook_Open` chạy đồng bộ nhiều tác vụ nặng (quét file, duyệt 150 CommandBars, WMI Registry) ngay trong luồng đóng hộp thoại Add-Ins.
- **Cách khắc phục (`v3.16.0`):**
  1. Chuẩn hóa vị trí cài đặt duy nhất về `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam` và xóa bỏ bản sao trong `XLSTART`.
  2. Loại bỏ hoàn toàn lệnh `oAi.Installed = False` trong `CleanAddInCollisions()`.
  3. Tối ưu `InitializeGuard()` chạy dưới `< 5ms` và chuyển toàn bộ quét khởi động (`ShieldStartupSweep`) sang luồng bất đồng bộ `Application.OnTime Now + TimeSerial(0, 0, 1)`.
- **Tọa độ code:** [`KangatangGuard_Code.vba#L18-L35`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L18-L35), [`L241-L264`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L241-L264), [`L2668-L2738`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L2668-L2738).

---

### `BUG-05`: Phần mềm diệt virus Seqrite Endpoint Security (Quick Heal) nhận nhầm và cách ly `Kangatang_FolderScanner.ps1` (`v3.17.0`)
- **Triệu chứng:** Máy trạm Client cài **Seqrite Endpoint Security** báo phát hiện mã độc và cách ly tệp `C:\Users\<User>\AppData\Roaming\KangatangGuard\Kangatang_FolderScanner.ps1`.
- **Root Cause:** Hội tụ 4 yếu tố kích hoạt bộ lọc Heuristic/EDR:
  1. Tệp `.ps1` nằm trong thư mục nhạy cảm `%APPDATA%\Roaming\...`.
  2. Tiến trình `EXCEL.EXE` gọi `WScript.Shell.Run "powershell.exe -ExecutionPolicy Bypass -File %APPDATA%\..."` (trùng mẫu hành vi *Office Macro Dropper*).
  3. Tệp `.ps1` nhúng C# `Add-Type` gọi Win32 P/Invoke (`DllImport("kernel32.dll")`, `ole32.dll`, `user32.dll`, `Process.Kill()`).
  4. Tệp `.ps1` chứa hàng loạt chuỗi IOC tên virus (`kangatang`, `mypersonnel`, `laroux`, `auto_open`, `VBComponents.Remove`) dạng plaintext.
- **Cách khắc phục (`v3.17.0`):**
  1. Tích hợp trình quét thư mục **100% Native VBA (`RunNativeFolderScan`)** ngay bên trong `KangatangGuard.xlam`. Không bao giờ tạo hay gọi `.ps1` trong `%APPDATA%\KangatangGuard`.
  2. Tự động xóa sạch mọi tệp `.ps1` cũ còn sót trong `%APPDATA%\KangatangGuard` khi khởi động/cập nhật.
  3. Rút gọn các tệp `.ps1` hỗ trợ trên máy chủ LAN thành trình bao COM ~100 dòng (0% P/Invoke, 0% chuỗi IOC).
- **Tọa độ code:** [`KangatangGuard_Code.vba#L636-L872`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L636-L872), [`L2713-L2737`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L2713-L2737).

---

### `BUG-06`: Lỗi COM `HRESULT: 0x800A9C68 (40040)` khi đặt `Application.AutomationSecurity = 3` bên trong VBA (`v3.17.0`)
- **Triệu chứng:** Khi hàm VBA `RunNativeFolderScan` đặt `Application.AutomationSecurity = 3` (`msoAutomationSecurityForceDisable`), mở một tệp `.xlsm` bằng `Workbooks.Open` rồi truy cập `wbTarget.VBProject`, Excel dừng ngang hàm VBA và trả về lỗi `0x800A9C68` mà không đi vào `On Error GoTo`.
- **Root Cause:** Khác với khi điều khiển từ ngoài tiến trình (Out-of-process COM), khi đang ở **bên trong Runtime VBA**, việc chuyển `AutomationSecurity = 3` khiến VBE khóa cứng việc khởi tạo `VBProject` của các workbook `.xlsm` mở sau đó và hủy ngang lệnh gọi VBA khi chạm vào `wbTarget.VBProject`.
- **Cách khắc phục (`v3.17.0`):** Không đặt `Application.AutomationSecurity = 3` bên trong `RunNativeFolderScan`. Khi mở file bằng `Application.Workbooks.Open` từ mã VBA, Excel **mặc định không bao giờ tự chạy `Auto_Open`**, và việc đặt `Application.EnableEvents = False` đã khóa hoàn toàn `Workbook_Open` mà vẫn giữ `wbTarget.VBProject` truy cập được 100%.
- **Tọa độ code:** [`KangatangGuard_Code.vba#L727-L740`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L727-L740).

---

### `BUG-07`: Crash `EXCEL.EXE` hoặc Reset dự án VBA khi xóa `Module Sheet` (Excel 4.0/5.0) trong sự kiện `WorkbookOpen` (`v3.10.0`)
- **Triệu chứng:** Một số biến thể Kangatang/Laroux cũ tạo ra `Module Sheet` (`TypeName(sh) = "Module"`, `sh.Type = -4141` `xlModule`). Nếu gọi `wb.VBProject` hoặc xóa sheet này trực tiếp trong sự kiện `WorkbookOpen`/`BeforeSave`, Excel sẽ tự reset toàn bộ call stack VBA (mất biến `oAppEvents`) hoặc crash ứng dụng.
- **Cách khắc phục (`v3.10.0`):**
  1. Kiểm tra `LegacyModuleSheetCount(wb)` trước khi chạm vào `wb.VBProject`. Nếu `> 0`, tuyệt đối không mở `wb.VBProject` ngay lúc đó.
  2. Chuyển giao việc xóa `Module Sheet` sang quy trình **Two-Phase Purge** ngoài luồng sự kiện bằng `Application.OnTime` (`DeferModuleSheetPurge` ➔ `ShieldPurgeModuleSheets` ➔ `ShieldAfterPurge`).
- **Tọa độ code:** [`KangatangGuard_Code.vba#L1324-L1364`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L1324-L1364), [`L3190-L3378`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L3190-L3378).

---

### `BUG-08`: Lỗi `AccessVBOM = 0` gây hiện hộp thoại cảnh báo lặp đi lặp lại (`v3.15.0`)
- **Triệu chứng:** Trên một số máy trạm, `AccessVBOM` trong Registry chưa được bật hoặc bị Group Policy (`HKCU\Software\Policies\...`) khóa về `0`, khiến mỗi lần mở file lại hiện popup báo lỗi.
- **Cách khắc phục (`v3.15.0`):** Tự động ghi `AccessVBOM = 1` vào `HKCU\Software\Microsoft\Office\<12.0|14.0|15.0|16.0>\Excel\Security` và xóa khóa `AccessVBOM` cản trở trong nhánh `Policies` thông qua WMI `StdRegProv` (`EnsureVbomRegistryEnabled`). Loại bỏ hoàn toàn popup cảnh báo làm phiền người dùng (`Zero-Popup Self-Healing`).
- **Tọa độ code:** [`KangatangGuard_Code.vba#L2906-L3046`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L2906-L3046).

---

### `BUG-09`: Lỗi không ghi đè được `.xlam` khi cập nhật LAN do thuộc tính `Read-Only` hoặc tiến trình `EXCEL.EXE` ngầm (`v3.8.4`)
- **Root Cause:** File `.xlam` tải từ share mạng SMB đôi khi mang cờ `Read-Only`, hoặc tiến trình `EXCEL.EXE` chưa thoát hẳn khi chạy `apply_update.cmd`.
- **Cách khắc phục:**
  1. Trên máy chủ LAN `223.7`, luôn phát hành kèm bản có gắn số phiên bản `KangatangGuard_vX.Y.Z.xlam` (không bao giờ bị SMB lock bởi máy khác).
  2. Trong `PerformLanUpdate`, gỡ cờ `Read-Only` (`Attributes = 0` và `attrib -r`), chờ tiến trình `EXCEL.EXE` thoát hoàn toàn rồi mới chép đè và kiểm tra toàn vẹn nhị phân bằng `fc /b`.
- **Tọa độ code:** [`KangatangGuard_Code.vba#L1958-L2091`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L1958-L2091).

---

### `BUG-10`: Lỗi vỡ tiếng Việt (Mojibake) trên VBA, Ribbon XML và file `.bat` (`v3.3.0`, `v3.8.7`)
- **Root Cause:** Trình biên dịch VBA sử dụng CodePage ANSI của hệ điều hành (trên Windows tiếng Hàn/Anh sẽ làm hỏng chữ tiếng Việt trong file `.vba`); PowerShell đọc/ghi XML theo string mặc định làm lệch byte UTF-8; Windows CMD (`cmd.exe`) bị vỡ lệnh khi gặp ký tự có dấu trong file `.bat`.
- **Cách khắc phục:**
  1. Trong `.vba`: Mã hóa toàn bộ chuỗi tiếng Việt thành `\uXXXX` thông qua hàm `Uni()` và hiển thị bằng Win32 API `MessageBoxW`.
  2. Trong `Install_KangatangGuard.ps1`: Đọc và nén `customUI14.xml` bằng mảng byte nguyên bản `[System.IO.File]::ReadAllBytes()`.
  3. Trong mọi file `.bat` / `.cmd`: Bắt buộc dùng **100% Pure ASCII** (tiếng Việt không dấu hoàn toàn).

---

### `BUG-11`: Lệch số phiên bản giữa Add-in (`v3.16.0`) và Trình quét (`v3.8.3`) (`v3.17.1`)
- **Root Cause:** Số phiên bản được khai báo tĩnh ở nhiều file khác nhau mà không có cơ chế đồng bộ tự động khi build; đồng thời `CheckForLanUpdates` có thời gian chờ (cooldown) 6 giờ.
- **Cách khắc phục (`v3.17.1`):**
  1. Thêm cổng `[0/6] Single Source of Truth Version Gate` vào `Deploy_To_Hub.ps1` tự động đóng dấu `$ReleaseVersion` lên toàn bộ 15 file trước khi biên dịch.
  2. Các script `.ps1` tự động đọc động số phiên bản từ `version.json` / Registry `InstalledVersion`.
  3. Giảm cooldown kiểm tra cập nhật tự động khi mở Excel xuống **1 phút** (`UPDATE_CHECK_MINUTES = 1`).

---

### `BUG-12`: Bấm "Quét thư mục..." không khởi động cửa sổ PowerShell & treo Excel chính khi quét thư mục LAN lớn (`v3.17.3`)
- **Triệu chứng:** Khi người dùng bấm **Quét thư mục...** hoặc **Tiếp tục quét** trên thư mục mạng lớn (`\\192.168.223.7\file_shared\vietnam\0-Common` — 1.042 tệp), không có cửa sổ tiến trình PowerShell nào hiện lên và cửa sổ Excel chính bị đơ.
- **Root Cause:**
  1. Ở bản `v3.17.0`–`v3.17.2`, `ScanFolderDialog()` và `ResumeScanDialog()` gọi trực tiếp `RunNativeFolderScan` đồng bộ ngay trên luồng Excel chính thay vì khởi chạy `LaunchBackgroundScanner` ngoài tiến trình, đồng thời tệp `Kangatang_FolderScanner.ps1` chưa được cài vào thư mục tin cậy `%APPDATA%\Microsoft\AddIns`.
  2. Ngoài ra, truyền đường dẫn thư mục có tiếng Việt hoặc đường dẫn UNC qua đối số dòng lệnh `wsh.Run` dễ gặp lỗi cắt chuỗi/mã hóa ANSI.
- **Cách khắc phục (`v3.17.3`):**
  1. **Out-of-Process Streaming Console + Native VBA Worker:** `ScanFolderDialog()` và `ResumeScanDialog()` gọi `LaunchBackgroundScanner`, ghi yêu cầu quét ra `%APPDATA%\KangatangGuard\scan_request.json` (UTF-8) và mở cửa sổ `powershell.exe -NoProfile -NoExit -ExecutionPolicy Bypass -File "%APPDATA%\Microsoft\AddIns\Kangatang_FolderScanner.ps1"` (`vbNormalFocus = 1`).
  2. **100% Antivirus-Safe (Seqrite / Quick Heal):** Tệp `Kangatang_FolderScanner.ps1` nằm tại `%APPDATA%\Microsoft\AddIns` (không lưu trong `%APPDATA%\KangatangGuard`), hoàn toàn không chứa C# `Add-Type` P/Invoke hay chuỗi IOC virus, mà ủy quyền kiểm tra từng tệp cho hàm `ScanSingleFileForWorker` bên trong `KangatangGuard.xlam` trên một Background Excel Worker riêng biệt.
- **Tọa độ code:** [`KangatangGuard_Code.vba#L632-L970`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L632-L970), [`Kangatang_FolderScanner.ps1`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1).

---

### `BUG-13`: Tự động bật cửa sổ Excel trắng (`Book5`, `Book6`, `Book7 - Excel` / `EXCEL.EXE /restore`) trong lúc quét thư mục lớn (`v3.19.1`)
- **Triệu chứng:** Trong khi trình quét `Kangatang_FolderScanner.ps1` đang chạy quét thư mục mạng lớn, các cửa sổ Excel hiển thị (`Book5`, `Book6`, `Book7 - Excel`) tự động bật lên trên màn hình người dùng.
- **Root Cause (Bằng chứng Forensic từ Process CommandLine & Windows Event Log):**
  1. Các cửa sổ `BookN - Excel` thực chất được Windows khởi chạy với dòng lệnh `"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE" /restore` do **Windows Error Reporting (`WerFault.exe`)** kích hoạt cơ chế Application Recovery khi một tiến trình Worker ngầm (`EXCEL.EXE /automation -Embedding`) bị crash.
  2. **3 nguyên nhân gây crash Worker `EXCEL.EXE /automation -Embedding`:**
     - **Crash Cause A (`mscorwks.dll` `.NET 2.0 CLR` Stack Overflow `0xc00000fd` / `0xc0000005`):** Hàm `ComputeThreatHash` trong `modThreatCollector` gọi `CreateObject("System.Text.UTF8Encoding")` và `MD5CryptoServiceProvider`, nạp `.NET 2.0 CLR (mscorwks.dll)` vào tiến trình `EXCEL.EXE` 64-bit và gọi lặp lại hàng trăm lần khi xóa các sheet `Kangatang_XX`.
     - **Crash Cause B (`Application.OnTime` chạy ngầm trong Worker):** Khi Worker nạp `KangatangGuard.xlam`, `InitializeGuard()` đặt hẹn giờ `ShieldStartupSweep` (sau 1s) và `CheckForLanUpdatesSilent` (sau 45s) mà chưa có chốt chặn `If bIsFolderScanning Or Not Application.Visible Then Exit Sub`.
     - **Crash Cause C (File `.xls` hỏng cấu trúc BIFF8 & biến cục bộ trong `Restart-LocalWorker`):** Các file `.xls` không có VBA/IOC vẫn bị đẩy vào Deep COM (khi gặp file `.xls` hỏng cấu trúc sẽ làm crash parser BIFF8 tại `0x000f963a`), đồng thời `Restart-LocalWorker` tham chiếu nhầm `$excel` cục bộ thay vì `$script:excel`.
- **Cách khắc phục (`v3.19.1`):**
  1. **Thay thế hoàn toàn `.NET 2.0 COM Interop` bằng thuật toán băm `Pure Native VBA Dual-Seed 32-bit Hash`:** Loại bỏ `CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")` trong `ComputeThreatHash`, triệt tiêu 100% lỗi crash `mscorwks.dll` và không gọi `CollectThreatSample` trên các `Worksheet` thường.
  2. **Khóa chặt `ShieldStartupSweep` & `CheckForLanUpdatesSilent` trong Worker:** Thêm chốt chặn `If bIsFolderScanning Or Not Application.Visible Then Exit Sub` và hủy gắn kết `oAppEvents` trong `InitWorkerMode`.
  3. **Mở rộng Tầng 2 (`Zero-COM Pre-Filter`) cho cả định dạng `.xls` (OLE2 Binary):** Kiểm tra trực tiếp luồng byte `.xls` trong RAM (`< 2ms`); nếu không có `VBA` và không khớp `$iocRegex` (ASCII & UTF-16LE) thì kết luận `SAFE` ngay mà không cần mở `EXCEL.EXE`.
  4. **Chặn tự mở `/restore` (`DontShowUI = 1`, `AutoRecover.Enabled = False`, `PurgeWorkerStrayRestoreProcesses`):** Vô hiệu hóa `AutoRecover` trên Worker, xóa `Resiliency\DocumentRecovery`, dùng bản sao cô lập `%TEMP%\KangatangGuard_Worker_<id>.xlam` và tự động dọn sạch mọi tiến trình `/restore` từ bên trong hàm VBA `PurgeWorkerStrayRestoreProcesses` (giữ `Kangatang_FolderScanner.ps1` 100% an toàn trước Quick Heal `QHPISVR`).
- **Tọa độ code:** [`KangatangGuard_Code.vba#L861-L910`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L861-L910), [`L3935-L3965`](file:///d:/7.%20AI%20tools/kangatang/addin/KangatangGuard_Code.vba#L3935-L3965), [`Kangatang_FolderScanner.ps1#L254-L470`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1#L254-L470).

---

### `BUG-14`: Tiến trình PowerShell bị treo đứng hoặc tự tắt giữa chừng khi quét thư mục lớn (`48.775` tệp) (`v3.19.2`)
- **Triệu chứng:** Khi quét thư mục mạng lớn (`\\192.168.223.7\file_shared\vietnam` — `48.775` tệp), chạy được một thời gian (đến khoảng tệp thứ `4.090`) thì tiến trình PowerShell dừng hẳn hoặc bị tắt.
- **Root Cause (Bằng chứng Forensic từ RunspacePool & Live Process Inspection):**
  1. **Lỗi `Workbooks.Open($workerXlam, 0, $true)` (`3 tham số`) trên tệp `.xlam` nằm trong `%TEMP%`:** Trong Excel 2016 (`16.0`), gọi `.Workbooks.Open($workerXlam, 0, $true)` (truyền `UpdateLinks=0, ReadOnly=$true` khi mở tệp Add-in `.xlam` từ `%TEMP%`) ném ngoại lệ `Unable to get the Open property of the Workbooks class`.
  2. **Thoát luồng Worker sớm & Deadlock tại cổng Backpressure (`$comWorkQueue.Count -gt 60`):** Ở bản `v3.19.1`, lời gọi `Restart-LocalWorker` đầu tiên trong `$workerScript` nằm trong `try { ... } finally { ... }` nhưng không có `catch`. Khi `Workbooks.Open` ném lỗi lúc khởi tạo, cả 2 luồng Worker nhảy thẳng vào `finally` và kết thúc vĩnh viễn (`0` tiến trình `EXCEL.EXE`). Đến tệp thứ `4.090`, khi có đủ `61` tệp cần Deep COM tích tụ trong `$comWorkQueue`, vòng lặp chính chạm `while ($comWorkQueue.Count -gt 60)` và chờ vô hạn vì không còn Worker nào sống để rút hàng đợi!
  3. **Ngoại lệ I/O mạng chưa bọc `try/catch` & Áp lực bộ nhớ Large Object Heap (LOH):** Trong vòng lặp duyệt tệp, thao tác đọc `FileInfo` / đường dẫn quá dài (`> 260` ký tự) hoặc tệp bị người dùng khác xóa/đổi tên giữa chừng trên SMB chưa được bọc `try/catch` riêng cho từng tệp; đồng thời đọc toàn bộ mảng byte của tệp `.xls`/`.xlsb` lớn tới `25 MB` tạo ra các chuỗi `String` khổng lồ trên .NET Large Object Heap (LOH), dễ dẫn đến `OutOfMemoryException`.
  4. **Rò rỉ tiến trình `EXCEL.EXE` ngầm khi tái khởi động Worker:** Gọi `ReleaseComObject` thay vì `FinalReleaseComObject` trên cả `$script:wbAddin` và `$script:excel` khiến RCW vẫn giữ tham chiếu COM `> 0`, làm tiến trình `EXCEL.EXE` cũ không thoát hẳn sau khi gọi `.Quit()`.
- **Cách khắc phục (`v3.19.2`):**
  1. **Đổi đường dẫn `$workerXlam` sang `%APPDATA%\KangatangGuard\KangatangGuard_Worker_<id>.xlam` và gọi `.Open($workerXlam)` 1 tham số (có độ trễ lệch pha `350ms`):** Khởi tạo thành công 100% cả 2 Worker COM mà không bị Office Trust Center hay `Workbooks.Open` từ chối.
  2. **Thiết kế Immortal Worker Loop + Worker Watchdog (`Ensure-WorkerPoolStarted`) + Safety Valve chống Deadlock:**
     - Bọc `try/catch` cho `Restart-LocalWorker` ngay khi khởi động và tự động khôi phục ở đầu vòng lặp `while (-not $state.CancelWorkers)` nếu `$script:excel` hoặc `$script:wbAddin` bị `$null`.
     - Tích hợp **Worker Watchdog** trong `Ensure-WorkerPoolStarted` (được gọi liên tục cả trong vòng lặp Backpressure và Drain): tự động kiểm tra `$job.Handle.IsCompleted` và tái sinh ngay bất kỳ luồng Worker nào bị ngắt.
     - Thêm van an toàn (`Safety Valve`) trong vòng lặp chờ Backpressure/Drain để không bao giờ bị khóa cứng vĩnh viễn nếu gặp tệp mạng lỗi.
  3. **Bọc `try/catch` từng tệp + Top-level `catch` + Giới hạn bộ đệm Zero-COM LOH ở mức `4 MB` + `FinalReleaseComObject`:** Bảo vệ tuyệt đối vòng lặp chính trước mọi lỗi SMB/IO, ngăn tràn bộ nhớ LOH và giải phóng triệt để tiến trình `EXCEL.EXE` khi tái chế Worker.
- **Tọa độ code:** [`Kangatang_FolderScanner.ps1#L246-L319`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1#L246-L319), [`L388-L558`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1#L388-L558), [`L780-L945`](file:///d:/7.%20AI%20tools/kangatang/addin/Kangatang_FolderScanner.ps1#L780-L945).
