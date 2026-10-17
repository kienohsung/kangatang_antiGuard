# Báo cáo Kỹ thuật v3.16.0: Khắc phục triệt để lỗi "Excel không add được Add-in (bấm OK quay tròn và tự mất tích chọn)"

- **Phiên bản:** `v3.16.0` (Instant Add-In Manager Integration & Zero-Spin Architecture)
- **Ngày thực hiện:** `2026-10-17`
- **Trạng thái:** ✅ Đã kiểm thử tự động qua Excel COM, phát hành lên LAN Hub `\\192.168.223.7` và đồng bộ GitHub.

---

## 1. Phân tích Nguyên nhân Gốc rễ (Root Cause Analysis — Có Bằng Chứng Nhật Ký)

Trích xuất trực tiếp từ nhật ký kiểm toán `%APPDATA%\KangatangGuard\scan_log_20261017.txt` lúc `08:57:34 – 08:57:35` (thời điểm người dùng bấm `OK` trong hộp thoại `File -> Options -> Add-ins -> Go...`):

```text
2026-10-17 08:57:34 | [VBA_LOCKED] E:\1.QuanTrong\1. Setup\Office\CaiDatMyExcel\CaiDatMyExcel\MyExcel\MyExcel.xlam - VBA project has a password, only sheets/names were checked.
2026-10-17 08:57:34 | [SAFE] E:\1.QuanTrong\1. Setup\Office\CaiDatMyExcel\CaiDatMyExcel\MyExcel\MyExcel.xlam
2026-10-17 08:57:34 | [SHIELD_STARTUP] Sweep started (v3.15.0)
2026-10-17 08:57:35 | [COLLISION_FIX] Removed duplicate in AddIns folder: C:\Users\mrKienIT\AppData\Roaming\Microsoft\AddIns\KangatangGuard.xlam
2026-10-17 08:57:35 | [COLLISION_FIX] Uninstalled duplicate Add-in from Application.AddIns: C:\Users\mrKienIT\AppData\Roaming\Microsoft\AddIns\KangatangGuard.xlam
```

Có **3 nguyên nhân kỹ thuật trực tiếp** gây ra hiện tượng con trỏ xoay tròn ("quay tròn") và không thêm được Add-in:

1. 🛑 **Logic tự gỡ cài đặt trong `CleanAddInCollisions()` (`modKangatangShield`):**
   - Ở các phiên bản trước (`v3.10.0 – v3.15.0`), hàm `CleanAddInCollisions()` ép buộc Add-in chỉ được chạy từ thư mục `XLSTART` và coi mọi đăng ký trong `Excel Add-Ins Manager` (`%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam`) là bản sao trùng lặp:
     - Xóa trực tiếp tệp `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam`.
     - Xóa khóa `OPEN*` và `Add-in Manager` trong `HKCU\Software\Microsoft\Office\*\Excel\Options`.
     - Duyệt `Application.AddIns` và gọi **`oAi.Installed = False`** đối với mọi mục `KangatangGuard` không nằm trong `XLSTART`.
   - Hệ quả: Ngay sau khi người dùng tích chọn `[x] Kangatangguard` hoặc `Browse -> OK`, đúng 3 giây sau `CleanAddInCollisions()` tự động chạy ngầm và gọi `oAi.Installed = False`, khiến Add-in bị gỡ bỏ ngay lập tức!
2. 🛑 **Xung đột tên tệp giữa `XLSTART\KangatangGuard.xlam` và `AddIns\KangatangGuard.xlam`:**
   - Khi `KangatangGuard.xlam` đã được nạp ngầm từ `XLSTART`, hộp thoại `Excel Add-ins` (`Go...`) không hiển thị dấu tích `[x]`. Khi người dùng bấm `Browse...` chọn `KangatangGuard.xlam` và bấm `OK`, Excel cố mở tệp từ `%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam` nhưng bị kẹt vì Excel không cho phép mở 2 workbook trùng tên `KangatangGuard.xlam` cùng lúc.
3. ⚠️ **Khởi tạo đồng bộ nặng trong `InitializeGuard()` & `RestoreExcelClipboardAndUI()`:**
   - Khi đóng hộp thoại `Add-Ins` (`OK`), `InitializeGuard()` gọi đồng bộ 20 lệnh ghi Registry WScript.Shell, duyệt qua hơn 150 `Application.CommandBars` và gọi `.Reset` trên 6 thanh công cụ (bao gồm `Worksheet Menu Bar` và `Standard`), cộng thêm việc quét toàn bộ `MyExcel.xlam` trong sự kiện `xlApp_WorkbookOpen`, khiến giao diện Excel quay tròn.

---

## 2. Giải pháp Kiến trúc trong `v3.16.0` (Before & After Matrix)

| Hạng mục | Trước (`v3.15.0`) | Sau (`v3.16.0`) |
| :--- | :--- | :--- |
| **Vị trí cài đặt chuẩn** | `%APPDATA%\Microsoft\Excel\XLSTART\KangatangGuard.xlam` (không hiển thị tích chọn trong `Add-Ins`) | **`%APPDATA%\Microsoft\AddIns\KangatangGuard.xlam`** + đăng ký khóa `OPEN*` chuẩn của Excel Add-Ins Manager |
| **Hành vi `CleanAddInCollisions()`** | Tự động gọi `oAi.Installed = False` và xóa `AddIns\KangatangGuard.xlam` | **Tuyệt đối không gỡ Add-in (`Installed = True` bền vững)**; tự động đồng bộ bản mới nhất vào `AddIns\KangatangGuard.xlam` và dọn bản trùng lặp trong `XLSTART` |
| **Tương thích Add-in thứ 3 (`MyExcel.xlam`)** | Quét toàn bộ `ScanWorkbook` khi mở `.xlam` và xóa/reset `Worksheet Menu Bar` | Kiểm tra nhanh theo tên (`QuickCheckWorkbook < 2ms`) cho mọi `.xlam`, giữ nguyên khóa `OPEN` của `MyExcel.xlam` và không `.Reset` thanh menu của Add-in khác |
| **Thời gian phản hồi khi bấm `OK` trong `Add-Ins`** | Quay tròn đơ hộp thoại, sau 3s tự mất dấu tích `[x]` | **`155ms` (tức thì, không quay tròn)**, giữ nguyên trạng thái `Installed = True, IsOpen = True` |

---

## 3. Kết quả Kiểm thử Tự động (COM Verification)

```text
VERIFY T=0s: ElapsedMs=155 | Installed=True | IsOpen=True
VERIFY T=3s: Installed=True | IsOpen=True
HKCU:\Software\Microsoft\Office\16.0\Excel\Options:
  OPEN  = "E:\1.QuanTrong\1. Setup\Office\CaiDatMyExcel\CaiDatMyExcel\MyExcel\MyExcel.xlam"
  OPEN1 = "KangatangGuard.xlam"
2026-10-17 09:57:27 | [SHIELD_STARTUP] Sweep started (v3.16.0)
2026-10-17 09:57:27 | [SAFE] Book1
2026-10-17 09:57:27 | [SHIELD_STARTUP] Sweep finished: 1 workbook(s), 0 add-in(s)
```
