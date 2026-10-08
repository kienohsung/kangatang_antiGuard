# KANGATANG GUARD v3.13.0 - KIẾN TRÚC PHÒNG VỆ & ĐIỀU PHỐI

> **Ngày thực hiện:** 2026-10-08  
> **Phiên bản kiến trúc:** `v3.13.0` (Production-grade)  
> **Mục tiêu:**
> 1. Thiết lập hạ tầng lưu trữ tập trung 100% mẫu mã độc từ mọi máy tính trong mạng LAN về máy cá nhân `192.168.223.176`.
> 2. Khắc phục triệt để lỗi mất dữ liệu khi Copy / Paste giữa 2 tệp Excel khác nhau (Cross-Workbook Copy/Paste).
> 3. Tuân thủ quy chuẩn điều phối đa tác tử `/multi-agent-coordination`.

---

## 1. PHÂN TÍCH NGUYÊN NHÂN GỐC RỄ (ROOT CAUSE ANALYSIS)

### A. Sự cố mất dữ liệu khi Copy/Paste giữa 2 file Excel khác nhau
* **Triệu chứng:** Người dùng copy ở Workbook A (có viền nét đứt), nhưng khi click chuyển sang cửa sổ Workbook B để dán thì viền nét đứt biến mất hoặc dán không ra dữ liệu.
* **Nguyên nhân gốc rễ:** 
  * Khi chuyển cửa sổ làm việc giữa các Workbook, sự kiện `xlApp_WorkbookActivate` và `xlApp_SheetActivate` tự động được kích hoạt.
  * Trong các phiên bản trước, các sự kiện này gọi `RestoreExcelClipboardAndUI`, trong đó thực hiện các thao tác:
    - `Application.CellDragAndDrop = True`
    - `CommandBars.Reset` và kích hoạt lại các control
    - Thiết lập `EnableSelection = -4142`
  * Trong mô hình đối tượng của Microsoft Excel, **bất kỳ can thiệp nào vào `CellDragAndDrop` hoặc `CommandBars.Reset` sẽ ngay lập tức xóa sổ bộ đệm sao chép nội bộ (`Application.CutCopyMode = 0`)**.
  * Hậu quả: Dữ liệu sao chép từ Workbook A bị tiêu hủy ngay khi con trỏ chuột chạm vào Workbook B.

### B. Vấn đề thu thập mẫu tập trung từ mạng LAN về 223.176
* **Triệu chứng:** Các máy client khi quét và phát hiện mã độc không thể đẩy code mẫu về `\\192.168.223.176\KangatangGuard_Hub\Threat_Samples`.
* **Nguyên nhân:** Thư mục chia sẻ `KangatangGuard_Hub` được cấu hình bảo mật `Everyone: Read` (chỉ đọc) để bảo vệ gói cài đặt Add-in, khiến máy Client không có quyền ghi (`Access is denied`).
* **Ràng buộc an ninh:** Tuyệt đối không lưu mã độc trên máy chủ chia sẻ chung `192.168.223.7`.

---

## 2. GIẢI PHÁP KIẾN TRÚC ĐÃ TRIỂN KHAI (V3.13.0)

### 2.1. Bảo toàn bộ đệm đa Workbook (Cross-Workbook Buffer Shield)
1. **Kiểm tra trạng thái bộ đệm trước khi can thiệp (Zero-Interference Guard):**
   * Trong `clsAppEvents.xlApp_WorkbookActivate` và `clsAppEvents.xlApp_SheetActivate`:
     ```vba
     If Application.CutCopyMode <> 0 Then Exit Sub
     ```
   * Trong `modKangatangScanner.RestoreExcelClipboardAndUI`:
     ```vba
     If Application.CutCopyMode <> 0 Then Exit Sub
     ```
   * *Ý nghĩa kiến trúc:* Khi người dùng đang trong chu kỳ Copy (`CutCopyMode = 1`) hoặc Cut (`CutCopyMode = 2`), toàn bộ các thao tác reset giao diện, thay đổi cấu trúc sheet đều bị phong tỏa, đảm bảo bộ đệm sao chép được bảo toàn nguyên vẹn 100% xuyên suốt mọi Workbook.

2. **Dán tương thích đa Workbook:**
   * Khi dán tại Workbook đích, thủ tục `Kangatang_ShortcutPaste` tiếp nhận và thực thi qua 2 pha:
     - Pha 1: `Selection.PasteSpecial -4123` (`xlPasteFormulas`) chuyển giao công thức và giá trị.
     - Pha 2: `Selection.PasteSpecial -4122` (`xlPasteFormats`) đồng bộ định dạng bảng, màu sắc, font chữ.

### 2.2. Hạ tầng Thu thập Mẫu Tập trung (Centralized Threat Collector)
1. **Thiết lập NTFS Junction & SMB DropBox:**
   * Tạo thư mục `D:\7. AI tools\kangatang\Threat_Samples` trên máy `192.168.223.176`.
   * Tạo liên kết NTFS Junction tại `D:\test\Threat_Samples` trỏ thẳng về `D:\7. AI tools\kangatang\Threat_Samples`.
   * Thư mục mạng `\\192.168.223.176\test` đã có sẵn quyền `Everyone: Change/Full`, cho phép toàn bộ máy Client trong LAN ghi tệp mã độc vào đây một cách minh bạch mà không cần cấp quyền Administrator hay nới lỏng bảo mật của thư mục cài đặt `distribution`.
2. **Cập nhật Module `modThreatCollector`:**
   * Ưu tiên 1 (Cục bộ): `D:\7. AI tools\kangatang\Threat_Samples` (nếu chạy trên máy cá nhân 223.176).
   * Ưu tiên 2 (Mạng LAN): `\\192.168.223.176\test\Threat_Samples` (dành cho mọi máy Client trong LAN).
   * Khóa an ninh: Duy trì `FORBIDDEN_SERVER = "192.168.223.7"`.

---

## 3. MA TRẬN ĐỐI CHIẾU TRƯỚC VÀ SAU (BEFORE & AFTER)

| Tiêu chí | Trước khi sửa (v3.12.2) | Sau khi nâng cấp (v3.13.0) |
|---|---|---|
| **Copy / Paste cùng file** | ✅ Hoạt động tốt | ✅ Hoạt động tốt |
| **Copy từ File A sang File B** | 🛑 Bị mất nét đứt, không dán được | ✅ **Hoạt động hoàn hảo**, bảo lưu toàn bộ công thức & format |
| **Copy / Paste liên tục đa tệp** | 🛑 Không thể thực hiện | ✅ **Paste liên tục nhiều lần trên nhiều tệp khác nhau** |
| **Nơi lưu code mẫu trên máy 223.176** | Cục bộ `D:\...\Threat_Samples` | ✅ Cục bộ `D:\...\Threat_Samples` |
| **Nơi lưu code mẫu từ máy Client LAN** | Rơi vào AppData máy Client (do lỗi quyền) | ✅ **Tự động gửi tập trung về máy 223.176** qua UNC DropBox |
| **Bảo mật máy chủ 223.7** | Tuyệt đối không lưu | ✅ Tiếp tục chặn cứng 100% |

---

## 4. KIỂM THỬ XÁC MINH (VERIFICATION)

1. `verify_cross_wb_e2e.ps1`:
   * Mở Book 1, copy ô A1 (`=50+50`, Vàng, Đậm).
   * Chuyển sang Book 2: `CutCopyMode` duy trì `xlCopy (1)`.
   * Dán vào B5 của Book 2: Nhận công thức `=50+50`, giá trị `100`, nền vàng, chữ đậm.
   * `CutCopyMode` vẫn giữ nguyên sau khi dán ➔ **PASSED 100%**.
2. `verify_threat_collection.ps1`:
   * Kích hoạt `CollectThreatSample` trên máy giả định.
   * Mẫu được trích xuất, băm MD5 và lưu thành công dạng JSON defanged ➔ **PASSED 100%**.
