# Kiến trúc KangatangGuard v3.8.6: Core-Only Startup & Horizontal Icon Toolbar

## 1. Tổng quan & Mục tiêu Kỹ thuật

Phiên bản `v3.8.6` tập trung giải quyết triệt để 2 vấn đề lớn được người dùng phản ánh trong quá trình vận hành thực tế:
1. **Khắc phục tình trạng treo/đơ (lag) và loading lâu lúc khởi động Excel:** Rà soát và chỉ nạp duy nhất mã nguồn cốt lõi (Core-only) tại chu trình khởi động `Workbook_Open` / `InitializeGuard`.
2. **Thiết kế lại giao diện người dùng:** Chuyển đổi từ menu thả xuống (dropdown) thành Thanh công cụ ngang (`CommandBar` dạng `msoBarTop`) với các nút tính năng độc lập, mỗi tính năng mang 1 biểu tượng (Icon `FaceId`) trực quan và tiêu đề rõ ràng trải dài theo chiều ngang trên tab **Add-ins**.

---

## 2. Phân tích Nguyên nhân Gốc rễ & Giải pháp Tối ưu

### 🛑 Các điểm nghẽn gây chậm lúc khởi động:
1. **WMI Provider Call (`CleanDocumentRecoveryRegistry`):**
   - *Nguyên nhân:* Gọi `GetObject("winmgmts:\\.\root\default:StdRegProv")` ngay khi add-in nạp khiến các hệ thống EDR / Antivirus hook vào Excel, gây đơ từ 5–15 giây.
   - *Giải pháp:* Loại bỏ hoàn toàn WMI khỏi chu trình khởi động. Chỉ gọi dọn dẹp Document Recovery khi phát hiện và xử lý xong file nhiễm thực tế.
2. **Quét cưỡng bức các Workbook đang mở (`For Each Wb In Workbooks`):**
   - *Nguyên nhân:* Quét lặp đồng bộ lúc Excel đang dựng canvas giao diện làm nghẽn luồng UI chính.
   - *Giải pháp:* Loại bỏ vòng lặp quét khởi động. `clsAppEvents` (`xlApp_WorkbookOpen`) đã đảm nhiệm hoàn toàn việc quét bảo vệ tự động khi người dùng mở hoặc tạo workbook mới.
3. **Kiểm tra Cập nhật Mạng SMB không có Rate-Limit:**
   - *Nguyên nhân:* Gọi I/O kiểm tra đường dẫn UNC (`\\192.168.223.7\...`) sau 3 giây. Khi máy tính mất mạng hoặc dùng Wi-Fi yếu, Windows SMB timeout giữ chặt luồng UI từ 20–45 giây.
   - *Giải pháp:*
     - Trì hoãn 45 giây (`Now + TimeValue("00:00:45")`) để Excel hoàn toàn nhàn rỗi.
     - Rate-limit 24 giờ thông qua Registry `HKCU\Software\KangatangGuard\LastUpdateCheck`. Nếu trong ngày đã kiểm tra rồi, thoát ngay tức thì (0ms).

---

## 3. Thiết kế Giao diện Thanh công cụ Ngang (Horizontal Icon Toolbar)

Sử dụng `Application.CommandBars.Add(Name:="KangatangGuard", Position:=msoBarTop, Temporary:=True)` kết hợp thuộc tính `btn.Style = msoButtonIconAndCaption`. Giao diện tự động tích hợp vào tab **Add-ins** -> nhóm **Custom Toolbars** với trải nghiệm 1 chạm trực tiếp:

| STT | Nút Chức năng | Icon (FaceId) | Ý nghĩa Biểu trưng | Macro điều khiển |
| :---: | :--- | :---: | :--- | :--- |
| **1** | **Quét tệp này** | `1088` | Lá chắn bảo vệ kiểm tra file | `ScanActiveWorkbook` |
| **2** | **Quét thư mục...** | `23` | Thư mục mở chọn quét ngầm | `ScanFolderDialog` |
| **3** | **Tiếp tục quét** | `38` | Biểu tượng Play tiếp tục quét nhanh | `ResumeScanDialog` |
| **4** | **Nhật ký (Log)** | `40` | Sổ tay ghi chép nhật ký kiểm toán | `OpenLogFolder` |
| **5** | **Cập nhật LAN** | `463` | Biểu tượng đồng bộ mạng LAN | `CheckForLanUpdatesManual` |
| **6** | **Thông tin** | `487` | Biểu tượng Trợ giúp / Giới thiệu | `ShowAbout` |

---

## 4. Bảng Đối chiếu Trước & Sau Thay đổi (Before & After Matrix)

| Tiêu chí | Trước khi Sửa (v3.8.5) | Sau khi Sửa (v3.8.6 Production-grade) |
| :--- | :--- | :--- |
| **Thời gian khởi động Excel** | Lag / Treo từ 5 đến 45 giây (WMI & SMB timeout) | ⚡ **Tức thì (< 0.05s)** |
| **Tác vụ lúc Workbook_Open** | WMI Registry + Quét Workbooks + Log I/O + Check Hub | 🛡️ **Core-only:** Chỉ nạp Cache + Events + Toolbar |
| **Tần suất kiểm tra cập nhật** | Mọi lần mở Excel (nguy cơ treo khi offline) | 🛡️ **Tối đa 1 lần/24h**, hoãn 45s khi Excel rảnh |
| **Giao diện người dùng** | Menu dọc xổ xuống (dropdown), ẩn các tính năng | 🎨 **Thanh công cụ ngang**, mỗi tính năng 1 icon rõ ràng |
| **Tương tác sử dụng** | 2 nhịp click (bấm menu -> rê chuột chọn tính năng) | ⚡ **1 nhịp click**, trực quan và tiện dụng |
| **Phiên bản kiến trúc** | `v3.8.5` | **`v3.8.6`** |
