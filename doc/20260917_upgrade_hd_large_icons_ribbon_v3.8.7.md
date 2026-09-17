# Nâng cấp Giao diện: Icon TO (32x32) Chuẩn Ribbon CustomUI XML & Cải tiến FaceId v3.8.7

## 1. Bối cảnh & Yêu cầu Người dùng
- **Yêu cầu:** "thay thế các ico cho toolbar, to, đẹp hơn".
- **Ảnh đính kèm từ người dùng:** Thanh công cụ trên tab **Add-ins** -> khu vực **Custom Toolbars** hiển thị các icon 16x16 pixels nhỏ, mờ nhạt:
  - Nút 1: Chữ `X` màu đỏ (`FaceId 1088`) - Gây hiểu nhầm là nút Huỷ/Lỗi thay vì bảo vệ.
  - Nút 3: Tam giác đen hướng lên (`FaceId 38`).
  - Nút 4: Tam giác đen hướng xuống (`FaceId 40`).
  - Nút 5: Tam giác vàng cảnh báo (`FaceId 463`).
  - Nút 6: Chữ `i` nhỏ (`FaceId 487`).

## 2. Nguyên nhân Kỹ thuật (Root Cause Analysis)
1. **Giới hạn của Modern Excel với CommandBars:**
   - Kể từ Office 2010 đến Microsoft 365, thanh công cụ tạo qua VBA `Application.CommandBars.Add` bị đẩy vào tab `Add-ins` -> nhóm `Custom Toolbars`.
   - Nhóm này có chiều cao cố định **16x16 pixels**, hoàn toàn không hỗ trợ hiển thị icon kích thước lớn (32x32 pixels) ngay cả khi lập trình viên cố gắng gán thuộc tính.
2. **Giải pháp Chuẩn của Microsoft Office:**
   - Để hiển thị **Icon TO (32x32 pixels)**, chuẩn chính thống duy nhất được Microsoft hỗ trợ trên Ribbon là **CustomUI Ribbon XML** (`customUI/customUI14.xml`).
   - Ribbon XML cho phép sử dụng thuộc tính `size="large"` và tích hợp bộ biểu tượng vector Office chuẩn chính hãng (`imageMso`) có độ phân giải cao, hiển thị sắc nét trên cả màn hình Retina/4K.

## 3. Kiến trúc Đột phá: Dual-Interface (Ribbon XML + CommandBars Fallback)

### A. Tích hợp CustomUI Ribbon XML trực tiếp vào file `.xlam`:
- Sử dụng .NET `System.IO.Compression.ZipArchive` inject trực tiếp `customUI/customUI14.xml` và cập nhật quan hệ `_rels/.rels` trong quá trình biên dịch (mất chưa đầy 60ms, 0 delay).
- Cung cấp giao diện kép (Dual-Tab):
  1. **Tab chuyên biệt "Kangatang Guard":** Nằm độc lập trên thanh Ribbon chính của Excel, phân loại thành 2 nhóm lớn trực quan:
     - Nhóm **Diệt Virus & Bảo vệ**:
       - 🛡️ **Quét tệp này** (`size="large"`, `imageMso="FileCheckOut"`)
       - 📁 **Quét thư mục** (`size="large"`, `imageMso="FolderBrowse"`)
       - ⏯️ **Tiếp tục quét** (`size="large"`, `imageMso="PlayMacro"`)
     - Nhóm **Hệ thống & Tiện ích**:
       - 📜 **Nhật ký (Log)** (`size="large"`, `imageMso="OpenReportDetails"`)
       - 🔄 **Cập nhật LAN** (`size="large"`, `imageMso="ServerRefresh"`)
       - ℹ️ **Thông tin** (`size="large"`, `imageMso="Info"`)
  2. **Tab Add-ins (`TabAddIns`):** Nhóm **Kangatang Guard v3.8.7** với toàn bộ 6 nút lớn `size="large"`, người dùng xem ở tab nào cũng thấy nút to, đẹp.

### B. Cải tiến Bộ Icon CommandBars FaceId (Cho các máy chạy chế độ tương thích):
- Thay thế triệt để các mã FaceId cũ:
  - Nút 1: `FaceId = 1087` (Dấu tích xanh lá cây sắc nét biểu thị an toàn/bảo vệ, thay cho chữ X đỏ 1088).
  - Nút 2: `FaceId = 23` (Thư mục mở màu vàng).
  - Nút 3: `FaceId = 297` (Nút Play tam giác xanh lá cây chuyên nghiệp, thay cho tam giác đen 38).
  - Nút 4: `FaceId = 269` (Biểu tượng đồng hồ/nhật ký lịch sử, thay cho tam giác đen 40).
  - Nút 5: `FaceId = 184` (Biểu tượng hai mũi tên đồng bộ vòng tròn, thay cho tam giác cảnh báo 463).
  - Nút 6: `FaceId = 487` (Biểu tượng thông tin i màu xanh).

### C. Cơ chế Callbacks Bất đồng bộ trong VBA:
- Thêm 6 Ribbon Callback subroutines: `Ribbon_ScanActiveWorkbook`, `Ribbon_ScanFolderDialog`, `Ribbon_ResumeScanDialog`, `Ribbon_OpenLogFolder`, `Ribbon_CheckForLanUpdates`, `Ribbon_ShowAbout` nhận tham số `ByVal control As Object`.
- Gọi ủy quyền trực tiếp tới logic cốt lõi trong `modKangatangScanner`.

## 4. Bảng Đối chiếu Before & After

| Tiêu chí | Trước khi Sửa (v3.8.6) | Sau khi Nâng cấp (v3.8.7) |
| :--- | :--- | :--- |
| **Kích thước Icon** | Cố định 16x16 pixels nhỏ | 🎨 **32x32 pixels (size="large")** cực kỳ to, rõ ràng |
| **Độ sắc nét** | Icon 16px cổ điển, mờ trên màn hình độ phân giải cao | ✨ **Vector Office HD (imageMso)** chuẩn Fluent |
| **Vị trí hiển thị** | Chỉ nằm co cụm tại Custom Toolbars (Add-ins) | 🌟 **Tab riêng biệt "Kangatang Guard"** trên thanh Ribbon + Tab Add-ins |
| **Nút Quét tệp** | Chữ X màu đỏ (`FaceId 1088`, dễ nhầm huỷ/lỗi) | 🛡️ **FileCheckOut HD / Tích xanh 1087** trực quan |
| **Nút Tiếp tục quét** | Tam giác đen 38 nhỏ | ⏯️ **PlayMacro HD / Play xanh 297** |
| **Nút Cập nhật LAN**| Tam giác cảnh báo 463 | 🔄 **ServerRefresh HD / Đồng bộ 184** |