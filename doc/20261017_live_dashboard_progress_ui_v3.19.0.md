# KẾ HOẠCH KIẾN TRÚC: GIAO DIỆN BẢNG ĐIỀU KHIỂN QUÉT CỐ ĐỊNH (`v3.19.0` — NON-SCROLLING LIVE DASHBOARD UI)

## 1. Mục Tiêu Nghiệp Vụ & Phạm Vi Giải Quyết
- **Vấn đề hiện tại (từ ảnh chụp thực tế của người dùng):**
  1. Giao diện cửa sổ PowerShell in nối tiếp từng thư mục và từng tệp theo chiều dọc (`Timeline Scrolling`), khiến màn hình trôi liên tục khi quét hàng trăm thư mục con, người dùng bị cuốn theo dòng chữ chạy và khó nắm bắt bức tranh tổng thể.
  2. Thiếu thanh trạng thái phần trăm tổng thể (`Đã xử lý %` / `Còn lại bao nhiêu %` / `ETA dự kiến`).
  3. Chưa có bảng tổng hợp số liệu trực tuyến ngay trong lúc quét (phải chờ quét xong mới thấy bảng tổng kết).
  4. Một số tên tệp tiếng Việt sử dụng bảng mã Unicode tổ hợp (`NFD`) hiển thị thành ký tự ô vuông (`□`) trên Windows Console (`conhost.exe`).
- **Mục tiêu thiết kế (`v3.19.0`):**
  - Chuyển đổi hoàn toàn sang **Giao diện Bảng Điều Khiển Cố Định Tại Chỗ (`In-Place Fixed Console Dashboard`)** — không cuộn dòng (`Zero-Scroll`), tự động làm mới mượt mà tại cùng một vị trí tọa độ con trỏ màn hình:
    1. **Dòng 1 (Thông tin vị trí đang thực hiện):** Hiển thị rõ **Đường dẫn thư mục đang thực hiện** và **Tên tệp đang xử lý** (tự động rút gọn thông minh ở giữa `\\192.168.223.7\...\11. IT_Report\10. Mua ban` để không bao giờ tràn dòng).
    2. **Dòng 2 (Process Status & Phần trăm còn lại):** Thanh tiến trình trực quan `[################..........] 62.4% | Còn lại: 37.6% (ETA: 01p 15s) | Tốc độ: 38.5 tệp/s`.
    3. **Khối 3 (Bảng Kết Quả Thống Kê Trực Tuyến):** Bảng kẻ khung cố định thể hiện đầy đủ:
       - Tổng số tệp phát hiện & số thư mục đã duyệt
       - Số tệp đã kiểm tra (phân tách chi tiết: *Fast Resume*, *Zero-COM OpenXML*, *Deep Excel COM*)
       - Số tệp **An toàn (Sạch)**
       - Số tệp **Phát hiện nhiễm virus**
       - Số tệp **Đã diệt & sao lưu thành công** (về máy `223.176`)
       - Số tệp **Bỏ qua / Khóa mật khẩu**
    4. **Khối 4 (Nhật ký 5 tệp nhiễm vừa xử lý gần nhất):** Giữ cố định 5 dòng cuối dưới đáy bảng để người dùng theo dõi được tên các tệp nhiễm virus vừa bị tiêu diệt mà vẫn không làm trôi màn hình.

---

## 2. Bảng Đối Chiếu Hiện Trạng (`Before & After Matrix`)

| Thành phần giao diện | Hiện trạng (`v3.17.3` / `v3.18.0`) | Thiết kế mới (`v3.19.0` — Live Dashboard UI) |
| :--- | :--- | :--- |
| **Cơ chế hiển thị Console** | In nối tiếp từng thư mục/tệp xuống dòng mới (`Write-Host`), gây trôi màn hình liên tục. | 🛡️ **Khóa tọa độ con trỏ (`[Console]::SetCursorPosition`)**: Toàn bộ giao diện cố định trên 1 khung màn hình duy nhất (`Zero-Scroll`). |
| **Dòng 1: Thư mục & Tệp đang quét** | Nằm rải rác theo từng khối `Folder [N]: ...` và `[OK] [M] ...`, bị trôi mất ngay khi sang thư mục khác. | ✅ **Cố định ở đầu Dashboard:**<br>`[THU MUC DANG QUET] : [Folder #66] \\192.168.223.7\...\10. Mua ban`<br>`[TEP DANG XU LY]    : don de nghi sua chua - RAM Laptop.xlsx` |
| **Dòng 2: Process Status & % Còn lại** | Chỉ hiển thị tiến độ cục bộ trong từng thư mục (`10/11`), chưa có `% hoàn thành` và `% còn lại` của toàn đợt quét. | ✅ **Thanh Process Status toàn cục:**<br>`[TIEN TRINH XU LY]  : [############........] 62.4% \| Con lai: 37.6% (ETA: 01p 12s)`<br>Sử dụng luồng đếm tổng số tệp song song (`Background Directory Counter Runspace`) để tính `%` chính xác mà không làm chậm lúc khởi động. |
| **Bảng kết quả tổng hợp** | Chỉ in ra ở cuối chương trình sau khi đã quét xong 100%. | ✅ **Bảng Kết Quả Trực Tuyến (Live Table):** Cập nhật tức thời số tệp đã quét, tệp sạch, tệp nhiễm, tệp đã xử lý/sao lưu và tệp lỗi ngay trong lúc quét. |
| **Hiển thị tên tệp tiếng Việt (`□`)** | Tên tệp dùng Unicode tổ hợp (`NFD`) bị lỗi ô vuông `□` trên `conhost.exe`. | 🛡️ Chuẩn hóa `chcp 65001` + chuẩn hóa chuỗi Unicode `FormC` (`NFC`) và rút gọn độ dài cố định giúp khung bảng không bao giờ bị lệch cột. |

---

## 3. Thiết Kế Bố Cục Màn Hình Chi Tiết (`Wireframe v3.19.0`)

```text
================================================================================
   KANGATANG GUARD v3.19.0 - BANG DIEU KHIEN QUET THU MUC (LIVE DASHBOARD)
   Day la addin phong chong virus kangatang - Phat trien boi OhsungIT team !
================================================================================
 [THU MUC DANG QUET] : [Thu muc #66] \\192.168.223.7\...\11. IT_Report\10. Mua ban
 [TEP DANG XU LY]    : don de nghi sua chua - RAM Laptop mr. Ba.xlsx
--------------------------------------------------------------------------------
 [TRANG THAI XU LY]  : [################..........]  62.4% | Con lai: 37.6% (ETA: 00p 45s)
================================================================================
 | DANH MUC THONG KE KET QUA                  |       SO LUONG | CHI TIET / TY LE             |
 +--------------------------------------------+----------------+------------------------------+
 | 1. Tong so tep Excel trong thu muc         |      1,042 tep | 142 thu muc (Toc do: 38.5/s) |
 | 2. So tep da kiem tra xong                 |        650 tep | Con lai: 392 tep (37.6%)     |
 |    - Bo qua nhanh (Fast Resume)            |        210 tep | Da sach tu phien truoc       |
 |    - Quet sieu toc Zero-COM (.xlsx)        |        415 tep | 1.5 ms/tep (Khong mo Excel)  |
 |    - Quet sau Excel COM (.xls / Macro)     |         25 tep | 4 Parallel Excel Workers     |
 +--------------------------------------------+----------------+------------------------------+
 | 3. Tep AN TOAN (Sach)                      |        640 tep | 98.5%                        |
 | 4. Tep PHAT HIEN NHIEM VIRUS               |         10 tep | Phat hien virus Kangatang    |
 | 5. Tep DA DIET & SAO LUU (ve 223.176)      |         10 tep | Da lam sach & luu an toan    |
 | 6. Tep BO QUA (Khoa mat khau / ReadOnly)   |          0 tep | Khong the ghi de             |
 +--------------------------------------------+----------------+------------------------------+
 [NHAT KY PHAT HIEN & XU LY VIRUS GAN NHAT]:
   > [13:15:02] DA DIET: Infected_Report_01.xls (- Module doc hai: Kangatang)
   > [13:15:09] DA DIET: Plan_2026.xlsm (- Sheet an doc hai: Kangatang)
================================================================================
```

---

## 4. Thiết Kế Kỹ Thuật & Chiến Lược Phòng Thủ (Defensive Design)

1. **Tính toán chính xác `% Hoàn thành` và `% Còn lại` không gây trễ khởi động (`Async Background File Counter`):**
   - Ngay khi khởi động `Kangatang_FolderScanner.ps1`, một luồng nhẹ (`Runspace` độc lập) sẽ chạy ngầm để đếm tổng số tệp Excel trong toàn bộ cây thư mục đích (`$sharedState.TotalTargetFiles`), trong khi luồng chính **bắt đầu quét ngay lập tức từ giây đầu tiên (`0.1s`)**.
   - Nếu là phiên **Tiếp tục quét (`-Resume`)** và `last_session.json` đã có sẵn `TotalFiles`, hệ thống nạp ngay con số này làm mẫu số ban đầu và tự động tinh chỉnh khi luồng đếm ngầm hoàn tất.
   - Công thức:
     - `PercentDone = Round((TotalScanned / TotalTargetFiles) * 100, 1)`
     - `PercentRemaining = Round(100.0 - PercentDone, 1)`
2. **Cơ chế Vẽ Lại Khung Cố Định Không Nhấp Nháy (`Flicker-Free In-Place Redraw`):**
   - Tắt con trỏ nhấp nháy trong lúc quét (`[Console]::CursorVisible = $false`) và bật lại trong khối `finally` (`[Console]::CursorVisible = $true`).
   - Sử dụng hàm `Pad-FixedWidth($text, 78)` kết hợp `[Console]::SetCursorPosition(0, $script:dashboardTop)` để ghi đè đúng từng dòng mà không dùng `Clear-Host` (nhờ đó màn hình **không bị chớp/nhấp nháy**).
   - Điều tiết tần suất vẽ giao diện (`UI Throttle = 120ms/lần` hoặc khi kết thúc thư mục / phát hiện virus) để không tốn CPU cho việc vẽ màn hình.
   - Hỗ trợ Fallback tự động nếu chạy trong môi trường перенаправление luồng (non-interactive pipe/test) không có Console Buffer.
3. **Chuẩn hóa Hiển thị Ký Tự Tiếng Việt trên Console (`Unicode NFC + Safe Truncation`):**
   - Kích hoạt `chcp 65001` và chuẩn hóa tên tệp/thư mục về dạng Unicode dựng sẵn (`NormalizationForm.FormC`) trước khi in ra màn hình.

---

## 5. Kế Hoạch Kiểm Thử & Triển Khai

1. Cập nhật `addin/Kangatang_FolderScanner.ps1`, `addin/KangatangGuard_Code.vba` và `addin/Deploy_To_Hub.ps1` lên phiên bản **`v3.19.0`**.
2. Chạy `addin/Deploy_To_Hub.ps1` để biên dịch `.xlam`, đồng bộ `Kangatang_FolderScanner.ps1` vào `%APPDATA%\Microsoft\AddIns\`, `distribution/` và máy chủ LAN `223.7`.
3. Chạy kiểm thử thực tế trên thư mục mẫu để xác nhận bảng Dashboard cập nhật mượt tại chỗ, hiển thị đúng `% Đã xử lý`, `% Còn lại` và bảng thống kê chi tiết.
