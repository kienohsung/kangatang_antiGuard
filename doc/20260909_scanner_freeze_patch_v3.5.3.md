# KẾ HOẠCH BẢN VÁ KIẾN TRÚC: CHỐNG TREO DỪNG TIẾN TRÌNH QUÉT MẠNG & TỰ ĐỘNG PHỤC HỒI (v3.5.3)

**Ngày lập:** 2026-09-09  
**Phiên bản mục tiêu:** 3.5.3 (Zero-Hang Watchdog & Safe Network Cleaning Architecture)  
**Tác giả:** Tech Lead kiêm Senior Full-stack Architect  

---

## 1. PHÂN TÍCH HIỆN TƯỢNG & NGUYÊN NHÂN GỐC RỄ (ROOT CAUSE ANALYSIS)

Dựa trên bằng chứng màn hình thực tế từ người dùng (KangatangGuard v3.5.1 - quét được 6585 tệp, diệt 652 tệp):
- Quá trình quét đang chạy trên đường dẫn mạng nội bộ UNC: \\192.168.223.7\file_shared\....
- Tại tệp thứ 6584 (OSV-SW202606.xlsx), hệ thống diệt thành công 12 sheet virus Kangatang.
- Ngay sau đó, tiến trình bước vào tệp thứ 6585 (OSV-SW202607.xlsx hoặc tệp kế tiếp) và **bị dừng/treo hoàn toàn suốt 12 phút**, không in thêm bất kỳ dòng nào, tiêu đề dừng ở Da quet: 6585.
- Đồng thời, trước đó có các lỗi rải rác:
  - The server threw an exception (RPC_E_SERVERFAULT)
  - [LOI] Loi trong qua trinh lam sach: You cannot call a method on a null-valued expression

### Các nguyên nhân cốt lõi:
1. **Thiếu cơ chế Giám sát Thời gian Chờ (Watchdog / Hard Timeout):**
   - Lệnh gọi COM Workbooks.Open() hoặc $wb.Save() được thực hiện đồng bộ trên luồng chính của PowerShell.
   - Khi gặp một tệp bị khóa bởi người dùng khác trên mạng (File in Use), tệp có mật khẩu bảo vệ, hoặc đường truyền mạng SMB bị nghẽn: Excel COM mở một hộp thoại ngầm (Modal Dialog) chờ người dùng bấm nút (Password prompt, File locked prompt). Do Visible = False, hộp thoại này vô hình, và tiến trình bị dừng vô thời hạn.
2. **Xung đột Khóa tệp qua mạng SMB khi Diệt Virus:**
   - Sau khi phát hiện virus, script đóng tệp Read-Only và ngay lập tức mở lại ở chế độ Read-Write. Trên mạng SMB, server chưa kịp nhả khóa (Oplock/File Handle) dẫn đến việc mở ghi bị từ chối hoặc trả về $null.
   - Script không kiểm tra $wb -eq  mà gọi ngay $wb.VBProject và $wb.Sheets, gây lỗi You cannot call a method on a null-valued expression.
3. **Cảnh báo Bảo mật & Tương thích khi Lưu tệp (Privacy / Compatibility Alerts):**
   - Khi lưu tệp đã khử virus, Excel có thể bật cảnh báo Privacy Warning hoặc Compatibility Checker, làm treo lệnh .Save().

---

## 2. GIẢI PHÁP KIẾN TRÚC BẢN VÁ v3.5.3 (VACCINES)

1. 🛡️ **Cơ chế Watchdog Timeout Chống Treo Tuyệt Đối (Per-File Hard Timeout):**
   - Bọc lệnh xử lý tệp trong cơ chế kiểm soát thời gian tối đa (Timeout: **20 giây** cho mỗi tệp).
   - Sử dụng PowerShell Runspace / Job / Timer bất đồng bộ.
   - Nếu một tệp xử lý vượt quá 20 giây: Tự động cưỡng chế kết liễu tiến trình Excel (dựa trên PID), ghi log [TIMEOUT] Tệp bị treo quá 20s (Do khóa mạng/mật khẩu/hộp thoại), tự động kích hoạt [AUTO-RECOVERY] tạo phiên Excel mới và **chuyển ngay sang tệp tiếp theo, không bao giờ bị dừng lại**.
2. 🛡️ **Kiểm Tra Khóa Tệp Trước Khi Mở Ghi (Pre-flight Write-Lock Check):**
   - Trước khi mở lại chế độ Read-Write để diệt virus trên mạng SMB, script thử mở khóa tệp bằng .NET [System.IO.File]::Open(, 'Open', 'ReadWrite', 'None').
   - Nếu tệp đang bị nhân viên khác mở: Không để Excel bật hộp thoại File in Use, mà lập tức ghi log [FILE_LOCKED_BY_USER] Tệp đang được mở bởi người khác, sao lưu và xử lý sau để bảo toàn tiến trình.
3. 🛡️ **Phòng Thủ Kiểm Tra NULL Tuyệt Đối Trong Khối Làm Sạch:**
   - Thêm kiểm tra nghiêm ngặt if ( -eq ) { ... return } sau khi mở lại tệp để làm sạch.
   - Khóa toàn bộ các hộp thoại cảnh báo khi lưu: $wb.CheckCompatibility = False, $wb.RemovePersonalInformation = False.
4. 🛡️ **Mở Tệp Phòng Vệ Chặn Hộp Thoại Mật Khẩu:**
   - Cung cấp mật khẩu giả ngẫu nhiên khi mở: nếu tệp có mật khẩu bảo vệ, Excel sẽ báo lỗi sai mật khẩu ngay lập tức thay vì dừng lại chờ người dùng nhập.

---

## 3. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Tiêu chí | Trước bản vá (v3.5.1) | Sau bản vá (v3.5.3) |
| :--- | :--- | :--- |
| **Kiểm soát thời gian mở/lưu** | 🛑 Đồng bộ vô thời hạn, treo vĩnh viễn nếu gặp hộp thoại | 🛡️ **Watchdog Timer (20s)**: Tự động ngắt và sang tệp tiếp theo |
| **Tệp bị khóa trên mạng (SMB)** | 🛑 Bật hộp thoại File in use vô hình làm dừng tiến trình | 🛡️ **Pre-flight Check**: Phát hiện tệp khóa và bỏ qua an toàn |
| **Xử lý tệp có mật khẩu** | 🛑 Hiện hộp thoại nhập mật khẩu vô hình, treo máy | 🛡️ Bắn lỗi mật khẩu ngay lập tức, không bật hộp thoại |
| **Bảo vệ COM khi làm sạch** | 🛑 Dễ văng lỗi 
ull-valued expression làm hỏng COM | 🛡️ Kiểm tra NULL nghiêm ngặt, tự động hồi sinh COM nếu lỗi |
| **Độ ổn định khi quét > 10,000 tệp** | 🛑 Có thể bị dừng bất kỳ lúc nào nếu gặp 1 tệp lỗi | ✅ **Chạy xuyên suốt không gián đoạn (Unattended Run)** |
