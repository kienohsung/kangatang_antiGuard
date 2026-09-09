# KẾ HOẠCH NÂNG CẤP & SỬA LỖI KIẾN TRÚC: KANGATANGGUARD v3.3.0
**Mã tài liệu:** `doc/20260908_fix_folder_scan_unicode_v3.3.0.md`  
**Phiên bản:** `v3.3.0`  
**Chức danh:** Tech Lead kiêm Senior Full-stack Architect  
**Trạng thái:** Chờ phê duyệt (Awaiting Approval)

---

## 1. NGUYÊN NHÂN GỐC RỄ (ROOT CAUSE ANALYSIS)

Dựa trên việc đối chiếu nhật ký hệ thống thực tế tại `%APPDATA%\KangatangGuard\scan_log_20260908.txt` và phân tích mã nguồn nhị phân đang nạp trong Excel, chúng tôi xác định được **3 nguyên nhân gốc rễ**:

### 🛑 1. Lỗi quét thư mục trả về 0 tệp (0 trường hợp) trong 1 giây
- **Nguyên nhân 1.1 (Mã nhị phân chưa được nạp mới):** Tệp `.xlam` đang chạy trong Excel thực tế vẫn là bản biên dịch ngày `2026-09-03 09:35` (`v3.2.0`). Bản này vẫn sử dụng lệnh VBA `Dir(folderPath & "*.*")`.
- **Nguyên nhân 1.2 (Lỗi giới hạn của lệnh `Dir()` trong VBA):**
  - Trên các đường dẫn mạng UNC (`\\192.168.223.7\file_shared\...`) hoặc đường dẫn chứa tiếng Việt có dấu (`D:\4. Security check\2. Ohsungvina\10. Mua bán\`), hàm `Dir()` của VBA **ngay lập tức trả về chuỗi rỗng `""`**, khiến vòng lặp duyệt tệp kết thúc ngay trong 0.1 giây mà không kiểm tra được tệp nào.
- **Nguyên nhân 1.3 (Thiếu cơ chế quét đệ quy thư mục con - Recursive Scan):** Bản code cũ chỉ quét tệp ở thư mục gốc, nếu người dùng chọn thư mục cha chứa các thư mục con thì toàn bộ tệp Excel bên trong thư mục con đều bị bỏ sót.

### 🛑 2. Giao diện bị lỗi font tiếng Việt (Mất dấu, hiển thị `?`)
- **Nguyên nhân:** Trình soạn thảo VBA (VBA IDE) trên hệ điều hành Windows mặc định lưu trữ mã nguồn theo bảng mã **ANSI (Windows-1252)**. Khi inject chuỗi tiếng Việt có dấu trực tiếp vào `CodeModule`, các ký tự Unicode ngoài bảng mã ANSI (`Đ, ệ, ế, ư, ơ, ấ...`) bị VBA tự động chuyển thành dấu hỏi chấm (`?`) hoặc ký tự rác.
- Hộp thoại `MsgBox` mặc định của VBA cũng là hộp thoại ANSI, không thể hiển thị Unicode nguyên bản.

### 🛑 3. Lỗi Backup thất bại với tệp có dấu: `Bad file name or number`
- **Nguyên nhân:** Lệnh `FileCopy` trong VBA không hỗ trợ Unicode tiếng Việt, dẫn đến việc khi phát hiện virus, hệ thống không thể tạo bản sao lưu an toàn nên buộc phải dừng thao tác làm sạch.

---

## 2. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Thành phần | Phiên bản v3.2.0 (Hiện tại) | Phiên bản v3.3.0 (Đề xuất Mới) | Đánh giá Giá trị |
| :--- | :--- | :--- | :--- |
| **Duyệt tệp quét thư mục** | Dùng `Dir()` (chết trên UNC & Unicode, không quét thư mục con) | Dùng **FileSystemObject (FSO)** + **Duyệt đệ quy (Recursive)** toàn bộ thư mục con | 🛡️ Quét sạch 100% mọi tệp trên ổ mạng UNC & thư mục con |
| **Hiển thị Tiếng Việt** | Dùng chuỗi thường + `MsgBox` ANSI (bị biến thành dấu `?`) | **Bộ giải mã `Uni()`** (chuỗi thoát mã Unicode không phụ thuộc ANSI) + **`MessageBoxW` (API Windows 32/64-bit)** | ✅ Tiếng Việt có dấu 100% sắc nét, chuẩn giao diện Windows |
| **Menu điều khiển Excel** | Caption ANSI (lỗi dấu `?`) | Caption dùng chuẩn chuỗi giải mã Unicode | ✅ Đẹp mắt, chuẩn xác ngữ nghĩa |
| **Sao lưu tệp dự phòng** | Dùng `FileCopy` (báo lỗi `Bad file name`) | Dùng **`FileSystemObject.CopyFile`** | 🛡️ Sao lưu thành công mọi tệp Unicode |
| **Quy trình Triển khai** | Chỉ sửa file text `.vba`, chưa nạp vào `.xlam` | **Tự động đóng Excel COM và nạp nhị phân trực tiếp** vào cả `XLSTART` và `AddIns` | 🛡️ Hiệu lực tức thì, không cần thao tác thủ công |

---

## 3. THIẾT KẾ KIẾN TRÚC MẪU PHÒNG THỦ (VACCINE PATTERNS v3.3.0)

### 🛡️ Mẫu Thiết kế 1: Động cơ quét đệ quy FSO chuyên dụng (Recursive FSO Engine)
```vba
Private Sub CollectExcelFiles(ByVal oFolder As Object, ByRef fileList As Object, ByVal excludeFileName As String)
    ' 1. Quét tệp trong thư mục hiện tại
    For Each f In oFolder.Files
        ' Kiểm tra phần mở rộng: .xls, .xlsm, .xlsb, .xltm, .xlt, .xla, .xlam
        ' Bỏ qua tệp excludeFileName và chính KangatangGuard.xlam
    Next f
    ' 2. Đệ quy duyệt tất cả thư mục con (Bỏ qua thư mục _Backup_Kangatang)
    For Each subF In oFolder.SubFolders
        If UCase(subF.Name) <> UCase(BACKUP_FOLDER_NAME) Then
            Call CollectExcelFiles(subF, fileList, excludeFileName)
        End If
    Next subF
End Sub
```
- **Ưu điểm:** Khắc phục triệt để lỗi của `Dir()`, thu thập toàn bộ danh sách tệp vào `Dictionary` trước khi mở file, hỗ trợ cả đường dẫn mạng UNC (`\\192.168.223.7\...`) và đường dẫn nội bộ.

### 🛡️ Mẫu Thiết kế 2: Chuẩn hóa Tiếng Việt có dấu bằng `Uni()` & `MessageBoxW`
- Để tránh việc VBA IDE trên Windows biến tiếng Việt thành dấu `?`, toàn bộ thông điệp giao diện sẽ được mã hóa an toàn qua hàm `Uni()` với mã thoát Unicode chuẩn:
  ```vba
  #If VBA7 Then
      Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
  #Else
      Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
  #End If
  
  Public Function MsgBoxW(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "KangatangGuard") As VbMsgBoxResult
      MsgBoxW = MessageBoxW(Application.Hwnd, StrPtr(prompt), StrPtr(title), buttons)
  End Function
  ```
- **Kết quả:** Hiển thị 100% tiếng Việt có dấu chuẩn giao diện Windows native, không phụ thuộc vào Region/Locale hay CodePage của máy tính.

---

## 4. KẾ HOẠCH TRIỂN KHAI CHI TIẾT

1. **Task 1:** Cập nhật `KangatangGuard_Code.vba` lên phiên bản `v3.3.0`:
   - Bổ sung `MsgBoxW` và bộ giải mã `Uni()`.
   - Thay thế hoàn toàn cơ chế quét thư mục bằng `CollectExcelFiles` đệ quy qua FSO.
   - Thay thế `FileCopy` bằng `FSO.CopyFile`.
   - Chuẩn hóa toàn bộ thông báo, thực đơn Menu sang tiếng Việt có dấu.
2. **Task 2:** Thực thi quy trình biên dịch tự động:
   - Đóng các tiến trình Excel đang chạy.
   - Chạy `Install_KangatangGuard.ps1` để biên dịch trực tiếp ra file nhị phân `KangatangGuard.xlam` mới.
   - Phân phối đồng thời vào `%APPDATA%\Microsoft\Excel\XLSTART` và `%APPDATA%\Microsoft\AddIns`.
3. **Task 3:** Xác thực thực tế:
   - Kiểm tra thuộc tính `LastWriteTime` của file `.xlam` đảm bảo đã cập nhật thành công.
   - Đồng bộ đầy đủ sang thư mục Desktop của người dùng.
