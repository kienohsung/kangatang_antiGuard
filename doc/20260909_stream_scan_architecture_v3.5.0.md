# THIẾT KẾ KIẾN TRÚC QUÉT LUỒNG TRỰC TIẾP (STREAMING SCANNER): KANGATANGGUARD v3.5.0
**Mã tài liệu:** `doc/20260909_stream_scan_architecture_v3.5.0.md`  
**Phiên bản:** `v3.5.0` (Real-Time Directory Streaming Architecture)  
**Chức danh:** Tech Lead kiêm Senior Full-stack Architect  
**Trạng thái:** Chờ phê duyệt (Awaiting Approval)

---

## 1. PHÂN TÍCH NHU CẦU CẢI TIẾN (REQUIREMENT ANALYSIS)

### 🛑 Hạn chế của cơ chế 2 giai đoạn (Pre-collection Phase):
- **Cơ chế cũ (v3.4.0):**
  - **Giai đoạn 1:** Chạy `Get-ChildItem -Recurse` duyệt qua toàn bộ cây thư mục để gom tất cả file vào danh sách mảng (`$filesToScan`).
  - **Giai đoạn 2:** Sau khi gom xong 100% tệp mới bắt đầu mở Excel COM và quét tệp đầu tiên.
- **Vấn đề phát sinh:**
  - Đối với các cây thư mục rất lớn hoặc các **ổ mạng UNC doanh nghiệp (`\\192.168.223.7\...`)** có độ trễ mạng (Network Latency), giai đoạn 1 gom tệp có thể mất từ vài chục giây đến vài phút. Người dùng phải ngồi chờ mà chưa thấy tệp nào được quét.
  - Tốn bộ nhớ RAM để lưu trữ mảng đối tượng FileInfo của toàn bộ ổ đĩa trước khi xử lý.

### 🛡️ Giải pháp Kiến trúc v3.5.0: Động cơ Quét Luồng Trực Tiếp (Real-Time Streaming Pipeline)
- Loại bỏ hoàn toàn giai đoạn gom tệp trước (Zero Pre-indexing Delay).
- Áp dụng mô hình **Duyệt Thư mục Luồng Tuần tự (Folder-by-Folder Streaming)**:
  - Khi bắt đầu: Khởi động ngay Excel COM độc lập trong **0.2 giây**.
  - Duyệt vào thư mục cấp 1: Phát hiện tệp Excel nào lập tức nạp và **quét ngay tệp đó**!
  - Xong các tệp trong thư mục hiện tại: Lần lượt đi tiếp vào từng thư mục con để quét tiếp (Depth-First / Breadth-First Stream).
  - **Độ trễ khởi động = 0 giây**: Người dùng nhìn thấy tiến trình kiểm tra tệp đầu tiên ngay lập tức!

---

## 2. BẢNG ĐỐI CHIẾU HIỆN TRẠNG (BEFORE & AFTER MATRIX)

| Tiêu chí Đánh giá | Bản v3.4.0 (Gom tệp rồi mới quét) | Bản v3.5.0 (Quét luồng trực tiếp ngay) | Giá trị Vượt trội |
| :--- | :--- | :--- | :--- |
| **Độ trễ bắt đầu quét** | Chờ gom đủ 100% tệp (mất 10s - 2 phút trên ổ mạng) | **0 giây (Bắt đầu quét ngay tệp đầu tiên)** | 🛡️ Phản hồi tức thì |
| **Cơ chế duyệt** | Chờ toàn bộ cây thư mục hoàn thành | **Quét lần lượt từng thư mục, gặp file là quét ngay** | ✅ Minh bạch, quan sát được tiến độ từng thư mục |
| **Tiêu thụ bộ nhớ** | Giữ toàn bộ mảng hàng nghìn FileInfo trong RAM | **Xử lý luồng (Stream), giải phóng bộ nhớ liên tục** | 🛡️ Tối ưu tài nguyên hệ thống |
| **Thông tin hiển thị** | Tiến độ % trên tổng số tệp ước lượng | **Hiển thị trực quan: Tên thư mục hiện tại + Bộ đếm tệp thời gian thực (Real-time Counter)** | ✅ Trực quan, sinh động |

---

## 3. THIẾT KẾ CHI TIẾT ĐỘNG CƠ STREAMING (`Kangatang_FolderScanner.ps1`)

```powershell
function Scan-FolderStream {
    param ([string]$CurrentFolder)
    
    # Bo qua thu muc sao luu _Backup_Kangatang
    if ($CurrentFolder -like "*_Backup_Kangatang*") { return }

    Write-Host "`n📂 [QUÉT THƯ MỤC] $CurrentFolder" -ForegroundColor Yellow

    # 1. Quet ngay lap tuc tat ca file Excel trong thu muc nay
    $files = Get-ChildItem -Path $CurrentFolder -File -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        if ($excelExtensions -contains $file.Extension.ToLower() -and $file.Name -ne "KangatangGuard.xlam") {
            $Script:TotalScanned++
            # Goi ham kiem tra va lam sach ngay lap tuc!
            Scan-And-Clean-SingleFile $file
        }
    }

    # 2. Duyet lan luot cac thu muc con va de quy quet tiep
    $subDirs = Get-ChildItem -Path $CurrentFolder -Directory -ErrorAction SilentlyContinue
    foreach ($sub in $subDirs) {
        if ($sub.Name -ne "_Backup_Kangatang") {
            Scan-FolderStream -CurrentFolder $sub.FullName
        }
    }
}
```

---

## 4. KẾ HOẠCH TRIỂN KHAI
1. **Task 1:** Cập nhật `addin\Kangatang_FolderScanner.ps1` sang kiến trúc quét luồng trực tiếp (Streaming Pipeline) không chờ gom tệp.
2. **Task 2:** Cập nhật phiên bản đồng bộ `v3.5.0` trong `KangatangGuard_Code.vba`, `Install_KangatangGuard.ps1` và `Diet_Virus_Kangatang.bat`.
3. **Task 3:** Biên dịch và triển khai tự động bản `v3.5.0` vào `%APPDATA%\KangatangGuard\` và `%APPDATA%\Microsoft\Excel\XLSTART\`.
4. **Task 4:** Đồng bộ mã nguồn ra thư mục Desktop.
