# 🧠 BẢN ĐỒ TRI THỨC KỸ THUẬT & TRA CỨU NHANH CHO AGENT (`LEARN.md`)
> **MỤC ĐÍCH:** Cổng chỉ mục phân tách giúp AI Agent tra cứu **Root Cause của Bug (`BUG-xx`)**, **Thuật toán tối ưu (`ALG-xx`)** và **Bài học kiến trúc qua các phiên bản (`ARC-xx`)** chỉ với **1 lần đọc (< 800 token)**, không cần đọc lại 23 tài liệu lịch sử trong `doc/`.

---

## 1. CẤU TRÚC PHÂN TÁCH THƯ MỤC `learn/`

| Tệp chuyên đề | Phạm vi tra cứu | Khi nào Agent cần đọc? |
| :--- | :--- | :--- |
| **[`learn/01_bugs_and_root_causes.md`](file:///d:/7.%20AI%20tools/kangatang/learn/01_bugs_and_root_causes.md)** | Từ điển 13 Bug thực tế (`BUG-01` ➔ `BUG-13`), nguyên nhân gốc rễ (Root Cause) & tọa độ code đã fix | Khi xử lý lỗi hồi quy, lỗi Copy/Paste, lỗi Add-in quay tròn, lỗi Antivirus cách ly, lỗi VBE/COM, lỗi tự bật cửa sổ Excel trắng |
| **[`learn/02_algorithms_and_optimizations.md`](file:///d:/7.%20AI%20tools/kangatang/learn/02_algorithms_and_optimizations.md)** | 10 Thuật toán & Kỹ thuật tối ưu (`ALG-01` ➔ `ALG-10`) đã kiểm chứng Production | Khi tối ưu hiệu năng quét, xử lý Clipboard, Fast Resume, Zero-COM OpenXML, Multi-Worker Pool, MD5 Dedup, Unicode |
| **[`learn/03_version_lessons_and_architecture.md`](file:///d:/7.%20AI%20tools/kangatang/learn/03_version_lessons_and_architecture.md)** | Tiến hóa kiến trúc `v3.1` ➔ `v3.19.1` & 6 Bài học thiết kế hệ thống LAN/EDR (`ARC-01` ➔ `ARC-06`) | Khi thiết kế tính năng mới, thay đổi cơ chế tự động cập nhật LAN hoặc triển khai sang máy chủ mới |

---

## 2. BẢNG ÁNH XẠ TRIỆU CHỨNG ➔ MÃ TRA CỨU NHANH (SYMPTOM ROUTING TABLE)

| Triệu chứng / Vấn đề kỹ thuật | Mã Bug liên quan (`01_bugs...`) | Mã Thuật toán (`02_algorithms...`) | Mã Kiến trúc (`03_version...`) |
| :--- | :--- | :--- | :--- |
| **Copy (`Ctrl+C`) có nét đứt nhưng Paste (`Ctrl+V`) mất nét đứt / không ra dữ liệu (cùng file hoặc giữa 2 file)** | `BUG-01`, `BUG-02` | `ALG-01` | `ARC-04` |
| **Mất thanh công cụ của Add-in khác (như `MyExcel.xlam`) khi cài `KangatangGuard`** | `BUG-03` | `ALG-02` | `ARC-04` |
| **Bấm OK trong `Excel Options -> Add-ins -> Go` bị quay tròn (Spinning) hoặc không tick chọn được Add-in** | `BUG-04` | `ALG-02` | `ARC-05` |
| **Bị Antivirus (Seqrite Endpoint Security / Quick Heal) nhận nhầm là mã độc và cách ly file `.ps1`** | `BUG-05`, `BUG-06` | `ALG-05` | `ARC-01`, `ARC-06` |
| **Lỗi COM `0x800A9C68 (40040)` khi quét thư mục bằng VBA trên tệp `.xlsm`** | `BUG-06` | `ALG-05` | `ARC-06` |
| **Excel bị crash cứng hoặc mất biến toàn cục VBA khi xóa sheet virus Excel 4.0/5.0 (`Module Sheet`)** | `BUG-07` | `ALG-03` | `ARC-03` |
| **Hiện popup lặp lại báo lỗi `Programmatic access to Visual Basic Project is not trusted` (`AccessVBOM`)** | `BUG-08` | `ALG-02` | `ARC-05` |
| **Lưu file (`Ctrl+S`) hoặc AutoSave bị khựng/lag, hoặc Excel đơ khi mạng LAN mất kết nối** | `BUG-09` | `ALG-04`, `ALG-06` | `ARC-02` |
| **Lỗi font tiếng Việt (Mojibake) trên hộp thoại thông báo, thanh Ribbon hoặc khi chạy file `.bat`** | `BUG-10` | `ALG-07` | `ARC-03` |
| **Số phiên bản giữa hộp thoại Giới thiệu Add-in và Trình quét virus bị lệch nhau** | `BUG-11` | — | `ARC-02`, `ARC-06` |
| **Bấm "Quét thư mục..." không hiện cửa sổ PowerShell hoặc quét thư mục lớn quá chậm (> 24 giờ)** | `BUG-12` | `ALG-09`, `ALG-10` | `ARC-01` |
| **Trong lúc quét thư mục, cửa sổ Excel trắng (`Book5`, `Book6`, `Book7 - Excel` / `EXCEL.EXE /restore`) tự bật lên** | `BUG-13` | `ALG-09`, `ALG-10` | `ARC-01`, `ARC-06` |

---

> 📌 **Tài liệu phối hợp:** Xem thêm [`AGENT_CONTEXT.md`](file:///d:/7.%20AI%20tools/kangatang/AGENT_CONTEXT.md) để tra cứu tọa độ dòng (`StartLine`–`EndLine`) của từng hàm trong `addin/KangatangGuard_Code.vba`.
