==============================================================================
   HƯỚNG DẪN SỬ DỤNG BỘ CÔNG CỤ DIỆT VIRUS EXCEL KANGATANG (v3.8.1)
   Máy chủ Tệp Trung tâm: \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang
==============================================================================

Tùy theo nhu cầu và chính sách bảo mật máy tính, bạn có 2 cách sử dụng:

------------------------------------------------------------------------------
CÁCH 1: CÀI ĐẶT ADD-IN BẢO VỆ THƯỜNG TRỰC (KHUYÊN DÙNG)
------------------------------------------------------------------------------
Áp dụng cho máy trạm sử dụng Excel thường xuyên, muốn Excel tự động diệt virus ngầm
khi mở file và tự động nhận bản cập nhật mới nhất từ máy chủ 223.7:

1. Bấm tổ hợp phím Windows + R (mở hộp thoại Run).
2. Dán đường dẫn máy chủ:
   \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang
3. Bấm Enter để mở thư mục chia sẻ.
4. Bấm đúp chuột vào tệp:
   Install_Client_Kangatang.bat
5. Cửa sổ cài đặt tự động hoàn tất trong 2-3 giây.
6. Mở Excel lên, bạn sẽ thấy menu "KangatangGuard" trên thanh công cụ.

------------------------------------------------------------------------------
CÁCH 2: CHẠY QUÉT ĐỘC LẬP (DÀNH CHO MÁY KHÔNG CÀI ĐƯỢC ADD-IN)
------------------------------------------------------------------------------
Áp dụng khi máy tính bị chính sách công ty chặn cài Add-in, hoặc máy bạn chỉ muốn
quét sạch virus 1 lần cho máy tính/thư mục dữ liệu mà không muốn can thiệp vào Excel:

1. Mở thư mục chia sẻ máy chủ (hoặc tải gói addin_kangatang.zip về giải nén):
   \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang
2. Bấm đúp chuột vào tệp:
   Chay_Diet_Virus_Ngoai.bat
3. Cửa sổ dòng lệnh sẽ hiển thị Menu lựa chọn:
   [1] Quét Thư mục Tùy chọn (Hỗ trợ hộp thoại chọn thư mục hoặc nhập đường dẫn UNC)
   [2] Quét Toàn bộ Máy tính (Tất cả ổ đĩa C:, D:, E:...)
   [3] Tiếp tục Phiên quét dở dang (Fast Resume từ Checkpoint nếu từng bị ngắt)
   [4] Dọn dẹp Mầm bệnh Hệ thống (XLSTART, Registry, AddIns)
   [5] Mở Thư mục Nhật ký Kiểm toán (Audit Logs)
   [0] Thoát chương trình

* LƯU Ý KHI CHẠY:
- Không dùng chuột bôi đen văn bản trong cửa sổ đen (tránh Windows QuickEdit tạm dừng tiến trình).
- Nếu vô tình nhấp chuột làm màn hình đứng yên, chỉ cần nhấn phím ENTER để chương trình chạy tiếp.
- Mọi tệp nhiễm virus trước khi xử lý đều được tự động sao lưu an toàn vào thư mục "_Backup_Kangatang".
==============================================================================
