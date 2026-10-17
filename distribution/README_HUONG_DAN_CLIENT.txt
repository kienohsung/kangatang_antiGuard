==============================================================================
   HƯỚNG DẪN SỬ DỤNG BỘ CÔNG CỤ DIỆT VIRUS EXCEL KANGATANG (v3.20.2)
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
6. Mở Excel lên, bạn sẽ thấy thanh công cụ "KangatangGuard" trên Ribbon.

------------------------------------------------------------------------------
CÁCH 2: CHẠY QUÉT ĐỘC LẬP (DÀNH CHO MÁY KHÔNG CÀI ĐƯỢC ADD-IN)
------------------------------------------------------------------------------
Áp dụng khi máy tính chỉ muốn quét sạch virus 1 lần cho thư mục dữ liệu:

1. Mở thư mục chia sẻ máy chủ (hoặc tải gói addin_kangatang.zip về giải nén):
   \\192.168.223.7\file_shared\vietnam\z. ETC\1. addinKangatang
2. Bấm đúp chuột vào tệp:
   Chay_Diet_Virus_Ngoai.bat
3. Chọn thư mục cần quét trên hộp thoại để chương trình tự động kiểm tra và diệt sạch.

* LƯU Ý:
- Phiên bản v3.20.2 tích hợp 100% Native VBA Folder Scanner bên trong KangatangGuard.xlam,
  hoàn toàn tương thích với mọi phần mềm Antivirus/EDR (Seqrite Endpoint Security / Quick Heal).
- Mọi tệp nhiễm virus trước khi xử lý đều được tự động sao lưu cách ly an toàn về máy chủ 192.168.223.176.
==============================================================================
