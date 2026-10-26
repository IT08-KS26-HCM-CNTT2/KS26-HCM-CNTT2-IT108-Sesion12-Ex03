
PHẦN 1. PHÂN TÍCH: MAPPING 4 THAM SỐ VLOOKUP VÀO MÔ HÌNH (NHÂN VIÊN KHO TRA HỒ SƠ)
Hàm VLOOKUP trong Excel được ví như nghiệp vụ tìm kiếm hồ sơ của một nhân viên kho. 
Dưới đây là bảng ánh xạ (Mapping) chi tiết 4 tham số của VLOOKUP theo mô hình thực tế:
1. Lookup_value (Giá trị cần tìm):
   - Ý nghĩa thực tế: "Mã vận đơn" hoặc "Mã đơn hàng" trên kiện hàng mà nhân viên kho đang cầm trên tay để tra cứu giá.
   - Ví dụ: Ô chứa mã vận đơn cần tra là `A2` (ví dụ: `FM1001`).
2. Table_array (Bảng chứa dữ liệu dò tìm):
   - Ý nghĩa thực tế: "Cuốn sổ tay tổng hợp bảng giá" hoặc "Danh mục cước phí" đặt sẵn trên bàn làm việc của kho.
   - Ví dụ: Vùng bảng giá nằm ở Sheet2, từ ô `A1:B50` (cần cố định tuyệt đối bằng dấu đô-la thành `$Sheet2!$A$1:$B$50`).
3. Col_index_num (Số thứ tự cột chứa kết quả cần lấy):
   - Ý nghĩa thực tế: "Số trang" hoặc "Cột số mấy" trong cuốn sổ tay ghi giá cước tương ứng với loại dịch vụ.
   - Ví dụ: Cột 1 là Mã vận đơn, Cột 2 là Giá cước -> Nhập số `2`.
4. Range_lookup (Kiểu tìm kiếm chính xác hay tương đối):
   - Ý nghĩa thực tế: "Quy tắc dò" (0/FALSE để tìm chính xác tuyệt đối mã đơn hàng; 1/TRUE để tìm khoảng giá tương đối).
   - Ví dụ: Bắt buộc chọn `FALSE` (hoặc `0`) để tìm kiếm chính xác tuyệt đối từng mã vận đơn, tránh lấy nhầm cước phí đơn khác.
PHẦN 2. TRIỂN KHAI NỘI DUNG GIẢI PHÁP CHO FASTMOVE
1. Tài liệu Word (Quy trình vận hành kho)
   - Thiết lập phong cách thương hiệu: Sử dụng tính năng Styles trong Microsoft Word 
     để đồng bộ hệ thống Heading (Heading 1, Heading 2) theo bộ nhận diện màu sắc của FastMove.
   - Đảm bảo tài liệu rõ ràng, chuẩn hóa bố cục cho shipper dễ đọc hiểu.
2. Bảng tính Excel (Tra cứu giá cước qua 2 Sheet khác nhau)
   - Nguyên tắc tuân thủ Database: Không dùng Merge Cell, không để trống hàng/cột trong vùng dữ liệu.
   - Công thức VLOOKUP kết hợp cố định tuyệt đối ($) để kéo xuống 500 dòng:
     + Giả sử cột A ở Sheet chính chứa Mã vận đơn (`A2`), bảng giá chuẩn nằm ở Sheet `BangGia` (cột A là mã, cột B là giá cước).
     + Công thức chuẩn tại ô B2: `=VLOOKUP(A2, BangGia!$A$1:$B$500, 2, FALSE)`
     + Ý nghĩa dấu `$`: Giúp cố định cố định vùng dữ liệu dò tìm của Sheet `BangGia` không bị trôi/lệch khi kéo copy công thức xuống 500 dòng bên dưới.
3. Slide PowerPoint (Báo cáo hiệu suất giao hàng KPI)
   - Quy tắc thiết kế theo Miller's Law: Giới hạn tối đa 3-4 bullet points trên mỗi slide để người xem dễ nhớ, dễ nắm bắt.
   - Tiêu đề Slide dạng Assertion Title (Khẳng định kết quả):
     + Ví dụ: "Tối ưu hóa quy trình giúp FastMove tiết kiệm 16 tiếng tra cứu cước mỗi ngày"
   - Trực quan hóa dữ liệu: Sử dụng Bar Chart (Biểu đồ thanh) để highlight rõ các mốc KPI giao hàng quan trọng.
PHẦN 3. KIỂM THỬ: XỬ LÝ LỖI #N/A KHI MÃ VẬN ĐƠN CÓ KÝ TỰ LẠ BẰNG IFERROR
Tình huống: Khi thực hiện hàm VLOOKUP tra cứu giá cước, một số mã vận đơn có chứa ký tự lạ hoặc lỗi phát sinh khiến hàm trả về giá trị lỗi `#N/A` làm mất thẩm mỹ bảng tính và gián đoạn báo cáo.
Giải pháp tích hợp hàm IFERROR:
1. Cách hoạt động: 
   - Hàm `IFERROR(value, value_if_error)` sẽ kiểm tra nếu biểu thức `value` (hàm VLOOKUP) trả về lỗi (như `#N/A`), nó sẽ tự động thay thế bằng một thông báo thân thiện hoặc giá trị rỗng, ngược lại sẽ giữ nguyên kết quả đúng.
2. Công thức áp dụng thực tế:
   - Sửa lại công thức tại ô B2 thành:
     `=IFERROR(VLOOKUP(A2, BangGia!$A$1:$B$500, 2, FALSE), "Chưa cập nhật giá")`
   - Hoặc để trống ô nếu chưa có dữ liệu:
     `=IFERROR(VLOOKUP(A2, BangGia!$A$1:$B$500, 2, FALSE), "")`

