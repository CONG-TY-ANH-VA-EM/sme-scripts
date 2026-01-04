# Hướng dẫn Xử lý Sự cố & Câu hỏi Thường gặp (FAQ)

Tài liệu này giúp bạn tự giải quyết các vấn đề nhanh chóng khi sử dụng công cụ gửi phiếu lương tự động.

---

### 1. Tại sao tôi bấm nút "Gửi" mà không có chuyện gì xảy ra?
- **Nguyên nhân**: Google yêu cầu bạn cấp quyền truy cập lần đầu tiên.
- **Cách khắc phục**: Nhìn lên góc trên màn hình, nếu thấy thông báo "Authorization Required", hãy bấm **Continue** > Chọn tài khoản Gmail của bạn > Bấm **Advanced** (Nâng cao) > Bấm **Go to SME Scripts (unsafe)** > Bấm **Allow** (Cho phép).

### 2. Lỗi: "Không tìm thấy sheet dữ liệu T1"
- **Nguyên nhân**: Tên tab tháng của bạn không khớp với cấu hình.
- **Cách khắc phục**: Kiểm tra xem tab chứa lương tháng 1 của bạn có tên chính xác là `T1` hay không (không có dấu cách, không có chữ "Tháng"). Nếu bạn muốn đặt là `Tháng 1`, hãy sửa biến `SHEET_NAME_PREFIX` trong code từ `"T"` thành `"Tháng "`.

### 3. Tại sao trong phiếu PDF hiện thẻ {{HOTEN}} mà không hiện tên người?
- **Nguyên nhân**: Bạn đã gõ sai thẻ Tag trong tab `TEMPLATE`.
- **Cách khắc phục**: Thẻ phải có **đôi ngoặc nhọn** `{{` và `}}`, và viết **HOA toàn bộ** (ví dụ: `{{HOTEN}}`). Kiểm tra xem có thừa dấu cách bên trong ngoặc không.

### 4. Sau khi gửi được vài chục người thì script dừng lại và báo "Tạm dừng"?
- **Nguyên nhân**: Đây là tính năng tự động để tránh lỗi giới hạn 6 phút của Google.
- **Cách khắc phục**: Bạn **không cần làm gì cả**. Script đã tự hẹn giờ để 1 phút sau tự chạy tiếp. Bạn sẽ nhận được email báo cáo tổng kết khi mọi thứ hoàn tất.

### 5. Tôi thêm một cột "Thưởng chuyên cần" vào bảng lương thì làm sao để hiện trong PDF?
- **Nguyên nhân**: Script chưa biết cột mới đó nằm ở đâu.
- **Cách khắc phục**: 
    1. Mở phần code, tìm đến mục `MAP`.
    2. Thêm một dòng mới, ví dụ: `CHUYENCAN: "AP"` (trong đó AP là tên cột trong Excel).
    3. Trong tab `TEMPLATE`, hãy nhập thẻ `{{CHUYENCAN}}` vào vị trí bạn muốn.

### 6. Script báo "Lỗi Quota" hoặc "Hạn mức email"?
- **Nguyên nhân**: Bạn đã dùng hết số lượng email mà Google cho phép gửi trong 1 ngày (100 với Gmail thường).
- **Cách khắc phục**: Đợi sang ngày hôm sau. Script sẽ tự động nhận diện những người đã gửi thành công và chỉ gửi nốt phần còn lại cho bạn.

### 7. File PDF bị tràn lề hoặc quá bé khi in?
- **Nguyên nhân**: Bạn thiết kế tab `TEMPLATE` quá rộng hoặc quá hẹp.
- **Cách khắc phục**: Đảm bảo tổng độ rộng các cột trong tab `TEMPLATE` nằm trong khoảng **700px - 750px**. Bạn có thể dùng tính năng **"Tạo Sheet Mẫu"** trong menu để có khung chuẩn A4.

### 8. Tôi nên sửa cấu hình trong code hay tạo tab CONFIG?
- **Độ ưu tiên**: Tab **CONFIG** có ưu tiên **CAO NHẤT**. Nếu bạn tạo tab `CONFIG`, script sẽ lấy giá trị tại đó và bỏ qua các giá trị mặc định trong code.
- **Cách thiết lập tab CONFIG**:
    - Cột A: Tên cấu hình (Ví dụ: `SENDER_NAME`, `SHEET_NAME_PREFIX`).
    - Cột B: Giá trị bạn muốn đặt.
- **Lời khuyên**: Nếu bạn không rành về code, hãy sử dụng tab **CONFIG** để thay đổi thông tin công ty hoặc tên sheet tháng. Việc này giúp bạn tránh lỡ tay làm hỏng mã nguồn.

---
**Bạn vẫn gặp khó khăn?** Hãy liên hệ bộ phận kỹ thuật hoặc để lại Issue trên Github nhé!
