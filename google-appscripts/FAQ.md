# Hướng dẫn Xử lý Sự cố & Câu hỏi Thường gặp (FAQ)

Tài liệu này giúp bạn tự giải quyết các vấn đề nhanh chóng khi sử dụng công cụ gửi phiếu lương tự động.

> 📄 **File mẫu dùng thử ngay:** [Mở xem](https://docs.google.com/spreadsheets/d/13IGNBDButCrmq3_459AF_pe2ATI-w2x5Vc634phaslE/edit) · [Tạo bản sao để dùng](https://docs.google.com/spreadsheets/d/13IGNBDButCrmq3_459AF_pe2ATI-w2x5Vc634phaslE/copy) — gồm sẵn tab `T1` (dữ liệu mẫu), `CONFIG`, `TEMPLATE`.

---

### 1. Tại sao tôi bấm nút "Gửi" mà không có chuyện gì xảy ra?
- **Nguyên nhân**: Google yêu cầu bạn cấp quyền truy cập lần đầu tiên.
- **Cách khắc phục**: Nhìn lên góc trên màn hình, nếu thấy thông báo "Authorization Required", hãy bấm **Continue** > Chọn tài khoản Gmail của bạn > Bấm **Advanced** (Nâng cao) > Bấm **Go to SME Scripts (unsafe)** > Bấm **Allow** (Cho phép).

### 2. Lỗi: "Không tìm thấy sheet dữ liệu T1"
- **Nguyên nhân**: Tên tab tháng của bạn không khớp với cấu hình.
- **Cách khắc phục**: Kiểm tra xem tab chứa lương tháng 1 của bạn có tên chính xác là `T1` hay không (không có dấu cách, không có chữ "Tháng"). Nếu bạn muốn đặt là `Tháng 1`, hãy mở tab `CONFIG` và đổi giá trị `SHEET_NAME_PREFIX` từ `T` thành `Tháng ` (có dấu cách cuối).

### 3. Tại sao trong phiếu PDF hiện thẻ {{HOTEN}} mà không hiện tên người?
- **Nguyên nhân**: Bạn đã gõ sai thẻ Tag trong tab `TEMPLATE`.
- **Cách khắc phục**: Thẻ phải có **đôi ngoặc nhọn** `{{` và `}}`, và viết **HOA toàn bộ** (ví dụ: `{{HOTEN}}`). Kiểm tra xem có thừa dấu cách bên trong ngoặc không.

### 4. Sau khi gửi được vài chục người thì script dừng lại và báo "Tạm dừng"?
- **Nguyên nhân**: Đây là tính năng tự động để tránh lỗi giới hạn 6 phút của Google.
- **Cách khắc phục**: Bạn **không cần làm gì cả**. Script đã tự hẹn giờ để 1 phút sau tự chạy tiếp. Bạn sẽ nhận được email báo cáo tổng kết khi mọi thứ hoàn tất.

### 5. Tôi thêm một cột "Thưởng chuyên cần" vào bảng lương thì làm sao để hiện trong PDF?
- **Cách làm (không cần mở code)**: 
    1. Mở tab `CONFIG`. Ở cột A (THAM SỐ), thêm một dòng mới ghi `MAP.CHUYENCAN`; cột B (GIÁ TRỊ) ghi tên cột Excel, ví dụ `AP`.
    2. Trong tab `TEMPLATE`, nhập thẻ `{{CHUYENCAN}}` vào vị trí bạn muốn.
    3. Gửi lại — phiếu sẽ tự điền dữ liệu cột đó.
- **Lưu ý về định dạng tiền**: cột mới sẽ hiển thị nguyên văn. Nếu muốn nó được định dạng kiểu tiền tệ (vd `1.000.000`), hãy đặt tên bắt đầu bằng `L_` hoặc `PC_` (ví dụ `MAP.L_CHUYENCAN` + thẻ `{{L_CHUYENCAN}}`).

### 6. Script báo "Lỗi Quota" hoặc "Hạn mức email"?
- **Nguyên nhân**: Bạn đã dùng hết số lượng email mà Google cho phép gửi trong 1 ngày (100 với Gmail thường).
- **Cách khắc phục**: Đợi sang ngày hôm sau. Script sẽ tự động nhận diện những người đã gửi thành công và chỉ gửi nốt phần còn lại cho bạn.

### 7. File PDF bị tràn lề hoặc quá bé khi in?
- **Nguyên nhân**: Bạn thiết kế tab `TEMPLATE` quá rộng hoặc quá hẹp.
- **Cách khắc phục**: Đảm bảo tổng độ rộng các cột trong tab `TEMPLATE` nằm trong khoảng **700px - 750px**. Bạn có thể dùng tính năng **"Tạo Sheet Mẫu"** trong menu để có khung chuẩn A4.

### 8. Tôi nên sửa cấu hình trong code hay tạo tab CONFIG?
- **Độ ưu tiên**: Tab **CONFIG** có ưu tiên **CAO NHẤT**. Nếu bạn tạo tab `CONFIG`, script sẽ lấy giá trị tại đó và bỏ qua các giá trị mặc định trong code.
- **Cách thiết lập tab CONFIG**:
    1. Vào menu **SME Tools** > **Tạo Sheet Cấu hình (CONFIG)**.
    2. Cập nhật các giá trị tại cột B (màu vàng).
- **Lời khuyên**: Nếu bạn không rành về code, hãy sử dụng tab **CONFIG** để thay đổi thông tin công ty hoặc tên sheet tháng. Việc này giúp bạn tránh lỡ tay làm hỏng mã nguồn.

### 9. Tôi muốn đổi vị trí các cột (Vd: cột trạng thái từ AO sang AP) thì làm sao?
- **Đã hỗ trợ qua tab `CONFIG`** (không cần sửa code). Trong tab `CONFIG`, tìm (hoặc thêm) dòng tương ứng và đổi giá trị ở cột B:
    - Đổi cột trạng thái: `MAP.SENT_STATUS` → `AP`
    - Đổi cột email: `MAP.EMAIL` → cột mới
    - Đổi cột họ tên: `MAP.HOTEN`, lương chuyển khoản: `MAP.LUONGCK`, v.v.
- Nguyên tắc: bất kỳ dòng nào bắt đầu bằng `MAP.` trong CONFIG sẽ ghi đè vị trí cột tương ứng.

### 10. "Gửi THỬ" khác gì với gửi thật?
- **Gửi THỬ** (menu **SME Tools** > **Gửi THỬ về email tôi...**) chỉ tạo một phiếu và gửi về email bạn chọn (mặc định email của bạn) để xem trước. Nó **không** đánh dấu "Thành công", **không** lưu Drive, **không** gửi cho nhân viên. Dùng thoải mái để kiểm tra mẫu trước khi gửi thật.

### 11. Phiếu lương được lưu ở đâu trên Drive?
- Khi `SAVE_TO_DRIVE = true`, mỗi phiếu lưu vào: `[Thư mục gốc] / Năm YYYY / Tháng MM-YYYY / PhieuLuong_TenNV_...pdf`.
- Thư mục gốc tạo tự động trong **My Drive của bạn** (mặc định chỉ mình bạn xem). ⚠️ **Đừng chia sẻ thư mục này cho người ngoài** — bên trong là lương toàn bộ nhân viên.
- Muốn dùng thư mục có sẵn? Điền ID thư mục vào `DRIVE_ROOT_FOLDER_ID` trong CONFIG.

### 12. Làm sao đặt mật khẩu cho file PDF phiếu lương?
- Google Apps Script **không** hỗ trợ đặt mật khẩu / mã hoá PDF. Thay vào đó, bật `SECURE_SHARE = true` trong CONFIG: file PDF trên Drive sẽ **chỉ chia sẻ riêng cho đúng email nhân viên**, và email kèm link riêng tư. Đây là cách bảo mật thay thế. Nếu việc chia sẻ riêng thất bại (vd domain chặn), hệ thống tự đính kèm PDF để nhân viên vẫn nhận được phiếu.

### 13. Tôi muốn xem lại lịch sử đã gửi cho ai, lúc nào?
- Khi `ENABLE_LOG = true`, công cụ tạo tab `LOG` ghi lại từng lần gửi: thời gian, kỳ lương, hàng, họ tên, email, trạng thái và link Drive. Mở tab này để đối chiếu hoặc kiểm toán.

### 14. Tôi gửi lương tháng trước, nhưng máy lại nhận tháng hiện tại?
- Dùng **SME Tools** > **Gửi Phiếu Lương (Chọn Tháng...)** và nhập tháng cần gửi (vd `5` hoặc `5/2025`). Tùy chọn "Tháng Hiện Tại" luôn lấy theo lịch của ngày hôm nay.

---
**Bạn vẫn gặp khó khăn?** Hãy liên hệ bộ phận kỹ thuật hoặc để lại Issue trên Github nhé!
