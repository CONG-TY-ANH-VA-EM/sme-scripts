# Hướng dẫn Sử dụng Công cụ Gửi Phiếu Lương Tự động (SME Tools)

Chào mừng bạn! Tài liệu này sẽ hướng dẫn bạn cách cài đặt và sử dụng công cụ gửi phiếu lương qua Email hoàn toàn tự động. 

Chúng tôi đã thiết kế tài liệu này dành riêng cho người dùng không chuyên về máy tính (non-tech), vì vậy các bước sẽ được mô tả rất chi tiết.

---

## BƯỚC 1: CÀI ĐẶT MÃ NGUỒN (CHỈ LÀM 1 LẦN DUY NHẤT)

Trước khi bắt đầu, bạn cần "dán" bộ mã xử lý vào bảng tính Google Sheets của bạn.

1.  Mở tệp **Google Sheets** bảng lương của bạn.
2.  Trên thanh thực đơn phía trên, chọn **Tiện ích mở rộng (Extensions)** > **Apps Script**. một trang web mới sẽ hiện ra.
3.  Nếu có mã cũ trong đó (thường là vài dòng chữ `function myFunction() { ... }`), hãy xóa hết đi để trang trắng tinh.
4.  Mở tệp mã nguồn [salary_send.js](file:///opt/workspace/sme-scripts/google-appscripts/salary_send.js) trên máy tính của bạn, chọn tất cả (Ctrl + A) và sao chép (Ctrl + C).
5.  Quay lại trang Apps Script vừa mở, dán (Ctrl + V) toàn bộ mã vào đó.
6.  Nhấn nút **Lưu (biểu tượng đĩa mềm)** ở phía trên và đặt tên dự án là "SME Payroll Tool" rồi nhấn **Đồng ý**.

---

## BƯỚC 2: CẤP QUYỀN VÀ KÍCH HOẠT MENU

Sau khi dán mã, bạn cần cho phép Google chạy mã này trên tài khoản của bạn. Đây là bước quan trọng nhất.

1.  Tại trang Apps Script, nhìn lên thanh thực đơn, chỗ có chữ **"sendMonthlyPayrollEmails_WithPDF"**, bạn hãy nhấn nút **Chạy (Run)** bên cạnh nó.
2.  Một cửa sổ "Cần có quyền truy cập" hiện lên, hãy nhấn **Xem quyền (Review Permissions)**.
3.  Chọn tài khoản Email của bạn.
4.  **QUAN TRỌNG:** Một cảnh báo hiện lên nói rằng **"Google chưa xác minh ứng dụng này"**. Đừng lo lắng, đây là mã nguồn nội bộ của bạn nên Google chưa kiểm duyệt.
    - Nhấn vào chữ **Nâng cao (Advanced)** ở góc dưới bên trái.
    - Nhấn tiếp vào dòng chữ nhỏ phía dưới: **Đi tới SME Payroll Tool (không an toàn)** hoặc **Go to SME Payroll Tool (unsafe)**.
5.  Một danh sách các quyền hiện ra, bạn kéo xuống dưới cùng và nhấn **Cho phép (Allow)**.
6.  Khi thấy dòng chữ "Đã hoàn thành thực thi" ở phía dưới là xong. 
7.  **Bây giờ, hãy quay lại tab Google Sheets của bạn và nhấn phím F5 (tải lại trang).** Bạn sẽ thấy một thực đơn mới tên là **SME Tools** hiện lên bên cạnh menu "Trợ giúp".

---

## BƯỚC 3: CHUẨN BỊ DỮ LIỆU LƯƠNG

Để công cụ hoạt động đúng, bảng tính của bạn cần được sắp xếp theo đúng quy chuẩn.

1.  Bạn nên sử dụng file mẫu để đảm bảo chính xác nhất: [payroll_sample_2025.csv](./payroll_sample_2025.csv).
2.  Cách đưa file mẫu vào Google Sheets:
    - Vào **Tệp (File)** > **Nhập (Import)** > **Tải lên (Upload)** và chọn file CSV mẫu.
    - Đặt tên Sheet (tab) là `T1` (nếu là lương tháng 1), `T2` (lương tháng 2), v.v.
3.  **Lưu ý cực kỳ quan trọng**:
    - Dòng số 1 phải là tiêu đề.
    - Dữ liệu nhân viên bắt đầu từ dòng số 2.
    - Cột Email (mặc định là cột F) phải chứa địa chỉ email chính xác để gửi thư.

---

## BƯỚC 4: THIẾT LẬP CẤU HÌNH VÀ MẪU PHIẾU LƯƠNG

Công cụ này cho phép bạn tự thiết kế mẫu phiếu lương cực kỳ đẹp mắt ngay trên Google Sheet.

### 4.1. Tạo trang Cấu hình (CONFIG)
1.  Vào menu **SME Tools** > **Tạo Sheet Cấu hình (CONFIG)**.
2.  Một trang mới tên là `CONFIG` sẽ tự động hiện lên.
3.  Bạn chỉ cần sửa các thông tin ở **cột B (màu vàng)** như: Tên công ty, địa chỉ, số điện thoại... Những thông tin này sẽ tự động hiện lên trên phiếu lương gửi cho nhân viên.

### 4.2. Tạo trang Mẫu phiếu lương (TEMPLATE)
1.  Vào menu **SME Tools** > **Tạo Sheet Mẫu (TEMPLATE)**.
2.  Một trang mới tên là `TEMPLATE` sẽ hiện lên với khung mẫu phiếu lương chuyên nghiệp.
3.  **Bạn có thể tự do chỉnh sửa**: Đổi màu, chèn Logo công ty, thay đổi font chữ...
4.  **Tuyệt đối không xóa các thẻ Tag**: Các chữ nằm trong dấu ngoặc nhọn kép như `{{HOTEN}}`, `{{TONGLUONG}}`... là nơi máy tính sẽ tự động điền dữ liệu vào. Bạn có thể di chuyển chúng sang ô khác nhưng đừng xóa hoặc sửa tên chúng.

---

## BƯỚC 5: GỬI PHIẾU LƯƠNG

Khi mọi thứ đã sẵn sàng, hãy thực hiện gửi:

1.  Mở sheet chứa dữ liệu lương của tháng bạn muốn gửi (Ví dụ sheet `T1`).
2.  Vào menu **SME Tools** > **Gửi Phiếu Lương PDF (Tháng Hiện Tại)**.
3.  Hệ thống sẽ bắt đầu chạy. Bạn sẽ thấy cột cuối cùng của mỗi hàng lần lượt hiện chữ **"Thành công"** và kèm theo link file PDF đã gửi.
4.  Nếu danh sách quá dài (vài trăm người), Google có thể tạm dừng. Đừng lo, mã nguồn có chế độ **tự động chạy tiếp** sau 1 phút mà bạn không cần nhấn gì thêm.

---

## CÁC CÂU HỎI THƯỜNG GẶP (FAQ)

> [!TIP]
> **Tôi gặp lỗi "Hết hạn mức gửi thư" (Quota)?**
> Mỗi ngày Google chỉ cho phép gửi tối đa một số lượng email nhất định (tài khoản thường khoảng 100 email/ngày). Nếu hết hạn mức, bạn chỉ cần đợi đến ngày hôm sau và nhấn gửi lại, máy sẽ tự động gửi tiếp cho những người chưa nhận được thư.

> [!NOTE]
> **Làm sao để biết ai chưa nhận được thư?**
> Bạn hãy xem ở cột cuối cùng (`SentStatus`). Những ai có chữ "Thành công" là đã nhận được thư. Những ai để trống là chưa gửi.

> [!IMPORTANT]
> **Tôi muốn sửa lại mẫu phiếu lương?**
> Bạn cứ vào sheet `TEMPLATE` để sửa. Sau khi sửa xong, chỉ cần vào gửi lại, hệ thống sẽ lấy mẫu mới nhất để áp dụng.

---
*Nếu gặp bất kỳ khó khăn nào, hãy liên hệ bộ phận hỗ trợ kỹ thuật hoặc xem file [FAQ chi tiết tại đây](file:///opt/workspace/sme-scripts/google-appscripts/FAQ.md).*

