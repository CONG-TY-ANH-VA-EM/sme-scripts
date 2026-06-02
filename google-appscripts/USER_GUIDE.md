# Hướng dẫn Sử dụng Công cụ Gửi Phiếu Lương Tự động (SME Tools)

Chào mừng bạn! Tài liệu này sẽ hướng dẫn bạn cách cài đặt và sử dụng công cụ gửi phiếu lương qua Email hoàn toàn tự động. 

Chúng tôi đã thiết kế tài liệu này dành riêng cho người dùng không chuyên về máy tính (non-tech), vì vậy các bước sẽ được mô tả rất chi tiết.

---

## BƯỚC 1: CÀI ĐẶT MÃ NGUỒN (CHỈ LÀM 1 LẦN DUY NHẤT)

Trước khi bắt đầu, bạn cần "dán" bộ mã xử lý vào bảng tính Google Sheets của bạn.

1.  Mở tệp **Google Sheets** bảng lương của bạn.
2.  Trên thanh thực đơn phía trên, chọn **Tiện ích mở rộng (Extensions)** > **Apps Script**. một trang web mới sẽ hiện ra.
3.  Nếu có mã cũ trong đó (thường là vài dòng chữ `function myFunction() { ... }`), hãy xóa hết đi để trang trắng tinh.
4.  Mở tệp mã nguồn [salary_send.js](./salary_send.js) trên máy tính của bạn, chọn tất cả (Ctrl + A) và sao chép (Ctrl + C).
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

0.  **Nhanh nhất**: [tạo bản sao file Google Sheets mẫu](https://docs.google.com/spreadsheets/d/13IGNBDButCrmq3_459AF_pe2ATI-w2x5Vc634phaslE/copy) (đã có sẵn `T1`, `CONFIG`, `TEMPLATE`) rồi thay dữ liệu thật.
1.  Hoặc dùng file CSV mẫu để đảm bảo chính xác: [payroll_sample_2025.csv](./payroll_sample_2025.csv).
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
2.  Một trang mới tên là `CONFIG` sẽ tự động hiện lên, chia thành các nhóm rõ ràng.
3.  Bạn chỉ cần sửa các giá trị ở **cột B (ô màu vàng)**. Không cần mở code. Tab CONFIG cho phép điều chỉnh:
    - **Thông tin công ty**: tên, địa chỉ, hotline, email liên hệ (hiện trên phiếu).
    - **Cấu trúc dữ liệu**: tiền tố tab tháng (`SHEET_NAME_PREFIX`), hàng bắt đầu (`START_ROW`).
    - **Email**: tiền tố tiêu đề/tên file, có đính kèm PDF hay không (`ATTACH_PDF_TO_EMAIL`).
    - **Lưu trữ Drive**: bật/tắt (`SAVE_TO_DRIVE`), tên/ID thư mục gốc.
    - **Bảo mật & nhật ký**: `SECURE_SHARE`, `ENABLE_LOG`.
    - **Vị trí cột**: các dòng dạng `MAP.HOTEN`, `MAP.EMAIL`, `MAP.LUONGCK`... — đổi khi bảng lương của bạn xếp cột khác mẫu (xem FAQ mục 5 & 9).

### 4.2. Tạo trang Mẫu phiếu lương (TEMPLATE)
1.  Vào menu **SME Tools** > **Tạo Sheet Mẫu (TEMPLATE)**.
2.  Một trang mới tên là `TEMPLATE` sẽ hiện lên với khung mẫu phiếu lương chuyên nghiệp.
3.  **Bạn có thể tự do chỉnh sửa**: Đổi màu, chèn Logo công ty, thay đổi font chữ...
4.  **Tuyệt đối không xóa các thẻ Tag**: Các chữ nằm trong dấu ngoặc nhọn kép như `{{HOTEN}}`, `{{TONGLUONG}}`... là nơi máy tính sẽ tự động điền dữ liệu vào. Bạn có thể di chuyển chúng sang ô khác nhưng đừng xóa hoặc sửa tên chúng.

---

## BƯỚC 5: GỬI THỬ TRƯỚC (RẤT NÊN LÀM)

Trước khi gửi cho toàn bộ nhân viên, hãy gửi thử **một phiếu** về email của chính bạn để kiểm tra phiếu PDF có đẹp và đúng số liệu không.

1.  Vào menu **SME Tools** > **Gửi THỬ về email tôi...**.
2.  Nhập email nhận bản thử (để trống = email của bạn) và số hàng muốn lấy dữ liệu (mặc định hàng 2).
3.  Mở email, xem file PDF đính kèm. Nếu đẹp và đúng → sang Bước 6 gửi thật.
4.  Lưu ý: gửi thử **không** đánh dấu "Thành công", **không** lưu Drive — hoàn toàn an toàn để thử nhiều lần.

---

## BƯỚC 6: GỬI PHIẾU LƯƠNG THẬT

Khi mọi thứ đã sẵn sàng:

1.  Mở sheet chứa dữ liệu lương của tháng bạn muốn gửi (Ví dụ sheet `T1`).
2.  Chọn một trong hai cách:
    - **SME Tools** > **Gửi Phiếu Lương (Tháng Hiện Tại)** — gửi cho đúng tháng theo lịch hôm nay.
    - **SME Tools** > **Gửi Phiếu Lương (Chọn Tháng...)** — nhập tháng cần gửi (vd `5` hoặc `5/2025`). Dùng khi cần gửi lương tháng cũ trong khi đã sang tháng mới.
3.  Hệ thống sẽ bắt đầu chạy. Bạn sẽ thấy cột trạng thái (mặc định cột `AO`) của mỗi hàng lần lượt hiện chữ **"Thành công"** (nền xanh) hoặc **"Lỗi: ..."** (nền đỏ nếu có trục trặc).
4.  Nếu danh sách quá dài (vài trăm người), Google có thể tạm dừng. Đừng lo, mã nguồn có chế độ **tự động chạy tiếp** sau 1 phút mà bạn không cần nhấn gì thêm. Khi xong, bạn nhận được **email báo cáo tổng kết**.
5.  Gửi nhầm/cần gửi lại từ đầu? Vào **SME Tools** > **Xóa trạng thái gửi (Reset)** rồi gửi lại.

---

## BƯỚC 7: LƯU TRỮ & NHẬT KÝ (TỰ ĐỘNG)

Ngoài việc gửi email, công cụ còn giúp bạn lưu trữ và theo dõi:

- **Lưu Drive** (khi `SAVE_TO_DRIVE = true`): mỗi phiếu PDF được lưu vào Google Drive theo cấu trúc
  `[Thư mục gốc] / Năm 2025 / Tháng 05-2025 / PhieuLuong_Ten_Nhan_Vien_...pdf`.
  Thư mục gốc tạo tự động trong **My Drive của bạn** và **mặc định chỉ mình bạn xem được**.
  > ⚠️ **Không chia sẻ thư mục gốc này cho người ngoài** — bên trong là lương của toàn bộ nhân viên.
- **Nhật ký LOG** (khi `ENABLE_LOG = true`): một tab tên `LOG` ghi lại từng lần gửi — thời gian, kỳ lương, hàng, họ tên, email, trạng thái, link Drive. Dùng để đối chiếu/kiểm toán.
- **Bảo mật `SECURE_SHARE`**: nếu bật (`true`), mỗi file PDF trên Drive chỉ được chia sẻ riêng cho **đúng email nhân viên đó**, và email sẽ kèm link riêng tư. (Lưu ý kỹ thuật: Google Apps Script **không** đặt được mật khẩu PDF; `SECURE_SHARE` là cách bảo mật thay thế. Nếu chia sẻ thất bại, hệ thống tự động đính kèm PDF để nhân viên vẫn nhận được phiếu.)

---

## CÁC CÂU HỎI THƯỜNG GẶP (FAQ)

> [!TIP]
> **Tôi gặp lỗi "Hết hạn mức gửi thư" (Quota)?**
> Mỗi ngày Google chỉ cho phép gửi tối đa một số lượng email nhất định (tài khoản thường khoảng 100 email/ngày). Nếu hết hạn mức, bạn chỉ cần đợi đến ngày hôm sau và nhấn gửi lại, máy sẽ tự động gửi tiếp cho những người chưa nhận được thư.

> [!NOTE]
**Làm sao để biết ai chưa nhận được thư?**
Bạn hãy xem ở cột cuối cùng (`SentStatus`). Những ai có chữ "Thành công" là đã nhận được thư. Những ai để trống là chưa gửi.

> [!TIP]
> **Thư được gửi từ địa chỉ Email nào?**
> Hệ thống sẽ sử dụng chính tài khoản Email mà bạn đang dùng để chạy công cụ này để gửi cho nhân viên. Vì vậy, bạn không cần điền "Email gửi đi". Ô `CONTACT_EMAIL` trong phần cấu hình chỉ là địa chỉ hiển thị trên phiếu lương để nhân viên liên hệ khi có thắc mắc.

> [!IMPORTANT]
> **Tôi muốn sửa lại mẫu phiếu lương?**
> Bạn cứ vào sheet `TEMPLATE` để sửa. Sau khi sửa xong, chỉ cần vào gửi lại, hệ thống sẽ lấy mẫu mới nhất để áp dụng.

---
*Nếu gặp bất kỳ khó khăn nào, hãy liên hệ bộ phận hỗ trợ kỹ thuật hoặc xem file [FAQ chi tiết tại đây](./FAQ.md).*

