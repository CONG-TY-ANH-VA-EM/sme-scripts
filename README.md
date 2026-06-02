# SME Scripts Collection

Dự án này chứa các script hỗ trợ quản lý doanh nghiệp nhỏ và vừa (SME), tập trung vào các giải pháp trên Google Apps Script.

## Nội dung chính

### 1. Hệ thống Gửi Bảng Lương (google-appscripts)
Script tự động hóa quy trình gửi phiếu lương hàng tháng tới từng nhân viên qua email dưới dạng file PDF được định dạng chuyên nghiệp.

- **File**: [salary_send.js](./google-appscripts/salary_send.js)
- **📄 File mẫu dùng thử ngay (Google Sheets, public)**: [**Mở xem**](https://docs.google.com/spreadsheets/d/13IGNBDButCrmq3_459AF_pe2ATI-w2x5Vc634phaslE/edit) · [**Tạo bản sao để dùng**](https://docs.google.com/spreadsheets/d/13IGNBDButCrmq3_459AF_pe2ATI-w2x5Vc634phaslE/copy)
  *(Gồm sẵn tab `T1` dữ liệu mẫu, `CONFIG`, `TEMPLATE` — dữ liệu giả định, bấm "Tạo bản sao" rồi dán code để chạy.)*
- **Tài liệu**:
  - [📘 Cẩm nang chi tiết cho người không rành công nghệ](./google-appscripts/CAM_NANG_CHI_TIET.md)
  - [Hướng dẫn sử dụng](./google-appscripts/USER_GUIDE.md)
  - [Câu hỏi thường gặp & Xử lý sự cố (FAQ)](./google-appscripts/FAQ.md)
- **Tính năng nổi bật** (v2.0):
  - Gửi phiếu lương PDF hàng loạt qua email, tự động resume khi chạm giới hạn 6 phút.
  - **Lưu trữ Drive**: mỗi phiếu lưu vào thư mục theo `Năm / Tháng`, đặt tên theo nhân viên.
  - **Gửi thử** một phiếu về email người vận hành trước khi gửi thật.
  - **Chọn tháng** thủ công (gửi lương tháng cũ trong khi đang ở tháng mới).
  - **Nhật ký**: ghi lịch sử gửi vào tab `LOG` (thời gian, người nhận, trạng thái, link Drive).
  - **Bảo mật**: `SECURE_SHARE` chia sẻ file Drive riêng cho đúng email nhân viên (thay cho mật khẩu PDF — Apps Script không hỗ trợ mã hoá PDF).
  - **Dễ dàng cấu hình**: đổi thông tin công ty, vị trí cột (`MAP.*`), tùy chọn Drive/bảo mật qua tab `CONFIG` mà không cần mở code.

## Cách sử dụng nhanh
1. Mở Sheet bảng lương.
2. Cài đặt script từ file `salary_send.js`.
3. Xem [Hướng dẫn chi tiết](./google-appscripts/USER_GUIDE.md) để thiết kế bảng lương chuẩn.

---
*Phát triển bởi: AI Assistant - SME SOLUTIONS*