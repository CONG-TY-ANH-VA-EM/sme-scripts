# SME Scripts Collection

Dự án này chứa các script hỗ trợ quản lý doanh nghiệp nhỏ và vừa (SME), tập trung vào các giải pháp trên Google Apps Script.

## Nội dung chính

### 1. Hệ thống Gửi Bảng Lương (google-appscripts)
Script tự động hóa quy trình gửi phiếu lương hàng tháng tới từng nhân viên qua email dưới dạng file PDF được định dạng chuyên nghiệp.

- **File**: [salary_send.js](file:///opt/workspace/sme-scripts/google-appscripts/salary_send.js)
- **Tài liệu**:
  - [Hướng dẫn sử dụng chi tiết](file:///opt/workspace/sme-scripts/google-appscripts/USER_GUIDE.md)
  - [Câu hỏi thường gặp & Xử lý sự cố (FAQ)](file:///opt/workspace/sme-scripts/google-appscripts/FAQ.md)
- **Tính năng nổi bật**:
  - Tự động nhận diện danh sách nhân viên linh hoạt.
  - Tối ưu hiệu suất tạo PDF hàng loạt.
  - Sửa lỗi lệch tháng tự động theo thời gian thực.
  - **Dễ dàng cấu hình**: Thay đổi thông số qua tab `CONFIG` ngay trên Sheet mà không cần mở code.

## Cách sử dụng nhanh
1. Mở Sheet bảng lương.
2. Cài đặt script từ file `salary_send.js`.
3. Xem [Hướng dẫn chi tiết](./google-appscripts/USER_GUIDE.md) để thiết kế bảng lương chuẩn.

---
*Phát triển bởi: AI Assistant - SME SOLUTIONS*