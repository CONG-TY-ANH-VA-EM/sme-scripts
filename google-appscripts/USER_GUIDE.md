# Hướng dẫn Sử dụng Hệ thống Gửi Phiếu Lương Tự động 2025

Chào mừng bạn đến với phiên bản nâng cấp 2025 của công cụ gửi phiếu lương. Phiên bản này hỗ trợ cơ chế **Template Sheet**, cho phép bạn tự thiết kế mẫu phiếu lương chuyên nghiệp ngay trên Google Sheets.

## 1. Chuẩn bị Bảng tính lương
Bảng lương của bạn cần có cấu trúc 40 cột để khớp với các thẻ Tag có sẵn.

**Cách nhanh nhất**:
1. Tải file mẫu [payroll_sample_2025.csv](./payroll_sample_2025.csv).
2. Mở Google Sheets > **Tệp (File)** > **Nhập (Import)** > **Tải lên** file CSV này.
3. Đổi tên sheet vừa nhập thành `T + <số tháng>` (Ví dụ: `T1`).

*Lưu ý: Dữ liệu nhân viên bắt đầu từ hàng 2.*

## 2. Thiết kế Mẫu Phiếu Lương (Tab TEMPLATE)
Thay vì vẽ bằng code, giờ đây bạn có thể tự thiết kế phiếu lương ngay trên Google Sheets.

1. Tạo một Sheet mới và đặt tên chính xác là **TEMPLATE**.
2. Thiết kế giao diện (Font, màu sắc, logo, bảng biểu) theo ý muốn.
3. Tại những vị trí cần điền dữ liệu động, hãy nhập các thẻ **{{TAG}}** tương ứng bên dưới.

### Danh sách các thẻ (Tags) khả dụng:
- **Thông tin chung**: `{{THANGNAM}}`, `{{SENDER_NAME}}`, `{{SENDER_ADDRESS}}`, `{{SENDER_HOTLINE}}`, `{{CONTACT_EMAIL}}`
- **Thông tin nhân viên**: `{{HOTEN}}`, `{{VITRI}}`, `{{STK}}`, `{{NGANHANG}}`, `{{NGAYGUI}}`
- **Công xá**: `{{NGAYCONGCHUAN}}`, `{{TONGGIOTT}}`, `{{GIO150}}`, `{{GIO200}}`, `{{GIO300}}`, `{{RO}}`, `{{P}}`, `{{L}}`
- **Lương & Phụ cấp**: `{{L_COBAN}}`, `{{PC_ANTRUA}}`, `{{PC_DIENTHOAI}}`, `{{PC_DILAI}}`, `{{L_NGAYCONG}}`, `{{L_NGHIPHEP}}`, `{{L_OT}}`, `{{L_KPI}}`, `{{PC_MAYTINH}}`, `{{CONGTACPHI}}`, `{{THUONG}}`
- **Khấu trừ & Thực lĩnh**: `{{BHXH}}`, `{{THUETNCN}}`, `{{TONGLUONG}}`, `{{THUCLINH}}`, `{{TAMUNG}}`, `{{DANHAN}}`, `{{LUONGCK}}`

> [!TIP]
> Script sẽ tự động định dạng số tiền (thêm dấu chấm phân cách hàng nghìn) cho các thẻ liên quan đến lương thưởng.

---

## 3. Cài đặt Script
1. Mở Sheet > **Tiện ích mở rộng (Extensions)** > **Apps Script**.
2. Copy toàn bộ nội dung file `salary_send.js` vào trình soạn thảo.
3. Nhấn **Lưu** và làm mới trang Google Sheets.

## 4. Cách chạy Gửi Phiếu Lương
1. Bấm vào menu **SME Tools** > **Gửi Phiếu Lương PDF (Tháng Hiện Tại)**.
2. Cấp quyền (lần đầu) và chờ thông báo kết quả.

---

## 5. Tùy chỉnh Cấu hình nâng cao
Bạn có thể tạo sheet **CONFIG** để ghi đè các tham số sau:
- `TEMPLATE_SHEET_NAME`: Nếu bạn muốn đặt tên tab mẫu khác.
- `START_ROW`: Nếu dữ liệu không bắt đầu từ hàng 2.

---
*Phát triển bởi: AI Assistant - SME SOLUTIONS*
